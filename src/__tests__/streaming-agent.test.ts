/**
 * StreamingAgentExecutor 测试
 */

import { StreamingAgentExecutor, StreamChunk } from "../agent/StreamingAgentExecutor";
import { RecoveryManager, RecoveryAction, RecoverableStep } from "../agent/RecoveryManager";
import { AgentTracer, getTracer, createTracer } from "../agent/tracing";

// ========== StreamingAgentExecutor 测试 ==========

describe("StreamingAgentExecutor", () => {
  describe("流式输出", () => {
    it("应该创建流式执行器实例", () => {
      const executor = new StreamingAgentExecutor();
      expect(executor).toBeDefined();
    });

    it("executeStream 应该返回 AsyncGenerator", async () => {
      const executor = new StreamingAgentExecutor();
      const stream = executor.executeStream({
        userMessage: "读取当前选区",
      });

      expect(stream[Symbol.asyncIterator]).toBeDefined();
    });

    it("流式输出应该包含 status 类型的 chunk", async () => {
      const executor = new StreamingAgentExecutor();
      const chunks: StreamChunk[] = [];

      const stream = executor.executeStream({
        userMessage: "读取数据",
      });

      // 只收集前几个 chunk（避免完整执行）
      let count = 0;
      for await (const chunk of stream) {
        chunks.push(chunk);
        count++;
        if (count >= 3) break;
      }

      expect(chunks.length).toBeGreaterThan(0);
      expect(chunks[0].type).toBe("status");
      expect(chunks[0].content).toContain("理解");
    });

    it("chunk 应该包含时间戳和进度", async () => {
      const executor = new StreamingAgentExecutor();
      const stream = executor.executeStream({
        userMessage: "测试",
      });

      const firstChunk = (await stream.next()).value as StreamChunk;

      expect(firstChunk.timestamp).toBeDefined();
      expect(typeof firstChunk.timestamp).toBe("number");
      expect(firstChunk.progress).toBeDefined();
    });
  });

  describe("取消操作", () => {
    it("应该支持通过 AbortSignal 取消", async () => {
      const executor = new StreamingAgentExecutor();
      const controller = new AbortController();

      const stream = executor.executeStream(
        { userMessage: "长任务" },
        { signal: controller.signal }
      );

      // 立即取消
      controller.abort();

      const result = await (async () => {
        for await (const chunk of stream) {
          // 应该很快结束
          return chunk;
        }
      })();

      // 结果可能是第一个 chunk 或取消结果
      expect(result).toBeDefined();
    });
  });
});

// ========== RecoveryManager 测试 ==========

describe("RecoveryManager", () => {
  let manager: RecoveryManager;

  beforeEach(() => {
    manager = new RecoveryManager();
  });

  describe("策略匹配", () => {
    it("应该返回已注册的策略列表", () => {
      const strategies = manager.getStrategies();
      expect(strategies.length).toBeGreaterThan(0);
      expect(strategies.some((s) => s.name === "network_error")).toBe(true);
      expect(strategies.some((s) => s.name === "range_not_found")).toBe(true);
    });

    it("网络错误应该匹配重试策略", async () => {
      const step = {
        id: "step1",
        action: "excel_read_range",
        description: "读取范围",
        parameters: { range: "A1:B10" },
        isWriteOperation: false,
        dependsOn: [],
      };

      const result = await manager.recover(step, new Error("Network timeout"));

      expect(result).not.toBeNull();
      expect(result?.type).toBe("retry");
      expect(result?.delay).toBeGreaterThan(0);
    });

    it("范围不存在应该匹配替代策略", async () => {
      const step = {
        id: "step1",
        action: "excel_read_range",
        description: "读取范围",
        parameters: { range: "A1:B10" },
        isWriteOperation: false,
        dependsOn: [],
      };

      const result = await manager.recover(step, new Error("Range not found"));

      expect(result).not.toBeNull();
      expect(result?.type).toBe("substitute");
      expect(result?.alternativeStep).toBeDefined();
    });

    it("权限错误 + 非关键操作应该跳过", async () => {
      const step = {
        id: "step1",
        action: "excel_read_info",
        description: "读取信息",
        parameters: {},
        isWriteOperation: false,
        dependsOn: [],
      };

      const result = await manager.recover(step, new Error("Permission denied"));

      expect(result).not.toBeNull();
      expect(result?.type).toBe("skip");
    });

    it("权限错误 + 关键操作应该中止", async () => {
      const step = {
        id: "step1",
        action: "excel_write_range",
        description: "写入范围",
        parameters: {},
        isWriteOperation: true,
        dependsOn: [],
      };

      const result = await manager.recover(step, new Error("Access denied"));

      expect(result).not.toBeNull();
      expect(result?.type).toBe("abort");
    });
  });

  describe("重试限制", () => {
    it("应该限制最大重试次数", async () => {
      const step = {
        id: "step_retry_test",
        action: "excel_api_call",
        description: "API 调用",
        parameters: {},
        isWriteOperation: false,
        dependsOn: [],
      };

      // 第 1 次
      const r1 = await manager.recover(step, new Error("Network error"));
      expect(r1?.type).toBe("retry");

      // 第 2 次
      const r2 = await manager.recover(step, new Error("Network error"));
      expect(r2?.type).toBe("retry");

      // 第 3 次
      const r3 = await manager.recover(step, new Error("Network error"));
      expect(r3?.type).toBe("retry");

      // 第 4 次应该降级到其他策略
      const r4 = await manager.recover(step, new Error("Network error"));
      // 由于达到重试上限，会匹配默认策略（skip）
      expect(r4?.type).not.toBe("retry");
    });

    it("resetRetryCount 应该重置计数", async () => {
      const step = {
        id: "step_reset_test",
        action: "excel_api",
        description: "测试",
        parameters: {},
        isWriteOperation: false,
        dependsOn: [],
      };

      // 触发几次重试
      await manager.recover(step, new Error("Network timeout"));
      await manager.recover(step, new Error("Network timeout"));

      // 重置
      manager.resetRetryCount("step_reset_test");

      // 应该又可以重试
      const result = await manager.recover(step, new Error("Network timeout"));
      expect(result?.type).toBe("retry");
    });
  });
});

// ========== AgentTracer 测试 ==========

describe("AgentTracer", () => {
  let tracer: AgentTracer;

  beforeEach(() => {
    tracer = createTracer({ consoleOutput: false });
  });

  describe("Span 管理", () => {
    it("应该创建和结束 Span", () => {
      const span = tracer.startSpan("test-span");
      expect(span.id).toBeDefined();
      expect(span.name).toBe("test-span");
      expect(span.status).toBe("running");

      tracer.endSpan("success");

      const exported = tracer.export();
      expect(exported.spans.length).toBe(1);
      expect(exported.spans[0].status).toBe("success");
      expect(exported.spans[0].duration).toBeDefined();
    });

    it("应该支持嵌套 Span", () => {
      tracer.startSpan("parent");
      const childSpan = tracer.startSpan("child");

      expect(childSpan.parentId).toBeDefined();

      tracer.endSpan(); // child
      tracer.endSpan(); // parent

      const exported = tracer.export();
      expect(exported.spans.length).toBe(2);
    });

    it("应该支持 Span 属性", () => {
      tracer.startSpan("with-attrs", { key1: "value1" });
      tracer.setSpanAttribute("key2", 42);
      tracer.endSpan();

      const exported = tracer.export();
      expect(exported.spans[0].attributes.key1).toBe("value1");
      expect(exported.spans[0].attributes.key2).toBe(42);
    });

    it("应该支持 Span 事件", () => {
      tracer.startSpan("with-events");
      tracer.addSpanEvent("checkpoint", { progress: 50 });
      tracer.endSpan();

      const exported = tracer.export();
      expect(exported.spans[0].events.length).toBe(1);
      expect(exported.spans[0].events[0].name).toBe("checkpoint");
    });
  });

  describe("日志", () => {
    it("应该记录不同级别的日志", () => {
      // 创建日志级别为 debug 的 tracer
      const debugTracer = createTracer({ logLevel: "debug", consoleOutput: false });
      debugTracer.debug("debug message");
      debugTracer.info("info message");
      debugTracer.warn("warn message");
      debugTracer.error("error message");

      const exported = debugTracer.export();
      expect(exported.logs.length).toBe(4);
    });

    it("应该根据配置过滤日志级别", () => {
      const warnTracer = createTracer({ logLevel: "warn", consoleOutput: false });

      warnTracer.debug("debug");
      warnTracer.info("info");
      warnTracer.warn("warn");
      warnTracer.error("error");

      const exported = warnTracer.export();
      // 只有 warn 和 error
      expect(exported.logs.length).toBe(2);
    });

    it("日志应该关联到当前 Span", () => {
      tracer.startSpan("logged-span");
      tracer.info("inside span");
      tracer.endSpan();

      const exported = tracer.export();
      expect(exported.logs[0].spanId).toBe(exported.spans[0].id);
    });
  });

  describe("trace 辅助函数", () => {
    it("trace 应该自动创建和结束 Span", async () => {
      const result = await tracer.trace("async-op", async () => {
        return 42;
      });

      expect(result).toBe(42);

      const exported = tracer.export();
      expect(exported.spans.length).toBe(1);
      expect(exported.spans[0].status).toBe("success");
    });

    it("trace 应该捕获错误并标记 Span", async () => {
      await expect(
        tracer.trace("failing-op", async () => {
          throw new Error("test error");
        })
      ).rejects.toThrow("test error");

      const exported = tracer.export();
      expect(exported.spans[0].status).toBe("error");
      expect(exported.spans[0].error).toBe("test error");
    });

    it("traceSync 应该同步执行", () => {
      const result = tracer.traceSync("sync-op", () => {
        return "sync result";
      });

      expect(result).toBe("sync result");

      const exported = tracer.export();
      expect(exported.spans.length).toBe(1);
    });
  });

  describe("性能摘要", () => {
    it("应该返回正确的性能摘要", () => {
      tracer.startSpan("span1");
      tracer.endSpan();

      tracer.startSpan("span2");
      tracer.endSpan("error", "some error");

      tracer.warn("a warning");
      tracer.error("an error");

      const summary = tracer.getPerformanceSummary();

      expect(summary.totalSpans).toBe(2);
      expect(summary.errorCount).toBe(1);
      expect(summary.warningCount).toBe(1);
    });
  });

  describe("导出与清理", () => {
    it("export 应该返回完整数据", () => {
      tracer.startSpan("test");
      tracer.info("log");
      tracer.endSpan();

      const exported = tracer.export();

      expect(exported.traceId).toBeDefined();
      expect(exported.spans).toBeDefined();
      expect(exported.logs).toBeDefined();
      expect(exported.metadata.version).toBe("4.1");
    });

    it("clear 应该清空所有数据", () => {
      tracer.startSpan("test");
      tracer.info("log");
      tracer.endSpan();

      tracer.clear();

      const exported = tracer.export();
      expect(exported.spans.length).toBe(0);
      expect(exported.logs.length).toBe(0);
    });
  });
});

// ========== 全局 Tracer 测试 ==========

describe("全局 Tracer", () => {
  afterEach(() => {
    // 测试后重置
    const { resetTracer } = require("../agent/tracing");
    resetTracer();
  });

  it("getTracer 应该返回单例", () => {
    const t1 = getTracer();
    const t2 = getTracer();
    expect(t1).toBe(t2);
  });
});
