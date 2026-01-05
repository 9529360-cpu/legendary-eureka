/**
 * TraceContext 测试
 *
 * 覆盖：
 * - 链路追踪创建与管理
 * - Span 层级结构
 * - 性能统计
 * - 导出功能
 */

import { TraceContext, Trace, Span, SpanType } from "../core/TraceContext";

describe("TraceContext", () => {
  beforeEach(() => {
    TraceContext.reset();
  });

  describe("Trace 管理", () => {
    it("应该启动新的追踪", () => {
      const trace = TraceContext.startTrace("test-operation");

      expect(trace).toBeDefined();
      expect(trace.name).toBe("test-operation");
      expect(trace.traceId).toBeDefined();
      expect(trace.status).toBe("running");
      expect(trace.spans.length).toBe(0);
    });

    it("应该获取当前追踪", () => {
      TraceContext.startTrace("current-test");

      const current = TraceContext.getCurrentTrace();

      expect(current).toBeDefined();
      expect(current!.name).toBe("current-test");
    });

    it("应该结束追踪", () => {
      const trace = TraceContext.startTrace("to-end");

      TraceContext.endTrace();

      const ended = TraceContext.getTrace(trace.traceId);
      expect(ended!.status).toBe("completed");
      expect(ended!.endTime).toBeDefined();
      expect(ended!.duration).toBeGreaterThanOrEqual(0);
    });

    it("应该标记追踪失败", () => {
      const trace = TraceContext.startTrace("to-fail");

      TraceContext.failTrace(new Error("Test error"));

      const failed = TraceContext.getTrace(trace.traceId);
      expect(failed!.status).toBe("error");
      expect(failed!.error).toBe("Test error");
    });

    it("应该保留多个追踪历史", () => {
      TraceContext.startTrace("trace-1");
      TraceContext.endTrace();

      TraceContext.startTrace("trace-2");
      TraceContext.endTrace();

      const all = TraceContext.getAllTraces();
      expect(all.length).toBe(2);
    });
  });

  describe("Span 管理", () => {
    beforeEach(() => {
      TraceContext.startTrace("span-test");
    });

    afterEach(() => {
      TraceContext.endTrace();
    });

    it("应该创建根 Span", () => {
      const span = TraceContext.startSpan("root-span", SpanType.TOOL);

      expect(span).toBeDefined();
      expect(span.name).toBe("root-span");
      expect(span.type).toBe(SpanType.TOOL);
      expect(span.parentSpanId).toBeNull();
    });

    it("应该创建嵌套 Span", () => {
      const parent = TraceContext.startSpan("parent", SpanType.TOOL);
      const child = TraceContext.startSpan("child", SpanType.EXCEL);

      expect(child.parentSpanId).toBe(parent.spanId);
    });

    it("应该结束 Span 并回到父级", () => {
      TraceContext.startSpan("parent", SpanType.TOOL);
      TraceContext.startSpan("child", SpanType.EXCEL);

      TraceContext.endSpan();

      const current = TraceContext.getCurrentSpan();
      expect(current!.name).toBe("parent");
    });

    it("应该在 Span 上设置属性", () => {
      const span = TraceContext.startSpan("with-attrs", SpanType.HTTP);

      TraceContext.setSpanAttribute("url", "https://api.example.com");
      TraceContext.setSpanAttribute("method", "POST");

      expect(span.attributes["url"]).toBe("https://api.example.com");
      expect(span.attributes["method"]).toBe("POST");
    });

    it("应该记录 Span 事件", () => {
      TraceContext.startSpan("with-events", SpanType.AI);

      TraceContext.addSpanEvent("Processing started", { step: 1 });
      TraceContext.addSpanEvent("Processing completed", { step: 2 });

      const span = TraceContext.getCurrentSpan();
      expect(span!.events.length).toBe(2);
    });

    it("应该正确计算 Span 持续时间", async () => {
      TraceContext.startSpan("timed-span", SpanType.TOOL);

      // 模拟一些工作
      await new Promise((resolve) => setTimeout(resolve, 50));

      TraceContext.endSpan();

      const trace = TraceContext.getCurrentTrace();
      const span = trace!.spans[0];
      expect(span.duration).toBeGreaterThanOrEqual(45);
    });
  });

  describe("Span 类型", () => {
    beforeEach(() => {
      TraceContext.startTrace("type-test");
    });

    afterEach(() => {
      TraceContext.endTrace();
    });

    it("应该支持 HTTP 类型", () => {
      const span = TraceContext.startSpan("http-call", SpanType.HTTP);
      expect(span.type).toBe(SpanType.HTTP);
    });

    it("应该支持 EXCEL 类型", () => {
      const span = TraceContext.startSpan("excel-op", SpanType.EXCEL);
      expect(span.type).toBe(SpanType.EXCEL);
    });

    it("应该支持 AI 类型", () => {
      const span = TraceContext.startSpan("ai-call", SpanType.AI);
      expect(span.type).toBe(SpanType.AI);
    });

    it("应该支持 TOOL 类型", () => {
      const span = TraceContext.startSpan("tool-exec", SpanType.TOOL);
      expect(span.type).toBe(SpanType.TOOL);
    });

    it("应该支持 USER 类型", () => {
      const span = TraceContext.startSpan("user-action", SpanType.USER);
      expect(span.type).toBe(SpanType.USER);
    });
  });

  describe("错误追踪", () => {
    beforeEach(() => {
      TraceContext.startTrace("error-test");
    });

    afterEach(() => {
      if (TraceContext.getCurrentTrace()) {
        TraceContext.endTrace();
      }
    });

    it("应该在 Span 上记录错误", () => {
      TraceContext.startSpan("error-span", SpanType.TOOL);

      TraceContext.setSpanError(new Error("Test error"));

      const span = TraceContext.getCurrentSpan();
      expect(span!.status).toBe("error");
      expect(span!.error).toBe("Test error");
    });

    it("应该记录错误事件", () => {
      TraceContext.startSpan("error-event-span", SpanType.HTTP);

      TraceContext.setSpanError(new Error("Connection failed"));

      const span = TraceContext.getCurrentSpan();
      const errorEvent = span!.events.find((e) => e.name === "error");
      expect(errorEvent).toBeDefined();
    });
  });

  describe("性能统计", () => {
    beforeEach(() => {
      // 创建多个追踪用于统计
      for (let i = 0; i < 5; i++) {
        const trace = TraceContext.startTrace(`perf-trace-${i}`);

        TraceContext.startSpan("http-span", SpanType.HTTP);
        TraceContext.endSpan();

        TraceContext.startSpan("tool-span", SpanType.TOOL);
        TraceContext.startSpan("excel-span", SpanType.EXCEL);
        TraceContext.endSpan();
        TraceContext.endSpan();

        if (i % 2 === 0) {
          TraceContext.endTrace();
        } else {
          TraceContext.failTrace(new Error("Simulated failure"));
        }
      }
    });

    it("应该计算成功率", () => {
      const stats = TraceContext.getStatistics();

      expect(stats.totalTraces).toBe(5);
      expect(stats.successRate).toBeCloseTo(0.6, 1);
    });

    it("应该统计 Span 类型分布", () => {
      const stats = TraceContext.getStatistics();

      expect(stats.spansByType[SpanType.HTTP]).toBe(5);
      expect(stats.spansByType[SpanType.TOOL]).toBe(5);
      expect(stats.spansByType[SpanType.EXCEL]).toBe(5);
    });

    it("应该计算平均持续时间", () => {
      const stats = TraceContext.getStatistics();

      expect(stats.avgDuration).toBeGreaterThanOrEqual(0);
    });
  });

  describe("追踪导出", () => {
    beforeEach(() => {
      TraceContext.startTrace("export-test");
      TraceContext.startSpan("span-1", SpanType.TOOL);
      TraceContext.startSpan("span-1-1", SpanType.EXCEL);
      TraceContext.endSpan();
      TraceContext.endSpan();
      TraceContext.startSpan("span-2", SpanType.HTTP);
      TraceContext.endSpan();
      TraceContext.endTrace();
    });

    it("应该导出为 JSON", () => {
      const traces = TraceContext.getAllTraces();
      const json = TraceContext.exportToJson(traces[0].traceId);

      expect(json).toBeDefined();
      const parsed = JSON.parse(json);
      expect(parsed.name).toBe("export-test");
      expect(parsed.spans.length).toBe(3);
    });

    it("应该导出为树形结构", () => {
      const traces = TraceContext.getAllTraces();
      const tree = TraceContext.exportToTree(traces[0].traceId);

      expect(tree).toBeDefined();
      expect(tree.name).toBe("export-test");
      expect(tree.children.length).toBe(2);
    });

    it("应该导出时间线数据", () => {
      const traces = TraceContext.getAllTraces();
      const timeline = TraceContext.exportToTimeline(traces[0].traceId);

      expect(timeline).toBeDefined();
      expect(timeline.events.length).toBeGreaterThan(0);
    });
  });

  describe("追踪上下文传递", () => {
    it("应该支持跨函数追踪", async () => {
      const trace = TraceContext.startTrace("cross-function");

      await simulateAsyncOperation();

      TraceContext.endTrace();

      const completed = TraceContext.getTrace(trace.traceId);
      expect(completed!.spans.length).toBeGreaterThan(0);
    });

    it("应该正确维护 Span 栈", () => {
      TraceContext.startTrace("stack-test");

      TraceContext.startSpan("level-1", SpanType.TOOL);
      TraceContext.startSpan("level-2", SpanType.EXCEL);
      TraceContext.startSpan("level-3", SpanType.HTTP);

      // 验证层级
      const level3 = TraceContext.getCurrentSpan();
      expect(level3!.name).toBe("level-3");

      TraceContext.endSpan();
      expect(TraceContext.getCurrentSpan()!.name).toBe("level-2");

      TraceContext.endSpan();
      expect(TraceContext.getCurrentSpan()!.name).toBe("level-1");

      TraceContext.endSpan();
      expect(TraceContext.getCurrentSpan()).toBeNull();

      TraceContext.endTrace();
    });
  });

  describe("追踪清理", () => {
    it("应该按时间清理旧追踪", () => {
      // 创建一些追踪
      for (let i = 0; i < 10; i++) {
        TraceContext.startTrace(`old-trace-${i}`);
        TraceContext.endTrace();
      }

      // 验证追踪存在
      expect(TraceContext.getAllTraces().length).toBe(10);

      // 清理（假设清理0毫秒前的追踪）
      TraceContext.cleanup(0);

      // 所有追踪都应该被清理
      expect(TraceContext.getAllTraces().length).toBe(0);
    });

    it("应该限制追踪数量", () => {
      // 设置最大追踪数
      TraceContext.setMaxTraces(5);

      // 创建超过限制的追踪
      for (let i = 0; i < 10; i++) {
        TraceContext.startTrace(`limited-trace-${i}`);
        TraceContext.endTrace();
      }

      // 应该只保留最新的5个
      expect(TraceContext.getAllTraces().length).toBe(5);
    });
  });
});

// 辅助函数：模拟异步操作
async function simulateAsyncOperation(): Promise<void> {
  TraceContext.startSpan("async-op", SpanType.HTTP);

  await new Promise((resolve) => setTimeout(resolve, 10));

  TraceContext.startSpan("inner-op", SpanType.EXCEL);
  await new Promise((resolve) => setTimeout(resolve, 5));
  TraceContext.endSpan();

  TraceContext.endSpan();
}
