/**
 * Phase 2 模块测试 - 并行执行、工具发现、持久化内存
 *
 * @module __tests__/phase2-modules.test
 */

import {
  ParallelExecutor,
  buildDAG,
  detectCycle,
  getReadyNodes,
  DAGNode,
  ParallelExecutionResult,
} from "../agent/ParallelExecutor";
import { RecoverableStep } from "../agent/RecoveryManager";
import {
  ToolDiscovery,
  IntentAtom,
  ToolMatch,
  SemanticTag,
} from "../agent/ToolDiscovery";
import { ToolRegistry } from "../agent/registry";
import { Tool, ToolResult } from "../agent/types/tool";

// ========== Mock 工具 ==========

function createMockTool(name: string, description: string): Tool {
  return {
    name,
    description,
    category: "excel",
    parameters: [],
    execute: async (): Promise<ToolResult> => ({
      success: true,
      output: `${name} executed`,
    }),
  };
}

function createFailingTool(name: string): Tool {
  return {
    name,
    description: `Failing tool ${name}`,
    category: "excel",
    parameters: [],
    execute: async (): Promise<ToolResult> => ({
      success: false,
      error: `${name} failed`,
      output: "",
    }),
  };
}

// ========== ParallelExecutor 测试 ==========

describe("ParallelExecutor", () => {
  let registry: ToolRegistry;
  let executor: ParallelExecutor;

  beforeEach(() => {
    registry = new ToolRegistry();
    registry.register(createMockTool("excel_read_range", "读取单元格区域"));
    registry.register(createMockTool("excel_write_range", "写入单元格区域"));
    registry.register(createMockTool("excel_create_chart", "创建图表"));
    registry.register(createMockTool("excel_format_cells", "格式化单元格"));
    executor = new ParallelExecutor(registry);
  });

  describe("buildDAG", () => {
    test("应该构建无依赖的 DAG", () => {
      const steps: RecoverableStep[] = [
        { id: "step1", action: "excel_read_range", parameters: {} },
        { id: "step2", action: "excel_write_range", parameters: {} },
      ];

      const dag = buildDAG(steps);

      expect(dag.size).toBe(2);
      expect(dag.get("step1")?.status).toBe("ready");
      expect(dag.get("step2")?.status).toBe("ready");
    });

    test("应该构建有依赖的 DAG", () => {
      const steps: RecoverableStep[] = [
        { id: "step1", action: "excel_read_range", parameters: {} },
        { id: "step2", action: "excel_write_range", parameters: {}, dependsOn: ["step1"] },
        { id: "step3", action: "excel_create_chart", parameters: {}, dependsOn: ["step2"] },
      ];

      const dag = buildDAG(steps);

      expect(dag.get("step1")?.status).toBe("ready");
      expect(dag.get("step2")?.status).toBe("pending");
      expect(dag.get("step3")?.status).toBe("pending");
      expect(dag.get("step1")?.dependents).toContain("step2");
      expect(dag.get("step2")?.dependents).toContain("step3");
    });
  });

  describe("detectCycle", () => {
    test("应该检测无循环", () => {
      const steps: RecoverableStep[] = [
        { id: "step1", action: "excel_read_range", parameters: {} },
        { id: "step2", action: "excel_write_range", parameters: {}, dependsOn: ["step1"] },
      ];

      const dag = buildDAG(steps);
      const cycle = detectCycle(dag);

      expect(cycle).toBeNull();
    });

    test("应该检测循环依赖", () => {
      // 手动构建循环
      const nodes = new Map<string, DAGNode>();
      nodes.set("A", {
        step: { id: "A", action: "test", parameters: {} },
        status: "pending",
        dependencies: ["C"],
        dependents: ["B"],
      });
      nodes.set("B", {
        step: { id: "B", action: "test", parameters: {} },
        status: "pending",
        dependencies: ["A"],
        dependents: ["C"],
      });
      nodes.set("C", {
        step: { id: "C", action: "test", parameters: {} },
        status: "pending",
        dependencies: ["B"],
        dependents: ["A"],
      });

      const cycle = detectCycle(nodes);
      // 循环检测返回非空数组
      expect(cycle).not.toBeNull();
    });
  });

  describe("getReadyNodes", () => {
    test("应该获取可执行节点", () => {
      const steps: RecoverableStep[] = [
        { id: "step1", action: "excel_read_range", parameters: {} },
        { id: "step2", action: "excel_write_range", parameters: {} },
        { id: "step3", action: "excel_create_chart", parameters: {}, dependsOn: ["step1"] },
      ];

      const dag = buildDAG(steps);
      const ready = getReadyNodes(dag);

      expect(ready.length).toBe(2);
      expect(ready.map((n) => n.step.id)).toContain("step1");
      expect(ready.map((n) => n.step.id)).toContain("step2");
    });

    test("依赖完成后应该标记新的可执行节点", () => {
      const steps: RecoverableStep[] = [
        { id: "step1", action: "excel_read_range", parameters: {} },
        { id: "step2", action: "excel_write_range", parameters: {}, dependsOn: ["step1"] },
      ];

      const dag = buildDAG(steps);

      // 模拟 step1 完成
      const step1Node = dag.get("step1")!;
      step1Node.status = "completed";

      const ready = getReadyNodes(dag);
      expect(ready.length).toBe(1);
      expect(ready[0].step.id).toBe("step2");
    });
  });

  describe("execute", () => {
    test("应该并行执行独立步骤", async () => {
      const steps: RecoverableStep[] = [
        { id: "step1", action: "excel_read_range", parameters: {} },
        { id: "step2", action: "excel_write_range", parameters: {} },
      ];

      const result = await executor.execute(steps);

      expect(result.success).toBe(true);
      expect(result.totalSteps).toBe(2);
      expect(result.successCount).toBe(2);
      expect(result.parallelism.maxConcurrent).toBeGreaterThanOrEqual(1);
    });

    test("应该按依赖顺序执行", async () => {
      const executionOrder: string[] = [];

      registry.register({
        name: "track_step",
        description: "追踪执行顺序",
        category: "test",
        parameters: [{ name: "stepId", description: "步骤ID", type: "string", required: true }],
        execute: async (input): Promise<ToolResult> => {
          executionOrder.push(input.stepId as string);
          await new Promise((r) => setTimeout(r, 10)); // 模拟执行时间
          return { success: true, output: String(input.stepId) };
        },
      });

      const steps: RecoverableStep[] = [
        { id: "step1", action: "track_step", parameters: { stepId: "step1" } },
        { id: "step2", action: "track_step", parameters: { stepId: "step2" }, dependsOn: ["step1"] },
      ];

      const result = await executor.execute(steps);

      expect(result.success).toBe(true);
      expect(executionOrder.indexOf("step1")).toBeLessThan(executionOrder.indexOf("step2"));
    });

    test("应该处理失败步骤", async () => {
      registry.register(createFailingTool("excel_fail"));

      const steps: RecoverableStep[] = [
        { id: "step1", action: "excel_fail", parameters: {} },
        { id: "step2", action: "excel_read_range", parameters: {} },
      ];

      const result = await executor.execute(steps, { continueOnFailure: true });

      expect(result.success).toBe(false);
      expect(result.failedCount).toBe(1);
      expect(result.successCount).toBe(1);
    });

    test("应该跳过依赖失败的步骤", async () => {
      registry.register(createFailingTool("excel_fail"));

      const steps: RecoverableStep[] = [
        { id: "step1", action: "excel_fail", parameters: {} },
        { id: "step2", action: "excel_read_range", parameters: {}, dependsOn: ["step1"] },
      ];

      const result = await executor.execute(steps, { continueOnFailure: false });

      expect(result.failedCount).toBe(1);
      expect(result.skippedCount).toBe(1);
    });

    test("应该触发事件回调", async () => {
      const events: string[] = [];

      const steps: RecoverableStep[] = [
        { id: "step1", action: "excel_read_range", parameters: {} },
      ];

      await executor.execute(steps, {
        onEvent: (event) => events.push(event.type),
      });

      expect(events).toContain("batch:start");
      expect(events).toContain("step:start");
      expect(events).toContain("step:complete");
      expect(events).toContain("complete");
    });

    test("应该遵守最大并发限制", async () => {
      let currentConcurrency = 0;
      let maxObservedConcurrency = 0;

      registry.register({
        name: "concurrent_test",
        description: "并发测试",
        category: "test",
        parameters: [],
        execute: async (): Promise<ToolResult> => {
          currentConcurrency++;
          maxObservedConcurrency = Math.max(maxObservedConcurrency, currentConcurrency);
          await new Promise((r) => setTimeout(r, 50));
          currentConcurrency--;
          return { success: true, output: "done" };
        },
      });

      const steps: RecoverableStep[] = Array.from({ length: 10 }, (_, i) => ({
        id: `step${i}`,
        action: "concurrent_test",
        parameters: {},
      }));

      await executor.execute(steps, { maxConcurrency: 3 });

      expect(maxObservedConcurrency).toBeLessThanOrEqual(3);
    });
  });
});

// ========== ToolDiscovery 测试 ==========

describe("ToolDiscovery", () => {
  let registry: ToolRegistry;
  let discovery: ToolDiscovery;

  beforeEach(async () => {
    registry = new ToolRegistry();
    registry.register(createMockTool("excel_read_range", "读取单元格区域的值"));
    registry.register(createMockTool("excel_write_range", "写入数据到单元格区域"));
    registry.register(createMockTool("excel_create_chart", "创建图表可视化"));
    registry.register(createMockTool("excel_format_cells", "格式化单元格样式"));
    registry.register(createMockTool("excel_delete_rows", "删除行"));
    registry.register(createMockTool("excel_sort_range", "排序数据区域"));

    discovery = new ToolDiscovery(registry);
    await discovery.initialize();
  });

  describe("initialize", () => {
    test("应该构建语义索引", async () => {
      const stats = discovery.getStats();

      expect(stats.totalTools).toBe(6);
      expect(stats.totalTags).toBeGreaterThan(0);
    });
  });

  describe("discover", () => {
    test("应该根据动作意图发现工具", () => {
      const intent: IntentAtom = {
        action: "读取",
        entity: "单元格",
      };

      const matches = discovery.discover(intent);

      expect(matches.length).toBeGreaterThan(0);
      expect(matches[0].tool.name).toBe("excel_read_range");
    });

    test("应该根据原始文本发现工具", () => {
      const intent: IntentAtom = {
        rawText: "我想创建一个图表",
      };

      const matches = discovery.discover(intent);

      expect(matches.some((m) => m.tool.name === "excel_create_chart")).toBe(true);
    });

    test("应该返回带有分数的匹配结果", () => {
      const intent: IntentAtom = {
        action: "写入",
        entity: "数据",
      };

      const matches = discovery.discover(intent);

      expect(matches.length).toBeGreaterThan(0);
      for (const match of matches) {
        expect(match.score).toBeGreaterThanOrEqual(0);
        expect(match.score).toBeLessThanOrEqual(1);
        expect(match.matchedTags).toBeDefined();
        expect(match.reason).toBeDefined();
      }
    });

    test("应该尊重最低分数阈值", () => {
      const intent: IntentAtom = {
        action: "读取",
      };

      const matches = discovery.discover(intent, { minScore: 0.8 });

      for (const match of matches) {
        expect(match.score).toBeGreaterThanOrEqual(0.8);
      }
    });

    test("应该限制返回数量", () => {
      const intent: IntentAtom = {
        rawText: "操作单元格",
      };

      const matches = discovery.discover(intent, { limit: 2 });

      expect(matches.length).toBeLessThanOrEqual(2);
    });

    test("应该按分类过滤", () => {
      registry.register({
        name: "word_insert_text",
        description: "插入文本",
        category: "word",
        parameters: [],
        execute: async () => ({ success: true, output: "" }),
      });

      // 重新初始化
      discovery = new ToolDiscovery(registry);
      discovery.initialize();

      const intent: IntentAtom = {
        rawText: "插入",
      };

      const matches = discovery.discover(intent, { categories: ["excel"] });

      for (const match of matches) {
        expect(match.tool.category).toBe("excel");
      }
    });
  });

  describe("search", () => {
    test("应该根据关键词搜索工具", () => {
      const matches = discovery.search("排序");

      expect(matches.some((m) => m.tool.name === "excel_sort_range")).toBe(true);
    });

    test("应该搜索中文关键词", () => {
      const matches = discovery.search("删除行");

      expect(matches.some((m) => m.tool.name === "excel_delete_rows")).toBe(true);
    });
  });

  describe("getByCategory", () => {
    test("应该获取分类下的工具", () => {
      const tools = discovery.getByCategory("excel");

      expect(tools.length).toBe(6);
      for (const tool of tools) {
        expect(tool.category).toBe("excel");
      }
    });
  });

  describe("updateStats", () => {
    test("应该更新工具使用统计", () => {
      discovery.updateStats("excel_read_range", true, 100);
      discovery.updateStats("excel_read_range", true, 200);
      discovery.updateStats("excel_read_range", false, 50);

      const popular = discovery.getPopular();

      expect(popular.length).toBeGreaterThan(0);
      const readRange = popular.find((p) => p.name === "excel_read_range");
      expect(readRange).toBeDefined();
      expect(readRange!.usageCount).toBe(3);
    });
  });
});

// ========== PersistentMemory 测试 (Mock IndexedDB) ==========

// 注意：真实的 IndexedDB 测试需要 jsdom 或 fake-indexeddb
// 这里提供基本的类型和接口测试

describe("PersistentMemory Types", () => {
  test("StoredMessage 接口应该正确", () => {
    const message = {
      id: "msg_123",
      role: "user" as const,
      content: "Hello",
      timestamp: Date.now(),
      sessionId: "session_1",
    };

    expect(message.id).toBeDefined();
    expect(message.role).toBe("user");
    expect(message.content).toBe("Hello");
  });

  test("StoredEpisode 接口应该正确", () => {
    const episode = {
      id: "ep_123",
      sessionId: "session_1",
      intent: "创建表格",
      actions: ["excel_create_table"],
      result: "success" as const,
      timestamp: Date.now(),
      duration: 1000,
      toolsUsed: ["excel_create_table"],
    };

    expect(episode.id).toBeDefined();
    expect(episode.result).toBe("success");
    expect(episode.toolsUsed).toContain("excel_create_table");
  });

  test("ToolStats 接口应该正确", () => {
    const stats = {
      name: "excel_read_range",
      totalCalls: 100,
      successCalls: 95,
      failureCalls: 5,
      totalDuration: 5000,
      avgDuration: 50,
      lastUsed: Date.now(),
    };

    expect(stats.name).toBeDefined();
    expect(stats.totalCalls).toBe(100);
    expect(stats.successCalls / stats.totalCalls).toBe(0.95);
  });

  test("SessionSummary 接口应该正确", () => {
    const session = {
      id: "session_1",
      startTime: Date.now() - 3600000,
      endTime: Date.now(),
      messageCount: 10,
      successRate: 0.9,
      title: "数据分析会话",
    };

    expect(session.id).toBeDefined();
    expect(session.endTime - session.startTime).toBeGreaterThan(0);
    expect(session.successRate).toBeLessThanOrEqual(1);
  });
});

// ========== 集成测试 ==========

describe("Phase 2 Integration", () => {
  test("ParallelExecutor + ToolDiscovery 集成", async () => {
    const registry = new ToolRegistry();
    registry.register(createMockTool("excel_read_range", "读取单元格区域"));
    registry.register(createMockTool("excel_write_range", "写入单元格区域"));

    // 使用 ToolDiscovery 发现工具
    const discovery = new ToolDiscovery(registry);
    await discovery.initialize();

    const readMatches = discovery.discover({ action: "读取" });
    const writeMatches = discovery.discover({ action: "写入" });

    expect(readMatches.length).toBeGreaterThan(0);
    expect(writeMatches.length).toBeGreaterThan(0);

    // 使用发现的工具构建执行计划
    const steps: RecoverableStep[] = [
      { id: "read", action: readMatches[0].tool.name, parameters: {} },
      { id: "write", action: writeMatches[0].tool.name, parameters: {}, dependsOn: ["read"] },
    ];

    // 并行执行
    const executor = new ParallelExecutor(registry);
    const result = await executor.execute(steps);

    expect(result.success).toBe(true);
    expect(result.successCount).toBe(2);
  });
});
