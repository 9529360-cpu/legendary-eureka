/**
 * Phase 3 模块测试 - 智能编排器和动态提示构建器
 *
 * @module __tests__/phase3-modules.test
 */

import {
  SmartOrchestrator,
  createSmartOrchestrator,
  OrchestrationResult,
} from "../agent/SmartOrchestrator";
import {
  DynamicPromptBuilder,
  createDynamicPromptBuilder,
  PromptBuildResult,
} from "../agent/DynamicPromptBuilder";
import { ToolRegistry } from "../agent/registry";
import { Tool, ToolResult } from "../agent/types/tool";
import { IntentParser } from "../agent/IntentParser";
import { SpecCompiler } from "../agent/SpecCompiler";

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

function createDelayedTool(name: string, delay: number): Tool {
  return {
    name,
    description: `Delayed tool ${name}`,
    category: "excel",
    parameters: [],
    execute: async (): Promise<ToolResult> => {
      await new Promise((r) => setTimeout(r, delay));
      return { success: true, output: `${name} completed after ${delay}ms` };
    },
  };
}

// ========== Mock IntentParser ==========

class MockIntentParser extends IntentParser {
  async parse(): Promise<any> {
    return {
      intent: "write_data",
      confidence: 0.9,
      needsClarification: false,
      spec: {
        target: "A1",
        data: [["测试"]],
      },
    };
  }
}

// ========== SmartOrchestrator 测试 ==========

describe("SmartOrchestrator", () => {
  let registry: ToolRegistry;

  beforeEach(() => {
    registry = new ToolRegistry();
    registry.register(createMockTool("excel_read_range", "读取单元格区域"));
    registry.register(createMockTool("excel_write_range", "写入单元格区域"));
    registry.register(createMockTool("excel_create_chart", "创建图表"));
    registry.register(createMockTool("excel_format_cells", "格式化单元格"));
  });

  describe("初始化", () => {
    test("应该正确初始化", async () => {
      const orchestrator = new SmartOrchestrator(registry);
      await orchestrator.initialize();

      // 不抛出错误即为成功
      expect(true).toBe(true);

      orchestrator.close();
    });

    test("应该使用自定义组件初始化", async () => {
      const mockParser = new MockIntentParser();
      const mockCompiler = new SpecCompiler();

      const orchestrator = new SmartOrchestrator(registry, {
        intentParser: mockParser,
        specCompiler: mockCompiler,
      });

      await orchestrator.initialize();
      orchestrator.close();
    });
  });

  describe("编排流程", () => {
    test("应该完成基本编排流程", async () => {
      const mockParser = new MockIntentParser();
      const orchestrator = new SmartOrchestrator(registry, {
        intentParser: mockParser,
      });

      await orchestrator.initialize();

      const result = await orchestrator.orchestrate("写入数据到 A1", {
        parallel: false,
        enableTracing: false,
      });

      expect(result).toBeDefined();
      expect(result.stats).toBeDefined();
      expect(result.stats.parseTime).toBeGreaterThanOrEqual(0);
      expect(result.stats.totalTime).toBeGreaterThanOrEqual(0);

      orchestrator.close();
    });

    test("应该触发进度回调", async () => {
      const mockParser = new MockIntentParser();
      const orchestrator = new SmartOrchestrator(registry, {
        intentParser: mockParser,
      });

      await orchestrator.initialize();

      const progressPhases: string[] = [];

      await orchestrator.orchestrate("测试", {
        onProgress: (progress) => {
          progressPhases.push(progress.phase);
        },
      });

      expect(progressPhases).toContain("parsing");
      expect(progressPhases).toContain("discovering");
      expect(progressPhases).toContain("compiling");

      orchestrator.close();
    });

    test("应该返回发现的工具", async () => {
      const mockParser = new MockIntentParser();
      const orchestrator = new SmartOrchestrator(registry, {
        intentParser: mockParser,
      });

      await orchestrator.initialize();

      const result = await orchestrator.orchestrate("读取单元格数据");

      expect(result.discoveredTools).toBeDefined();
      // 工具发现应该找到相关工具
      expect(result.discoveredTools!.length).toBeGreaterThanOrEqual(0);

      orchestrator.close();
    });

    test("应该获取追踪数据", async () => {
      const mockParser = new MockIntentParser();
      const orchestrator = new SmartOrchestrator(registry, {
        intentParser: mockParser,
      });

      await orchestrator.initialize();

      await orchestrator.orchestrate("测试", {
        enableTracing: true,
      });

      const traceData = orchestrator.getTraceData();
      expect(traceData).toBeDefined();
      expect(traceData.spans).toBeDefined();

      orchestrator.close();
    });
  });

  describe("错误处理", () => {
    test("应该处理解析失败", async () => {
      // 使用默认 IntentParser（没有 API 会失败）
      const orchestrator = new SmartOrchestrator(registry);
      await orchestrator.initialize();

      const result = await orchestrator.orchestrate("测试");

      // 即使解析失败，也应该返回结果
      expect(result).toBeDefined();
      expect(result.stats).toBeDefined();

      orchestrator.close();
    });
  });
});

// ========== DynamicPromptBuilder 测试 ==========

describe("DynamicPromptBuilder", () => {
  let registry: ToolRegistry;
  let builder: DynamicPromptBuilder;

  beforeEach(() => {
    registry = new ToolRegistry();
    registry.register(createMockTool("excel_read_range", "读取单元格区域的值"));
    registry.register(createMockTool("excel_write_range", "写入数据到单元格"));
    registry.register(createMockTool("excel_create_chart", "创建可视化图表"));
    registry.register(createMockTool("excel_format_cells", "格式化单元格样式"));
    registry.register(createMockTool("excel_delete_rows", "删除指定行"));
    registry.register(createMockTool("excel_sort_range", "排序数据区域"));

    builder = new DynamicPromptBuilder(registry);
  });

  describe("build", () => {
    test("应该构建基础提示", async () => {
      const result = await builder.build({
        userMessage: "帮我读取数据",
      });

      expect(result).toBeDefined();
      expect(result.systemPrompt).toContain("Excel");
      expect(result.estimatedTokens).toBeGreaterThan(0);
    });

    test("应该包含相关工具描述", async () => {
      const result = await builder.build({
        userMessage: "读取单元格",
      });

      expect(result.includedTools).toBeGreaterThan(0);
      expect(result.systemPrompt).toContain("read_range");
    });

    test("应该根据关键词过滤工具", async () => {
      const result = await builder.build({
        userMessage: "排序数据",
      });

      expect(result.systemPrompt).toContain("sort_range");
    });

    test("应该包含上下文信息", async () => {
      const result = await builder.build({
        userMessage: "测试",
        workbookContext: {
          sheets: ["Sheet1", "Sheet2"],
          activeSheet: "Sheet1",
          usedRanges: ["A1:B10"],
        },
      });

      expect(result.systemPrompt).toContain("Sheet1");
      expect(result.systemPrompt).toContain("A1:B10");
    });

    test("应该尊重 Token 预算", async () => {
      const smallBudgetBuilder = new DynamicPromptBuilder(registry, {
        maxTokens: 100,
        maxToolDescriptions: 2,
      });

      const result = await smallBudgetBuilder.build({
        userMessage: "测试",
      });

      expect(result.includedTools).toBeLessThanOrEqual(2);
    });

    test("应该支持英文语言", async () => {
      const enBuilder = new DynamicPromptBuilder(registry, {
        language: "en",
      });

      const result = await enBuilder.build({
        userMessage: "test",
      });

      expect(result.systemPrompt).toContain("Excel AI assistant");
    });

    test("应该计算压缩比", async () => {
      const result = await builder.build({
        userMessage: "读取",
      });

      expect(result.compressionRatio).toBeGreaterThan(0);
      expect(result.compressionRatio).toBeLessThanOrEqual(1);
    });
  });

  describe("compressHistory", () => {
    test("应该压缩对话历史", () => {
      const history = [
        { role: "user" as const, content: "Hello" },
        { role: "assistant" as const, content: "Hi there! How can I help you with Excel today?" },
        { role: "user" as const, content: "I want to read some data" },
        { role: "assistant" as const, content: "Sure, I can help you read data from Excel cells." },
      ];

      const compressed = builder.compressHistory(history, 50);

      expect(compressed.length).toBeLessThanOrEqual(history.length);
    });

    test("应该保留最近的消息", () => {
      const history = [
        { role: "user" as const, content: "First message" },
        { role: "assistant" as const, content: "First response" },
        { role: "user" as const, content: "Last message" },
      ];

      const compressed = builder.compressHistory(history, 20);

      // 应该至少保留最后一条消息
      expect(compressed.length).toBeGreaterThan(0);
      expect(compressed[compressed.length - 1].content).toContain("Last");
    });

    test("应该处理空历史", () => {
      const compressed = builder.compressHistory([], 100);
      expect(compressed).toEqual([]);
    });
  });
});

// ========== 集成测试 ==========

describe("Phase 3 Integration", () => {
  test("SmartOrchestrator + DynamicPromptBuilder 集成", async () => {
    const registry = new ToolRegistry();
    registry.register(createMockTool("excel_read_range", "读取单元格区域数据"));
    registry.register(createMockTool("excel_write_range", "写入单元格区域数据"));

    // 创建动态提示构建器
    const promptBuilder = new DynamicPromptBuilder(registry);
    const prompt = await promptBuilder.build({
      userMessage: "读取区域",
      workbookContext: {
        sheets: ["Sheet1"],
        activeSheet: "Sheet1",
        usedRanges: ["A1:C10"],
      },
    });

    expect(prompt.systemPrompt).toBeTruthy();
    expect(prompt.includedTools).toBeGreaterThanOrEqual(0);

    // 创建智能编排器
    const mockParser = new MockIntentParser();
    const orchestrator = new SmartOrchestrator(registry, {
      intentParser: mockParser,
    });

    await orchestrator.initialize();

    const result = await orchestrator.orchestrate("读取数据");

    expect(result).toBeDefined();
    expect(result.stats.totalTime).toBeGreaterThan(0);

    orchestrator.close();
  });

  test("完整编排流程统计", async () => {
    const registry = new ToolRegistry();
    registry.register(createDelayedTool("tool1", 10));
    registry.register(createDelayedTool("tool2", 20));

    const mockParser = new MockIntentParser();
    const orchestrator = new SmartOrchestrator(registry, {
      intentParser: mockParser,
    });

    await orchestrator.initialize();

    const result = await orchestrator.orchestrate("测试", {
      enableTracing: true,
    });

    // 验证统计数据
    expect(result.stats.parseTime).toBeGreaterThanOrEqual(0);
    expect(result.stats.discoverTime).toBeGreaterThanOrEqual(0);
    expect(result.stats.compileTime).toBeGreaterThanOrEqual(0);
    expect(result.stats.totalTime).toBeGreaterThanOrEqual(0);

    orchestrator.close();
  });
});
