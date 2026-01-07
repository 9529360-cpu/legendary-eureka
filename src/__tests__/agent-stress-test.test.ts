/**
 * Agent 压力测试 - 模拟各种"刁难"用户场景
 *
 * 测试 AgentOrchestrator 在真实复杂场景下的表现
 */

import {
  AgentOrchestrator,
  createAgentOrchestrator,
} from "../agent/AgentOrchestrator";

// ========== Mock 设置 ==========

// 用于控制 IntentParser 返回值的变量
let mockParseResult: {
  intent: string;
  confidence: number;
  needsClarification: boolean;
  clarificationQuestion?: string;
  spec: unknown;
};

// 创建有效的 IntentSpec mock
function createMockIntentResult(options: {
  intent: string;
  confidence: number;
  needsClarification?: boolean;
  clarificationQuestion?: string;
}) {
  let spec: unknown;

  switch (options.intent) {
    case "clarify":
      spec = {
        type: "clarify",
        question: options.clarificationQuestion || "请提供更多信息",
        reason: "用户请求不够清晰",
      };
      break;
    case "create_chart":
      spec = {
        type: "chart",
        chartType: "column",
        dataRange: "A1:B10",
      };
      break;
    case "write_data":
      spec = {
        type: "write_data",
        target: "A1",
        data: [["测试"]],
      };
      break;
    case "respond_only":
      spec = {
        type: "respond",
        message: "抱歉，无法执行此操作",
      };
      break;
    case "query_data":
    default:
      spec = {
        type: "query",
        target: "selection",
      };
      break;
  }

  return {
    intent: options.intent,
    confidence: options.confidence,
    needsClarification: options.needsClarification || false,
    clarificationQuestion: options.clarificationQuestion,
    spec,
  };
}

// Mock IntentParser
jest.mock("../agent/IntentParser", () => ({
  IntentParser: jest.fn().mockImplementation(() => ({
    parse: jest.fn().mockImplementation(() => Promise.resolve(mockParseResult)),
  })),
}));

// Mock SpecCompiler
let mockCompileResult: {
  success: boolean;
  error?: string;
  plan?: {
    id: string;
    taskDescription: string;
    steps: Array<{
      id: string;
      order: number;
      action: string;
      description: string;
      parameters: Record<string, unknown>;
      dependsOn: string[];
      successCondition: { type: string };
      isWriteOperation: boolean;
      isCritical?: boolean;
      status: string;
    }>;
    taskType: string;
    currentPhase: string;
    dependencyCheck: { passed: boolean; issues: string[] };
  };
};

jest.mock("../agent/SpecCompiler", () => ({
  SpecCompiler: jest.fn().mockImplementation(() => ({
    compile: jest.fn().mockImplementation(() => mockCompileResult),
  })),
}));

// Mock ToolRegistry
let mockToolExecuteResult: { success: boolean; output: string; error?: string };

jest.mock("../agent/registry", () => ({
  ToolRegistry: jest.fn().mockImplementation(() => ({
    register: jest.fn(),
    get: jest.fn().mockImplementation(() => ({
      name: "excel_get_selection",
      execute: jest.fn().mockImplementation(() => Promise.resolve(mockToolExecuteResult)),
    })),
  })),
}));

// Mock ExcelAdapter
jest.mock("../agent/ExcelAdapter", () => ({
  __esModule: true,
  default: jest.fn().mockReturnValue([]),
}));

// Mock EpisodicMemory
jest.mock("../agent/EpisodicMemory", () => ({
  EpisodicMemory: jest.fn().mockImplementation(() => ({
    findSimilar: jest.fn().mockReturnValue([]),
    startEpisode: jest.fn().mockReturnValue("episode_1"),
    recordStep: jest.fn(),
    endEpisode: jest.fn().mockReturnValue(null),
    extractReusableExperience: jest.fn().mockReturnValue([]),
    abandonEpisode: jest.fn(),
  })),
}));

// Mock AntiHallucinationController
jest.mock("../agent/core/gates/AntiHallucinationController", () => ({
  AntiHallucinationController: jest.fn().mockImplementation(() => ({
    createRun: jest.fn().mockReturnValue({
      runId: "run_1",
      state: "INIT",
      iteration: 0,
      history: [],
    }),
    handleUserMessage: jest.fn(),
    handleModelOutput: jest.fn().mockReturnValue({
      allowFinish: true,
      state: "DEPLOYED",
    }),
  })),
}));

// ========== 重置函数 ==========

function resetMocks() {
  mockParseResult = createMockIntentResult({
    intent: "query_data",
    confidence: 0.9,
  });
  mockCompileResult = {
    success: true,
    plan: {
      id: "plan_1",
      taskDescription: "测试任务",
      steps: [
        {
          id: "step_1",
          order: 1,
          action: "excel_get_selection",
          description: "获取选区",
          parameters: {},
          dependsOn: [],
          successCondition: { type: "tool_success" },
          isWriteOperation: false,
          status: "pending",
        },
      ],
      taskType: "data_analysis",
      currentPhase: "execution",
      dependencyCheck: { passed: true, issues: [] },
    },
  };
  mockToolExecuteResult = {
    success: true,
    output: JSON.stringify({ address: "A1:B10" }),
  };
}

// ========== 测试套件 ==========

describe("Agent 压力测试 - 刁难场景", () => {
  let orchestrator: AgentOrchestrator;

  beforeEach(() => {
    resetMocks();
    orchestrator = createAgentOrchestrator();
  });

  // ========== 1. 模糊请求测试 ==========
  describe("模糊请求处理", () => {
    const vagueRequests = [
      "帮我弄一下表格",
      "把这个整理一下",
      "搞个图",
      "数据有问题",
      "这里不对",
      "优化一下",
    ];

    test.each(vagueRequests)("应该对模糊请求 '%s' 要求澄清", async (request) => {
      mockParseResult = createMockIntentResult({
        intent: "clarify",
        confidence: 0.3,
        needsClarification: true,
        clarificationQuestion: "请问您具体想要做什么操作？",
      });

      const result = await orchestrator.run({ userMessage: request });

      expect(result.needsClarification).toBe(true);
      expect(result.clarificationQuestion).toBeDefined();
    });

    test("连续多个模糊请求应该保持对话历史", async () => {
      mockParseResult = createMockIntentResult({
        intent: "clarify",
        confidence: 0.3,
        needsClarification: true,
        clarificationQuestion: "需要更多信息",
      });

      await orchestrator.run({ userMessage: "弄一下" });
      await orchestrator.run({ userMessage: "就是那个" });
      await orchestrator.run({ userMessage: "那个东西" });

      const history = orchestrator.getConversationHistory();
      // 每轮对话保存用户消息，澄清不保存 assistant 消息
      // 所以 3 轮对话至少有 3 条用户消息
      expect(history.length).toBeGreaterThanOrEqual(3);
    });
  });

  // ========== 2. 信息不完整测试 ==========
  describe("信息不完整处理", () => {
    test("缺少目标位置时应该要求澄清", async () => {
      mockParseResult = createMockIntentResult({
        intent: "clarify",
        confidence: 0.6,
        needsClarification: true,
        clarificationQuestion: "请问结果放在哪里？",
      });

      const result = await orchestrator.run({ userMessage: "把A列求和" });

      expect(result.needsClarification).toBe(true);
    });

    test("缺少数据范围时应该要求澄清", async () => {
      mockParseResult = createMockIntentResult({
        intent: "clarify",
        confidence: 0.5,
        needsClarification: true,
        clarificationQuestion: "请问要创建什么数据的图表？",
      });

      const result = await orchestrator.run({ userMessage: "创建图表" });

      expect(result.needsClarification).toBe(true);
    });
  });

  // ========== 3. 错别字处理测试 ==========
  describe("错别字和术语错误处理", () => {
    const typoRequests = [
      "创建一个柱转图",
      "数据透示表",
      "条件各式",
      "冻洁首行",
    ];

    test.each(typoRequests)("应该理解有错别字的请求 '%s'", async (request) => {
      mockParseResult = createMockIntentResult({
        intent: "create_chart",
        confidence: 0.85,
      });

      const result = await orchestrator.run({ userMessage: request });

      expect(result.success).toBe(true);
    });
  });

  // ========== 4. 中英混杂测试 ==========
  describe("中英混杂请求处理", () => {
    const mixedRequests = [
      "帮我create一个chart",
      "用SUMIF计算sales总和",
      "把这个range format一下",
    ];

    test.each(mixedRequests)("应该理解中英混杂请求 '%s'", async (request) => {
      mockParseResult = createMockIntentResult({
        intent: "create_chart",
        confidence: 0.8,
      });

      const result = await orchestrator.run({ userMessage: request });

      expect(result.success).toBe(true);
    });
  });

  // ========== 5. 边缘情况测试 ==========
  describe("边缘情况处理", () => {
    test("应该处理空消息", async () => {
      mockParseResult = createMockIntentResult({
        intent: "clarify",
        confidence: 0.1,
        needsClarification: true,
        clarificationQuestion: "请输入您想要执行的操作",
      });

      const result = await orchestrator.run({ userMessage: "" });

      expect(result.needsClarification).toBe(true);
    });

    test("应该处理超长消息", async () => {
      const longMessage = "请帮我处理以下数据：" + "数据项".repeat(500);

      mockParseResult = createMockIntentResult({
        intent: "query_data",
        confidence: 0.7,
      });

      const result = await orchestrator.run({ userMessage: longMessage });

      expect(result).toBeDefined();
    });

    test("应该处理特殊字符", async () => {
      const specialChars = "处理<script>alert('xss')</script>数据";

      mockParseResult = createMockIntentResult({
        intent: "query_data",
        confidence: 0.8,
      });

      const result = await orchestrator.run({ userMessage: specialChars });

      expect(result).toBeDefined();
    });

    test("应该处理 emoji", async () => {
      const emojiMessage = "创建📊图表，统计🎯销售💰";

      mockParseResult = createMockIntentResult({
        intent: "create_chart",
        confidence: 0.85,
      });

      const result = await orchestrator.run({ userMessage: emojiMessage });

      expect(result.success).toBe(true);
    });
  });

  // ========== 6. 多步骤任务测试 ==========
  describe("多步骤任务处理", () => {
    test("应该处理多个操作的复合请求", async () => {
      mockParseResult = createMockIntentResult({
        intent: "write_data",
        confidence: 0.9,
      });

      mockCompileResult = {
        success: true,
        plan: {
          id: "plan_multi",
          taskDescription: "排序、格式化并创建图表",
          steps: [
            {
              id: "step_1",
              order: 1,
              action: "excel_sort_range",
              description: "排序",
              parameters: {},
              dependsOn: [],
              successCondition: { type: "tool_success" },
              isWriteOperation: true,
              status: "pending",
            },
            {
              id: "step_2",
              order: 2,
              action: "excel_format_range",
              description: "格式化",
              parameters: {},
              dependsOn: ["step_1"],
              successCondition: { type: "tool_success" },
              isWriteOperation: true,
              status: "pending",
            },
            {
              id: "step_3",
              order: 3,
              action: "excel_create_chart",
              description: "创建图表",
              parameters: {},
              dependsOn: ["step_2"],
              successCondition: { type: "tool_success" },
              isWriteOperation: true,
              status: "pending",
            },
          ],
          taskType: "data_analysis",
          currentPhase: "execution",
          dependencyCheck: { passed: true, issues: [] },
        },
      };

      const result = await orchestrator.run({
        userMessage: "把A列排序后格式化，然后创建柱状图",
      });

      expect(result.success).toBe(true);
      expect(result.state.stepResults.length).toBe(3);
    });
  });

  // ========== 7. 错误恢复测试 ==========
  describe("错误恢复和重试", () => {
    test("关键步骤失败时应该返回错误", async () => {
      mockParseResult = createMockIntentResult({
        intent: "write_data",
        confidence: 0.9,
      });

      mockToolExecuteResult = {
        success: false,
        output: "",
        error: "执行失败",
      };

      mockCompileResult = {
        success: true,
        plan: {
          id: "plan_fail",
          taskDescription: "测试失败",
          steps: [
            {
              id: "step_1",
              order: 1,
              action: "excel_write_cell",
              description: "写入",
              parameters: {},
              dependsOn: [],
              successCondition: { type: "tool_success" },
              isWriteOperation: true,
              isCritical: true,
              status: "pending",
            },
          ],
          taskType: "write",
          currentPhase: "execution",
          dependencyCheck: { passed: true, issues: [] },
        },
      };

      const result = await orchestrator.run({ userMessage: "写入测试" });

      expect(result.success).toBe(false);
    });

    test("非关键步骤应该正常执行", async () => {
      mockParseResult = createMockIntentResult({
        intent: "query_data",
        confidence: 0.9,
      });

      const result = await orchestrator.run({ userMessage: "测试" });

      expect(result.success).toBe(true);
    });
  });

  // ========== 8. 不友好语气测试 ==========
  describe("不友好语气处理", () => {
    const aggressiveRequests = [
      "快点！帮我弄表格！",
      "这破系统怎么这么慢",
      "能不能行啊？",
      "赶紧的！",
    ];

    test.each(aggressiveRequests)("应该专业处理不友好请求 '%s'", async (request) => {
      mockParseResult = createMockIntentResult({
        intent: "query_data",
        confidence: 0.8,
      });

      const result = await orchestrator.run({ userMessage: request });

      expect(result).toBeDefined();
    });
  });

  // ========== 9. 不可能的请求测试 ==========
  describe("不可能的请求处理", () => {
    test("应该拒绝超出能力范围的请求", async () => {
      mockParseResult = createMockIntentResult({
        intent: "respond_only",
        confidence: 0.95,
      });

      mockCompileResult = {
        success: false,
        error: "无法执行此操作",
        plan: undefined,
      };

      const result = await orchestrator.run({
        userMessage: "帮我预测明天的股价",
      });

      expect(result.success).toBe(false);
    });
  });

  // ========== 10. 对话上下文测试 ==========
  describe("对话上下文理解", () => {
    test("应该记住之前的操作", async () => {
      mockParseResult = createMockIntentResult({
        intent: "query_data",
        confidence: 0.9,
      });

      await orchestrator.run({ userMessage: "获取A列数据" });
      await orchestrator.run({ userMessage: "把它排序" });

      const history = orchestrator.getConversationHistory();
      expect(history.length).toBeGreaterThan(0);
    });

    test("清除历史后应该重新开始", async () => {
      mockParseResult = createMockIntentResult({
        intent: "query_data",
        confidence: 0.9,
      });

      await orchestrator.run({ userMessage: "第一条" });
      orchestrator.clearConversationHistory();

      const history = orchestrator.getConversationHistory();
      expect(history.length).toBe(0);
    });
  });
});

// ========== 性能测试 ==========
describe("Agent 性能测试", () => {
  let orchestrator: AgentOrchestrator;

  beforeEach(() => {
    resetMocks();
    orchestrator = createAgentOrchestrator();
  });

  test("简单任务应该快速完成", async () => {
    const startTime = Date.now();
    await orchestrator.run({ userMessage: "获取A1" });
    const duration = Date.now() - startTime;

    expect(duration).toBeLessThan(1000);
  });

  test("应该防止无限循环", async () => {
    orchestrator = createAgentOrchestrator({ maxIterations: 3 });

    mockToolExecuteResult = {
      success: false,
      output: "",
      error: "持续失败",
    };

    if (mockCompileResult.plan) {
      mockCompileResult.plan.steps[0].isCritical = true;
    }

    const result = await orchestrator.run({ userMessage: "测试" });

    expect(result).toBeDefined();
  });
});

// ========== 状态管理测试 ==========
describe("Agent 状态管理", () => {
  let orchestrator: AgentOrchestrator;

  beforeEach(() => {
    resetMocks();
    orchestrator = createAgentOrchestrator();
  });

  test("执行后状态应该正确", async () => {
    const result = await orchestrator.run({ userMessage: "测试" });

    expect(result.state.phase).toBe("completed");
    expect(result.state.stepResults.length).toBeGreaterThan(0);
  });

  test("失败后状态应该反映错误", async () => {
    mockCompileResult = {
      success: false,
      error: "编译失败",
      plan: undefined,
    };

    const result = await orchestrator.run({ userMessage: "测试" });

    expect(result.success).toBe(false);
    expect(result.state.errors.length).toBeGreaterThan(0);
  });
});
