/**
 * AgentOrchestrator 单元测试
 *
 * 测试智能闭环控制中心的核心功能：
 * - 配置管理
 * - 对话历史
 * - 状态管理
 * - 修复策略
 */

import {
  AgentOrchestrator,
  createAgentOrchestrator,
  DEFAULT_ORCHESTRATOR_CONFIG,
  OrchestratorConfig,
  AgentPhase,
} from "../agent/AgentOrchestrator";

// ========== Mock 依赖 ==========

// Mock IntentParser
jest.mock("../agent/IntentParser", () => ({
  IntentParser: jest.fn().mockImplementation(() => ({
    parse: jest.fn().mockResolvedValue({
      intent: "query",
      confidence: 0.9,
      spec: { type: "query", target: "selection" },
      needsClarification: false,
    }),
  })),
}));

// Mock SpecCompiler
jest.mock("../agent/SpecCompiler", () => ({
  SpecCompiler: jest.fn().mockImplementation(() => ({
    compile: jest.fn().mockReturnValue({
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
    }),
  })),
}));

// Mock ToolRegistry
jest.mock("../agent/registry", () => ({
  ToolRegistry: jest.fn().mockImplementation(() => ({
    register: jest.fn(),
    get: jest.fn().mockReturnValue({
      name: "excel_get_selection",
      execute: jest.fn().mockResolvedValue({
        success: true,
        output: JSON.stringify({ address: "A1:B10", values: [[1, 2]] }),
      }),
    }),
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
    endEpisode: jest.fn().mockReturnValue({
      id: "episode_1",
      userRequest: "测试请求",
      steps: [],
      outcome: "success",
      startTime: Date.now(),
      endTime: Date.now(),
    }),
    extractReusableExperience: jest.fn().mockReturnValue([]),
    abandonEpisode: jest.fn(),
  })),
}));

// Mock AntiHallucinationController
jest.mock("../agent/core/gates/AntiHallucinationController", () => ({
  AntiHallucinationController: jest.fn().mockImplementation(() => ({
    createRun: jest.fn().mockReturnValue({
      id: "run_1",
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

// ========== 测试套件 ==========

describe("AgentOrchestrator", () => {
  let orchestrator: AgentOrchestrator;

  beforeEach(() => {
    orchestrator = createAgentOrchestrator();
  });

  // ========== 初始化测试 ==========

  describe("初始化", () => {
    test("应该使用默认配置创建实例", () => {
      const config = orchestrator.getConfig();
      expect(config.maxRetries).toBe(DEFAULT_ORCHESTRATOR_CONFIG.maxRetries);
      expect(config.maxIterations).toBe(DEFAULT_ORCHESTRATOR_CONFIG.maxIterations);
      expect(config.enableLearning).toBe(true);
      expect(config.enableAutoFix).toBe(true);
      expect(config.enableAntiHallucination).toBe(true);
      expect(config.enableConversationContext).toBe(true);
    });

    test("应该允许自定义配置", () => {
      const customOrchestrator = createAgentOrchestrator({
        maxRetries: 5,
        enableLearning: false,
      });
      const config = customOrchestrator.getConfig();
      expect(config.maxRetries).toBe(5);
      expect(config.enableLearning).toBe(false);
    });

    test("应该能获取工具注册表", () => {
      const registry = orchestrator.getToolRegistry();
      expect(registry).toBeDefined();
    });

    test("应该能获取经验记忆", () => {
      const memory = orchestrator.getMemory();
      expect(memory).toBeDefined();
    });
  });

  // ========== 配置管理测试 ==========

  describe("配置管理", () => {
    test("应该能更新配置", () => {
      orchestrator.updateConfig({ maxRetries: 10 });
      const config = orchestrator.getConfig();
      expect(config.maxRetries).toBe(10);
    });

    test("更新配置应该保留其他配置项", () => {
      const originalConfig = orchestrator.getConfig();
      orchestrator.updateConfig({ maxRetries: 10 });
      const newConfig = orchestrator.getConfig();

      expect(newConfig.maxIterations).toBe(originalConfig.maxIterations);
      expect(newConfig.enableLearning).toBe(originalConfig.enableLearning);
    });
  });

  // ========== 对话历史测试 ==========

  describe("对话历史", () => {
    test("初始对话历史应该为空", () => {
      const history = orchestrator.getConversationHistory();
      expect(history).toEqual([]);
    });

    test("清除对话历史应该正常工作", () => {
      orchestrator.clearConversationHistory();
      const history = orchestrator.getConversationHistory();
      expect(history).toEqual([]);
    });

    test("对话历史应该是只读副本", () => {
      const history1 = orchestrator.getConversationHistory();
      history1.push({ role: "user", content: "test" });
      const history2 = orchestrator.getConversationHistory();
      expect(history2).toEqual([]);
    });
  });

  // ========== 反假完成状态测试 ==========

  describe("反假完成状态", () => {
    test("初始状态应该显示已启用但无运行实例", () => {
      const status = orchestrator.getAntiHallucinationStatus();
      expect(status.enabled).toBe(true);
      expect(status.runId).toBeUndefined();
    });

    test("禁用反假完成时状态应该正确", () => {
      const disabledOrchestrator = createAgentOrchestrator({
        enableAntiHallucination: false,
      });
      const status = disabledOrchestrator.getAntiHallucinationStatus();
      expect(status.enabled).toBe(false);
    });
  });

  // ========== 事件系统测试 ==========

  describe("事件系统", () => {
    test("应该能注册事件监听器", () => {
      const handler = jest.fn();
      orchestrator.on("phase:changed", handler);
      // 验证不抛出错误
      expect(true).toBe(true);
    });

    test("应该能移除事件监听器", () => {
      const handler = jest.fn();
      orchestrator.on("phase:changed", handler);
      orchestrator.off("phase:changed", handler);
      // 验证不抛出错误
      expect(true).toBe(true);
    });
  });

  // ========== 执行流程测试 ==========

  describe("执行流程", () => {
    test("应该能执行简单查询", async () => {
      const result = await orchestrator.run({
        userMessage: "获取当前选区",
      });

      expect(result.success).toBe(true);
      expect(result.message).toBeDefined();
      expect(result.state).toBeDefined();
      expect(result.state.phase).toBe("completed");
    });

    test("执行后应该保存对话历史", async () => {
      await orchestrator.run({
        userMessage: "测试消息",
      });

      const history = orchestrator.getConversationHistory();
      expect(history.length).toBeGreaterThan(0);
      expect(history[0].role).toBe("user");
      expect(history[0].content).toBe("测试消息");
    });

    test("应该触发事件", async () => {
      const phaseHandler = jest.fn();
      const intentHandler = jest.fn();

      orchestrator.on("phase:changed", phaseHandler);
      orchestrator.on("intent:parsed", intentHandler);

      await orchestrator.run({
        userMessage: "测试消息",
      });

      expect(phaseHandler).toHaveBeenCalled();
      expect(intentHandler).toHaveBeenCalled();
    });
  });

  // ========== 澄清处理测试 ==========

  describe("澄清处理", () => {
    test("需要澄清时应该返回澄清结果", async () => {
      // 临时修改 mock
      const IntentParser = require("../agent/IntentParser").IntentParser;
      IntentParser.mockImplementationOnce(() => ({
        parse: jest.fn().mockResolvedValue({
          intent: "unknown",
          confidence: 0.3,
          needsClarification: true,
          clarificationQuestion: "请问您想做什么？",
        }),
      }));

      const newOrchestrator = createAgentOrchestrator();
      const result = await newOrchestrator.run({
        userMessage: "做点什么",
      });

      expect(result.needsClarification).toBe(true);
      expect(result.clarificationQuestion).toBeDefined();
    });
  });

  // ========== 错误处理测试 ==========

  describe("错误处理", () => {
    test("编译失败时应该返回错误结果", async () => {
      // 临时修改 mock
      const SpecCompiler = require("../agent/SpecCompiler").SpecCompiler;
      SpecCompiler.mockImplementationOnce(() => ({
        compile: jest.fn().mockReturnValue({
          success: false,
          error: "无法编译意图",
        }),
      }));

      const newOrchestrator = createAgentOrchestrator();
      const result = await newOrchestrator.run({
        userMessage: "做一个复杂操作",
      });

      expect(result.success).toBe(false);
      expect(result.message).toContain("失败");
    });
  });
});

// ========== 配置常量测试 ==========

describe("DEFAULT_ORCHESTRATOR_CONFIG", () => {
  test("应该有合理的默认值", () => {
    expect(DEFAULT_ORCHESTRATOR_CONFIG.maxRetries).toBe(3);
    expect(DEFAULT_ORCHESTRATOR_CONFIG.maxIterations).toBe(10);
    expect(DEFAULT_ORCHESTRATOR_CONFIG.enableLearning).toBe(true);
    expect(DEFAULT_ORCHESTRATOR_CONFIG.enableAutoFix).toBe(true);
    expect(DEFAULT_ORCHESTRATOR_CONFIG.verificationTimeout).toBe(5000);
    expect(DEFAULT_ORCHESTRATOR_CONFIG.confirmBeforeWrite).toBe(false);
    expect(DEFAULT_ORCHESTRATOR_CONFIG.enableAntiHallucination).toBe(true);
    expect(DEFAULT_ORCHESTRATOR_CONFIG.enableConversationContext).toBe(true);
  });
});

// ========== 工厂函数测试 ==========

describe("createAgentOrchestrator", () => {
  test("应该创建 AgentOrchestrator 实例", () => {
    const orchestrator = createAgentOrchestrator();
    expect(orchestrator).toBeInstanceOf(AgentOrchestrator);
  });

  test("应该接受部分配置", () => {
    const orchestrator = createAgentOrchestrator({ maxRetries: 5 });
    expect(orchestrator.getConfig().maxRetries).toBe(5);
  });
});
