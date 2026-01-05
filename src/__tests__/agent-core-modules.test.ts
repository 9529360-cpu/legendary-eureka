/**
 * Agent Core 模块化架构测试
 *
 * 测试目标：
 * 1. SemanticExtractor 语义提取
 * 2. DiagnosticEngine 诊断引擎
 * 3. SolutionBuilder 解决方案构建
 * 4. AgentOrchestrator 编排器
 */

import { SemanticExtractor, semanticExtractor } from "../agent/core/semantic/SemanticExtractor";
import { DiagnosticEngine, diagnosticEngine } from "../agent/core/semantic/DiagnosticEngine";
import { SolutionBuilder, solutionBuilder } from "../agent/core/solutions/SolutionBuilder";
import { AgentOrchestrator } from "../agent/core/AgentOrchestrator";
import { getToolContract, getAllToolContracts } from "../agent/core/contracts/ToolContract";

// ========== SemanticExtractor 测试 ==========

describe("SemanticExtractor", () => {
  const extractor = new SemanticExtractor();

  describe("意图提取", () => {
    test("识别跨表汇总意图", () => {
      const result = extractor.extract("帮我把多表汇总数据整理一下");
      expect(result.intent).toBe("cross_sheet_summary");
      expect(result.confidence).toBeGreaterThan(0.5);
    });

    test("识别格式化意图", () => {
      const result = extractor.extract("帮我格式化这个范围");
      expect(result.intent).toBe("format_range");
      expect(result.confidence).toBeGreaterThan(0.5);
    });

    test("识别数据清洗意图", () => {
      const result = extractor.extract("清洗一下这些数据");
      expect(result.intent).toBe("data_cleanup");
    });

    test("识别图表创建意图", () => {
      const result = extractor.extract("创建一个销售数据的柱状图");
      expect(result.intent).toBe("chart_creation");
    });

    test("识别诊断意图", () => {
      const result = extractor.extract("为什么公式结果是0");
      expect(result.intent).toBe("diagnose_zero");
    });
  });

  describe("实体提取", () => {
    test("提取列名（通过范围模式）", () => {
      const result = extractor.extract("计算A1:B10的数据");
      expect(result.entities.ranges).toContain("A1:B10");
    });

    test("提取范围", () => {
      const result = extractor.extract("计算A1:B10的平均值");
      expect(result.entities.ranges).toContain("A1:B10");
    });

    test("提取工作表名（Sheet格式）", () => {
      const result = extractor.extract("在'数据表'sheet中查找");
      expect(result.entities.sheets).toContain("数据表");
    });
  });

  describe("约束提取", () => {
    test("识别紧急约束", () => {
      const result = extractor.extract("紧急！马上帮我修复这个公式");
      expect(result.constraints.urgent).toBe(true);
    });

    test("识别保持格式约束", () => {
      const result = extractor.extract("修改公式但保持原有格式");
      expect(result.constraints.preserveFormat).toBe(true);
    });

    test("识别只读约束", () => {
      const result = extractor.extract("只查看数据，不要修改");
      expect(result.constraints.readOnly).toBe(true);
    });
  });

  describe("单例导出", () => {
    test("semanticExtractor 单例可用", () => {
      const result = semanticExtractor.extract("测试输入");
      expect(result).toBeDefined();
      expect(result.intent).toBeDefined();
    });
  });
});

// ========== DiagnosticEngine 测试 ==========

describe("DiagnosticEngine", () => {
  const engine = new DiagnosticEngine();

  describe("症状诊断", () => {
    test("诊断结果为0的问题", () => {
      const result = engine.diagnose("为什么SUM公式结果是0");
      expect(result.possibleCauses.length).toBeGreaterThan(0);
      expect(result.possibleCauses[0].rank).toBe(1);
      expect(result.possibleCauses[0].probability).toBeGreaterThan(0);
    });

    test("诊断循环引用", () => {
      const result = engine.diagnose("公式显示循环引用错误");
      expect(result.possibleCauses.length).toBeGreaterThan(0);
      expect(result.possibleCauses[0].shortestValidation).toBeDefined();
    });

    test("诊断 IMPORTRANGE 问题", () => {
      const result = engine.diagnose("IMPORTRANGE不工作，需要允许访问");
      expect(result.possibleCauses.length).toBeGreaterThan(0);
    });
  });

  describe("验证步骤", () => {
    test("提供验证步骤", () => {
      const result = engine.diagnose("数字被当成文本了");
      expect(result.validationSteps.length).toBeGreaterThan(0);
      expect(result.validationSteps[0].order).toBe(1);
    });
  });

  describe("格式化输出", () => {
    test("formatDiagnosis 生成可读文本", () => {
      const result = engine.diagnose("公式返回0");
      const formatted = engine.formatDiagnosis(result);
      expect(formatted).toContain("Top3 可能原因");
      expect(formatted).toContain("验证步骤");
    });
  });

  describe("单例导出", () => {
    test("diagnosticEngine 单例可用", () => {
      const result = diagnosticEngine.diagnose("测试");
      expect(result).toBeDefined();
    });
  });
});

// ========== SolutionBuilder 测试 ==========

describe("SolutionBuilder", () => {
  const builder = new SolutionBuilder();

  describe("从语义提取构建解决方案", () => {
    test("构建公式创建解决方案", () => {
      const extraction = semanticExtractor.extract("创建一个求和公式");
      const solution = builder.buildFromSemanticExtraction(extraction);

      expect(solution.minimal).toBeDefined();
      expect(solution.recommended).toBeDefined();
      expect(solution.structural).toBeDefined();

      expect(solution.minimal.tier).toBe("minimal");
      expect(solution.recommended.tier).toBe("recommended");
      expect(solution.structural!.tier).toBe("structural");
    });

    test("解决方案包含步骤", () => {
      const extraction = semanticExtractor.extract("格式化表格");
      const solution = builder.buildFromSemanticExtraction(extraction);

      expect(solution.minimal.steps).toBeDefined();
      expect(solution.minimal.steps!.length).toBeGreaterThan(0);
    });
  });

  describe("从诊断结果构建解决方案", () => {
    test("根据诊断构建解决方案", () => {
      const diagnosis = diagnosticEngine.diagnose("公式返回0");
      const solution = builder.buildFromDiagnosis(diagnosis);

      expect(solution.minimal.emoji).toBe("🚀");
      expect(solution.recommended.emoji).toBe("✅");
      expect(solution.structural!.emoji).toBe("🏗️");
    });
  });

  describe("格式化输出", () => {
    test("formatSolution 生成分层文本", () => {
      const extraction = semanticExtractor.extract("分析数据");
      const solution = builder.buildFromSemanticExtraction(extraction);
      const formatted = builder.formatSolution(solution);

      expect(formatted).toContain("🚀");
      expect(formatted).toContain("✅");
      expect(formatted).toContain("🏗️");
    });
  });

  describe("单例导出", () => {
    test("solutionBuilder 单例可用", () => {
      const extraction = semanticExtractor.extract("测试");
      const solution = solutionBuilder.buildFromSemanticExtraction(extraction);
      expect(solution).toBeDefined();
    });
  });
});

// ========== AgentOrchestrator 测试 ==========

describe("AgentOrchestrator", () => {
  describe("工作流处理", () => {
    test("处理用户输入并返回结果", async () => {
      const orchestrator = new AgentOrchestrator();
      const result = await orchestrator.process("帮我计算A列的总和");

      expect(result.phase).toBe("completed");
      expect(result.semanticExtraction).toBeDefined();
    });

    test("低置信度时请求澄清", async () => {
      const orchestrator = new AgentOrchestrator({
        confidenceThreshold: 0.99, // 设置很高的阈值
      });
      const result = await orchestrator.process("xyz");

      expect(result.phase).toBe("awaiting_clarification");
      expect(result.clarificationNeeded).toBeDefined();
    });

    test("问题类输入触发诊断", async () => {
      const orchestrator = new AgentOrchestrator({
        enableDiagnosis: true,
        confidenceThreshold: 0.3, // 降低阈值确保不被拦截
      });
      const result = await orchestrator.process("这个公式有错误#REF!");

      expect(result.diagnosis).toBeDefined();
    });
  });

  describe("事件系统", () => {
    test("注册和触发事件", async () => {
      const orchestrator = new AgentOrchestrator();
      const events: string[] = [];

      orchestrator.on("phase_change", (e) => {
        events.push(e.type);
      });

      await orchestrator.process("测试输入");
      expect(events.length).toBeGreaterThan(0);
    });

    test("移除事件监听", () => {
      const orchestrator = new AgentOrchestrator();
      const handler = () => {};

      orchestrator.on("test", handler);
      orchestrator.off("test", handler);
      // 无异常即为通过
    });
  });

  describe("配置管理", () => {
    test("更新配置", () => {
      const orchestrator = new AgentOrchestrator();
      orchestrator.updateConfig({ confidenceThreshold: 0.8 });

      const config = orchestrator.getConfig();
      expect(config.confidenceThreshold).toBe(0.8);
    });
  });

  describe("响应格式化", () => {
    test("formatResponse 生成完整响应", async () => {
      const orchestrator = new AgentOrchestrator();
      const result = await orchestrator.process("创建一个求和公式");
      const formatted = orchestrator.formatResponse(result);

      expect(formatted).toContain("理解您的需求");
    });
  });
});

// ========== ToolContract 测试 ==========

describe("ToolContract", () => {
  test("获取单个工具契约", () => {
    const contract = getToolContract("read_sheet");
    expect(contract).toBeDefined();
    expect(contract!.name).toBe("read_sheet");
    expect(contract!.inputSchema).toBeDefined();
    expect(contract!.outputSchema).toBeDefined();
  });

  test("获取所有工具契约", () => {
    const contracts = getAllToolContracts();
    expect(contracts.length).toBeGreaterThan(0);
  });

  test("工具契约包含失败模式", () => {
    const contract = getToolContract("read_sheet");
    expect(contract!.failureModes).toBeDefined();
    expect(contract!.failureModes!.length).toBeGreaterThan(0);
  });

  test("工具契约包含类别", () => {
    const contract = getToolContract("write_sheet");
    expect(contract!.category).toBeDefined();
    expect(contract!.category).toBe("write");
  });
});
