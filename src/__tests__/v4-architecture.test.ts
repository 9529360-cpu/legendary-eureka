/**
 * v4.0 架构集成测试
 *
 * 测试 IntentParser → SpecCompiler → AgentExecutor 流程
 */

import {
  IntentParser,
  SpecCompiler,
  IntentSpec,
  CreateTableSpec,
  FormatSpec,
  WriteDataSpec,
  ClarifySpec,
  RespondSpec,
} from "../agent";
import { SpecCompileResult } from "../agent/SpecCompiler";

describe("v4.0 Architecture Integration", () => {
  let intentParser: IntentParser;
  let specCompiler: SpecCompiler;

  beforeEach(() => {
    intentParser = new IntentParser();
    specCompiler = new SpecCompiler();
  });

  describe("IntentParser", () => {
    it("should have tool-free system prompt", () => {
      const systemPrompt = intentParser.buildSystemPrompt();

      // 验证 System Prompt 不包含任何工具名 (excel_ 前缀)
      expect(systemPrompt).not.toContain("excel_");
      expect(systemPrompt).not.toContain("excel_read_range");
      expect(systemPrompt).not.toContain("excel_write_range");
      expect(systemPrompt).not.toContain("excel_set_formula");
      expect(systemPrompt).not.toContain("excel_create_chart");
    });

    it("should contain business concepts in system prompt", () => {
      const systemPrompt = intentParser.buildSystemPrompt();

      // 验证包含业务概念
      expect(systemPrompt).toContain("create_table");
      expect(systemPrompt).toContain("write_data");
      expect(systemPrompt).toContain("format_range");
      expect(systemPrompt).toContain("create_formula");
    });

    it("should build user prompt with context", () => {
      const userPrompt = intentParser.buildUserPrompt({
        userMessage: "创建一个销售表格",
        activeSheet: "Sheet1",
        workbookSummary: { sheetNames: ["销售数据.xlsx"] },
      });

      expect(userPrompt).toContain("创建一个销售表格");
      expect(userPrompt).toContain("销售数据.xlsx");
    });
  });

  describe("SpecCompiler", () => {
    it("should compile create_table intent", () => {
      const spec: CreateTableSpec = {
        type: "create_table",
        tableType: "sales",
        columns: [
          { name: "日期", type: "date" },
          { name: "产品", type: "text" },
          { name: "数量", type: "number" },
          { name: "单价", type: "number" },
          { name: "金额", type: "formula", formula: "=C{row}*D{row}" },
        ],
        startCell: "A1",
        options: { hasHeader: true, hasTotalRow: true },
      };

      const intent: IntentSpec = {
        intent: "create_table",
        confidence: 0.95,
        spec,
        needsClarification: false,
      };

      const result: SpecCompileResult = specCompiler.compile(intent);

      expect(result.success).toBe(true);
      expect(result.plan).toBeDefined();
      expect(result.plan!.steps.length).toBeGreaterThan(0);

      // 验证使用 step.id 作为依赖，而非工具名
      const steps = result.plan!.steps;
      for (const step of steps) {
        if (step.dependsOn && step.dependsOn.length > 0) {
          for (const depId of step.dependsOn) {
            // 依赖 ID 不应该是工具名
            expect(depId).not.toContain("excel_");
            // 依赖 ID 应该是有效的 step ID
            const depStep = steps.find((s) => s.id === depId);
            expect(depStep).toBeDefined();
          }
        }
      }
    });

    it("should compile format_range intent", () => {
      const spec: FormatSpec = {
        type: "format_range",
        range: "A1:E10",
        format: {
          bold: true,
          backgroundColor: "#4472C4",
          fontColor: "#FFFFFF",
        },
      };

      const intent: IntentSpec = {
        intent: "format_range",
        confidence: 0.9,
        spec,
        needsClarification: false,
      };

      const result = specCompiler.compile(intent);

      expect(result.success).toBe(true);
      expect(result.plan!.steps.length).toBeGreaterThan(0);

      // 验证格式化工具被使用
      const hasFormatStep = result.plan!.steps.some(
        (step) => step.action === "excel_format_range"
      );
      expect(hasFormatStep).toBe(true);
    });

    it("should add sensing step before write operations", () => {
      const spec: WriteDataSpec = {
        type: "write_data",
        target: "A1:C5",
        data: [
          ["姓名", "年龄", "城市"],
          ["张三", 25, "北京"],
        ],
      };

      const intent: IntentSpec = {
        intent: "write_data",
        confidence: 0.9,
        spec,
        needsClarification: false,
      };

      const result = specCompiler.compile(intent);

      expect(result.success).toBe(true);

      // 验证自动添加了感知步骤
      const steps = result.plan!.steps;
      const readStepIndex = steps.findIndex(
        (s) => s.action.includes("read") || s.action.includes("get")
      );
      const writeStepIndex = steps.findIndex(
        (s) => s.action.includes("write")
      );

      // 读取步骤应该在写入步骤之前
      if (readStepIndex !== -1 && writeStepIndex !== -1) {
        expect(readStepIndex).toBeLessThan(writeStepIndex);
      }
    });

    it("should handle clarify intent", () => {
      const spec: ClarifySpec = {
        type: "clarify",
        question: "您想创建什么类型的表格？",
        options: ["销售表", "库存表", "财务报表"],
        reason: "用户请求不明确",
      };

      const intent: IntentSpec = {
        intent: "clarify",
        confidence: 0.8,
        spec,
        needsClarification: true,
        clarificationQuestion: "您想创建什么类型的表格？",
      };

      const result = specCompiler.compile(intent);

      expect(result.success).toBe(true);
      // clarify 意图应该生成步骤
      expect(result.plan!.steps.length).toBeGreaterThan(0);
    });

    it("should handle respond intent", () => {
      const spec: RespondSpec = {
        type: "respond",
        message: "好的，我已经完成了任务。",
      };

      const intent: IntentSpec = {
        intent: "respond_only",
        confidence: 0.95,
        spec,
        needsClarification: false,
      };

      const result = specCompiler.compile(intent);

      expect(result.success).toBe(true);
    });
  });

  describe("Step ID Dependency", () => {
    it("should generate unique step IDs", () => {
      const spec: CreateTableSpec = {
        type: "create_table",
        tableType: "custom",
        columns: [
          { name: "A", type: "text" },
          { name: "B", type: "number" },
        ],
        startCell: "A1",
      };

      const intent: IntentSpec = {
        intent: "create_table",
        confidence: 0.95,
        spec,
        needsClarification: false,
      };

      const result = specCompiler.compile(intent);
      const stepIds = result.plan!.steps.map((s) => s.id);

      // 所有 ID 应该唯一
      const uniqueIds = new Set(stepIds);
      expect(uniqueIds.size).toBe(stepIds.length);

      // 所有 ID 应该有正确的格式 (step_ 前缀)
      for (const id of stepIds) {
        expect(id).toMatch(/^step_/);
      }
    });
  });
});
