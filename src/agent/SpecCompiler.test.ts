import { SpecCompiler, SpecCompileContext } from "./SpecCompiler";
import { IntentSpec, WriteDataSpec, FormulaSpec, SheetSpec } from "./types/intent";

describe("SpecCompiler", () => {
  let compiler: SpecCompiler;

  beforeEach(() => {
    compiler = new SpecCompiler();
  });

  describe("compressedIntent routing", () => {
    it("sets routing hint for failure intent", () => {
      const spec: IntentSpec = {
        intent: "write_data",
        confidence: 0.9,
        needsClarification: false,
        spec: {
          type: "write_data",
          target: "A1",
          data: [["test"]],
        } as WriteDataSpec,
        compressedIntent: "failure",
      };

      const context: SpecCompileContext = {};
      compiler.compile(spec, context);

      // 验证 context 中设置了路由提示
      expect((context as Record<string, unknown>).__routingHint).toBeDefined();
      const hint = (context as Record<string, unknown>).__routingHint as {
        priority: string;
        addDiagnosticStep: boolean;
      };
      expect(hint.priority).toBe("diagnose");
      expect(hint.addDiagnosticStep).toBe(true);
    });

    it("sets routing hint for automation intent", () => {
      const spec: IntentSpec = {
        intent: "create_formula",
        confidence: 0.85,
        needsClarification: false,
        spec: {
          type: "formula",
          targetCell: "B2",
          formulaType: "sum",
          sourceRange: "A:A",
        } as FormulaSpec,
        compressedIntent: "automation",
      };

      const context: SpecCompileContext = {};
      compiler.compile(spec, context);

      const hint = (context as Record<string, unknown>).__routingHint as {
        priority: string;
        suggestedTools: string[];
      };
      expect(hint.priority).toBe("batch");
      expect(hint.suggestedTools).toContain("excel_fill_formula");
    });

    it("sets routing hint for structure intent", () => {
      const spec: IntentSpec = {
        intent: "create_sheet",
        confidence: 0.8,
        needsClarification: false,
        spec: {
          type: "sheet",
          operation: "create",
          sheetName: "NewSheet",
        } as SheetSpec,
        compressedIntent: "structure",
      };

      const context: SpecCompileContext = {};
      compiler.compile(spec, context);

      const hint = (context as Record<string, unknown>).__routingHint as {
        priority: string;
        message: string;
      };
      expect(hint.priority).toBe("refactor");
      expect(hint.message).toContain("结构");
    });

    it("sets routing hint for maintainability intent", () => {
      const spec: IntentSpec = {
        intent: "write_data",
        confidence: 0.9,
        needsClarification: false,
        spec: {
          type: "write_data",
          target: "A1",
          data: [["protected data"]],
        } as WriteDataSpec,
        compressedIntent: "maintainability",
      };

      const context: SpecCompileContext = {};
      compiler.compile(spec, context);

      const hint = (context as Record<string, unknown>).__routingHint as {
        priority: string;
        suggestedTools: string[];
      };
      expect(hint.priority).toBe("protect");
      expect(hint.suggestedTools).toContain("excel_protect_sheet");
    });

    it("handles missing compressedIntent gracefully", () => {
      const spec: IntentSpec = {
        intent: "write_data",
        confidence: 0.9,
        needsClarification: false,
        spec: {
          type: "write_data",
          target: "A1",
          data: [["test"]],
        } as WriteDataSpec,
        // no compressedIntent
      };

      const context: SpecCompileContext = {};
      // 不应该抛出错误
      expect(() => compiler.compile(spec, context)).not.toThrow();
      // 不应该设置路由提示
      expect((context as Record<string, unknown>).__routingHint).toBeUndefined();
    });
  });
});
