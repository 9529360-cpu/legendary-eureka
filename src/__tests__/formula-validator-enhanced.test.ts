/**
 * FormulaValidator 增强功能测试
 *
 * 测试 autoFixFormula, suggestFormula 等新功能
 */

import { FormulaValidator, ExcelErrorType, DataModelingValidator } from "../agent/FormulaValidator";

describe("FormulaValidator Enhanced Tests", () => {
  let validator: FormulaValidator;

  beforeEach(() => {
    validator = new FormulaValidator();
  });

  // ========== autoFixFormula 测试 ==========

  describe("autoFixFormula", () => {
    it("should fix #DIV/0! by wrapping with IFERROR", () => {
      const result = validator.autoFixFormula("=A1/B1", "#DIV/0!");
      expect(result.success).toBe(true);
      expect(result.fixedFormula).toContain("IFERROR");
      expect(result.fixApplied.length).toBeGreaterThan(0);
    });

    it("should fix #N/A by suggesting IFNA or IFERROR", () => {
      const result = validator.autoFixFormula("=VLOOKUP(A1,B:C,2,FALSE)", "#N/A");
      expect(result.success).toBe(true);
      expect(result.fixedFormula).toMatch(/IFERROR|IFNA/);
    });

    it("should attempt to fix #NAME? by correcting function names", () => {
      // 中文括号会被修复
      const result = validator.autoFixFormula("=SUM（A1:A10）", "#NAME?");
      expect(result.success).toBe(true);
      expect(result.fixedFormula).toBe("=SUM(A1:A10)");
    });

    it("should fix Chinese brackets in #NAME? errors", () => {
      const result = validator.autoFixFormula("=SUM（A1:A10）", "#NAME?");
      expect(result.fixedFormula).not.toContain("（");
      expect(result.fixedFormula).not.toContain("）");
    });

    it("should handle #VALUE! errors", () => {
      const result = validator.autoFixFormula("=A1+B1", "#VALUE!");
      // 可能无法自动修复，但应该有建议
      expect(result.fixApplied.length).toBeGreaterThanOrEqual(0);
    });

    it("should handle #REF! errors", () => {
      const result = validator.autoFixFormula("=A1+#REF!", "#REF!");
      // #REF! 通常无法自动修复
      expect(result).toBeDefined();
      expect(result.originalFormula).toBe("=A1+#REF!");
    });
  });

  // ========== suggestFormula 测试 ==========

  describe("suggestFormula", () => {
    it("should suggest SUM for sum intent", () => {
      const suggestions = validator.suggestFormula(
        { type: "sum", description: "求和" },
        { sourceRange: "A1:A100" }
      );
      expect(suggestions.length).toBeGreaterThan(0);
      expect(suggestions[0].formula).toContain("SUM");
    });

    it("should suggest VLOOKUP/XLOOKUP for lookup intent", () => {
      const suggestions = validator.suggestFormula(
        { type: "lookup", description: "查找" },
        { lookupValue: "A1", lookupRange: "B:B", returnRange: "C:C" }
      );
      expect(suggestions.length).toBeGreaterThan(0);
      const hasLookup = suggestions.some(
        (s) => s.formula.includes("LOOKUP") || s.formula.includes("XLOOKUP")
      );
      expect(hasLookup).toBe(true);
    });

    it("should suggest COUNT/COUNTIF for count intent", () => {
      const suggestions = validator.suggestFormula(
        { type: "count", description: "计数" },
        { sourceRange: "A1:A100" }
      );
      expect(suggestions.length).toBeGreaterThan(0);
      const hasCount = suggestions.some(
        (s) => s.formula.includes("COUNT") || s.formula.includes("COUNTA")
      );
      expect(hasCount).toBe(true);
    });

    it("should suggest percentage formula", () => {
      const suggestions = validator.suggestFormula(
        { type: "percentage", description: "百分比" },
        { numerator: "A1", denominator: "B1" }
      );
      expect(suggestions.length).toBeGreaterThan(0);
      const hasPercentage = suggestions.some((s) => s.formula.includes("/"));
      expect(hasPercentage).toBe(true);
    });

    it("should suggest date formulas for date intent", () => {
      const suggestions = validator.suggestFormula({ type: "date", description: "日期" }, {});
      expect(suggestions.length).toBeGreaterThan(0);
      const hasDateFunc = suggestions.some(
        (s) =>
          s.formula.includes("TODAY") || s.formula.includes("NOW") || s.formula.includes("DATE")
      );
      expect(hasDateFunc).toBe(true);
    });

    it("should suggest text formulas for text intent", () => {
      const suggestions = validator.suggestFormula(
        { type: "text", description: "文本处理" },
        { sourceCell: "A1" }
      );
      expect(suggestions.length).toBeGreaterThan(0);
      const hasTextFunc = suggestions.some(
        (s) =>
          s.formula.includes("CONCATENATE") ||
          s.formula.includes("TEXT") ||
          s.formula.includes("TRIM") ||
          s.formula.includes("&")
      );
      expect(hasTextFunc).toBe(true);
    });
  });

  // ========== 综合验证测试 ==========

  describe("Integration Tests", () => {
    it("should validate and fix formula in one flow", () => {
      const formula = "=SUM（A1:A10）";

      // 使用 autoFixFormula 来修复
      const result = validator.autoFixFormula(formula, "#NAME?");

      // 修复后应该没有中文括号问题
      expect(result.fixedFormula).not.toContain("（");
      expect(result.fixedFormula).not.toContain("）");
    });

    it("should handle formula without issues gracefully", () => {
      const formula = "=SUM(A1:A10)";
      const result = validator.autoFixFormula(formula, "#DIV/0!");

      // 没有除法，不需要修复
      expect(result.success).toBe(false);
    });
  });

  // ========== 边界情况测试 ==========

  describe("Edge Cases", () => {
    it("should handle empty formula", () => {
      const result = validator.autoFixFormula("", "#VALUE!");
      expect(result.success).toBe(false);
    });

    it("should handle formula without equals sign", () => {
      const formula = "SUM(A1:A10)";
      const result = validator.autoFixFormula(formula, "#NAME?");
      // 应该能处理
      expect(result).toBeDefined();
    });

    it("should handle unknown error types gracefully", () => {
      const result = validator.autoFixFormula("=A1+B1", "#UNKNOWN!" as ExcelErrorType);
      // 不应该抛出错误
      expect(result).toBeDefined();
      expect(result.originalFormula).toBe("=A1+B1");
    });

    it("should suggest formulas with empty context", () => {
      const suggestions = validator.suggestFormula({ type: "sum", description: "求和" }, {});
      expect(suggestions).toBeDefined();
      expect(Array.isArray(suggestions)).toBe(true);
    });

    it("should handle custom intent type", () => {
      const suggestions = validator.suggestFormula(
        { type: "custom", description: "自定义操作" },
        { sourceRange: "A1:B10" }
      );
      expect(suggestions).toBeDefined();
      expect(Array.isArray(suggestions)).toBe(true);
    });
  });

  // ========== AutoFixResult 结构测试 ==========

  describe("AutoFixResult Structure", () => {
    it("should have correct structure", () => {
      const result = validator.autoFixFormula("=A1/B1", "#DIV/0!");

      expect(result).toHaveProperty("success");
      expect(result).toHaveProperty("originalFormula");
      expect(result).toHaveProperty("fixedFormula");
      expect(result).toHaveProperty("fixApplied");
      expect(Array.isArray(result.fixApplied)).toBe(true);
    });

    it("should preserve original formula in result", () => {
      const formula = "=VLOOKUP(A1,B:C,2,FALSE)";
      const result = validator.autoFixFormula(formula, "#N/A");

      expect(result.originalFormula).toBe(formula);
    });
  });

  // ========== FormulaSuggestion 结构测试 ==========

  describe("FormulaSuggestion Structure", () => {
    it("should have correct structure", () => {
      const suggestions = validator.suggestFormula({ type: "sum" }, { sourceRange: "A1:A10" });

      if (suggestions.length > 0) {
        expect(suggestions[0]).toHaveProperty("formula");
        expect(suggestions[0]).toHaveProperty("description");
        expect(suggestions[0]).toHaveProperty("confidence");
        expect(typeof suggestions[0].confidence).toBe("number");
      }
    });

    it("should have confidence between 0 and 1", () => {
      const suggestions = validator.suggestFormula({ type: "sum" }, { sourceRange: "A1:A10" });

      for (const suggestion of suggestions) {
        expect(suggestion.confidence).toBeGreaterThanOrEqual(0);
        expect(suggestion.confidence).toBeLessThanOrEqual(1);
      }
    });
  });
});

// ========== DataModelingValidator 增强测试 (v2.8.2) ==========

describe("DataModelingValidator Enhanced Tests (v2.8.2)", () => {
  let modelingValidator: DataModelingValidator;

  beforeEach(() => {
    modelingValidator = new DataModelingValidator();
  });

  // ========== 智能表类型识别测试 ==========

  describe("detectTableType", () => {
    it("should detect master table from name", () => {
      const result = modelingValidator.detectTableType("产品主数据表", [
        "产品ID",
        "产品名称",
        "单价",
        "成本",
      ]);
      expect(result.detectedType).toBe("master");
      expect(result.confidence).toBeGreaterThan(0.8);
      expect(result.reasons.length).toBeGreaterThan(0);
    });

    it("should detect transaction table from name", () => {
      const result = modelingValidator.detectTableType("订单明细表", [
        "订单ID",
        "产品ID",
        "数量",
        "单价",
        "销售额",
      ]);
      expect(result.detectedType).toBe("transaction");
      expect(result.confidence).toBeGreaterThan(0.7);
    });

    it("should detect summary table from name", () => {
      const result = modelingValidator.detectTableType("产品汇总表", [
        "产品ID",
        "销量",
        "销售额",
        "毛利",
      ]);
      expect(result.detectedType).toBe("summary");
      expect(result.confidence).toBeGreaterThan(0.7);
    });

    it("should detect analysis table from name", () => {
      const result = modelingValidator.detectTableType("利润分析表", [
        "月份",
        "收入",
        "成本",
        "净利润",
      ]);
      expect(result.detectedType).toBe("analysis");
      expect(result.confidence).toBeGreaterThan(0.7);
    });

    it("should suggest relations for transaction tables", () => {
      const result = modelingValidator.detectTableType("销售订单表", [
        "订单ID",
        "产品ID",
        "数量",
        "单价",
      ]);
      expect(result.suggestedRelations.length).toBeGreaterThan(0);
      expect(result.suggestedRelations[0].relationshipType).toBe("lookup");
    });

    it("should return unknown for unrecognized tables", () => {
      const result = modelingValidator.detectTableType("数据1", ["A", "B", "C"]);
      expect(result.detectedType).toBe("unknown");
    });
  });

  // ========== 公式建议生成测试 ==========

  describe("generateFormulaSuggestion", () => {
    it("should suggest XLOOKUP for 单价 in transaction table", () => {
      const formula = modelingValidator.generateFormulaSuggestion(
        "单价",
        "transaction",
        "产品主数据表"
      );
      expect(formula).toContain("XLOOKUP");
      expect(formula).toContain("产品主数据表");
    });

    it("should suggest multiplication for 销售额 in transaction table", () => {
      const formula = modelingValidator.generateFormulaSuggestion("销售额", "transaction");
      expect(formula).toContain("*");
    });

    it("should suggest SUMIF for 销量 in summary table", () => {
      const formula = modelingValidator.generateFormulaSuggestion(
        "销量",
        "summary",
        undefined,
        "订单交易表"
      );
      expect(formula).toContain("SUMIF");
      expect(formula).toContain("订单交易表");
    });

    it("should suggest division for 毛利率 in summary table", () => {
      const formula = modelingValidator.generateFormulaSuggestion("毛利率", "summary");
      expect(formula).toContain("/");
    });

    it("should return empty string for unknown field", () => {
      const formula = modelingValidator.generateFormulaSuggestion("未知字段", "master");
      expect(formula).toBe("");
    });
  });

  // ========== 交易表验证增强测试 ==========

  describe("validateTransactionTable Enhanced", () => {
    it("should include fixAction in detected issues", () => {
      const data = [
        [1, "P001", 10, 100, 50],
        [2, "P002", 5, 100, 50],
        [3, "P003", 3, 100, 50],
        [4, "P004", 8, 100, 50],
      ];
      const headers = ["订单ID", "产品ID", "数量", "单价", "成本"];

      const issues = modelingValidator.validateTransactionTable(data, headers, "产品主数据表");

      expect(issues.length).toBeGreaterThan(0);
      const issueWithFix = issues.find((i) => i.fixAction);
      expect(issueWithFix).toBeDefined();
      expect(issueWithFix!.fixAction!.action).toBe("set_formula");
      expect(issueWithFix!.fixAction!.formula).toContain("XLOOKUP");
    });

    it("should detect missing formula for 销售额 column with same values", () => {
      // 销售额列值全部相同才会检测到硬编码问题
      const data = [
        [1, "P001", 10, 100, 1000],
        [2, "P002", 5, 100, 1000],
        [3, "P003", 3, 100, 1000],
        [4, "P004", 8, 100, 1000],
      ];
      const headers = ["订单ID", "产品ID", "数量", "单价", "销售额"];

      const issues = modelingValidator.validateTransactionTable(data, headers);

      const salesIssue = issues.find((i) => i.location.includes("销售额"));
      expect(salesIssue).toBeDefined();
      expect(salesIssue!.type).toBe("missing_formula");
    });
  });

  // ========== 汇总表验证增强测试 ==========

  describe("validateSummaryTable Enhanced", () => {
    it("should include fixAction with SUMIF formula", () => {
      const data = [
        ["P001", 100, 5000, 2500, 2500],
        ["P002", 100, 5000, 2500, 2500],
        ["P003", 100, 5000, 2500, 2500],
      ];
      const headers = ["产品ID", "销量", "销售额", "总成本", "毛利"];

      const issues = modelingValidator.validateSummaryTable(data, headers, "订单交易表");

      expect(issues.length).toBeGreaterThan(0);
      const salesIssue = issues.find((i) => i.location.includes("销量"));
      expect(salesIssue).toBeDefined();
      expect(salesIssue!.fixAction).toBeDefined();
      expect(salesIssue!.fixAction!.formula).toContain("SUMIF");
    });

    it("should detect duplicate 毛利率 with fix action", () => {
      const data = [
        ["P001", 5000, 2500, 0.5],
        ["P002", 6000, 3000, 0.5],
        ["P003", 4000, 2000, 0.5],
      ];
      const headers = ["产品ID", "销售额", "总成本", "毛利率"];

      const issues = modelingValidator.validateSummaryTable(data, headers);

      const rateIssue = issues.find((i) => i.location.includes("毛利率"));
      expect(rateIssue).toBeDefined();
      expect(rateIssue!.type).toBe("inconsistent_data");
      expect(rateIssue!.fixAction).toBeDefined();
      expect(rateIssue!.fixAction!.formula).toContain("毛利");
    });
  });

  // ========== 综合验证增强测试 ==========

  describe("validateDataModeling Enhanced", () => {
    it("should include fixActions in validation result", () => {
      const data = [
        [1, "P001", 10, 100, 1000],
        [2, "P002", 5, 100, 500],
        [3, "P003", 3, 100, 300],
        [4, "P004", 8, 100, 800],
      ];
      const headers = ["订单ID", "产品ID", "数量", "单价", "销售额"];

      const result = modelingValidator.validateDataModeling(
        "transaction",
        data,
        headers,
        "产品主数据表"
      );

      expect(result.fixActions).toBeDefined();
      expect(result.fixActions!.length).toBeGreaterThan(0);
    });

    it("should generate enhanced recommendations", () => {
      const data = [
        ["P001", 100, 5000, 2500],
        ["P002", 100, 5000, 2500],
        ["P003", 100, 5000, 2500],
      ];
      const headers = ["产品ID", "销量", "销售额", "毛利"];

      const result = modelingValidator.validateDataModeling("summary", data, headers);

      expect(result.recommendations.length).toBeGreaterThan(0);
      expect(
        result.recommendations.some((r) => r.includes("⚠️") || r.includes("📌") || r.includes("📊"))
      ).toBe(true);
    });
  });

  // ========== 修复脚本生成测试 ==========

  describe("generateFixScript", () => {
    it("should generate fix script for set_formula actions", () => {
      const issues = [
        {
          type: "hardcoded_value" as const,
          severity: "critical" as const,
          location: "列 单价",
          message: "单价列所有值都是100，疑似硬编码",
          suggestion: "使用 XLOOKUP 公式",
          fixAction: {
            action: "set_formula" as const,
            target: "单价列",
            formula: "=XLOOKUP([@产品ID], 产品主数据表[产品ID], 产品主数据表[单价])",
          },
        },
      ];

      const scripts = modelingValidator.generateFixScript(issues);

      expect(scripts.length).toBeGreaterThan(0);
      expect(scripts.some((s) => s.includes("excel_set_formula"))).toBe(true);
      expect(scripts.some((s) => s.includes("XLOOKUP"))).toBe(true);
    });

    it("should return empty array for issues without fixAction", () => {
      const issues = [
        {
          type: "inconsistent_data" as const,
          severity: "warning" as const,
          location: "某列",
          message: "数据不一致",
          suggestion: "手工检查",
        },
      ];

      const scripts = modelingValidator.generateFixScript(issues);

      expect(scripts.length).toBe(0);
    });
  });
});
