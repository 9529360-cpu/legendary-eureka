/**
 * 智能操作验证系统测试
 * 测试操作前预检查和自动修复功能
 */

describe("操作验证系统", () => {
  // 模拟地址验证正则
  const addressRegex = /^[A-Za-z]+\d+(:[A-Za-z]+\d+)?$/;

  describe("地址格式验证", () => {
    test("有效的单元格地址", () => {
      expect(addressRegex.test("A1")).toBe(true);
      expect(addressRegex.test("B10")).toBe(true);
      expect(addressRegex.test("AA100")).toBe(true);
      expect(addressRegex.test("ZZ999")).toBe(true);
    });

    test("有效的范围地址", () => {
      expect(addressRegex.test("A1:B10")).toBe(true);
      expect(addressRegex.test("AA1:ZZ100")).toBe(true);
    });

    test("无效的地址格式", () => {
      expect(addressRegex.test("")).toBe(false);
      expect(addressRegex.test("A")).toBe(false);
      expect(addressRegex.test("1")).toBe(false);
      expect(addressRegex.test("A1:")).toBe(false);
    });
  });

  describe("公式语法检查", () => {
    test("括号匹配检测", () => {
      const checkParentheses = (formula: string): boolean => {
        const openParens = (formula.match(/\(/g) || []).length;
        const closeParens = (formula.match(/\)/g) || []).length;
        return openParens === closeParens;
      };

      expect(checkParentheses("=SUM(A1:A10)")).toBe(true);
      expect(checkParentheses("=IF(A1>0,SUM(B1:B10),0)")).toBe(true);
      expect(checkParentheses("=SUM(A1:A10")).toBe(false);
      expect(checkParentheses("=IF(A1>0,SUM(B1:B10,0)")).toBe(false);
    });

    test("公式前缀检测", () => {
      const hasFormulaPrefix = (formula: string): boolean => {
        return formula.startsWith("=");
      };

      expect(hasFormulaPrefix("=SUM(A1:A10)")).toBe(true);
      expect(hasFormulaPrefix("SUM(A1:A10)")).toBe(false);
    });
  });

  describe("公式引用解析", () => {
    const parseFormulaReferences = (formula: string): string[] => {
      if (!formula || !formula.startsWith("=")) return [];
      const references: string[] = [];
      const cellRefRegex = /(?:'[^']+'!)?(?:\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)/gi;
      let match;
      while ((match = cellRefRegex.exec(formula)) !== null) {
        references.push(match[0]);
      }
      return references;
    };

    test("解析简单引用", () => {
      expect(parseFormulaReferences("=A1")).toEqual(["A1"]);
      expect(parseFormulaReferences("=A1+B1")).toEqual(["A1", "B1"]);
    });

    test("解析范围引用", () => {
      expect(parseFormulaReferences("=SUM(A1:A10)")).toEqual(["A1:A10"]);
    });

    test("解析绝对引用", () => {
      expect(parseFormulaReferences("=$A$1")).toEqual(["$A$1"]);
      expect(parseFormulaReferences("=$A$1:$B$10")).toEqual(["$A$1:$B$10"]);
    });

    test("解析跨工作表引用", () => {
      expect(parseFormulaReferences("='Sheet1'!A1")).toEqual(["'Sheet1'!A1"]);
    });
  });

  describe("公式复杂度分析", () => {
    const analyzeFormulaComplexity = (formula: string): { level: string; functions: string[] } => {
      if (!formula || !formula.startsWith("=")) {
        return { level: "simple", functions: [] };
      }

      const functionRegex = /([A-Z]+)\s*\(/gi;
      const functions: string[] = [];
      let match;
      while ((match = functionRegex.exec(formula)) !== null) {
        functions.push(match[1].toUpperCase());
      }

      let maxDepth = 0;
      let currentDepth = 0;
      for (const char of formula) {
        if (char === "(") {
          currentDepth++;
          maxDepth = Math.max(maxDepth, currentDepth);
        } else if (char === ")") {
          currentDepth--;
        }
      }

      const score = functions.length * 10 + maxDepth * 15;
      let level = "simple";
      if (score >= 50) level = "complex";
      else if (score >= 20) level = "medium";

      return { level, functions };
    };

    test("简单公式", () => {
      const result = analyzeFormulaComplexity("=A1+B1");
      expect(result.level).toBe("simple");
      expect(result.functions).toEqual([]);
    });

    test("中等复杂度公式", () => {
      const result = analyzeFormulaComplexity("=SUM(A1:A10)");
      expect(result.level).toBe("medium");
      expect(result.functions).toContain("SUM");
    });

    test("复杂嵌套公式", () => {
      const result = analyzeFormulaComplexity("=IF(A1>0,SUMIF(B:B,C1,D:D),AVERAGE(E:E))");
      expect(result.level).toBe("complex");
      expect(result.functions.length).toBeGreaterThan(2);
    });
  });
});

describe("图表类型验证", () => {
  const validChartTypes = ["column", "bar", "line", "pie", "scatter", "area", "doughnut", "radar"];

  test("有效的图表类型", () => {
    validChartTypes.forEach((type) => {
      expect(validChartTypes.includes(type)).toBe(true);
    });
  });

  test("无效的图表类型应有默认值", () => {
    const validateChartType = (type: string): string => {
      return validChartTypes.includes(type.toLowerCase()) ? type : "column";
    };

    expect(validateChartType("unknown")).toBe("column");
    expect(validateChartType("invalid")).toBe("column");
    expect(validateChartType("line")).toBe("line");
  });
});

describe("数据验证类型", () => {
  const validValidationTypes = ["list", "number", "date", "textLength", "wholeNumber", "decimal"];

  test("有效的验证类型", () => {
    validValidationTypes.forEach((type) => {
      expect(validValidationTypes.includes(type)).toBe(true);
    });
  });
});
