/**
 * 公式验证引擎 - FormulaValidator
 *
 * 职责：
 * 1. 验证公式是否存在自引用（循环引用）
 * 2. 验证是否支持自动扩展（ARRAYFORMULA/动态数组）
 * 3. 平台特定规则（Excel vs Google Sheets）
 *
 * 永远不让模型验证自己，由系统规则判定
 */

import { Artifact, Platform, Validation, ValidationStatus } from "./types";

// ========== 验证规则 ==========

/**
 * 验证规则定义
 */
interface ValidationRule {
  id: string;
  name: string;
  description: string;
  platforms: Platform[];
  validate: (artifact: Artifact) => Validation;
}

// ========== FormulaValidator 类 ==========

/**
 * 公式验证器
 */
export class FormulaValidator {
  private rules: ValidationRule[] = [];

  constructor() {
    this.initializeRules();
  }

  /**
   * 初始化所有规则
   */
  private initializeRules(): void {
    // R2: 不允许结果列自引用
    this.rules.push({
      id: "R2_SELF_REFERENCE",
      name: "自引用检查",
      description: "不允许结果列自引用（循环引用高发）",
      platforms: [Platform.EXCEL, Platform.GOOGLE_SHEETS],
      validate: this.checkSelfReference.bind(this),
    });

    // R3: 必须具备自动扩展策略
    this.rules.push({
      id: "R3_AUTO_EXPAND",
      name: "自动扩展检查",
      description: "必须具备自动扩展策略（不拖拽）",
      platforms: [Platform.EXCEL, Platform.GOOGLE_SHEETS],
      validate: this.checkAutoExpand.bind(this),
    });

    // GS1: ARRAYFORMULA 输出列必须干净
    this.rules.push({
      id: "GS1_ARRAYFORMULA_OUTPUT",
      name: "ARRAYFORMULA 输出检查",
      description: "ARRAYFORMULA 输出列必须干净",
      platforms: [Platform.GOOGLE_SHEETS],
      validate: this.checkArrayFormulaOutput.bind(this),
    });

    // GS4: QUERY/ARRAYFORMULA 范围写法
    this.rules.push({
      id: "GS4_OPEN_RANGE",
      name: "开放区间检查",
      description: "推荐使用开放区间（如 A2:A）",
      platforms: [Platform.GOOGLE_SHEETS],
      validate: this.checkOpenRange.bind(this),
    });

    // XL1: Excel 不拖拽实现
    this.rules.push({
      id: "XL1_NO_DRAG",
      name: "Excel 不拖拽检查",
      description: "使用结构化引用或动态数组",
      platforms: [Platform.EXCEL],
      validate: this.checkExcelNoDrag.bind(this),
    });
  }

  /**
   * 验证单个产物
   */
  validate(artifact: Artifact): Validation[] {
    const results: Validation[] = [];

    for (const rule of this.rules) {
      // 只应用适用于当前平台的规则
      if (rule.platforms.includes(artifact.platform)) {
        results.push(rule.validate(artifact));
      }
    }

    return results;
  }

  /**
   * 验证所有产物
   */
  validateAll(artifacts: Artifact[]): Validation[] {
    const results: Validation[] = [];

    for (const artifact of artifacts) {
      results.push(...this.validate(artifact));
    }

    return results;
  }

  /**
   * R2: 检查自引用
   */
  private checkSelfReference(artifact: Artifact): Validation {
    const formula = artifact.content.toUpperCase();
    const target = artifact.target;

    // 提取公式所在列
    let outputColumn: string | null = null;

    if (target.column) {
      outputColumn = target.column.toUpperCase();
    } else if (target.cell) {
      // 从单元格地址提取列（如 A1 → A）
      const match = target.cell.match(/^([A-Za-z]+)/);
      if (match) {
        outputColumn = match[1].toUpperCase();
      }
    } else if (target.range) {
      // 从范围提取列（如 A2:A100 → A）
      const match = target.range.match(/^([A-Za-z]+)/);
      if (match) {
        outputColumn = match[1].toUpperCase();
      }
    }

    if (!outputColumn) {
      return {
        name: "自引用检查",
        ruleId: "R2_SELF_REFERENCE",
        status: ValidationStatus.WARN,
        reason: "无法确定输出列，请检查是否存在自引用",
      };
    }

    // 检查公式是否引用了输出列
    // 正则匹配列引用：A1, A2:A100, A:A, $A$1 等
    const columnRefPattern = new RegExp(
      `\\b${outputColumn}\\d*\\b|\\$${outputColumn}\\$?\\d*`,
      "g"
    );

    if (columnRefPattern.test(formula)) {
      return {
        name: "自引用检查",
        ruleId: "R2_SELF_REFERENCE",
        status: ValidationStatus.FAIL,
        reason: `公式引用了输出列 ${outputColumn}，存在循环引用风险`,
        details: { outputColumn, formula: artifact.content },
      };
    }

    return {
      name: "自引用检查",
      ruleId: "R2_SELF_REFERENCE",
      status: ValidationStatus.PASS,
    };
  }

  /**
   * R3: 检查自动扩展
   */
  private checkAutoExpand(artifact: Artifact): Validation {
    const formula = artifact.content.toUpperCase();

    if (artifact.platform === Platform.GOOGLE_SHEETS) {
      // Google Sheets: 检查是否使用 ARRAYFORMULA 或 QUERY
      if (formula.includes("ARRAYFORMULA") || formula.includes("QUERY")) {
        return {
          name: "自动扩展检查",
          ruleId: "R3_AUTO_EXPAND",
          status: ValidationStatus.PASS,
        };
      }
    } else if (artifact.platform === Platform.EXCEL) {
      // Excel: 检查是否使用结构化引用或动态数组函数
      const dynamicFunctions = [
        "FILTER",
        "SORT",
        "UNIQUE",
        "SORTBY",
        "XLOOKUP",
        "LET",
        "LAMBDA",
        "SEQUENCE",
        "RANDARRAY",
      ];

      const usesStructuredRef = /\[.*\]/.test(formula);
      const usesDynamicArray = dynamicFunctions.some((f) => formula.includes(f));

      if (usesStructuredRef || usesDynamicArray) {
        return {
          name: "自动扩展检查",
          ruleId: "R3_AUTO_EXPAND",
          status: ValidationStatus.PASS,
        };
      }
    }

    return {
      name: "自动扩展检查",
      ruleId: "R3_AUTO_EXPAND",
      status: ValidationStatus.WARN,
      reason: "公式可能不支持自动扩展，建议使用 ARRAYFORMULA（Sheets）或动态数组函数（Excel）",
    };
  }

  /**
   * GS1: 检查 ARRAYFORMULA 输出
   */
  private checkArrayFormulaOutput(artifact: Artifact): Validation {
    const formula = artifact.content.toUpperCase();

    if (!formula.includes("ARRAYFORMULA")) {
      return {
        name: "ARRAYFORMULA 输出检查",
        ruleId: "GS1_ARRAYFORMULA_OUTPUT",
        status: ValidationStatus.PASS,
        reason: "非 ARRAYFORMULA 公式",
      };
    }

    // 检查公式范围是否包含输出列
    const target = artifact.target;
    let outputColumn: string | null = null;

    if (target.column) {
      outputColumn = target.column.toUpperCase();
    } else if (target.cell) {
      const match = target.cell.match(/^([A-Za-z]+)/);
      if (match) {
        outputColumn = match[1].toUpperCase();
      }
    }

    if (outputColumn && formula.includes(outputColumn)) {
      return {
        name: "ARRAYFORMULA 输出检查",
        ruleId: "GS1_ARRAYFORMULA_OUTPUT",
        status: ValidationStatus.FAIL,
        reason: `ARRAYFORMULA 范围包含输出列 ${outputColumn}，可能导致循环引用`,
      };
    }

    return {
      name: "ARRAYFORMULA 输出检查",
      ruleId: "GS1_ARRAYFORMULA_OUTPUT",
      status: ValidationStatus.PASS,
    };
  }

  /**
   * GS4: 检查开放区间
   */
  private checkOpenRange(artifact: Artifact): Validation {
    const formula = artifact.content;

    // 检查是否使用硬编码行数（如 A2:A100）
    const hardcodedRangePattern = /[A-Z]+\d+:[A-Z]+(\d{2,})/gi;
    const matches = formula.match(hardcodedRangePattern);

    if (matches && matches.length > 0) {
      return {
        name: "开放区间检查",
        ruleId: "GS4_OPEN_RANGE",
        status: ValidationStatus.WARN,
        reason: `发现硬编码行数范围：${matches.join(", ")}。推荐使用开放区间（如 A2:A）`,
        details: { hardcodedRanges: matches },
      };
    }

    return {
      name: "开放区间检查",
      ruleId: "GS4_OPEN_RANGE",
      status: ValidationStatus.PASS,
    };
  }

  /**
   * XL1: 检查 Excel 不拖拽实现
   */
  private checkExcelNoDrag(artifact: Artifact): Validation {
    const formula = artifact.content.toUpperCase();

    // 检查是否使用结构化引用 [列名]
    const usesStructuredRef = /\[.*\]/.test(formula);

    // 检查是否使用动态数组函数
    const dynamicFunctions = ["FILTER", "SORT", "UNIQUE", "SORTBY", "XLOOKUP", "LET", "LAMBDA"];
    const usesDynamicArray = dynamicFunctions.some((f) => formula.includes(f));

    if (usesStructuredRef || usesDynamicArray) {
      return {
        name: "Excel 不拖拽检查",
        ruleId: "XL1_NO_DRAG",
        status: ValidationStatus.PASS,
      };
    }

    return {
      name: "Excel 不拖拽检查",
      ruleId: "XL1_NO_DRAG",
      status: ValidationStatus.WARN,
      reason: "公式可能需要拖拽填充。推荐使用 Excel 表格（结构化引用）或动态数组函数",
    };
  }
}

// ========== 导出单例 ==========

export const formulaValidator = new FormulaValidator();
