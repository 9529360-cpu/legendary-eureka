/**
 * FormulaValidator - 公式验证和错误检测引擎 v2.9.57
 *
 * 职责：
 * 1. 验证公式引用的有效性
 * 2. 执行前检测"写死行号"等致命模式
 * 3. 执行后检查 #VALUE!, #REF!, #NAME? 等错误
 * 4. 列级统计检测"整列同值"问题
 * 5. 提供错误诊断和修复建议
 *
 * v2.9.57 修复：
 * - 新增 validateFillFormulaPattern：拦截"写死行号"模式
 * - shouldRollback 使用比例而非绝对数
 * - totalCells 实际计算
 * - 抽样从 5 行提升到 50 行 + 列级 uniqueCount
 * - detectDistributionAnomaly 使用 IQR 方法
 */

// ========== 类型定义 ==========

/**
 * Excel 错误类型
 */
export type ExcelErrorType =
  | "#VALUE!"
  | "#REF!"
  | "#NAME?"
  | "#DIV/0!"
  | "#NULL!"
  | "#NUM!"
  | "#N/A"
  | "#GETTING_DATA"
  | "#SPILL!"
  | "#CALC!";

/**
 * v2.9.57: 公式填充模式风险等级
 */
export type FillPatternRisk = "safe" | "warning" | "critical";

/**
 * v2.9.57: 公式填充模式验证结果
 */
export interface FillPatternValidation {
  isValid: boolean;
  risk: FillPatternRisk;
  formula: string;
  targetRange: string;
  issues: FillPatternIssue[];
  suggestions: string[];
}

/**
 * v2.9.57: 公式填充模式问题
 */
export interface FillPatternIssue {
  type: "fixed_row_ref" | "absolute_row_only" | "hardcoded_value" | "cross_sheet_fixed";
  severity: "critical" | "warning";
  message: string;
  problematicPart: string;
  suggestedFix: string;
}

/**
 * 单元格检查结果
 */
export interface CellCheckResult {
  address: string;
  value: unknown;
  hasError: boolean;
  errorType?: ExcelErrorType;
  formula?: string;
}

/**
 * 范围检查结果 v2.9.57: totalCells 实际计算
 */
export interface RangeCheckResult {
  range: string;
  totalCells: number;
  errorCells: CellCheckResult[];
  hasErrors: boolean;
  errorSummary: Map<ExcelErrorType, number>;
  errorRate: number; // v2.9.57: 错误率
}

/**
 * 公式验证结果
 */
export interface FormulaValidationResult {
  isValid: boolean;
  formula: string;
  errors: FormulaError[];
  suggestions: string[];
}

/**
 * 公式错误
 */
export interface FormulaError {
  type: "syntax" | "reference" | "function" | "circular" | "type";
  message: string;
  position?: number;
}

/**
 * 执行后验证结果
 */
export interface PostExecutionValidation {
  success: boolean;
  sheet: string;
  range: string;
  operation: string;
  errors: ExecutionError[];
  shouldRollback: boolean;
  fixSuggestions: FixSuggestion[];
}

/**
 * 执行错误
 */
export interface ExecutionError {
  cell: string;
  errorType: ExcelErrorType;
  formula?: string;
  expectedType?: string;
  actualValue?: unknown;
}

/**
 * 修复建议
 */
export interface FixSuggestion {
  priority: "high" | "medium" | "low";
  description: string;
  action: string;
  details?: Record<string, unknown>;
}

// ========== v2.7/v2.9.57 抽样校验类型 ==========

/**
 * 抽样值
 */
export interface SampleValue {
  rowIndex: number;
  colIndex: number;
  value: unknown;
  formula?: string;
}

/**
 * v2.9.57: 列级统计结果
 */
export interface ColumnStats {
  columnIndex: number;
  columnLetter: string;
  uniqueCount: number;
  totalCount: number;
  uniqueRatio: number; // uniqueCount / totalCount
  topValues: Array<{ value: unknown; count: number }>; // 前3个最常见值
  hasFormulas: boolean;
  allSameValue: boolean;
  sameValue?: unknown;
}

/**
 * 抽样问题 v2.9.57: 新增 low_unique_ratio
 */
export interface SampleIssue {
  type:
    | "all_zeros"
    | "single_value"
    | "low_unique_ratio" // v2.9.57: 唯一值比例过低
    | "distribution_anomaly"
    | "type_mismatch"
    | "empty_data"
    | "read_error"
    | "sheet_not_found";
  severity: "critical" | "warning" | "info";
  message: string;
  affectedRange: string;
  details?: Record<string, unknown>;
}

/**
 * 抽样验证结果 v2.9.57: 增加列级统计
 */
export interface SampleValidationResult {
  isValid: boolean;
  sampledRows: number;
  issues: SampleIssue[];
  sampledValues: SampleValue[];
  columnStats?: ColumnStats[]; // v2.9.57: 列级统计
}

/**
 * 全零检测结果
 */
interface AllZeroCheckResult {
  detected: boolean;
  percentage: number;
  zeroCount?: number;
  totalCount?: number;
  [key: string]: unknown; // Index signature for Record compatibility
}

/**
 * 单一值检测结果
 */
interface SingleValueCheckResult {
  detected: boolean;
  value?: unknown;
  columnIndex?: number;
  sampleCount?: number;
  [key: string]: unknown;
}

/**
 * 分布异常检测结果
 */
interface DistributionAnomalyResult {
  hasAnomaly: boolean;
  message: string;
  min?: number;
  max?: number;
  mean?: number;
  outliers?: number[];
  [key: string]: unknown;
}

/**
 * 类型不匹配检测结果
 */
interface TypeMismatchResult {
  hasMismatch: boolean;
  message: string;
  columnIndex?: number;
  types?: string[];
  [key: string]: unknown;
}

// ========== 公式验证引擎 ==========

export class FormulaValidator {
  /**
   * v2.9.57: 验证公式填充模式（在写公式之前调用！）
   *
   * 核心功能：拦截"写死行号"导致整列同值的致命模式
   *
   * @param formula - 要填充的公式（如 "=B2*C2"）
   * @param targetRange - 目标范围（如 "D2:D100"）
   * @returns 验证结果，包含风险等级和修复建议
   */
  validateFillFormulaPattern(formula: string, targetRange: string): FillPatternValidation {
    const issues: FillPatternIssue[] = [];
    const suggestions: string[] = [];

    // 解析目标范围
    const rangeMatch = targetRange.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/i);
    if (!rangeMatch) {
      // 单格写入，不需要检查填充模式
      return {
        isValid: true,
        risk: "safe",
        formula,
        targetRange,
        issues: [],
        suggestions: [],
      };
    }

    const startRow = parseInt(rangeMatch[2]);
    const endRow = parseInt(rangeMatch[4]);
    const rowCount = endRow - startRow + 1;

    // 只有填充多行时才检查
    if (rowCount <= 1) {
      return {
        isValid: true,
        risk: "safe",
        formula,
        targetRange,
        issues: [],
        suggestions: [],
      };
    }

    // 提取公式中的所有单元格引用
    // 匹配模式: A1, $A1, A$1, $A$1, Sheet1!A1, 'Sheet Name'!A1
    const cellRefPattern = /(?:(?:'[^']+'|[A-Za-z0-9_]+)!)?\$?([A-Z]+)\$?(\d+)/gi;
    const refs: Array<{
      full: string;
      col: string;
      row: string;
      hasAbsRow: boolean;
      hasAbsCol: boolean;
    }> = [];

    let match;
    while ((match = cellRefPattern.exec(formula)) !== null) {
      const fullRef = match[0];
      const col = match[1];
      const row = match[2];

      // 检查是否有绝对引用符号
      const hasAbsCol = fullRef.includes("$" + col);
      const hasAbsRow = fullRef.includes("$" + row) || fullRef.includes(col + "$");

      refs.push({ full: fullRef, col, row, hasAbsRow, hasAbsCol });
    }

    // 检测问题模式
    for (const ref of refs) {
      const refRow = parseInt(ref.row);

      // 模式 1: 固定行号引用（如 B2、C2）在多行填充中
      // 如果行号等于起始行，且没有绝对行引用符号，这是相对引用，Excel 会自动调整 - 这是安全的
      // 但如果用 range.formulas = [[formula], [formula], ...] 方式写入，每格都是同样的文本，就不会调整！
      if (refRow === startRow && !ref.hasAbsRow) {
        // 这种情况取决于写入方式：
        // - 如果用 Excel 的自动填充，会自动调整
        // - 如果用 range.formulas = 二维数组方式，每格文本相同，不会调整
        issues.push({
          type: "fixed_row_ref",
          severity: "critical",
          message: `\u516C\u5F0F\u5F15\u7528 ${ref.full} \u5728\u586B\u5145 ${rowCount} \u884C\u65F6\u53EF\u80FD\u5BFC\u81F4\u6574\u5217\u90FD\u5F15\u7528\u7B2C ${refRow} \u884C`,
          problematicPart: ref.full,
          suggestedFix: `\u4F7F\u7528\u7ED3\u6784\u5316\u5F15\u7528 @[\u5217\u540D] \u6216\u884C\u6A21\u677F ${ref.col}{row}`,
        });
      }

      // 模式 2: 绝对行引用（如 B$2）- 可能是故意的，但需要警告
      if (ref.hasAbsRow && !ref.hasAbsCol) {
        issues.push({
          type: "absolute_row_only",
          severity: "warning",
          message: `\u7EDD\u5BF9\u884C\u5F15\u7528 ${ref.full} \u4F1A\u5BFC\u81F4\u6240\u6709\u884C\u90FD\u5F15\u7528\u540C\u4E00\u884C`,
          problematicPart: ref.full,
          suggestedFix: `\u5982\u679C\u8FD9\u662F\u6545\u610F\u7684\uFF08\u5982\u5F15\u7528\u8868\u5934\uFF09\uFF0C\u53EF\u4EE5\u5FFD\u7565\uFF1B\u5426\u5219\u8BF7\u4F7F\u7528\u76F8\u5BF9\u5F15\u7528`,
        });
      }
    }

    // 检测硬编码常量
    const formulaBody = formula.replace(/^=/, "");
    if (/^\d+(\.\d+)?$/.test(formulaBody)) {
      issues.push({
        type: "hardcoded_value",
        severity: "critical",
        message: `\u516C\u5F0F\u662F\u786C\u7F16\u7801\u5E38\u91CF ${formulaBody}\uFF0C\u586B\u5145 ${rowCount} \u884C\u5C06\u5168\u90E8\u76F8\u540C`,
        problematicPart: formulaBody,
        suggestedFix:
          "\u5982\u679C\u9700\u8981\u5E38\u91CF\uFF0C\u8BF7\u786E\u8BA4\u8FD9\u662F\u6545\u610F\u7684\uFF1B\u5426\u5219\u8BF7\u4F7F\u7528\u5355\u5143\u683C\u5F15\u7528",
      });
    }

    // 生成建议
    if (issues.some((i) => i.type === "fixed_row_ref")) {
      suggestions.push(
        "\u5EFA\u8BAE\u4F7F\u7528 excel_smart_formula \u5DE5\u5177\uFF0C\u5B83\u4F1A\u6309\u884C\u5C55\u5F00\u516C\u5F0F"
      );
      suggestions.push(
        "\u6216\u8005\u4F7F\u7528\u7ED3\u6784\u5316\u5F15\u7528\uFF1A=@[\u5355\u4EF7]*@[\u6570\u91CF]"
      );
      suggestions.push("\u6216\u8005\u4F7F\u7528\u884C\u6A21\u677F\uFF1A=B{row}*C{row}");
    }

    // 判断风险等级
    let risk: FillPatternRisk = "safe";
    if (issues.some((i) => i.severity === "critical")) {
      risk = "critical";
    } else if (issues.some((i) => i.severity === "warning")) {
      risk = "warning";
    }

    return {
      isValid: issues.filter((i) => i.severity === "critical").length === 0,
      risk,
      formula,
      targetRange,
      issues,
      suggestions,
    };
  }

  /**
   * 验证公式语法（执行前）
   */
  validateFormulaSyntax(formula: string): FormulaValidationResult {
    const errors: FormulaError[] = [];
    const suggestions: string[] = [];

    // 检查等号前缀
    if (!formula.startsWith("=")) {
      errors.push({
        type: "syntax",
        message: "公式必须以 = 开头",
        position: 0,
      });
      suggestions.push(`修复: =${formula}`);
    }

    // 检查括号匹配
    const openParens = (formula.match(/\(/g) || []).length;
    const closeParens = (formula.match(/\)/g) || []).length;
    if (openParens !== closeParens) {
      errors.push({
        type: "syntax",
        message: `括号不匹配: ${openParens} 个 ( 和 ${closeParens} 个 )`,
      });
      if (openParens > closeParens) {
        suggestions.push(`添加 ${openParens - closeParens} 个右括号`);
      } else {
        suggestions.push(`移除 ${closeParens - openParens} 个右括号`);
      }
    }

    // 检查引号匹配
    const quotes = (formula.match(/"/g) || []).length;
    if (quotes % 2 !== 0) {
      errors.push({
        type: "syntax",
        message: "引号不匹配",
      });
    }

    // 检查常见的函数名拼写
    const commonMisspellings: Record<string, string> = {
      SUMIF: "SUMIFS",
      COUNTIF: "COUNTIFS",
      AVERAGEIF: "AVERAGEIFS",
      VLOOK: "VLOOKUP",
      XLOOK: "XLOOKUP",
      IFERR: "IFERROR",
    };

    for (const [wrong, correct] of Object.entries(commonMisspellings)) {
      if (formula.includes(wrong) && !formula.includes(correct)) {
        suggestions.push(`检查函数名: 你是否想用 ${correct}?`);
      }
    }

    // 检查空的函数调用
    if (/\(\s*\)/.test(formula)) {
      errors.push({
        type: "syntax",
        message: "检测到空的函数调用 ()",
      });
    }

    return {
      isValid: errors.length === 0,
      formula,
      errors,
      suggestions,
    };
  }

  /**
   * 验证公式引用（需要工作簿上下文）
   */
  validateFormulaReferences(
    formula: string,
    currentSheet: string,
    workbookContext: WorkbookContext
  ): FormulaValidationResult {
    const errors: FormulaError[] = [];
    const suggestions: string[] = [];

    // 提取跨表引用
    const crossSheetRefs = this.extractCrossSheetReferences(formula);

    for (const ref of crossSheetRefs) {
      // 检查工作表是否存在
      const sheetExists = workbookContext.sheets.some(
        (s) => s.name.toLowerCase() === ref.sheet.toLowerCase()
      );

      if (!sheetExists) {
        errors.push({
          type: "reference",
          message: `工作表 "${ref.sheet}" 不存在`,
        });

        // 建议相似的工作表名
        const similarSheets = workbookContext.sheets
          .map((s) => s.name)
          .filter((name) => this.isSimilar(name, ref.sheet));

        if (similarSheets.length > 0) {
          suggestions.push(`你是否想引用: ${similarSheets.join(", ")}?`);
        }
      }
    }

    // 检查本表引用
    const localRefs = this.extractLocalReferences(formula);
    const currentSheetInfo = workbookContext.sheets.find((s) => s.name === currentSheet);

    if (currentSheetInfo) {
      for (const ref of localRefs) {
        const colIndex = this.columnToIndex(ref.column);
        const rowIndex = ref.row;

        // 检查是否超出已使用范围（可能是错误引用）
        if (
          colIndex > currentSheetInfo.columnCount + 10 ||
          rowIndex > currentSheetInfo.rowCount + 1000
        ) {
          suggestions.push(
            `引用 ${ref.column}${ref.row} 可能超出数据范围（当前表有 ${currentSheetInfo.columnCount} 列，${currentSheetInfo.rowCount} 行）`
          );
        }
      }
    }

    return {
      isValid: errors.length === 0,
      formula,
      errors,
      suggestions,
    };
  }

  /**
   * 检查执行后的结果（关键！）
   */
  async checkPostExecution(
    sheet: string,
    range: string,
    operation: string
  ): Promise<PostExecutionValidation> {
    const errors: ExecutionError[] = [];
    const fixSuggestions: FixSuggestion[] = [];

    try {
      // 读取执行后的单元格值
      const checkResult = await this.readRangeForErrors(sheet, range);

      for (const cell of checkResult.errorCells) {
        errors.push({
          cell: cell.address,
          errorType: cell.errorType!,
          formula: cell.formula,
        });

        // 根据错误类型生成修复建议
        const fix = this.generateFixSuggestion(cell);
        if (fix) {
          fixSuggestions.push(fix);
        }
      }

      // v2.9.57: 传递 totalCells 以使用比例判断
      const shouldRollback = this.shouldRollback(errors, operation, checkResult.totalCells);

      return {
        success: errors.length === 0,
        sheet,
        range,
        operation,
        errors,
        shouldRollback,
        fixSuggestions,
      };
    } catch (error) {
      return {
        success: false,
        sheet,
        range,
        operation,
        errors: [
          {
            cell: range,
            errorType: "#VALUE!",
            formula: undefined,
          },
        ],
        shouldRollback: true,
        fixSuggestions: [
          {
            priority: "high",
            description: "无法读取执行结果",
            action: "retry",
            details: { error: error instanceof Error ? error.message : String(error) },
          },
        ],
      };
    }
  }

  /**
   * 读取范围内的错误
   * v2.9.57: 修复 totalCells 和 errorRate 计算
   */
  private async readRangeForErrors(sheet: string, range: string): Promise<RangeCheckResult> {
    const errorCells: CellCheckResult[] = [];
    const errorSummary = new Map<ExcelErrorType, number>();
    let totalCells = 0;

    try {
      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getItem(sheet);
        const targetRange = worksheet.getRange(range);

        targetRange.load(["values", "formulas", "address"]);
        await context.sync();

        const values = targetRange.values;
        const formulas = targetRange.formulas;

        // v2.9.57: 正确计算总单元格数
        totalCells = values.length * (values[0]?.length || 0);

        // 解析起始地址
        const startMatch = range.match(/([A-Z]+)(\d+)/);
        const startCol = startMatch ? this.columnToIndex(startMatch[1]) : 0;
        const startRow = startMatch ? parseInt(startMatch[2]) : 1;

        for (let row = 0; row < values.length; row++) {
          for (let col = 0; col < values[row].length; col++) {
            const value = values[row][col];
            const formula = formulas[row][col];

            if (this.isExcelError(value)) {
              const errorType = value as ExcelErrorType;
              const cellAddress = this.indexToColumn(startCol + col) + (startRow + row);

              errorCells.push({
                address: cellAddress,
                value,
                hasError: true,
                errorType,
                formula:
                  typeof formula === "string" && formula.startsWith("=") ? formula : undefined,
              });

              errorSummary.set(errorType, (errorSummary.get(errorType) || 0) + 1);
            }
          }
        }
      });
    } catch (error) {
      console.error("[FormulaValidator] Error reading range:", error);
    }

    // v2.9.57: 计算错误率
    const errorRate = totalCells > 0 ? errorCells.length / totalCells : 0;

    return {
      range,
      totalCells,
      errorCells,
      hasErrors: errorCells.length > 0,
      errorSummary,
      errorRate,
    };
  }

  /**
   * 判断值是否是 Excel 错误
   */
  private isExcelError(value: unknown): boolean {
    if (typeof value !== "string") return false;

    const errorPatterns = [
      "#VALUE!",
      "#REF!",
      "#NAME?",
      "#DIV/0!",
      "#NULL!",
      "#NUM!",
      "#N/A",
      "#GETTING_DATA",
      "#SPILL!",
      "#CALC!",
    ];

    return errorPatterns.includes(value);
  }

  /**
   * 生成修复建议
   */
  private generateFixSuggestion(cell: CellCheckResult): FixSuggestion | null {
    switch (cell.errorType) {
      case "#VALUE!":
        return {
          priority: "high",
          description: `单元格 ${cell.address} 出现 #VALUE! 错误`,
          action: "check_data_types",
          details: {
            message: "数据类型不匹配，检查公式中引用的单元格是否包含正确的数据类型",
            formula: cell.formula,
            suggestions: [
              "确保数值运算的输入是数字",
              "检查是否有空单元格被用于计算",
              "检查引用的数据源是否存在",
            ],
          },
        };

      case "#REF!":
        return {
          priority: "high",
          description: `单元格 ${cell.address} 出现 #REF! 错误`,
          action: "fix_reference",
          details: {
            message: "无效的单元格引用，可能引用了被删除的单元格或工作表",
            formula: cell.formula,
          },
        };

      case "#NAME?":
        return {
          priority: "high",
          description: `单元格 ${cell.address} 出现 #NAME? 错误`,
          action: "fix_function_name",
          details: {
            message: "无法识别的函数名或命名范围",
            formula: cell.formula,
            suggestions: ["检查函数名拼写是否正确", "检查命名范围是否存在", "文本需要用引号包裹"],
          },
        };

      case "#DIV/0!":
        return {
          priority: "medium",
          description: `单元格 ${cell.address} 出现 #DIV/0! 错误`,
          action: "add_zero_check",
          details: {
            message: "除数为零",
            formula: cell.formula,
            suggestion: "使用 IFERROR 或 IF 函数处理除数为零的情况",
          },
        };

      case "#N/A":
        return {
          priority: "medium",
          description: `单元格 ${cell.address} 出现 #N/A 错误`,
          action: "check_lookup",
          details: {
            message: "查找函数未找到匹配值",
            formula: cell.formula,
            suggestions: [
              "检查查找值是否存在于数据源中",
              '使用 IFERROR(XLOOKUP(...), "") 处理未找到的情况',
            ],
          },
        };

      default:
        return null;
    }
  }

  /**
   * 判断是否需要回滚（公开方法，供外部调用）
   * v2.9.57: 使用错误比例而非绝对数量，避免大范围操作误判
   *
   * @param errors - 错误列表
   * @param operation - 操作类型
   * @param totalCells - 总单元格数（可选，用于计算比例）
   */
  shouldRollback(
    errors: ExecutionError[],
    operation: string = "unknown",
    totalCells?: number
  ): boolean {
    // 如果有 #REF! 错误，通常意味着结构性问题，需要回滚
    if (errors.some((e) => e.errorType === "#REF!")) {
      return true;
    }

    // v2.9.57: 基于错误比例判断
    // 如果提供了总数，用比例；否则用兼容逻辑
    if (totalCells && totalCells > 0) {
      const errorRate = errors.length / totalCells;
      // 超过 10% 的单元格出错，需要回滚
      if (errorRate > 0.1) {
        return true;
      }
    } else {
      // 兼容旧调用：如果大量单元格出错（超过 50 个），需要回滚
      // 注意：这是降级逻辑，新代码应该传 totalCells
      if (errors.length > 50) {
        return true;
      }
    }

    // 公式设置操作如果出现 #VALUE!，需要回滚
    // v2.9.57: 修正操作名匹配（excel_set_formula 而不是 set_formula）
    const formulaOperations = ["set_formula", "excel_set_formula", "excel_smart_formula"];
    if (formulaOperations.includes(operation) && errors.some((e) => e.errorType === "#VALUE!")) {
      return true;
    }

    return false;
  }

  // ========== v2.7 硬约束: 数值抽样校验 ==========

  /**
   * 抽样验证结果 - 程序层强制检测
   */
  async sampleValidation(
    sheet: string,
    range: string,
    sampleSize: number = 50 // v2.9.57: 从 5 增加到 50
  ): Promise<SampleValidationResult> {
    const issues: SampleIssue[] = [];
    let isValid = true;
    const sampledValues: SampleValue[] = [];
    const columnStats: ColumnStats[] = []; // v2.9.57: 列级统计

    try {
      await Excel.run(async (context) => {
        // v2.7.3: 先检查工作表是否存在，避免 ItemNotFound 错误
        const worksheet = context.workbook.worksheets.getItemOrNullObject(sheet);
        await context.sync();

        if (worksheet.isNullObject) {
          issues.push({
            type: "sheet_not_found",
            severity: "warning",
            message: `工作表 "${sheet}" 不存在，跳过校验`,
            affectedRange: `${sheet}!${range}`,
          });
          // 工作表不存在不算验证失败，只是跳过
          return;
        }

        const targetRange = worksheet.getRange(range);

        targetRange.load(["values", "formulas", "rowCount", "columnCount"]);
        await context.sync();

        const values = targetRange.values;
        const formulas = targetRange.formulas;
        const totalRows = values.length;
        const totalCols = values[0]?.length || 0;

        if (totalRows === 0) {
          issues.push({
            type: "empty_data",
            severity: "critical",
            message: "目标范围没有数据",
            affectedRange: range,
          });
          isValid = false;
          return;
        }

        // v2.9.57: 计算列级统计（在采样之前，用全量数据）
        for (let col = 0; col < totalCols; col++) {
          const colValues = values.map((row) => row[col]);
          const colFormulas = formulas.map((row) => row[col]);
          const stats = this.calculateColumnStats(colValues, col, colFormulas);
          columnStats.push(stats);

          // 检测低唯一率问题
          if (stats.allSameValue && totalRows > 1) {
            issues.push({
              type: "low_unique_ratio",
              severity: stats.hasFormulas ? "critical" : "warning",
              message: `第 ${col + 1} 列所有 ${totalRows} 行值相同: "${stats.topValues[0]?.value}"${stats.hasFormulas ? " (可能公式固定引用导致)" : ""}`,
              affectedRange: range,
              details: {
                columnIndex: col,
                uniqueCount: stats.uniqueCount,
                uniqueRatio: stats.uniqueRatio,
                topValue: stats.topValues[0],
              },
            });

            if (stats.hasFormulas) {
              isValid = false;
            }
          }
        }

        // 抽样: 头部 + 中间 + 尾部
        const sampleIndices = this.getSampleIndices(totalRows, sampleSize);

        for (const idx of sampleIndices) {
          if (idx >= values.length) continue;

          const row = values[idx];
          for (let col = 0; col < row.length; col++) {
            const value = row[col];
            const formula = formulas[idx]?.[col];

            sampledValues.push({
              rowIndex: idx,
              colIndex: col,
              value,
              formula: typeof formula === "string" && formula.startsWith("=") ? formula : undefined,
            });
          }
        }

        // 1. 检测全零问题
        const allZeroCheck = this.detectAllZeros(sampledValues);
        if (allZeroCheck.detected) {
          issues.push({
            type: "all_zeros",
            severity: "warning",
            message: `检测到全零数据 (${allZeroCheck.percentage.toFixed(0)}% 的抽样值为0)`,
            affectedRange: range,
            details: allZeroCheck,
          });
          // 100% 全零是critical
          if (allZeroCheck.percentage >= 100) {
            isValid = false;
            issues[issues.length - 1].severity = "critical";
          }
        }

        // 2. 检测单一值问题（已在列级统计中检测，这里是抽样补充）
        const singleValueCheck = this.detectSingleValue(sampledValues);
        if (singleValueCheck.detected) {
          // 如果列统计没有检测到，才添加
          if (!issues.some((i) => i.type === "low_unique_ratio")) {
            issues.push({
              type: "single_value",
              severity: "warning",
              message: `检测到所有值相同: "${singleValueCheck.value}"`,
              affectedRange: range,
              details: singleValueCheck,
            });
            // 如果是公式列但所有值相同，可能是公式写成了常量
            if (sampledValues.some((v) => v.formula)) {
              issues[issues.length - 1].severity = "critical";
              issues[issues.length - 1].message += " (公式可能被误写为常量)";
              isValid = false;
            }
          }
        }

        // 3. 检测异常分布
        const distributionCheck = this.detectDistributionAnomaly(sampledValues);
        if (distributionCheck.hasAnomaly) {
          issues.push({
            type: "distribution_anomaly",
            severity: "warning",
            message: distributionCheck.message,
            affectedRange: range,
            details: distributionCheck,
          });
        }

        // 4. 检测数据类型一致性
        const typeCheck = this.detectTypeMismatch(sampledValues);
        if (typeCheck.hasMismatch) {
          issues.push({
            type: "type_mismatch",
            severity: "warning",
            message: typeCheck.message,
            affectedRange: range,
            details: typeCheck,
          });
        }
      });
    } catch (error) {
      issues.push({
        type: "read_error",
        severity: "critical",
        message: `无法读取范围 ${range}: ${error instanceof Error ? error.message : String(error)}`,
        affectedRange: range,
      });
      isValid = false;
    }

    return {
      isValid,
      sampledRows: sampledValues.length,
      issues,
      sampledValues,
      columnStats, // v2.9.57: 返回列级统计
    };
  }

  /**
   * v2.9.57: 计算单列统计信息
   */
  private calculateColumnStats(
    values: unknown[],
    columnIndex: number,
    formulas?: unknown[]
  ): ColumnStats {
    const valueCount = new Map<string, number>();

    for (const v of values) {
      const key = String(v);
      valueCount.set(key, (valueCount.get(key) || 0) + 1);
    }

    // 按出现次数降序排列
    const sorted = Array.from(valueCount.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5);

    const uniqueCount = valueCount.size;
    const totalCount = values.length;
    const uniqueRatio = totalCount > 0 ? uniqueCount / totalCount : 0;
    const allSameValue = uniqueCount === 1 && totalCount > 1;

    // 检查是否有公式
    const hasFormulas = formulas
      ? formulas.some((f) => typeof f === "string" && f.startsWith("="))
      : false;

    return {
      columnIndex,
      columnLetter: this.indexToColumn(columnIndex),
      uniqueCount,
      totalCount,
      uniqueRatio,
      topValues: sorted.map(([value, count]) => ({ value, count })),
      hasFormulas,
      allSameValue,
      sameValue: allSameValue ? sorted[0]?.[0] : undefined,
    };
  }

  /**
   * 获取抽样索引 (头部 + 中间 + 尾部)
   */
  private getSampleIndices(totalRows: number, sampleSize: number): number[] {
    if (totalRows <= sampleSize) {
      return Array.from({ length: totalRows }, (_, i) => i);
    }

    const indices: number[] = [];

    // 头部 (约40%)
    const headCount = Math.ceil(sampleSize * 0.4);
    for (let i = 0; i < headCount && i < totalRows; i++) {
      indices.push(i);
    }

    // 中间 (约30%)
    const midCount = Math.floor(sampleSize * 0.3);
    const midStart = Math.floor(totalRows / 2) - Math.floor(midCount / 2);
    for (let i = 0; i < midCount; i++) {
      const idx = midStart + i;
      if (idx >= 0 && idx < totalRows && !indices.includes(idx)) {
        indices.push(idx);
      }
    }

    // 尾部 (约30%)
    const tailCount = sampleSize - indices.length;
    for (let i = 0; i < tailCount; i++) {
      const idx = totalRows - 1 - i;
      if (idx >= 0 && !indices.includes(idx)) {
        indices.push(idx);
      }
    }

    return indices.sort((a, b) => a - b);
  }

  /**
   * 检测全零问题
   */
  private detectAllZeros(samples: SampleValue[]): AllZeroCheckResult {
    const numericSamples = samples.filter((s) => typeof s.value === "number");
    if (numericSamples.length === 0) {
      return { detected: false, percentage: 0 };
    }

    const zeroCount = numericSamples.filter((s) => s.value === 0).length;
    const percentage = (zeroCount / numericSamples.length) * 100;

    return {
      detected: percentage >= 80, // 80% 以上为0视为异常
      percentage,
      zeroCount,
      totalCount: numericSamples.length,
    };
  }

  /**
   * 检测单一值问题
   */
  private detectSingleValue(samples: SampleValue[]): SingleValueCheckResult {
    if (samples.length < 2) {
      return { detected: false };
    }

    // 按列分组检查
    const byColumn = new Map<number, unknown[]>();
    for (const s of samples) {
      if (!byColumn.has(s.colIndex)) {
        byColumn.set(s.colIndex, []);
      }
      byColumn.get(s.colIndex)!.push(s.value);
    }

    // 检查是否有列的所有值相同
    for (const [colIdx, values] of byColumn) {
      if (values.length < 2) continue;

      const firstValue = values[0];
      const allSame = values.every((v) => v === firstValue);

      if (allSame && firstValue !== "" && firstValue !== null) {
        return {
          detected: true,
          value: firstValue,
          columnIndex: colIdx,
          sampleCount: values.length,
        };
      }
    }

    return { detected: false };
  }

  /**
   * 检测分布异常
   * v2.9.57: 使用 IQR (四分位距) 方法检测离群值，更稳健
   */
  private detectDistributionAnomaly(samples: SampleValue[]): DistributionAnomalyResult {
    const numericSamples = samples.filter((s) => typeof s.value === "number") as Array<
      SampleValue & { value: number }
    >;

    if (numericSamples.length < 5) {
      // v2.9.57: 至少需要 5 个样本才能计算 IQR
      return { hasAnomaly: false, message: "" };
    }

    const values = numericSamples.map((s) => s.value).sort((a, b) => a - b);
    const n = values.length;

    const min = values[0];
    const max = values[n - 1];
    const mean = values.reduce((a, b) => a + b, 0) / n;

    // v2.9.57: 使用 IQR 方法检测离群值
    // Q1 = 25th percentile, Q3 = 75th percentile
    const q1Index = Math.floor(n * 0.25);
    const q3Index = Math.floor(n * 0.75);
    const q1 = values[q1Index];
    const q3 = values[q3Index];
    const iqr = q3 - q1;

    // 离群值定义: < Q1 - 1.5*IQR 或 > Q3 + 1.5*IQR
    const lowerBound = q1 - 1.5 * iqr;
    const upperBound = q3 + 1.5 * iqr;

    const outliers = values.filter((v) => v < lowerBound || v > upperBound);

    if (outliers.length > 0) {
      const outlierRate = outliers.length / n;
      // 超过 10% 的数据是离群值才报警
      if (outlierRate > 0.1) {
        return {
          hasAnomaly: true,
          message: `检测到 ${outliers.length} 个离群值 (${(outlierRate * 100).toFixed(1)}%)，范围: [${lowerBound.toFixed(2)}, ${upperBound.toFixed(2)}]`,
          min,
          max,
          mean,
          outliers,
          // v2.9.57: 添加 IQR 统计信息
          iqrStats: { q1, q3, iqr, lowerBound, upperBound },
        };
      }
    }

    // 检测负数（如果预期应该是正数的场景）
    const negatives = values.filter((v) => v < 0);
    const positives = values.filter((v) => v > 0);
    if (negatives.length > 0 && positives.length > 0) {
      // 混合正负数可能需要注意
      // v2.9.57: 只有当正负比例悬殊时才报警（1:10 或 10:1）
      const ratio = Math.max(
        positives.length / negatives.length,
        negatives.length / positives.length
      );
      if (ratio >= 10) {
        return {
          hasAnomaly: true,
          message: `数据包含正负混合值，比例悬殊 (${positives.length} 正, ${negatives.length} 负)`,
          min,
          max,
          mean,
        };
      }
    }

    return { hasAnomaly: false, message: "" };
  }

  /**
   * 检测数据类型不一致
   */
  private detectTypeMismatch(samples: SampleValue[]): TypeMismatchResult {
    if (samples.length < 2) {
      return { hasMismatch: false, message: "" };
    }

    // 按列分组
    const byColumn = new Map<number, Set<string>>();
    for (const s of samples) {
      if (!byColumn.has(s.colIndex)) {
        byColumn.set(s.colIndex, new Set());
      }
      const type = s.value === null ? "null" : s.value === "" ? "empty" : typeof s.value;
      byColumn.get(s.colIndex)!.add(type);
    }

    // 检查是否有列混合了不同类型
    for (const [colIdx, types] of byColumn) {
      // 移除空值类型
      types.delete("null");
      types.delete("empty");

      if (types.size > 1) {
        return {
          hasMismatch: true,
          message: `列 ${colIdx + 1} 包含混合数据类型: ${Array.from(types).join(", ")}`,
          columnIndex: colIdx,
          types: Array.from(types),
        };
      }
    }

    return { hasMismatch: false, message: "" };
  }

  /**
   * 提取跨表引用
   */
  private extractCrossSheetReferences(formula: string): Array<{ sheet: string; range: string }> {
    const refs: Array<{ sheet: string; range: string }> = [];

    // 匹配 'Sheet Name'!A1:B2 或 Sheet1!A1
    const pattern = /'?([^'!]+)'?!([A-Z]+\d+(?::[A-Z]+\d+)?)/g;
    let match;

    while ((match = pattern.exec(formula)) !== null) {
      refs.push({
        sheet: match[1],
        range: match[2],
      });
    }

    return refs;
  }

  /**
   * 提取本表引用
   */
  private extractLocalReferences(formula: string): Array<{ column: string; row: number }> {
    const refs: Array<{ column: string; row: number }> = [];

    // 匹配 A1, B2 等（排除跨表引用）
    const pattern = /(?<![A-Z!'])([A-Z]+)(\d+)/g;
    let match;

    while ((match = pattern.exec(formula)) !== null) {
      // 确保不是函数名的一部分
      const precedingChar = formula[match.index - 1];
      if (precedingChar && /[A-Z]/.test(precedingChar)) {
        continue;
      }

      refs.push({
        column: match[1],
        row: parseInt(match[2]),
      });
    }

    return refs;
  }

  /**
   * 列字母转索引
   */
  private columnToIndex(column: string): number {
    let index = 0;
    for (let i = 0; i < column.length; i++) {
      index = index * 26 + (column.charCodeAt(i) - 64);
    }
    return index;
  }

  /**
   * 索引转列字母
   */
  private indexToColumn(index: number): string {
    let column = "";
    while (index > 0) {
      const remainder = (index - 1) % 26;
      column = String.fromCharCode(65 + remainder) + column;
      index = Math.floor((index - 1) / 26);
    }
    return column || "A";
  }

  /**
   * 检查两个字符串是否相似
   */
  private isSimilar(a: string, b: string): boolean {
    const la = a.toLowerCase();
    const lb = b.toLowerCase();

    // 完全匹配
    if (la === lb) return true;

    // 包含关系
    if (la.includes(lb) || lb.includes(la)) return true;

    // Levenshtein 距离
    if (this.levenshteinDistance(la, lb) <= 2) return true;

    return false;
  }

  /**
   * v2.8.0 新增: 自动修复公式
   * 尝试自动修复常见的公式问题
   */
  autoFixFormula(formula: string, errorType: ExcelErrorType): AutoFixResult {
    const result: AutoFixResult = {
      success: false,
      originalFormula: formula,
      fixedFormula: formula,
      fixApplied: [],
    };

    switch (errorType) {
      case "#DIV/0!":
        // 用 IFERROR 包装除法
        if (formula.includes("/")) {
          result.fixedFormula = `=IFERROR(${formula.substring(1)}, 0)`;
          result.success = true;
          result.fixApplied.push("添加 IFERROR 处理除零错误");
        }
        break;

      case "#N/A":
        // 用 IFERROR 包装查找函数
        if (/XLOOKUP|VLOOKUP|HLOOKUP|INDEX|MATCH/i.test(formula)) {
          result.fixedFormula = `=IFERROR(${formula.substring(1)}, "")`;
          result.success = true;
          result.fixApplied.push("添加 IFERROR 处理查找失败");
        }
        break;

      case "#VALUE!":
        // 尝试添加类型转换
        if (!formula.includes("VALUE(") && !formula.includes("TEXT(")) {
          // 无法自动修复，但提供建议
          result.fixApplied.push("建议: 检查数据类型，确保数值运算使用数字");
        }
        break;

      case "#NAME?": {
        // 尝试修复常见的函数名拼写错误
        const fixedFormula = this.fixFunctionNames(formula);
        if (fixedFormula !== formula) {
          result.fixedFormula = fixedFormula;
          result.success = true;
          result.fixApplied.push("修正函数名拼写");
        }
        break;
      }

      default:
        break;
    }

    return result;
  }

  /**
   * 修复常见的函数名拼写错误
   */
  private fixFunctionNames(formula: string): string {
    const corrections: Record<string, string> = {
      SUMIF: "SUMIF",
      SUMIFS: "SUMIFS",
      COUNTIF: "COUNTIF",
      COUNTIFS: "COUNTIFS",
      VLOOKUP: "VLOOKUP",
      HLOOKUP: "HLOOKUP",
      XLOOKUP: "XLOOKUP",
      IFERROR: "IFERROR",
      AVERAGE: "AVERAGE",
      AVERAGEIF: "AVERAGEIF",
      CONCATENATE: "CONCATENATE",
      TEXTJOIN: "TEXTJOIN",
      // 常见中文输入法导致的错误
      "SUM（": "SUM(",
      "IF（": "IF(",
      "VLOOKUP（": "VLOOKUP(",
    };

    let fixed = formula;
    for (const [wrong, correct] of Object.entries(corrections)) {
      const regex = new RegExp(wrong.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "gi");
      fixed = fixed.replace(regex, correct);
    }

    // 修复中文括号
    fixed = fixed.replace(/（/g, "(").replace(/）/g, ")");
    fixed = fixed.replace(/，/g, ",");

    return fixed;
  }

  /**
   * v2.8.0 新增: 智能公式建议
   * 根据数据特征和用户意图推荐公式
   */
  suggestFormula(intent: FormulaIntent, context: FormulaContext): FormulaSuggestion[] {
    const suggestions: FormulaSuggestion[] = [];

    switch (intent.type) {
      case "sum":
        suggestions.push({
          formula: `=SUM(${context.sourceRange})`,
          description: "求和",
          confidence: 0.95,
        });
        if (context.conditionColumn) {
          suggestions.push({
            formula: `=SUMIF(${context.conditionColumn}, "${context.conditionValue || "*"}", ${context.sourceRange})`,
            description: "条件求和",
            confidence: 0.85,
          });
        }
        break;

      case "lookup":
        suggestions.push({
          formula: `=XLOOKUP(${context.lookupValue}, ${context.lookupRange}, ${context.returnRange}, "", 0)`,
          description: "XLOOKUP 精确查找",
          confidence: 0.9,
        });
        suggestions.push({
          formula: `=IFERROR(XLOOKUP(${context.lookupValue}, ${context.lookupRange}, ${context.returnRange}), "未找到")`,
          description: "XLOOKUP 带错误处理",
          confidence: 0.85,
        });
        break;

      case "count":
        suggestions.push({
          formula: `=COUNTA(${context.sourceRange})`,
          description: "计数非空单元格",
          confidence: 0.9,
        });
        if (context.conditionColumn) {
          suggestions.push({
            formula: `=COUNTIF(${context.sourceRange}, "${context.conditionValue || '<>""'}")`,
            description: "条件计数",
            confidence: 0.85,
          });
        }
        break;

      case "percentage":
        suggestions.push({
          formula: `=${context.numerator}/${context.denominator}`,
          description: "百分比计算",
          confidence: 0.9,
        });
        suggestions.push({
          formula: `=IFERROR(${context.numerator}/${context.denominator}, 0)`,
          description: "百分比计算（防除零）",
          confidence: 0.95,
        });
        break;

      case "date":
        suggestions.push({
          formula: `=TODAY()`,
          description: "当前日期",
          confidence: 0.9,
        });
        suggestions.push({
          formula: `=EOMONTH(${context.dateCell || "TODAY()"}, 0)`,
          description: "月末日期",
          confidence: 0.8,
        });
        break;

      case "text":
        suggestions.push({
          formula: `=TEXTJOIN(", ", TRUE, ${context.sourceRange})`,
          description: "合并文本",
          confidence: 0.85,
        });
        suggestions.push({
          formula: `=LEFT(${context.sourceCell}, ${context.charCount || 5})`,
          description: "提取左侧字符",
          confidence: 0.8,
        });
        break;
    }

    return suggestions.sort((a, b) => b.confidence - a.confidence);
  }

  /**
   * 计算 Levenshtein 距离
   */
  private levenshteinDistance(a: string, b: string): number {
    const matrix: number[][] = [];

    for (let i = 0; i <= b.length; i++) {
      matrix[i] = [i];
    }

    for (let j = 0; j <= a.length; j++) {
      matrix[0][j] = j;
    }

    for (let i = 1; i <= b.length; i++) {
      for (let j = 1; j <= a.length; j++) {
        if (b.charAt(i - 1) === a.charAt(j - 1)) {
          matrix[i][j] = matrix[i - 1][j - 1];
        } else {
          matrix[i][j] = Math.min(
            matrix[i - 1][j - 1] + 1,
            matrix[i][j - 1] + 1,
            matrix[i - 1][j] + 1
          );
        }
      }
    }

    return matrix[b.length][a.length];
  }
}

// ========== v2.8.0 新增类型 ==========

/**
 * 自动修复结果
 */
export interface AutoFixResult {
  success: boolean;
  originalFormula: string;
  fixedFormula: string;
  fixApplied: string[];
}

/**
 * 公式意图
 */
export interface FormulaIntent {
  type: "sum" | "lookup" | "count" | "percentage" | "date" | "text" | "custom";
  description?: string;
}

/**
 * 公式上下文
 */
export interface FormulaContext {
  sourceRange?: string;
  sourceCell?: string;
  lookupValue?: string;
  lookupRange?: string;
  returnRange?: string;
  conditionColumn?: string;
  conditionValue?: string;
  numerator?: string;
  denominator?: string;
  dateCell?: string;
  charCount?: number;
}

/**
 * 公式建议
 */
export interface FormulaSuggestion {
  formula: string;
  description: string;
  confidence: number;
}

// ========== v2.8.1 数据建模验证 ==========

/**
 * 表类型定义
 */
export type TableType = "master" | "transaction" | "summary" | "analysis" | "unknown";

/**
 * 数据建模问题类型
 */
export type DataModelingIssueType =
  | "hardcoded_value"
  | "missing_lookup"
  | "missing_aggregation"
  | "duplicate_data"
  | "inconsistent_data"
  | "missing_formula"
  | "circular_reference"
  | "wrong_table_order"
  | "missing_master_reference";

/**
 * 数据建模问题
 */
export interface DataModelingIssue {
  type: DataModelingIssueType;
  severity: "critical" | "warning" | "info";
  location: string;
  message: string;
  suggestion: string;
  fixAction?: FixAction;
}

/**
 * 自动修复动作
 */
export interface FixAction {
  action: "set_formula" | "delete_and_recreate" | "fill_formula" | "add_lookup";
  target: string; // 目标范围
  formula?: string; // 要设置的公式
  parameters?: Record<string, unknown>;
}

/**
 * 表关系定义
 */
export interface TableRelation {
  sourceTable: string;
  targetTable: string;
  relationshipType: "one_to_many" | "many_to_one" | "lookup";
  keyColumn: string;
  lookupColumn?: string;
}

/**
 * 表类型识别结果
 */
export interface TableTypeDetection {
  tableName: string;
  detectedType: TableType;
  confidence: number; // 0-1
  reasons: string[];
  suggestedRelations: TableRelation[];
}

/**
 * 数据建模验证结果
 */
export interface DataModelingValidation {
  isValid: boolean;
  score: number; // 0-100
  issues: DataModelingIssue[];
  recommendations: string[];
  fixActions?: FixAction[];
}

/**
 * 表结构信息
 */
export interface TableStructure {
  name: string;
  type: TableType;
  headers: string[];
  rowCount: number;
  keyColumn?: string;
  formulaColumns: string[];
  valueColumns: string[];
}

/**
 * 数据建模验证器
 * 检查表格是否遵循正确的数据建模原则
 */
export class DataModelingValidator {
  // 表类型识别关键词
  private readonly masterKeywords = [
    "产品",
    "客户",
    "员工",
    "供应商",
    "物料",
    "目录",
    "主数据",
    "基础",
    "master",
    "product",
    "customer",
  ];
  private readonly transactionKeywords = [
    "订单",
    "交易",
    "销售",
    "采购",
    "记录",
    "明细",
    "order",
    "transaction",
    "sales",
  ];
  private readonly summaryKeywords = [
    "汇总",
    "统计",
    "报表",
    "月度",
    "年度",
    "summary",
    "report",
    "total",
  ];
  private readonly analysisKeywords = ["分析", "KPI", "利润", "预测", "趋势", "analysis", "profit"];

  // 需要用 XLOOKUP 引用的字段
  private readonly lookupFields = [
    "单价",
    "成本",
    "价格",
    "单位成本",
    "产品名称",
    "客户名称",
    "供应商名称",
  ];

  // 需要用公式计算的字段
  private readonly formulaFields = [
    "销售额",
    "金额",
    "总成本",
    "总价",
    "利润",
    "毛利",
    "毛利率",
    "净利率",
  ];

  // 需要用聚合函数的字段
  private readonly aggregateFields = ["销量", "销售额", "总成本", "总数", "总金额", "数量合计"];

  /**
   * 智能识别表类型
   */
  detectTableType(tableName: string, headers: string[]): TableTypeDetection {
    const name = tableName.toLowerCase();
    const headerStr = headers.join(" ").toLowerCase();
    let detectedType: TableType = "unknown";
    let confidence = 0;
    const reasons: string[] = [];
    const suggestedRelations: TableRelation[] = [];

    // 检查主数据表特征
    if (this.masterKeywords.some((kw) => name.includes(kw) || headerStr.includes(kw))) {
      if (headers.some((h) => h.includes("ID") || h.includes("编号") || h.includes("代码"))) {
        detectedType = "master";
        confidence = 0.9;
        reasons.push("表名包含主数据关键词");
        reasons.push("包含ID/编号字段");
      }
    }

    // 检查交易表特征
    if (this.transactionKeywords.some((kw) => name.includes(kw) || headerStr.includes(kw))) {
      detectedType = "transaction";
      confidence = 0.85;
      reasons.push("表名包含交易/订单关键词");

      // 检测可能需要 LOOKUP 的字段
      headers.forEach((h) => {
        if (this.lookupFields.some((f) => h.includes(f))) {
          suggestedRelations.push({
            sourceTable: tableName,
            targetTable: "产品主数据表",
            relationshipType: "lookup",
            keyColumn: "产品ID",
            lookupColumn: h,
          });
        }
      });
    }

    // 检查汇总表特征
    if (this.summaryKeywords.some((kw) => name.includes(kw) || headerStr.includes(kw))) {
      detectedType = "summary";
      confidence = 0.85;
      reasons.push("表名包含汇总/统计关键词");

      suggestedRelations.push({
        sourceTable: tableName,
        targetTable: "交易表",
        relationshipType: "many_to_one",
        keyColumn: "产品ID",
      });
    }

    // 检查分析表特征
    if (this.analysisKeywords.some((kw) => name.includes(kw) || headerStr.includes(kw))) {
      detectedType = "analysis";
      confidence = 0.8;
      reasons.push("表名包含分析/KPI关键词");
    }

    return {
      tableName,
      detectedType,
      confidence,
      reasons,
      suggestedRelations,
    };
  }

  /**
   * 生成正确的公式建议
   */
  generateFormulaSuggestion(
    fieldName: string,
    tableType: TableType,
    masterTableName?: string,
    transactionTableName?: string
  ): string {
    const masterRef = masterTableName || "产品主数据表";
    const transRef = transactionTableName || "订单交易表";

    // 交易表的 LOOKUP 公式
    if (tableType === "transaction") {
      if (fieldName.includes("单价") || fieldName.includes("价格")) {
        return `=XLOOKUP([@产品ID], ${masterRef}[产品ID], ${masterRef}[单价])`;
      }
      if (fieldName.includes("成本") && !fieldName.includes("总")) {
        return `=XLOOKUP([@产品ID], ${masterRef}[产品ID], ${masterRef}[成本])`;
      }
      if (fieldName.includes("产品名称")) {
        return `=XLOOKUP([@产品ID], ${masterRef}[产品ID], ${masterRef}[产品名称])`;
      }
      if (fieldName.includes("销售额") || fieldName.includes("金额")) {
        return "=[@数量]*[@单价]";
      }
      if (fieldName.includes("总成本")) {
        return "=[@数量]*[@成本]";
      }
      if (fieldName.includes("利润") || fieldName.includes("毛利")) {
        return "=[@销售额]-[@总成本]";
      }
    }

    // 汇总表的聚合公式
    if (tableType === "summary") {
      if (fieldName.includes("销量") || fieldName.includes("数量")) {
        return `=SUMIF(${transRef}[产品ID], [@产品ID], ${transRef}[数量])`;
      }
      if (fieldName.includes("销售额")) {
        return `=SUMIF(${transRef}[产品ID], [@产品ID], ${transRef}[销售额])`;
      }
      if (fieldName.includes("总成本")) {
        return `=SUMIF(${transRef}[产品ID], [@产品ID], ${transRef}[总成本])`;
      }
      if (fieldName.includes("毛利") && !fieldName.includes("率")) {
        return "=[@销售额]-[@总成本]";
      }
      if (fieldName.includes("毛利率") || fieldName.includes("利润率")) {
        return "=[@毛利]/[@销售额]";
      }
    }

    return "";
  }

  /**
   * 检查交易表是否正确引用主数据
   */
  validateTransactionTable(
    transactionData: unknown[][],
    headers: string[],
    masterTableName?: string
  ): DataModelingIssue[] {
    const issues: DataModelingIssue[] = [];

    for (let colIdx = 0; colIdx < headers.length; colIdx++) {
      const header = headers[colIdx];

      // 检查应该是 LOOKUP 的字段
      if (this.lookupFields.some((h) => header.includes(h))) {
        const values = transactionData.map((row) => row[colIdx]);
        const uniqueValues = new Set(values.filter((v) => v !== null && v !== undefined));

        if (uniqueValues.size === 1 && transactionData.length > 3) {
          const suggestedFormula = this.generateFormulaSuggestion(
            header,
            "transaction",
            masterTableName
          );
          issues.push({
            type: "hardcoded_value",
            severity: "critical",
            location: `列 ${header}`,
            message: `"${header}" 列所有值都是 ${[...uniqueValues][0]}，疑似硬编码`,
            suggestion: `使用公式: ${suggestedFormula}`,
            fixAction: {
              action: "set_formula",
              target: `${header}列`,
              formula: suggestedFormula,
            },
          });
        }
      }

      // 检查应该是公式计算的字段
      if (this.formulaFields.some((h) => header.includes(h))) {
        const values = transactionData.map((row) => row[colIdx]);
        const uniqueValues = new Set(values.filter((v) => v !== null && v !== undefined));

        if (uniqueValues.size === 1 && transactionData.length > 3) {
          const suggestedFormula = this.generateFormulaSuggestion(
            header,
            "transaction",
            masterTableName
          );
          issues.push({
            type: "missing_formula",
            severity: "critical",
            location: `列 ${header}`,
            message: `"${header}" 列所有值都是 ${[...uniqueValues][0]}，应该用公式计算`,
            suggestion: suggestedFormula
              ? `使用公式: ${suggestedFormula}`
              : "使用公式如 =数量*单价 来计算",
            fixAction: suggestedFormula
              ? {
                  action: "set_formula",
                  target: `${header}列`,
                  formula: suggestedFormula,
                }
              : undefined,
          });
        }
      }
    }

    return issues;
  }

  /**
   * 检查汇总表是否正确使用聚合函数
   */
  validateSummaryTable(
    summaryData: unknown[][],
    headers: string[],
    transactionTableName?: string
  ): DataModelingIssue[] {
    const issues: DataModelingIssue[] = [];

    for (let colIdx = 0; colIdx < headers.length; colIdx++) {
      const header = headers[colIdx];

      if (this.aggregateFields.some((h) => header.includes(h))) {
        const values = summaryData.map((row) => row[colIdx]);
        const uniqueValues = new Set(values.filter((v) => v !== null && v !== undefined));

        // 如果所有汇总值都相同，很可能是错误的
        if (uniqueValues.size === 1 && summaryData.length > 2) {
          const suggestedFormula = this.generateFormulaSuggestion(
            header,
            "summary",
            undefined,
            transactionTableName
          );
          issues.push({
            type: "duplicate_data",
            severity: "critical",
            location: `列 ${header}`,
            message: `汇总表 "${header}" 列所有值相同 (${[...uniqueValues][0]})，这是不正常的`,
            suggestion: suggestedFormula ? `使用公式: ${suggestedFormula}` : `使用 SUMIF 聚合数据`,
            fixAction: suggestedFormula
              ? {
                  action: "set_formula",
                  target: `${header}列`,
                  formula: suggestedFormula,
                }
              : undefined,
          });
        }
      }
    }

    // 检查毛利率是否所有相同
    const profitRateIdx = headers.findIndex((h) => h.includes("毛利率") || h.includes("利润率"));
    if (profitRateIdx !== -1) {
      const values = summaryData.map((row) => row[profitRateIdx]);
      const uniqueValues = new Set(values.filter((v) => v !== null && v !== undefined));

      if (uniqueValues.size === 1 && summaryData.length > 2) {
        issues.push({
          type: "inconsistent_data",
          severity: "critical",
          location: "毛利率列",
          message: `所有产品毛利率都是 ${[...uniqueValues][0]}，这不符合实际`,
          suggestion: "毛利率应该用公式 =[@毛利]/[@销售额] 计算，不同产品应该有不同的毛利率",
          fixAction: {
            action: "set_formula",
            target: "毛利率列",
            formula: "=[@毛利]/[@销售额]",
          },
        });
      }
    }

    return issues;
  }

  /**
   * 验证表创建顺序是否正确
   */
  validateTableOrder(tables: TableStructure[]): DataModelingIssue[] {
    const issues: DataModelingIssue[] = [];
    const typeOrder = ["master", "transaction", "summary", "analysis"];

    let lastTypeIndex = -1;
    for (const table of tables) {
      const currentIndex = typeOrder.indexOf(table.type);
      if (currentIndex !== -1 && currentIndex < lastTypeIndex) {
        issues.push({
          type: "wrong_table_order",
          severity: "warning",
          location: table.name,
          message: `表 "${table.name}" (${table.type}) 创建顺序不正确`,
          suggestion: "应该先创建主数据表，再创建交易表，然后是汇总表，最后是分析表",
        });
      }
      if (currentIndex !== -1) {
        lastTypeIndex = currentIndex;
      }
    }

    return issues;
  }

  /**
   * 检测可能需要的表关系
   */
  detectMissingRelations(tables: TableStructure[]): DataModelingIssue[] {
    const issues: DataModelingIssue[] = [];

    const transactionTables = tables.filter((t) => t.type === "transaction");
    const masterTables = tables.filter((t) => t.type === "master");

    for (const trans of transactionTables) {
      // 检查交易表是否有对应的主数据表
      const hasLookupColumns = trans.headers.some((h) =>
        this.lookupFields.some((f) => h.includes(f))
      );

      if (hasLookupColumns && masterTables.length === 0) {
        issues.push({
          type: "missing_master_reference",
          severity: "critical",
          location: trans.name,
          message: `交易表 "${trans.name}" 包含需要从主数据引用的字段，但没有找到主数据表`,
          suggestion: "请先创建主数据表（如产品表），包含产品ID、单价、成本等字段",
        });
      }
    }

    return issues;
  }

  /**
   * 综合验证数据建模质量
   */
  validateDataModeling(
    tableType: TableType,
    data: unknown[][],
    headers: string[],
    masterTableName?: string,
    transactionTableName?: string
  ): DataModelingValidation {
    let issues: DataModelingIssue[] = [];

    if (tableType === "transaction") {
      issues = this.validateTransactionTable(data, headers, masterTableName);
    } else if (tableType === "summary") {
      issues = this.validateSummaryTable(data, headers, transactionTableName);
    }

    // 计算分数
    const criticalCount = issues.filter((i) => i.severity === "critical").length;
    const warningCount = issues.filter((i) => i.severity === "warning").length;

    let score = 100;
    score -= criticalCount * 25; // 每个严重问题扣25分
    score -= warningCount * 10; // 每个警告扣10分
    score = Math.max(0, score);

    // 生成建议
    const recommendations: string[] = [];
    if (criticalCount > 0) {
      recommendations.push("⚠️ 发现严重的数据建模问题，需要立即修复");
    }
    if (issues.some((i) => i.type === "hardcoded_value")) {
      recommendations.push("📌 使用 XLOOKUP 从主数据表引用数据，避免硬编码");
    }
    if (issues.some((i) => i.type === "duplicate_data")) {
      recommendations.push("📊 使用 SUMIF/COUNTIF 聚合数据，而不是手工填写");
    }
    if (issues.some((i) => i.type === "missing_formula")) {
      recommendations.push("🔢 计算字段必须使用公式，确保数据一致性");
    }

    // 收集修复动作
    const fixActions = issues.filter((i) => i.fixAction).map((i) => i.fixAction!);

    return {
      isValid: criticalCount === 0,
      score,
      issues,
      recommendations,
      fixActions: fixActions.length > 0 ? fixActions : undefined,
    };
  }

  /**
   * 生成完整的修复脚本
   */
  generateFixScript(issues: DataModelingIssue[]): string[] {
    const scripts: string[] = [];

    for (const issue of issues) {
      if (issue.fixAction) {
        switch (issue.fixAction.action) {
          case "set_formula":
            scripts.push(`// 修复: ${issue.message}`);
            scripts.push(
              `await adapter.excel_set_formula("${issue.fixAction.target}", "${issue.fixAction.formula}");`
            );
            break;
          case "fill_formula":
            scripts.push(`// 填充公式到整列`);
            scripts.push(`await adapter.excel_fill_formula("${issue.fixAction.target}");`);
            break;
        }
      }
    }

    return scripts;
  }
}

// 导出单例
export const dataModelingValidator = new DataModelingValidator();

// ========== 辅助类型 ==========

interface WorkbookContext {
  sheets: Array<{
    name: string;
    rowCount: number;
    columnCount: number;
  }>;
}

// 导出单例
export const formulaValidator = new FormulaValidator();
