/**
 * FormulaCompiler - 公式编译器 v2.9.53
 *
 * 职责：将逻辑公式（使用字段名）编译为 Excel 可执行公式
 *
 * 输入: logicalFormula（如 "=单价*数量"）+ 表结构信息
 * 输出: excelFormula（如 "=C2*D2" 或 "=Table1[@单价]*Table1[@数量]"）
 *
 * 两种输出模式：
 * 1. A1 引用模式: 适用于普通范围数据
 * 2. Table 结构化引用模式: 适用于 Excel Table
 */

// ========== 类型定义 ==========

/**
 * 编译模式
 */
export type CompileMode = "A1" | "Table";

/**
 * 字段映射条目
 */
export interface FieldMapping {
  fieldName: string; // 字段名（中文/英文）
  columnLetter: string; // 列字母 (A, B, C...)
  columnIndex: number; // 列索引 (0-based)
  tableName?: string; // Excel Table 名称（可选）
}

/**
 * 表结构信息
 */
export interface TableSchema {
  tableName: string; // 逻辑表名
  excelTableName?: string; // Excel Table 名称（如果是 Table）
  headerRow: number; // 表头行号（1-based）
  dataStartRow: number; // 数据起始行（1-based）
  fields: FieldMapping[]; // 字段映射列表
}

/**
 * 编译上下文
 */
export interface CompileContext {
  targetRow: number; // 目标行号（1-based，用于 A1 模式）
  mode: CompileMode; // 编译模式
  currentTable: TableSchema; // 当前表
  referenceTables?: TableSchema[]; // 引用的其他表（用于 XLOOKUP 等）
  useAbsoluteRef?: boolean; // 是否使用绝对引用（默认 false）
}

/**
 * 编译结果
 */
export interface CompileResult {
  success: boolean;
  excelFormula?: string;
  errors?: CompileError[];
  warnings?: CompileWarning[];
  usedFields?: string[]; // 使用的字段列表
}

/**
 * 编译错误
 */
export interface CompileError {
  type: "unknown_field" | "invalid_syntax" | "circular_ref" | "missing_table";
  message: string;
  position?: number; // 公式中的位置
  fieldName?: string;
}

/**
 * 编译警告
 */
export interface CompileWarning {
  type: "ambiguous_field" | "implicit_type_cast" | "deprecated_syntax";
  message: string;
  fieldName?: string;
}

// ========== 公式编译器 ==========

export class FormulaCompiler {
  // Excel 函数列表（避免把函数名当字段名）
  private static readonly EXCEL_FUNCTIONS = new Set([
    "SUM",
    "SUMIF",
    "SUMIFS",
    "SUMPRODUCT",
    "AVERAGE",
    "AVERAGEIF",
    "AVERAGEIFS",
    "COUNT",
    "COUNTA",
    "COUNTIF",
    "COUNTIFS",
    "COUNTBLANK",
    "MAX",
    "MAXIFS",
    "MIN",
    "MINIFS",
    "VLOOKUP",
    "HLOOKUP",
    "XLOOKUP",
    "LOOKUP",
    "INDEX",
    "MATCH",
    "IF",
    "IFS",
    "IFERROR",
    "IFNA",
    "AND",
    "OR",
    "NOT",
    "TRUE",
    "FALSE",
    "LEFT",
    "RIGHT",
    "MID",
    "LEN",
    "TRIM",
    "UPPER",
    "LOWER",
    "CONCATENATE",
    "CONCAT",
    "TEXTJOIN",
    "DATE",
    "YEAR",
    "MONTH",
    "DAY",
    "TODAY",
    "NOW",
    "ROUND",
    "ROUNDUP",
    "ROUNDDOWN",
    "INT",
    "MOD",
    "ABS",
    "UNIQUE",
    "SORT",
    "FILTER",
    "SEQUENCE",
    "OFFSET",
    "INDIRECT",
    "ROW",
    "COLUMN",
    "ADDRESS",
    "EOMONTH",
    "EDATE",
    "DATEDIF",
    "WEEKDAY",
    "WORKDAY",
  ]);

  // 中文运算符映射（可选，用于处理用户输入中文标点）
  // 使用 Unicode 转义避免编码问题
  private static readonly CN_OPERATORS: Record<string, string> = {
    "\u00D7": "*", // × 乘号
    "\u00F7": "/", // ÷ 除号
    "\uFF08": "(", // （ 中文左括号
    "\uFF09": ")", // ） 中文右括号
    "\uFF0C": ",", // ， 中文逗号
    "\u201C": '"', // " 中文左双引号
    "\u201D": '"', // " 中文右双引号
    "\u2018": "'", // ' 中文左单引号
    "\u2019": "'", // ' 中文右单引号
  };

  /**
   * 编译逻辑公式为 Excel 公式
   */
  compile(logicalFormula: string, context: CompileContext): CompileResult {
    const errors: CompileError[] = [];
    const warnings: CompileWarning[] = [];
    const usedFields: string[] = [];

    if (!logicalFormula) {
      return {
        success: false,
        errors: [{ type: "invalid_syntax", message: "公式为空" }],
      };
    }

    // 预处理：移除开头的 = 号，后面再加回来
    let formula = logicalFormula.trim();
    const hasEquals = formula.startsWith("=");
    if (hasEquals) {
      formula = formula.substring(1);
    }

    // 预处理：转换中文标点
    formula = this.normalizePunctuation(formula);

    // 构建字段名到映射的索引
    const fieldIndex = this.buildFieldIndex(context);

    // 解析并替换字段引用
    const result = this.parseAndReplace(formula, fieldIndex, context, usedFields, errors, warnings);

    if (errors.length > 0) {
      return { success: false, errors, warnings, usedFields };
    }

    // 加回 = 号
    const excelFormula = hasEquals ? `=${result}` : result;

    return {
      success: true,
      excelFormula,
      warnings: warnings.length > 0 ? warnings : undefined,
      usedFields,
    };
  }

  /**
   * 批量编译：为多行生成公式
   */
  compileForRange(
    logicalFormula: string,
    schema: TableSchema,
    startRow: number,
    endRow: number,
    mode: CompileMode = "A1"
  ): Map<number, CompileResult> {
    const results = new Map<number, CompileResult>();

    // Table 模式只需编译一次
    if (mode === "Table") {
      const context: CompileContext = {
        targetRow: startRow,
        mode: "Table",
        currentTable: schema,
      };
      const result = this.compile(logicalFormula, context);
      for (let row = startRow; row <= endRow; row++) {
        results.set(row, result);
      }
      return results;
    }

    // A1 模式需要为每行编译
    for (let row = startRow; row <= endRow; row++) {
      const context: CompileContext = {
        targetRow: row,
        mode: "A1",
        currentTable: schema,
      };
      results.set(row, this.compile(logicalFormula, context));
    }

    return results;
  }

  /**
   * 验证逻辑公式语法（不编译，只检查）
   */
  validate(logicalFormula: string, fieldNames: string[]): CompileResult {
    const errors: CompileError[] = [];
    const warnings: CompileWarning[] = [];

    if (!logicalFormula) {
      return { success: false, errors: [{ type: "invalid_syntax", message: "公式为空" }] };
    }

    let formula = logicalFormula.trim();
    if (formula.startsWith("=")) formula = formula.substring(1);

    // 检查括号匹配
    let parenCount = 0;
    for (let i = 0; i < formula.length; i++) {
      if (formula[i] === "(") parenCount++;
      if (formula[i] === ")") parenCount--;
      if (parenCount < 0) {
        errors.push({
          type: "invalid_syntax",
          message: `位置 ${i + 1} 处多余的右括号`,
          position: i,
        });
        break;
      }
    }
    if (parenCount > 0) {
      errors.push({ type: "invalid_syntax", message: "缺少右括号" });
    }

    // 提取字段引用，检查是否存在
    const tokens = this.tokenize(formula);
    const fieldSet = new Set(fieldNames.map((f) => f.toLowerCase()));

    for (const token of tokens) {
      if (
        token.type === "identifier" &&
        !FormulaCompiler.EXCEL_FUNCTIONS.has(token.value.toUpperCase())
      ) {
        if (!fieldSet.has(token.value.toLowerCase())) {
          errors.push({
            type: "unknown_field",
            message: `未知字段: ${token.value}`,
            fieldName: token.value,
            position: token.position,
          });
        }
      }
    }

    return {
      success: errors.length === 0,
      errors: errors.length > 0 ? errors : undefined,
      warnings: warnings.length > 0 ? warnings : undefined,
    };
  }

  /**
   * 从公式提取所有字段引用
   */
  extractFieldReferences(logicalFormula: string): string[] {
    const fields: string[] = [];

    if (!logicalFormula) return fields;

    let formula = logicalFormula.trim();
    if (formula.startsWith("=")) formula = formula.substring(1);

    const tokens = this.tokenize(formula);

    for (const token of tokens) {
      if (
        token.type === "identifier" &&
        !FormulaCompiler.EXCEL_FUNCTIONS.has(token.value.toUpperCase())
      ) {
        if (!fields.includes(token.value)) {
          fields.push(token.value);
        }
      }
    }

    return fields;
  }

  // ========== 私有方法 ==========

  /**
   * 构建字段名索引（支持模糊匹配）
   */
  private buildFieldIndex(context: CompileContext): Map<string, FieldMapping> {
    const index = new Map<string, FieldMapping>();

    // 添加当前表的字段
    for (const field of context.currentTable.fields) {
      index.set(field.fieldName.toLowerCase(), field);
    }

    // 添加引用表的字段（带表名前缀）
    if (context.referenceTables) {
      for (const table of context.referenceTables) {
        for (const field of table.fields) {
          // "表名.字段名" 格式
          index.set(`${table.tableName.toLowerCase()}.${field.fieldName.toLowerCase()}`, field);
        }
      }
    }

    return index;
  }

  /**
   * 解析并替换字段引用
   */
  private parseAndReplace(
    formula: string,
    fieldIndex: Map<string, FieldMapping>,
    context: CompileContext,
    usedFields: string[],
    errors: CompileError[],
    warnings: CompileWarning[]
  ): string {
    const tokens = this.tokenize(formula);
    let result = "";
    let lastEnd = 0;

    for (const token of tokens) {
      // 添加 token 之前的内容（运算符、空格等）
      result += formula.substring(lastEnd, token.position);

      if (token.type === "identifier") {
        const lowerValue = token.value.toLowerCase();

        // 检查是否是 Excel 函数
        if (FormulaCompiler.EXCEL_FUNCTIONS.has(token.value.toUpperCase())) {
          result += token.value;
        }
        // 检查是否是已知字段
        else if (fieldIndex.has(lowerValue)) {
          const field = fieldIndex.get(lowerValue)!;
          usedFields.push(field.fieldName);
          result += this.generateReference(field, context);
        }
        // 尝试模糊匹配
        else {
          const matched = this.fuzzyMatchField(token.value, fieldIndex);
          if (matched) {
            usedFields.push(matched.fieldName);
            warnings.push({
              type: "ambiguous_field",
              message: `字段 "${token.value}" 匹配到 "${matched.fieldName}"`,
              fieldName: token.value,
            });
            result += this.generateReference(matched, context);
          } else {
            errors.push({
              type: "unknown_field",
              message: `未知字段: ${token.value}`,
              fieldName: token.value,
              position: token.position,
            });
            result += token.value; // 保留原文
          }
        }
      } else {
        result += token.value;
      }

      lastEnd = token.position + token.value.length;
    }

    // 添加剩余内容
    result += formula.substring(lastEnd);

    return result;
  }

  /**
   * 生成单元格引用
   */
  private generateReference(field: FieldMapping, context: CompileContext): string {
    if (context.mode === "Table") {
      // Table 结构化引用: Table1[@字段名]
      const tableName =
        field.tableName || context.currentTable.excelTableName || context.currentTable.tableName;
      return `${tableName}[@${field.fieldName}]`;
    } else {
      // A1 引用: C2
      const col = field.columnLetter;
      const row = context.targetRow;

      if (context.useAbsoluteRef) {
        return `$${col}$${row}`;
      }
      return `${col}${row}`;
    }
  }

  /**
   * 模糊匹配字段名
   */
  private fuzzyMatchField(
    name: string,
    fieldIndex: Map<string, FieldMapping>
  ): FieldMapping | null {
    const lowerName = name.toLowerCase();

    // 1. 尝试部分匹配
    for (const [key, field] of fieldIndex) {
      if (key.includes(lowerName) || lowerName.includes(key)) {
        return field;
      }
    }

    // 2. 同义词映射
    const synonyms: Record<string, string[]> = {
      价格: ["单价", "售价", "price"],
      数量: ["qty", "quantity", "amount"],
      金额: ["销售额", "总额", "total"],
      成本: ["cost", "进价"],
      利润: ["profit", "毛利"],
    };

    for (const [canonical, aliases] of Object.entries(synonyms)) {
      if (aliases.includes(lowerName) || lowerName === canonical) {
        for (const alias of [canonical, ...aliases]) {
          if (fieldIndex.has(alias.toLowerCase())) {
            return fieldIndex.get(alias.toLowerCase())!;
          }
        }
      }
    }

    return null;
  }

  /**
   * 词法分析：将公式拆分为 token
   */
  private tokenize(formula: string): Token[] {
    const tokens: Token[] = [];
    let i = 0;

    while (i < formula.length) {
      const char = formula[i];

      // 跳过空白
      if (/\s/.test(char)) {
        i++;
        continue;
      }

      // 字符串字面量
      if (char === '"') {
        const start = i;
        i++;
        while (i < formula.length && formula[i] !== '"') {
          if (formula[i] === "\\") i++; // 跳过转义
          i++;
        }
        i++; // 跳过结束引号
        tokens.push({ type: "string", value: formula.substring(start, i), position: start });
        continue;
      }

      // 数字
      if (/\d/.test(char) || (char === "." && /\d/.test(formula[i + 1] || ""))) {
        const start = i;
        while (i < formula.length && /[\d.eE+-]/.test(formula[i])) {
          i++;
        }
        tokens.push({ type: "number", value: formula.substring(start, i), position: start });
        continue;
      }

      // 标识符（字段名或函数名）- 支持中文
      if (/[a-zA-Z_\u4e00-\u9fa5]/.test(char)) {
        const start = i;
        while (i < formula.length && /[\w\u4e00-\u9fa5]/.test(formula[i])) {
          i++;
        }
        tokens.push({ type: "identifier", value: formula.substring(start, i), position: start });
        continue;
      }

      // 运算符和其他字符
      tokens.push({ type: "operator", value: char, position: i });
      i++;
    }

    return tokens;
  }

  /**
   * 中文标点转英文
   */
  private normalizePunctuation(formula: string): string {
    let result = formula;
    for (const [cn, en] of Object.entries(FormulaCompiler.CN_OPERATORS)) {
      result = result.split(cn).join(en);
    }
    return result;
  }

  // ========== 辅助方法：从表结构创建 Schema ==========

  /**
   * 从 header 数组创建 TableSchema
   */
  static createSchemaFromHeaders(
    tableName: string,
    headers: string[],
    headerRow: number = 1,
    excelTableName?: string
  ): TableSchema {
    const fields: FieldMapping[] = headers.map((header, index) => ({
      fieldName: String(header || ""),
      columnLetter: FormulaCompiler.indexToColumn(index),
      columnIndex: index,
      tableName: excelTableName,
    }));

    return {
      tableName,
      excelTableName,
      headerRow,
      dataStartRow: headerRow + 1,
      fields,
    };
  }

  /**
   * 列索引转字母
   */
  static indexToColumn(index: number): string {
    let result = "";
    let n = index + 1;
    while (n > 0) {
      n--;
      result = String.fromCharCode(65 + (n % 26)) + result;
      n = Math.floor(n / 26);
    }
    return result;
  }

  /**
   * 列字母转索引
   */
  static columnToIndex(column: string): number {
    let result = 0;
    for (let i = 0; i < column.length; i++) {
      result = result * 26 + (column.charCodeAt(i) - 64);
    }
    return result - 1;
  }
}

// ========== Token 类型 ==========

interface Token {
  type: "identifier" | "number" | "string" | "operator";
  value: string;
  position: number;
}

// ========== 导出 ==========

export const formulaCompiler = new FormulaCompiler();
