/**
 * FormulaTranslator - 公式引用翻译器 v1.0
 *
 * 借鉴自 sv-excel-agent helpers.py
 *
 * 用于 auto_fill 等功能，将公式中的单元格引用按偏移量调整。
 *
 * 处理的引用类型：
 * - 相对引用 (A1) - 按偏移调整
 * - 绝对行 ($A1) - 列调整，行固定
 * - 绝对列 (A$1) - 行调整，列固定
 * - 完全绝对 ($A$1) - 不调整
 *
 * @see https://github.com/Sylvian/sv-excel-agent
 */

// ==================== 列号转换 ====================

/**
 * 列字母转数字 (A=1, B=2, ..., Z=26, AA=27, ...)
 */
export function columnLetterToNumber(letters: string): number {
  let result = 0;
  for (const char of letters.toUpperCase()) {
    result = result * 26 + (char.charCodeAt(0) - "A".charCodeAt(0) + 1);
  }
  return result;
}

/**
 * 数字转列字母 (1=A, 2=B, ..., 26=Z, 27=AA, ...)
 */
export function numberToColumnLetter(num: number): string {
  let result = "";
  while (num > 0) {
    num--;
    result = String.fromCharCode("A".charCodeAt(0) + (num % 26)) + result;
    num = Math.floor(num / 26);
  }
  return result;
}

// ==================== 公式翻译 ====================

/**
 * 单元格引用正则表达式
 *
 * 匹配模式：
 * - A1, B2, C10 (相对引用)
 * - $A1, $B$2 (绝对引用)
 * - Sheet1!A1 (跨表引用) - 只翻译单元格部分
 */
const CELL_REF_PATTERN = /(\$?)([A-Z]+)(\$?)(\d+)/gi;

/**
 * 翻译公式中的单元格引用
 *
 * @param formula - 原始公式（如 "=A1+B1"）
 * @param rowOffset - 行偏移量（正数向下，负数向上）
 * @param colOffset - 列偏移量（正数向右，负数向左）
 * @returns 翻译后的公式
 *
 * @example
 * ```typescript
 * // 相对引用
 * translateFormula("=A1+B1", 1, 0);  // "=A2+B2"
 * translateFormula("=A1+B1", 0, 1);  // "=B1+C1"
 *
 * // 绝对引用
 * translateFormula("=$A$1+B1", 1, 1);  // "=$A$1+C2"
 * translateFormula("=$A1+B$1", 1, 1);  // "=$A2+C$1"
 *
 * // 复杂公式
 * translateFormula("=SUM(A1:A10)", 5, 0);  // "=SUM(A6:A15)"
 * translateFormula("=VLOOKUP(A1,$B$1:$C$100,2)", 1, 0);  // "=VLOOKUP(A2,$B$1:$C$100,2)"
 * ```
 */
export function translateFormula(formula: string, rowOffset: number, colOffset: number): string {
  if (!formula.startsWith("=")) {
    return formula;
  }

  return formula.replace(
    CELL_REF_PATTERN,
    (match, colAbsolute, colLetters, rowAbsolute, rowNum) => {
      const isColAbsolute = colAbsolute === "$";
      const isRowAbsolute = rowAbsolute === "$";

      let newColLetters = colLetters.toUpperCase();
      let newRowNum = parseInt(rowNum, 10);

      // 调整列（如果不是绝对列）
      if (!isColAbsolute && colOffset !== 0) {
        let colIndex = columnLetterToNumber(colLetters);
        colIndex += colOffset;
        // 确保列号有效（最小为1）
        colIndex = Math.max(1, colIndex);
        newColLetters = numberToColumnLetter(colIndex);
      }

      // 调整行（如果不是绝对行）
      if (!isRowAbsolute && rowOffset !== 0) {
        newRowNum += rowOffset;
        // 确保行号有效（最小为1）
        newRowNum = Math.max(1, newRowNum);
      }

      // 重构引用
      return `${isColAbsolute ? "$" : ""}${newColLetters}${isRowAbsolute ? "$" : ""}${newRowNum}`;
    }
  );
}

// ==================== 数组公式检测 ====================

/**
 * 检测公式是否需要作为数组公式 (CSE) 输入
 *
 * 常见需要 CSE 的模式：
 * - MATCH(1, (range=value)*(range=value), 0) - 数组 AND 逻辑
 * - INDEX(..., MATCH(1, ...*..., 0)) - 带数组匹配的查找
 *
 * 不需要 CSE 的：
 * - SUMPRODUCT - 原生处理数组
 * - 现代 Excel 动态数组函数
 *
 * @param formula - 公式字符串
 * @returns 是否需要 CSE
 */
export function isArrayFormula(formula: string): boolean {
  if (!formula.startsWith("=")) {
    return false;
  }

  const upper = formula.toUpperCase();

  // SUMPRODUCT 原生处理数组
  if (upper.startsWith("=SUMPRODUCT")) {
    return false;
  }

  // 动态数组函数（Excel 365+）不需要 CSE
  const dynamicArrayFunctions = [
    "FILTER",
    "SORT",
    "SORTBY",
    "UNIQUE",
    "SEQUENCE",
    "RANDARRAY",
    "XLOOKUP",
    "XMATCH",
    "LET",
    "LAMBDA",
  ];
  for (const fn of dynamicArrayFunctions) {
    if (upper.includes(fn + "(")) {
      return false;
    }
  }

  // 模式：MATCH(1, ...) 内部有乘法 - 数组 AND 逻辑
  if (/MATCH\s*\(\s*1\s*,/i.test(formula)) {
    // 检查是否有范围比较的乘法
    if (/\([^)]*[A-Z]+\d*:[A-Z]+\d*[^)]*\)\s*\*\s*\(/i.test(formula)) {
      return true;
    }
  }

  // 模式：使用布尔运算符连接范围比较
  if (/\([^)]*[A-Z]+\d*:[A-Z]+\d*[^)]*[=<>][^)]*\)\s*\*/i.test(formula)) {
    return true;
  }

  return false;
}

// ==================== 区域解析 ====================

/**
 * 解析单元格地址
 *
 * @param address - 单元格地址（如 "A1", "$B$2", "Sheet1!C3"）
 * @returns { sheet?: string, col: string, row: number, colAbsolute: boolean, rowAbsolute: boolean }
 */
export function parseCellAddress(address: string): {
  sheet?: string;
  col: string;
  row: number;
  colAbsolute: boolean;
  rowAbsolute: boolean;
} {
  let sheet: string | undefined;
  let cellPart = address;

  // 处理跨表引用
  if (address.includes("!")) {
    const parts = address.split("!");
    sheet = parts[0].replace(/^'|'$/g, ""); // 移除引号
    cellPart = parts[1];
  }

  // 解析单元格
  const match = cellPart.match(/^(\$?)([A-Z]+)(\$?)(\d+)$/i);
  if (!match) {
    throw new Error(`Invalid cell address: ${address}`);
  }

  return {
    sheet,
    col: match[2].toUpperCase(),
    row: parseInt(match[4], 10),
    colAbsolute: match[1] === "$",
    rowAbsolute: match[3] === "$",
  };
}

/**
 * 解析区域地址
 *
 * @param range - 区域地址（如 "A1:B10", "Sheet1!C1:D5"）
 * @returns { sheet?: string, startCol: number, startRow: number, endCol: number, endRow: number }
 */
export function parseRangeAddress(range: string): {
  sheet?: string;
  startCol: number;
  startRow: number;
  endCol: number;
  endRow: number;
} {
  let sheet: string | undefined;
  let rangePart = range;

  // 处理跨表引用
  if (range.includes("!")) {
    const parts = range.split("!");
    sheet = parts[0].replace(/^'|'$/g, "");
    rangePart = parts[1];
  }

  // 分割起止单元格
  const [startCell, endCell] = rangePart.split(":");

  if (!endCell) {
    // 单个单元格
    const start = parseCellAddress(startCell);
    return {
      sheet,
      startCol: columnLetterToNumber(start.col),
      startRow: start.row,
      endCol: columnLetterToNumber(start.col),
      endRow: start.row,
    };
  }

  const start = parseCellAddress(startCell);
  const end = parseCellAddress(endCell);

  return {
    sheet,
    startCol: columnLetterToNumber(start.col),
    startRow: start.row,
    endCol: columnLetterToNumber(end.col),
    endRow: end.row,
  };
}

/**
 * 构建单元格地址
 */
export function buildCellAddress(
  col: number | string,
  row: number,
  options?: { colAbsolute?: boolean; rowAbsolute?: boolean; sheet?: string }
): string {
  const colLetter = typeof col === "number" ? numberToColumnLetter(col) : col;
  const colPrefix = options?.colAbsolute ? "$" : "";
  const rowPrefix = options?.rowAbsolute ? "$" : "";
  const sheetPrefix = options?.sheet ? `'${options.sheet}'!` : "";

  return `${sheetPrefix}${colPrefix}${colLetter}${rowPrefix}${row}`;
}

/**
 * 构建区域地址
 */
export function buildRangeAddress(
  startCol: number,
  startRow: number,
  endCol: number,
  endRow: number,
  sheet?: string
): string {
  const sheetPrefix = sheet ? `'${sheet}'!` : "";
  const startCell = `${numberToColumnLetter(startCol)}${startRow}`;
  const endCell = `${numberToColumnLetter(endCol)}${endRow}`;

  if (startCol === endCol && startRow === endRow) {
    return `${sheetPrefix}${startCell}`;
  }

  return `${sheetPrefix}${startCell}:${endCell}`;
}

// ==================== 导出 ====================

export default {
  columnLetterToNumber,
  numberToColumnLetter,
  translateFormula,
  isArrayFormula,
  parseCellAddress,
  parseRangeAddress,
  buildCellAddress,
  buildRangeAddress,
};
