/**
 * Excel 工具共享代码
 *
 * 从 ExcelAdapter.ts 抽取的通用辅助函数和类型
 *
 * @packageDocumentation
 */

import type { Tool, ToolResult } from "../../types";

// 重导出类型供工具模块使用
export type { Tool, ToolResult };

// ========== Excel.run 封装 ==========

/**
 * 安全执行 Excel.run 并处理错误
 */
export async function excelRun<T extends ToolResult>(
  callback: (ctx: Excel.RequestContext) => Promise<T>
): Promise<T> {
  try {
    return await Excel.run(async (ctx) => {
      return await callback(ctx);
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return {
      success: false,
      output: `Excel 操作失败: ${message}`,
      error: message,
    } as T;
  }
}

// ========== Sheet 辅助函数 ==========

/**
 * 获取指定工作表或活动工作表
 */
export function getTargetSheet(
  ctx: Excel.RequestContext,
  sheetName?: string | null
): Excel.Worksheet {
  if (sheetName && typeof sheetName === "string" && sheetName.trim()) {
    return ctx.workbook.worksheets.getItem(sheetName.trim());
  }
  return ctx.workbook.worksheets.getActiveWorksheet();
}

/**
 * 从 input 中提取 sheet 参数
 */
export function extractSheetName(input: Record<string, unknown>): string | null {
  const sheet = input.sheet || input.sheetName || input.worksheet || input.targetSheet;
  if (sheet && typeof sheet === "string" && sheet.trim()) {
    return sheet.trim();
  }
  return null;
}

// ========== 范围解析辅助函数 ==========

/**
 * 智能提取地址参数（兼容多种写法）
 */
export function extractAddress(input: Record<string, unknown>, defaultValue = "A1:A10"): string {
  return String(
    input.address || input.range || input.cell || input.area || input.targetRange || defaultValue
  );
}

/**
 * 解析范围地址获取行列信息
 */
export function parseRangeAddress(address: string): {
  startCol: string;
  startRow: number;
  endCol: string;
  endRow: number;
} | null {
  const match = address.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/i);
  if (!match) return null;
  return {
    startCol: match[1].toUpperCase(),
    startRow: parseInt(match[2], 10),
    endCol: match[3].toUpperCase(),
    endRow: parseInt(match[4], 10),
  };
}

/**
 * 列字母转数字 (A=1, B=2, ..., Z=26, AA=27)
 */
export function columnLetterToNumber(letter: string): number {
  let result = 0;
  for (let i = 0; i < letter.length; i++) {
    result = result * 26 + (letter.charCodeAt(i) - 64);
  }
  return result;
}

/**
 * 数字转列字母 (1=A, 2=B, ..., 26=Z, 27=AA)
 */
export function numberToColumnLetter(num: number): string {
  let result = "";
  while (num > 0) {
    const remainder = (num - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    num = Math.floor((num - 1) / 26);
  }
  return result;
}

// ========== 数据格式化辅助函数 ==========

/**
 * 格式化数据预览
 */
export function formatDataPreview(values: unknown[][], maxRows = 10, maxCols = 10): string {
  return values
    .slice(0, maxRows)
    .map((row) =>
      row
        .slice(0, maxCols)
        .map((cell) => (cell === null || cell === undefined ? "" : String(cell)))
        .join("\t")
    )
    .join("\n");
}

/**
 * 安全获取字符串值
 */
export function safeString(value: unknown, defaultValue = ""): string {
  if (value === null || value === undefined) return defaultValue;
  return String(value);
}

/**
 * 安全获取数字值
 */
export function safeNumber(value: unknown, defaultValue = 0): number {
  if (value === null || value === undefined) return defaultValue;
  const num = Number(value);
  return isNaN(num) ? defaultValue : num;
}

/**
 * 安全获取布尔值
 */
export function safeBoolean(value: unknown, defaultValue = false): boolean {
  if (value === null || value === undefined) return defaultValue;
  if (typeof value === "boolean") return value;
  if (typeof value === "string") {
    return value.toLowerCase() === "true" || value === "1";
  }
  return Boolean(value);
}

// ========== 颜色处理辅助函数 ==========

/**
 * 标准化颜色值
 */
export function normalizeColor(color: unknown): string | null {
  if (!color) return null;
  const colorStr = String(color).trim();

  // 如果已经是 # 开头的十六进制颜色
  if (/^#[0-9A-Fa-f]{6}$/.test(colorStr)) {
    return colorStr;
  }

  // 如果是不带 # 的十六进制颜色
  if (/^[0-9A-Fa-f]{6}$/.test(colorStr)) {
    return `#${colorStr}`;
  }

  // 常见颜色名称映射
  const colorMap: Record<string, string> = {
    red: "#FF0000",
    green: "#00FF00",
    blue: "#0000FF",
    yellow: "#FFFF00",
    orange: "#FFA500",
    purple: "#800080",
    pink: "#FFC0CB",
    black: "#000000",
    white: "#FFFFFF",
    gray: "#808080",
    grey: "#808080",
  };

  return colorMap[colorStr.toLowerCase()] || null;
}
