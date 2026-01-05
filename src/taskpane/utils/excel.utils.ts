/**
 * Excel 辅助函数
 * v2.9.8: 从 App.tsx 提取的纯函数
 */

import type { SelectionResult, CellValue } from "../types";

// ========== 常量 ==========
export const MAX_CONTEXT_ROWS = 50;
export const MAX_CONTEXT_COLUMNS = 20;

// ========== 基础工具函数 ==========

/**
 * 生成唯一 ID
 */
export function uid(): string {
  return Math.random().toString(16).slice(2);
}

/**
 * 读取 Excel 当前选区
 */
export async function readSelection(): Promise<SelectionResult> {
  if (typeof Excel === "undefined") {
    throw new Error("当前不在 Excel 环境中运行");
  }

  return await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["address", "values", "rowCount", "columnCount"]);
    await context.sync();
    return {
      address: range.address,
      values: range.values as (string | number | boolean | null)[][],
      rowCount: range.rowCount,
      columnCount: range.columnCount,
    };
  });
}

/**
 * 切片选区值（限制行列数）
 */
export function sliceSelectionValues(
  values: CellValue[][],
  maxRows: number = MAX_CONTEXT_ROWS,
  maxColumns: number = MAX_CONTEXT_COLUMNS
): CellValue[][] {
  return values.slice(0, maxRows).map((row) => row.slice(0, maxColumns));
}

/**
 * 构建选区上下文信息
 */
export function buildSelectionContext(selection: SelectionResult) {
  const values = sliceSelectionValues(selection.values);
  const sampleRowCount = values.length;
  const sampleColumnCount = values[0]?.length ?? 0;
  const isTruncated =
    sampleRowCount < selection.rowCount || sampleColumnCount < selection.columnCount;

  return {
    address: selection.address,
    values,
    rowCount: selection.rowCount,
    columnCount: selection.columnCount,
    sampleRowCount,
    sampleColumnCount,
    isTruncated,
  };
}

// ========== 命令参数处理 ==========

/**
 * 从参数中获取范围地址
 */
export function getCommandRangeAddress(parameters: Record<string, unknown>): string | undefined {
  const p = parameters as Record<string, string | undefined>;
  return p.address || p.range || p.rangeAddress || p.dataRange;
}

/**
 * 强制转换值为单元格值类型
 */
export function coerceCellValue(value: unknown): CellValue {
  if (value === null || value === undefined) {
    return null;
  }
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return value;
  }
  return JSON.stringify(value);
}

/**
 * 规范化表头列表
 */
export function normalizeHeaderList(value: unknown): string[] | null {
  if (!Array.isArray(value)) {
    return null;
  }

  const headers = value
    .map((item) => {
      if (typeof item === "string" || typeof item === "number" || typeof item === "boolean") {
        return String(item);
      }
      if (item && typeof item === "object" && "name" in item) {
        return String((item as { name: unknown }).name);
      }
      return null;
    })
    .filter((item): item is string => Boolean(item));

  return headers.length > 0 ? headers : null;
}

/**
 * 从参数中提取表头
 */
export function extractHeaders(parameters: Record<string, unknown>): string[] | null {
  const p = parameters as Record<string, unknown>;
  return (
    normalizeHeaderList(p.headers) ||
    normalizeHeaderList(p.columns) ||
    normalizeHeaderList(p.fields)
  );
}

/**
 * 检查是否存在匹配的表头行
 */
export function hasMatchingHeaderRow(values: CellValue[][], headers: string[]): boolean {
  if (!values.length) return false;
  const firstRow = values[0].map((v) => (v === null ? "" : String(v).toLowerCase().trim()));
  const headerSet = new Set(headers.map((h) => h.toLowerCase().trim()));
  const matchCount = firstRow.filter((c) => headerSet.has(c)).length;
  return matchCount >= headerSet.size * 0.8;
}

/**
 * 获取请求的样本行数
 */
export function getRequestedSampleCount(parameters: Record<string, unknown>): number | null {
  const p = parameters as Record<string, string | number | undefined>;
  const rawCount = p.sampleCount ?? p.rowCount ?? p.count ?? p.rows;

  if (typeof rawCount === "number") {
    return rawCount;
  }
  if (typeof rawCount === "string") {
    const parsed = parseInt(rawCount, 10);
    if (!isNaN(parsed)) return parsed;
  }
  return null;
}

/**
 * 构建样本行数据
 */
export function buildSampleRows(
  headers: string[],
  count: number,
  offset: number = 0
): CellValue[][] {
  const rows: CellValue[][] = [];
  for (let i = 0; i < count; i++) {
    rows.push(headers.map((h) => inferSampleValue(h, i + offset)));
  }
  return rows;
}

/**
 * 合并样本行
 */
export function mergeSampleRows(
  existingRows: CellValue[][],
  additionalRows: CellValue[][],
  maxRows: number = 20
): CellValue[][] {
  const merged = [...existingRows, ...additionalRows];
  return merged.slice(0, maxRows);
}

/**
 * 根据列名推断样本值
 */
export function inferSampleValue(header: string, index: number): CellValue {
  const h = header.toLowerCase();
  const n = index + 1;

  // 日期类
  if (h.includes("date") || h.includes("日期") || h.includes("时间") || h.includes("time")) {
    const d = new Date();
    d.setDate(d.getDate() + index);
    return d.toLocaleDateString("zh-CN");
  }

  // 金额类
  if (
    h.includes("price") ||
    h.includes("价格") ||
    h.includes("金额") ||
    h.includes("amount") ||
    h.includes("cost") ||
    h.includes("费用")
  ) {
    return Math.round(100 + Math.random() * 900);
  }

  // 数量类
  if (
    h.includes("数量") ||
    h.includes("quantity") ||
    h.includes("qty") ||
    h.includes("count") ||
    h.includes("num")
  ) {
    return Math.floor(1 + Math.random() * 100);
  }

  // 百分比
  if (
    h.includes("percent") ||
    h.includes("百分比") ||
    h.includes("比例") ||
    h.includes("rate") ||
    h.includes("ratio")
  ) {
    return `${(Math.random() * 100).toFixed(1)}%`;
  }

  // ID类
  if (h.includes("id") || h.includes("编号") || h.includes("序号") || h.includes("编码")) {
    return `ID-${1000 + n}`;
  }

  // 名称类
  if (h.includes("name") || h.includes("名称") || h.includes("姓名") || h.includes("产品")) {
    return `项目${n}`;
  }

  // 类别类
  if (
    h.includes("category") ||
    h.includes("type") ||
    h.includes("分类") ||
    h.includes("类型") ||
    h.includes("类别")
  ) {
    const categories = ["A类", "B类", "C类"];
    return categories[index % categories.length];
  }

  // 状态类
  if (h.includes("status") || h.includes("状态") || h.includes("state")) {
    const statuses = ["完成", "进行中", "待处理"];
    return statuses[index % statuses.length];
  }

  // 默认：文本
  return `${header}${n}`;
}
