/**
 * Excel 格式化类工具
 *
 * 包含工具（未完全迁移，从 ExcelAdapter 导入）：
 * - createFormatRangeTool: 格式化范围
 * - createAutoFitTool: 自动调整列宽
 * - createConditionalFormatTool: 条件格式
 * - createMergeCellsTool: 合并单元格
 * - createBorderTool: 边框设置
 * - createNumberFormatTool: 数字格式
 *
 * @packageDocumentation
 */

import { Tool } from "../../types";

// TODO: 从 ExcelAdapter.ts 迁移工具函数到此文件
// 当前通过 ExcelAdapter.ts 的 createExcelTools() 统一提供

/**
 * 创建所有格式化类工具
 * 临时实现：返回空数组，实际工具由 ExcelAdapter 提供
 */
export function createFormatTools(): Tool[] {
  // 未迁移，由 ExcelAdapter.createExcelTools() 提供
  return [];
}
