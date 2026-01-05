/**
 * Excel 数据操作类工具
 *
 * 包含工具（未完全迁移，从 ExcelAdapter 导入）：
 * - createSortTool: 排序
 * - createSortRangeTool: 范围排序
 * - createFilterTool: 筛选
 * - createClearRangeTool: 清除范围
 * - createRemoveDuplicatesTool: 删除重复项
 * - createFindReplaceTool: 查找替换
 * - createFillSeriesTool: 填充序列
 * - createInsertRowsTool: 插入行
 * - createDeleteRowsTool: 删除行
 * - createInsertColumnsTool: 插入列
 * - createDeleteColumnsTool: 删除列
 * - createMoveRangeTool: 移动范围
 * - createCopyRangeTool: 复制范围
 *
 * @packageDocumentation
 */

import { Tool } from "../../types";

// TODO: 从 ExcelAdapter.ts 迁移工具函数到此文件
// 当前通过 ExcelAdapter.ts 的 createExcelTools() 统一提供

/**
 * 创建所有数据操作类工具
 * 临时实现：返回空数组，实际工具由 ExcelAdapter 提供
 */
export function createDataTools(): Tool[] {
  // 未迁移，由 ExcelAdapter.createExcelTools() 提供
  return [];
}
