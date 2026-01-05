/**
 * Excel 工作表类工具
 *
 * 包含工具（未完全迁移，从 ExcelAdapter 导入）：
 * - createSheetTool: 获取工作表信息
 * - createCreateSheetTool: 创建工作表
 * - createSwitchSheetTool: 切换工作表
 * - createDeleteSheetTool: 删除工作表
 * - createCopySheetTool: 复制工作表
 * - createRenameSheetTool: 重命名工作表
 * - createProtectSheetTool: 保护工作表
 *
 * @packageDocumentation
 */

import { Tool } from "../../types";

// TODO: 从 ExcelAdapter.ts 迁移工具函数到此文件
// 当前通过 ExcelAdapter.ts 的 createExcelTools() 统一提供

/**
 * 创建所有工作表类工具
 * 临时实现：返回空数组，实际工具由 ExcelAdapter 提供
 */
export function createSheetTools(): Tool[] {
  // 未迁移，由 ExcelAdapter.createExcelTools() 提供
  return [];
}
