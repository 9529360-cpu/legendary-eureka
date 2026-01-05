/**
 * Excel 其他工具
 *
 * 包含工具（未完全迁移，从 ExcelAdapter 导入）：
 * - createRespondToUserTool: 响应用户
 * - createClarifyRequestTool: 澄清请求
 *
 * @packageDocumentation
 */

import { Tool } from "../../types";

// TODO: 从 ExcelAdapter.ts 迁移工具函数到此文件
// 当前通过 ExcelAdapter.ts 的 createExcelTools() 统一提供

/**
 * 创建所有其他工具
 * 临时实现：返回空数组，实际工具由 ExcelAdapter 提供
 */
export function createMiscTools(): Tool[] {
  // 未迁移，由 ExcelAdapter.createExcelTools() 提供
  return [];
}
