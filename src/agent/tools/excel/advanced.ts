/**
 * Excel 高级工具
 *
 * 包含工具（未完全迁移，从 ExcelAdapter 导入）：
 * - createTableTool: 创建表格
 * - createPivotTableTool: 创建透视表
 * - createFreezePanesTool: 冻结窗格
 * - createGroupRowsTool: 分组行
 * - createGroupColumnsTool: 分组列
 * - createCommentTool: 添加批注
 * - createHyperlinkTool: 添加超链接
 * - createPageSetupTool: 页面设置
 * - createPrintAreaTool: 打印区域
 * - createBatchWriteOptimizedTool: 批量写入优化
 * - createPerformanceModeTool: 性能模式
 * - createRecalculateTool: 重新计算
 * - createAdvancedConditionalFormatTool: 高级条件格式
 * - createClearConditionalFormatsTool: 清除条件格式
 * - createQuickReportTool: 快速报表
 * - createDataChangeListenerTool: 数据变更监听
 * - createGeometricShapeTool: 几何形状
 * - createInsertImageTool: 插入图片
 * - createFindAllTool: 全局查找
 * - createAdvancedCopyTool: 高级复制
 * - createMoveRangeAdvancedTool: 高级移动
 * - createNamedRangeTool: 命名范围
 * - createInsertExternalSheetsTool: 插入外部工作表
 * - createDataValidationTool: 数据验证
 *
 * @packageDocumentation
 */

import { Tool } from "../../types";

// TODO: 从 ExcelAdapter.ts 迁移工具函数到此文件
// 当前通过 ExcelAdapter.ts 的 createExcelTools() 统一提供

/**
 * 创建所有高级工具
 * 临时实现：返回空数组，实际工具由 ExcelAdapter 提供
 */
export function createAdvancedTools(): Tool[] {
  // 未迁移，由 ExcelAdapter.createExcelTools() 提供
  return [];
}
