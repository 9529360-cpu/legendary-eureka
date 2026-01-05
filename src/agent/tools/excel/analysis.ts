/**
 * Excel 分析类工具
 *
 * 包含工具（未完全迁移，从 ExcelAdapter 导入）：
 * - createAnalyzeDataTool: 数据分析
 * - createGoalSeekTool: 目标求解
 * - createTrendAnalysisTool: 趋势分析
 * - createAnomalyDetectionTool: 异常检测
 * - createDataInsightsTool: 数据洞察
 * - createStatisticalAnalysisTool: 统计分析
 * - createPredictiveAnalysisTool: 预测分析
 * - createProactiveSuggestionsTool: 主动建议
 *
 * @packageDocumentation
 */

import { Tool } from "../../types";

// TODO: 从 ExcelAdapter.ts 迁移工具函数到此文件
// 当前通过 ExcelAdapter.ts 的 createExcelTools() 统一提供

/**
 * 创建所有分析类工具
 * 临时实现：返回空数组，实际工具由 ExcelAdapter 提供
 */
export function createAnalysisTools(): Tool[] {
  // 未迁移，由 ExcelAdapter.createExcelTools() 提供
  return [];
}
