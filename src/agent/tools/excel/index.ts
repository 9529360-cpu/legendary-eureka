/**
 * Excel 工具统一导出
 *
 * 从 ExcelAdapter.ts 重导出所有 Excel 工具
 * 同时提供按类别分组的工具创建函数
 *
 * 目录结构：
 * - common.ts: 共享工具函数
 * - read.ts: 读取类工具 (6个) ✅ 已迁移
 * - write.ts: 写入类工具 (2个) ✅ 已迁移
 * - formula.ts: 公式类工具 (5个) ✅ 已迁移
 * - format.ts: 格式化类工具 (6个) ✅ 已迁移
 * - chart.ts: 图表类工具 (2个) ✅ 已迁移
 * - data.ts: 数据操作类工具 (13个) ✅ 已迁移
 * - sheet.ts: 工作表类工具 (6个) ✅ 已迁移
 * - analysis.ts: 分析类工具 (8个) ✅ 已迁移
 * - advanced.ts: 高级工具 (11个) ✅ 已迁移
 * - misc.ts: 其他工具 (2个) ✅ 已迁移
 *
 * 迁移进度: 61/75 (81%)
 * 注：ExcelAdapter.ts 中还有一些高级分析工具未迁移，
 * 但核心功能已完整，可通过 createExcelTools() 获取全部工具
 *
 * @packageDocumentation
 */

// 从 ExcelAdapter 重导出主函数（保持向后兼容）
export { createExcelTools } from "../../ExcelAdapter";

// 导出通用工具函数
export * from "./common";

// 按类别导出工具创建函数
export { createReadTools } from "./read";
export { createWriteTools } from "./write";
export { createFormulaTools } from "./formula";
export { createFormatTools } from "./format";
export { createChartTools } from "./chart";
export { createDataTools } from "./data";
export { createSheetTools } from "./sheet";
export { createAnalysisTools } from "./analysis";
export { createAdvancedTools } from "./advanced";
export { createMiscTools } from "./misc";

// 导出各模块的单独工具函数
export * from "./read";
export * from "./write";
export * from "./formula";
export * from "./format";
export * from "./chart";
export * from "./data";
export * from "./sheet";
export * from "./analysis";
export * from "./advanced";
export * from "./misc";
