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
 * - format.ts: 格式化类工具 (6个) 🔄 骨架
 * - chart.ts: 图表类工具 (2个) 🔄 骨架
 * - data.ts: 数据操作类工具 (13个) 🔄 骨架
 * - sheet.ts: 工作表类工具 (7个) 🔄 骨架
 * - analysis.ts: 分析类工具 (8个) 🔄 骨架
 * - advanced.ts: 高级工具 (24个) 🔄 骨架
 * - misc.ts: 其他工具 (2个) 🔄 骨架
 *
 * 迁移进度: 13/75 (17%)
 *
 * @packageDocumentation
 */

// 从 ExcelAdapter 重导出主函数（保持向后兼容）
export { createExcelTools } from "../../ExcelAdapter";

// 导出通用工具函数
export * from "./common";

// 按类别导出（已完成迁移的）
export { createReadTools } from "./read";
export { createWriteTools } from "./write";
export { createFormulaTools } from "./formula";

// 按类别导出（骨架文件，实际工具由 ExcelAdapter 提供）
export { createFormatTools } from "./format";
export { createChartTools } from "./chart";
export { createDataTools } from "./data";
export { createSheetTools } from "./sheet";
export { createAnalysisTools } from "./analysis";
export { createAdvancedTools } from "./advanced";
export { createMiscTools } from "./misc";
