/**
 * Excel 工具统一导出
 *
 * 从 ExcelAdapter.ts 重导出所有 Excel 工具
 * 未来可以逐步将工具函数迁移到各个类别文件中
 *
 * 目录结构（规划中）：
 * - read.ts: 读取类工具 (6个)
 * - write.ts: 写入类工具 (2个)
 * - formula.ts: 公式类工具 (5个)
 * - format.ts: 格式化类工具 (6个)
 * - chart.ts: 图表类工具 (2个)
 * - data.ts: 数据操作类工具 (13个)
 * - sheet.ts: 工作表类工具 (7个)
 * - table.ts: 表格类工具 (2个)
 * - view.ts: 视图类工具 (3个)
 * - analysis.ts: 分析类工具 (7个)
 * - advanced.ts: 高级工具 (16个)
 * - misc.ts: 其他工具 (2个)
 *
 * @packageDocumentation
 */

// 从 ExcelAdapter 重导出主函数
export { createExcelTools } from "../../ExcelAdapter";

// 导出通用工具函数
export * from "./common";

// 按类别导出（逐步迁移中）
// TODO: 迁移完成后取消注释
// export * from "./read";
// export * from "./write";
// export * from "./formula";
// export * from "./format";
// export * from "./chart";
// export * from "./data";
// export * from "./sheet";
// export * from "./table";
// export * from "./view";
// export * from "./analysis";
// export * from "./advanced";
// export * from "./misc";
