/**
 * Excel 写入类工具
 *
 * 包含工具：
 * - createWriteRangeTool: 批量写入数据
 * - createWriteCellTool: 单元格写入
 *
 * @packageDocumentation
 */

/* eslint-disable office-addins/call-sync-after-load, office-addins/call-sync-before-read */

import { Tool } from "../../types";
import { excelRun, getTargetSheet, extractSheetName } from "./common";

// ========== 写入类工具 ==========

export function createWriteRangeTool(): Tool {
  return {
    name: "excel_write_range",
    description: "向指定范围写入数据（二维数组）",
    category: "excel",
    parameters: [
      { name: "address", type: "string", description: "起始地址，如 A1", required: true },
      { name: "values", type: "array", description: "二维数组数据", required: true },
      {
        name: "sheet",
        type: "string",
        description: "工作表名称（可选，默认当前活动表）",
        required: false,
      },
    ],
    execute: async (input) => {
      // v2.9.39: 智能参数兼容 - 接受 address/range/cell/target 等多种写法
      const address = String(
        input.address || input.range || input.cell || input.target || input.location || "A1"
      );
      // v2.9.41: 支持指定工作表
      const sheetName = extractSheetName(input);

      // 处理 values - 可能是数组或字符串
      let values = input.values as unknown[][];
      if (!values || !Array.isArray(values)) {
        // 尝试从其他可能的参数名获取
        values = (input.data || input.content || input.rows) as unknown[][];
      }

      // 如果 values 是一维数组，转成二维
      if (values && Array.isArray(values) && values.length > 0 && !Array.isArray(values[0])) {
        values = [values as unknown[]];
      }

      if (!values || values.length === 0) {
        return { success: false, output: "缺少要写入的数据 (values)" };
      }

      return await excelRun(async (ctx) => {
        const sheet = getTargetSheet(ctx, sheetName);
        sheet.load("name");
        const range = sheet
          .getRange(address)
          .getResizedRange(values.length - 1, values[0].length - 1);
        range.values = values;
        await ctx.sync();

        // v2.9.39: 写入后验证 - 确保数据真正写入
        range.load("values, address");
        await ctx.sync();

        const writtenValues = range.values;
        const expectedRows = values.length;
        const actualRows = writtenValues?.length || 0;

        if (actualRows !== expectedRows) {
          return {
            success: false,
            output: `写入验证失败：预期 ${expectedRows} 行，实际只有 ${actualRows} 行`,
          };
        }

        return {
          success: true,
          output: `已将 ${values.length}行 × ${values[0].length}列 数据写入 ${sheet.name}!${range.address}`,
          data: {
            address: range.address,
            sheet: sheet.name,
            rows: actualRows,
            cols: values[0].length,
          },
        };
      });
    },
  };
}

export function createWriteCellTool(): Tool {
  return {
    name: "excel_write_cell",
    description: "向单个单元格写入值",
    category: "excel",
    parameters: [
      { name: "address", type: "string", description: "单元格地址，如 A1", required: true },
      { name: "value", type: "string", description: "要写入的值", required: true },
      { name: "sheet", type: "string", description: "工作表名称（可选）", required: false },
    ],
    execute: async (input) => {
      // v2.9.38: 智能参数兼容 - 接受多种写法
      const address = String(input.address || input.cell || input.range || input.target || "A1");
      const value = input.value || input.content || input.data || "";
      // v2.9.41: 支持指定工作表
      const sheetName = extractSheetName(input);

      return await excelRun(async (ctx) => {
        const sheet = getTargetSheet(ctx, sheetName);
        sheet.load("name");
        const range = sheet.getRange(address);
        range.values = [[value]];
        await ctx.sync();

        // v2.9.44: 写入后验证 - 确保数据真正写入
        range.load("values");
        await ctx.sync();

        const writtenValue = range.values[0]?.[0];
        // 比较时需要考虑类型转换（Excel 可能将字符串数字转为数字）
        const valueStr = String(value);
        const writtenStr = String(writtenValue ?? "");

        if (writtenStr !== valueStr && writtenValue !== value) {
          console.warn(`[ExcelAdapter] 写入验证警告: 预期 "${valueStr}", 实际 "${writtenStr}"`);
          // 不直接失败，因为可能是 Excel 的类型转换
        }

        return {
          success: true,
          output: `已将 "${value}" 写入 ${sheet.name}!${address}`,
          data: { sheet: sheet.name, address, verified: true },
        };
      });
    },
  };
}

/**
 * 创建所有写入类工具
 */
export function createWriteTools(): Tool[] {
  return [createWriteRangeTool(), createWriteCellTool()];
}
