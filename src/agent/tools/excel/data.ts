/**
 * Excel 数据操作类工具
 *
 * 包含工具：
 * - createSortTool: 排序
 * - createSortRangeTool: 范围排序（别名）
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

/* eslint-disable office-addins/call-sync-after-load, office-addins/call-sync-before-read */

import { Tool } from "../../types";
import { excelRun } from "./common";

// ========== 排序筛选工具 ==========

/**
 * 排序工具
 */
export function createSortTool(): Tool {
  return {
    name: "excel_sort",
    description: "对数据进行排序",
    category: "excel",
    parameters: [
      { name: "address", type: "string", description: "排序范围", required: true },
      {
        name: "column",
        type: "number",
        description: "按第几列排序（从0开始，或列字母如A/B/C）",
        required: true,
      },
      { name: "ascending", type: "boolean", description: "是否升序", required: false },
    ],
    execute: async (input) => {
      const address = String(input.address || input.range || input.data || "A1:D10");

      let column = 0;
      const colInput = input.column || input.sortColumn || input.key || 0;
      if (typeof colInput === "string" && /^[A-Za-z]+$/.test(colInput)) {
        column = colInput.toUpperCase().charCodeAt(0) - 65;
      } else {
        column = Number(colInput);
      }

      const ascending = input.ascending !== false && input.descending !== true;

      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(address);

        range.load("address");
        await ctx.sync();
        const rangeAddress = range.address;

        range.sort.apply([{ key: column, ascending }]);
        await ctx.sync();

        return {
          success: true,
          output: `已按第${column + 1}列（${String.fromCharCode(65 + column)}列）${ascending ? "升序" : "降序"}排序 ${rangeAddress}`,
          data: { address: rangeAddress, column, ascending },
        };
      });
    },
  };
}

/**
 * 排序工具别名
 */
export function createSortRangeTool(): Tool {
  const baseTool = createSortTool();
  return {
    ...baseTool,
    name: "excel_sort_range",
    description: "对数据范围进行排序（excel_sort 的别名）",
  };
}

/**
 * 筛选工具
 */
export function createFilterTool(): Tool {
  return {
    name: "excel_filter",
    description: "筛选数据",
    category: "excel",
    parameters: [
      { name: "address", type: "string", description: "数据范围", required: true },
      { name: "column", type: "number", description: "筛选列（从0开始）", required: true },
      { name: "criteria", type: "string", description: "筛选条件值", required: true },
    ],
    execute: async (input) => {
      const address = String(input.address || input.range || input.data || "A1:D10");
      const colInput = input.column || input.filterColumn || 0;
      let column = 0;
      if (typeof colInput === "string" && /^[A-Za-z]+$/.test(colInput)) {
        column = colInput.toUpperCase().charCodeAt(0) - 65;
      } else {
        column = Number(colInput);
      }
      const criteria = String(input.criteria || input.value || input.filter || "");

      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(address);
        range.load("address, rowCount, columnCount");
        await ctx.sync();
        const rangeAddress = range.address;

        try {
          sheet.autoFilter.apply(range, column);
          await ctx.sync();

          return {
            success: true,
            output: `已在 ${rangeAddress} 上应用自动筛选：第${column + 1}列`,
            data: { address: rangeAddress, column, criteria },
          };
        } catch (filterError) {
          try {
            sheet.autoFilter.remove();
            await ctx.sync();

            sheet.autoFilter.apply(range, column);
            await ctx.sync();

            return {
              success: true,
              output: `已在 ${rangeAddress} 上应用筛选：第${column + 1}列 = "${criteria}"`,
              data: { address: rangeAddress, column, criteria },
            };
          } catch (_retryError) {
            return {
              success: false,
              output: `筛选失败：${filterError instanceof Error ? filterError.message : String(filterError)}`,
              error: String(filterError),
            };
          }
        }
      });
    },
  };
}

// ========== 清除与去重工具 ==========

/**
 * 清空范围工具
 */
export function createClearRangeTool(): Tool {
  return {
    name: "excel_clear_range",
    description: "清空指定范围",
    category: "excel",
    parameters: [
      { name: "address", type: "string", description: "要清空的范围", required: true },
      {
        name: "type",
        type: "string",
        description: "清空类型: contents/formats/all",
        required: false,
      },
    ],
    execute: async (input) => {
      const address = String(input.address || input.range || input.cell || "A1:A10");
      const type = String(input.type || input.clearType || "contents");

      return await excelRun(async (ctx) => {
        const range = ctx.workbook.worksheets.getActiveWorksheet().getRange(address);

        if (type === "all") {
          range.clear();
        } else if (type === "formats") {
          range.clear(Excel.ClearApplyTo.formats);
        } else {
          range.clear(Excel.ClearApplyTo.contents);
        }

        await ctx.sync();

        return {
          success: true,
          output: `已清空 ${address} 的${type === "all" ? "全部内容" : type === "formats" ? "格式" : "内容"}`,
        };
      });
    },
  };
}

/**
 * 删除重复行工具
 */
export function createRemoveDuplicatesTool(): Tool {
  return {
    name: "excel_remove_duplicates",
    description: "删除重复行",
    category: "excel",
    parameters: [
      { name: "address", type: "string", description: "数据范围", required: true },
      { name: "columns", type: "array", description: "用于判断重复的列索引", required: false },
    ],
    execute: async (input) => {
      const address = String(input.address || input.range || input.data || "A1:D10");
      const columns = (input.columns as number[]) || (input.compareColumns as number[]) || [0];

      return await excelRun(async (ctx) => {
        const range = ctx.workbook.worksheets.getActiveWorksheet().getRange(address);
        range.removeDuplicates(columns, true);
        await ctx.sync();

        return {
          success: true,
          output: `已从 ${address} 删除重复行`,
        };
      });
    },
  };
}

// ========== 查找替换与填充工具 ==========

/**
 * 查找替换工具
 */
export function createFindReplaceTool(): Tool {
  return {
    name: "excel_find_replace",
    description: "在指定范围内查找并替换文本",
    category: "excel",
    parameters: [
      { name: "sheet", type: "string", description: "工作表名称", required: true },
      {
        name: "range",
        type: "string",
        description: "搜索范围，如 A1:Z1000，留空表示整个工作表",
        required: false,
      },
      { name: "find", type: "string", description: "要查找的文本", required: true },
      { name: "replace", type: "string", description: "替换为的文本", required: true },
      { name: "matchCase", type: "boolean", description: "是否区分大小写", required: false },
    ],
    execute: async (params) => {
      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(params.sheet as string);
        const rangeAddress = (params.range as string) || null;
        const range = rangeAddress ? sheet.getRange(rangeAddress) : sheet.getUsedRange();

        range.load("values");
        await ctx.sync();

        const findText = params.find as string;
        const replaceText = params.replace as string;
        const matchCase = (params.matchCase as boolean) ?? false;
        let replaceCount = 0;

        const values = range.values;
        for (let r = 0; r < values.length; r++) {
          for (let c = 0; c < values[r].length; c++) {
            const cellValue = String(values[r][c]);
            if (
              matchCase
                ? cellValue.includes(findText)
                : cellValue.toLowerCase().includes(findText.toLowerCase())
            ) {
              values[r][c] = matchCase
                ? cellValue.replace(new RegExp(findText, "g"), replaceText)
                : cellValue.replace(new RegExp(findText, "gi"), replaceText);
              replaceCount++;
            }
          }
        }

        range.values = values;
        await ctx.sync();

        return {
          success: true,
          output: `已替换 ${replaceCount} 处: "${findText}"  "${replaceText}"`,
        };
      });
    },
  };
}

/**
 * 填充序列工具
 */
export function createFillSeriesTool(): Tool {
  return {
    name: "excel_fill_series",
    description: "自动填充序列（日期、数字等）",
    category: "excel",
    parameters: [
      { name: "sheet", type: "string", description: "工作表名称", required: true },
      { name: "startCell", type: "string", description: "起始单元格，如 A1", required: true },
      { name: "endCell", type: "string", description: "结束单元格，如 A10", required: true },
      {
        name: "type",
        type: "string",
        description: "序列类型: number(数字)/date(日期)/text(文本+数字)",
        required: false,
      },
      { name: "step", type: "number", description: "步长（数字序列）", required: false },
    ],
    execute: async (params) => {
      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(params.sheet as string);
        const startCell = sheet.getRange(params.startCell as string);
        const endCell = sheet.getRange(params.endCell as string);

        startCell.load("values, address");
        endCell.load("address");
        await ctx.sync();
        const startValue = startCell.values[0][0];
        const rangeAddress = `${params.startCell}:${params.endCell}`;
        const range = sheet.getRange(rangeAddress);
        range.load("rowCount, columnCount");
        await ctx.sync();
        const step = (params.step as number) || 1;
        const rowCount = range.rowCount;
        const columnCount = range.columnCount;
        const isVertical = rowCount > columnCount;
        const count = isVertical ? rowCount : columnCount;

        const values: unknown[][] = [];
        for (let i = 0; i < (isVertical ? count : 1); i++) {
          const row: unknown[] = [];
          for (let j = 0; j < (isVertical ? 1 : count); j++) {
            const idx = isVertical ? i : j;
            if (typeof startValue === "number") {
              row.push(startValue + idx * step);
            } else {
              const match = String(startValue).match(/^(.*)(\d+)$/);
              if (match) {
                row.push(`${match[1]}${parseInt(match[2]) + idx * step}`);
              } else {
                row.push(startValue);
              }
            }
          }
          values.push(row);
        }

        range.values = values;
        await ctx.sync();

        return {
          success: true,
          output: `已填充序列 ${rangeAddress}，共 ${count} 个值`,
        };
      });
    },
  };
}

// ========== 行列操作工具 ==========

/**
 * 插入行工具
 */
export function createInsertRowsTool(): Tool {
  return {
    name: "excel_insert_rows",
    description: "在指定位置插入空白行，原数据会向下移动",
    category: "excel",
    parameters: [
      {
        name: "rowIndex",
        type: "number",
        description: "要插入的行号（从1开始），如 2 表示在第2行位置插入",
        required: true,
      },
      { name: "count", type: "number", description: "要插入的行数，默认为1", required: false },
      {
        name: "sheet",
        type: "string",
        description: "工作表名称，默认为当前活动工作表",
        required: false,
      },
    ],
    execute: async (params) => {
      const rowIndex = Number(params.rowIndex);
      const count = Number(params.count || 1);
      const sheetName = params.sheet as string | undefined;

      if (rowIndex < 1) {
        return { success: false, output: "行号必须大于等于1", error: "Invalid row index" };
      }

      return await excelRun(async (ctx) => {
        const sheet = sheetName
          ? ctx.workbook.worksheets.getItem(sheetName)
          : ctx.workbook.worksheets.getActiveWorksheet();

        const range = sheet.getRange(`${rowIndex}:${rowIndex + count - 1}`);
        range.insert(Excel.InsertShiftDirection.down);
        await ctx.sync();

        return {
          success: true,
          output: `已在第 ${rowIndex} 行位置插入 ${count} 行空白行，原数据已向下移动`,
        };
      });
    },
  };
}

/**
 * 删除行工具
 */
export function createDeleteRowsTool(): Tool {
  return {
    name: "excel_delete_rows",
    description: "删除指定的行，下方数据会向上移动",
    category: "excel",
    parameters: [
      { name: "startRow", type: "number", description: "起始行号（从1开始）", required: true },
      {
        name: "endRow",
        type: "number",
        description: "结束行号，默认与起始行相同（只删除一行）",
        required: false,
      },
      {
        name: "sheet",
        type: "string",
        description: "工作表名称，默认为当前活动工作表",
        required: false,
      },
    ],
    execute: async (params) => {
      const startRow = Number(params.startRow);
      const endRow = Number(params.endRow || startRow);
      const sheetName = params.sheet as string | undefined;

      if (startRow < 1 || endRow < startRow) {
        return { success: false, output: "无效的行号范围", error: "Invalid row range" };
      }

      return await excelRun(async (ctx) => {
        const sheet = sheetName
          ? ctx.workbook.worksheets.getItem(sheetName)
          : ctx.workbook.worksheets.getActiveWorksheet();

        const range = sheet.getRange(`${startRow}:${endRow}`);
        range.delete(Excel.DeleteShiftDirection.up);
        await ctx.sync();

        const count = endRow - startRow + 1;
        return {
          success: true,
          output: `已删除第 ${startRow}${endRow > startRow ? ` 到第 ${endRow}` : ""} 行，共 ${count} 行`,
        };
      });
    },
  };
}

/**
 * 插入列工具
 */
export function createInsertColumnsTool(): Tool {
  return {
    name: "excel_insert_columns",
    description: "在指定位置插入空白列，原数据会向右移动",
    category: "excel",
    parameters: [
      {
        name: "column",
        type: "string",
        description: "要插入的列字母，如 B 表示在B列位置插入",
        required: true,
      },
      { name: "count", type: "number", description: "要插入的列数，默认为1", required: false },
      {
        name: "sheet",
        type: "string",
        description: "工作表名称，默认为当前活动工作表",
        required: false,
      },
    ],
    execute: async (params) => {
      const column = String(params.column).toUpperCase();
      const count = Number(params.count || 1);
      const sheetName = params.sheet as string | undefined;

      return await excelRun(async (ctx) => {
        const sheet = sheetName
          ? ctx.workbook.worksheets.getItem(sheetName)
          : ctx.workbook.worksheets.getActiveWorksheet();

        const startColCode = column.charCodeAt(0);
        const endCol = String.fromCharCode(startColCode + count - 1);
        const range = sheet.getRange(`${column}:${endCol}`);
        range.insert(Excel.InsertShiftDirection.right);
        await ctx.sync();

        return {
          success: true,
          output: `已在 ${column} 列位置插入 ${count} 列空白列，原数据已向右移动`,
        };
      });
    },
  };
}

/**
 * 删除列工具
 */
export function createDeleteColumnsTool(): Tool {
  return {
    name: "excel_delete_columns",
    description: "删除指定的列，右侧数据会向左移动",
    category: "excel",
    parameters: [
      { name: "startColumn", type: "string", description: "起始列字母，如 A", required: true },
      {
        name: "endColumn",
        type: "string",
        description: "结束列字母，默认与起始列相同（只删除一列）",
        required: false,
      },
      {
        name: "sheet",
        type: "string",
        description: "工作表名称，默认为当前活动工作表",
        required: false,
      },
    ],
    execute: async (params) => {
      const startCol = String(params.startColumn).toUpperCase();
      const endCol = params.endColumn ? String(params.endColumn).toUpperCase() : startCol;
      const sheetName = params.sheet as string | undefined;

      return await excelRun(async (ctx) => {
        const sheet = sheetName
          ? ctx.workbook.worksheets.getItem(sheetName)
          : ctx.workbook.worksheets.getActiveWorksheet();

        const range = sheet.getRange(`${startCol}:${endCol}`);
        range.delete(Excel.DeleteShiftDirection.left);
        await ctx.sync();

        return {
          success: true,
          output: `已删除 ${startCol}${endCol !== startCol ? ` 到 ${endCol}` : ""} 列`,
        };
      });
    },
  };
}

// ========== 移动复制工具 ==========

/**
 * 移动范围工具（剪切+粘贴）
 */
export function createMoveRangeTool(): Tool {
  return {
    name: "excel_move_range",
    description:
      "将一个范围的数据移动到另一个位置（剪切粘贴操作）。注意：如果需要将行数据移到上方，应先用 excel_insert_rows 插入空行，再移动数据，最后删除原行",
    category: "excel",
    parameters: [
      {
        name: "sourceAddress",
        type: "string",
        description: "源范围地址，如 A22:J22",
        required: true,
      },
      { name: "targetCell", type: "string", description: "目标起始单元格，如 A2", required: true },
      {
        name: "sheet",
        type: "string",
        description: "工作表名称，默认为当前活动工作表",
        required: false,
      },
    ],
    execute: async (params) => {
      const sourceAddress = String(params.sourceAddress);
      const targetCell = String(params.targetCell);
      const sheetName = params.sheet as string | undefined;

      return await excelRun(async (ctx) => {
        const sheet = sheetName
          ? ctx.workbook.worksheets.getItem(sheetName)
          : ctx.workbook.worksheets.getActiveWorksheet();

        const sourceRange = sheet.getRange(sourceAddress);
        const targetRange = sheet.getRange(targetCell);

        sourceRange.load("values, formulas, numberFormat, rowCount, columnCount");
        await ctx.sync();
        const _values = sourceRange.values;
        const formulas = sourceRange.formulas;
        const formats = sourceRange.numberFormat;
        const rowCount = sourceRange.rowCount;
        const colCount = sourceRange.columnCount;

        const targetFullRange = targetRange.getResizedRange(rowCount - 1, colCount - 1);

        targetFullRange.formulas = formulas;
        targetFullRange.numberFormat = formats;
        await ctx.sync();

        sourceRange.clear(Excel.ClearApplyTo.contents);
        await ctx.sync();

        return {
          success: true,
          output: `已将 ${sourceAddress} 的数据移动到 ${targetCell}`,
        };
      });
    },
  };
}

/**
 * 复制范围工具
 */
export function createCopyRangeTool(): Tool {
  return {
    name: "excel_copy_range",
    description: "将一个范围的数据复制到另一个位置",
    category: "excel",
    parameters: [
      {
        name: "sourceAddress",
        type: "string",
        description: "源范围地址，如 A1:D10",
        required: true,
      },
      { name: "targetCell", type: "string", description: "目标起始单元格，如 F1", required: true },
      {
        name: "sheet",
        type: "string",
        description: "工作表名称，默认为当前活动工作表",
        required: false,
      },
      {
        name: "targetSheet",
        type: "string",
        description: "目标工作表名称，默认与源工作表相同",
        required: false,
      },
    ],
    execute: async (params) => {
      const sourceAddress = String(params.sourceAddress);
      const targetCell = String(params.targetCell);
      const sheetName = params.sheet as string | undefined;
      const targetSheetName = params.targetSheet as string | undefined;

      return await excelRun(async (ctx) => {
        const sourceSheet = sheetName
          ? ctx.workbook.worksheets.getItem(sheetName)
          : ctx.workbook.worksheets.getActiveWorksheet();

        const targetSheet = targetSheetName
          ? ctx.workbook.worksheets.getItem(targetSheetName)
          : sourceSheet;

        const sourceRange = sourceSheet.getRange(sourceAddress);
        const targetRange = targetSheet.getRange(targetCell);

        sourceRange.load("values, formulas, numberFormat, rowCount, columnCount");
        await ctx.sync();
        const formulas = sourceRange.formulas;
        const formats = sourceRange.numberFormat;
        const rowCount = sourceRange.rowCount;
        const colCount = sourceRange.columnCount;

        const targetFullRange = targetRange.getResizedRange(rowCount - 1, colCount - 1);

        targetFullRange.formulas = formulas;
        targetFullRange.numberFormat = formats;
        await ctx.sync();

        return {
          success: true,
          output: `已将 ${sourceAddress} 的数据复制到 ${targetSheetName ? targetSheetName + "!" : ""}${targetCell}`,
        };
      });
    },
  };
}

/**
 * 创建所有数据操作类工具
 */
export function createDataTools(): Tool[] {
  return [
    createSortTool(),
    createSortRangeTool(),
    createFilterTool(),
    createClearRangeTool(),
    createRemoveDuplicatesTool(),
    createFindReplaceTool(),
    createFillSeriesTool(),
    createInsertRowsTool(),
    createDeleteRowsTool(),
    createInsertColumnsTool(),
    createDeleteColumnsTool(),
    createMoveRangeTool(),
    createCopyRangeTool(),
  ];
}
