/**
 * Excel 读取类工具
 *
 * 包含工具：
 * - createReadSelectionTool: 读取当前选区
 * - createReadRangeTool: 读取指定范围
 * - createGetWorkbookInfoTool: 获取工作簿信息
 * - createGetTableSchemaTool: 获取表格结构
 * - createSampleRowsTool: 获取样本数据
 * - createGetSheetInfoTool: 获取工作表信息
 *
 * @packageDocumentation
 */

/* eslint-disable office-addins/call-sync-after-load, office-addins/call-sync-before-read */

import { Tool } from "../../types";
import { excelRun, getTargetSheet, extractSheetName } from "./common";

// ========== 读取类工具 ==========

export function createReadSelectionTool(): Tool {
  return {
    name: "excel_read_selection",
    description: "读取当前 Excel 选区的数据",
    category: "excel",
    parameters: [],
    execute: async () => {
      return await excelRun(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        range.load("address, values, rowCount, columnCount");
        await ctx.sync();

        const preview = range.values
          .slice(0, 10)
          .map((row: unknown[]) => row.slice(0, 10).map(String).join("\t"))
          .join("\n");

        return {
          success: true,
          output: `选区 ${range.address}（${range.rowCount}行 × ${range.columnCount}列）\n数据预览:\n${preview}`,
          data: {
            address: range.address,
            values: range.values,
            rowCount: range.rowCount,
            columnCount: range.columnCount,
          },
        };
      });
    },
  };
}

export function createReadRangeTool(): Tool {
  return {
    name: "excel_read_range",
    description: "读取指定范围的数据",
    category: "excel",
    parameters: [
      { name: "address", type: "string", description: "范围地址，如 A1:D10", required: true },
      {
        name: "sheet",
        type: "string",
        description: "工作表名称（可选，默认当前活动表）",
        required: false,
      },
    ],
    execute: async (input) => {
      // v2.9.38: 智能参数兼容 - 接受 address/range/selection 等多种写法
      const address = String(input.address || input.range || input.cell || input.area || "A1:A10");
      // v2.9.41: 支持指定工作表
      const sheetName = extractSheetName(input);

      return await excelRun(async (ctx) => {
        const sheet = getTargetSheet(ctx, sheetName);
        const range = sheet.getRange(address);
        range.load("address, values, rowCount, columnCount");
        sheet.load("name");
        await ctx.sync();

        return {
          success: true,
          output: `已读取 ${sheet.name}!${range.address}（${range.rowCount}行 × ${range.columnCount}列）`,
          data: { values: range.values, sheet: sheet.name },
        };
      });
    },
  };
}

export function createGetWorkbookInfoTool(): Tool {
  return {
    name: "excel_get_workbook_info",
    description: "获取工作簿的整体信息（工作表列表、表格、图表等）",
    category: "excel",
    parameters: [],
    execute: async () => {
      return await excelRun(async (ctx) => {
        const sheets = ctx.workbook.worksheets;
        sheets.load("items/name, items/position");

        const tables = ctx.workbook.tables;
        tables.load("items/name");

        const activeSheet = ctx.workbook.worksheets.getActiveWorksheet();
        activeSheet.load("name");

        await ctx.sync();

        const sheetNames = sheets.items.map((s: { name: string }) => s.name);
        const tableNames = tables.items.map((t: { name: string }) => t.name);

        return {
          success: true,
          output: `工作簿信息:\n- 工作表: ${sheetNames.join(", ")}\n- 当前工作表: ${activeSheet.name}\n- 表格: ${tableNames.length > 0 ? tableNames.join(", ") : "无"}`,
          data: {
            sheets: sheetNames,
            activeSheet: activeSheet.name,
            tables: tableNames,
          },
        };
      });
    },
  };
}

// ========== 辅助函数 ==========

/**
 * v3.0.2: 推断数据类型
 */
function inferDataType(values: unknown[]): string {
  const validValues = values.filter((v) => v !== null && v !== undefined && v !== "");
  if (validValues.length === 0) return "empty";

  const firstValue = validValues[0];

  // 数字类型
  if (typeof firstValue === "number") {
    // 检查是否为日期序列号 (Excel 存储日期为数字)
    if (firstValue > 25569 && firstValue < 50000) {
      return "date (number)";
    }
    return "number";
  }

  // 字符串类型，检查是否为日期格式
  if (typeof firstValue === "string") {
    const str = String(firstValue);
    if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}$/.test(str)) {
      return "date (YYYY-MM-DD)";
    }
    if (/^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/.test(str)) {
      return "date (MM/DD/YYYY)";
    }
    if (/^[\d,.]+$/.test(str) && !isNaN(parseFloat(str.replace(/,/g, "")))) {
      return "number (string)";
    }
    return "text";
  }

  if (typeof firstValue === "boolean") {
    return "boolean";
  }

  return "unknown";
}

// ========== v2.7.3 按需上下文获取工具 ==========

/**
 * 获取表格/工作表结构（列名、列类型、行数）
 * 支持两种模式：
 * 1. 表格名称 - 从 Excel Table 获取
 * 2. 工作表名称 - 从工作表的已用区域推断
 */
export function createGetTableSchemaTool(): Tool {
  return {
    name: "get_table_schema",
    description: "获取表格详细结构（列名、数据类型、格式、样本值）。用于理解数据结构后再操作。",
    category: "excel",
    parameters: [
      { name: "name", type: "string", description: "表格名称或工作表名称", required: true },
    ],
    execute: async (input) => {
      // v3.0.7: 参数验证增强 - 如果没有提供 name，返回友好错误
      if (!input.name || String(input.name).trim() === "" || input.name === "undefined") {
        return {
          success: false,
          output:
            "请指定要查看的表格名称或工作表名称。可以使用 excel_get_tables 获取所有表格列表，或指定工作表名如 'Sheet1'。",
          error: "Missing required parameter: name",
        };
      }

      const name = String(input.name).trim();

      return await excelRun(async (ctx) => {
        // 首先尝试作为表格获取
        try {
          const table = ctx.workbook.tables.getItem(name);
          table.load("name, rowCount");

          const headerRange = table.getHeaderRowRange();
          headerRange.load("values");

          const dataRange = table.getDataBodyRange();
          dataRange.load("address, rowCount, columnCount, values, numberFormat");

          await ctx.sync();

          const columns = headerRange.values[0] as string[];
          const dataValues = dataRange.values as unknown[][];
          const numberFormats = dataRange.numberFormat as string[][];

          // v3.0.2: 推断每列的数据类型和格式
          const columnSchemas = columns.map((colName, colIdx) => {
            const sampleValues = dataValues.slice(0, 5).map((row) => row[colIdx]);
            const formats = numberFormats.slice(0, 5).map((row) => row[colIdx]);
            const dataType = inferDataType(sampleValues);
            const formatExample = formats[0] || "General";

            return {
              name: colName,
              type: dataType,
              format: formatExample,
              samples: sampleValues.slice(0, 3).map((v) => String(v ?? "")),
            };
          });

          const schemaOutput = columnSchemas
            .map(
              (col, idx) =>
                `  ${String.fromCharCode(65 + idx)}列「${col.name}」: ${col.type}, 格式=${col.format}, 示例=[${col.samples.join(", ")}]`
            )
            .join("\n");

          return {
            success: true,
            output:
              `表格「${name}」详细结构:\n` +
              `- 类型: Excel Table\n` +
              `- 行数: ${dataRange.rowCount}\n` +
              `- 列数: ${columns.length}\n` +
              `- 数据区域: ${dataRange.address}\n` +
              `- 列定义:\n${schemaOutput}`,
            data: {
              name,
              type: "table",
              columns: columnSchemas,
              rowCount: dataRange.rowCount,
              columnCount: columns.length,
              dataAddress: dataRange.address,
            },
          };
        } catch {
          // 表格不存在，尝试作为工作表获取
        }

        // 尝试作为工作表获取
        try {
          const sheet = ctx.workbook.worksheets.getItem(name);
          const usedRange = sheet.getUsedRangeOrNullObject();
          usedRange.load("values, rowCount, columnCount, address, numberFormat");

          await ctx.sync();

          if (usedRange.isNullObject) {
            return {
              success: false,
              output: `工作表「${name}」没有数据`,
              error: "Empty sheet",
            };
          }

          // 假设第一行是标题
          const allValues = usedRange.values as unknown[][];
          const numberFormats = usedRange.numberFormat as string[][];
          const firstRow = allValues[0];
          const columns = firstRow.map((h) => String(h ?? ""));
          const dataRowCount = usedRange.rowCount - 1;
          const dataValues = allValues.slice(1); // 跳过标题行

          // v3.0.2: 推断每列的数据类型和格式
          const columnSchemas = columns.map((colName, colIdx) => {
            const sampleValues = dataValues.slice(0, 5).map((row) => row[colIdx]);
            const formats = numberFormats.slice(1, 6).map((row) => row[colIdx]);
            const dataType = inferDataType(sampleValues);
            const formatExample = formats[0] || "General";

            return {
              name: colName,
              type: dataType,
              format: formatExample,
              samples: sampleValues.slice(0, 3).map((v) => String(v ?? "")),
            };
          });

          const schemaOutput = columnSchemas
            .map(
              (col, idx) =>
                `  ${String.fromCharCode(65 + idx)}列「${col.name}」: ${col.type}, 格式=${col.format}, 示例=[${col.samples.join(", ")}]`
            )
            .join("\n");

          return {
            success: true,
            output:
              `工作表「${name}」详细结构:\n` +
              `- 类型: 普通数据区域（非 Table）\n` +
              `- 数据行数: ${dataRowCount}\n` +
              `- 列数: ${columns.length}\n` +
              `- 已用区域: ${usedRange.address}\n` +
              `- 列定义:\n${schemaOutput}`,
            data: {
              name,
              type: "sheet",
              columns: columnSchemas,
              rowCount: dataRowCount,
              columnCount: columns.length,
              address: usedRange.address,
            },
          };
        } catch (error) {
          return {
            success: false,
            output: `未找到表格或工作表「${name}」`,
            error: String(error),
          };
        }
      });
    },
  };
}

/**
 * 获取表格的样本数据（前 n 行）
 * 支持两种模式：
 * 1. 表格名称 - 从 Excel Table 获取
 * 2. 工作表名称 - 从工作表的已用区域获取
 */
export function createSampleRowsTool(): Tool {
  return {
    name: "sample_rows",
    description:
      "获取样本数据（前 n 行）。可传入表格名称或工作表名称。如果是工作表，会读取已用区域的数据。",
    category: "excel",
    parameters: [
      { name: "name", type: "string", description: "表格名称或工作表名称", required: true },
      { name: "n", type: "number", description: "要获取的行数（默认 5）", required: false },
    ],
    execute: async (input) => {
      const name = String(input.name);
      const n = Math.min(Number(input.n) || 5, 20); // 最多 20 行

      return await excelRun(async (ctx) => {
        // 首先尝试作为表格获取
        try {
          const table = ctx.workbook.tables.getItem(name);

          const headerRange = table.getHeaderRowRange();
          headerRange.load("values");

          const dataRange = table.getDataBodyRange();
          dataRange.load("values, rowCount");

          await ctx.sync();

          const headers = headerRange.values[0] as string[];
          const allData = dataRange.values as unknown[][];
          const sampleData = allData.slice(0, n);

          // 格式化为可读表格
          const tableOutput = [
            headers.join(" | "),
            headers.map(() => "---").join(" | "),
            ...sampleData.map((row) => row.map((cell) => String(cell ?? "")).join(" | ")),
          ].join("\n");

          return {
            success: true,
            output: `表格「${name}」样本数据（前 ${sampleData.length} 行，共 ${allData.length} 行）:\n\n${tableOutput}`,
            data: {
              name,
              type: "table",
              headers,
              sampleData,
              totalRows: allData.length,
              sampledRows: sampleData.length,
            },
          };
        } catch {
          // 表格不存在，尝试作为工作表获取
        }

        // 尝试作为工作表获取
        try {
          const sheet = ctx.workbook.worksheets.getItem(name);
          const usedRange = sheet.getUsedRangeOrNullObject();
          usedRange.load("values, rowCount, columnCount, address");

          await ctx.sync();

          if (usedRange.isNullObject) {
            return {
              success: false,
              output: `工作表「${name}」没有数据`,
              error: "Empty sheet",
            };
          }

          const allValues = usedRange.values as unknown[][];

          // 假设第一行是标题
          const headers = allValues[0].map((h) => String(h ?? ""));
          const dataRows = allValues.slice(1);
          const sampleData = dataRows.slice(0, n);

          // 格式化为可读表格
          const tableOutput = [
            headers.join(" | "),
            headers.map(() => "---").join(" | "),
            ...sampleData.map((row) => row.map((cell) => String(cell ?? "")).join(" | ")),
          ].join("\n");

          return {
            success: true,
            output: `工作表「${name}」样本数据（前 ${sampleData.length} 行，共 ${dataRows.length} 行数据）:\n区域: ${usedRange.address}\n\n${tableOutput}`,
            data: {
              name,
              type: "sheet",
              headers,
              sampleData,
              totalRows: dataRows.length,
              sampledRows: sampleData.length,
              address: usedRange.address,
            },
          };
        } catch (error) {
          return {
            success: false,
            output: `未找到表格或工作表「${name}」`,
            error: String(error),
          };
        }
      });
    },
  };
}

/**
 * 获取工作表详细信息
 */
export function createGetSheetInfoTool(): Tool {
  return {
    name: "get_sheet_info",
    description: "获取指定工作表的详细信息（已用区域、表格列表、图表列表）",
    category: "excel",
    parameters: [
      {
        name: "sheetName",
        type: "string",
        description: "工作表名称（留空则用当前表）",
        required: false,
      },
    ],
    execute: async (input) => {
      const sheetName = input.sheetName ? String(input.sheetName) : undefined;

      return await excelRun(async (ctx) => {
        try {
          const sheet = sheetName
            ? ctx.workbook.worksheets.getItem(sheetName)
            : ctx.workbook.worksheets.getActiveWorksheet();

          sheet.load("name");

          const usedRange = sheet.getUsedRangeOrNullObject();
          usedRange.load("address, rowCount, columnCount");

          const tables = sheet.tables;
          tables.load("items/name");

          const charts = sheet.charts;
          charts.load("items/name, items/chartType");

          await ctx.sync();

          const tableNames = tables.items.map((t: { name: string }) => t.name);
          const chartInfo = charts.items.map(
            (c: { name: string; chartType: string }) => `${c.name} (${c.chartType})`
          );

          const hasData = !usedRange.isNullObject;

          return {
            success: true,
            output:
              `工作表「${sheet.name}」信息:\n` +
              `- 已用区域: ${hasData ? usedRange.address : "空"}\n` +
              `- 数据规模: ${hasData ? `${usedRange.rowCount}行 × ${usedRange.columnCount}列` : "无数据"}\n` +
              `- 表格: ${tableNames.length > 0 ? tableNames.join(", ") : "无"}\n` +
              `- 图表: ${chartInfo.length > 0 ? chartInfo.join(", ") : "无"}`,
            data: {
              sheetName: sheet.name,
              usedRange: hasData ? usedRange.address : null,
              rowCount: hasData ? usedRange.rowCount : 0,
              columnCount: hasData ? usedRange.columnCount : 0,
              tables: tableNames,
              charts: chartInfo,
            },
          };
        } catch (error) {
          return {
            success: false,
            output: `获取工作表信息失败: ${error}`,
            error: String(error),
          };
        }
      });
    },
  };
}

/**
 * 创建所有读取类工具
 */
export function createReadTools(): Tool[] {
  return [
    createReadSelectionTool(),
    createReadRangeTool(),
    createGetWorkbookInfoTool(),
    createGetTableSchemaTool(),
    createSampleRowsTool(),
    createGetSheetInfoTool(),
  ];
}
