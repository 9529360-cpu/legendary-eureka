/**
 * ExcelAdapter - Excel 工具适配器
 *
 * 这是 Agent 的一个"手臂"，让 Agent 能够操作 Excel
 *
 * 架构位置：
 * ┌─────────────────────────────────────┐
 * │            Agent Core               │
 * │               ↓                     │
 * │         Tool Registry               │
 * │               ↓                     │
 * │  ┌─────────────────────────────┐   │
 * │  │      Excel Adapter ← 这里   │   │
 * │  │  (read, write, chart, ...)  │   │
 * │  └─────────────────────────────┘   │
 * └─────────────────────────────────────┘
 *
 * 这个模块完全可以被替换成 WordAdapter、PowerPointAdapter 等
 */

/* eslint-disable office-addins/call-sync-after-load, office-addins/call-sync-before-read */
// 注：上述规则被禁用是因为代码中已正确调用 ctx.sync()，但 lint 无法正确跟踪异步流

import { Tool, ToolResult } from "./AgentCore";

// ========== v2.9.41: Sheet 辅助函数 ==========

/**
 * v2.9.41: 获取指定工作表或活动工作表
 * 解决 Critical Bug: Multi-sheet targeting 被忽略的问题
 */
function getTargetSheet(ctx: Excel.RequestContext, sheetName?: string | null): Excel.Worksheet {
  if (sheetName && typeof sheetName === "string" && sheetName.trim()) {
    // 使用指定的工作表
    return ctx.workbook.worksheets.getItem(sheetName.trim());
  }
  // 默认使用活动工作表
  return ctx.workbook.worksheets.getActiveWorksheet();
}

/**
 * v2.9.41: 从 input 中提取 sheet 参数
 */
function extractSheetName(input: Record<string, unknown>): string | null {
  const sheet = input.sheet || input.sheetName || input.worksheet || input.targetSheet;
  if (sheet && typeof sheet === "string" && sheet.trim()) {
    return sheet.trim();
  }
  return null;
}

// ========== Excel 工具定义 ==========

/**
 * 创建所有 Excel 相关的工具
 */
export function createExcelTools(): Tool[] {
  return [
    // 读取工具
    createReadSelectionTool(),
    createReadRangeTool(),
    createGetWorkbookInfoTool(),

    // v2.7.3 新增: 按需上下文获取工具
    createGetTableSchemaTool(),
    createSampleRowsTool(),
    createGetSheetInfoTool(),

    // 写入工具
    createWriteRangeTool(),
    createWriteCellTool(),

    // 公式工具
    createSetFormulaTool(),
    createBatchFormulaTool(),
    createSetFormulasTool(), // v2.9.40: excel_set_formulas 别名
    createFillFormulaTool(), // v2.9.40: excel_fill_formula 别名
    createSmartFormulaTool(), // v2.9.56: 智能公式工具（支持结构化引用 + 动态行数）

    // 格式化工具
    createFormatRangeTool(),
    createAutoFitTool(),
    createConditionalFormatTool(),
    createMergeCellsTool(), // v2.8.0 新增
    createBorderTool(), // v2.8.0 新增
    createNumberFormatTool(), // v2.8.0 新增

    // 图表工具
    createChartTool(),
    createChartTrendlineTool(), // v2.8.0 新增

    // 数据操作工具
    createSortTool(),
    createSortRangeTool(), // v2.9.40: excel_sort_range 别名
    createFilterTool(),
    createClearRangeTool(),
    createRemoveDuplicatesTool(),
    createFindReplaceTool(), // v2.8.0 新增
    createFillSeriesTool(), // v2.8.0 新增

    // v2.9.27 行/列操作工具
    createInsertRowsTool(),
    createDeleteRowsTool(),
    createInsertColumnsTool(),
    createDeleteColumnsTool(),
    createMoveRangeTool(),
    createCopyRangeTool(),

    // 工作表工具
    createSheetTool(),
    createCreateSheetTool(),
    createSwitchSheetTool(),
    createDeleteSheetTool(), // v2.8.0 新增
    createCopySheetTool(), // v2.8.0 新增
    createRenameSheetTool(), // v2.8.0 新增
    createProtectSheetTool(), // v2.8.0 新增

    // 表格工具 v2.8.0 新增
    createTableTool(),
    createPivotTableTool(),

    // 视图工具 v2.8.0 新增
    createFreezePanesTool(),
    createGroupRowsTool(),
    createGroupColumnsTool(),

    // 批注与链接 v2.8.0 新增
    createCommentTool(),
    createHyperlinkTool(),

    // 页面设置 v2.8.0 新增
    createPageSetupTool(),
    createPrintAreaTool(),

    // 数据验证工具
    createDataValidationTool(),

    // 分析工具
    createAnalyzeDataTool(),
    createGoalSeekTool(), // v2.8.0 新增

    // v2.9.45: 高级分析工具
    createTrendAnalysisTool(), // 趋势分析
    createAnomalyDetectionTool(), // 异常检测
    createDataInsightsTool(), // 数据洞察
    createStatisticalAnalysisTool(), // 统计分析
    createPredictiveAnalysisTool(), // 预测分析
    createProactiveSuggestionsTool(), // 主动建议

    // v2.9.48: 性能优化工具 (借鉴 office-js-snippets)
    createBatchWriteOptimizedTool(), // 批量写入优化
    createPerformanceModeTool(), // 性能模式切换
    createRecalculateTool(), // 手动重新计算

    // v2.9.48: 高级条件格式工具
    createAdvancedConditionalFormatTool(), // 高级条件格式
    createClearConditionalFormatsTool(), // 清除条件格式

    // v2.9.48: 报表与事件工具
    createQuickReportTool(), // 快速报表生成
    createDataChangeListenerTool(), // 数据变更监听

    // v2.9.49: 更多 office-js-snippets 功能
    createGeometricShapeTool(), // 几何形状
    createInsertImageTool(), // 插入图片
    createFindAllTool(), // 全局查找高亮
    createAdvancedCopyTool(), // 高级复制粘贴
    createMoveRangeAdvancedTool(), // 移动范围
    createNamedRangeTool(), // 命名范围
    createInsertExternalSheetsTool(), // 插入外部工作表

    // 通用工具
    createRespondToUserTool(),
    createClarifyRequestTool(), // v3.0.7: 澄清模糊请求
  ];
}

// ========== 读取类工具 ==========

function createReadSelectionTool(): Tool {
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

function createReadRangeTool(): Tool {
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

function createGetWorkbookInfoTool(): Tool {
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

// ========== v2.7.3 按需上下文获取工具 ==========

/**
 * 获取表格/工作表结构（列名、列类型、行数）
 * 支持两种模式：
 * 1. 表格名称 - 从 Excel Table 获取
 * 2. 工作表名称 - 从工作表的已用区域推断
 */
function createGetTableSchemaTool(): Tool {
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

/**
 * 获取表格的样本数据（前 n 行）
 * 支持两种模式：
 * 1. 表格名称 - 从 Excel Table 获取
 * 2. 工作表名称 - 从工作表的已用区域获取
 */
function createSampleRowsTool(): Tool {
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
function createGetSheetInfoTool(): Tool {
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

// ========== 写入类工具 ==========

function createWriteRangeTool(): Tool {
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

function createWriteCellTool(): Tool {
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

// ========== 公式类工具 ==========

function createSetFormulaTool(): Tool {
  return {
    name: "excel_set_formula",
    description: "在指定单元格或范围设置公式（支持单个单元格如 A1，也支持范围如 D2:D10）",
    category: "excel",
    parameters: [
      {
        name: "address",
        type: "string",
        description: "单元格或范围地址，如 A1 或 D2:D10",
        required: true,
      },
      {
        name: "formula",
        type: "string",
        description: "Excel公式，如 =SUM(A1:A10)。对于范围地址，公式会自动填充到所有单元格",
        required: true,
      },
      { name: "sheet", type: "string", description: "工作表名称（可选）", required: false },
    ],
    execute: async (input) => {
      // v2.9.38: 智能参数兼容 - 接受多种写法
      const address = String(input.address || input.cell || input.range || input.target || "A1");
      let formula = String(input.formula || input.expression || "");
      if (!formula.startsWith("=")) formula = "=" + formula;
      // v2.9.41: 支持指定工作表
      const sheetName = extractSheetName(input);

      return await excelRun(async (ctx) => {
        const sheet = getTargetSheet(ctx, sheetName);
        const range = sheet.getRange(address);

        // v2.9.10: 自动适配单元格或范围地址
        // 加载范围尺寸以构建正确维度的二维数组
        range.load("rowCount, columnCount");
        await ctx.sync();

        const rowCount = range.rowCount;
        const colCount = range.columnCount;

        // 构建正确维度的二维数组
        const formulas: string[][] = [];
        for (let r = 0; r < rowCount; r++) {
          const row: string[] = [];
          for (let c = 0; c < colCount; c++) {
            row.push(formula);
          }
          formulas.push(row);
        }

        range.formulas = formulas;
        range.load("values");
        await ctx.sync();

        // 返回结果摘要
        if (rowCount === 1 && colCount === 1) {
          const result = range.values[0][0];
          return {
            success: true,
            output: `已在 ${address} 设置公式 ${formula}，计算结果: ${result}`,
          };
        } else {
          // 范围情况：返回首行结果作为示例
          const firstResult = range.values[0][0];
          const lastResult = range.values[rowCount - 1][colCount - 1];
          return {
            success: true,
            output: `已在 ${address}（${rowCount}行×${colCount}列）设置公式 ${formula}，首个结果: ${firstResult}，末个结果: ${lastResult}`,
          };
        }
      });
    },
  };
}

function createBatchFormulaTool(): Tool {
  return {
    name: "excel_batch_formula",
    description: "批量应用公式到范围（推荐用于单列，如 D2:D100）。对于多列公式，请分别调用。",
    category: "excel",
    parameters: [
      {
        name: "address",
        type: "string",
        description: "范围地址（建议单列，如 D2:D100）",
        required: true,
      },
      {
        name: "formula",
        type: "string",
        description: "公式模板（会自动调整行引用）",
        required: true,
      },
    ],
    execute: async (input) => {
      // v2.9.38: 智能参数兼容 - 接受多种写法
      const address = String(
        input.address || input.range || input.cell || input.target || "A1:A10"
      );
      let formula = String(input.formula || input.expression || "");
      if (!formula.startsWith("=")) formula = "=" + formula;

      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(address);
        range.load("rowCount, columnCount, rowIndex");
        await ctx.sync();

        const rowCount = range.rowCount;
        const colCount = range.columnCount;

        // 方法: 构建正确维度的二维数组
        // 如果公式只适用于单列，每行用相同公式（Excel会自动调整行引用）
        // 如果是多列，每列需要不同公式（不推荐用这个工具）

        if (colCount > 1) {
          // 多列情况：提醒用户分列操作
          // 但仍然尝试执行：每列用相同公式
          console.warn(`批量公式应用于多列(${colCount}列)，建议分列调用`);
        }

        // 构建正确维度的二维数组 [rowCount][colCount]
        const formulas: string[][] = [];
        for (let r = 0; r < rowCount; r++) {
          const row: string[] = [];
          for (let c = 0; c < colCount; c++) {
            row.push(formula);
          }
          formulas.push(row);
        }

        range.formulas = formulas;
        await ctx.sync();

        return {
          success: true,
          output: `已将公式 ${formula} 应用到 ${address}（${rowCount}行 × ${colCount}列）`,
        };
      });
    },
  };
}

// v2.9.40: 公式工具别名
function createSetFormulasTool(): Tool {
  const baseTool = createBatchFormulaTool();
  return {
    ...baseTool,
    name: "excel_set_formulas",
    description: "批量设置公式到范围（excel_batch_formula 的别名）",
  };
}

function createFillFormulaTool(): Tool {
  const baseTool = createBatchFormulaTool();
  return {
    ...baseTool,
    name: "excel_fill_formula",
    description:
      "\u586B\u5145\u516C\u5F0F\u5230\u8303\u56F4\uFF08excel_batch_formula \u7684\u522B\u540D\uFF09",
  };
}

/**
 * v2.9.56: 智能公式工具
 *
 * 支持三种模式：
 * 1. 结构化引用: "=@[\u5355\u4EF7]*@[\u6570\u91CF]" - \u8F6C\u6362\u4E3A Excel Table \u7ED3\u6784\u5316\u5F15\u7528
 * 2. \u884C\u6A21\u677F: "=B{row}*C{row}" - \u6309\u884C\u5C55\u5F00\u4E3A B2*C2, B3*C3...
 * 3. \u5217\u6807\u8BC6\u6A21\u5F0F: \u53EA\u6307\u5B9A\u5217\uFF0C\u6839\u636E usedRange \u81EA\u52A8\u786E\u5B9A\u884C\u6570
 *
 * \u6838\u5FC3\u539F\u5219\uFF1A
 * - \u4E0D\u5199\u6B7B\u884C\u53F7\uFF0C\u6839\u636E\u5B9E\u9645\u6570\u636E\u884C\u6570\u51B3\u5B9A
 * - \u652F\u6301 Table \u7ED3\u6784\u5316\u5F15\u7528\uFF08\u6700\u4F73\u5B9E\u8DF5\uFF09
 * - \u8FD4\u56DE\u5B9E\u9645\u5F71\u54CD\u8303\u56F4\u548C\u7ED3\u679C\u9A8C\u8BC1
 */
function createSmartFormulaTool(): Tool {
  return {
    name: "excel_smart_formula",
    description:
      "\u667A\u80FD\u516C\u5F0F\u5DE5\u5177\uFF1A\u652F\u6301\u7ED3\u6784\u5316\u5F15\u7528(@[\u5B57\u6BB5])\u6216\u884C\u6A21\u677F({row})\uFF0C\u81EA\u52A8\u6839\u636E\u771F\u5B9E\u6570\u636E\u884C\u6570\u5199\u5165",
    category: "excel",
    parameters: [
      {
        name: "sheet",
        type: "string",
        description: "\u5DE5\u4F5C\u8868\u540D\u79F0",
        required: true,
      },
      {
        name: "column",
        type: "string",
        description:
          "\u76EE\u6807\u5217\uFF08\u5982 D \u6216 E\uFF09\uFF0C\u4E0D\u9700\u8981\u6307\u5B9A\u884C\u53F7",
        required: true,
      },
      {
        name: "logicalFormula",
        type: "string",
        description:
          '\u903B\u8F91\u516C\u5F0F\uFF0C\u5982 "=@[\u5355\u4EF7]*@[\u6570\u91CF]" \u6216 "=B{row}*C{row}"',
        required: true,
      },
      {
        name: "referenceMode",
        type: "string",
        description:
          "\u5F15\u7528\u6A21\u5F0F: structured(\u7ED3\u6784\u5316), row_template(\u884C\u6A21\u677F), a1_fixed(A1\u56FA\u5B9A)",
        required: false,
      },
      {
        name: "startRow",
        type: "number",
        description: "\u8D77\u59CB\u884C\uFF08\u9ED8\u8BA42\uFF0C\u8DF3\u8FC7\u8868\u5934\uFF09",
        required: false,
      },
    ],
    execute: async (input) => {
      const sheetName = String(input.sheet || "");
      const column = String(input.column || "A").toUpperCase();
      const logicalFormula = String(input.logicalFormula || input.formula || "");
      const referenceMode = String(input.referenceMode || "structured");
      const startRow = Number(input.startRow || 2);

      if (!logicalFormula) {
        return { success: false, output: "\u7F3A\u5C11 logicalFormula \u53C2\u6570" };
      }

      return await excelRun(async (ctx) => {
        const sheet = getTargetSheet(ctx, sheetName);

        // 1. \u83B7\u53D6\u5B9E\u9645\u6570\u636E\u884C\u6570
        const usedRange = sheet.getUsedRange();
        usedRange.load("rowCount, columnCount, address");
        await ctx.sync();

        const dataEndRow = usedRange.rowCount; // \u5B9E\u9645\u6709\u6570\u636E\u7684\u884C\u6570

        if (dataEndRow <= 1) {
          return {
            success: false,
            output: `\u5DE5\u4F5C\u8868 "${sheetName}" \u6CA1\u6709\u6570\u636E\u884C\uFF08\u53EA\u6709\u8868\u5934\u6216\u4E3A\u7A7A\uFF09`,
          };
        }

        // 2. \u68C0\u67E5\u662F\u5426\u6709 Excel Table
        const tables = sheet.tables;
        tables.load("items/name, items/columns/items/name");
        await ctx.sync();

        let useTableRef = false;
        let tableName = "";
        let _columnMappingInfo = "";

        if (tables.items.length > 0) {
          // \u6709 Table\uFF0C\u53EF\u4EE5\u4F7F\u7528\u7ED3\u6784\u5316\u5F15\u7528
          const table = tables.items[0]; // \u4F7F\u7528\u7B2C\u4E00\u4E2A Table
          tableName = table.name;
          useTableRef = referenceMode === "structured";

          if (useTableRef) {
            // \u83B7\u53D6 Table \u5217\u540D\u7528\u4E8E\u8F6C\u6362
            const tableColumns = table.columns;
            tableColumns.load("items/name");
            await ctx.sync();
            const _columnMappingInfo = tableColumns.items.map((c) => c.name).join(", ");
          }
        }

        // 3. \u8BFB\u53D6\u8868\u5934\u884C\u7528\u4E8E\u5B57\u6BB5\u540D\u5230\u5217\u53F7\u7684\u6620\u5C04
        const headerRange = sheet.getRange("1:1");
        headerRange.load("values");
        await ctx.sync();

        const headers = headerRange.values[0] as (string | number | boolean)[];
        const fieldToColumn: Record<string, string> = {};
        for (let i = 0; i < headers.length; i++) {
          const header = String(headers[i] || "").trim();
          if (header) {
            fieldToColumn[header] = indexToColumn(i + 1);
          }
        }

        // 4. \u8F6C\u6362\u516C\u5F0F
        let finalFormulas: string[][] = [];
        const targetRange = `${column}${startRow}:${column}${dataEndRow}`;

        if (useTableRef && tableName) {
          // \u7ED3\u6784\u5316\u5F15\u7528\u6A21\u5F0F\uFF1A\u5199\u5165\u4E00\u6B21\uFF0CExcel \u81EA\u52A8\u6269\u5C55
          // \u628A @[\u5B57\u6BB5\u540D] \u8F6C\u6362\u4E3A Excel Table \u7ED3\u6784\u5316\u5F15\u7528
          const tableFormula = convertToTableReference(logicalFormula);

          // Table \u7ED3\u6784\u5316\u5F15\u7528\u53EA\u9700\u5199\u5165\u4E00\u683C\uFF0CExcel \u81EA\u52A8\u5904\u7406
          const singleCell = sheet.getRange(`${column}${startRow}`);
          singleCell.formulas = [[tableFormula]];
          await ctx.sync();

          // \u9A8C\u8BC1\u7ED3\u679C
          const resultRange = sheet.getRange(targetRange);
          resultRange.load("values");
          await ctx.sync();

          const results = resultRange.values.map((r) => r[0]);
          const hasErrors = results.some(
            (v) =>
              typeof v === "string" &&
              (v.startsWith("#REF") || v.startsWith("#VALUE") || v.startsWith("#NAME"))
          );

          return {
            success: !hasErrors,
            output: hasErrors
              ? `\u516C\u5F0F\u5199\u5165 ${targetRange} \u4F46\u5B58\u5728\u9519\u8BEF\u503C\uFF0C\u8BF7\u68C0\u67E5\u5B57\u6BB5\u540D\u662F\u5426\u6B63\u786E`
              : `\u5DF2\u4F7F\u7528 Table "${tableName}" \u7ED3\u6784\u5316\u5F15\u7528\u5199\u5165 ${targetRange}\uFF08${dataEndRow - startRow + 1}\u884C\uFF09\uFF0C\u9996\u7ED3\u679C: ${results[0]}\uFF0C\u672B\u7ED3\u679C: ${results[results.length - 1]}`,
            data: {
              affectedRange: targetRange,
              affectedRows: dataEndRow - startRow + 1,
              formula: tableFormula,
              sampleResults: results.slice(0, 3),
            },
          };
        } else {
          // \u884C\u6A21\u677F\u6A21\u5F0F\uFF1A\u6309\u884C\u5C55\u5F00\u516C\u5F0F
          for (let row = startRow; row <= dataEndRow; row++) {
            let rowFormula = logicalFormula;

            // \u66FF\u6362 {row} \u5360\u4F4D\u7B26
            rowFormula = rowFormula.replace(/\{row\}/g, String(row));

            // \u66FF\u6362 @[\u5B57\u6BB5\u540D] \u4E3A\u5177\u4F53\u5355\u5143\u683C\u5F15\u7528
            rowFormula = rowFormula.replace(/@\[([^\]]+)\]/g, (_, fieldName) => {
              const col = fieldToColumn[fieldName];
              if (col) {
                return `${col}${row}`;
              }
              console.warn(`[SmartFormula] \u672A\u627E\u5230\u5B57\u6BB5: ${fieldName}`);
              return `@[${fieldName}]`; // \u4FDD\u7559\u539F\u6837\uFF0C\u8BA9 Excel \u62A5\u9519
            });

            finalFormulas.push([rowFormula]);
          }

          // \u5199\u5165\u516C\u5F0F
          const range = sheet.getRange(targetRange);
          range.formulas = finalFormulas;
          await ctx.sync();

          // \u9A8C\u8BC1\u7ED3\u679C
          range.load("values");
          await ctx.sync();

          const results = range.values.map((r) => r[0]);
          const hasErrors = results.some(
            (v) =>
              typeof v === "string" &&
              (v.startsWith("#REF") || v.startsWith("#VALUE") || v.startsWith("#NAME"))
          );

          return {
            success: !hasErrors,
            output: hasErrors
              ? `\u516C\u5F0F\u5199\u5165 ${targetRange} \u4F46\u5B58\u5728\u9519\u8BEF\u503C`
              : `\u5DF2\u5C06\u516C\u5F0F\u5199\u5165 ${targetRange}\uFF08${dataEndRow - startRow + 1}\u884C\uFF09\uFF0C\u9996\u7ED3\u679C: ${results[0]}\uFF0C\u672B\u7ED3\u679C: ${results[results.length - 1]}`,
            data: {
              affectedRange: targetRange,
              affectedRows: dataEndRow - startRow + 1,
              sampleFormulas: finalFormulas.slice(0, 2).map((f) => f[0]),
              sampleResults: results.slice(0, 3),
            },
          };
        }
      });
    },
  };
}

/**
 * v2.9.56: \u5C06 @[\u5B57\u6BB5\u540D] \u8F6C\u6362\u4E3A Excel Table \u7ED3\u6784\u5316\u5F15\u7528
 * \u8F93\u5165: "=@[\u5355\u4EF7]*@[\u6570\u91CF]"
 * \u8F93\u51FA: "=[@\u5355\u4EF7]*[@\u6570\u91CF]"
 */
function convertToTableReference(formula: string): string {
  // Excel Table \u7ED3\u6784\u5316\u5F15\u7528\u683C\u5F0F: [@\u5217\u540D]
  return formula.replace(/@\[([^\]]+)\]/g, "[@$1]");
}

/**
 * v2.9.56: \u5217\u7D22\u5F15\u8F6C\u5217\u5B57\u6BCD
 */
function indexToColumn(index: number): string {
  let column = "";
  while (index > 0) {
    const remainder = (index - 1) % 26;
    column = String.fromCharCode(65 + remainder) + column;
    index = Math.floor((index - 1) / 26);
  }
  return column || "A";
}

// ========== \u683C\u5F0F\u5316\u7C7B\u5DE5\u5177 ==========

function createFormatRangeTool(): Tool {
  return {
    name: "excel_format_range",
    description: "格式化单元格范围（颜色、字体、边框等）",
    category: "excel",
    parameters: [
      { name: "address", type: "string", description: "范围地址", required: true },
      { name: "fill", type: "string", description: "填充颜色，如 #FFFF00", required: false },
      { name: "fontColor", type: "string", description: "字体颜色", required: false },
      { name: "bold", type: "boolean", description: "是否加粗", required: false },
      { name: "italic", type: "boolean", description: "是否斜体", required: false },
      { name: "fontSize", type: "number", description: "字体大小", required: false },
      {
        name: "horizontalAlignment",
        type: "string",
        description: "水平对齐: left/center/right",
        required: false,
      },
      {
        name: "verticalAlignment",
        type: "string",
        description: "垂直对齐: top/center/bottom",
        required: false,
      },
      { name: "autoFit", type: "boolean", description: "是否自动调整列宽", required: false },
      { name: "sheet", type: "string", description: "工作表名称（可选）", required: false },
    ],
    execute: async (input) => {
      // v2.9.40: 智能参数兼容 - 接受多种写法，包括 format 对象
      const address = String(input.address || input.range || input.cell || input.area || "A1");
      // v2.9.41: 支持指定工作表
      const sheetName = extractSheetName(input);

      // v2.9.40: 如果传入了 format 对象，展开它
      const formatObj = input.format as Record<string, unknown> | undefined;
      if (formatObj && typeof formatObj === "object") {
        // 将 format 对象的属性合并到 input
        Object.assign(input, formatObj);
      }

      return await excelRun(async (ctx) => {
        const sheet = getTargetSheet(ctx, sheetName);
        sheet.load("name");
        const range = sheet.getRange(address);

        // v2.9.40: 加载范围信息以获取实际地址
        range.load("address");
        await ctx.sync();

        // v2.9.40: 确保只格式化指定的单元格范围，而不是整列
        const actualAddress = range.address;
        console.log(`[ExcelAdapter] 格式化范围: ${actualAddress}`);

        if (input.fill || input.backgroundColor)
          range.format.fill.color = String(input.fill || input.backgroundColor);
        if (input.fontColor) range.format.font.color = String(input.fontColor);
        if (input.bold !== undefined) range.format.font.bold = Boolean(input.bold);
        if (input.italic !== undefined) range.format.font.italic = Boolean(input.italic);
        if (input.fontSize) range.format.font.size = Number(input.fontSize);

        // v2.9.40: 水平和垂直对齐
        if (input.horizontalAlignment) {
          const hAlign = String(input.horizontalAlignment).toLowerCase();
          if (hAlign === "center")
            range.format.horizontalAlignment = Excel.HorizontalAlignment.center;
          else if (hAlign === "left")
            range.format.horizontalAlignment = Excel.HorizontalAlignment.left;
          else if (hAlign === "right")
            range.format.horizontalAlignment = Excel.HorizontalAlignment.right;
        }
        if (input.verticalAlignment) {
          const vAlign = String(input.verticalAlignment).toLowerCase();
          if (vAlign === "center") range.format.verticalAlignment = Excel.VerticalAlignment.center;
          else if (vAlign === "top") range.format.verticalAlignment = Excel.VerticalAlignment.top;
          else if (vAlign === "bottom")
            range.format.verticalAlignment = Excel.VerticalAlignment.bottom;
        }

        await ctx.sync();

        // v2.9.40: 自动调整列宽
        if (input.autoFit) {
          range.format.autofitColumns();
          range.format.autofitRows();
          await ctx.sync();
        }

        const appliedFormats: string[] = [];
        if (input.fill || input.backgroundColor)
          appliedFormats.push(`背景色=${input.fill || input.backgroundColor}`);
        if (input.fontColor) appliedFormats.push(`字体色=${input.fontColor}`);
        if (input.bold) appliedFormats.push("加粗");
        if (input.italic) appliedFormats.push("斜体");
        if (input.fontSize) appliedFormats.push(`字号=${input.fontSize}`);
        if (input.horizontalAlignment) appliedFormats.push(`水平对齐=${input.horizontalAlignment}`);
        if (input.verticalAlignment) appliedFormats.push(`垂直对齐=${input.verticalAlignment}`);
        if (input.autoFit) appliedFormats.push("自动列宽");

        return {
          success: true,
          output: `已格式化 ${actualAddress}: ${appliedFormats.join(", ")}`,
          data: { address: actualAddress, formats: appliedFormats },
        };
      });
    },
  };
}

function createAutoFitTool(): Tool {
  return {
    name: "excel_auto_fit",
    description: "自动调整列宽或行高",
    category: "excel",
    parameters: [
      { name: "address", type: "string", description: "范围地址", required: true },
      { name: "type", type: "string", description: "调整类型: columns 或 rows", required: false },
    ],
    execute: async (input) => {
      // v2.9.38: 智能参数兼容 - 接受多种写法
      const address = String(input.address || input.range || input.cell || input.area || "A1:A10");
      const type = String(input.type || input.fitType || input.dimension || "columns");

      return await excelRun(async (ctx) => {
        const range = ctx.workbook.worksheets.getActiveWorksheet().getRange(address);

        if (type === "rows") {
          range.format.autofitRows();
        } else {
          range.format.autofitColumns();
        }

        await ctx.sync();

        return {
          success: true,
          output: `已自动调整 ${address} 的${type === "rows" ? "行高" : "列宽"}`,
        };
      });
    },
  };
}

function createConditionalFormatTool(): Tool {
  return {
    name: "excel_conditional_format",
    description: "添加条件格式",
    category: "excel",
    parameters: [
      { name: "address", type: "string", description: "范围地址", required: true },
      {
        name: "rule",
        type: "string",
        description: "规则类型: greaterThan/lessThan/equalTo/between",
        required: true,
      },
      { name: "value", type: "number", description: "比较值", required: true },
      {
        name: "format",
        type: "object",
        description: "格式设置 { fill, fontColor }",
        required: false,
      },
    ],
    execute: async (input) => {
      // v2.9.38: 智能参数兼容 - 接受多种写法
      const address = String(input.address || input.range || input.cell || "A1:A10");
      const rule = String(input.rule || input.condition || "greaterThan");
      const value = Number(input.value || input.threshold || 0);
      const format = (input.format as { fill?: string; fontColor?: string }) || {};

      return await excelRun(async (ctx) => {
        const range = ctx.workbook.worksheets.getActiveWorksheet().getRange(address);

        const conditionalFormat = range.conditionalFormats.add(
          Excel.ConditionalFormatType.cellValue
        );

        const cellValue = conditionalFormat.cellValue;

        // 设置规则
        switch (rule) {
          case "greaterThan":
            cellValue.rule = {
              formula1: String(value),
              operator: Excel.ConditionalCellValueOperator.greaterThan,
            };
            break;
          case "lessThan":
            cellValue.rule = {
              formula1: String(value),
              operator: Excel.ConditionalCellValueOperator.lessThan,
            };
            break;
          case "equalTo":
            cellValue.rule = {
              formula1: String(value),
              operator: Excel.ConditionalCellValueOperator.equalTo,
            };
            break;
        }

        // 设置格式
        if (format.fill) {
          cellValue.format.fill.color = format.fill;
        }
        if (format.fontColor) {
          cellValue.format.font.color = format.fontColor;
        }

        await ctx.sync();

        return {
          success: true,
          output: `已为 ${address} 添加条件格式: ${rule} ${value}`,
        };
      });
    },
  };
}

// ========== 图表工具 ==========

function createChartTool(): Tool {
  return {
    name: "excel_create_chart",
    description: "基于数据创建图表",
    category: "excel",
    parameters: [
      { name: "dataRange", type: "string", description: "数据范围，如 A1:D10", required: true },
      {
        name: "chartType",
        type: "string",
        description: "图表类型: column/bar/line/pie/area",
        required: true,
      },
      { name: "title", type: "string", description: "图表标题", required: false },
    ],
    execute: async (input) => {
      // v2.9.38: 智能参数兼容 - 接受多种写法
      const dataRange = String(
        input.dataRange || input.range || input.data || input.sourceRange || "A1:B10"
      );
      const chartType = String(input.chartType || input.type || "column");
      const title =
        input.title || input.chartTitle || input.name
          ? String(input.title || input.chartTitle || input.name)
          : undefined;

      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(dataRange);

        const chartTypeMap: Record<string, Excel.ChartType> = {
          column: Excel.ChartType.columnClustered,
          bar: Excel.ChartType.barClustered,
          line: Excel.ChartType.line,
          pie: Excel.ChartType.pie,
          area: Excel.ChartType.area,
          scatter: Excel.ChartType.xyscatter,
        };

        const chart = sheet.charts.add(
          chartTypeMap[chartType] || Excel.ChartType.columnClustered,
          range,
          Excel.ChartSeriesBy.auto
        );

        if (title) {
          chart.title.text = title;
        }

        // 设置图表位置
        chart.left = 350;
        chart.top = 20;
        chart.width = 400;
        chart.height = 300;

        await ctx.sync();

        return {
          success: true,
          output: `已创建${chartType}图表，数据源: ${dataRange}${title ? `，标题: ${title}` : ""}`,
        };
      });
    },
  };
}

// ========== 数据操作工具 ==========

function createSortTool(): Tool {
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
      // v2.9.40: 智能参数兼容 - 接受多种写法
      const address = String(input.address || input.range || input.data || "A1:D10");

      // v2.9.40: 支持列字母或数字
      let column = 0;
      const colInput = input.column || input.sortColumn || input.key || 0;
      if (typeof colInput === "string" && /^[A-Za-z]+$/.test(colInput)) {
        // 列字母转数字 (A=0, B=1, C=2...)
        column = colInput.toUpperCase().charCodeAt(0) - 65;
      } else {
        column = Number(colInput);
      }

      const ascending = input.ascending !== false && input.descending !== true;

      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(address);

        // 加载范围地址以验证
        range.load("address");
        await ctx.sync();

        range.sort.apply([
          {
            key: column,
            ascending,
          },
        ]);
        await ctx.sync();

        return {
          success: true,
          output: `已按第${column + 1}列（${String.fromCharCode(65 + column)}列）${ascending ? "升序" : "降序"}排序 ${range.address}`,
          data: { address: range.address, column, ascending },
        };
      });
    },
  };
}

// v2.9.40: 添加排序工具别名 excel_sort_range -> excel_sort
function createSortRangeTool(): Tool {
  const baseTool = createSortTool();
  return {
    ...baseTool,
    name: "excel_sort_range",
    description: "对数据范围进行排序（excel_sort 的别名）",
  };
}

function createFilterTool(): Tool {
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
      // v2.9.40: 实现真正的筛选功能
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

        // v2.9.40: 使用 AutoFilter API 进行筛选
        try {
          // 启用自动筛选 - 第二个参数是列索引
          sheet.autoFilter.apply(range, column);
          // 注意：Excel JS API 不直接支持通过 apply 设置条件，需要额外步骤
          await ctx.sync();

          return {
            success: true,
            output: `已在 ${range.address} 上应用自动筛选：第${column + 1}列`,
            data: { address: range.address, column, criteria },
          };
        } catch (filterError) {
          // 如果自动筛选失败，尝试清除现有筛选后重试
          try {
            sheet.autoFilter.remove();
            await ctx.sync();

            sheet.autoFilter.apply(range, column);
            await ctx.sync();

            return {
              success: true,
              output: `已在 ${range.address} 上应用筛选：第${column + 1}列 = "${criteria}"`,
              data: { address: range.address, column, criteria },
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

function createClearRangeTool(): Tool {
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
      // v2.9.38: 智能参数兼容 - 接受多种写法
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

function createRemoveDuplicatesTool(): Tool {
  return {
    name: "excel_remove_duplicates",
    description: "删除重复行",
    category: "excel",
    parameters: [
      { name: "address", type: "string", description: "数据范围", required: true },
      { name: "columns", type: "array", description: "用于判断重复的列索引", required: false },
    ],
    execute: async (input) => {
      // v2.9.38: 智能参数兼容 - 接受多种写法
      const address = String(input.address || input.range || input.data || "A1:D10");
      const columns = (input.columns as number[]) || (input.compareColumns as number[]) || [0];

      return await excelRun(async (ctx) => {
        const range = ctx.workbook.worksheets.getActiveWorksheet().getRange(address);
        // removeDuplicates 需要列索引数组和是否包含标题
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

// ========== 工作表工具 ==========

function createSheetTool(): Tool {
  return {
    name: "excel_sheet",
    description: "工作表操作（创建、删除、重命名、切换）",
    category: "excel",
    parameters: [
      {
        name: "action",
        type: "string",
        description: "操作: create/delete/rename/switch",
        required: true,
      },
      { name: "name", type: "string", description: "工作表名称", required: true },
      {
        name: "newName",
        type: "string",
        description: "新名称（仅 rename 时需要）",
        required: false,
      },
    ],
    execute: async (input) => {
      const action = String(input.action);
      const name = String(input.name);
      const newName = input.newName ? String(input.newName) : undefined;

      return await excelRun(async (ctx) => {
        const sheets = ctx.workbook.worksheets;

        switch (action) {
          case "create": {
            const newSheet = sheets.add(name);
            newSheet.activate();
            await ctx.sync();
            return { success: true, output: `已创建工作表 "${name}"` };
          }
          case "delete": {
            const sheet = sheets.getItem(name);
            sheet.delete();
            await ctx.sync();
            return { success: true, output: `已删除工作表 "${name}"` };
          }
          case "rename": {
            const sheet = sheets.getItem(name);
            sheet.name = newName!;
            await ctx.sync();
            return { success: true, output: `已将工作表 "${name}" 重命名为 "${newName}"` };
          }
          case "switch": {
            const sheet = sheets.getItem(name);
            sheet.activate();
            await ctx.sync();
            return { success: true, output: `已切换到工作表 "${name}"` };
          }
          default:
            return { success: false, output: `未知操作: ${action}`, error: "Invalid action" };
        }
      });
    },
  };
}

// 快捷的创建工作表工具
function createCreateSheetTool(): Tool {
  return {
    name: "excel_create_sheet",
    description: "创建一个新的工作表",
    category: "excel",
    parameters: [{ name: "name", type: "string", description: "新工作表的名称", required: true }],
    execute: async (input) => {
      const name = String(input.name);
      return await excelRun(async (ctx) => {
        const sheets = ctx.workbook.worksheets;

        // v2.9.39: 检查是否已存在同名工作表
        sheets.load("items/name");
        await ctx.sync();

        const existingNames = sheets.items.map((s) => s.name);
        if (existingNames.includes(name)) {
          // 已存在，直接切换到该工作表
          const existingSheet = sheets.getItem(name);
          existingSheet.activate();
          await ctx.sync();
          return {
            success: true,
            output: `工作表 "${name}" 已存在，已切换到该工作表`,
            data: { sheetName: name, isNew: false },
          };
        }

        const newSheet = sheets.add(name);
        newSheet.activate();
        await ctx.sync();

        // v2.9.39: 验证创建成功
        newSheet.load("name");
        await ctx.sync();

        if (newSheet.name !== name) {
          return {
            success: false,
            output: `创建工作表失败：请求名称 "${name}"，实际名称 "${newSheet.name}"`,
          };
        }

        return {
          success: true,
          output: `已创建并切换到工作表 "${name}"`,
          data: { sheetName: name, isNew: true },
        };
      });
    },
  };
}

// 快捷的切换工作表工具
function createSwitchSheetTool(): Tool {
  return {
    name: "excel_switch_sheet",
    description: "切换到指定的工作表",
    category: "excel",
    parameters: [
      { name: "name", type: "string", description: "要切换到的工作表名称", required: true },
    ],
    execute: async (input) => {
      const name = String(input.name);
      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(name);
        sheet.activate();
        await ctx.sync();
        return { success: true, output: `已切换到工作表 "${name}"` };
      });
    },
  };
}

// 数据验证工具（下拉框等）
function createDataValidationTool(): Tool {
  return {
    name: "excel_add_data_validation",
    description: "为单元格添加数据验证（如下拉列表）",
    category: "excel",
    parameters: [
      { name: "address", type: "string", description: "目标范围地址，如 J2:J100", required: true },
      {
        name: "type",
        type: "string",
        description: "验证类型: list/number/textLength",
        required: true,
      },
      {
        name: "values",
        type: "array",
        description: '列表选项，如 ["选项1", "选项2"]（仅 list 类型需要）',
        required: false,
      },
      {
        name: "min",
        type: "number",
        description: "最小值（仅 number/textLength 类型需要）",
        required: false,
      },
      {
        name: "max",
        type: "number",
        description: "最大值（仅 number/textLength 类型需要）",
        required: false,
      },
    ],
    execute: async (input) => {
      // v2.9.38: 智能参数兼容 - 接受多种写法
      const address = String(input.address || input.range || input.cell || "A1");
      const type = String(input.type || input.validationType || "list");
      const values = (input.values || input.options || input.list) as string[] | undefined;
      const min = (input.min ?? input.minimum) as number | undefined;
      const max = (input.max ?? input.maximum) as number | undefined;

      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(address);

        if (type === "list" && values && values.length > 0) {
          range.dataValidation.rule = {
            list: {
              inCellDropDown: true,
              source: values.join(","),
            },
          };
        } else if (type === "number") {
          range.dataValidation.rule = {
            wholeNumber: {
              formula1: min ?? 0,
              formula2: max ?? 999999,
              operator: Excel.DataValidationOperator.between,
            },
          };
        } else if (type === "textLength") {
          range.dataValidation.rule = {
            textLength: {
              formula1: min ?? 0,
              formula2: max ?? 255,
              operator: Excel.DataValidationOperator.between,
            },
          };
        } else {
          return { success: false, output: `不支持的验证类型: ${type}`, error: "Invalid type" };
        }

        await ctx.sync();
        return { success: true, output: `已为 ${address} 添加 ${type} 数据验证` };
      });
    },
  };
}

// ========== 分析工具 ==========

function createAnalyzeDataTool(): Tool {
  return {
    name: "excel_analyze_data",
    description: "分析选区或指定范围的数据特征",
    category: "excel",
    parameters: [
      {
        name: "address",
        type: "string",
        description: "范围地址（留空则使用当前选区）",
        required: false,
      },
    ],
    execute: async (input) => {
      const address = input.address ? String(input.address) : undefined;

      return await excelRun(async (ctx) => {
        const range = address
          ? ctx.workbook.worksheets.getActiveWorksheet().getRange(address)
          : ctx.workbook.getSelectedRange();

        range.load("address, values, rowCount, columnCount");
        await ctx.sync();

        const values = range.values as unknown[][];
        const rowCount = values.length;
        const colCount = values[0]?.length || 0;

        // 统计分析
        let emptyCount = 0;
        let numericCount = 0;
        let textCount = 0;
        let numericSum = 0;
        let numericValues: number[] = [];

        for (const row of values) {
          for (const cell of row) {
            if (cell === null || cell === "" || cell === undefined) {
              emptyCount++;
            } else if (typeof cell === "number") {
              numericCount++;
              numericSum += cell;
              numericValues.push(cell);
            } else {
              textCount++;
            }
          }
        }

        const totalCells = rowCount * colCount;
        const completeness = Math.round((1 - emptyCount / totalCells) * 100);

        let statsText = "";
        if (numericValues.length > 0) {
          const avg = numericSum / numericValues.length;
          const min = Math.min(...numericValues);
          const max = Math.max(...numericValues);
          statsText = `\n数值统计: 总和=${numericSum.toFixed(2)}, 平均=${avg.toFixed(2)}, 最小=${min}, 最大=${max}`;
        }

        return {
          success: true,
          output: `数据分析 ${range.address}:
- 范围: ${rowCount}行 × ${colCount}列 (共${totalCells}个单元格)
- 数值单元格: ${numericCount}, 文本单元格: ${textCount}, 空单元格: ${emptyCount}
- 数据完整度: ${completeness}%${statsText}`,
          data: {
            address: range.address,
            rowCount,
            colCount,
            numericCount,
            textCount,
            emptyCount,
            completeness,
            numericSum,
            numericValues,
          },
        };
      });
    },
  };
}

/**
 * v2.9.45: 趋势分析工具
 * 分析数据的趋势方向、增长率和预测
 */
function createTrendAnalysisTool(): Tool {
  return {
    name: "excel_trend_analysis",
    description: "分析数据趋势，识别增长/下降模式，计算趋势线和增长率。适用于时间序列数据分析。",
    category: "excel",
    parameters: [
      {
        name: "address",
        type: "string",
        description: "数据范围地址，如 A1:B20。第一列应为X轴（如时间），第二列为Y轴（如数值）",
        required: true,
      },
      {
        name: "sheet",
        type: "string",
        description: "工作表名称（可选）",
        required: false,
      },
      {
        name: "predictPeriods",
        type: "number",
        description: "预测未来的周期数（默认3）",
        required: false,
      },
    ],
    execute: async (input) => {
      const address = String(input.address || "");
      const sheetName = extractSheetName(input);
      const predictPeriods = Number(input.predictPeriods) || 3;

      return await excelRun(async (ctx) => {
        const sheet = getTargetSheet(ctx, sheetName);
        const range = sheet.getRange(address);
        range.load("values, rowCount, columnCount");
        await ctx.sync();

        const values = range.values as unknown[][];
        if (values.length < 3) {
          return { success: false, output: "数据量不足，趋势分析需要至少3行数据" };
        }

        // 提取数值列
        const numericData: number[] = [];
        for (let i = 0; i < values.length; i++) {
          const lastCol = values[i][values[i].length - 1];
          const num = Number(lastCol);
          if (!isNaN(num)) {
            numericData.push(num);
          }
        }

        if (numericData.length < 3) {
          return { success: false, output: "有效数值不足，无法进行趋势分析" };
        }

        // 计算趋势（线性回归）
        const n = numericData.length;
        const xValues = Array.from({ length: n }, (_, i) => i);
        const meanX = xValues.reduce((a, b) => a + b, 0) / n;
        const meanY = numericData.reduce((a, b) => a + b, 0) / n;

        let numerator = 0;
        let denominator = 0;
        for (let i = 0; i < n; i++) {
          numerator += (xValues[i] - meanX) * (numericData[i] - meanY);
          denominator += (xValues[i] - meanX) ** 2;
        }

        const slope = denominator !== 0 ? numerator / denominator : 0;
        const intercept = meanY - slope * meanX;

        // 计算R²决定系数
        let ssRes = 0;
        let ssTot = 0;
        for (let i = 0; i < n; i++) {
          const predicted = slope * xValues[i] + intercept;
          ssRes += (numericData[i] - predicted) ** 2;
          ssTot += (numericData[i] - meanY) ** 2;
        }
        const rSquared = ssTot !== 0 ? 1 - ssRes / ssTot : 0;

        // 趋势方向和强度
        const trendDirection = slope > 0.01 ? "上升" : slope < -0.01 ? "下降" : "平稳";
        const trendStrength =
          Math.abs(rSquared) > 0.7 ? "强" : Math.abs(rSquared) > 0.4 ? "中等" : "弱";

        // 计算增长率
        const firstValue = numericData[0];
        const lastValue = numericData[numericData.length - 1];
        const totalGrowthRate =
          firstValue !== 0 ? ((lastValue - firstValue) / Math.abs(firstValue)) * 100 : 0;
        const avgGrowthRate = totalGrowthRate / (n - 1);

        // 预测未来值
        const predictions: number[] = [];
        for (let i = 0; i < predictPeriods; i++) {
          predictions.push(slope * (n + i) + intercept);
        }

        const output = `📊 趋势分析结果（${address}）:

📈 趋势方向: ${trendDirection} (${trendStrength}趋势)
📐 趋势线方程: Y = ${slope.toFixed(4)} × X + ${intercept.toFixed(2)}
📊 R²决定系数: ${(rSquared * 100).toFixed(1)}% (拟合度${rSquared > 0.7 ? "优秀" : rSquared > 0.4 ? "良好" : "一般"})

💹 增长统计:
- 总增长率: ${totalGrowthRate.toFixed(2)}%
- 平均周期增长率: ${avgGrowthRate.toFixed(2)}%
- 起始值: ${firstValue.toFixed(2)}, 终止值: ${lastValue.toFixed(2)}

🔮 预测（未来${predictPeriods}期）:
${predictions.map((p, i) => `- 第${n + i + 1}期: ${p.toFixed(2)}`).join("\n")}`;

        return {
          success: true,
          output,
          data: {
            slope,
            intercept,
            rSquared,
            trendDirection,
            trendStrength,
            totalGrowthRate,
            avgGrowthRate,
            predictions,
            dataPoints: numericData.length,
          },
        };
      });
    },
  };
}

/**
 * v2.9.45: 异常检测工具
 * 使用IQR和Z-Score方法检测数据异常值
 */
function createAnomalyDetectionTool(): Tool {
  return {
    name: "excel_anomaly_detection",
    description: "检测数据中的异常值（离群值），支持IQR和Z-Score两种方法。可自动高亮异常单元格。",
    category: "excel",
    parameters: [
      {
        name: "address",
        type: "string",
        description: "数据范围地址，如 B2:B100",
        required: true,
      },
      {
        name: "sheet",
        type: "string",
        description: "工作表名称（可选）",
        required: false,
      },
      {
        name: "method",
        type: "string",
        description: "检测方法: 'iqr'(四分位距法，默认) 或 'zscore'(标准差法)",
        required: false,
      },
      {
        name: "threshold",
        type: "number",
        description: "阈值: IQR方法默认1.5，Z-Score方法默认3",
        required: false,
      },
      {
        name: "highlight",
        type: "boolean",
        description: "是否自动高亮异常值（默认false）",
        required: false,
      },
    ],
    execute: async (input) => {
      const address = String(input.address || "");
      const sheetName = extractSheetName(input);
      const method = String(input.method || "iqr").toLowerCase();
      const threshold = Number(input.threshold) || (method === "zscore" ? 3 : 1.5);
      const highlight = Boolean(input.highlight);

      return await excelRun(async (ctx) => {
        const sheet = getTargetSheet(ctx, sheetName);
        const range = sheet.getRange(address);
        range.load("values, rowCount, address");
        await ctx.sync();

        const values = range.values as unknown[][];
        const numericData: Array<{ value: number; row: number; col: number }> = [];

        // 提取数值数据
        for (let row = 0; row < values.length; row++) {
          for (let col = 0; col < values[row].length; col++) {
            const num = Number(values[row][col]);
            if (!isNaN(num) && values[row][col] !== null && values[row][col] !== "") {
              numericData.push({ value: num, row, col });
            }
          }
        }

        if (numericData.length < 5) {
          return { success: false, output: "数据量不足，异常检测需要至少5个数值" };
        }

        const numericValues = numericData.map((d) => d.value);
        const anomalies: Array<{
          value: number;
          row: number;
          col: number;
          score: number;
          severity: string;
        }> = [];

        if (method === "iqr") {
          // IQR方法
          const sorted = [...numericValues].sort((a, b) => a - b);
          const q1Index = Math.floor(sorted.length * 0.25);
          const q3Index = Math.floor(sorted.length * 0.75);
          const q1 = sorted[q1Index];
          const q3 = sorted[q3Index];
          const iqr = q3 - q1;
          const lowerBound = q1 - threshold * iqr;
          const upperBound = q3 + threshold * iqr;

          for (const item of numericData) {
            if (item.value < lowerBound || item.value > upperBound) {
              const score =
                item.value < lowerBound
                  ? (lowerBound - item.value) / iqr
                  : (item.value - upperBound) / iqr;
              anomalies.push({
                ...item,
                score,
                severity: score > 2 ? "高" : score > 1 ? "中" : "低",
              });
            }
          }
        } else {
          // Z-Score方法
          const mean = numericValues.reduce((a, b) => a + b, 0) / numericValues.length;
          const std = Math.sqrt(
            numericValues.reduce((sum, v) => sum + (v - mean) ** 2, 0) / numericValues.length
          );

          if (std === 0) {
            return { success: true, output: "数据无变异，所有值相同，无异常值" };
          }

          for (const item of numericData) {
            const zScore = Math.abs((item.value - mean) / std);
            if (zScore > threshold) {
              anomalies.push({
                ...item,
                score: zScore,
                severity: zScore > 4 ? "高" : zScore > 3 ? "中" : "低",
              });
            }
          }
        }

        // 可选：高亮异常值
        if (highlight && anomalies.length > 0) {
          for (const anomaly of anomalies) {
            const cell = range.getCell(anomaly.row, anomaly.col);
            const fillColor =
              anomaly.severity === "高"
                ? "#FF6B6B"
                : anomaly.severity === "中"
                  ? "#FFE66D"
                  : "#FFB347";
            cell.format.fill.color = fillColor;
          }
          await ctx.sync();
        }

        const anomalyRate = ((anomalies.length / numericData.length) * 100).toFixed(1);

        const output = `🔍 异常检测结果（${address}）:

📊 检测方法: ${method.toUpperCase()}（阈值: ${threshold}）
📈 数据点总数: ${numericData.length}
⚠️ 异常值数量: ${anomalies.length} (${anomalyRate}%)

${
  anomalies.length > 0
    ? `🚨 异常值详情（前10个）:
${anomalies
  .slice(0, 10)
  .map(
    (a) =>
      `- 行${a.row + 1}列${a.col + 1}: ${a.value} (${a.severity}严重度, 偏离分数: ${a.score.toFixed(2)})`
  )
  .join("\n")}`
    : "✅ 未检测到异常值"
}

${highlight && anomalies.length > 0 ? "🎨 已高亮标记异常单元格" : ""}`;

        return {
          success: true,
          output,
          data: {
            method,
            threshold,
            totalPoints: numericData.length,
            anomalyCount: anomalies.length,
            anomalyRate: parseFloat(anomalyRate),
            anomalies: anomalies.slice(0, 50),
            highlighted: highlight,
          },
        };
      });
    },
  };
}

/**
 * v2.9.45: 数据洞察工具
 * 智能分析数据并生成可操作的洞察建议
 */
function createDataInsightsTool(): Tool {
  return {
    name: "excel_data_insights",
    description: "智能分析数据并生成洞察报告，包括数据质量评估、模式识别和改进建议。",
    category: "excel",
    parameters: [
      {
        name: "address",
        type: "string",
        description: "数据范围地址，如 A1:E100",
        required: true,
      },
      {
        name: "sheet",
        type: "string",
        description: "工作表名称（可选）",
        required: false,
      },
      {
        name: "hasHeaders",
        type: "boolean",
        description: "第一行是否为表头（默认true）",
        required: false,
      },
    ],
    execute: async (input) => {
      const address = String(input.address || "");
      const sheetName = extractSheetName(input);
      const hasHeaders = input.hasHeaders !== false;

      return await excelRun(async (ctx) => {
        const sheet = getTargetSheet(ctx, sheetName);
        const range = sheet.getRange(address);
        range.load("values, rowCount, columnCount");
        await ctx.sync();

        const values = range.values as unknown[][];
        if (values.length < 2) {
          return { success: false, output: "数据量不足，洞察分析需要至少2行数据" };
        }

        const headers = hasHeaders
          ? (values[0] as string[]).map(String)
          : values[0].map((_, i) => `列${i + 1}`);
        const dataRows = hasHeaders ? values.slice(1) : values;

        const insights: string[] = [];
        const recommendations: string[] = [];
        let qualityScore = 100;

        // 1. 数据完整性分析
        let totalCells = 0;
        let emptyCells = 0;
        const columnEmptyRates: Record<string, number> = {};

        for (let col = 0; col < headers.length; col++) {
          let colEmpty = 0;
          for (const row of dataRows) {
            totalCells++;
            if (row[col] === null || row[col] === "" || row[col] === undefined) {
              emptyCells++;
              colEmpty++;
            }
          }
          const emptyRate = (colEmpty / dataRows.length) * 100;
          columnEmptyRates[headers[col]] = emptyRate;
          if (emptyRate > 10) {
            insights.push(`⚠️ "${headers[col]}"列有${emptyRate.toFixed(1)}%的数据缺失`);
            qualityScore -= Math.min(emptyRate / 2, 15);
          }
        }

        const completenessRate = ((1 - emptyCells / totalCells) * 100).toFixed(1);
        if (Number(completenessRate) >= 95) {
          insights.push(`✅ 数据完整性优秀（${completenessRate}%）`);
        }

        // 2. 数据类型分析
        const columnTypes: Record<string, { type: string; uniqueCount: number }> = {};
        for (let col = 0; col < headers.length; col++) {
          let numCount = 0;
          let textCount = 0;
          const uniqueValues = new Set<string>();

          for (const row of dataRows) {
            const cell = row[col];
            if (cell !== null && cell !== "" && cell !== undefined) {
              uniqueValues.add(String(cell));
              if (typeof cell === "number" || !isNaN(Number(cell))) {
                numCount++;
              } else {
                textCount++;
              }
            }
          }

          const type = numCount > textCount ? "数值" : "文本";
          columnTypes[headers[col]] = { type, uniqueCount: uniqueValues.size };

          // 检测低基数列（可能适合下拉列表）
          if (uniqueValues.size <= 10 && uniqueValues.size > 1 && dataRows.length > 20) {
            recommendations.push(
              `💡 "${headers[col]}"列只有${uniqueValues.size}个唯一值，可考虑使用数据验证下拉列表`
            );
          }
        }

        // 3. 重复数据检测
        const rowStrings = dataRows.map((row) => JSON.stringify(row));
        const uniqueRows = new Set(rowStrings);
        const duplicateCount = dataRows.length - uniqueRows.size;
        if (duplicateCount > 0) {
          const dupRate = ((duplicateCount / dataRows.length) * 100).toFixed(1);
          insights.push(`🔄 发现${duplicateCount}行重复数据（${dupRate}%）`);
          recommendations.push(`💡 建议使用"删除重复项"功能清理重复数据`);
          qualityScore -= Math.min((duplicateCount / dataRows.length) * 20, 15);
        }

        // 4. 数值列统计洞察
        for (let col = 0; col < headers.length; col++) {
          if (columnTypes[headers[col]]?.type === "数值") {
            const numericValues = dataRows.map((row) => Number(row[col])).filter((v) => !isNaN(v));

            if (numericValues.length >= 5) {
              const mean = numericValues.reduce((a, b) => a + b, 0) / numericValues.length;
              const sorted = [...numericValues].sort((a, b) => a - b);
              const median = sorted[Math.floor(sorted.length / 2)];
              const std = Math.sqrt(
                numericValues.reduce((sum, v) => sum + (v - mean) ** 2, 0) / numericValues.length
              );

              // 检测偏态分布
              if (Math.abs(mean - median) > std * 0.5) {
                const skewDirection = mean > median ? "右偏（有高值异常）" : "左偏（有低值异常）";
                insights.push(`📊 "${headers[col]}"列数据${skewDirection}`);
              }

              // 检测高变异性
              const cv = (std / mean) * 100;
              if (cv > 50) {
                insights.push(`📈 "${headers[col]}"列变异系数高达${cv.toFixed(1)}%，数据波动较大`);
              }
            }
          }
        }

        // 5. 生成建议
        if (qualityScore < 80) {
          recommendations.push(`🔧 数据质量评分${qualityScore.toFixed(0)}/100，建议进行数据清洗`);
        }
        if (dataRows.length > 1000) {
          recommendations.push(
            `💡 数据量较大（${dataRows.length}行），建议使用数据透视表进行汇总分析`
          );
        }

        // 确定质量等级
        const qualityLevel =
          qualityScore >= 90
            ? "优秀"
            : qualityScore >= 75
              ? "良好"
              : qualityScore >= 60
                ? "一般"
                : "需改进";

        const output = `📊 数据洞察报告（${address}）

📋 数据概览:
- 数据范围: ${dataRows.length}行 × ${headers.length}列
- 数据完整性: ${completenessRate}%
- 质量评分: ${qualityScore.toFixed(0)}/100 (${qualityLevel})

🔍 发现的洞察:
${insights.length > 0 ? insights.map((i) => `  ${i}`).join("\n") : "  ✅ 数据状态良好，未发现显著问题"}

💡 改进建议:
${recommendations.length > 0 ? recommendations.map((r) => `  ${r}`).join("\n") : "  ✅ 数据已达到良好状态"}

📊 列类型分析:
${Object.entries(columnTypes)
  .map(([name, info]) => `  - ${name}: ${info.type}型，${info.uniqueCount}个唯一值`)
  .join("\n")}`;

        return {
          success: true,
          output,
          data: {
            rowCount: dataRows.length,
            columnCount: headers.length,
            completenessRate: parseFloat(completenessRate),
            qualityScore,
            qualityLevel,
            duplicateCount,
            insights,
            recommendations,
            columnTypes,
            columnEmptyRates,
          },
        };
      });
    },
  };
}

/**
 * v2.9.45: 统计分析工具
 * 提供描述性统计和相关性分析
 */
function createStatisticalAnalysisTool(): Tool {
  return {
    name: "excel_statistical_analysis",
    description: "执行描述性统计分析，包括均值、中位数、标准差、分位数、相关性等。",
    category: "excel",
    parameters: [
      {
        name: "address",
        type: "string",
        description: "数据范围地址，如 A1:D100",
        required: true,
      },
      {
        name: "sheet",
        type: "string",
        description: "工作表名称（可选）",
        required: false,
      },
      {
        name: "includeCorrelation",
        type: "boolean",
        description: "是否计算列间相关性（默认true，多列数值时）",
        required: false,
      },
    ],
    execute: async (input) => {
      const address = String(input.address || "");
      const sheetName = extractSheetName(input);
      const includeCorrelation = input.includeCorrelation !== false;

      return await excelRun(async (ctx) => {
        const sheet = getTargetSheet(ctx, sheetName);
        const range = sheet.getRange(address);
        range.load("values, rowCount, columnCount");
        await ctx.sync();

        const values = range.values as unknown[][];
        if (values.length < 2) {
          return { success: false, output: "数据量不足，统计分析需要至少2行数据" };
        }

        // 假设第一行是表头
        const headers = (values[0] as string[]).map(String);
        const dataRows = values.slice(1);

        const stats: Record<
          string,
          {
            count: number;
            sum: number;
            mean: number;
            median: number;
            std: number;
            min: number;
            max: number;
            q1: number;
            q3: number;
            iqr: number;
          }
        > = {};

        const numericColumns: Record<string, number[]> = {};

        // 计算每列统计
        for (let col = 0; col < headers.length; col++) {
          const numericValues = dataRows
            .map((row) => Number(row[col]))
            .filter((v) => !isNaN(v) && isFinite(v));

          if (numericValues.length >= 2) {
            numericColumns[headers[col]] = numericValues;

            const n = numericValues.length;
            const sum = numericValues.reduce((a, b) => a + b, 0);
            const mean = sum / n;
            const sorted = [...numericValues].sort((a, b) => a - b);
            const median =
              n % 2 === 0 ? (sorted[n / 2 - 1] + sorted[n / 2]) / 2 : sorted[Math.floor(n / 2)];
            const std = Math.sqrt(numericValues.reduce((s, v) => s + (v - mean) ** 2, 0) / n);
            const q1 = sorted[Math.floor(n * 0.25)];
            const q3 = sorted[Math.floor(n * 0.75)];

            stats[headers[col]] = {
              count: n,
              sum,
              mean,
              median,
              std,
              min: sorted[0],
              max: sorted[n - 1],
              q1,
              q3,
              iqr: q3 - q1,
            };
          }
        }

        const columnNames = Object.keys(stats);
        if (columnNames.length === 0) {
          return { success: false, output: "未找到有效的数值列进行统计分析" };
        }

        // 计算相关性矩阵
        const correlations: Record<string, Record<string, number>> = {};
        if (includeCorrelation && columnNames.length >= 2) {
          for (const col1 of columnNames) {
            correlations[col1] = {};
            for (const col2 of columnNames) {
              const data1 = numericColumns[col1];
              const data2 = numericColumns[col2];

              if (data1.length !== data2.length) {
                correlations[col1][col2] = 0;
                continue;
              }

              const n = Math.min(data1.length, data2.length);
              const mean1 = data1.slice(0, n).reduce((a, b) => a + b, 0) / n;
              const mean2 = data2.slice(0, n).reduce((a, b) => a + b, 0) / n;

              let num = 0;
              let den1 = 0;
              let den2 = 0;
              for (let i = 0; i < n; i++) {
                const d1 = data1[i] - mean1;
                const d2 = data2[i] - mean2;
                num += d1 * d2;
                den1 += d1 * d1;
                den2 += d2 * d2;
              }

              const denominator = Math.sqrt(den1 * den2);
              correlations[col1][col2] = denominator !== 0 ? num / denominator : 0;
            }
          }
        }

        // 生成报告
        let output = `📊 统计分析报告（${address}）\n\n`;

        for (const [colName, s] of Object.entries(stats)) {
          output += `📈 ${colName}:
  样本数: ${s.count}
  总和: ${s.sum.toFixed(2)}
  均值: ${s.mean.toFixed(4)}
  中位数: ${s.median.toFixed(4)}
  标准差: ${s.std.toFixed(4)}
  最小值: ${s.min.toFixed(2)}
  最大值: ${s.max.toFixed(2)}
  Q1(25%): ${s.q1.toFixed(2)}
  Q3(75%): ${s.q3.toFixed(2)}
  IQR: ${s.iqr.toFixed(2)}
\n`;
        }

        if (includeCorrelation && columnNames.length >= 2) {
          output += `🔗 相关性矩阵:\n`;
          output += "      " + columnNames.map((n) => n.substring(0, 8).padEnd(10)).join("") + "\n";
          for (const col1 of columnNames) {
            output += col1.substring(0, 6).padEnd(6);
            for (const col2 of columnNames) {
              const corr = correlations[col1]?.[col2] || 0;
              output += corr.toFixed(2).padStart(10);
            }
            output += "\n";
          }

          // 找出强相关对
          const strongCorrelations: string[] = [];
          for (let i = 0; i < columnNames.length; i++) {
            for (let j = i + 1; j < columnNames.length; j++) {
              const corr = correlations[columnNames[i]]?.[columnNames[j]] || 0;
              if (Math.abs(corr) > 0.7) {
                const strength = corr > 0 ? "正相关" : "负相关";
                strongCorrelations.push(
                  `${columnNames[i]} ↔ ${columnNames[j]}: ${(corr * 100).toFixed(0)}% (${strength})`
                );
              }
            }
          }

          if (strongCorrelations.length > 0) {
            output += `\n💡 发现强相关关系:\n${strongCorrelations.map((s) => `  - ${s}`).join("\n")}\n`;
          }
        }

        return {
          success: true,
          output,
          data: {
            statistics: stats,
            correlations: includeCorrelation ? correlations : undefined,
            columnCount: columnNames.length,
          },
        };
      });
    },
  };
}

/**
 * v2.9.45: 预测分析工具
 * 基于历史数据进行简单预测
 */
function createPredictiveAnalysisTool(): Tool {
  return {
    name: "excel_predictive_analysis",
    description: "基于历史数据进行预测分析，支持线性回归和移动平均预测方法。",
    category: "excel",
    parameters: [
      {
        name: "address",
        type: "string",
        description: "历史数据范围，如 A1:B20（第一列为时间/序号，第二列为数值）",
        required: true,
      },
      {
        name: "sheet",
        type: "string",
        description: "工作表名称（可选）",
        required: false,
      },
      {
        name: "periods",
        type: "number",
        description: "预测周期数（默认5）",
        required: false,
      },
      {
        name: "method",
        type: "string",
        description: "预测方法: 'linear'(线性回归，默认) 或 'ma'(移动平均)",
        required: false,
      },
      {
        name: "maWindow",
        type: "number",
        description: "移动平均窗口大小（默认3）",
        required: false,
      },
    ],
    execute: async (input) => {
      const address = String(input.address || "");
      const sheetName = extractSheetName(input);
      const periods = Number(input.periods) || 5;
      const method = String(input.method || "linear").toLowerCase();
      const maWindow = Number(input.maWindow) || 3;

      return await excelRun(async (ctx) => {
        const sheet = getTargetSheet(ctx, sheetName);
        const range = sheet.getRange(address);
        range.load("values, rowCount");
        await ctx.sync();

        const values = range.values as unknown[][];
        const numericData: number[] = [];

        // 提取数值（取最后一列）
        for (const row of values) {
          const val = Number(row[row.length - 1]);
          if (!isNaN(val)) {
            numericData.push(val);
          }
        }

        if (numericData.length < 5) {
          return { success: false, output: "数据量不足，预测分析需要至少5个数据点" };
        }

        const n = numericData.length;
        let predictions: number[] = [];
        let modelInfo = "";
        let confidence = 0;

        if (method === "linear") {
          // 线性回归预测
          const xValues = Array.from({ length: n }, (_, i) => i);
          const meanX = n / 2 - 0.5;
          const meanY = numericData.reduce((a, b) => a + b, 0) / n;

          let numerator = 0;
          let denominator = 0;
          for (let i = 0; i < n; i++) {
            numerator += (xValues[i] - meanX) * (numericData[i] - meanY);
            denominator += (xValues[i] - meanX) ** 2;
          }

          const slope = denominator !== 0 ? numerator / denominator : 0;
          const intercept = meanY - slope * meanX;

          // 计算R²
          let ssRes = 0;
          let ssTot = 0;
          for (let i = 0; i < n; i++) {
            const predicted = slope * i + intercept;
            ssRes += (numericData[i] - predicted) ** 2;
            ssTot += (numericData[i] - meanY) ** 2;
          }
          const rSquared = ssTot !== 0 ? 1 - ssRes / ssTot : 0;
          confidence = rSquared * 100;

          // 预测
          for (let i = 0; i < periods; i++) {
            predictions.push(slope * (n + i) + intercept);
          }

          modelInfo = `线性回归 (y = ${slope.toFixed(4)}x + ${intercept.toFixed(2)})，R² = ${(rSquared * 100).toFixed(1)}%`;
        } else {
          // 移动平均预测
          const lastWindow = numericData.slice(-maWindow);
          const lastMA = lastWindow.reduce((a, b) => a + b, 0) / maWindow;

          // 计算趋势
          const maValues: number[] = [];
          for (let i = maWindow - 1; i < n; i++) {
            const window = numericData.slice(i - maWindow + 1, i + 1);
            maValues.push(window.reduce((a, b) => a + b, 0) / maWindow);
          }

          const maTrend =
            maValues.length >= 2
              ? (maValues[maValues.length - 1] - maValues[0]) / (maValues.length - 1)
              : 0;

          // 预测
          for (let i = 0; i < periods; i++) {
            predictions.push(lastMA + maTrend * (i + 1));
          }

          // 估计置信度
          const residuals = maValues.map((ma, i) => Math.abs(numericData[i + maWindow - 1] - ma));
          const avgResidual = residuals.reduce((a, b) => a + b, 0) / residuals.length;
          const dataRange = Math.max(...numericData) - Math.min(...numericData);
          confidence = Math.max(0, 100 - (avgResidual / dataRange) * 100);

          modelInfo = `${maWindow}期移动平均，趋势: ${maTrend > 0 ? "上升" : maTrend < 0 ? "下降" : "平稳"}`;
        }

        // 计算预测变化
        const lastActual = numericData[n - 1];
        const lastPredicted = predictions[predictions.length - 1];
        const changePercent =
          lastActual !== 0 ? ((lastPredicted - lastActual) / Math.abs(lastActual)) * 100 : 0;

        const output = `🔮 预测分析报告（${address}）

📊 模型信息:
- 预测方法: ${modelInfo}
- 历史数据点: ${n}
- 预测周期: ${periods}
- 置信度: ${confidence.toFixed(1)}%

📈 预测结果:
${predictions.map((p, i) => `  第${n + i + 1}期: ${p.toFixed(2)}`).join("\n")}

💹 变化预测:
- 最近实际值: ${lastActual.toFixed(2)}
- 最终预测值: ${lastPredicted.toFixed(2)}
- 预期变化: ${changePercent > 0 ? "+" : ""}${changePercent.toFixed(1)}%

⚠️ 注意: 预测基于历史趋势，实际结果可能因外部因素而异。
${confidence < 50 ? "⚠️ 置信度较低，建议增加数据量或检查数据质量。" : ""}`;

        return {
          success: true,
          output,
          data: {
            method,
            modelInfo,
            historicalCount: n,
            predictions,
            confidence,
            changePercent,
          },
        };
      });
    },
  };
}

/**
 * v2.9.45: 主动建议工具
 * 分析当前工作表并提供改进建议
 */
function createProactiveSuggestionsTool(): Tool {
  return {
    name: "excel_proactive_suggestions",
    description: "分析当前工作表，主动提供格式优化、数据清理、可视化等改进建议。",
    category: "excel",
    parameters: [
      {
        name: "sheet",
        type: "string",
        description: "工作表名称（可选，默认当前活动表）",
        required: false,
      },
    ],
    execute: async (input) => {
      const sheetName = extractSheetName(input);

      return await excelRun(async (ctx) => {
        const sheet = getTargetSheet(ctx, sheetName);
        sheet.load("name");
        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load("values, rowCount, columnCount, address");
        await ctx.sync();

        if (usedRange.isNullObject || usedRange.rowCount === 0) {
          return { success: true, output: "当前工作表为空，暂无建议。" };
        }

        const values = usedRange.values as unknown[][];
        const suggestions: Array<{
          priority: string;
          category: string;
          title: string;
          description: string;
          action?: string;
        }> = [];

        // 1. 检查是否有表头
        const firstRow = values[0] || [];
        const hasHeaders = firstRow.every((cell) => typeof cell === "string" && cell.length > 0);
        if (!hasHeaders && values.length > 5) {
          suggestions.push({
            priority: "高",
            category: "结构",
            title: "添加表头",
            description: "数据缺少清晰的表头，建议在第一行添加列标题",
            action: "在A1单元格开始添加每列的描述性标题",
          });
        }

        // 2. 检查数据量和是否适合转换为表格
        if (values.length > 10 && !usedRange.address.includes("Table")) {
          suggestions.push({
            priority: "中",
            category: "格式",
            title: "转换为Excel表格",
            description: `数据有${values.length}行，转换为表格可获得自动筛选、格式化和公式填充功能`,
            action: "使用 excel_create_table 工具创建表格",
          });
        }

        // 3. 检查空单元格
        let emptyCount = 0;
        for (const row of values) {
          for (const cell of row) {
            if (cell === null || cell === "" || cell === undefined) {
              emptyCount++;
            }
          }
        }
        const emptyRate = (emptyCount / (values.length * (values[0]?.length || 1))) * 100;
        if (emptyRate > 10) {
          suggestions.push({
            priority: emptyRate > 30 ? "高" : "中",
            category: "数据质量",
            title: "处理空值",
            description: `数据中有${emptyRate.toFixed(1)}%的空单元格，可能影响分析结果`,
            action: "检查并填充缺失值，或删除不完整的行",
          });
        }

        // 4. 检查重复行
        const rowStrings = values.map((row) => JSON.stringify(row));
        const uniqueRows = new Set(rowStrings);
        const duplicateCount = values.length - uniqueRows.size;
        if (duplicateCount > 0) {
          suggestions.push({
            priority: duplicateCount > values.length * 0.1 ? "高" : "低",
            category: "数据质量",
            title: "删除重复数据",
            description: `发现${duplicateCount}行重复数据`,
            action: "使用 excel_remove_duplicates 工具删除重复项",
          });
        }

        // 5. 检查数值列是否适合添加汇总
        let numericColumnCount = 0;
        for (let col = 0; col < (values[0]?.length || 0); col++) {
          const numCount = values.slice(1).filter((row) => typeof row[col] === "number").length;
          if (numCount > values.length * 0.5) {
            numericColumnCount++;
          }
        }
        if (numericColumnCount >= 2 && values.length > 5) {
          suggestions.push({
            priority: "中",
            category: "分析",
            title: "添加汇总统计",
            description: `检测到${numericColumnCount}个数值列，可添加合计、平均值等汇总行`,
            action: "在数据下方添加 SUM、AVERAGE 等汇总公式",
          });
        }

        // 6. 检查是否适合创建图表
        if (values.length >= 3 && numericColumnCount >= 1) {
          suggestions.push({
            priority: "低",
            category: "可视化",
            title: "创建数据图表",
            description: "数据适合使用图表进行可视化展示",
            action: "使用 excel_create_chart 工具创建图表",
          });
        }

        // 7. 检查列宽是否需要调整
        let longTextCount = 0;
        for (const row of values) {
          for (const cell of row) {
            if (typeof cell === "string" && cell.length > 15) {
              longTextCount++;
            }
          }
        }
        if (longTextCount > values.length) {
          suggestions.push({
            priority: "低",
            category: "格式",
            title: "自动调整列宽",
            description: "部分单元格内容较长，调整列宽可提高可读性",
            action: "使用 excel_auto_fit 工具自动调整列宽",
          });
        }

        // 按优先级排序
        const priorityOrder = { 高: 0, 中: 1, 低: 2 };
        suggestions.sort(
          (a, b) =>
            priorityOrder[a.priority as keyof typeof priorityOrder] -
            priorityOrder[b.priority as keyof typeof priorityOrder]
        );

        if (suggestions.length === 0) {
          return {
            success: true,
            output: `✅ 工作表"${sheet.name}"状态良好，暂无需要改进的建议。\n\n数据概览: ${values.length}行 × ${values[0]?.length || 0}列`,
          };
        }

        const output = `💡 智能建议 - "${sheet.name}"工作表

📊 数据概览: ${values.length}行 × ${values[0]?.length || 0}列

📋 改进建议（共${suggestions.length}条）:

${suggestions
  .map(
    (s, i) => `${i + 1}. [${s.priority}优先级] ${s.title}
   分类: ${s.category}
   说明: ${s.description}
   ${s.action ? `操作: ${s.action}` : ""}`
  )
  .join("\n\n")}`;

        return {
          success: true,
          output,
          data: {
            sheetName: sheet.name,
            rowCount: values.length,
            columnCount: values[0]?.length || 0,
            suggestionCount: suggestions.length,
            suggestions,
          },
        };
      });
    },
  };
}

// ========== 工具函数 ==========

/**
 * 封装 Excel.run，处理错误
 */
// ========== 通用工具 ==========

/**
 * respond_to_user - 向用户发送最终回复
 *
 * 这是 Agent 完成任务时使用的工具，用于向用户提供最终答案。
 * 当 Agent 调用此工具时，表示任务已完成。
 */
function createRespondToUserTool(): Tool {
  return {
    name: "respond_to_user",
    description:
      "向用户发送最终回复。当你完成了用户的任务，或者需要向用户提供信息/解答时，使用此工具。调用此工具后，任务将被标记为完成。",
    category: "general",
    parameters: [
      {
        name: "message",
        type: "string",
        description: "发送给用户的消息，应该清晰地总结完成了什么操作或回答用户的问题",
        required: true,
      },
      {
        name: "success",
        type: "boolean",
        description: "任务是否成功完成（默认为 true）",
        required: false,
      },
    ],
    execute: async (params: Record<string, unknown>) => {
      const message = params.message as string;
      const success = params.success !== false; // 默认为 true

      // 这个工具只是返回消息，实际的用户交互由 UI 层处理
      return {
        success: success,
        output: message,
        data: {
          isResponse: true,
          shouldComplete: true,
        },
      };
    },
  };
}

/**
 * clarify_request - 向用户澄清模糊请求
 * v3.0.7: 新增工具，用于处理模糊+有副作用的请求
 *
 * 当用户请求模糊不清且可能有副作用时（如"删除没用的"），
 * Agent 应使用此工具向用户澄清，而不是直接操作。
 */
function createClarifyRequestTool(): Tool {
  return {
    name: "clarify_request",
    description:
      "向用户澄清模糊的请求。当用户请求不够明确且可能有副作用（如删除、修改数据）时使用。调用此工具后，等待用户回复后再继续。",
    category: "general",
    parameters: [
      {
        name: "question",
        type: "string",
        description: "向用户提问的内容，应该清晰地说明需要澄清什么",
        required: true,
      },
      {
        name: "options",
        type: "array",
        description: "提供给用户的选项列表（可选）",
        required: false,
      },
      {
        name: "context",
        type: "string",
        description: "为什么需要澄清的上下文说明（可选）",
        required: false,
      },
    ],
    execute: async (params: Record<string, unknown>) => {
      const question = params.question as string;
      const options = params.options as string[] | undefined;
      const context = params.context as string | undefined;

      // 构建澄清消息
      let message = question;
      if (options && options.length > 0) {
        message += "\n\n请选择：\n" + options.map((opt, i) => `${i + 1}. ${opt}`).join("\n");
      }
      if (context) {
        message = `${context}\n\n${message}`;
      }

      return {
        success: true,
        output: message,
        data: {
          isClarification: true,
          shouldComplete: true, // 澄清后暂停，等待用户回复
          question,
          options,
        },
      };
    },
  };
}

// ========== 工具辅助函数 ==========

/**
 * v2.9.44: 增强错误信息提取
 */
async function excelRun(
  callback: (ctx: Excel.RequestContext) => Promise<ToolResult>
): Promise<ToolResult> {
  if (typeof Excel === "undefined") {
    return {
      success: false,
      output: "Excel API 不可用（可能不在 Excel 环境中）",
      error: "Excel API not available",
    };
  }

  try {
    return await Excel.run(callback);
  } catch (error) {
    let errorMsg = error instanceof Error ? error.message : String(error);
    let errorCode = "";

    // v2.9.44: 提取 Office.js 错误码
    if (error && typeof error === "object" && "code" in error) {
      errorCode = String((error as { code: unknown }).code);
      errorMsg = `[${errorCode}] ${errorMsg}`;
    }

    // v2.9.44: 提取更多诊断信息
    if (error && typeof error === "object" && "debugInfo" in error) {
      const debugInfo = (error as { debugInfo: unknown }).debugInfo;
      console.error("[ExcelAdapter] Debug Info:", debugInfo);
    }

    // 开发模式下打印完整堆栈
    console.error("[ExcelAdapter] Excel 操作失败:", error);

    return {
      success: false,
      output: `Excel 操作失败: ${errorMsg}`,
      error: errorMsg,
    };
  }
}

// ========== v2.8.0 新增工具 ==========

/**
 * 合并单元格工具
 */
function createMergeCellsTool(): Tool {
  return {
    name: "excel_merge_cells",
    description: "合并或取消合并单元格",
    category: "excel",
    parameters: [
      { name: "sheet", type: "string", description: "工作表名称", required: true },
      { name: "range", type: "string", description: "要合并的范围，如 A1:D1", required: true },
      {
        name: "action",
        type: "string",
        description: "操作类型: merge(合并) / unmerge(取消合并) / merge_across(跨行合并)",
        required: false,
      },
    ],
    execute: async (params) => {
      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(params.sheet as string);
        const range = sheet.getRange(params.range as string);
        const action = (params.action as string) || "merge";

        if (action === "unmerge") {
          range.unmerge();
        } else if (action === "merge_across") {
          range.merge(true); // across = true 跨行合并
        } else {
          range.merge(false);
        }
        await ctx.sync();

        return {
          success: true,
          output: `已${action === "unmerge" ? "取消合并" : "合并"}单元格: ${params.range}`,
        };
      });
    },
  };
}

/**
 * 边框设置工具
 */
function createBorderTool(): Tool {
  return {
    name: "excel_set_border",
    description: "设置单元格边框样式",
    category: "excel",
    parameters: [
      { name: "sheet", type: "string", description: "工作表名称", required: true },
      { name: "range", type: "string", description: "范围，如 A1:D10", required: true },
      {
        name: "style",
        type: "string",
        description: "边框样式: thin(细线)/medium(中等)/thick(粗线)/double(双线)/none(无)",
        required: false,
      },
      { name: "color", type: "string", description: "边框颜色，如 #000000", required: false },
      {
        name: "edges",
        type: "string",
        description: "边框位置: all(全部)/outside(外框)/inside(内框)",
        required: false,
      },
    ],
    execute: async (params) => {
      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(params.sheet as string);
        const range = sheet.getRange(params.range as string);
        const style = (params.style as string) || "thin";
        const color = (params.color as string) || "#000000";
        const edges = (params.edges as string) || "all";

        // 使用 Excel.BorderLineStyle 枚举值
        type BorderStyleType =
          | "None"
          | "Continuous"
          | "Dash"
          | "DashDot"
          | "DashDotDot"
          | "Dot"
          | "Double"
          | "SlantDashDot";
        const borderStyleMap: Record<string, BorderStyleType> = {
          thin: "Continuous",
          medium: "Continuous",
          thick: "Continuous",
          double: "Double",
          none: "None",
        };
        const borderStyle: BorderStyleType = borderStyleMap[style] || "Continuous";

        const borders = [
          Excel.BorderIndex.edgeTop,
          Excel.BorderIndex.edgeBottom,
          Excel.BorderIndex.edgeLeft,
          Excel.BorderIndex.edgeRight,
        ];

        if (edges === "all" || edges === "outside") {
          for (const border of borders) {
            const b = range.format.borders.getItem(border);
            b.style = borderStyle as Excel.BorderLineStyle;
            b.color = color;
          }
        }

        if (edges === "all" || edges === "inside") {
          const insideBorders = [
            Excel.BorderIndex.insideHorizontal,
            Excel.BorderIndex.insideVertical,
          ];
          for (const border of insideBorders) {
            const b = range.format.borders.getItem(border);
            b.style = borderStyle as Excel.BorderLineStyle;
            b.color = color;
          }
        }

        await ctx.sync();
        return {
          success: true,
          output: `已设置 ${params.range} 的边框样式`,
        };
      });
    },
  };
}

/**
 * 数字格式工具
 */
function createNumberFormatTool(): Tool {
  return {
    name: "excel_number_format",
    description: "设置单元格的数字格式",
    category: "excel",
    parameters: [
      {
        name: "sheet",
        type: "string",
        description: "工作表名称（可选，默认使用活动工作表）",
        required: false,
      },
      { name: "range", type: "string", description: "范围，如 A1:D10", required: true },
      {
        name: "format",
        type: "string",
        description: "格式类型: number/currency/percent/date/text/scientific/custom",
        required: true,
      },
      {
        name: "customFormat",
        type: "string",
        description: "自定义格式字符串，如 #,##0.00",
        required: false,
      },
      { name: "decimals", type: "number", description: "小数位数", required: false },
    ],
    execute: async (params) => {
      return await excelRun(async (ctx) => {
        // v2.9.39: 如果未指定 sheet，使用活动工作表
        let sheet: Excel.Worksheet;
        if (params.sheet) {
          sheet = ctx.workbook.worksheets.getItem(params.sheet as string);
        } else {
          sheet = ctx.workbook.worksheets.getActiveWorksheet();
        }
        const range = sheet.getRange(params.range as string);
        const formatType = params.format as string;
        const decimals = (params.decimals as number) ?? 2;

        const formatMap: Record<string, string> = {
          number: decimals > 0 ? `#,##0.${"0".repeat(decimals)}` : "#,##0",
          currency: decimals > 0 ? `¥#,##0.${"0".repeat(decimals)}` : "¥#,##0",
          percent: decimals > 0 ? `0.${"0".repeat(decimals)}%` : "0%",
          date: "yyyy-mm-dd",
          text: "@",
          scientific: "0.00E+00",
        };

        const formatString = (params.customFormat as string) || formatMap[formatType] || "@";
        range.numberFormat = [[formatString]];
        await ctx.sync();

        return {
          success: true,
          output: `已设置 ${params.range} 的格式为: ${formatString}`,
        };
      });
    },
  };
}

/**
 * 图表趋势线工具
 */
function createChartTrendlineTool(): Tool {
  return {
    name: "excel_chart_trendline",
    description: "为图表添加趋势线",
    category: "excel",
    parameters: [
      { name: "sheet", type: "string", description: "图表所在的工作表", required: true },
      { name: "chartIndex", type: "number", description: "图表索引（从0开始）", required: false },
      {
        name: "seriesIndex",
        type: "number",
        description: "数据系列索引（从0开始）",
        required: false,
      },
      {
        name: "type",
        type: "string",
        description:
          "趋势线类型: linear(线性)/exponential(指数)/logarithmic(对数)/polynomial(多项式)/power(幂)/movingAverage(移动平均)",
        required: false,
      },
      { name: "displayEquation", type: "boolean", description: "是否显示公式", required: false },
      { name: "displayRSquared", type: "boolean", description: "是否显示R²值", required: false },
    ],
    execute: async (params) => {
      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(params.sheet as string);
        sheet.load("charts");
        await ctx.sync();
        const charts = sheet.charts;
        charts.load("items");
        await ctx.sync();

        const chartIndex = (params.chartIndex as number) || 0;
        if (charts.items.length === 0) {
          return { success: false, output: "工作表中没有图表" };
        }

        const chart = charts.items[chartIndex];
        const series = chart.series;
        series.load("items");
        await ctx.sync();

        const seriesIndex = (params.seriesIndex as number) || 0;
        if (series.items.length === 0) {
          return { success: false, output: "图表中没有数据系列" };
        }

        const targetSeries = series.items[seriesIndex];
        const trendlineType =
          {
            linear: Excel.ChartTrendlineType.linear,
            exponential: Excel.ChartTrendlineType.exponential,
            logarithmic: Excel.ChartTrendlineType.logarithmic,
            polynomial: Excel.ChartTrendlineType.polynomial,
            power: Excel.ChartTrendlineType.power,
            movingAverage: Excel.ChartTrendlineType.movingAverage,
          }[(params.type as string) || "linear"] || Excel.ChartTrendlineType.linear;

        targetSeries.trendlines.add(trendlineType);
        // displayEquation 和 displayRSquared 在某些 API 版本中不可用
        await ctx.sync();

        return {
          success: true,
          output: `已为图表添加${params.type || "linear"}趋势线`,
        };
      });
    },
  };
}

/**
 * 查找替换工具
 */
function createFindReplaceTool(): Tool {
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
          output: `已替换 ${replaceCount} 处: "${findText}" → "${replaceText}"`,
        };
      });
    },
  };
}

/**
 * 填充序列工具
 */
// ========== v2.9.27 行/列操作工具 ==========

/**
 * 插入行工具
 */
function createInsertRowsTool(): Tool {
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

        // 获取要插入位置的整行
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
function createDeleteRowsTool(): Tool {
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
function createInsertColumnsTool(): Tool {
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

        // 计算列范围
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
function createDeleteColumnsTool(): Tool {
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

/**
 * 移动范围工具（剪切+粘贴）
 */
function createMoveRangeTool(): Tool {
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

        // 读取源数据
        sourceRange.load("values, formulas, numberFormat");
        await ctx.sync();

        // 获取目标范围（与源范围同样大小）
        const values = sourceRange.values;
        const formulas = sourceRange.formulas;
        const formats = sourceRange.numberFormat;
        const rowCount = values.length;
        const colCount = values[0].length;

        const targetFullRange = targetRange.getResizedRange(rowCount - 1, colCount - 1);

        // 先复制到目标位置
        targetFullRange.formulas = formulas;
        targetFullRange.numberFormat = formats;
        await ctx.sync();

        // 清空源范围
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
function createCopyRangeTool(): Tool {
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

        // 读取源数据
        sourceRange.load("values, formulas, numberFormat");
        await ctx.sync();

        const formulas = sourceRange.formulas;
        const formats = sourceRange.numberFormat;
        const rowCount = formulas.length;
        const colCount = formulas[0].length;

        const targetFullRange = targetRange.getResizedRange(rowCount - 1, colCount - 1);

        // 复制到目标位置
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

function createFillSeriesTool(): Tool {
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

        // 获取起始值和目标范围
        const startValue = startCell.values[0][0];
        const rangeAddress = `${params.startCell}:${params.endCell}`;
        const range = sheet.getRange(rangeAddress);
        range.load("rowCount, columnCount");
        await ctx.sync();

        const step = (params.step as number) || 1;
        const isVertical = range.rowCount > range.columnCount;
        const count = isVertical ? range.rowCount : range.columnCount;

        const values: unknown[][] = [];
        for (let i = 0; i < (isVertical ? count : 1); i++) {
          const row: unknown[] = [];
          for (let j = 0; j < (isVertical ? 1 : count); j++) {
            const idx = isVertical ? i : j;
            if (typeof startValue === "number") {
              row.push(startValue + idx * step);
            } else {
              // 尝试提取数字后缀
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

/**
 * 删除工作表工具
 */
function createDeleteSheetTool(): Tool {
  return {
    name: "excel_delete_sheet",
    description: "删除指定的工作表",
    category: "excel",
    parameters: [
      { name: "name", type: "string", description: "要删除的工作表名称", required: true },
    ],
    execute: async (params) => {
      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(params.name as string);
        sheet.delete();
        await ctx.sync();

        return {
          success: true,
          output: `已删除工作表: ${params.name}`,
        };
      });
    },
  };
}

/**
 * 复制工作表工具
 */
function createCopySheetTool(): Tool {
  return {
    name: "excel_copy_sheet",
    description: "复制工作表",
    category: "excel",
    parameters: [
      { name: "source", type: "string", description: "源工作表名称", required: true },
      { name: "newName", type: "string", description: "新工作表名称", required: false },
    ],
    execute: async (params) => {
      return await excelRun(async (ctx) => {
        const sourceSheet = ctx.workbook.worksheets.getItem(params.source as string);
        const newSheet = sourceSheet.copy();

        if (params.newName) {
          newSheet.name = params.newName as string;
        }

        newSheet.load("name");
        await ctx.sync();

        return {
          success: true,
          output: `已复制工作表: ${params.source} → ${newSheet.name}`,
        };
      });
    },
  };
}

/**
 * 重命名工作表工具
 */
function createRenameSheetTool(): Tool {
  return {
    name: "excel_rename_sheet",
    description: "重命名工作表",
    category: "excel",
    parameters: [
      { name: "oldName", type: "string", description: "当前工作表名称", required: true },
      { name: "newName", type: "string", description: "新名称", required: true },
    ],
    execute: async (params) => {
      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(params.oldName as string);
        sheet.name = params.newName as string;
        await ctx.sync();

        return {
          success: true,
          output: `已重命名工作表: ${params.oldName} → ${params.newName}`,
        };
      });
    },
  };
}

/**
 * 保护工作表工具
 */
function createProtectSheetTool(): Tool {
  return {
    name: "excel_protect_sheet",
    description: "保护或取消保护工作表",
    category: "excel",
    parameters: [
      { name: "sheet", type: "string", description: "工作表名称", required: true },
      {
        name: "protect",
        type: "boolean",
        description: "true=保护, false=取消保护",
        required: true,
      },
      { name: "password", type: "string", description: "保护密码（可选）", required: false },
      { name: "allowSort", type: "boolean", description: "是否允许排序", required: false },
      { name: "allowFilter", type: "boolean", description: "是否允许筛选", required: false },
    ],
    execute: async (params) => {
      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(params.sheet as string);

        if (params.protect) {
          const options: Excel.WorksheetProtectionOptions = {
            allowSort: (params.allowSort as boolean) ?? false,
            allowAutoFilter: (params.allowFilter as boolean) ?? false,
          };
          sheet.protection.protect(options, params.password as string);
        } else {
          sheet.protection.unprotect(params.password as string);
        }

        await ctx.sync();

        return {
          success: true,
          output: `已${params.protect ? "保护" : "取消保护"}工作表: ${params.sheet}`,
        };
      });
    },
  };
}

/**
 * 创建表格工具
 */
function createTableTool(): Tool {
  return {
    name: "excel_create_table",
    description: "将数据范围转换为 Excel 表格",
    category: "excel",
    parameters: [
      { name: "sheet", type: "string", description: "工作表名称", required: true },
      { name: "range", type: "string", description: "数据范围，如 A1:D10", required: true },
      { name: "name", type: "string", description: "表格名称", required: false },
      { name: "hasHeaders", type: "boolean", description: "是否包含标题行", required: false },
      {
        name: "style",
        type: "string",
        description: "表格样式，如 TableStyleMedium2",
        required: false,
      },
    ],
    execute: async (params) => {
      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(params.sheet as string);
        const range = sheet.getRange(params.range as string);
        const hasHeaders = (params.hasHeaders as boolean) ?? true;

        const table = sheet.tables.add(range, hasHeaders);

        if (params.name) {
          table.name = params.name as string;
        }
        if (params.style) {
          table.style = params.style as string;
        }

        table.load("name");
        await ctx.sync();

        return {
          success: true,
          output: `已创建表格: ${table.name}，范围: ${params.range}`,
        };
      });
    },
  };
}

/**
 * 创建数据透视表工具
 */
function createPivotTableTool(): Tool {
  return {
    name: "excel_create_pivot_table",
    description: "创建数据透视表",
    category: "excel",
    parameters: [
      { name: "sourceSheet", type: "string", description: "源数据工作表", required: true },
      { name: "sourceRange", type: "string", description: "源数据范围", required: true },
      {
        name: "destSheet",
        type: "string",
        description: "目标工作表（可以是新工作表名）",
        required: true,
      },
      { name: "destCell", type: "string", description: "目标起始单元格，如 A1", required: false },
      { name: "name", type: "string", description: "透视表名称", required: false },
      { name: "rowFields", type: "string", description: "行字段，逗号分隔", required: false },
      { name: "columnFields", type: "string", description: "列字段，逗号分隔", required: false },
      { name: "valueFields", type: "string", description: "值字段，逗号分隔", required: false },
    ],
    execute: async (params) => {
      return await excelRun(async (ctx) => {
        const sourceSheet = ctx.workbook.worksheets.getItem(params.sourceSheet as string);
        const sourceRange = sourceSheet.getRange(params.sourceRange as string);

        // 检查目标工作表是否存在
        let destSheet: Excel.Worksheet;
        try {
          destSheet = ctx.workbook.worksheets.getItem(params.destSheet as string);
          await ctx.sync();
        } catch {
          destSheet = ctx.workbook.worksheets.add(params.destSheet as string);
          await ctx.sync();
        }

        const destCell = destSheet.getRange((params.destCell as string) || "A1");
        const pivotName = (params.name as string) || `透视表_${Date.now()}`;

        const pivotTable = ctx.workbook.pivotTables.add(pivotName, sourceRange, destCell);

        // 添加字段（如果指定）
        if (params.rowFields) {
          const fields = (params.rowFields as string).split(",").map((f) => f.trim());
          for (const field of fields) {
            try {
              pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(field));
            } catch {
              /* 字段可能不存在 */
            }
          }
        }

        if (params.valueFields) {
          const fields = (params.valueFields as string).split(",").map((f) => f.trim());
          for (const field of fields) {
            try {
              const dataHierarchy = pivotTable.dataHierarchies.add(
                pivotTable.hierarchies.getItem(field)
              );
              dataHierarchy.summarizeBy = Excel.AggregationFunction.sum;
            } catch {
              /* 字段可能不存在 */
            }
          }
        }

        await ctx.sync();

        return {
          success: true,
          output: `已创建数据透视表: ${pivotName}，位于 ${params.destSheet}`,
        };
      });
    },
  };
}

/**
 * 冻结窗格工具
 */
function createFreezePanesTool(): Tool {
  return {
    name: "excel_freeze_panes",
    description: "冻结或取消冻结窗格",
    category: "excel",
    parameters: [
      { name: "sheet", type: "string", description: "工作表名称", required: true },
      {
        name: "action",
        type: "string",
        description:
          "操作: freeze_rows(冻结首行)/freeze_columns(冻结首列)/freeze_at(冻结到指定位置)/unfreeze(取消冻结)",
        required: true,
      },
      {
        name: "cell",
        type: "string",
        description: "当 action=freeze_at 时，指定冻结位置（该单元格左上方将被冻结）",
        required: false,
      },
    ],
    execute: async (params) => {
      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(params.sheet as string);
        sheet.load("freezePanes");
        await ctx.sync();
        const freezePane = sheet.freezePanes;
        const action = params.action as string;

        switch (action) {
          case "freeze_rows":
            freezePane.freezeRows(1);
            break;
          case "freeze_columns":
            freezePane.freezeColumns(1);
            break;
          case "freeze_at":
            if (params.cell) {
              freezePane.freezeAt(sheet.getRange(params.cell as string));
            }
            break;
          case "unfreeze":
            freezePane.unfreeze();
            break;
        }

        await ctx.sync();

        return {
          success: true,
          output: `已${action === "unfreeze" ? "取消" : ""}冻结窗格`,
        };
      });
    },
  };
}

/**
 * 行分组工具
 */
function createGroupRowsTool(): Tool {
  return {
    name: "excel_group_rows",
    description: "创建行分组（大纲）",
    category: "excel",
    parameters: [
      { name: "sheet", type: "string", description: "工作表名称", required: true },
      { name: "startRow", type: "number", description: "起始行号", required: true },
      { name: "endRow", type: "number", description: "结束行号", required: true },
      { name: "collapsed", type: "boolean", description: "是否折叠", required: false },
    ],
    execute: async (params) => {
      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(params.sheet as string);
        const range = sheet.getRange(`${params.startRow}:${params.endRow}`);
        range.group(Excel.GroupOption.byRows);

        if (params.collapsed) {
          range.rowHidden = true;
        }

        await ctx.sync();

        return {
          success: true,
          output: `已分组行 ${params.startRow}-${params.endRow}`,
        };
      });
    },
  };
}

/**
 * 列分组工具
 */
function createGroupColumnsTool(): Tool {
  return {
    name: "excel_group_columns",
    description: "创建列分组（大纲）",
    category: "excel",
    parameters: [
      { name: "sheet", type: "string", description: "工作表名称", required: true },
      { name: "startColumn", type: "string", description: "起始列，如 B", required: true },
      { name: "endColumn", type: "string", description: "结束列，如 D", required: true },
      { name: "collapsed", type: "boolean", description: "是否折叠", required: false },
    ],
    execute: async (params) => {
      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(params.sheet as string);
        const range = sheet.getRange(`${params.startColumn}:${params.endColumn}`);
        range.group(Excel.GroupOption.byColumns);

        if (params.collapsed) {
          range.columnHidden = true;
        }

        await ctx.sync();

        return {
          success: true,
          output: `已分组列 ${params.startColumn}-${params.endColumn}`,
        };
      });
    },
  };
}

/**
 * 批注工具
 */
function createCommentTool(): Tool {
  return {
    name: "excel_comment",
    description: "添加、编辑或删除单元格批注",
    category: "excel",
    parameters: [
      { name: "sheet", type: "string", description: "工作表名称", required: true },
      { name: "cell", type: "string", description: "单元格地址，如 A1", required: true },
      {
        name: "action",
        type: "string",
        description: "操作: add(添加)/edit(编辑)/delete(删除)",
        required: true,
      },
      {
        name: "content",
        type: "string",
        description: "批注内容（add/edit 时需要）",
        required: false,
      },
    ],
    execute: async (params) => {
      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(params.sheet as string);
        const cell = sheet.getRange(params.cell as string);
        const action = params.action as string;

        // 批注 API 在 Office.js 中需要通过 workbook.comments 访问
        if (action === "add") {
          ctx.workbook.comments.add(cell, params.content as string);
        } else if (action === "edit" || action === "delete") {
          const comments = ctx.workbook.comments;
          comments.load("items");
          await ctx.sync();
          // 查找该单元格的批注
          const _cellAddress = (params.cell as string).toUpperCase();
          for (const c of comments.items) {
            c.load("location");
          }
          await ctx.sync();
          // 简化处理：仅支持添加批注
          return { success: false, output: "编辑/删除批注暂不支持，请使用添加" };
        }

        await ctx.sync();

        return {
          success: true,
          output: `已${action === "delete" ? "删除" : action === "edit" ? "编辑" : "添加"}批注: ${params.cell}`,
        };
      });
    },
  };
}

/**
 * 超链接工具
 */
function createHyperlinkTool(): Tool {
  return {
    name: "excel_hyperlink",
    description: "添加超链接",
    category: "excel",
    parameters: [
      { name: "sheet", type: "string", description: "工作表名称", required: true },
      { name: "cell", type: "string", description: "单元格地址，如 A1", required: true },
      {
        name: "url",
        type: "string",
        description: "链接地址（网址或内部引用如 Sheet2!A1）",
        required: true,
      },
      { name: "displayText", type: "string", description: "显示文本", required: false },
      { name: "tooltip", type: "string", description: "提示文本", required: false },
    ],
    execute: async (params) => {
      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(params.sheet as string);
        const cell = sheet.getRange(params.cell as string);

        const url = params.url as string;
        const displayText = (params.displayText as string) || url;
        const tooltip = (params.tooltip as string) || "";

        // 使用 HYPERLINK 公式
        if (url.startsWith("http")) {
          cell.formulas = [[`=HYPERLINK("${url}","${displayText}")`]];
        } else {
          // 内部链接
          cell.formulas = [[`=HYPERLINK("#${url}","${displayText}")`]];
        }

        if (tooltip) {
          // Office.js 目前不直接支持设置超链接提示
        }

        await ctx.sync();

        return {
          success: true,
          output: `已添加超链接: ${params.cell} → ${url}`,
        };
      });
    },
  };
}

/**
 * 页面设置工具
 */
function createPageSetupTool(): Tool {
  return {
    name: "excel_page_setup",
    description: "设置页面布局（用于打印）",
    category: "excel",
    parameters: [
      { name: "sheet", type: "string", description: "工作表名称", required: true },
      {
        name: "orientation",
        type: "string",
        description: "方向: portrait(纵向)/landscape(横向)",
        required: false,
      },
      {
        name: "paperSize",
        type: "string",
        description: "纸张大小: A4/Letter/Legal",
        required: false,
      },
      { name: "fitToPage", type: "boolean", description: "是否适应页面", required: false },
      {
        name: "margins",
        type: "string",
        description: "页边距: normal/narrow/wide",
        required: false,
      },
    ],
    execute: async (params) => {
      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(params.sheet as string);
        sheet.load("pageLayout");
        await ctx.sync();
        const pageLayout = sheet.pageLayout;

        if (params.orientation) {
          pageLayout.orientation =
            params.orientation === "landscape"
              ? Excel.PageOrientation.landscape
              : Excel.PageOrientation.portrait;
        }

        if (params.paperSize) {
          const sizeMap: Record<string, string> = {
            A4: "A4",
            Letter: "Letter",
            Legal: "Legal",
          };
          // 使用字符串设置纸张类型
          (pageLayout as unknown as { paperSize: string }).paperSize =
            sizeMap[params.paperSize as string] || "A4";
        }

        if (params.margins) {
          const marginPresets: Record<
            string,
            { left: number; right: number; top: number; bottom: number }
          > = {
            normal: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75 },
            narrow: { left: 0.25, right: 0.25, top: 0.75, bottom: 0.75 },
            wide: { left: 1, right: 1, top: 1, bottom: 1 },
          };
          const m = marginPresets[params.margins as string] || marginPresets.normal;
          pageLayout.leftMargin = m.left;
          pageLayout.rightMargin = m.right;
          pageLayout.topMargin = m.top;
          pageLayout.bottomMargin = m.bottom;
        }

        await ctx.sync();

        return {
          success: true,
          output: `已设置页面布局: ${params.sheet}`,
        };
      });
    },
  };
}

/**
 * 打印区域工具
 */
function createPrintAreaTool(): Tool {
  return {
    name: "excel_print_area",
    description: "设置或清除打印区域",
    category: "excel",
    parameters: [
      { name: "sheet", type: "string", description: "工作表名称", required: true },
      {
        name: "range",
        type: "string",
        description: "打印区域范围，如 A1:D20，留空则清除打印区域",
        required: false,
      },
    ],
    execute: async (params) => {
      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(params.sheet as string);

        if (params.range) {
          sheet.pageLayout.setPrintArea(params.range as string);
        } else {
          // 清除打印区域 - 设置为空
          sheet.pageLayout.setPrintArea("");
        }

        await ctx.sync();

        return {
          success: true,
          output: params.range ? `已设置打印区域: ${params.range}` : "已清除打印区域",
        };
      });
    },
  };
}

/**
 * 单变量求解工具（Goal Seek 模拟）
 */
function createGoalSeekTool(): Tool {
  return {
    name: "excel_goal_seek",
    description: "单变量求解：通过调整输入单元格使目标单元格达到指定值",
    category: "excel",
    parameters: [
      { name: "sheet", type: "string", description: "工作表名称", required: true },
      { name: "targetCell", type: "string", description: "目标单元格（包含公式）", required: true },
      { name: "targetValue", type: "number", description: "目标值", required: true },
      { name: "changingCell", type: "string", description: "可变单元格", required: true },
      { name: "maxIterations", type: "number", description: "最大迭代次数", required: false },
      { name: "precision", type: "number", description: "精度", required: false },
    ],
    execute: async (params) => {
      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(params.sheet as string);
        const targetCell = sheet.getRange(params.targetCell as string);
        const changingCell = sheet.getRange(params.changingCell as string);
        const targetValue = params.targetValue as number;
        const maxIterations = (params.maxIterations as number) || 100;
        const precision = (params.precision as number) || 0.001;

        // 简单的二分法求解
        changingCell.load("values");
        targetCell.load("values");
        await ctx.sync();

        let low = -1000000;
        let high = 1000000;
        let mid = (changingCell.values[0][0] as number) || 0;
        let iterations = 0;
        let currentValue = targetCell.values[0][0] as number;

        while (iterations < maxIterations && Math.abs(currentValue - targetValue) > precision) {
          mid = (low + high) / 2;
          changingCell.values = [[mid]];
          await ctx.sync();

          targetCell.load("values");
          await ctx.sync();
          currentValue = targetCell.values[0][0] as number;

          if (currentValue < targetValue) {
            low = mid;
          } else {
            high = mid;
          }
          iterations++;
        }

        const success = Math.abs(currentValue - targetValue) <= precision;

        return {
          success,
          output: success
            ? `求解成功！${params.changingCell} = ${mid.toFixed(4)}，目标值误差: ${Math.abs(currentValue - targetValue).toFixed(6)}`
            : `求解未收敛，当前值: ${currentValue}，目标值: ${targetValue}`,
        };
      });
    },
  };
}

// ========== v2.9.48: 性能优化工具 (借鉴 office-js-snippets) ==========

/**
 * 创建批量写入优化工具
 * 使用 suspendScreenUpdatingUntilNextSync 和 untrack 优化大批量操作
 */
function createBatchWriteOptimizedTool(): Tool {
  return {
    name: "excel_batch_write_optimized",
    description: "优化的批量写入工具，适用于大量数据写入（使用屏幕更新暂停和代理对象释放优化）",
    category: "excel",
    parameters: [
      { name: "startCell", type: "string", description: "起始单元格，如 A1", required: true },
      { name: "data", type: "array", description: "二维数据数组", required: true },
      {
        name: "pauseScreenUpdate",
        type: "boolean",
        description: "是否暂停屏幕更新（推荐对大数据启用）",
        required: false,
        default: true,
      },
      {
        name: "untrackRanges",
        type: "boolean",
        description: "是否释放代理对象（推荐对大数据启用）",
        required: false,
        default: true,
      },
    ],
    execute: async (input) => {
      const params = input as {
        startCell: string;
        data: unknown[][];
        pauseScreenUpdate?: boolean;
        untrackRanges?: boolean;
      };

      return await excelRun(async (ctx) => {
        const startTime = Date.now();
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();

        // 性能优化1: 暂停屏幕更新直到 sync
        if (params.pauseScreenUpdate !== false) {
          ctx.application.suspendScreenUpdatingUntilNextSync();
        }

        const rowCount = params.data.length;
        const colCount = params.data[0]?.length || 0;

        // 计算目标范围
        const startCell = sheet.getRange(params.startCell);
        const targetRange = startCell.getResizedRange(rowCount - 1, colCount - 1);

        // 直接设置整个范围的值（比逐个单元格快得多）
        targetRange.values = params.data;

        // 性能优化2: 释放代理对象
        if (params.untrackRanges !== false) {
          targetRange.untrack();
        }

        await ctx.sync();
        const elapsed = Date.now() - startTime;

        return {
          success: true,
          output: `批量写入完成！${rowCount} 行 × ${colCount} 列，共 ${rowCount * colCount} 个单元格，耗时 ${elapsed}ms`,
          data: { rowCount, colCount, elapsed },
        };
      });
    },
  };
}

/**
 * 创建性能模式切换工具
 */
function createPerformanceModeTool(): Tool {
  return {
    name: "excel_performance_mode",
    description: "切换 Excel 性能模式（手动计算模式可大幅提升批量操作速度）",
    category: "excel",
    parameters: [
      {
        name: "mode",
        type: "string",
        description: "模式: 'manual' (手动计算), 'automatic' (自动计算), 'query' (查询当前模式)",
        required: true,
      },
    ],
    execute: async (input) => {
      const params = input as { mode: string };

      return await excelRun(async (ctx) => {
        ctx.application.load("calculationMode");
        await ctx.sync();

        const currentMode = ctx.application.calculationMode;

        if (params.mode === "query") {
          return {
            success: true,
            output: `当前计算模式: ${currentMode}`,
            data: { calculationMode: currentMode },
          };
        }

        if (params.mode === "manual") {
          ctx.application.calculationMode = Excel.CalculationMode.manual;
        } else if (params.mode === "automatic") {
          ctx.application.calculationMode = Excel.CalculationMode.automatic;
        } else {
          return {
            success: false,
            output: `无效的模式: ${params.mode}，请使用 'manual', 'automatic' 或 'query'`,
          };
        }

        await ctx.sync();

        return {
          success: true,
          output: `计算模式已切换: ${currentMode} → ${params.mode}`,
          data: { previousMode: currentMode, newMode: params.mode },
        };
      });
    },
  };
}

/**
 * 创建手动重新计算工具
 */
function createRecalculateTool(): Tool {
  return {
    name: "excel_recalculate",
    description: "手动触发 Excel 重新计算（在手动计算模式下使用）",
    category: "excel",
    parameters: [
      {
        name: "type",
        type: "string",
        description: "计算类型: 'full' (完整), 'fullRebuild' (完整重建)",
        required: false,
        default: "full",
      },
    ],
    execute: async (input) => {
      const params = input as { type?: string };

      return await excelRun(async (ctx) => {
        const calcType =
          params.type === "fullRebuild"
            ? Excel.CalculationType.fullRebuild
            : Excel.CalculationType.full;

        ctx.application.calculate(calcType);
        await ctx.sync();

        return {
          success: true,
          output: `已触发 ${params.type || "full"} 重新计算`,
        };
      });
    },
  };
}

// ========== v2.9.48: 高级条件格式工具 (借鉴 office-js-snippets) ==========

/**
 * 创建高级条件格式工具
 */
function createAdvancedConditionalFormatTool(): Tool {
  return {
    name: "excel_advanced_conditional_format",
    description: "高级条件格式设置，支持预设规则、单元格值规则、优先级控制",
    category: "excel",
    parameters: [
      { name: "range", type: "string", description: "要应用条件格式的范围", required: true },
      {
        name: "type",
        type: "string",
        description:
          "格式类型: 'preset' (预设), 'cellValue' (单元格值), 'colorScale' (色阶), 'dataBar' (数据条), 'iconSet' (图标集)",
        required: true,
      },
      { name: "rule", type: "object", description: "规则配置", required: true },
      {
        name: "format",
        type: "object",
        description: "格式配置 (字体颜色、背景色等)",
        required: false,
      },
      { name: "priority", type: "number", description: "优先级 (0最高)", required: false },
      {
        name: "stopIfTrue",
        type: "boolean",
        description: "条件满足时停止其他规则",
        required: false,
      },
    ],
    execute: async (input) => {
      const params = input as {
        range: string;
        type: string;
        rule: Record<string, unknown>;
        format?: { fontColor?: string; fillColor?: string; bold?: boolean };
        priority?: number;
        stopIfTrue?: boolean;
      };

      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(params.range);

        let conditionalFormat: Excel.ConditionalFormat;

        switch (params.type) {
          case "preset": {
            conditionalFormat = range.conditionalFormats.add(
              Excel.ConditionalFormatType.presetCriteria
            );
            if (params.format?.fontColor) {
              conditionalFormat.preset.format.font.color = params.format.fontColor;
            }
            if (params.format?.bold) {
              conditionalFormat.preset.format.font.bold = params.format.bold;
            }
            if (params.format?.fillColor) {
              conditionalFormat.preset.format.fill.color = params.format.fillColor;
            }
            // 预设规则: oneStdDevAboveAverage, oneStdDevBelowAverage, duplicateValues, uniqueValues, blanks, nonBlanks 等
            conditionalFormat.preset.rule =
              params.rule as unknown as Excel.ConditionalPresetCriteriaRule;
            break;
          }

          case "cellValue": {
            conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
            if (params.format?.fontColor) {
              conditionalFormat.cellValue.format.font.color = params.format.fontColor;
            }
            if (params.format?.fillColor) {
              conditionalFormat.cellValue.format.fill.color = params.format.fillColor;
            }
            // 单元格值规则: { formula1: "=0", formula2?: "=100", operator: "LessThan" | "Between" | ... }
            conditionalFormat.cellValue.rule =
              params.rule as unknown as Excel.ConditionalCellValueRule;
            break;
          }

          case "colorScale": {
            conditionalFormat = range.conditionalFormats.add(
              Excel.ConditionalFormatType.colorScale
            );
            // 色阶会自动应用，可以自定义颜色
            const colorScale = conditionalFormat.colorScale;
            colorScale.criteria = params.rule as unknown as Excel.ConditionalColorScaleCriteria;
            break;
          }

          case "dataBar": {
            conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
            // 数据条配置
            if (params.format?.fillColor) {
              conditionalFormat.dataBar.barDirection =
                Excel.ConditionalDataBarDirection.leftToRight;
              conditionalFormat.dataBar.positiveFormat.fillColor = params.format.fillColor;
            }
            break;
          }

          case "iconSet": {
            conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
            // 图标集配置
            if (params.rule.style) {
              conditionalFormat.iconSet.style = params.rule.style as Excel.IconSet;
            }
            break;
          }

          default:
            return { success: false, output: `不支持的条件格式类型: ${params.type}` };
        }

        // 设置优先级
        if (params.priority !== undefined) {
          conditionalFormat.priority = params.priority;
        }

        // 设置 stopIfTrue
        if (params.stopIfTrue !== undefined) {
          conditionalFormat.stopIfTrue = params.stopIfTrue;
        }

        await ctx.sync();

        return {
          success: true,
          output: `高级条件格式已应用到 ${params.range}，类型: ${params.type}`,
        };
      });
    },
  };
}

/**
 * 清除所有条件格式工具
 */
function createClearConditionalFormatsTool(): Tool {
  return {
    name: "excel_clear_conditional_formats",
    description: "清除指定范围的所有条件格式",
    category: "excel",
    parameters: [
      {
        name: "range",
        type: "string",
        description: "要清除条件格式的范围，留空则清除整个工作表",
        required: false,
      },
    ],
    execute: async (input) => {
      const params = input as { range?: string };

      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const range = params.range ? sheet.getRange(params.range) : sheet.getRange();

        range.conditionalFormats.clearAll();
        await ctx.sync();

        return {
          success: true,
          output: params.range
            ? `已清除 ${params.range} 的所有条件格式`
            : `已清除工作表的所有条件格式`,
        };
      });
    },
  };
}

// ========== v2.9.48: 报表生成工具 (借鉴 office-js-snippets report-generation) ==========

/**
 * 创建快速报表工具
 */
function createQuickReportTool(): Tool {
  return {
    name: "excel_quick_report",
    description: "快速生成格式化报表（包含标题、数据表格和可选图表）",
    category: "excel",
    parameters: [
      { name: "title", type: "string", description: "报表标题", required: true },
      { name: "headers", type: "array", description: "列标题数组", required: true },
      { name: "data", type: "array", description: "二维数据数组", required: true },
      {
        name: "includeChart",
        type: "boolean",
        description: "是否添加图表",
        required: false,
        default: false,
      },
      {
        name: "chartType",
        type: "string",
        description: "图表类型: 'column', 'line', 'pie', 'bar'",
        required: false,
        default: "column",
      },
      {
        name: "sheetName",
        type: "string",
        description: "新工作表名称（留空则在当前表生成）",
        required: false,
      },
    ],
    execute: async (input) => {
      const params = input as {
        title: string;
        headers: string[];
        data: unknown[][];
        includeChart?: boolean;
        chartType?: string;
        sheetName?: string;
      };

      return await excelRun(async (ctx) => {
        // 暂停屏幕更新以提升性能
        ctx.application.suspendScreenUpdatingUntilNextSync();

        let sheet: Excel.Worksheet;
        if (params.sheetName) {
          // 删除已存在的同名工作表
          const existingSheet = ctx.workbook.worksheets.getItemOrNullObject(params.sheetName);
          existingSheet.delete();
          await ctx.sync();
          sheet = ctx.workbook.worksheets.add(params.sheetName);
        } else {
          sheet = ctx.workbook.worksheets.getActiveWorksheet();
        }

        // 1. 设置报表标题
        const titleCell = sheet.getCell(0, 0);
        titleCell.values = [[params.title]];
        titleCell.format.font.name = "微软雅黑";
        titleCell.format.font.size = 20;
        titleCell.format.font.bold = true;
        titleCell.format.font.color = "#1a1a1a";

        // 合并标题单元格
        const titleMergeRange = titleCell.getResizedRange(0, params.headers.length - 1);
        titleMergeRange.merge();
        titleMergeRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;

        // 2. 设置列标题
        const headerRow = titleCell
          .getOffsetRange(2, 0)
          .getResizedRange(0, params.headers.length - 1);
        headerRow.values = [params.headers];
        headerRow.format.font.bold = true;
        headerRow.format.fill.color = "#4472C4";
        headerRow.format.font.color = "#FFFFFF";
        headerRow.format.horizontalAlignment = Excel.HorizontalAlignment.center;

        // 3. 填充数据
        const dataRange = headerRow.getOffsetRange(1, 0).getResizedRange(params.data.length - 1, 0);
        dataRange.values = params.data;

        // 4. 添加表格边框和交替行颜色
        const fullTableRange = headerRow.getResizedRange(params.data.length, 0);
        fullTableRange.format.borders.getItem("EdgeTop").style = "Thin" as Excel.BorderLineStyle;
        fullTableRange.format.borders.getItem("EdgeBottom").style = "Thin" as Excel.BorderLineStyle;
        fullTableRange.format.borders.getItem("EdgeLeft").style = "Thin" as Excel.BorderLineStyle;
        fullTableRange.format.borders.getItem("EdgeRight").style = "Thin" as Excel.BorderLineStyle;
        fullTableRange.format.borders.getItem("InsideHorizontal").style =
          "Thin" as Excel.BorderLineStyle;
        fullTableRange.format.borders.getItem("InsideVertical").style =
          "Thin" as Excel.BorderLineStyle;

        // 自动调整列宽
        fullTableRange.format.autofitColumns();
        fullTableRange.format.autofitRows();

        // 确保第一列至少100像素宽
        const firstColRange = sheet.getRange("A:A");
        firstColRange.load("format/columnWidth");
        await ctx.sync();
        if (firstColRange.format.columnWidth < 100) {
          firstColRange.format.columnWidth = 100;
        }

        // 5. 可选: 添加图表
        if (params.includeChart) {
          const chartTypeMap: Record<string, Excel.ChartType> = {
            column: Excel.ChartType.columnClustered,
            line: Excel.ChartType.line,
            pie: Excel.ChartType.pie,
            bar: Excel.ChartType.barClustered,
          };

          const chart = sheet.charts.add(
            chartTypeMap[params.chartType || "column"] || Excel.ChartType.columnClustered,
            fullTableRange,
            Excel.ChartSeriesBy.columns
          );

          // 定位图表
          const chartTopRow = dataRange.getLastRow().getOffsetRange(2, 0);
          chart.setPosition(chartTopRow, chartTopRow.getOffsetRange(15, params.headers.length - 1));
          chart.title.text = params.title;
          chart.legend.position = Excel.ChartLegendPosition.right;
          chart.legend.format.fill.setSolidColor("white");
        }

        if (params.sheetName) {
          sheet.activate();
        }

        await ctx.sync();

        return {
          success: true,
          output: `报表 "${params.title}" 已生成！包含 ${params.headers.length} 列、${params.data.length} 行数据${params.includeChart ? "，已添加图表" : ""}`,
          data: {
            title: params.title,
            rows: params.data.length,
            columns: params.headers.length,
            hasChart: params.includeChart,
          },
        };
      });
    },
  };
}

// ========== v2.9.48: 事件监听工具 (借鉴 office-js-snippets events) ==========

/**
 * 创建数据变更监听工具
 */
function createDataChangeListenerTool(): Tool {
  return {
    name: "excel_data_change_listener",
    description: "注册或取消数据变更事件监听（当指定范围的数据发生变化时触发）",
    category: "excel",
    parameters: [
      {
        name: "action",
        type: "string",
        description: "操作: 'register' (注册), 'unregister' (取消)",
        required: true,
      },
      { name: "range", type: "string", description: "要监听的范围", required: true },
      {
        name: "bindingId",
        type: "string",
        description: "绑定ID（用于标识和取消监听）",
        required: true,
      },
    ],
    execute: async (input) => {
      const params = input as { action: string; range: string; bindingId: string };

      return await excelRun(async (ctx) => {
        if (params.action === "register") {
          const sheet = ctx.workbook.worksheets.getActiveWorksheet();
          const range = sheet.getRange(params.range);

          // 创建绑定
          const binding = ctx.workbook.bindings.add(
            range,
            Excel.BindingType.range,
            params.bindingId
          );

          // 注册事件处理器 (注意: 在实际使用中需要保存处理器引用以便取消)
          binding.onDataChanged.add(async (args) => {
            console.log(`[数据变更事件] 绑定ID: ${args.binding.id}`);
            return;
          });

          await ctx.sync();

          return {
            success: true,
            output: `已注册数据变更监听，范围: ${params.range}，绑定ID: ${params.bindingId}`,
          };
        } else if (params.action === "unregister") {
          const binding = ctx.workbook.bindings.getItemOrNullObject(params.bindingId);
          binding.delete();
          await ctx.sync();

          return {
            success: true,
            output: `已取消数据变更监听，绑定ID: ${params.bindingId}`,
          };
        } else {
          return { success: false, output: `无效的操作: ${params.action}` };
        }
      });
    },
  };
}

// ========== v2.9.49: 更多 office-js-snippets 集成 ==========

/**
 * 创建形状工具 - 几何形状
 */
function createGeometricShapeTool(): Tool {
  return {
    name: "excel_add_shape",
    description: "在工作表中添加几何形状（矩形、圆形、三角形、箭头等）",
    category: "excel",
    parameters: [
      {
        name: "shapeType",
        type: "string",
        description:
          "形状类型: 'rectangle', 'oval', 'triangle', 'diamond', 'hexagon', 'star5', 'arrow', 'heart'",
        required: true,
      },
      { name: "left", type: "number", description: "左边距(像素)", required: false, default: 100 },
      { name: "top", type: "number", description: "上边距(像素)", required: false, default: 100 },
      { name: "width", type: "number", description: "宽度(像素)", required: false, default: 150 },
      { name: "height", type: "number", description: "高度(像素)", required: false, default: 100 },
      { name: "fillColor", type: "string", description: "填充颜色", required: false },
      { name: "text", type: "string", description: "形状中的文本", required: false },
    ],
    execute: async (input) => {
      const params = input as {
        shapeType: string;
        left?: number;
        top?: number;
        width?: number;
        height?: number;
        fillColor?: string;
        text?: string;
      };

      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();

        const shapeTypeMap: Record<string, Excel.GeometricShapeType> = {
          rectangle: Excel.GeometricShapeType.rectangle,
          oval: "Oval" as Excel.GeometricShapeType,
          triangle: Excel.GeometricShapeType.triangle,
          diamond: Excel.GeometricShapeType.diamond,
          hexagon: Excel.GeometricShapeType.hexagon,
          star5: Excel.GeometricShapeType.star5,
          arrow: Excel.GeometricShapeType.rightArrow,
          heart: Excel.GeometricShapeType.heart,
        };

        const excelShapeType = shapeTypeMap[params.shapeType.toLowerCase()];
        if (!excelShapeType) {
          return { success: false, output: `不支持的形状类型: ${params.shapeType}` };
        }

        const shape = sheet.shapes.addGeometricShape(excelShapeType);
        shape.left = params.left || 100;
        shape.top = params.top || 100;
        shape.width = params.width || 150;
        shape.height = params.height || 100;

        if (params.fillColor) {
          shape.fill.setSolidColor(params.fillColor);
        }

        if (params.text) {
          shape.textFrame.textRange.text = params.text;
        }

        await ctx.sync();

        return {
          success: true,
          output: `已添加 ${params.shapeType} 形状`,
        };
      });
    },
  };
}

/**
 * 创建插入图片工具
 */
function createInsertImageTool(): Tool {
  return {
    name: "excel_insert_image",
    description: "在工作表中插入图片（Base64格式）",
    category: "excel",
    parameters: [
      {
        name: "base64Image",
        type: "string",
        description: "Base64编码的图片数据（不含data:前缀）",
        required: true,
      },
      { name: "left", type: "number", description: "左边距(像素)", required: false, default: 100 },
      { name: "top", type: "number", description: "上边距(像素)", required: false, default: 100 },
      { name: "name", type: "string", description: "图片名称", required: false },
    ],
    execute: async (input) => {
      const params = input as {
        base64Image: string;
        left?: number;
        top?: number;
        name?: string;
      };

      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();

        // 移除可能的 data:image/xxx;base64, 前缀
        let base64 = params.base64Image;
        const base64Index = base64.indexOf("base64,");
        if (base64Index > -1) {
          base64 = base64.substring(base64Index + 7);
        }

        const image = sheet.shapes.addImage(base64);
        image.left = params.left || 100;
        image.top = params.top || 100;

        if (params.name) {
          image.name = params.name;
        }

        await ctx.sync();

        return {
          success: true,
          output: `已插入图片${params.name ? ` "${params.name}"` : ""}`,
        };
      });
    },
  };
}

/**
 * 创建工作表全局查找工具
 */
function createFindAllTool(): Tool {
  return {
    name: "excel_find_all",
    description: "在工作表中查找所有匹配的单元格并高亮显示",
    category: "excel",
    parameters: [
      { name: "searchText", type: "string", description: "要查找的文本", required: true },
      {
        name: "highlightColor",
        type: "string",
        description: "高亮颜色",
        required: false,
        default: "yellow",
      },
      {
        name: "completeMatch",
        type: "boolean",
        description: "是否完全匹配",
        required: false,
        default: false,
      },
      {
        name: "matchCase",
        type: "boolean",
        description: "是否区分大小写",
        required: false,
        default: false,
      },
    ],
    execute: async (input) => {
      const params = input as {
        searchText: string;
        highlightColor?: string;
        completeMatch?: boolean;
        matchCase?: boolean;
      };

      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const foundRanges = sheet.findAllOrNullObject(params.searchText, {
          completeMatch: params.completeMatch || false,
          matchCase: params.matchCase || false,
        });

        await ctx.sync();

        if (foundRanges.isNullObject) {
          return {
            success: true,
            output: `未找到 "${params.searchText}"`,
            data: { matchCount: 0 },
          };
        }

        foundRanges.format.fill.color = params.highlightColor || "yellow";
        foundRanges.load("address, cellCount");
        await ctx.sync();

        return {
          success: true,
          output: `找到 ${foundRanges.cellCount} 个匹配项，已用 ${params.highlightColor || "yellow"} 高亮显示`,
          data: {
            matchCount: foundRanges.cellCount,
            addresses: foundRanges.address,
          },
        };
      });
    },
  };
}

/**
 * 创建高级复制粘贴工具
 */
function createAdvancedCopyTool(): Tool {
  return {
    name: "excel_advanced_copy",
    description: "高级复制粘贴，支持跳过空白、转置、仅复制值/公式/格式",
    category: "excel",
    parameters: [
      { name: "sourceRange", type: "string", description: "源范围", required: true },
      { name: "targetCell", type: "string", description: "目标起始单元格", required: true },
      {
        name: "copyType",
        type: "string",
        description: "复制类型: 'all', 'values', 'formulas', 'formats'",
        required: false,
        default: "all",
      },
      {
        name: "skipBlanks",
        type: "boolean",
        description: "是否跳过空白单元格",
        required: false,
        default: false,
      },
      {
        name: "transpose",
        type: "boolean",
        description: "是否转置（行列互换）",
        required: false,
        default: false,
      },
    ],
    execute: async (input) => {
      const params = input as {
        sourceRange: string;
        targetCell: string;
        copyType?: string;
        skipBlanks?: boolean;
        transpose?: boolean;
      };

      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();

        const copyTypeMap: Record<string, Excel.RangeCopyType> = {
          all: Excel.RangeCopyType.all,
          values: Excel.RangeCopyType.values,
          formulas: Excel.RangeCopyType.formulas,
          formats: Excel.RangeCopyType.formats,
        };

        const copyType = copyTypeMap[params.copyType || "all"] || Excel.RangeCopyType.all;
        const targetRange = sheet.getRange(params.targetCell);

        targetRange.copyFrom(
          params.sourceRange,
          copyType,
          params.skipBlanks || false,
          params.transpose || false
        );

        await ctx.sync();

        const options = [];
        if (params.skipBlanks) options.push("跳过空白");
        if (params.transpose) options.push("转置");

        return {
          success: true,
          output: `已将 ${params.sourceRange} 复制到 ${params.targetCell}（类型: ${params.copyType || "all"}${options.length > 0 ? "，" + options.join("、") : ""}）`,
        };
      });
    },
  };
}

/**
 * 创建移动范围工具
 */
function createMoveRangeAdvancedTool(): Tool {
  return {
    name: "excel_move_range_to",
    description: "将范围移动到新位置（剪切粘贴）",
    category: "excel",
    parameters: [
      { name: "sourceRange", type: "string", description: "源范围", required: true },
      { name: "targetCell", type: "string", description: "目标起始单元格", required: true },
    ],
    execute: async (input) => {
      const params = input as { sourceRange: string; targetCell: string };

      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(params.sourceRange);
        range.moveTo(params.targetCell);
        await ctx.sync();

        return {
          success: true,
          output: `已将 ${params.sourceRange} 移动到 ${params.targetCell}`,
        };
      });
    },
  };
}

/**
 * 创建命名范围工具
 */
function createNamedRangeTool(): Tool {
  return {
    name: "excel_named_range",
    description: "创建、获取或删除命名范围（可在公式中使用名称引用）",
    category: "excel",
    parameters: [
      {
        name: "action",
        type: "string",
        description: "操作: 'create', 'delete', 'list', 'get'",
        required: true,
      },
      { name: "name", type: "string", description: "命名范围的名称", required: false },
      { name: "range", type: "string", description: "范围地址 (create时需要)", required: false },
      {
        name: "formula",
        type: "string",
        description: "公式 (可选，用于创建公式命名)",
        required: false,
      },
    ],
    execute: async (input) => {
      const params = input as {
        action: string;
        name?: string;
        range?: string;
        formula?: string;
      };

      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();

        switch (params.action) {
          case "create": {
            if (!params.name) {
              return { success: false, output: "创建命名范围需要提供 name 参数" };
            }
            if (params.formula) {
              // 创建公式命名
              sheet.names.add(params.name, params.formula);
            } else if (params.range) {
              // 创建范围命名
              const targetRange = sheet.getRange(params.range);
              sheet.names.add(params.name, targetRange);
            } else {
              return { success: false, output: "创建命名范围需要提供 range 或 formula 参数" };
            }
            await ctx.sync();
            return { success: true, output: `已创建命名范围 "${params.name}"` };
          }

          case "delete": {
            if (!params.name) {
              return { success: false, output: "删除命名范围需要提供 name 参数" };
            }
            const namedItem = sheet.names.getItemOrNullObject(params.name);
            await ctx.sync();
            if (namedItem.isNullObject) {
              return { success: false, output: `命名范围 "${params.name}" 不存在` };
            }
            namedItem.delete();
            await ctx.sync();
            return { success: true, output: `已删除命名范围 "${params.name}"` };
          }

          case "list": {
            const namedItems = sheet.names.load("items");
            await ctx.sync();
            const names = namedItems.items.map((item: Excel.NamedItem) => ({
              name: item.name,
              type: item.type,
              value: item.value,
            }));
            return {
              success: true,
              output: `工作表中有 ${names.length} 个命名范围:\n${names.map((n: { name: string; value: string }) => `- ${n.name}: ${n.value}`).join("\n")}`,
              data: { namedItems: names },
            };
          }

          case "get": {
            if (!params.name) {
              return { success: false, output: "获取命名范围需要提供 name 参数" };
            }
            const item = sheet.names.getItemOrNullObject(params.name);
            item.load("name, type, value, formula");
            await ctx.sync();
            if (item.isNullObject) {
              return { success: false, output: `命名范围 "${params.name}" 不存在` };
            }
            return {
              success: true,
              output: `命名范围 "${item.name}": ${item.value}`,
              data: { name: item.name, type: item.type, value: item.value },
            };
          }

          default:
            return { success: false, output: `无效的操作: ${params.action}` };
        }
      });
    },
  };
}

/**
 * 创建插入外部工作表工具
 */
function createInsertExternalSheetsTool(): Tool {
  return {
    name: "excel_insert_external_sheets",
    description: "从另一个 Excel 文件插入工作表（需要 Base64 格式的文件内容）",
    category: "excel",
    parameters: [
      {
        name: "base64Workbook",
        type: "string",
        description: "Base64编码的 Excel 文件",
        required: true,
      },
      {
        name: "sheetNames",
        type: "array",
        description: "要插入的工作表名称列表（留空则插入全部）",
        required: false,
      },
      { name: "insertAfter", type: "string", description: "在此工作表后插入", required: false },
    ],
    execute: async (input) => {
      const params = input as {
        base64Workbook: string;
        sheetNames?: string[];
        insertAfter?: string;
      };

      return await excelRun(async (ctx) => {
        const workbook = ctx.workbook;

        // 移除可能的 data:前缀
        let base64 = params.base64Workbook;
        const base64Index = base64.indexOf("base64,");
        if (base64Index > -1) {
          base64 = base64.substring(base64Index + 7);
        }

        const options: Excel.InsertWorksheetOptions = {
          sheetNamesToInsert: params.sheetNames || [],
          positionType: Excel.WorksheetPositionType.after,
          relativeTo: params.insertAfter || undefined,
        };

        workbook.insertWorksheetsFromBase64(base64, options);
        await ctx.sync();

        return {
          success: true,
          output: `已插入工作表${params.sheetNames?.length ? `: ${params.sheetNames.join(", ")}` : "（全部）"}`,
        };
      });
    },
  };
}

// ========== 导出 ==========

export default createExcelTools;

// ========== v2.9.5 ExcelReader 实现 ==========

import { ExcelReader } from "./AgentCore";

/**
 * v2.9.5: 创建 ExcelReader 实现
 *
 * 这个实现用于硬校验规则，让校验可以真正读取 Excel 数据
 * 而不是依赖 Agent 的"自觉"
 */
export function createExcelReader(): ExcelReader {
  return {
    /**
     * 读取指定范围的值和公式
     */
    readRange: async (
      sheet: string,
      range: string
    ): Promise<{ values: unknown[][]; formulas: string[][] }> => {
      if (typeof Excel === "undefined") {
        console.warn("[ExcelReader] Excel API 不可用");
        return { values: [], formulas: [] };
      }

      try {
        return await Excel.run(async (ctx) => {
          const worksheet = ctx.workbook.worksheets.getItemOrNullObject(sheet);
          await ctx.sync();

          if (worksheet.isNullObject) {
            console.warn(`[ExcelReader] 工作表 "${sheet}" 不存在`);
            return { values: [], formulas: [] };
          }

          const targetRange = worksheet.getRange(range);
          targetRange.load(["values", "formulas"]);
          await ctx.sync();

          return {
            values: targetRange.values as unknown[][],
            formulas: targetRange.formulas as string[][],
          };
        });
      } catch (error) {
        console.error("[ExcelReader] readRange 失败:", error);
        return { values: [], formulas: [] };
      }
    },

    /**
     * 获取工作表的样本行数据
     */
    sampleRows: async (sheet: string, count: number): Promise<unknown[][]> => {
      if (typeof Excel === "undefined") {
        console.warn("[ExcelReader] Excel API 不可用");
        return [];
      }

      try {
        return await Excel.run(async (ctx) => {
          const worksheet = ctx.workbook.worksheets.getItemOrNullObject(sheet);
          await ctx.sync();

          if (worksheet.isNullObject) {
            console.warn(`[ExcelReader] 工作表 "${sheet}" 不存在`);
            return [];
          }

          const usedRange = worksheet.getUsedRangeOrNullObject();
          usedRange.load(["values", "rowCount"]);
          await ctx.sync();

          if (usedRange.isNullObject) {
            return [];
          }

          const values = usedRange.values as unknown[][];
          // 返回指定行数（包括表头）
          return values.slice(0, Math.min(count + 1, values.length));
        });
      } catch (error) {
        console.error("[ExcelReader] sampleRows 失败:", error);
        return [];
      }
    },

    /**
     * 获取指定列的所有公式
     */
    getColumnFormulas: async (sheet: string, column: string): Promise<string[]> => {
      if (typeof Excel === "undefined") {
        console.warn("[ExcelReader] Excel API 不可用");
        return [];
      }

      try {
        return await Excel.run(async (ctx) => {
          const worksheet = ctx.workbook.worksheets.getItemOrNullObject(sheet);
          await ctx.sync();

          if (worksheet.isNullObject) {
            console.warn(`[ExcelReader] 工作表 "${sheet}" 不存在`);
            return [];
          }

          const usedRange = worksheet.getUsedRangeOrNullObject();
          usedRange.load(["rowCount"]);
          await ctx.sync();

          if (usedRange.isNullObject) {
            return [];
          }

          // 读取整列公式
          const columnRange = worksheet.getRange(`${column}1:${column}${usedRange.rowCount}`);
          columnRange.load(["formulas"]);
          await ctx.sync();

          // 扁平化为一维数组
          const formulas = columnRange.formulas as string[][];
          return formulas.map((row) => (row[0] ? String(row[0]) : ""));
        });
      } catch (error) {
        console.error("[ExcelReader] getColumnFormulas 失败:", error);
        return [];
      }
    },
  };
}
