/**
 * Excel 高级工具
 *
 * 包含工具：
 * - createHyperlinkTool: 添加超链接
 * - createPageSetupTool: 页面设置
 * - createPrintAreaTool: 打印区域
 * - createBatchWriteOptimizedTool: 批量写入优化
 * - createPerformanceModeTool: 性能模式
 * - createRecalculateTool: 重新计算
 * - createAdvancedConditionalFormatTool: 高级条件格式
 * - createClearConditionalFormatsTool: 清除条件格式
 * - createQuickReportTool: 快速报表
 * - createGeometricShapeTool: 几何形状
 * - createFindAllTool: 全局查找
 * - createAdvancedCopyTool: 高级复制
 * - createNamedRangeTool: 命名范围
 *
 * @packageDocumentation
 */

import { Tool } from "../../types";
import { excelRun } from "./common";

/**
 * 超链接工具
 */
export function createHyperlinkTool(): Tool {
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
    ],
    execute: async (params) => {
      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getItem(params.sheet as string);
        const cell = sheet.getRange(params.cell as string);

        const url = params.url as string;
        const displayText = (params.displayText as string) || url;

        if (url.startsWith("http")) {
          cell.formulas = [[`=HYPERLINK("${url}","${displayText}")`]];
        } else {
          cell.formulas = [[`=HYPERLINK("#${url}","${displayText}")`]];
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
export function createPageSetupTool(): Tool {
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
        const pageLayout = sheet.pageLayout;

        if (params.orientation) {
          pageLayout.orientation =
            params.orientation === "landscape"
              ? Excel.PageOrientation.landscape
              : Excel.PageOrientation.portrait;
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
export function createPrintAreaTool(): Tool {
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
 * 批量写入优化工具
 */
export function createBatchWriteOptimizedTool(): Tool {
  return {
    name: "excel_batch_write_optimized",
    description: "优化的批量写入工具，适用于大量数据写入",
    category: "excel",
    parameters: [
      { name: "startCell", type: "string", description: "起始单元格，如 A1", required: true },
      { name: "data", type: "array", description: "二维数据数组", required: true },
    ],
    execute: async (input) => {
      const params = input as { startCell: string; data: unknown[][] };

      return await excelRun(async (ctx) => {
        const startTime = Date.now();
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();

        ctx.application.suspendScreenUpdatingUntilNextSync();

        const rowCount = params.data.length;
        const colCount = params.data[0]?.length || 0;

        const startCell = sheet.getRange(params.startCell);
        const targetRange = startCell.getResizedRange(rowCount - 1, colCount - 1);

        targetRange.values = params.data;
        targetRange.untrack();

        await ctx.sync();
        const elapsed = Date.now() - startTime;

        return {
          success: true,
          output: `批量写入完成！${rowCount} 行 × ${colCount} 列，共 ${rowCount * colCount} 个单元格，耗时 ${elapsed}ms`,
        };
      });
    },
  };
}

/**
 * 性能模式工具
 */
export function createPerformanceModeTool(): Tool {
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
          };
        }

        if (params.mode === "manual") {
          ctx.application.calculationMode = Excel.CalculationMode.manual;
        } else if (params.mode === "automatic") {
          ctx.application.calculationMode = Excel.CalculationMode.automatic;
        } else {
          return {
            success: false,
            output: `无效的模式: ${params.mode}`,
          };
        }

        await ctx.sync();

        return {
          success: true,
          output: `计算模式已切换: ${currentMode} → ${params.mode}`,
        };
      });
    },
  };
}

/**
 * 重新计算工具
 */
export function createRecalculateTool(): Tool {
  return {
    name: "excel_recalculate",
    description: "手动触发 Excel 重新计算",
    category: "excel",
    parameters: [
      {
        name: "type",
        type: "string",
        description: "计算类型: 'full' (完整), 'fullRebuild' (完整重建)",
        required: false,
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

/**
 * 清除条件格式工具
 */
export function createClearConditionalFormatsTool(): Tool {
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

/**
 * 几何形状工具
 */
export function createGeometricShapeTool(): Tool {
  return {
    name: "excel_add_shape",
    description: "在工作表中添加几何形状",
    category: "excel",
    parameters: [
      {
        name: "shapeType",
        type: "string",
        description: "形状类型: 'rectangle', 'oval', 'triangle', 'diamond', 'star5', 'arrow', 'heart'",
        required: true,
      },
      { name: "left", type: "number", description: "左边距(像素)", required: false },
      { name: "top", type: "number", description: "上边距(像素)", required: false },
      { name: "width", type: "number", description: "宽度(像素)", required: false },
      { name: "height", type: "number", description: "高度(像素)", required: false },
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
 * 全局查找工具
 */
export function createFindAllTool(): Tool {
  return {
    name: "excel_find_all",
    description: "在工作表中查找所有匹配的单元格并高亮显示",
    category: "excel",
    parameters: [
      { name: "searchText", type: "string", description: "要查找的文本", required: true },
      { name: "highlightColor", type: "string", description: "高亮颜色", required: false },
      { name: "completeMatch", type: "boolean", description: "是否完全匹配", required: false },
      { name: "matchCase", type: "boolean", description: "是否区分大小写", required: false },
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
          };
        }

        foundRanges.format.fill.color = params.highlightColor || "yellow";
        foundRanges.load("address, cellCount");
        await ctx.sync();

        return {
          success: true,
          output: `找到 ${foundRanges.cellCount} 个匹配项，已高亮显示`,
        };
      });
    },
  };
}

/**
 * 高级复制工具
 */
export function createAdvancedCopyTool(): Tool {
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
      },
      { name: "skipBlanks", type: "boolean", description: "是否跳过空白单元格", required: false },
      { name: "transpose", type: "boolean", description: "是否转置", required: false },
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

        return {
          success: true,
          output: `已将 ${params.sourceRange} 复制到 ${params.targetCell}`,
        };
      });
    },
  };
}

/**
 * 命名范围工具
 */
export function createNamedRangeTool(): Tool {
  return {
    name: "excel_named_range",
    description: "创建、获取或删除命名范围",
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
    ],
    execute: async (input) => {
      const params = input as {
        action: string;
        name?: string;
        range?: string;
      };

      return await excelRun(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();

        switch (params.action) {
          case "create": {
            if (!params.name || !params.range) {
              return { success: false, output: "创建命名范围需要提供 name 和 range 参数" };
            }
            const targetRange = sheet.getRange(params.range);
            sheet.names.add(params.name, targetRange);
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
              value: item.value,
            }));
            return {
              success: true,
              output: `工作表中有 ${names.length} 个命名范围`,
            };
          }

          case "get": {
            if (!params.name) {
              return { success: false, output: "获取命名范围需要提供 name 参数" };
            }
            const item = sheet.names.getItemOrNullObject(params.name);
            item.load("name, value");
            await ctx.sync();
            if (item.isNullObject) {
              return { success: false, output: `命名范围 "${params.name}" 不存在` };
            }
            return {
              success: true,
              output: `命名范围 "${item.name}": ${item.value}`,
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
 * 创建所有高级工具
 */
export function createAdvancedTools(): Tool[] {
  return [
    createHyperlinkTool(),
    createPageSetupTool(),
    createPrintAreaTool(),
    createBatchWriteOptimizedTool(),
    createPerformanceModeTool(),
    createRecalculateTool(),
    createClearConditionalFormatsTool(),
    createGeometricShapeTool(),
    createFindAllTool(),
    createAdvancedCopyTool(),
    createNamedRangeTool(),
  ];
}
