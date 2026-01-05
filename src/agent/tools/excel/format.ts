/**
 * Excel 格式化类工具
 *
 * 包含工具：
 * - createFormatRangeTool: 格式化范围
 * - createAutoFitTool: 自动调整列宽
 * - createConditionalFormatTool: 条件格式
 * - createMergeCellsTool: 合并单元格
 * - createBorderTool: 边框设置
 * - createNumberFormatTool: 数字格式
 *
 * @packageDocumentation
 */

/* eslint-disable office-addins/call-sync-after-load, office-addins/call-sync-before-read */

import { Tool } from "../../types";
import { excelRun, getTargetSheet, extractSheetName } from "./common";

// ========== 格式化类工具 ==========

export function createFormatRangeTool(): Tool {
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
      const address = String(input.address || input.range || input.cell || input.area || "A1");
      const sheetName = extractSheetName(input);

      // 如果传入了 format 对象，展开它
      const formatObj = input.format as Record<string, unknown> | undefined;
      if (formatObj && typeof formatObj === "object") {
        Object.assign(input, formatObj);
      }

      return await excelRun(async (ctx) => {
        const sheet = getTargetSheet(ctx, sheetName);
        const range = sheet.getRange(address);
        range.load("address");
        await ctx.sync();

        const actualAddress = range.address;

        if (input.fill || input.backgroundColor)
          range.format.fill.color = String(input.fill || input.backgroundColor);
        if (input.fontColor) range.format.font.color = String(input.fontColor);
        if (input.bold !== undefined) range.format.font.bold = Boolean(input.bold);
        if (input.italic !== undefined) range.format.font.italic = Boolean(input.italic);
        if (input.fontSize) range.format.font.size = Number(input.fontSize);

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

export function createAutoFitTool(): Tool {
  return {
    name: "excel_auto_fit",
    description: "自动调整列宽或行高",
    category: "excel",
    parameters: [
      { name: "address", type: "string", description: "范围地址", required: true },
      { name: "type", type: "string", description: "调整类型: columns 或 rows", required: false },
    ],
    execute: async (input) => {
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

export function createConditionalFormatTool(): Tool {
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

export function createMergeCellsTool(): Tool {
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
          range.merge(true);
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

export function createBorderTool(): Tool {
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

export function createNumberFormatTool(): Tool {
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
 * 创建所有格式化类工具
 */
export function createFormatTools(): Tool[] {
  return [
    createFormatRangeTool(),
    createAutoFitTool(),
    createConditionalFormatTool(),
    createMergeCellsTool(),
    createBorderTool(),
    createNumberFormatTool(),
  ];
}
