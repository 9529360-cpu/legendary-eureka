/**
 * Excel 分析类工具
 *
 * 包含工具：
 * - createDataValidationTool: 数据验证
 * - createTableTool: 创建表格
 * - createPivotTableTool: 创建透视表
 * - createGoalSeekTool: 目标求解
 * - createFreezePanesTool: 冻结窗格
 * - createGroupRowsTool: 行分组
 * - createGroupColumnsTool: 列分组
 * - createCommentTool: 批注
 *
 * @packageDocumentation
 */

/* eslint-disable office-addins/call-sync-after-load, office-addins/call-sync-before-read, office-addins/load-object-before-read */

import { Tool } from "../../types";
import { excelRun } from "./common";

/**
 * 数据验证工具
 */
export function createDataValidationTool(): Tool {
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

/**
 * 创建表格工具
 */
export function createTableTool(): Tool {
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
export function createPivotTableTool(): Tool {
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
 * 目标求解工具
 */
export function createGoalSeekTool(): Tool {
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

/**
 * 冻结窗格工具
 */
export function createFreezePanesTool(): Tool {
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
export function createGroupRowsTool(): Tool {
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
export function createGroupColumnsTool(): Tool {
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
export function createCommentTool(): Tool {
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

        if (action === "add") {
          ctx.workbook.comments.add(cell, params.content as string);
        } else if (action === "edit" || action === "delete") {
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
 * 创建所有分析类工具
 */
export function createAnalysisTools(): Tool[] {
  return [
    createDataValidationTool(),
    createTableTool(),
    createPivotTableTool(),
    createGoalSeekTool(),
    createFreezePanesTool(),
    createGroupRowsTool(),
    createGroupColumnsTool(),
    createCommentTool(),
  ];
}
