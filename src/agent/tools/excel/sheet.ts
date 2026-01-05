/**
 * Excel 工作表类工具
 *
 * 包含工具：
 * - createCreateSheetTool: 创建工作表
 * - createSwitchSheetTool: 切换工作表
 * - createDeleteSheetTool: 删除工作表
 * - createCopySheetTool: 复制工作表
 * - createRenameSheetTool: 重命名工作表
 * - createProtectSheetTool: 保护工作表
 *
 * @packageDocumentation
 */

import { Tool } from "../../types";
import { excelRun } from "./common";

/**
 * 创建工作表工具
 */
export function createCreateSheetTool(): Tool {
  return {
    name: "excel_create_sheet",
    description: "创建一个新的工作表",
    category: "excel",
    parameters: [{ name: "name", type: "string", description: "新工作表的名称", required: true }],
    execute: async (input) => {
      const name = String(input.name);
      return await excelRun(async (ctx) => {
        const sheets = ctx.workbook.worksheets;

        sheets.load("items/name");
        await ctx.sync();

        const existingNames = sheets.items.map((s) => s.name);
        if (existingNames.includes(name)) {
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

/**
 * 切换工作表工具
 */
export function createSwitchSheetTool(): Tool {
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

/**
 * 删除工作表工具
 */
export function createDeleteSheetTool(): Tool {
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
export function createCopySheetTool(): Tool {
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
export function createRenameSheetTool(): Tool {
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
export function createProtectSheetTool(): Tool {
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
 * 创建所有工作表类工具
 */
export function createSheetTools(): Tool[] {
  return [
    createCreateSheetTool(),
    createSwitchSheetTool(),
    createDeleteSheetTool(),
    createCopySheetTool(),
    createRenameSheetTool(),
    createProtectSheetTool(),
  ];
}
