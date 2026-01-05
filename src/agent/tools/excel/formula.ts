/**
 * Excel 公式类工具
 *
 * 包含工具：
 * - createSetFormulaTool: 设置单个公式
 * - createBatchFormulaTool: 批量设置公式
 * - createSetFormulasTool: excel_set_formulas 别名
 * - createFillFormulaTool: excel_fill_formula 别名
 * - createSmartFormulaTool: 智能公式（支持结构化引用）
 *
 * @packageDocumentation
 */

import { Tool } from "../../types";
import { excelRun, getTargetSheet, extractSheetName } from "./common";

// ========== 公式类工具 ==========

export function createSetFormulaTool(): Tool {
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

export function createBatchFormulaTool(): Tool {
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
export function createSetFormulasTool(): Tool {
  const baseTool = createBatchFormulaTool();
  return {
    ...baseTool,
    name: "excel_set_formulas",
    description: "批量设置公式到范围（excel_batch_formula 的别名）",
  };
}

export function createFillFormulaTool(): Tool {
  const baseTool = createBatchFormulaTool();
  return {
    ...baseTool,
    name: "excel_fill_formula",
    description: "填充公式到范围（excel_batch_formula 的别名）",
  };
}

// ========== 辅助函数 ==========

/**
 * v2.9.56: 将 @[字段名] 转换为 Excel Table 结构化引用
 * 输入: "=@[单价]*@[数量]"
 * 输出: "=[@单价]*[@数量]"
 */
function convertToTableReference(formula: string): string {
  // Excel Table 结构化引用格式: [@列名]
  return formula.replace(/@\[([^\]]+)\]/g, "[@$1]");
}

/**
 * v2.9.56: 列索引转列字母
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

/**
 * v2.9.56: 智能公式工具
 *
 * 支持三种模式：
 * 1. 结构化引用: "=@[单价]*@[数量]" - 转换为 Excel Table 结构化引用
 * 2. 行模板: "=B{row}*C{row}" - 按行展开为 B2*C2, B3*C3...
 * 3. 列标识模式: 只指定列，根据 usedRange 自动确定行数
 *
 * 核心原则：
 * - 不写死行号，根据实际数据行数决定
 * - 支持 Table 结构化引用（最佳实践）
 * - 返回实际影响范围和结果验证
 */
export function createSmartFormulaTool(): Tool {
  return {
    name: "excel_smart_formula",
    description: "智能公式工具：支持结构化引用(@[字段])或行模板({row})，自动根据真实数据行数写入",
    category: "excel",
    parameters: [
      {
        name: "sheet",
        type: "string",
        description: "工作表名称",
        required: true,
      },
      {
        name: "column",
        type: "string",
        description: "目标列（如 D 或 E），不需要指定行号",
        required: true,
      },
      {
        name: "logicalFormula",
        type: "string",
        description: '逻辑公式，如 "=@[单价]*@[数量]" 或 "=B{row}*C{row}"',
        required: true,
      },
      {
        name: "referenceMode",
        type: "string",
        description: "引用模式: structured(结构化), row_template(行模板), a1_fixed(A1固定)",
        required: false,
      },
      {
        name: "startRow",
        type: "number",
        description: "起始行（默认2，跳过表头）",
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
        return { success: false, output: "缺少 logicalFormula 参数" };
      }

      return await excelRun(async (ctx) => {
        const sheet = getTargetSheet(ctx, sheetName);

        // 1. 获取实际数据行数
        const usedRange = sheet.getUsedRange();
        usedRange.load("rowCount, columnCount, address");
        await ctx.sync();

        const dataEndRow = usedRange.rowCount; // 实际有数据的行数

        if (dataEndRow <= 1) {
          return {
            success: false,
            output: `工作表 "${sheetName}" 没有数据行（只有表头或为空）`,
          };
        }

        // 2. 检查是否有 Excel Table
        const tables = sheet.tables;
        tables.load("items/name, items/columns/items/name");
        await ctx.sync();

        let useTableRef = false;
        let tableName = "";

        if (tables.items.length > 0) {
          // 有 Table，可以使用结构化引用
          const table = tables.items[0]; // 使用第一个 Table
          tableName = table.name;
          useTableRef = referenceMode === "structured";
        }

        // 3. 读取表头行用于字段名到列号的映射
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

        // 4. 转换公式
        const finalFormulas: string[][] = [];
        const targetRange = `${column}${startRow}:${column}${dataEndRow}`;

        if (useTableRef && tableName) {
          // 结构化引用模式：写入一次，Excel 自动扩展
          // 把 @[字段名] 转换为 Excel Table 结构化引用
          const tableFormula = convertToTableReference(logicalFormula);

          // Table 结构化引用只需写入一格，Excel 自动处理
          const singleCell = sheet.getRange(`${column}${startRow}`);
          singleCell.formulas = [[tableFormula]];
          await ctx.sync();

          // 验证结果
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
              ? `公式写入 ${targetRange} 但存在错误值，请检查字段名是否正确`
              : `已使用 Table "${tableName}" 结构化引用写入 ${targetRange}（${dataEndRow - startRow + 1}行），首结果: ${results[0]}，末结果: ${results[results.length - 1]}`,
            data: {
              affectedRange: targetRange,
              affectedRows: dataEndRow - startRow + 1,
              formula: tableFormula,
              sampleResults: results.slice(0, 3),
            },
          };
        } else {
          // 行模板模式：按行展开公式
          for (let row = startRow; row <= dataEndRow; row++) {
            let rowFormula = logicalFormula;

            // 替换 {row} 占位符
            rowFormula = rowFormula.replace(/\{row\}/g, String(row));

            // 替换 @[字段名] 为具体单元格引用
            rowFormula = rowFormula.replace(/@\[([^\]]+)\]/g, (_, fieldName) => {
              const col = fieldToColumn[fieldName];
              if (col) {
                return `${col}${row}`;
              }
              console.warn(`[SmartFormula] 未找到字段: ${fieldName}`);
              return `@[${fieldName}]`; // 保留原样，让 Excel 报错
            });

            finalFormulas.push([rowFormula]);
          }

          // 写入公式
          const range = sheet.getRange(targetRange);
          range.formulas = finalFormulas;
          await ctx.sync();

          // 验证结果
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
              ? `公式写入 ${targetRange} 但存在错误值`
              : `已将公式写入 ${targetRange}（${dataEndRow - startRow + 1}行），首结果: ${results[0]}，末结果: ${results[results.length - 1]}`,
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
 * 创建所有公式类工具
 */
export function createFormulaTools(): Tool[] {
  return [
    createSetFormulaTool(),
    createBatchFormulaTool(),
    createSetFormulasTool(),
    createFillFormulaTool(),
    createSmartFormulaTool(),
  ];
}
