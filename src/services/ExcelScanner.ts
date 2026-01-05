/**
 * ExcelScanner - 工作簿扫描服务
 * v2.9.12: 从 App.tsx 提取，纯工具层，不依赖 UI 或 Agent
 *
 * 职责：扫描 Excel 工作簿，返回结构化的 WorkbookContext
 */

import type {
  WorkbookContext,
  SheetInfo,
  NamedRangeInfo,
  TableInfo,
  ChartInfo,
  PivotTableInfo,
  FormulaDependency,
} from "../taskpane/types";
import { parseFormulaReferences, analyzeFormulaComplexity } from "../taskpane/utils/dataAnalysis";

export interface ScanProgress {
  progress: number; // 0-100
  phase: string;
}

export type ProgressCallback = (progress: ScanProgress) => void;

/**
 * 扫描工作簿，返回完整的工作簿上下文
 */
export async function scanWorkbook(onProgress?: ProgressCallback): Promise<WorkbookContext | null> {
  if (typeof Excel === "undefined") {
    return null;
  }

  const updateProgress = (progress: number, phase: string) => {
    onProgress?.({ progress, phase });
  };

  try {
    updateProgress(0, "初始化");

    const context = await Excel.run(async (ctx) => {
      const workbook = ctx.workbook;
      // 在测试环境中，ctx.workbook 可能是部分模拟对象，先检查 load 是否存在
      if (workbook && typeof workbook.load === "function") {
        workbook.load("name");
      }

      // 加载所有工作表
      const sheets = workbook && workbook.worksheets;
      if (sheets && typeof sheets.load === "function") {
        sheets.load("items/name, items/position");
      }

      // 加载命名范围
      const namedItems = workbook && workbook.names;
      if (namedItems && typeof namedItems.load === "function") {
        namedItems.load("items/name, items/value, items/comment, items/scope");
      }

      await ctx.sync();
      updateProgress(20, "扫描工作表");

      const sheetInfos: SheetInfo[] = [];
      const tableInfos: TableInfo[] = [];
      const chartInfos: ChartInfo[] = [];
      const pivotTableInfos: PivotTableInfo[] = [];
      let totalCellsWithData = 0;
      let totalFormulas = 0;
      const issues: WorkbookContext["issues"] = [];

      // 公式依赖收集
      const formulaDeps: FormulaDependency[] = [];
      const complexFormulas: Array<{ address: string; formula: string; complexity: string }> = [];

      // 扫描每个工作表
      for (let i = 0; i < sheets.items.length; i++) {
        const sheet = sheets.items[i];
        sheet.load("name, position");

        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load("address, rowCount, columnCount, values, formulas");

        const tables = sheet.tables;
        tables.load("items/name, items/showHeaders, items/style");

        const charts = sheet.charts;
        charts.load("items/name, items/chartType, items/title/text");

        const pivots = sheet.pivotTables;
        pivots.load("items/name");

        await ctx.sync();
        updateProgress(
          20 + Math.round((i / sheets.items.length) * 60),
          `扫描工作表: ${sheet.name}`
        );

        const hasData = !usedRange.isNullObject;
        let rowCount = 0;
        let columnCount = 0;
        let usedRangeAddress = "";

        if (hasData) {
          rowCount = usedRange.rowCount;
          columnCount = usedRange.columnCount;
          usedRangeAddress = usedRange.address;
          totalCellsWithData += rowCount * columnCount;

          // 统计公式并分析复杂度
          if (usedRange.formulas) {
            // 获取起始单元格地址
            const addressMatch = usedRange.address.match(/!?([A-Z]+)(\d+)/);
            const startCol = addressMatch ? addressMatch[1] : "A";
            const startRow = addressMatch ? parseInt(addressMatch[2], 10) : 1;

            for (let r = 0; r < usedRange.formulas.length; r++) {
              const formulaRow = usedRange.formulas[r];
              for (let c = 0; c < formulaRow.length; c++) {
                const cell = formulaRow[c];
                if (typeof cell === "string" && cell.startsWith("=")) {
                  totalFormulas++;

                  // 分析公式复杂度 (只分析前100个公式)
                  if (formulaDeps.length < 100) {
                    const cellAddr = `${String.fromCharCode(startCol.charCodeAt(0) + c)}${startRow + r}`;
                    const refs = parseFormulaReferences(cell);
                    const complexity = analyzeFormulaComplexity(cell);

                    formulaDeps.push({
                      cellAddress: `${sheet.name}!${cellAddr}`,
                      formula: cell.length > 100 ? cell.substring(0, 100) + "..." : cell,
                      dependsOn: refs,
                      usedBy: [], // 暂不计算反向依赖
                    });

                    if (complexity.level === "complex") {
                      complexFormulas.push({
                        address: `${sheet.name}!${cellAddr}`,
                        formula: cell.length > 50 ? cell.substring(0, 50) + "..." : cell,
                        complexity: `${complexity.functions.join(", ")} (嵌套深度高)`,
                      });
                    }
                  }
                }
              }
            }
          }
        }

        // 收集工作表信息
        sheetInfos.push({
          name: sheet.name,
          index: sheet.position,
          isActive: i === 0, // 暂时假定第一个是活动
          usedRangeAddress,
          rowCount,
          columnCount,
          hasData,
          hasTables: tables.items.length > 0,
          hasCharts: charts.items.length > 0,
          hasPivotTables: pivots.items.length > 0,
        });

        // 收集表格信息
        for (const table of tables.items) {
          const tableRange = table.getRange();
          tableRange.load("address, rowCount, columnCount");
          const headerRow = table.getHeaderRowRange();
          headerRow.load("values");

          await ctx.sync();

          tableInfos.push({
            name: table.name,
            sheetName: sheet.name,
            address: tableRange.address,
            rowCount: tableRange.rowCount,
            columnCount: tableRange.columnCount,
            hasHeaders: table.showHeaders,
            columns: headerRow.values[0]?.map((v) => String(v)) || [],
            style: table.style,
          });
        }

        // 收集图表信息
        for (const chart of charts.items) {
          chartInfos.push({
            name: chart.name,
            sheetName: sheet.name,
            chartType: chart.chartType,
            title: chart.title?.text,
          });
        }

        // 收集透视表信息
        for (const pivot of pivots.items) {
          pivotTableInfos.push({
            name: pivot.name,
            sheetName: sheet.name,
          });
        }
      }

      updateProgress(85, "分析数据质量");

      // 收集命名范围信息
      const namedRangeInfos: NamedRangeInfo[] = namedItems.items.map((item) => ({
        name: item.name,
        address: item.value,
        scope: item.scope === Excel.NamedItemScope.workbook ? "workbook" : "worksheet",
        comment: item.comment,
      }));

      // 创建快速查找索引
      const sheetByName: Record<string, SheetInfo> = {};
      sheetInfos.forEach((s) => {
        sheetByName[s.name] = s;
      });

      const tableByName: Record<string, TableInfo> = {};
      tableInfos.forEach((t) => {
        tableByName[t.name] = t;
      });

      // 检测潜在问题
      for (const sheet of sheetInfos) {
        if (sheet.rowCount > 10000) {
          issues.push({
            type: "warning",
            message: `工作表 "${sheet.name}" 数据量较大 (${sheet.rowCount}行)，可能影响性能`,
            location: sheet.name,
          });
        }
        if (!sheet.hasData) {
          issues.push({
            type: "suggestion",
            message: `工作表 "${sheet.name}" 为空`,
            location: sheet.name,
          });
        }
      }

      // 添加复杂公式警告
      for (const cf of complexFormulas) {
        issues.push({
          type: "warning",
          message: `复杂公式 ${cf.address}: ${cf.complexity}`,
          location: cf.address,
        });
      }

      // 计算数据质量分数
      const hasTablesBonus = tableInfos.length > 0 ? 10 : 0;
      const hasChartsBonus = chartInfos.length > 0 ? 10 : 0;
      const hasNamesBonus = namedRangeInfos.length > 0 ? 5 : 0;
      const sheetsOrganizedBonus = sheetInfos.every((s) => s.hasData) ? 10 : 0;
      const formulaComplexityPenalty = complexFormulas.length > 5 ? -5 : 0;
      const overallQualityScore = Math.min(
        100,
        Math.max(
          0,
          65 +
            hasTablesBonus +
            hasChartsBonus +
            hasNamesBonus +
            sheetsOrganizedBonus +
            formulaComplexityPenalty
        )
      );

      updateProgress(100, "完成");

      return {
        lastScanned: new Date(),
        fileName: workbook.name || "未命名",
        sheets: sheetInfos,
        namedRanges: namedRangeInfos,
        tables: tableInfos,
        charts: chartInfos,
        pivotTables: pivotTableInfos,
        totalCellsWithData,
        totalFormulas,
        formulaDependencies: formulaDeps,
        dataRelationships: [], // 数据关系，待后续扩展
        sheetByName,
        tableByName,
        overallQualityScore,
        issues,
      };
    });

    return context;
  } catch (error) {
    console.error("工作簿扫描失败:", error);
    return null;
  }
}

/**
 * 验证操作执行结果
 */
export interface OperationVerification {
  success: boolean;
  operationType: string;
  targetAddress: string;
  expectedResult?: unknown;
  actualResult?: unknown;
  matchesExpectation: boolean;
  details: string;
  timestamp: Date;
}

export async function verifyOperationResult(
  operationType: string,
  targetAddress: string,
  expectedValue?: unknown
): Promise<OperationVerification> {
  const verification: OperationVerification = {
    success: false,
    operationType,
    targetAddress,
    expectedResult: expectedValue,
    actualResult: undefined,
    matchesExpectation: false,
    details: "",
    timestamp: new Date(),
  };

  if (typeof Excel === "undefined") {
    verification.details = "Excel API 不可用";
    return verification;
  }

  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();

      // 解析地址，支持跨表地址
      let targetRange: Excel.Range;
      if (targetAddress.includes("!")) {
        const [sheetName, addr] = targetAddress.split("!");
        const targetSheet = ctx.workbook.worksheets.getItem(sheetName.replace(/'/g, ""));
        targetRange = targetSheet.getRange(addr);
      } else {
        targetRange = sheet.getRange(targetAddress);
      }

      targetRange.load("values, formulas, numberFormat");
      await ctx.sync();

      verification.actualResult = targetRange.values;

      // 根据操作类型验证结果
      switch (operationType.toLowerCase()) {
        case "write_range":
        case "writerange":
        case "insert_data": {
          // 验证数据已写入
          const hasData = targetRange.values.some((row) =>
            row.some((cell) => cell !== null && cell !== undefined && cell !== "")
          );
          verification.success = hasData;
          verification.details = hasData ? "数据已成功写入" : "写入后单元格为空";
          break;
        }

        case "set_formula":
        case "setformula": {
          // 验证公式已设置
          const formula = targetRange.formulas[0]?.[0];
          const hasFormula = typeof formula === "string" && formula.startsWith("=");
          verification.success = hasFormula;
          verification.details = hasFormula ? `公式已设置: ${formula}` : "公式未设置";

          // 检查公式错误
          const formulaResult = targetRange.values[0]?.[0];
          if (typeof formulaResult === "string" && formulaResult.startsWith("#")) {
            verification.success = false;
            verification.details = `公式错误: ${formulaResult}`;
          }
          break;
        }

        case "clear_range":
        case "clearrange": {
          // 验证已清空
          const isEmpty = targetRange.values.every((row) =>
            row.every((cell) => cell === null || cell === undefined || cell === "")
          );
          verification.success = isEmpty;
          verification.details = isEmpty ? "数据已清空" : "数据未完全清空";
          break;
        }

        default:
          // 默认检查是否有异常
          verification.success = true;
          verification.details = "操作已执行";
      }

      // 如果有预期值，进行比对
      if (expectedValue !== undefined) {
        const actualFirst = targetRange.values[0]?.[0];
        verification.matchesExpectation = actualFirst === expectedValue;
        if (!verification.matchesExpectation) {
          verification.details += ` (预期: ${expectedValue}, 实际: ${actualFirst})`;
        }
      }
    });
  } catch (error) {
    verification.success = false;
    verification.details = `验证失败: ${error instanceof Error ? error.message : String(error)}`;
  }

  return verification;
}
