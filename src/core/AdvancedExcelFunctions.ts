/**
 * AdvancedExcelFunctions - 高级 Excel 功能实现
 *
 * 功能：
 * - 智能表格美化
 * - 条件格式化
 * - 数据透视表
 * - 高级图表
 * - 数据验证
 *
 * @version 1.0.0
 */

import { Logger } from "../utils/Logger";
import { SecurityManager } from "./SecurityManager";
import { TraceContext, SpanType } from "./TraceContext";

// ============ 类型定义 ============

/**
 * 表格样式预设
 */
export interface TableStylePreset {
  id: string;
  name: string;
  description: string;
  headerStyle: CellStyle;
  bodyStyle: CellStyle;
  alternateRowStyle?: CellStyle;
  totalRowStyle?: CellStyle;
  borderStyle: BorderStyle;
}

/**
 * 单元格样式
 */
export interface CellStyle {
  fill?: {
    type: "solid" | "gradient" | "pattern";
    color?: string;
    colors?: string[];
    pattern?: string;
  };
  font?: {
    name?: string;
    size?: number;
    bold?: boolean;
    italic?: boolean;
    color?: string;
    underline?: boolean;
  };
  alignment?: {
    horizontal?: "left" | "center" | "right" | "justify";
    vertical?: "top" | "middle" | "bottom";
    wrapText?: boolean;
  };
  numberFormat?: string;
}

/**
 * 边框样式
 */
export interface BorderStyle {
  style: "none" | "thin" | "medium" | "thick" | "double" | "dashed";
  color?: string;
  sides?: ("top" | "bottom" | "left" | "right" | "inside" | "outside")[];
}

/**
 * 条件格式规则
 */
export interface ConditionalFormatRule {
  type: "cellValue" | "dataBar" | "colorScale" | "iconSet" | "topBottom" | "text" | "duplicate";
  range: string;
  // 条件
  operator?:
    | "greaterThan"
    | "lessThan"
    | "between"
    | "equalTo"
    | "notEqualTo"
    | "contains"
    | "beginsWith"
    | "endsWith";
  values?: (string | number)[];
  text?: string;
  // 格式
  format?: CellStyle;
  // 数据条
  dataBarColor?: string;
  dataBarShowValue?: boolean;
  // 色阶
  colorScaleColors?: string[];
  // 图标集
  iconSetType?: "threeArrows" | "threeFlags" | "threeTrafficLights" | "fourArrows" | "fiveArrows";
  // Top/Bottom
  topBottomCount?: number;
  topBottomPercent?: boolean;
  topBottomType?: "top" | "bottom";
}

/**
 * 图表配置
 */
export interface ChartConfig {
  type: "column" | "bar" | "line" | "pie" | "area" | "scatter" | "doughnut" | "combo";
  dataRange: string;
  title?: string;
  legend?: {
    position: "top" | "bottom" | "left" | "right" | "none";
  };
  series?: Array<{
    name?: string;
    dataRange: string;
    chartType?: string;
    color?: string;
  }>;
  axes?: {
    xAxis?: { title?: string; min?: number; max?: number };
    yAxis?: { title?: string; min?: number; max?: number };
  };
  style?: "default" | "colorful" | "monochrome" | "dark";
}

/**
 * 数据验证规则
 */
export interface DataValidationRule {
  type: "whole" | "decimal" | "list" | "date" | "time" | "textLength" | "custom";
  range: string;
  operator?:
    | "between"
    | "notBetween"
    | "equalTo"
    | "notEqualTo"
    | "greaterThan"
    | "lessThan"
    | "greaterThanOrEqualTo"
    | "lessThanOrEqualTo";
  values?: (string | number | Date)[];
  listItems?: string[];
  formula?: string;
  errorMessage?: {
    title: string;
    message: string;
    style: "stop" | "warning" | "information";
  };
  inputMessage?: {
    title: string;
    message: string;
  };
}

/**
 * 数据透视表配置
 */
export interface PivotTableConfig {
  sourceRange: string;
  targetCell: string;
  name?: string;
  rows: string[];
  columns: string[];
  values: Array<{
    field: string;
    summarizeBy: "sum" | "count" | "average" | "max" | "min" | "product";
    displayAs?: "default" | "percentOfRowTotal" | "percentOfColumnTotal" | "percentOfGrandTotal";
  }>;
  filters?: string[];
  showGrandTotals?: {
    rows: boolean;
    columns: boolean;
  };
}

// ============ 预设样式 ============

export const TABLE_STYLE_PRESETS: TableStylePreset[] = [
  {
    id: "professional-blue",
    name: "专业蓝",
    description: "经典的蓝色商务风格",
    headerStyle: {
      fill: { type: "solid", color: "#4472C4" },
      font: { bold: true, color: "#FFFFFF", size: 12 },
      alignment: { horizontal: "center", vertical: "middle" },
    },
    bodyStyle: {
      font: { size: 11 },
      alignment: { vertical: "middle" },
    },
    alternateRowStyle: {
      fill: { type: "solid", color: "#D6DCE5" },
    },
    borderStyle: {
      style: "thin",
      color: "#8EA9DB",
      sides: ["top", "bottom", "left", "right"],
    },
  },
  {
    id: "modern-green",
    name: "现代绿",
    description: "清新的绿色现代风格",
    headerStyle: {
      fill: { type: "solid", color: "#70AD47" },
      font: { bold: true, color: "#FFFFFF", size: 12 },
      alignment: { horizontal: "center", vertical: "middle" },
    },
    bodyStyle: {
      font: { size: 11 },
      alignment: { vertical: "middle" },
    },
    alternateRowStyle: {
      fill: { type: "solid", color: "#E2EFDA" },
    },
    borderStyle: {
      style: "thin",
      color: "#A9D08E",
      sides: ["top", "bottom", "left", "right"],
    },
  },
  {
    id: "elegant-gray",
    name: "优雅灰",
    description: "简约的灰色优雅风格",
    headerStyle: {
      fill: { type: "solid", color: "#404040" },
      font: { bold: true, color: "#FFFFFF", size: 12 },
      alignment: { horizontal: "center", vertical: "middle" },
    },
    bodyStyle: {
      font: { size: 11 },
      alignment: { vertical: "middle" },
    },
    alternateRowStyle: {
      fill: { type: "solid", color: "#F2F2F2" },
    },
    borderStyle: {
      style: "thin",
      color: "#BFBFBF",
      sides: ["top", "bottom", "left", "right"],
    },
  },
  {
    id: "vibrant-orange",
    name: "活力橙",
    description: "充满活力的橙色风格",
    headerStyle: {
      fill: { type: "solid", color: "#ED7D31" },
      font: { bold: true, color: "#FFFFFF", size: 12 },
      alignment: { horizontal: "center", vertical: "middle" },
    },
    bodyStyle: {
      font: { size: 11 },
      alignment: { vertical: "middle" },
    },
    alternateRowStyle: {
      fill: { type: "solid", color: "#FCE4D6" },
    },
    borderStyle: {
      style: "thin",
      color: "#F4B084",
      sides: ["top", "bottom", "left", "right"],
    },
  },
  {
    id: "minimal",
    name: "极简",
    description: "极简主义风格，仅边框",
    headerStyle: {
      font: { bold: true, size: 12 },
      alignment: { horizontal: "center", vertical: "middle" },
    },
    bodyStyle: {
      font: { size: 11 },
      alignment: { vertical: "middle" },
    },
    borderStyle: {
      style: "thin",
      color: "#D9D9D9",
      sides: ["outside"],
    },
  },
];

// ============ AdvancedExcelFunctions 类 ============

class AdvancedExcelFunctionsImpl {
  /**
   * 智能美化表格
   */
  async beautifyTable(
    range: string,
    styleId: string = "professional-blue",
    options?: {
      autoFitColumns?: boolean;
      freezeHeader?: boolean;
      addFilters?: boolean;
      detectHeaderRow?: boolean;
    }
  ): Promise<{ success: boolean; message: string }> {
    const _span = TraceContext.startSpan("beautifyTable", SpanType.EXCEL);
    TraceContext.setSpanAttribute("range", range);
    TraceContext.setSpanAttribute("styleId", styleId);

    try {
      // 验证输入
      const validation = SecurityManager.validateInput(range, { type: "range" });
      if (!validation.valid) {
        throw new Error(`无效的范围: ${validation.errors.join(", ")}`);
      }

      // 检查权限
      const permission = SecurityManager.checkPermission("excel_format_range");
      if (!permission.allowed) {
        throw new Error(permission.reason);
      }

      // 获取样式预设
      const preset = TABLE_STYLE_PRESETS.find((p) => p.id === styleId);
      if (!preset) {
        throw new Error(`未找到样式预设: ${styleId}`);
      }

      // 执行格式化
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const targetRange = sheet.getRange(range);

        // 加载范围属性
        targetRange.load(["rowCount", "columnCount", "values"]);
        await context.sync();

        const rowCount = targetRange.rowCount;
        const colCount = targetRange.columnCount;

        // 应用表头样式（第一行）
        const headerRange = targetRange.getRow(0);
        this.applyCellStyle(headerRange, preset.headerStyle);

        // 应用表体样式
        if (rowCount > 1) {
          const bodyRange = targetRange
            .getOffsetRange(1, 0)
            .getResizedRange(-(rowCount - (rowCount - 1)), 0);
          this.applyCellStyle(bodyRange, preset.bodyStyle);

          // 应用交替行样式
          if (preset.alternateRowStyle) {
            for (let i = 1; i < rowCount; i += 2) {
              const altRow = targetRange.getRow(i);
              this.applyCellStyle(altRow, preset.alternateRowStyle);
            }
          }
        }

        // 应用边框
        this.applyBorderStyle(targetRange, preset.borderStyle);

        // 可选：自动调整列宽
        if (options?.autoFitColumns !== false) {
          for (let c = 0; c < colCount; c++) {
            targetRange.getColumn(c).format.autofitColumns();
          }
        }

        // 可选：冻结表头
        if (options?.freezeHeader) {
          sheet.freezePanes.freezeRows(1);
        }

        // 可选：添加筛选器
        if (options?.addFilters) {
          targetRange.getRow(0).format.rowHeight = 20;
          sheet.tables.add(targetRange, true);
        }

        await context.sync();
      });

      Logger.info("[AdvancedExcelFunctions] 表格美化完成", { range, styleId });
      TraceContext.endSpan();

      return {
        success: true,
        message: `已使用"${preset.name}"样式美化表格`,
      };
    } catch (error) {
      Logger.error("[AdvancedExcelFunctions] 表格美化失败", error);
      TraceContext.setSpanError(error as Error);
      TraceContext.endSpan();

      return {
        success: false,
        message: `表格美化失败: ${(error as Error).message}`,
      };
    }
  }

  /**
   * 添加条件格式
   */
  async addConditionalFormat(
    rules: ConditionalFormatRule[]
  ): Promise<{ success: boolean; message: string; appliedRules: number }> {
    const _span = TraceContext.startSpan("addConditionalFormat", SpanType.EXCEL);

    try {
      let appliedCount = 0;

      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        for (const rule of rules) {
          const range = sheet.getRange(rule.range);

          switch (rule.type) {
            case "cellValue":
              this.addCellValueCondition(range, rule);
              break;
            case "dataBar":
              this.addDataBarCondition(range, rule);
              break;
            case "colorScale":
              this.addColorScaleCondition(range, rule);
              break;
            case "iconSet":
              this.addIconSetCondition(range, rule);
              break;
            case "topBottom":
              this.addTopBottomCondition(range, rule);
              break;
            case "text":
              this.addTextCondition(range, rule);
              break;
            case "duplicate":
              this.addDuplicateCondition(range, rule);
              break;
          }

          appliedCount++;
        }

        await context.sync();
      });

      Logger.info("[AdvancedExcelFunctions] 条件格式添加完成", {
        ruleCount: rules.length,
        appliedCount,
      });
      TraceContext.endSpan();

      return {
        success: true,
        message: `已应用 ${appliedCount} 条条件格式规则`,
        appliedRules: appliedCount,
      };
    } catch (error) {
      Logger.error("[AdvancedExcelFunctions] 条件格式添加失败", error);
      TraceContext.setSpanError(error as Error);
      TraceContext.endSpan();

      return {
        success: false,
        message: `条件格式添加失败: ${(error as Error).message}`,
        appliedRules: 0,
      };
    }
  }

  /**
   * 创建图表
   */
  async createChart(
    config: ChartConfig
  ): Promise<{ success: boolean; message: string; chartName?: string }> {
    const _span = TraceContext.startSpan("createChart", SpanType.EXCEL);
    TraceContext.setSpanAttribute("chartType", config.type);

    try {
      let chartName = "";

      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const dataRange = sheet.getRange(config.dataRange);

        // 映射图表类型
        const chartTypeMap: Record<string, Excel.ChartType> = {
          column: Excel.ChartType.columnClustered,
          bar: Excel.ChartType.barClustered,
          line: Excel.ChartType.line,
          pie: Excel.ChartType.pie,
          area: Excel.ChartType.area,
          scatter: Excel.ChartType.xyscatter,
          doughnut: Excel.ChartType.doughnut,
          combo: Excel.ChartType.columnClustered, // combo 需要额外处理
        };

        const excelChartType = chartTypeMap[config.type] || Excel.ChartType.columnClustered;

        // 创建图表
        const chart = sheet.charts.add(excelChartType, dataRange, Excel.ChartSeriesBy.auto);

        // 设置标题
        if (config.title) {
          chart.title.text = config.title;
          chart.title.visible = true;
        }

        // 设置图例
        if (config.legend) {
          chart.legend.visible = config.legend.position !== "none";
          if (config.legend.position !== "none") {
            const positionMap: Record<string, Excel.ChartLegendPosition> = {
              top: Excel.ChartLegendPosition.top,
              bottom: Excel.ChartLegendPosition.bottom,
              left: Excel.ChartLegendPosition.left,
              right: Excel.ChartLegendPosition.right,
            };
            chart.legend.position = positionMap[config.legend.position];
          }
        }

        // 加载图表名称
        chart.load("name");
        await context.sync();

        chartName = chart.name;
      });

      Logger.info("[AdvancedExcelFunctions] 图表创建完成", { chartName, type: config.type });
      TraceContext.setSpanAttribute("chartName", chartName);
      TraceContext.endSpan();

      return {
        success: true,
        message: `已创建${config.type}图表: ${chartName}`,
        chartName,
      };
    } catch (error) {
      Logger.error("[AdvancedExcelFunctions] 图表创建失败", error);
      TraceContext.setSpanError(error as Error);
      TraceContext.endSpan();

      return {
        success: false,
        message: `图表创建失败: ${(error as Error).message}`,
      };
    }
  }

  /**
   * 添加数据验证
   */
  async addDataValidation(
    rules: DataValidationRule[]
  ): Promise<{ success: boolean; message: string }> {
    const _span = TraceContext.startSpan("addDataValidation", SpanType.EXCEL);

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        for (const rule of rules) {
          const range = sheet.getRange(rule.range);
          // eslint-disable-next-line office-addins/load-object-before-read, office-addins/call-sync-before-read
          const validation = range.dataValidation;

          // 清除现有验证
          validation.clear();

          // 设置验证规则
          switch (rule.type) {
            case "whole":
            case "decimal":
              validation.rule = {
                wholeNumber:
                  rule.type === "whole"
                    ? {
                        formula1: rule.values?.[0] as number,
                        formula2: rule.values?.[1] as number,
                        operator: this.mapValidationOperator(rule.operator),
                      }
                    : undefined,
                decimal:
                  rule.type === "decimal"
                    ? {
                        formula1: rule.values?.[0] as number,
                        formula2: rule.values?.[1] as number,
                        operator: this.mapValidationOperator(rule.operator),
                      }
                    : undefined,
              };
              break;

            case "list":
              validation.rule = {
                list: {
                  inCellDropDown: true,
                  source: rule.listItems?.join(",") || "",
                },
              };
              break;

            case "date":
              validation.rule = {
                date: {
                  formula1: rule.values?.[0] as Date,
                  formula2: rule.values?.[1] as Date,
                  operator: this.mapValidationOperator(rule.operator),
                },
              };
              break;

            case "textLength":
              validation.rule = {
                textLength: {
                  formula1: rule.values?.[0] as number,
                  formula2: rule.values?.[1] as number,
                  operator: this.mapValidationOperator(rule.operator),
                },
              };
              break;

            case "custom":
              if (rule.formula) {
                validation.rule = {
                  custom: {
                    formula: rule.formula,
                  },
                };
              }
              break;
          }

          // 设置错误消息
          if (rule.errorMessage) {
            validation.errorAlert = {
              title: rule.errorMessage.title,
              message: rule.errorMessage.message,
              style: this.mapAlertStyle(rule.errorMessage.style),
              showAlert: true,
            };
          }

          // 设置输入消息
          if (rule.inputMessage) {
            validation.prompt = {
              title: rule.inputMessage.title,
              message: rule.inputMessage.message,
              showPrompt: true,
            };
          }
        }

        await context.sync();
      });

      Logger.info("[AdvancedExcelFunctions] 数据验证添加完成", { ruleCount: rules.length });
      TraceContext.endSpan();

      return {
        success: true,
        message: `已添加 ${rules.length} 条数据验证规则`,
      };
    } catch (error) {
      Logger.error("[AdvancedExcelFunctions] 数据验证添加失败", error);
      TraceContext.setSpanError(error as Error);
      TraceContext.endSpan();

      return {
        success: false,
        message: `数据验证添加失败: ${(error as Error).message}`,
      };
    }
  }

  /**
   * 创建数据透视表
   */
  async createPivotTable(
    config: PivotTableConfig
  ): Promise<{ success: boolean; message: string; pivotTableName?: string }> {
    const _span = TraceContext.startSpan("createPivotTable", SpanType.EXCEL);

    try {
      // 检查 API 支持
      const compatibility = SecurityManager.checkCompatibility();
      if (!compatibility.capabilities["pivot_tables"]) {
        throw new Error("当前 Excel 版本不支持数据透视表 API");
      }

      let pivotTableName = "";

      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const sourceRange = sheet.getRange(config.sourceRange);

        // 创建新工作表用于数据透视表
        const pivotSheet = context.workbook.worksheets.add(config.name || "透视表");
        const pivotLocation = pivotSheet.getRange(config.targetCell || "A1");

        // 创建数据透视表
        const pivotTable = pivotSheet.pivotTables.add(
          config.name || `PivotTable_${Date.now()}`,
          sourceRange,
          pivotLocation
        );

        // 添加行字段
        for (const rowField of config.rows) {
          pivotTable.hierarchies.add(
            pivotTable.hierarchies.getItem(rowField),
            Excel.PivotLayoutType.unknown // 使用正确的类型
          );
        }

        // 加载名称
        pivotTable.load("name");
        await context.sync();

        pivotTableName = pivotTable.name;
      });

      Logger.info("[AdvancedExcelFunctions] 数据透视表创建完成", { pivotTableName });
      TraceContext.endSpan();

      return {
        success: true,
        message: `已创建数据透视表: ${pivotTableName}`,
        pivotTableName,
      };
    } catch (error) {
      Logger.error("[AdvancedExcelFunctions] 数据透视表创建失败", error);
      TraceContext.setSpanError(error as Error);
      TraceContext.endSpan();

      return {
        success: false,
        message: `数据透视表创建失败: ${(error as Error).message}`,
      };
    }
  }

  /**
   * 获取可用样式预设
   */
  getStylePresets(): TableStylePreset[] {
    return [...TABLE_STYLE_PRESETS];
  }

  /**
   * 智能推荐样式
   */
  async recommendStyle(range: string): Promise<{
    recommended: string;
    alternatives: string[];
    reason: string;
  }> {
    const _span = TraceContext.startSpan("recommendStyle", SpanType.EXCEL);

    try {
      let dataCharacteristics = {
        hasNumbers: false,
        hasPercentages: false,
        hasCurrency: false,
        rowCount: 0,
        colCount: 0,
      };

      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const targetRange = sheet.getRange(range);
        targetRange.load(["values", "rowCount", "columnCount", "numberFormat"]);
        await context.sync();

        dataCharacteristics.rowCount = targetRange.rowCount;
        dataCharacteristics.colCount = targetRange.columnCount;

        const values = targetRange.values as unknown[][];
        for (const row of values) {
          for (const cell of row) {
            if (typeof cell === "number") {
              dataCharacteristics.hasNumbers = true;
            }
            if (typeof cell === "string") {
              if (cell.includes("%")) dataCharacteristics.hasPercentages = true;
              if (cell.includes("$") || cell.includes("¥")) dataCharacteristics.hasCurrency = true;
            }
          }
        }
      });

      // 基于数据特征推荐样式
      let recommended = "professional-blue";
      let reason = "默认商务风格，适合大多数场景";
      const alternatives: string[] = [];

      if (dataCharacteristics.hasCurrency) {
        recommended = "elegant-gray";
        reason = "灰色优雅风格适合财务数据展示";
        alternatives.push("professional-blue", "minimal");
      } else if (dataCharacteristics.rowCount > 50) {
        recommended = "minimal";
        reason = "大数据量表格使用极简风格更清晰";
        alternatives.push("elegant-gray", "professional-blue");
      } else if (dataCharacteristics.hasPercentages) {
        recommended = "modern-green";
        reason = "绿色风格适合展示增长和绩效数据";
        alternatives.push("vibrant-orange", "professional-blue");
      } else {
        alternatives.push("modern-green", "elegant-gray", "vibrant-orange");
      }

      TraceContext.endSpan();

      return {
        recommended,
        alternatives,
        reason,
      };
    } catch (error) {
      Logger.error("[AdvancedExcelFunctions] 样式推荐失败", error);
      TraceContext.setSpanError(error as Error);
      TraceContext.endSpan();

      return {
        recommended: "professional-blue",
        alternatives: ["modern-green", "elegant-gray"],
        reason: "使用默认推荐",
      };
    }
  }

  // ============ 私有辅助方法 ============

  private applyCellStyle(range: Excel.Range, style: CellStyle): void {
    if (style.fill?.color) {
      range.format.fill.color = style.fill.color;
    }

    if (style.font) {
      if (style.font.bold !== undefined) range.format.font.bold = style.font.bold;
      if (style.font.italic !== undefined) range.format.font.italic = style.font.italic;
      if (style.font.size !== undefined) range.format.font.size = style.font.size;
      if (style.font.color !== undefined) range.format.font.color = style.font.color;
      if (style.font.name !== undefined) range.format.font.name = style.font.name;
      if (style.font.underline !== undefined) {
        range.format.font.underline = style.font.underline
          ? Excel.RangeUnderlineStyle.single
          : Excel.RangeUnderlineStyle.none;
      }
    }

    if (style.alignment) {
      if (style.alignment.horizontal) {
        range.format.horizontalAlignment = style.alignment.horizontal as Excel.HorizontalAlignment;
      }
      if (style.alignment.vertical) {
        range.format.verticalAlignment = style.alignment.vertical as Excel.VerticalAlignment;
      }
      if (style.alignment.wrapText !== undefined) {
        range.format.wrapText = style.alignment.wrapText;
      }
    }

    if (style.numberFormat) {
      range.numberFormat = [[style.numberFormat]];
    }
  }

  private applyBorderStyle(range: Excel.Range, borderStyle: BorderStyle): void {
    const borderColor = borderStyle.color || "#000000";
    const _weight = this.mapBorderWeight(borderStyle.style);

    const sides = borderStyle.sides || ["top", "bottom", "left", "right"];

    for (const side of sides) {
      let borderIndex: Excel.BorderIndex;

      switch (side) {
        case "top":
          borderIndex = Excel.BorderIndex.edgeTop;
          break;
        case "bottom":
          borderIndex = Excel.BorderIndex.edgeBottom;
          break;
        case "left":
          borderIndex = Excel.BorderIndex.edgeLeft;
          break;
        case "right":
          borderIndex = Excel.BorderIndex.edgeRight;
          break;
        case "inside":
          // 内部边框需要单独处理
          range.format.borders.getItem(Excel.BorderIndex.insideHorizontal).style =
            this.mapBorderStyle(borderStyle.style);
          range.format.borders.getItem(Excel.BorderIndex.insideHorizontal).color = borderColor;
          range.format.borders.getItem(Excel.BorderIndex.insideVertical).style =
            this.mapBorderStyle(borderStyle.style);
          range.format.borders.getItem(Excel.BorderIndex.insideVertical).color = borderColor;
          continue;
        case "outside":
          // 外部边框
          for (const edge of [
            Excel.BorderIndex.edgeTop,
            Excel.BorderIndex.edgeBottom,
            Excel.BorderIndex.edgeLeft,
            Excel.BorderIndex.edgeRight,
          ]) {
            range.format.borders.getItem(edge).style = this.mapBorderStyle(borderStyle.style);
            range.format.borders.getItem(edge).color = borderColor;
          }
          continue;
        default:
          continue;
      }

      range.format.borders.getItem(borderIndex).style = this.mapBorderStyle(borderStyle.style);
      range.format.borders.getItem(borderIndex).color = borderColor;
    }
  }

  private mapBorderStyle(style: string): Excel.BorderLineStyle {
    const map: Record<string, Excel.BorderLineStyle> = {
      none: Excel.BorderLineStyle.none,
      thin: Excel.BorderLineStyle.thin,
      medium: Excel.BorderLineStyle.medium,
      thick: Excel.BorderLineStyle.thick,
      double: Excel.BorderLineStyle.double,
      dashed: Excel.BorderLineStyle.dash,
    };
    return map[style] || Excel.BorderLineStyle.thin;
  }

  private mapBorderWeight(style: string): Excel.BorderWeight {
    const map: Record<string, Excel.BorderWeight> = {
      thin: Excel.BorderWeight.thin,
      medium: Excel.BorderWeight.medium,
      thick: Excel.BorderWeight.thick,
    };
    return map[style] || Excel.BorderWeight.thin;
  }

  private addCellValueCondition(range: Excel.Range, rule: ConditionalFormatRule): void {
    const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);

    if (rule.format?.fill?.color) {
      cf.cellValue.format.fill.color = rule.format.fill.color;
    }
    if (rule.format?.font?.color) {
      cf.cellValue.format.font.color = rule.format.font.color;
    }

    cf.cellValue.rule = {
      formula1: String(rule.values?.[0] || ""),
      formula2: String(rule.values?.[1] || ""),
      operator: this.mapCfOperator(rule.operator),
    };
  }

  private addDataBarCondition(range: Excel.Range, rule: ConditionalFormatRule): void {
    const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);

    if (rule.dataBarColor) {
      cf.dataBar.barDirection = Excel.ConditionalDataBarDirection.leftToRight;
      cf.dataBar.positiveFormat.fillColor = rule.dataBarColor;
    }
  }

  private addColorScaleCondition(range: Excel.Range, rule: ConditionalFormatRule): void {
    const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);

    if (rule.colorScaleColors && rule.colorScaleColors.length >= 2) {
      cf.colorScale.criteria.minimum.color = rule.colorScaleColors[0];
      cf.colorScale.criteria.maximum.color =
        rule.colorScaleColors[rule.colorScaleColors.length - 1];

      if (rule.colorScaleColors.length >= 3) {
        cf.colorScale.criteria.midpoint = {
          formula: null,
          type: Excel.ConditionalFormatColorCriterionType.percentile,
          color: rule.colorScaleColors[1],
        };
      }
    }
  }

  private addIconSetCondition(range: Excel.Range, rule: ConditionalFormatRule): void {
    const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);

    const styleMap: Record<string, Excel.IconSet> = {
      threeArrows: Excel.IconSet.threeArrows,
      threeFlags: Excel.IconSet.threeFlags,
      threeTrafficLights: Excel.IconSet.threeTrafficLights1,
      fourArrows: Excel.IconSet.fourArrows,
      fiveArrows: Excel.IconSet.fiveArrows,
    };

    cf.iconSet.style = styleMap[rule.iconSetType || "threeArrows"] || Excel.IconSet.threeArrows;
  }

  private addTopBottomCondition(range: Excel.Range, rule: ConditionalFormatRule): void {
    const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.topBottom);

    cf.topBottom.rule = {
      rank: rule.topBottomCount || 10,
      type:
        rule.topBottomType === "bottom"
          ? Excel.ConditionalTopBottomCriterionType.bottomItems
          : Excel.ConditionalTopBottomCriterionType.topItems,
    };

    if (rule.format?.fill?.color) {
      cf.topBottom.format.fill.color = rule.format.fill.color;
    }
  }

  private addTextCondition(range: Excel.Range, rule: ConditionalFormatRule): void {
    const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.containsText);

    cf.textComparison.rule = {
      text: rule.text || "",
      operator: Excel.ConditionalTextOperator.contains,
    };

    if (rule.format?.fill?.color) {
      cf.textComparison.format.fill.color = rule.format.fill.color;
    }
  }

  private addDuplicateCondition(range: Excel.Range, rule: ConditionalFormatRule): void {
    const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.presetCriteria);

    cf.preset.rule = {
      criterion: Excel.ConditionalFormatPresetCriterion.duplicateValues,
    };

    if (rule.format?.fill?.color) {
      cf.preset.format.fill.color = rule.format.fill.color;
    }
  }

  private mapCfOperator(operator?: string): Excel.ConditionalCellValueOperator {
    const map: Record<string, Excel.ConditionalCellValueOperator> = {
      greaterThan: Excel.ConditionalCellValueOperator.greaterThan,
      lessThan: Excel.ConditionalCellValueOperator.lessThan,
      between: Excel.ConditionalCellValueOperator.between,
      equalTo: Excel.ConditionalCellValueOperator.equalTo,
      notEqualTo: Excel.ConditionalCellValueOperator.notEqualTo,
    };
    return map[operator || "greaterThan"] || Excel.ConditionalCellValueOperator.greaterThan;
  }

  private mapValidationOperator(operator?: string): Excel.DataValidationOperator {
    const map: Record<string, Excel.DataValidationOperator> = {
      between: Excel.DataValidationOperator.between,
      notBetween: Excel.DataValidationOperator.notBetween,
      equalTo: Excel.DataValidationOperator.equalTo,
      notEqualTo: Excel.DataValidationOperator.notEqualTo,
      greaterThan: Excel.DataValidationOperator.greaterThan,
      lessThan: Excel.DataValidationOperator.lessThan,
      greaterThanOrEqualTo: Excel.DataValidationOperator.greaterThanOrEqualTo,
      lessThanOrEqualTo: Excel.DataValidationOperator.lessThanOrEqualTo,
    };
    return map[operator || "between"] || Excel.DataValidationOperator.between;
  }

  private mapAlertStyle(style?: string): Excel.DataValidationAlertStyle {
    const map: Record<string, Excel.DataValidationAlertStyle> = {
      stop: Excel.DataValidationAlertStyle.stop,
      warning: Excel.DataValidationAlertStyle.warning,
      information: Excel.DataValidationAlertStyle.information,
    };
    return map[style || "stop"] || Excel.DataValidationAlertStyle.stop;
  }
}

// ============ 单例导出 ============

export const AdvancedExcelFunctions = new AdvancedExcelFunctionsImpl();

// 便捷方法导出
export const advanced = {
  beautifyTable: (
    range: string,
    styleId?: string,
    options?: Parameters<typeof AdvancedExcelFunctions.beautifyTable>[2]
  ) => AdvancedExcelFunctions.beautifyTable(range, styleId, options),
  addConditionalFormat: (rules: ConditionalFormatRule[]) =>
    AdvancedExcelFunctions.addConditionalFormat(rules),
  createChart: (config: ChartConfig) => AdvancedExcelFunctions.createChart(config),
  addDataValidation: (rules: DataValidationRule[]) =>
    AdvancedExcelFunctions.addDataValidation(rules),
  createPivotTable: (config: PivotTableConfig) => AdvancedExcelFunctions.createPivotTable(config),
  getStylePresets: () => AdvancedExcelFunctions.getStylePresets(),
  recommendStyle: (range: string) => AdvancedExcelFunctions.recommendStyle(range),
};
