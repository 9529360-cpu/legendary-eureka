/**
 * Excel 图表类工具
 *
 * 包含工具：
 * - createChartTool: 创建图表
 * - createChartTrendlineTool: 添加趋势线
 *
 * @packageDocumentation
 */

import { Tool } from "../../types";
import { excelRun } from "./common";

// ========== 图表类工具 ==========

export function createChartTool(): Tool {
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

export function createChartTrendlineTool(): Tool {
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
        // eslint-disable-next-line office-addins/call-sync-after-load, office-addins/call-sync-before-read
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
 * 创建所有图表类工具
 */
export function createChartTools(): Tool[] {
  return [createChartTool(), createChartTrendlineTool()];
}
