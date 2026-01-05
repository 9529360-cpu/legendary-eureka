/**
 * 数据分析工具函数
 * v2.9.12: 从 App.tsx 提取的纯函数，无副作用
 */

import type { SelectionResult, DataSummary, ProactiveSuggestion } from "../types";
import { uid } from "./excel.utils";

/**
 * 解析公式中的单元格引用
 */
export function parseFormulaReferences(formula: string): string[] {
  if (!formula || !formula.startsWith("=")) return [];

  const references: string[] = [];
  // 匹配单元格引用: A1, $A$1, Sheet1!A1, 'Sheet Name'!A1:B10
  const cellRefRegex = /(?:'[^']+'!)?(?:\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)/gi;
  let match;
  while ((match = cellRefRegex.exec(formula)) !== null) {
    references.push(match[0]);
  }
  // 匹配命名范围 (简化：排除常用函数名)
  const namedRangeRegex = /\b([A-Za-z_][A-Za-z0-9_]*)\b/g;
  const commonFunctions = new Set([
    "SUM",
    "AVERAGE",
    "COUNT",
    "MAX",
    "MIN",
    "IF",
    "VLOOKUP",
    "HLOOKUP",
    "INDEX",
    "MATCH",
    "SUMIF",
    "COUNTIF",
    "LEFT",
    "RIGHT",
    "MID",
    "LEN",
    "TRIM",
    "UPPER",
    "LOWER",
    "CONCATENATE",
    "TEXT",
    "VALUE",
    "DATE",
    "YEAR",
    "MONTH",
    "DAY",
    "NOW",
    "TODAY",
    "ROUND",
    "ABS",
    "SQRT",
  ]);
  while ((match = namedRangeRegex.exec(formula)) !== null) {
    if (!commonFunctions.has(match[1].toUpperCase())) {
      // 可能是命名范围 - 暂不添加，避免误判
    }
  }
  return references;
}

/**
 * 分析公式复杂度
 */
export function analyzeFormulaComplexity(formula: string): {
  level: "simple" | "medium" | "complex";
  score: number;
  functions: string[];
} {
  if (!formula || !formula.startsWith("=")) {
    return { level: "simple", score: 0, functions: [] };
  }

  // 提取函数名
  const functionRegex = /([A-Z]+)\s*\(/gi;
  const functions: string[] = [];
  let match;
  while ((match = functionRegex.exec(formula)) !== null) {
    functions.push(match[1].toUpperCase());
  }

  // 计算嵌套深度
  let maxDepth = 0;
  let currentDepth = 0;
  for (const char of formula) {
    if (char === "(") {
      currentDepth++;
      maxDepth = Math.max(maxDepth, currentDepth);
    } else if (char === ")") {
      currentDepth--;
    }
  }

  // 计算复杂度分数
  const score = functions.length * 10 + maxDepth * 15 + (formula.length > 100 ? 20 : 0);

  let level: "simple" | "medium" | "complex" = "simple";
  if (score >= 50) level = "complex";
  else if (score >= 20) level = "medium";

  return { level, score, functions };
}

/**
 * 生成数据摘要
 */
export function generateDataSummary(selection: SelectionResult): DataSummary {
  const values = selection.values;
  let numericCount = 0;
  let textCount = 0;
  let dateCount = 0;
  let emptyCount = 0;
  const columnTypes: string[] = [];

  // 分析每列的数据类型
  for (let col = 0; col < selection.columnCount; col++) {
    let colNumeric = 0;
    let _colText = 0;
    let colDate = 0;
    let _colEmpty = 0;

    for (let row = 1; row < values.length; row++) {
      // 跳过表头
      const cell = values[row]?.[col];
      if (cell === null || cell === undefined || cell === "") {
        _colEmpty++;
        emptyCount++;
      } else if (typeof cell === "number") {
        colNumeric++;
      } else if (typeof cell === "string") {
        // 检查是否是日期格式
        if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(cell)) {
          colDate++;
        } else {
          _colText++;
        }
      }
    }

    // 确定列类型
    const total = values.length - 1;
    if (colNumeric / total > 0.5) {
      columnTypes.push("numeric");
      numericCount++;
    } else if (colDate / total > 0.3) {
      columnTypes.push("date");
      dateCount++;
    } else {
      columnTypes.push("text");
      textCount++;
    }
  }

  // 检查是否有表头
  const firstRow = values[0];
  const hasHeaders =
    firstRow?.every((cell) => typeof cell === "string" && cell.length > 0) ?? false;

  // 计算数据质量
  const totalCells = selection.rowCount * selection.columnCount;
  const qualityScore = Math.round((1 - emptyCount / totalCells) * 100);

  return {
    rowCount: selection.rowCount,
    columnCount: selection.columnCount,
    dataTypes: columnTypes,
    hasHeaders,
    numericColumns: numericCount,
    textColumns: textCount,
    dateColumns: dateCount,
    emptyCount,
    qualityScore,
  };
}

/**
 * 生成主动建议
 * 注意：onSend 回调由调用方传入，保持纯函数特性
 */
export function generateProactiveSuggestions(
  selection: SelectionResult,
  summary: DataSummary,
  onSend: (text: string) => Promise<void>
): ProactiveSuggestion[] {
  const suggestions: ProactiveSuggestion[] = [];

  // 1. 如果有数值列，建议统计分析
  if (summary.numericColumns > 0) {
    suggestions.push({
      id: uid(),
      icon: "analyze",
      title: "统计分析",
      description: `分析 ${summary.numericColumns} 个数值列的统计信息`,
      action: async () => {
        await onSend("分析这些数据的统计信息，包括总和、平均值、最大值、最小值");
      },
      confidence: 0.9,
    });
  }

  // 2. 如果有多个数值列，建议创建图表
  if (summary.numericColumns >= 1 && summary.rowCount >= 3) {
    const chartType = summary.dateColumns > 0 ? "折线图" : "柱状图";
    suggestions.push({
      id: uid(),
      icon: "chart",
      title: `创建${chartType}`,
      description: `将数据可视化为${chartType}`,
      action: async () => {
        await onSend(`为这些数据创建一个${chartType}`);
      },
      confidence: 0.85,
    });
  }

  // 3. 如果有空值，建议数据清洗
  if (summary.emptyCount > 0) {
    const emptyPercent = (
      (summary.emptyCount / (summary.rowCount * summary.columnCount)) *
      100
    ).toFixed(1);
    suggestions.push({
      id: uid(),
      icon: "clean",
      title: "数据清洗",
      description: `发现 ${summary.emptyCount} 个空值 (${emptyPercent}%)`,
      action: async () => {
        await onSend("帮我处理这些数据中的空值");
      },
      confidence: 0.8,
    });
  }

  // 4. 如果数据量较大，建议汇总/求和
  if (summary.rowCount > 5 && summary.numericColumns > 0) {
    suggestions.push({
      id: uid(),
      icon: "formula",
      title: "添加汇总",
      description: "为数值列添加求和/平均值公式",
      action: async () => {
        await onSend("在最下方添加求和公式");
      },
      confidence: 0.75,
    });
  }

  // 5. 如果数据质量不高，建议格式化
  if (summary.qualityScore < 90 || summary.rowCount > 10) {
    suggestions.push({
      id: uid(),
      icon: "format",
      title: "格式优化",
      description: "优化数据的格式和样式",
      action: async () => {
        await onSend("帮我优化数据的格式，添加边框和颜色");
      },
      confidence: 0.7,
    });
  }

  // 按置信度排序，最多返回4个建议
  return suggestions.sort((a, b) => b.confidence - a.confidence).slice(0, 4);
}
