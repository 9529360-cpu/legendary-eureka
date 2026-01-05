/**
 * DataAnalyzer - 高级数据分析引擎
 *
 * 功能：
 * 1. 描述性统计分析
 * 2. 趋势分析和预测
 * 3. 异常检测
 * 4. 相关性分析
 * 5. 数据质量评估
 * 6. 智能洞察生成
 */

export interface AnalysisResult {
  summary: DataSummary;
  statistics: Record<string, ColumnStatistics>;
  insights: Insight[];
  quality: DataQuality;
  recommendations: Recommendation[];
}

export interface DataSummary {
  rowCount: number;
  columnCount: number;
  numericColumns: number;
  textColumns: number;
  dateColumns: number;
  emptyRows: number;
  duplicateRows: number;
}

export interface ColumnStatistics {
  type: "numeric" | "text" | "date" | "boolean" | "mixed";
  count: number;
  nullCount: number;
  uniqueCount: number;
  mean?: number;
  median?: number;
  mode?: any;
  std?: number;
  min?: any;
  max?: any;
  q1?: number;
  q3?: number;
  outliers?: number[];
}

export interface Insight {
  type: "trend" | "pattern" | "anomaly" | "correlation" | "recommendation";
  title: string;
  description: string;
  confidence: number;
  impact: "low" | "medium" | "high";
  data?: any;
}

export interface DataQuality {
  overall: "excellent" | "good" | "fair" | "poor";
  score: number;
  issues: QualityIssue[];
  completeness: number;
  consistency: number;
  accuracy: number;
}

export interface QualityIssue {
  type: "missing" | "duplicate" | "inconsistent" | "invalid" | "outlier";
  severity: "low" | "medium" | "high";
  description: string;
  affectedRows: number[];
  suggestion: string;
}

export interface Recommendation {
  category: "cleaning" | "transformation" | "visualization" | "analysis";
  title: string;
  description: string;
  priority: "low" | "medium" | "high";
  actionable: boolean;
  steps?: string[];
}

/**
 * DataAnalyzer类
 */
export class DataAnalyzer {
  /**
   * 执行完整的数据分析
   */
  async analyzeData(
    data: any[][],
    headers?: string[],
    options: {
      includeStatistics?: boolean;
      includeInsights?: boolean;
      includeQuality?: boolean;
      includeRecommendations?: boolean;
    } = {}
  ): Promise<AnalysisResult> {
    const {
      includeStatistics = true,
      includeInsights = true,
      includeQuality = true,
      includeRecommendations = true,
    } = options;

    // 生成默认表头
    const columnHeaders = headers || data[0]?.map((_, i) => `Column${i + 1}`) || [];

    // 1. 基础汇总
    const summary = this.generateSummary(data, columnHeaders);

    // 2. 列统计
    const statistics = includeStatistics ? this.calculateStatistics(data, columnHeaders) : {};

    // 3. 数据质量
    const quality = includeQuality
      ? this.assessDataQuality(data, statistics)
      : this.getDefaultQuality();

    // 4. 生成洞察
    const insights = includeInsights ? this.generateInsights(data, statistics, summary) : [];

    // 5. 生成建议
    const recommendations = includeRecommendations
      ? this.generateRecommendations(summary, statistics, quality, insights)
      : [];

    return {
      summary,
      statistics,
      insights,
      quality,
      recommendations,
    };
  }

  /**
   * 生成数据摘要
   */
  private generateSummary(data: any[][], headers: string[]): DataSummary {
    const rowCount = data.length;
    const columnCount = headers.length;

    // 分析列类型
    let numericColumns = 0;
    let textColumns = 0;
    let dateColumns = 0;

    for (let col = 0; col < columnCount; col++) {
      const columnType = this.detectColumnType(data.map((row) => row[col]));
      if (columnType === "numeric") numericColumns++;
      else if (columnType === "text") textColumns++;
      else if (columnType === "date") dateColumns++;
    }

    // 检测空行
    const emptyRows = data.filter((row) =>
      row.every((cell) => cell === null || cell === "" || cell === undefined)
    ).length;

    // 检测重复行
    const uniqueRows = new Set(data.map((row) => JSON.stringify(row)));
    const duplicateRows = rowCount - uniqueRows.size;

    return {
      rowCount,
      columnCount,
      numericColumns,
      textColumns,
      dateColumns,
      emptyRows,
      duplicateRows,
    };
  }

  /**
   * 计算列统计信息
   */
  private calculateStatistics(data: any[][], headers: string[]): Record<string, ColumnStatistics> {
    const statistics: Record<string, ColumnStatistics> = {};

    for (let col = 0; col < headers.length; col++) {
      const columnData = data.map((row) => row[col]);
      const columnName = headers[col];
      statistics[columnName] = this.calculateColumnStatistics(columnData);
    }

    return statistics;
  }

  /**
   * 计算单列统计信息
   */
  private calculateColumnStatistics(data: any[]): ColumnStatistics {
    const type = this.detectColumnType(data);
    const count = data.length;
    const nullCount = data.filter((v) => v === null || v === undefined || v === "").length;
    const uniqueCount = new Set(data.filter((v) => v !== null && v !== undefined && v !== "")).size;

    const stats: ColumnStatistics = {
      type,
      count,
      nullCount,
      uniqueCount,
    };

    // 如果是数值类型，计算统计指标
    if (type === "numeric") {
      const numericData = data
        .filter((v) => typeof v === "number" || !isNaN(Number(v)))
        .map((v) => Number(v))
        .filter((v) => !isNaN(v));

      if (numericData.length > 0) {
        stats.mean = this.calculateMean(numericData);
        stats.median = this.calculateMedian(numericData);
        stats.std = this.calculateStd(numericData);
        stats.min = Math.min(...numericData);
        stats.max = Math.max(...numericData);

        const quartiles = this.calculateQuartiles(numericData);
        stats.q1 = quartiles.q1;
        stats.q3 = quartiles.q3;

        // 检测离群值
        stats.outliers = this.detectOutliers(numericData, quartiles);
      }
    }

    // 计算众数
    stats.mode = this.calculateMode(data.filter((v) => v !== null && v !== undefined && v !== ""));

    return stats;
  }

  /**
   * 检测列类型
   */
  private detectColumnType(data: any[]): "numeric" | "text" | "date" | "boolean" | "mixed" {
    const nonNullData = data.filter((v) => v !== null && v !== undefined && v !== "");

    if (nonNullData.length === 0) return "text";

    let numericCount = 0;
    let dateCount = 0;
    let booleanCount = 0;
    let textCount = 0;

    for (const value of nonNullData) {
      if (typeof value === "boolean") {
        booleanCount++;
      } else if (typeof value === "number" || !isNaN(Number(value))) {
        numericCount++;
      } else if (this.isDateString(String(value))) {
        dateCount++;
      } else {
        textCount++;
      }
    }

    const total = nonNullData.length;
    const numericRatio = numericCount / total;
    const dateRatio = dateCount / total;
    const booleanRatio = booleanCount / total;

    if (numericRatio > 0.8) return "numeric";
    if (dateRatio > 0.8) return "date";
    if (booleanRatio > 0.8) return "boolean";
    if (textCount / total > 0.8) return "text";

    return "mixed";
  }

  /**
   * 判断是否为日期字符串
   */
  private isDateString(value: string): boolean {
    const date = new Date(value);
    return !isNaN(date.getTime()) && value.length > 5;
  }

  /**
   * 计算平均值
   */
  private calculateMean(data: number[]): number {
    return data.reduce((sum, v) => sum + v, 0) / data.length;
  }

  /**
   * 计算中位数
   */
  private calculateMedian(data: number[]): number {
    const sorted = [...data].sort((a, b) => a - b);
    const mid = Math.floor(sorted.length / 2);
    return sorted.length % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];
  }

  /**
   * 计算标准差
   */
  private calculateStd(data: number[]): number {
    const mean = this.calculateMean(data);
    const variance = data.reduce((sum, v) => sum + Math.pow(v - mean, 2), 0) / data.length;
    return Math.sqrt(variance);
  }

  /**
   * 计算众数
   */
  private calculateMode(data: any[]): any {
    const frequency: Record<string, number> = {};
    let maxFreq = 0;
    let mode = null;

    for (const value of data) {
      const key = String(value);
      frequency[key] = (frequency[key] || 0) + 1;
      if (frequency[key] > maxFreq) {
        maxFreq = frequency[key];
        mode = value;
      }
    }

    return mode;
  }

  /**
   * 计算四分位数
   */
  private calculateQuartiles(data: number[]): {
    q1: number;
    q2: number;
    q3: number;
  } {
    const sorted = [...data].sort((a, b) => a - b);
    const q2 = this.calculateMedian(sorted);

    const lowerHalf = sorted.slice(0, Math.floor(sorted.length / 2));
    const upperHalf = sorted.slice(Math.ceil(sorted.length / 2));

    const q1 = lowerHalf.length > 0 ? this.calculateMedian(lowerHalf) : sorted[0];
    const q3 = upperHalf.length > 0 ? this.calculateMedian(upperHalf) : sorted[sorted.length - 1];

    return { q1, q2, q3 };
  }

  /**
   * 检测离群值（IQR方法）
   */
  private detectOutliers(data: number[], quartiles: { q1: number; q3: number }): number[] {
    const iqr = quartiles.q3 - quartiles.q1;
    const lowerBound = quartiles.q1 - 1.5 * iqr;
    const upperBound = quartiles.q3 + 1.5 * iqr;

    return data.filter((v) => v < lowerBound || v > upperBound);
  }

  /**
   * 高级异常检测
   * 支持多种算法：IQR、Z-Score、Modified Z-Score
   */
  detectAnomalies(
    data: any[][],
    headers?: string[],
    options: AnomalyDetectionOptions = {}
  ): AnomalyDetectionResult {
    const {
      method = "iqr",
      threshold = method === "zscore" ? 3 : 1.5,
      includeDetails = true,
    } = options;

    const columnHeaders = headers || data[0]?.map((_, i) => `Column${i + 1}`) || [];
    const anomalies: AnomalyInfo[] = [];
    const columnAnomalySummary: Record<string, ColumnAnomalySummary> = {};

    for (let col = 0; col < columnHeaders.length; col++) {
      const columnData = data.map((row) => row[col]);
      const columnName = columnHeaders[col];

      // 只对数值列进行异常检测
      if (this.detectColumnType(columnData) !== "numeric") {
        continue;
      }

      const numericData = columnData
        .map((v, idx) => ({ value: Number(v), rowIndex: idx }))
        .filter((item) => !isNaN(item.value));

      const values = numericData.map((d) => d.value);
      if (values.length < 3) continue;

      // 根据方法检测异常
      let detectedAnomalies: Array<{ rowIndex: number; value: number; score: number }> = [];

      switch (method) {
        case "iqr":
          detectedAnomalies = this.detectAnomaliesIQR(numericData, threshold);
          break;
        case "zscore":
          detectedAnomalies = this.detectAnomaliesZScore(numericData, threshold);
          break;
        case "modified_zscore":
          detectedAnomalies = this.detectAnomaliesModifiedZScore(numericData, threshold);
          break;
      }

      // 统计列级异常信息
      const stats = this.calculateColumnStatistics(columnData);
      columnAnomalySummary[columnName] = {
        count: detectedAnomalies.length,
        percentage: (detectedAnomalies.length / numericData.length) * 100,
        method,
        threshold,
        bounds: this.getAnomalyBounds(values, method, threshold),
      };

      // 添加详细异常信息
      if (includeDetails) {
        for (const anomaly of detectedAnomalies) {
          anomalies.push({
            rowIndex: anomaly.rowIndex,
            columnIndex: col,
            columnName,
            value: anomaly.value,
            score: anomaly.score,
            severity: this.classifyAnomalySeverity(anomaly.score, method),
            suggestion: this.generateAnomalySuggestion(anomaly.value, stats),
          });
        }
      }
    }

    // 检测行级异常（多列同时异常）
    const rowAnomalyCounts = new Map<number, number>();
    for (const anomaly of anomalies) {
      rowAnomalyCounts.set(anomaly.rowIndex, (rowAnomalyCounts.get(anomaly.rowIndex) || 0) + 1);
    }

    const multiColumnAnomalyRows = Array.from(rowAnomalyCounts.entries())
      .filter(([, count]) => count > 1)
      .map(([rowIndex, count]) => ({ rowIndex, anomalyCount: count }));

    return {
      totalAnomalies: anomalies.length,
      anomalyRate: data.length > 0 ? (anomalies.length / data.length) * 100 : 0,
      anomalies,
      columnSummary: columnAnomalySummary,
      multiColumnAnomalyRows,
      overallQuality: this.assessAnomalyImpact(anomalies.length, data.length),
    };
  }

  /**
   * IQR方法检测异常
   */
  private detectAnomaliesIQR(
    data: Array<{ value: number; rowIndex: number }>,
    multiplier: number
  ): Array<{ rowIndex: number; value: number; score: number }> {
    const values = data.map((d) => d.value);
    const quartiles = this.calculateQuartiles(values);
    const iqr = quartiles.q3 - quartiles.q1;
    const lowerBound = quartiles.q1 - multiplier * iqr;
    const upperBound = quartiles.q3 + multiplier * iqr;

    return data
      .filter((d) => d.value < lowerBound || d.value > upperBound)
      .map((d) => ({
        rowIndex: d.rowIndex,
        value: d.value,
        score: d.value < lowerBound ? (lowerBound - d.value) / iqr : (d.value - upperBound) / iqr,
      }));
  }

  /**
   * Z-Score方法检测异常
   */
  private detectAnomaliesZScore(
    data: Array<{ value: number; rowIndex: number }>,
    threshold: number
  ): Array<{ rowIndex: number; value: number; score: number }> {
    const values = data.map((d) => d.value);
    const mean = this.calculateMean(values);
    const std = this.calculateStd(values);

    if (std === 0) return [];

    return data
      .map((d) => ({
        rowIndex: d.rowIndex,
        value: d.value,
        score: Math.abs((d.value - mean) / std),
      }))
      .filter((d) => d.score > threshold);
  }

  /**
   * Modified Z-Score方法检测异常（更稳健，使用MAD）
   */
  private detectAnomaliesModifiedZScore(
    data: Array<{ value: number; rowIndex: number }>,
    threshold: number
  ): Array<{ rowIndex: number; value: number; score: number }> {
    const values = data.map((d) => d.value);
    const median = this.calculateMedian(values);

    // 计算MAD (Median Absolute Deviation)
    const absoluteDeviations = values.map((v) => Math.abs(v - median));
    const mad = this.calculateMedian(absoluteDeviations);

    if (mad === 0) return [];

    const k = 0.6745; // 常数因子

    return data
      .map((d) => ({
        rowIndex: d.rowIndex,
        value: d.value,
        score: Math.abs((k * (d.value - median)) / mad),
      }))
      .filter((d) => d.score > threshold);
  }

  /**
   * 获取异常检测边界
   */
  private getAnomalyBounds(
    values: number[],
    method: string,
    threshold: number
  ): { lower: number; upper: number } {
    if (method === "iqr") {
      const quartiles = this.calculateQuartiles(values);
      const iqr = quartiles.q3 - quartiles.q1;
      return {
        lower: quartiles.q1 - threshold * iqr,
        upper: quartiles.q3 + threshold * iqr,
      };
    } else if (method === "zscore") {
      const mean = this.calculateMean(values);
      const std = this.calculateStd(values);
      return {
        lower: mean - threshold * std,
        upper: mean + threshold * std,
      };
    } else {
      const median = this.calculateMedian(values);
      const absoluteDeviations = values.map((v) => Math.abs(v - median));
      const mad = this.calculateMedian(absoluteDeviations);
      const k = 0.6745;
      return {
        lower: median - (threshold * mad) / k,
        upper: median + (threshold * mad) / k,
      };
    }
  }

  /**
   * 分类异常严重程度
   */
  private classifyAnomalySeverity(score: number, method: string): "low" | "medium" | "high" {
    const thresholds = method === "iqr" ? { low: 1, medium: 2 } : { low: 2, medium: 3 };

    if (score <= thresholds.low) return "low";
    if (score <= thresholds.medium) return "medium";
    return "high";
  }

  /**
   * 生成异常处理建议
   */
  private generateAnomalySuggestion(value: number, stats: ColumnStatistics): string {
    const mean = stats.mean || 0;
    const median = stats.median || 0;

    if (value > (stats.max || 0) * 0.9) {
      return "该值接近或超过最大值，建议核实数据录入是否正确";
    }
    if (value < (stats.min || 0) * 1.1) {
      return "该值接近或超过最小值，建议核实数据录入是否正确";
    }
    if (Math.abs(value - mean) > Math.abs(value - median)) {
      return "该值可能是异常值，建议使用中位数代替或删除";
    }
    return "建议进一步核实数据来源";
  }

  /**
   * 评估异常对数据质量的影响
   */
  private assessAnomalyImpact(
    anomalyCount: number,
    totalRows: number
  ): "excellent" | "good" | "fair" | "poor" {
    const rate = (anomalyCount / totalRows) * 100;
    if (rate <= 1) return "excellent";
    if (rate <= 5) return "good";
    if (rate <= 15) return "fair";
    return "poor";
  }

  /**
   * 评估数据质量
   */
  private assessDataQuality(
    data: any[][],
    statistics: Record<string, ColumnStatistics>
  ): DataQuality {
    const issues: QualityIssue[] = [];
    let totalScore = 100;

    // 1. 检查缺失值
    const totalCells = data.length * Object.keys(statistics).length;
    let missingCells = 0;

    for (const stats of Object.values(statistics)) {
      missingCells += stats.nullCount;

      if (stats.nullCount > data.length * 0.1) {
        issues.push({
          type: "missing",
          severity: stats.nullCount > data.length * 0.5 ? "high" : "medium",
          description: `列中有${stats.nullCount}个缺失值 (${((stats.nullCount / data.length) * 100).toFixed(1)}%)`,
          affectedRows: [],
          suggestion: "考虑填充缺失值或删除不完整的行",
        });
        totalScore -= stats.nullCount > data.length * 0.5 ? 20 : 10;
      }
    }

    const completeness = 1 - missingCells / totalCells;

    // 2. 检查离群值
    for (const [columnName, stats] of Object.entries(statistics)) {
      if (stats.outliers && stats.outliers.length > 0) {
        if (stats.outliers.length > data.length * 0.05) {
          issues.push({
            type: "outlier",
            severity: "medium",
            description: `${columnName}列发现${stats.outliers.length}个异常值`,
            affectedRows: [],
            suggestion: "检查异常值是否为错误数据",
          });
          totalScore -= 5;
        }
      }
    }

    // 3. 检查重复
    const uniqueRows = new Set(data.map((row) => JSON.stringify(row)));
    const duplicates = data.length - uniqueRows.size;
    if (duplicates > 0) {
      issues.push({
        type: "duplicate",
        severity: duplicates > data.length * 0.1 ? "high" : "low",
        description: `发现${duplicates}行重复数据`,
        affectedRows: [],
        suggestion: "移除重复行以确保数据唯一性",
      });
      totalScore -= duplicates > data.length * 0.1 ? 15 : 5;
    }

    const score = Math.max(0, Math.min(100, totalScore));
    let overall: "excellent" | "good" | "fair" | "poor";

    if (score >= 90) overall = "excellent";
    else if (score >= 75) overall = "good";
    else if (score >= 60) overall = "fair";
    else overall = "poor";

    return {
      overall,
      score,
      issues,
      completeness,
      consistency: 0.9, // 简化实现
      accuracy: 0.95, // 简化实现
    };
  }

  /**
   * 生成洞察
   */
  private generateInsights(
    data: any[][],
    statistics: Record<string, ColumnStatistics>,
    summary: DataSummary
  ): Insight[] {
    const insights: Insight[] = [];

    // 1. 数据规模洞察
    if (summary.rowCount > 10000) {
      insights.push({
        type: "recommendation",
        title: "大数据集",
        description: `数据集包含${summary.rowCount.toLocaleString()}行，建议使用数据透视表或筛选功能进行分析`,
        confidence: 1.0,
        impact: "medium",
      });
    }

    // 2. 缺失值洞察
    if (summary.emptyRows > 0) {
      insights.push({
        type: "pattern",
        title: "空行检测",
        description: `发现${summary.emptyRows}个空行，可能需要清理`,
        confidence: 1.0,
        impact: "low",
      });
    }

    // 3. 重复数据洞察
    if (summary.duplicateRows > 0) {
      insights.push({
        type: "pattern",
        title: "重复数据",
        description: `发现${summary.duplicateRows}行重复数据，建议检查并处理`,
        confidence: 1.0,
        impact: summary.duplicateRows > summary.rowCount * 0.1 ? "high" : "medium",
      });
    }

    // 4. 数值列分析
    for (const [columnName, stats] of Object.entries(statistics)) {
      if (stats.type === "numeric" && stats.mean !== undefined && stats.std !== undefined) {
        // 检测高变异性
        const cv = stats.std / stats.mean;
        if (cv > 1) {
          insights.push({
            type: "pattern",
            title: `${columnName}列变异性高`,
            description: `该列数据分布较为分散，变异系数为${cv.toFixed(2)}`,
            confidence: 0.8,
            impact: "medium",
          });
        }

        // 检测偏态
        if (stats.outliers && stats.outliers.length > 0) {
          insights.push({
            type: "anomaly",
            title: `${columnName}列包含异常值`,
            description: `检测到${stats.outliers.length}个潜在异常值`,
            confidence: 0.9,
            impact: "medium",
            data: { outliers: stats.outliers.slice(0, 5) },
          });
        }
      }
    }

    // 5. 数据类型建议
    if (summary.textColumns > summary.numericColumns && summary.numericColumns > 0) {
      insights.push({
        type: "recommendation",
        title: "可视化建议",
        description: "数据包含较多文本列，建议使用条形图或饼图进行可视化",
        confidence: 0.7,
        impact: "low",
      });
    } else if (summary.numericColumns >= 2) {
      insights.push({
        type: "recommendation",
        title: "可视化建议",
        description: "数据包含多个数值列，建议使用折线图或散点图展示趋势",
        confidence: 0.7,
        impact: "low",
      });
    }

    return insights;
  }

  /**
   * 生成改进建议
   */
  private generateRecommendations(
    summary: DataSummary,
    statistics: Record<string, ColumnStatistics>,
    quality: DataQuality,
    insights: Insight[]
  ): Recommendation[] {
    const recommendations: Recommendation[] = [];

    // 1. 数据清理建议
    if (summary.emptyRows > 0 || summary.duplicateRows > 0) {
      recommendations.push({
        category: "cleaning",
        title: "清理数据",
        description: "移除空行和重复数据以提高数据质量",
        priority: "high",
        actionable: true,
        steps: ["选择数据范围", "使用'删除重复项'功能", "删除空行"],
      });
    }

    // 2. 缺失值处理建议
    const columnsWithMissing = Object.entries(statistics).filter(
      ([_, stats]) => stats.nullCount > stats.count * 0.05
    );

    if (columnsWithMissing.length > 0) {
      recommendations.push({
        category: "cleaning",
        title: "处理缺失值",
        description: `${columnsWithMissing.length}列包含大量缺失值`,
        priority: "high",
        actionable: true,
        steps: ["识别缺失值模式", "使用平均值/中位数填充或删除行", "验证数据完整性"],
      });
    }

    // 3. 异常值处理建议
    const hasOutliers = insights.some((i) => i.type === "anomaly");
    if (hasOutliers) {
      recommendations.push({
        category: "cleaning",
        title: "检查异常值",
        description: "数据中存在异常值，可能影响分析结果",
        priority: "medium",
        actionable: true,
        steps: ["使用条件格式突出显示异常值", "验证异常值是否为错误", "决定保留或移除"],
      });
    }

    // 4. 可视化建议
    if (summary.numericColumns >= 2) {
      recommendations.push({
        category: "visualization",
        title: "创建图表",
        description: "数据适合使用图表进行可视化展示",
        priority: "medium",
        actionable: true,
        steps: ["选择数据范围", "插入推荐的图表类型", "自定义图表样式"],
      });
    }

    // 5. 性能优化建议
    if (summary.rowCount > 5000) {
      recommendations.push({
        category: "analysis",
        title: "优化性能",
        description: "大数据集建议使用表格和筛选功能提高性能",
        priority: "low",
        actionable: true,
        steps: ["将范围转换为表格", "使用切片器和筛选", "考虑使用数据透视表"],
      });
    }

    // 6. 数据验证建议
    if (quality.overall === "fair" || quality.overall === "poor") {
      recommendations.push({
        category: "cleaning",
        title: "提高数据质量",
        description: `当前数据质量评分：${quality.score}/100`,
        priority: "high",
        actionable: true,
        steps: ["审查数据质量问题", "实施数据验证规则", "建立数据清理流程"],
      });
    }

    return recommendations;
  }

  /**
   * 获取默认质量评估
   */
  private getDefaultQuality(): DataQuality {
    return {
      overall: "good",
      score: 80,
      issues: [],
      completeness: 1.0,
      consistency: 1.0,
      accuracy: 1.0,
    };
  }

  /**
   * 计算两列之间的相关性
   */
  calculateCorrelation(data1: number[], data2: number[]): number {
    if (data1.length !== data2.length || data1.length === 0) return 0;

    const mean1 = this.calculateMean(data1);
    const mean2 = this.calculateMean(data2);

    let numerator = 0;
    let sum1 = 0;
    let sum2 = 0;

    for (let i = 0; i < data1.length; i++) {
      const diff1 = data1[i] - mean1;
      const diff2 = data2[i] - mean2;
      numerator += diff1 * diff2;
      sum1 += diff1 * diff1;
      sum2 += diff2 * diff2;
    }

    const denominator = Math.sqrt(sum1 * sum2);
    return denominator === 0 ? 0 : numerator / denominator;
  }

  /**
   * 简单的趋势预测
   */
  predictTrend(data: number[], periods: number = 3): number[] {
    if (data.length < 2) return [];

    // 使用简单线性回归
    const n = data.length;
    const x = Array.from({ length: n }, (_, i) => i);
    const meanX = this.calculateMean(x);
    const meanY = this.calculateMean(data);

    let numerator = 0;
    let denominator = 0;

    for (let i = 0; i < n; i++) {
      numerator += (x[i] - meanX) * (data[i] - meanY);
      denominator += (x[i] - meanX) ** 2;
    }

    const slope = numerator / denominator;
    const intercept = meanY - slope * meanX;

    // 预测未来值
    const predictions: number[] = [];
    for (let i = 0; i < periods; i++) {
      predictions.push(slope * (n + i) + intercept);
    }

    return predictions;
  }

  /**
   * 智能图表推荐
   * 根据数据特征推荐最合适的图表类型
   */
  recommendChart(
    data: any[][],
    headers?: string[],
    purpose?: "comparison" | "composition" | "distribution" | "relationship" | "trend"
  ): ChartRecommendation {
    const columnHeaders = headers || data[0]?.map((_, i) => `Column${i + 1}`) || [];
    const statistics = this.calculateStatistics(data, columnHeaders);
    const summary = this.generateSummary(data, columnHeaders);

    // 分析数据特征
    const dataProfile = this.analyzeDataProfile(data, summary, statistics);

    // 根据数据特征和目的推荐图表
    const recommendations = this.generateChartRecommendations(dataProfile, purpose);

    return {
      primary: recommendations[0],
      alternatives: recommendations.slice(1, 4),
      dataProfile,
      explanation: this.generateChartExplanation(recommendations[0], dataProfile),
    };
  }

  /**
   * 分析数据特征
   */
  private analyzeDataProfile(
    data: any[][],
    summary: DataSummary,
    statistics: Record<string, ColumnStatistics>
  ): DataProfile {
    const columns = Object.entries(statistics);
    const numericCols = columns.filter(([, s]) => s.type === "numeric");
    const categoricalCols = columns.filter(([, s]) => s.type === "text");
    const dateCols = columns.filter(([, s]) => s.type === "date");

    // 检测时间序列特征
    const hasTimeSequence = dateCols.length > 0 || this.detectTimePattern(data);

    // 检测类别数量
    const categoryCount =
      categoricalCols.length > 0 ? Math.max(...categoricalCols.map(([, s]) => s.uniqueCount)) : 0;

    // 检测数值范围
    const numericRanges = numericCols.map(([name, s]) => ({
      name,
      range: (s.max || 0) - (s.min || 0),
      hasNegatives: (s.min || 0) < 0,
    }));

    // 检测比例性（是否适合饼图）
    const isProportion =
      numericCols.length === 1 && categoricalCols.length === 1 && summary.rowCount <= 10;

    // 检测相关性（是否适合散点图）
    let hasStrongCorrelation = false;
    if (numericCols.length >= 2) {
      const col1Data = data.map((row) => Number(row[0])).filter((v) => !isNaN(v));
      const col2Data = data.map((row) => Number(row[1])).filter((v) => !isNaN(v));
      if (col1Data.length === col2Data.length && col1Data.length > 5) {
        const correlation = Math.abs(this.calculateCorrelation(col1Data, col2Data));
        hasStrongCorrelation = correlation > 0.5;
      }
    }

    return {
      rowCount: summary.rowCount,
      numericColumnCount: numericCols.length,
      categoricalColumnCount: categoricalCols.length,
      dateColumnCount: dateCols.length,
      hasTimeSequence,
      categoryCount,
      numericRanges,
      isProportion,
      hasStrongCorrelation,
      hasOutliers: numericCols.some(([, s]) => (s.outliers?.length || 0) > 0),
    };
  }

  /**
   * 检测时间序列模式
   */
  private detectTimePattern(data: any[][]): boolean {
    if (data.length < 2) return false;

    // 检查第一列是否看起来像时间序列
    const firstColumn = data.map((row) => row[0]);

    // 检测数字递增模式（年份、月份等）
    const numericValues = firstColumn
      .filter((v) => typeof v === "number" || !isNaN(Number(v)))
      .map((v) => Number(v));

    if (numericValues.length >= 3) {
      let isIncreasing = true;
      for (let i = 1; i < numericValues.length; i++) {
        if (numericValues[i] <= numericValues[i - 1]) {
          isIncreasing = false;
          break;
        }
      }
      if (isIncreasing) return true;
    }

    // 检测日期字符串模式
    const datePatterns = [
      /^\d{4}[-/]\d{1,2}[-/]\d{1,2}$/, // YYYY-MM-DD
      /^\d{1,2}[-/]\d{1,2}[-/]\d{4}$/, // MM-DD-YYYY
      /^\d{4}年\d{1,2}月/, // 中文日期
      /^Q[1-4]\s*\d{4}$/, // Q1 2024
      /^\d{4}$/, // 年份
    ];

    const firstVal = String(firstColumn[0] || "");
    return datePatterns.some((pattern) => pattern.test(firstVal));
  }

  /**
   * 生成图表推荐列表
   */
  private generateChartRecommendations(profile: DataProfile, purpose?: string): ChartOption[] {
    const recommendations: ChartOption[] = [];

    // 根据数据特征评分
    const scores = this.calculateChartScores(profile, purpose);

    // 按分数排序
    const sortedCharts = Object.entries(scores)
      .sort(([, a], [, b]) => b - a)
      .filter(([, score]) => score > 0);

    for (const [chartType, score] of sortedCharts) {
      recommendations.push({
        type: chartType as ChartType,
        score,
        suitable: score >= 60,
        config: this.getDefaultChartConfig(chartType as ChartType, profile),
      });
    }

    return recommendations;
  }

  /**
   * 计算各种图表类型的适合度分数
   */
  private calculateChartScores(profile: DataProfile, purpose?: string): Record<ChartType, number> {
    const scores: Record<ChartType, number> = {
      column: 0,
      bar: 0,
      line: 0,
      pie: 0,
      doughnut: 0,
      scatter: 0,
      area: 0,
      radar: 0,
      combo: 0,
    };

    // 柱状图：适合类别比较
    if (profile.categoricalColumnCount >= 1 && profile.numericColumnCount >= 1) {
      scores.column = 70;
      if (profile.categoryCount <= 12) scores.column += 15;
      if (profile.rowCount <= 20) scores.column += 10;
      if (purpose === "comparison") scores.column += 20;
    }

    // 条形图：类似柱状图，但类别名称较长时更合适
    scores.bar = scores.column - 5;
    if (profile.categoryCount > 8) scores.bar += 10;

    // 折线图：适合时间序列和趋势
    if (profile.hasTimeSequence || profile.rowCount >= 5) {
      scores.line = 60;
      if (profile.hasTimeSequence) scores.line += 25;
      if (profile.rowCount >= 10) scores.line += 10;
      if (purpose === "trend") scores.line += 20;
    }

    // 饼图：适合显示比例
    if (profile.isProportion) {
      scores.pie = 80;
      if (profile.categoryCount <= 6) scores.pie += 15;
      if (purpose === "composition") scores.pie += 20;
    }
    scores.doughnut = scores.pie - 5;

    // 散点图：适合显示相关性
    if (profile.numericColumnCount >= 2) {
      scores.scatter = 50;
      if (profile.hasStrongCorrelation) scores.scatter += 30;
      if (profile.rowCount >= 10) scores.scatter += 10;
      if (purpose === "relationship") scores.scatter += 20;
    }

    // 面积图：适合累积数据
    if (profile.hasTimeSequence && profile.numericColumnCount >= 1) {
      scores.area = 55;
      if (profile.numericColumnCount >= 2) scores.area += 10;
    }

    // 雷达图：适合多维度比较
    if (profile.numericColumnCount >= 3 && profile.rowCount <= 10) {
      scores.radar = 60;
      if (profile.numericColumnCount >= 5) scores.radar += 15;
    }

    // 组合图：多数值列时
    if (profile.numericColumnCount >= 2 && profile.hasTimeSequence) {
      scores.combo = 65;
    }

    return scores;
  }

  /**
   * 获取默认图表配置
   */
  private getDefaultChartConfig(chartType: ChartType, profile: DataProfile): ChartConfig {
    const config: ChartConfig = {
      title: "",
      showLegend: profile.numericColumnCount > 1,
      showDataLabels: profile.rowCount <= 10,
      colorScheme: "default",
    };

    switch (chartType) {
      case "pie":
      case "doughnut":
        config.showLegend = true;
        config.showDataLabels = true;
        break;
      case "line":
      case "area":
        config.showMarkers = profile.rowCount <= 20;
        config.smooth = profile.rowCount >= 10;
        break;
      case "scatter":
        config.showTrendline = profile.hasStrongCorrelation;
        break;
    }

    return config;
  }

  /**
   * 生成图表推荐解释
   */
  private generateChartExplanation(chart: ChartOption, profile: DataProfile): string {
    const typeNames: Record<ChartType, string> = {
      column: "柱状图",
      bar: "条形图",
      line: "折线图",
      pie: "饼图",
      doughnut: "环形图",
      scatter: "散点图",
      area: "面积图",
      radar: "雷达图",
      combo: "组合图",
    };

    const typeName = typeNames[chart.type];
    let explanation = `推荐使用${typeName}，`;

    switch (chart.type) {
      case "column":
        explanation += `适合比较${profile.categoryCount}个类别的数值差异。`;
        break;
      case "bar":
        explanation += "适合横向比较类别，特别是类别名称较长时。";
        break;
      case "line":
        explanation += profile.hasTimeSequence
          ? "适合展示数据随时间的变化趋势。"
          : "适合展示连续数据的变化趋势。";
        break;
      case "pie":
      case "doughnut":
        explanation += "适合展示各部分占整体的比例关系。";
        break;
      case "scatter":
        explanation += profile.hasStrongCorrelation
          ? "数据显示出较强的相关性，散点图可以直观展示这种关系。"
          : "适合探索两个数值变量之间的关系。";
        break;
      case "area":
        explanation += "适合展示累积数据或多个数据系列的变化。";
        break;
      case "radar":
        explanation += "适合多维度对比分析。";
        break;
      case "combo":
        explanation += "适合同时展示不同度量单位的数据。";
        break;
    }

    return explanation;
  }
}

// 图表推荐相关类型
export type ChartType =
  | "column"
  | "bar"
  | "line"
  | "pie"
  | "doughnut"
  | "scatter"
  | "area"
  | "radar"
  | "combo";

export interface DataProfile {
  rowCount: number;
  numericColumnCount: number;
  categoricalColumnCount: number;
  dateColumnCount: number;
  hasTimeSequence: boolean;
  categoryCount: number;
  numericRanges: Array<{ name: string; range: number; hasNegatives: boolean }>;
  isProportion: boolean;
  hasStrongCorrelation: boolean;
  hasOutliers: boolean;
}

export interface ChartOption {
  type: ChartType;
  score: number;
  suitable: boolean;
  config: ChartConfig;
}

export interface ChartConfig {
  title: string;
  showLegend: boolean;
  showDataLabels: boolean;
  colorScheme: string;
  showMarkers?: boolean;
  smooth?: boolean;
  showTrendline?: boolean;
}

export interface ChartRecommendation {
  primary: ChartOption;
  alternatives: ChartOption[];
  dataProfile: DataProfile;
  explanation: string;
}

// 异常检测相关类型
export type AnomalyDetectionMethod = "iqr" | "zscore" | "modified_zscore";

export interface AnomalyDetectionOptions {
  method?: AnomalyDetectionMethod;
  threshold?: number;
  includeDetails?: boolean;
}

export interface AnomalyInfo {
  rowIndex: number;
  columnIndex: number;
  columnName: string;
  value: number;
  score: number;
  severity: "low" | "medium" | "high";
  suggestion: string;
}

export interface ColumnAnomalySummary {
  count: number;
  percentage: number;
  method: string;
  threshold: number;
  bounds: { lower: number; upper: number };
}

export interface AnomalyDetectionResult {
  totalAnomalies: number;
  anomalyRate: number;
  anomalies: AnomalyInfo[];
  columnSummary: Record<string, ColumnAnomalySummary>;
  multiColumnAnomalyRows: Array<{ rowIndex: number; anomalyCount: number }>;
  overallQuality: "excellent" | "good" | "fair" | "poor";
}
