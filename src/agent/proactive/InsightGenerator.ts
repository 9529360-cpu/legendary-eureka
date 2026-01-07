/**
 * InsightGenerator - 洞察生成器
 *
 * 根据工作表分析结果，生成人类可读的洞察和建议
 * 像一个有经验的数据分析师一样思考和表达
 *
 * @module agent/proactive/InsightGenerator
 */

import {
  WorksheetAnalysis,
  AnalysisIssue,
  ColumnAnalysis,
  TableStructure,
} from "./WorksheetAnalyzer";

// ========== 类型定义 ==========

/**
 * 洞察类型
 */
export type InsightType =
  | "structure"       // 结构洞察
  | "quality"         // 质量洞察
  | "pattern"         // 模式洞察
  | "anomaly"         // 异常洞察
  | "opportunity";    // 优化机会

/**
 * 单条洞察
 */
export interface Insight {
  type: InsightType;
  title: string;
  description: string;
  confidence: number; // 0-1
  importance: "low" | "medium" | "high";
  relatedColumns?: string[];
  relatedIssues?: string[];
}

/**
 * 建议操作
 */
export interface Suggestion {
  id: string;
  title: string;
  description: string;
  category: "cleanup" | "format" | "structure" | "enhance";
  priority: number; // 1-10
  estimatedImpact: "low" | "medium" | "high";
  autoExecutable: boolean;
  actions: SuggestedAction[];
}

/**
 * 具体操作
 */
export interface SuggestedAction {
  type: string;
  description: string;
  target?: string;
  parameters?: Record<string, unknown>;
}

/**
 * 完整洞察报告
 */
export interface InsightReport {
  // 简短总结（一句话）
  summary: string;

  // 详细描述（像人说话一样）
  narrativeDescription: string;

  // 洞察列表
  insights: Insight[];

  // 建议列表（已排序）
  suggestions: Suggestion[];

  // 快速操作菜单
  quickActions: QuickAction[];

  // 对话建议（如何询问用户）
  conversationPrompt: string;
}

/**
 * 快速操作
 */
export interface QuickAction {
  label: string;
  action: string;
  icon?: string;
}

// ========== 生成器类 ==========

export class InsightGenerator {
  private language: "zh" | "en" = "zh";

  constructor(options?: { language?: "zh" | "en" }) {
    if (options?.language) this.language = options.language;
  }

  /**
   * 从分析结果生成洞察报告
   */
  generate(analysis: WorksheetAnalysis): InsightReport {
    const insights = this.generateInsights(analysis);
    const suggestions = this.generateSuggestions(analysis);
    const summary = this.generateSummary(analysis);
    const narrative = this.generateNarrative(analysis, insights);
    const quickActions = this.generateQuickActions(analysis, suggestions);
    const conversationPrompt = this.generateConversationPrompt(analysis, suggestions);

    return {
      summary,
      narrativeDescription: narrative,
      insights,
      suggestions,
      quickActions,
      conversationPrompt,
    };
  }

  /**
   * 生成一句话总结
   */
  private generateSummary(analysis: WorksheetAnalysis): string {
    const { structure, totalRows, totalColumns, issues, overallQuality } = analysis;

    const structureDesc = this.getStructureDescription(structure);
    const qualityDesc = overallQuality >= 80 ? "质量不错" : overallQuality >= 60 ? "有一些问题" : "需要整理";
    const issueCount = issues.length;

    if (issueCount === 0) {
      return `这是一个${structureDesc}，共 ${totalRows} 行 ${totalColumns} 列，${qualityDesc}。`;
    }

    return `这是一个${structureDesc}，共 ${totalRows} 行 ${totalColumns} 列，发现 ${issueCount} 个可优化的地方。`;
  }

  /**
   * 生成自然语言描述（像人说话一样）
   */
  private generateNarrative(analysis: WorksheetAnalysis, insights: Insight[]): string {
    const lines: string[] = [];
    const { structure, totalRows, totalColumns, issues, columns, headerRowIndex } = analysis;

    // 开场白
    lines.push(`我刚看了这个 ${analysis.sheetName || "工作表"}。`);

    // 描述结构
    const structureDesc = this.getStructureDescription(structure);
    if (headerRowIndex !== null) {
      lines.push(`这是一个${structureDesc}，第 ${headerRowIndex + 1} 行是表头。`);
    } else {
      lines.push(`看起来像是${structureDesc}，但没有明显的表头行。`);
    }

    // 描述列
    const numericCols = columns.filter((c) => c.dataType === "number");
    const textCols = columns.filter((c) => c.dataType === "text");
    const dateCols = columns.filter((c) => c.seemsLikeDate);

    if (dateCols.length > 0) {
      lines.push(`有 ${dateCols.length} 个时间相关的列（${dateCols.map((c) => c.header || c.letter).slice(0, 3).join("、")}）。`);
    }
    if (numericCols.length > 0) {
      lines.push(`${numericCols.length} 个数值列，适合做计算和汇总。`);
    }

    // 描述问题
    const highPriorityIssues = issues.filter((i) => i.severity === "high");
    const mediumPriorityIssues = issues.filter((i) => i.severity === "medium");

    if (highPriorityIssues.length > 0) {
      lines.push("");
      lines.push("⚠️ 发现一些需要注意的问题：");
      for (const issue of highPriorityIssues.slice(0, 3)) {
        lines.push(`  • ${issue.location}：${issue.description}`);
      }
    }

    if (mediumPriorityIssues.length > 0 && highPriorityIssues.length < 2) {
      lines.push("");
      lines.push("有几个可以优化的地方：");
      for (const issue of mediumPriorityIssues.slice(0, 2)) {
        lines.push(`  • ${issue.location}：${issue.description}`);
      }
    }

    return lines.join("\n");
  }

  /**
   * 生成洞察列表
   */
  private generateInsights(analysis: WorksheetAnalysis): Insight[] {
    const insights: Insight[] = [];
    const { columns, issues, structure, overallQuality, qualityFactors } = analysis;

    // 结构洞察
    insights.push({
      type: "structure",
      title: "表格结构",
      description: this.getDetailedStructureDescription(analysis),
      confidence: 0.9,
      importance: "medium",
    });

    // 数据类型洞察
    const numericCols = columns.filter((c) => c.dataType === "number");
    const mixedCols = columns.filter((c) => c.dataType === "mixed");

    if (mixedCols.length > 0) {
      insights.push({
        type: "anomaly",
        title: "混合数据类型",
        description: `${mixedCols.length} 列包含混合数据类型，可能影响计算和分析`,
        confidence: 0.85,
        importance: "high",
        relatedColumns: mixedCols.map((c) => c.letter),
      });
    }

    // 格式问题洞察
    const textNumberIssues = issues.filter((i) => i.type === "text_formatted_numbers");
    if (textNumberIssues.length > 0) {
      insights.push({
        type: "quality",
        title: "数值格式问题",
        description: "部分数值列被存储为文本格式，会影响公式计算和透视表",
        confidence: 0.95,
        importance: "high",
        relatedIssues: textNumberIssues.map((i) => i.location),
      });
    }

    // 数据完整性洞察
    if (qualityFactors.dataCompleteness < 80) {
      const sparseColumns = columns.filter((c) => c.fillRate < 0.5);
      insights.push({
        type: "quality",
        title: "数据不完整",
        description: `${sparseColumns.length} 列数据填充率较低`,
        confidence: 0.8,
        importance: "medium",
        relatedColumns: sparseColumns.map((c) => c.letter),
      });
    }

    // 优化机会洞察
    if (structure === "free_form" || structure === "simple_list") {
      insights.push({
        type: "opportunity",
        title: "可以转换为正式表格",
        description: "将数据转换为 Excel 表格可以获得自动筛选、格式化和更好的公式支持",
        confidence: 0.75,
        importance: "medium",
      });
    }

    return insights;
  }

  /**
   * 生成建议列表
   */
  private generateSuggestions(analysis: WorksheetAnalysis): Suggestion[] {
    const suggestions: Suggestion[] = [];
    const { issues, structure, columns } = analysis;

    // 1. 转换为正式表格
    if (structure !== "standard_table" && analysis.totalRows > 1) {
      suggestions.push({
        id: "convert_to_table",
        title: "转换为正式表格",
        description: "将数据区域转换为 Excel 表格，获得自动筛选和格式化",
        category: "structure",
        priority: 8,
        estimatedImpact: "high",
        autoExecutable: true,
        actions: [
          {
            type: "excel_create_table",
            description: "创建表格",
            target: analysis.usedRange,
          },
        ],
      });
    }

    // 2. 修复文本格式的数字
    const textNumberIssues = issues.filter((i) => i.type === "text_formatted_numbers");
    if (textNumberIssues.length > 0) {
      suggestions.push({
        id: "fix_text_numbers",
        title: "修正数值格式",
        description: `将 ${textNumberIssues.length} 列的文本转换为数值`,
        category: "format",
        priority: 9,
        estimatedImpact: "high",
        autoExecutable: true,
        actions: textNumberIssues.map((issue) => ({
          type: "excel_convert_to_number",
          description: `转换 ${issue.location}`,
          target: issue.affectedRange,
        })),
      });
    }

    // 3. 统一格式
    const formatIssues = issues.filter((i) => i.type === "inconsistent_format");
    if (formatIssues.length > 0) {
      suggestions.push({
        id: "unify_format",
        title: "统一列格式",
        description: `统一 ${formatIssues.length} 列的格式`,
        category: "format",
        priority: 6,
        estimatedImpact: "medium",
        autoExecutable: true,
        actions: formatIssues.map((issue) => ({
          type: "excel_format_column",
          description: `格式化 ${issue.location}`,
          target: issue.affectedRange,
        })),
      });
    }

    // 4. 删除空行
    const emptyRowIssue = issues.find((i) => i.type === "empty_rows");
    if (emptyRowIssue) {
      suggestions.push({
        id: "remove_empty_rows",
        title: "删除空行",
        description: emptyRowIssue.description,
        category: "cleanup",
        priority: 5,
        estimatedImpact: "low",
        autoExecutable: true,
        actions: [
          {
            type: "excel_delete_empty_rows",
            description: "删除空行",
          },
        ],
      });
    }

    // 5. 添加条件格式
    const numericCols = columns.filter((c) => c.dataType === "number" && c.header);
    if (numericCols.length > 0) {
      suggestions.push({
        id: "add_conditional_format",
        title: "添加条件格式",
        description: `为数值列添加数据条或色阶，便于快速识别趋势`,
        category: "enhance",
        priority: 4,
        estimatedImpact: "medium",
        autoExecutable: true,
        actions: numericCols.slice(0, 3).map((col) => ({
          type: "excel_add_conditional_format",
          description: `为 ${col.header || col.letter} 列添加数据条`,
          target: `${col.letter}:${col.letter}`,
        })),
      });
    }

    // 按优先级排序
    return suggestions.sort((a, b) => b.priority - a.priority);
  }

  /**
   * 生成快速操作
   */
  private generateQuickActions(
    analysis: WorksheetAnalysis,
    suggestions: Suggestion[]
  ): QuickAction[] {
    const actions: QuickAction[] = [];

    // 取前3个最高优先级的建议作为快速操作
    for (const suggestion of suggestions.slice(0, 3)) {
      actions.push({
        label: suggestion.title,
        action: suggestion.id,
        icon: this.getActionIcon(suggestion.category),
      });
    }

    // 添加"全部执行"选项
    if (suggestions.filter((s) => s.autoExecutable).length > 1) {
      actions.push({
        label: "一键优化全部",
        action: "execute_all",
        icon: "✨",
      });
    }

    return actions;
  }

  /**
   * 生成对话提示（Agent 如何询问用户）
   */
  private generateConversationPrompt(
    analysis: WorksheetAnalysis,
    suggestions: Suggestion[]
  ): string {
    const lines: string[] = [];
    const autoSuggestions = suggestions.filter((s) => s.autoExecutable);

    if (autoSuggestions.length === 0) {
      return "这个表格结构还不错，有什么我可以帮你的吗？";
    }

    lines.push("我可以帮你：");
    for (const s of autoSuggestions.slice(0, 4)) {
      lines.push(`• ${s.title}`);
    }

    lines.push("");

    if (autoSuggestions.length === 1) {
      lines.push("要我帮你做这个吗？");
    } else if (autoSuggestions.length <= 3) {
      lines.push("你想全部一起做，还是先做某几个？");
    } else {
      lines.push("你是想全部一起做，还是先改某几项？");
    }

    return lines.join("\n");
  }

  // ========== 辅助方法 ==========

  private getStructureDescription(structure: TableStructure): string {
    const descriptions: Record<TableStructure, string> = {
      simple_list: "简单列表",
      standard_table: "标准表格",
      multi_header: "多行表头表格",
      pivot_style: "透视表风格的汇总表",
      matrix: "矩阵表格",
      free_form: "自由格式的数据区域",
      empty: "空表格",
    };
    return descriptions[structure] || "数据区域";
  }

  private getDetailedStructureDescription(analysis: WorksheetAnalysis): string {
    const { structure, totalRows, totalColumns, headerRowIndex, columns } = analysis;
    const base = this.getStructureDescription(structure);

    const parts = [base];
    parts.push(`${totalRows} 行 × ${totalColumns} 列`);

    if (headerRowIndex !== null) {
      const headeredCols = columns.filter((c) => c.header);
      parts.push(`表头在第 ${headerRowIndex + 1} 行（${headeredCols.length} 个有效列名）`);
    }

    return parts.join("，");
  }

  private getActionIcon(category: Suggestion["category"]): string {
    const icons: Record<Suggestion["category"], string> = {
      cleanup: "🧹",
      format: "🎨",
      structure: "📊",
      enhance: "✨",
    };
    return icons[category] || "📌";
  }
}

// ========== 导出工厂函数 ==========

export function createInsightGenerator(options?: {
  language?: "zh" | "en";
}): InsightGenerator {
  return new InsightGenerator(options);
}
