/**
 * PromptBuilder - 高级AI提示构建系统
 *
 * 功能：
 * 1. 构建结构化的AI提示，确保高质量响应
 * 2. 支持多种场景：意图分析、任务规划、数据分析、代码生成
 * 3. 上下文管理和对话历史整合
 * 4. 提供示例和约束，引导AI输出
 */

import { ConversationMessage, ToolDefinition, UserIntent, WorkbookContext } from "../types";

/**
 * 提示模板类型
 */
export type PromptTemplate =
  | "intent_analysis"
  | "task_planning"
  | "data_analysis"
  | "code_generation"
  | "error_diagnosis"
  | "suggestion_generation"
  | "formula_creation"
  | "chart_recommendation";

/**
 * 提示构建选项
 */
export interface PromptOptions {
  includeExamples?: boolean;
  includeConstraints?: boolean;
  includeContext?: boolean;
  temperature?: number;
  maxTokens?: number;
  format?: "json" | "text" | "code";
}

/**
 * PromptBuilder类
 */
export class PromptBuilder {
  private systemPrompts: Map<PromptTemplate, string>;
  private examples: Map<PromptTemplate, string[]>;

  constructor() {
    this.systemPrompts = new Map();
    this.examples = new Map();
    this.initializePrompts();
  }

  /**
   * 初始化提示模板
   */
  private initializePrompts(): void {
    // 意图分析提示
    this.systemPrompts.set(
      "intent_analysis",
      `你是一个专业的Excel操作助手，专门分析用户的自然语言指令并识别其意图。

你的任务：
1. 理解用户想要在Excel中完成什么操作
2. 提取关键参数（如范围、数据、格式要求等）
3. 评估操作的复杂度和风险等级
4. 返回结构化的意图信息

支持的操作类型：
- create_table: 创建表格或数据区域
- format_cells: 格式化单元格（颜色、字体、边框等）
- create_chart: 创建各种类型的图表
- insert_data: 插入或更新数据
- apply_filter: 应用数据筛选
- insert_formula: 插入公式或计算
- sort_data: 数据排序
- clear_range: 清除数据
- analyze_data: 数据分析和统计
- generate_summary: 生成数据摘要
- create_pivot: 创建透视表
- conditional_format: 条件格式化
- data_validation: 数据验证
- merge_cells: 合并单元格

输出格式（必须是有效的JSON）：
{
  "intent": "操作类型",
  "confidence": 0.0-1.0,
  "description": "操作描述",
  "parameters": {
    "range": "A1:B10",
    "具体参数": "值"
  },
  "complexity": "low|medium|high",
  "risks": ["潜在风险列表"],
  "suggestions": ["改进建议"]
}

注意：
- 如果用户描述模糊，提供最合理的解释
- 对于复杂操作，考虑分解为多个步骤
- 标注可能的风险和注意事项`
    );

    // 任务规划提示
    this.systemPrompts.set(
      "task_planning",
      `你是一个Excel任务规划专家，负责将用户意图转换为可执行的操作计划。

你的任务：
1. 分析用户意图和当前Excel状态
2. 制定详细的执行步骤
3. 识别步骤间的依赖关系
4. 评估每步的风险和执行时间
5. 提供优化建议

规划原则：
- 操作顺序合理，考虑依赖关系
- 每步操作清晰可执行
- 风险最小化，提供回退方案
- 性能优化，避免重复操作
- 用户体验友好，提供进度反馈

输出格式（必须是有效的JSON）：
{
  "plan_id": "唯一标识",
  "steps": [
    {
      "step_number": 1,
      "operation": "操作类型",
      "description": "详细描述",
      "parameters": {},
      "dependencies": [前置步骤编号],
      "estimated_time": 2,
      "risk_level": "low|medium|high",
      "rollback_strategy": "回退策略"
    }
  ],
  "total_estimated_time": 10,
  "overall_risk": "low|medium|high",
  "warnings": ["警告信息"],
  "optimization_suggestions": ["优化建议"]
}`
    );

    // 数据分析提示
    this.systemPrompts.set(
      "data_analysis",
      `你是一个数据分析专家，专门分析Excel数据并提供洞察。

你的任务：
1. 理解数据的结构和特征
2. 识别数据模式、趋势和异常
3. 执行统计分析
4. 提供可视化建议
5. 生成分析报告

分析能力：
- 描述性统计（均值、中位数、标准差等）
- 趋势分析和预测
- 相关性分析
- 异常值检测
- 数据质量评估
- 分组和聚合分析
- 时间序列分析

输出格式（必须是有效的JSON）：
{
  "summary": {
    "row_count": 100,
    "column_count": 5,
    "data_quality": "good|fair|poor",
    "completeness": 0.95
  },
  "statistics": {
    "column_name": {
      "mean": 0,
      "median": 0,
      "std": 0,
      "min": 0,
      "max": 0
    }
  },
  "insights": [
    {
      "type": "trend|pattern|anomaly|correlation",
      "description": "发现描述",
      "confidence": 0.8,
      "recommendation": "建议行动"
    }
  ],
  "visualizations": [
    {
      "chart_type": "line|bar|pie|scatter",
      "reason": "推荐理由",
      "data_range": "A1:B10"
    }
  ]
}`
    );

    // 公式创建提示
    this.systemPrompts.set(
      "formula_creation",
      `你是一个Excel公式专家，专门创建和优化Excel公式。

你的能力：
- 所有Excel内置函数（数学、统计、文本、日期、逻辑等）
- 数组公式和动态数组
- 嵌套函数和复杂逻辑
- 性能优化的公式
- 错误处理和边界情况

常用函数类别：
- 数学：SUM, AVERAGE, COUNT, MAX, MIN, ROUND, ABS
- 统计：STDEV, VAR, MEDIAN, MODE, PERCENTILE
- 逻辑：IF, AND, OR, NOT, IFS, SWITCH
- 查找：VLOOKUP, HLOOKUP, INDEX, MATCH, XLOOKUP
- 文本：CONCATENATE, LEFT, RIGHT, MID, FIND, SUBSTITUTE
- 日期：TODAY, NOW, DATE, YEAR, MONTH, DAY, DATEDIF
- 数组：FILTER, SORT, UNIQUE, SEQUENCE

输出格式（必须是有效的JSON）：
{
  "formula": "=SUM(A1:A10)",
  "explanation": "公式说明",
  "alternative_formulas": ["备选公式"],
  "performance_note": "性能说明",
  "examples": [
    {
      "input": "输入示例",
      "output": "输出结果"
    }
  ]
}`
    );

    // 图表推荐提示
    this.systemPrompts.set(
      "chart_recommendation",
      `你是一个数据可视化专家，专门推荐最适合的图表类型。

图表类型及适用场景：
- 柱状图(column): 比较不同类别的数值
- 条形图(bar): 横向比较，类别名称较长
- 折线图(line): 显示趋势和时间序列
- 饼图(pie): 显示部分与整体的关系
- 散点图(scatter): 显示两个变量的相关性
- 面积图(area): 显示数量随时间的变化
- 雷达图(radar): 多维度比较
- 组合图(combo): 结合多种图表类型

推荐原则：
- 根据数据类型和维度选择
- 考虑数据量和可读性
- 符合业务场景和目的
- 美观且信息传达清晰

输出格式（必须是有效的JSON）：
{
  "primary_recommendation": {
    "chart_type": "column",
    "reason": "推荐理由",
    "data_range": "A1:B10",
    "settings": {
      "title": "图表标题",
      "x_axis_title": "X轴标题",
      "y_axis_title": "Y轴标题",
      "legend_position": "right"
    }
  },
  "alternatives": [
    {
      "chart_type": "line",
      "reason": "备选理由"
    }
  ]
}`
    );

    // 错误诊断提示
    this.systemPrompts.set(
      "error_diagnosis",
      `你是一个Excel问题诊断专家，专门识别和解决Excel中的问题。

常见问题类型：
- 公式错误（#REF!, #VALUE!, #DIV/0!, #NAME?, #N/A）
- 数据质量问题（缺失值、重复值、格式不一致）
- 性能问题（计算缓慢、文件过大）
- 格式问题（单元格格式、条件格式）
- 引用问题（循环引用、断开链接）

诊断流程：
1. 识别问题类型和根本原因
2. 评估影响范围和严重程度
3. 提供解决方案和预防措施
4. 建议最佳实践

输出格式（必须是有效的JSON）：
{
  "problem_type": "formula_error|data_quality|performance|format",
  "root_cause": "根本原因",
  "affected_range": "A1:B10",
  "severity": "low|medium|high|critical",
  "solutions": [
    {
      "approach": "解决方案",
      "steps": ["步骤1", "步骤2"],
      "pros": ["优点"],
      "cons": ["缺点"]
    }
  ],
  "prevention": ["预防措施"],
  "best_practices": ["最佳实践"]
}`
    );
  }

  /**
   * 构建意图分析提示
   */
  buildIntentAnalysisPrompt(
    userInput: string,
    conversationHistory: ConversationMessage[],
    availableTools: ToolDefinition[],
    options: PromptOptions = {}
  ): { system: string; user: string; options: PromptOptions } {
    const systemPrompt = this.systemPrompts.get("intent_analysis") || "";

    // 构建上下文
    let contextSection = "";
    if (options.includeContext && conversationHistory.length > 0) {
      contextSection = "\n\n对话历史（最近5条）：\n";
      conversationHistory.slice(-5).forEach((msg) => {
        contextSection += `${msg.role}: ${msg.content}\n`;
      });
    }

    // 添加可用工具信息
    let toolsSection = "";
    if (availableTools.length > 0) {
      toolsSection = "\n\n可用的Excel工具：\n";
      availableTools.forEach((tool) => {
        toolsSection += `- ${tool.name}: ${tool.description}\n`;
      });
    }

    // 构建示例
    let examplesSection = "";
    if (options.includeExamples) {
      examplesSection = `\n\n示例：
输入: "把A1到A10的数据求和"
输出: {"intent":"insert_formula","confidence":0.95,"description":"在指定单元格插入求和公式","parameters":{"range":"A1:A10","formula":"SUM"},"complexity":"low","risks":[],"suggestions":["可以使用自动求和功能"]}

输入: "创建一个销售数据的柱状图"
输出: {"intent":"create_chart","confidence":0.9,"description":"创建柱状图展示销售数据","parameters":{"chartType":"column"},"complexity":"medium","risks":["需要确保数据范围正确"],"suggestions":["选择包含标题的数据区域"]}`;
    }

    const fullSystemPrompt = systemPrompt + contextSection + toolsSection + examplesSection;

    return {
      system: fullSystemPrompt,
      user: `用户请求：${userInput}\n\n请分析用户意图并返回JSON格式的结果。`,
      options: {
        temperature: options.temperature || 0.2,
        maxTokens: options.maxTokens || 800,
        format: "json",
      },
    };
  }

  /**
   * 构建任务规划提示
   */
  buildTaskPlanningPrompt(
    intent: UserIntent,
    workbookContext: WorkbookContext,
    options: PromptOptions = {}
  ): { system: string; user: string; options: PromptOptions } {
    const systemPrompt = this.systemPrompts.get("task_planning") || "";

    // 添加工作簿上下文
    const contextSection = `\n\n当前Excel状态：
- 活动工作表：${workbookContext.activeSheet}
- 选中范围：${workbookContext.selectedRange}
- 工作簿名称：${workbookContext.workbookName}
- 所有工作表：${workbookContext.sheetNames.join(", ")}`;

    const userPrompt = `用户意图：
类型：${intent.type}
描述：${intent.rawInput}
参数：${JSON.stringify(intent.parameters, null, 2)}
置信度：${intent.confidence}

请制定详细的执行计划，返回JSON格式。`;

    return {
      system: systemPrompt + contextSection,
      user: userPrompt,
      options: {
        temperature: options.temperature || 0.3,
        maxTokens: options.maxTokens || 1200,
        format: "json",
      },
    };
  }

  /**
   * 构建数据分析提示
   */
  buildDataAnalysisPrompt(
    data: any[][],
    headers: string[],
    analysisType: "summary" | "trend" | "correlation" | "anomaly" = "summary",
    options: PromptOptions = {}
  ): { system: string; user: string; options: PromptOptions } {
    const systemPrompt = this.systemPrompts.get("data_analysis") || "";

    // 数据预览
    const dataPreview = data
      .slice(0, 5)
      .map((row) => row.join("\t"))
      .join("\n");

    const userPrompt = `请分析以下Excel数据：

列名：${headers.join(", ")}
数据行数：${data.length}
数据列数：${headers.length}

数据预览（前5行）：
${dataPreview}

分析类型：${analysisType}
${analysisType === "summary" ? "请提供数据的整体统计摘要" : ""}
${analysisType === "trend" ? "请识别数据的趋势和模式" : ""}
${analysisType === "correlation" ? "请分析变量之间的相关性" : ""}
${analysisType === "anomaly" ? "请检测异常值和数据质量问题" : ""}

请返回JSON格式的分析结果。`;

    return {
      system: systemPrompt,
      user: userPrompt,
      options: {
        temperature: options.temperature || 0.4,
        maxTokens: options.maxTokens || 1500,
        format: "json",
      },
    };
  }

  /**
   * 构建公式创建提示
   */
  buildFormulaCreationPrompt(
    requirement: string,
    dataContext?: {
      range: string;
      sampleData: any[][];
      dataType: "numeric" | "text" | "date" | "mixed";
    },
    options: PromptOptions = {}
  ): { system: string; user: string; options: PromptOptions } {
    const systemPrompt = this.systemPrompts.get("formula_creation") || "";

    let contextSection = "";
    if (dataContext) {
      contextSection = `\n\n数据上下文：
范围：${dataContext.range}
数据类型：${dataContext.dataType}
示例数据：${JSON.stringify(dataContext.sampleData.slice(0, 3))}`;
    }

    const userPrompt = `需求：${requirement}${contextSection}

请创建合适的Excel公式，返回JSON格式。`;

    return {
      system: systemPrompt,
      user: userPrompt,
      options: {
        temperature: options.temperature || 0.2,
        maxTokens: options.maxTokens || 600,
        format: "json",
      },
    };
  }

  /**
   * 构建图表推荐提示
   */
  buildChartRecommendationPrompt(
    data: any[][],
    headers: string[],
    purpose: string,
    options: PromptOptions = {}
  ): { system: string; user: string; options: PromptOptions } {
    const systemPrompt = this.systemPrompts.get("chart_recommendation") || "";

    const userPrompt = `数据信息：
列名：${headers.join(", ")}
行数：${data.length}
目的：${purpose}

请推荐最适合的图表类型，返回JSON格式。`;

    return {
      system: systemPrompt,
      user: userPrompt,
      options: {
        temperature: options.temperature || 0.3,
        maxTokens: options.maxTokens || 800,
        format: "json",
      },
    };
  }

  /**
   * 构建错误诊断提示
   */
  buildErrorDiagnosisPrompt(
    errorInfo: {
      type: string;
      message: string;
      location?: string;
      context?: string;
    },
    options: PromptOptions = {}
  ): { system: string; user: string; options: PromptOptions } {
    const systemPrompt = this.systemPrompts.get("error_diagnosis") || "";

    const userPrompt = `Excel问题诊断：
错误类型：${errorInfo.type}
错误消息：${errorInfo.message}
${errorInfo.location ? `位置：${errorInfo.location}` : ""}
${errorInfo.context ? `上下文：${errorInfo.context}` : ""}

请诊断问题并提供解决方案，返回JSON格式。`;

    return {
      system: systemPrompt,
      user: userPrompt,
      options: {
        temperature: options.temperature || 0.3,
        maxTokens: options.maxTokens || 1000,
        format: "json",
      },
    };
  }

  /**
   * 构建建议生成提示
   */
  buildSuggestionPrompt(
    context: {
      currentData?: any[][];
      userHistory?: string[];
      workbookState?: WorkbookContext;
    },
    options: PromptOptions = {}
  ): { system: string; user: string; options: PromptOptions } {
    const systemPrompt = `你是一个Excel智能助手，根据用户当前的工作状态提供有用的建议。

建议类型：
- 数据清理和优化
- 公式和计算改进
- 可视化建议
- 最佳实践
- 快捷操作
- 性能优化

输出格式（必须是有效的JSON）：
{
  "suggestions": [
    {
      "type": "data_quality|formula|visualization|best_practice|shortcut",
      "title": "建议标题",
      "description": "详细描述",
      "priority": "low|medium|high",
      "action": "具体操作",
      "benefit": "预期收益"
    }
  ]
}`;

    let userPrompt = "基于当前Excel状态，请提供有用的建议：\n";

    if (context.currentData) {
      userPrompt += `\n当前数据：${context.currentData.length} 行`;
    }
    if (context.userHistory && context.userHistory.length > 0) {
      userPrompt += `\n最近操作：${context.userHistory.slice(-3).join(", ")}`;
    }
    if (context.workbookState) {
      userPrompt += `\n活动工作表：${context.workbookState.activeSheet}`;
    }

    userPrompt += "\n\n请返回JSON格式的建议列表。";

    return {
      system: systemPrompt,
      user: userPrompt,
      options: {
        temperature: options.temperature || 0.5,
        maxTokens: options.maxTokens || 1000,
        format: "json",
      },
    };
  }

  /**
   * 解析AI响应的JSON
   */
  parseAIResponse<T>(response: string): T | null {
    try {
      // 清理响应，移除可能的markdown代码块标记
      const cleaned = response
        .replace(/```json\s*/gi, "")
        .replace(/```\s*$/gi, "")
        .trim();

      return JSON.parse(cleaned) as T;
    } catch (error) {
      console.error("Failed to parse AI response:", error);
      return null;
    }
  }
}
