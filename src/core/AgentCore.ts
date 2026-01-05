/**
 * AgentCore - 对话管理、意图理解、任务规划核心模块
 *
 * @deprecated v2.9.5 - 请使用 src/agent/AgentCore.ts 中的新版 Agent！
 *
 * 迁移说明：
 * - 新版 Agent 使用 ReAct 循环，更灵活
 * - 新版支持硬校验、回滚、问题追踪
 * - 导入方式: import { Agent } from "../../agent";
 *
 * 保留原因：
 * - 测试文件 core-integration.test.ts 仍在使用
 * - 包含有价值的意图理解逻辑
 *
 * 原始设计原则：
 * 1. 严格区分Agent和ChatBot：Agent有明确目标，能规划多步操作
 * 2. 意图理解必须结构化，不能直接传递自然语言给ExcelService
 * 3. 任务规划必须考虑操作依赖和风险
 * 4. 所有LLM交互必须通过PromptBuilder，确保安全可控
 * 5. 通过 WorkbookContext 深度感知 Excel 结构（Excel 感知层）
 */

import {
  ConversationMessage,
  UserIntent,
  IntentType,
  IntentParameters,
  ExecutionPlan,
  ExcelOperation,
  ValidationResult,
  ToolCall,
  ChartType,
  TaskGoal as _TaskGoal,
  TaskReflection as _TaskReflection,
} from "../types";
import { PromptBuilder } from "./PromptBuilder";
import { ExcelService, ExcelOperationResult } from "./ExcelService";
import { DataAnalyzer, AnalysisResult } from "./DataAnalyzer";
import { getAllTools, validateToolParameters } from "./ToolRegistry";
import { WorkbookContext, WorkbookContextData, createWorkbookContext } from "./WorkbookContext";

/**
 * Agent状态
 */
export enum AgentState {
  IDLE = "idle",
  ANALYZING_INTENT = "analyzing_intent",
  PLANNING = "planning",
  EXECUTING = "executing",
  WAITING_FOR_CONFIRMATION = "waiting_for_confirmation",
  ERROR = "error",
  COMPLETED = "completed",
}

/**
 * 操作历史记录
 */
export interface OperationRecord {
  id: string;
  timestamp: Date;
  operation: ExcelOperation;
  result: ExcelOperationResult;
  canUndo: boolean;
  undoData?: any;
}

/**
 * Agent配置
 */
export interface AgentConfig {
  maxConversationHistory: number;
  requireConfirmation: boolean;
  maxPlanSteps: number;
  enableReasoning: boolean;
  allowedIntentTypes: IntentType[];
  enableDataAnalysis: boolean;
  maxOperationHistory: number;
  enableWorkbookContext: boolean;
  contextDepth: "shallow" | "medium" | "deep";
}

/**
 * Agent核心类
 */
export class AgentCore {
  private state: AgentState = AgentState.IDLE;
  private conversationHistory: ConversationMessage[] = [];
  private currentIntent: UserIntent | null = null;
  private currentPlan: ExecutionPlan | null = null;
  private promptBuilder: PromptBuilder;
  private excelService: ExcelService;
  private dataAnalyzer: DataAnalyzer;
  private operationHistory: OperationRecord[] = [];
  private config: AgentConfig;
  private workbookContext: WorkbookContext | null = null;
  private cachedWorkbookData: WorkbookContextData | null = null;

  constructor(excelService: ExcelService, config?: Partial<AgentConfig>) {
    this.excelService = excelService;
    this.promptBuilder = new PromptBuilder();
    this.dataAnalyzer = new DataAnalyzer();

    this.config = {
      maxConversationHistory: 20,
      requireConfirmation: true,
      maxPlanSteps: 10,
      enableReasoning: true,
      enableDataAnalysis: true,
      maxOperationHistory: 50,
      enableWorkbookContext: true,
      contextDepth: "medium",
      allowedIntentTypes: [
        "create_table",
        "format_cells",
        "create_chart",
        "insert_data",
        "apply_filter",
        "insert_formula",
        "sort_data",
        "clear_range",
        "copy_range",
        "merge_cells",
        "analyze_data",
        "generate_summary",
      ],
      ...config,
    };
  }

  /**
   * 初始化工作簿上下文（需要在 Excel.run 内部调用）
   */
  initializeWorkbookContext(context: Excel.RequestContext): void {
    this.workbookContext = createWorkbookContext(context);
  }

  /**
   * 获取工作簿上下文数据
   */
  async getWorkbookContextData(): Promise<WorkbookContextData | null> {
    if (!this.workbookContext || !this.config.enableWorkbookContext) {
      return null;
    }

    try {
      this.cachedWorkbookData = await this.workbookContext.getFullContext(this.config.contextDepth);
      return this.cachedWorkbookData;
    } catch (error) {
      console.error("获取工作簿上下文失败:", error);
      return null;
    }
  }

  /**
   * 获取工作簿上下文摘要（用于 AI Prompt）
   */
  async getWorkbookContextSummary(): Promise<string> {
    if (!this.workbookContext || !this.config.enableWorkbookContext) {
      return "";
    }

    try {
      return await this.workbookContext.getContextSummary();
    } catch (error) {
      console.error("获取工作簿上下文摘要失败:", error);
      return "";
    }
  }

  /**
   * 处理用户输入
   */
  async processUserInput(input: string): Promise<AgentResponse> {
    try {
      // 1. 更新对话历史
      this.addUserMessage(input);

      // 2. 分析用户意图
      this.setState(AgentState.ANALYZING_INTENT);
      const intent = await this.analyzeIntent(input);
      this.currentIntent = intent;

      // 3. 验证意图是否允许
      if (!this.isIntentAllowed(intent.type)) {
        return this.createErrorResponse(
          `不支持的操作类型: ${intent.type}`,
          "请尝试其他类型的Excel操作"
        );
      }

      // 4. 生成执行计划
      this.setState(AgentState.PLANNING);
      const plan = await this.generatePlan(intent);
      this.currentPlan = plan;

      // 5. 验证计划
      const validation = this.validatePlan(plan);
      if (!validation.isValid) {
        return this.createErrorResponse(
          "计划验证失败",
          validation.errors.map((e) => e.message).join(", ")
        );
      }

      // 6. 如果需要确认，等待用户确认
      if (this.config.requireConfirmation && plan.riskLevel !== "low") {
        this.setState(AgentState.WAITING_FOR_CONFIRMATION);
        return this.createConfirmationResponse(plan);
      }

      // 7. 执行计划
      return await this.executePlan(plan);
    } catch (error) {
      this.setState(AgentState.ERROR);
      return this.createErrorResponse(
        "处理用户输入时发生错误",
        error instanceof Error ? error.message : String(error)
      );
    }
  }

  /**
   * 确认并执行计划
   */
  async confirmAndExecute(): Promise<AgentResponse> {
    if (this.state !== AgentState.WAITING_FOR_CONFIRMATION || !this.currentPlan) {
      return this.createErrorResponse("无效状态", "当前没有等待确认的计划");
    }

    try {
      this.setState(AgentState.EXECUTING);
      return await this.executePlan(this.currentPlan);
    } catch (error) {
      this.setState(AgentState.ERROR);
      return this.createErrorResponse(
        "执行计划时发生错误",
        error instanceof Error ? error.message : String(error)
      );
    }
  }

  /**
   * 获取当前状态
   */
  getState(): AgentState {
    return this.state;
  }

  /**
   * 获取对话历史
   */
  getConversationHistory(): ConversationMessage[] {
    return [...this.conversationHistory];
  }

  /**
   * 获取当前计划
   */
  getCurrentPlan(): ExecutionPlan | null {
    return this.currentPlan ? { ...this.currentPlan } : null;
  }

  /**
   * 获取当前意图
   */
  getCurrentIntent(): UserIntent | null {
    return this.currentIntent ? { ...this.currentIntent } : null;
  }

  /**
   * 重置Agent状态
   */
  reset(): void {
    this.state = AgentState.IDLE;
    this.conversationHistory = [];
    this.currentIntent = null;
    this.currentPlan = null;
  }

  /**
   * 私有方法：分析用户意图
   */
  private async analyzeIntent(input: string): Promise<UserIntent> {
    // 使用PromptBuilder构建意图分析Prompt（保留以备将来LLM集成）
    // 注意：当前使用规则引擎，但PromptBuilder已准备好用于LLM集成
    this.promptBuilder.buildIntentAnalysisPrompt(input, this.conversationHistory, getAllTools());

    // 这里应该调用LLM API，但为了简化，我们使用规则引擎
    // 在实际项目中，这里会调用DeepSeek API
    return this.analyzeIntentWithRules(input);
  }

  /**
   * 私有方法：使用规则分析意图（简化实现）
   */
  private analyzeIntentWithRules(input: string): UserIntent {
    const lowerInput = input.toLowerCase();

    // 意图类型映射规则
    const intentRules: Array<{
      type: IntentType;
      keywords: string[];
      extractor: (input: string) => IntentParameters;
    }> = [
      {
        type: "create_table",
        keywords: ["创建表格", "新建表格", "制作表格", "table", "create table"],
        extractor: (input) => this.extractTableParameters(input),
      },
      {
        type: "format_cells",
        keywords: ["格式化", "设置格式", "加粗", "颜色", "format", "bold", "color"],
        extractor: (input) => this.extractFormatParameters(input),
      },
      {
        type: "create_chart",
        keywords: ["创建图表", "制作图表", "图表", "chart", "graph"],
        extractor: (input) => this.extractChartParameters(input),
      },
      {
        type: "insert_data",
        keywords: ["插入数据", "输入数据", "填写", "insert", "add data"],
        extractor: (input) => this.extractDataParameters(input),
      },
      {
        type: "apply_filter",
        keywords: ["筛选", "过滤", "filter"],
        extractor: (input) => this.extractFilterParameters(input),
      },
      {
        type: "insert_formula",
        keywords: ["公式", "计算", "求和", "平均", "formula", "sum", "average"],
        extractor: (input) => this.extractFormulaParameters(input),
      },
      {
        type: "sort_data",
        keywords: ["排序", "sort", "order by"],
        extractor: (input) => this.extractSortParameters(input),
      },
      {
        type: "clear_range",
        keywords: ["清除", "清空", "删除内容", "clear", "delete"],
        extractor: (input) => this.extractRangeParameters(input),
      },
      {
        type: "analyze_data",
        keywords: ["分析", "统计", "分析数据", "analyze", "statistics"],
        extractor: (_input) => ({}),
      },
      {
        type: "generate_summary",
        keywords: ["总结", "汇总", "摘要", "summary"],
        extractor: (_input) => ({}),
      },
    ];

    // 查找匹配的意图
    for (const rule of intentRules) {
      if (rule.keywords.some((keyword) => lowerInput.includes(keyword))) {
        return {
          type: rule.type,
          confidence: 0.8, // 基于规则匹配的置信度
          parameters: rule.extractor(input),
          rawInput: input,
        };
      }
    }

    // 默认返回未知意图
    return {
      type: "unknown",
      confidence: 0.1,
      parameters: {},
      rawInput: input,
    };
  }

  /**
   * 私有方法：提取表格参数
   */
  private extractTableParameters(input: string): IntentParameters {
    // 简化实现：提取范围和数据
    const rangeMatch = input.match(/([A-Z]+[0-9]+:[A-Z]+[0-9]+)/);
    return {
      range: rangeMatch ? rangeMatch[1] : "A1",
      headers: ["列1", "列2", "列3"], // 默认值
    };
  }

  /**
   * 私有方法：提取格式参数
   */
  private extractFormatParameters(input: string): IntentParameters {
    const format: any = {};

    if (input.includes("加粗") || input.includes("bold")) {
      format.bold = true;
    }
    if (input.includes("红色") || input.includes("red")) {
      format.fontColor = "#FF0000";
    }
    if (input.includes("黄色") || input.includes("yellow")) {
      format.fillColor = "#FFFF00";
    }

    const rangeMatch = input.match(/([A-Z]+[0-9]+:[A-Z]+[0-9]+)/);
    return {
      range: rangeMatch ? rangeMatch[1] : "A1",
      format,
    };
  }

  /**
   * 私有方法：提取图表参数
   */
  private extractChartParameters(input: string): IntentParameters {
    let chartType: ChartType = "column";
    if (input.includes("折线") || input.includes("line")) {
      chartType = "line";
    } else if (input.includes("饼") || input.includes("pie")) {
      chartType = "pie";
    } else if (input.includes("条形") || input.includes("bar")) {
      chartType = "bar";
    }

    const rangeMatch = input.match(/([A-Z]+[0-9]+:[A-Z]+[0-9]+)/);
    return {
      range: rangeMatch ? rangeMatch[1] : "A1:B10",
      chartType,
    };
  }

  /**
   * 私有方法：提取数据参数
   */
  private extractDataParameters(_input: string): IntentParameters {
    // 从输入中提取范围
    const rangeMatch = _input.match(/([A-Z]+[0-9]+:[A-Z]+[0-9]+)/);
    return {
      range: rangeMatch ? rangeMatch[1] : "A1",
    };
  }

  /**
   * 私有方法：提取筛选参数
   */
  private extractFilterParameters(_input: string): IntentParameters {
    // 从输入中提取范围
    const rangeMatch = _input.match(/([A-Z]+[0-9]+:[A-Z]+[0-9]+)/);
    return {
      range: rangeMatch ? rangeMatch[1] : "A1:D100",
    };
  }

  /**
   * 私有方法：提取公式参数
   */
  private extractFormulaParameters(input: string): IntentParameters {
    let formula = "SUM";
    if (input.includes("平均") || input.includes("average")) {
      formula = "AVERAGE";
    } else if (input.includes("最大") || input.includes("max")) {
      formula = "MAX";
    } else if (input.includes("最小") || input.includes("min")) {
      formula = "MIN";
    }

    const rangeMatch = input.match(/([A-Z]+[0-9]+:[A-Z]+[0-9]+)/);
    return {
      range: rangeMatch ? rangeMatch[1] : "A1:A10",
      formula,
    };
  }

  /**
   * 私有方法：提取排序参数
   */
  private extractSortParameters(input: string): IntentParameters {
    const ascending = !(input.includes("降序") || input.includes("desc"));
    const rangeMatch = input.match(/([A-Z]+[0-9]+:[A-Z]+[0-9]+)/);
    return {
      range: rangeMatch ? rangeMatch[1] : "A1:D100",
      ascending,
    };
  }

  /**
   * 私有方法：提取范围参数
   */
  private extractRangeParameters(input: string): IntentParameters {
    const rangeMatch = input.match(/([A-Z]+[0-9]+:[A-Z]+[0-9]+)/);
    return {
      range: rangeMatch ? rangeMatch[1] : "A1",
    };
  }

  /**
   * 私有方法：生成执行计划
   */
  private async generatePlan(intent: UserIntent): Promise<ExecutionPlan> {
    // 将意图转换为Excel操作序列
    const operations = this.intentToOperations(intent);

    // 分析操作依赖
    const dependencies = this.analyzeDependencies(operations);

    // 评估风险等级
    const riskLevel = this.assessRiskLevel(operations);

    return {
      id: `plan_${Date.now()}`,
      operations,
      dependencies,
      estimatedTime: operations.length * 2, // 每步操作估计2秒
      riskLevel,
      validationResults: [],
    };
  }

  /**
   * 私有方法：将意图转换为操作序列
   */
  private intentToOperations(intent: UserIntent): ExcelOperation[] {
    const operations: ExcelOperation[] = [];

    switch (intent.type) {
      case "create_table":
        operations.push({
          id: "create_table_1",
          type: "create_table",
          description: "创建表格",
          parameters: {
            range: intent.parameters.range || "A1",
            data: intent.parameters.data,
          },
          validationRules: [{ field: "range", type: "required", message: "必须指定范围" }],
          executable: true,
        });
        break;

      case "format_cells":
        operations.push({
          id: "format_cells_1",
          type: "format_cells",
          description: "格式化单元格",
          parameters: {
            range: intent.parameters.range || "A1",
            format: intent.parameters.format,
          },
          validationRules: [{ field: "range", type: "required", message: "必须指定范围" }],
          executable: true,
        });
        break;

      case "create_chart":
        operations.push({
          id: "create_chart_1",
          type: "create_chart",
          description: "创建图表",
          parameters: {
            range: intent.parameters.range || "A1:B10",
            chartType: intent.parameters.chartType || "column",
          },
          validationRules: [{ field: "range", type: "required", message: "必须指定数据范围" }],
          executable: true,
        });
        break;

      // 其他意图类型...
    }

    return operations;
  }

  /**
   * 私有方法：分析操作依赖
   */
  private analyzeDependencies(_operations: ExcelOperation[]): any[] {
    // 简化实现：假设操作按顺序执行，没有复杂依赖
    // 注意：在实际实现中，应分析操作之间的依赖关系
    void _operations; // 明确表示参数未使用
    return [];
  }

  /**
   * 私有方法：评估风险等级
   */
  private assessRiskLevel(operations: ExcelOperation[]): "low" | "medium" | "high" {
    // 简化风险评估
    const riskyOperations = operations.filter((op) => op.type === "clear_range");

    if (riskyOperations.length > 0) {
      return "high";
    } else if (operations.length > 3) {
      return "medium";
    } else {
      return "low";
    }
  }

  /**
   * 私有方法：验证计划
   */
  private validatePlan(plan: ExecutionPlan): ValidationResult {
    const errors: any[] = [];
    const warnings: any[] = [];

    // 检查操作数量限制
    if (plan.operations.length > this.config.maxPlanSteps) {
      errors.push({
        field: "operations",
        message: `操作步骤过多（${plan.operations.length} > ${this.config.maxPlanSteps}）`,
        code: "MAX_STEPS_EXCEEDED",
      });
    }

    // 检查高风险操作
    if (plan.riskLevel === "high") {
      warnings.push({
        field: "risk",
        message: "计划包含高风险操作",
        severity: "high",
      });
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings,
    };
  }

  /**
   * 私有方法：执行计划
   */
  private async executePlan(plan: ExecutionPlan): Promise<AgentResponse> {
    const results = [];

    for (const operation of plan.operations) {
      try {
        // 将Excel操作转换为工具调用
        const toolCall = this.operationToToolCall(operation);

        // 验证工具参数
        const validation = validateToolParameters(toolCall.name, toolCall.arguments);
        if (!validation.isValid) {
          throw new Error(`参数验证失败: ${validation.errors.join(", ")}`);
        }

        // 执行工具
        const result = await this.excelService.executeTool(toolCall.name, toolCall.arguments);

        results.push({
          operationId: operation.id,
          success: result.success,
          result: result.data,
          error: result.error,
          executionTime: 0, // 简化实现
        });
      } catch (error) {
        results.push({
          operationId: operation.id,
          success: false,
          result: null,
          error: error instanceof Error ? error.message : String(error),
          executionTime: 0,
        });
      }
    }

    // 更新状态
    this.setState(AgentState.COMPLETED);
    this.addAssistantMessage("计划执行完成", { results });

    return {
      success: true,
      message: "计划执行完成",
      data: {
        planId: plan.id,
        results,
        totalOperations: plan.operations.length,
        successfulOperations: results.filter((r) => r.success).length,
      },
      requiresConfirmation: false,
    };
  }

  /**
   * 私有方法：将Excel操作转换为工具调用
   */
  private operationToToolCall(operation: ExcelOperation): ToolCall {
    // 简化实现：根据操作类型映射到工具
    const toolMapping: Record<string, string> = {
      create_table: "excel.set_range_values",
      format_cells: "excel.format_range",
      create_chart: "excel.create_chart",
      insert_data: "excel.set_range_values",
      apply_filter: "analysis.filter_range",
      insert_formula: "analysis.sum_range", // 简化
      sort_data: "analysis.sort_range",
      clear_range: "excel.clear_range",
    };

    const toolName = toolMapping[operation.type] || "excel.set_cell_value";

    return {
      id: `tool_${Date.now()}_${operation.id}`,
      name: toolName,
      arguments: operation.parameters,
    };
  }

  /**
   * 私有方法：设置状态
   */
  private setState(state: AgentState): void {
    this.state = state;
  }

  /**
   * 私有方法：添加用户消息
   */
  private addUserMessage(content: string): void {
    this.conversationHistory.push({
      id: `msg_${Date.now()}_user`,
      role: "user",
      content,
      timestamp: new Date(),
    });

    // 限制历史记录长度
    if (this.conversationHistory.length > this.config.maxConversationHistory) {
      this.conversationHistory = this.conversationHistory.slice(
        -this.config.maxConversationHistory
      );
    }
  }

  /**
   * 私有方法：添加助手消息
   */
  private addAssistantMessage(content: string, metadata?: any): void {
    this.conversationHistory.push({
      id: `msg_${Date.now()}_assistant`,
      role: "assistant",
      content,
      timestamp: new Date(),
      metadata,
    });
  }

  /**
   * 私有方法：检查意图是否允许
   */
  private isIntentAllowed(intentType: IntentType): boolean {
    return this.config.allowedIntentTypes.includes(intentType);
  }

  /**
   * 私有方法：创建错误响应
   */
  private createErrorResponse(title: string, details: string): AgentResponse {
    this.setState(AgentState.ERROR);
    this.addAssistantMessage(`错误: ${title} - ${details}`);

    return {
      success: false,
      message: title,
      error: details,
      requiresConfirmation: false,
    };
  }

  /**
   * 私有方法：创建确认响应
   */
  private createConfirmationResponse(plan: ExecutionPlan): AgentResponse {
    const operationDescriptions = plan.operations
      .map((op) => `• ${op.description} (${op.type})`)
      .join("\n");

    const message = `请确认以下操作计划：\n\n${operationDescriptions}\n\n风险等级: ${plan.riskLevel}\n预计时间: ${plan.estimatedTime}秒`;

    this.addAssistantMessage(message, { planId: plan.id });

    return {
      success: true,
      message: "需要确认",
      data: {
        planId: plan.id,
        operations: plan.operations,
        riskLevel: plan.riskLevel,
        estimatedTime: plan.estimatedTime,
      },
      requiresConfirmation: true,
    };
  }

  // ==================== 数据分析整合 ====================

  /**
   * 执行数据分析
   */
  async analyzeSelectedData(): Promise<AnalysisResult | null> {
    if (!this.config.enableDataAnalysis) {
      return null;
    }

    try {
      // 获取当前选区数据
      const selectedRange = await this.excelService.getSelectedRange();
      if (!selectedRange.success || !selectedRange.data) {
        return null;
      }

      const { values, address } = selectedRange.data;
      if (!values || values.length === 0) {
        return null;
      }

      // 假设第一行是表头
      const headers = values[0]?.map(String);
      const dataRows = values.slice(1);

      // 执行分析
      const analysisResult = await this.dataAnalyzer.analyzeData(dataRows, headers, {
        includeStatistics: true,
        includeInsights: true,
        includeQuality: true,
        includeRecommendations: true,
      });

      // 记录分析操作
      this.addAssistantMessage(
        `已分析 ${address} 区域的数据：${analysisResult.summary.rowCount} 行 × ${analysisResult.summary.columnCount} 列`
      );

      return analysisResult;
    } catch (error) {
      console.error("数据分析失败:", error);
      return null;
    }
  }

  /**
   * 根据分析结果生成建议消息
   */
  generateInsightMessage(analysis: AnalysisResult): string {
    const lines: string[] = [];

    // 数据概览
    lines.push(`📊 **数据概览**`);
    lines.push(`- 行数: ${analysis.summary.rowCount}`);
    lines.push(`- 列数: ${analysis.summary.columnCount}`);
    lines.push(`- 数值列: ${analysis.summary.numericColumns}`);
    lines.push(`- 文本列: ${analysis.summary.textColumns}`);
    lines.push("");

    // 数据质量
    lines.push(`📋 **数据质量评分**: ${analysis.quality.score}/100 (${analysis.quality.overall})`);
    if (analysis.quality.issues.length > 0) {
      lines.push(`发现 ${analysis.quality.issues.length} 个问题：`);
      analysis.quality.issues.slice(0, 3).forEach((issue) => {
        lines.push(`  - ${issue.description}`);
      });
    }
    lines.push("");

    // 洞察
    if (analysis.insights.length > 0) {
      lines.push(`💡 **关键洞察**`);
      analysis.insights.slice(0, 5).forEach((insight) => {
        lines.push(`- **${insight.title}**: ${insight.description}`);
      });
      lines.push("");
    }

    // 建议
    if (analysis.recommendations.length > 0) {
      lines.push(`🎯 **建议操作**`);
      analysis.recommendations.slice(0, 3).forEach((rec, index) => {
        lines.push(`${index + 1}. ${rec.title}: ${rec.description}`);
      });
    }

    return lines.join("\n");
  }

  // ==================== 操作历史管理 ====================

  /**
   * 记录操作到历史
   */
  private recordOperation(
    operation: ExcelOperation,
    result: ExcelOperationResult,
    undoData?: any
  ): void {
    const record: OperationRecord = {
      id: `op_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
      timestamp: new Date(),
      operation,
      result,
      canUndo: undoData !== undefined,
      undoData,
    };

    this.operationHistory.unshift(record);

    // 限制历史记录数量
    if (this.operationHistory.length > this.config.maxOperationHistory) {
      this.operationHistory = this.operationHistory.slice(0, this.config.maxOperationHistory);
    }
  }

  /**
   * 获取操作历史
   */
  getOperationHistory(): OperationRecord[] {
    return [...this.operationHistory];
  }

  /**
   * 撤销上一个操作
   */
  async undoLastOperation(): Promise<AgentResponse> {
    const lastUndoable = this.operationHistory.find((op) => op.canUndo);
    if (!lastUndoable) {
      return this.createErrorResponse("无法撤销", "没有可撤销的操作");
    }

    try {
      // 根据操作类型执行撤销
      const undoResult = await this.performUndo(lastUndoable);
      if (undoResult.success) {
        // 移除已撤销的操作
        this.operationHistory = this.operationHistory.filter((op) => op.id !== lastUndoable.id);
        this.addAssistantMessage(`已撤销操作: ${lastUndoable.operation.description}`);
        return {
          success: true,
          message: `已撤销: ${lastUndoable.operation.description}`,
          requiresConfirmation: false,
        };
      } else {
        return this.createErrorResponse("撤销失败", undoResult.error || "未知错误");
      }
    } catch (error) {
      return this.createErrorResponse(
        "撤销失败",
        error instanceof Error ? error.message : String(error)
      );
    }
  }

  /**
   * 执行撤销操作
   */
  private async performUndo(record: OperationRecord): Promise<ExcelOperationResult> {
    const { operation, undoData } = record;

    switch (operation.type) {
      case "set_cell_value":
      case "set_range_values":
        // 恢复原始值
        if (undoData?.originalValues && undoData?.range) {
          return await this.excelService.setRangeValues(undoData.range, undoData.originalValues);
        }
        break;

      case "format_range":
        // 格式化操作的撤销比较复杂，暂时返回失败
        return {
          success: false,
          error: "格式化操作暂不支持撤销",
          timestamp: Date.now(),
        };

      case "clear_range":
        // 恢复清除前的数据
        if (undoData?.originalValues && undoData?.range) {
          return await this.excelService.setRangeValues(undoData.range, undoData.originalValues);
        }
        break;

      default:
        return {
          success: false,
          error: `${operation.type} 类型的操作不支持撤销`,
          timestamp: Date.now(),
        };
    }

    return {
      success: false,
      error: "缺少撤销所需的数据",
      timestamp: Date.now(),
    };
  }

  /**
   * 清除操作历史
   */
  clearOperationHistory(): void {
    this.operationHistory = [];
  }

  // ==================== 上下文管理 ====================

  /**
   * 获取对话上下文摘要（用于多轮对话）
   */
  getContextSummary(): string {
    if (this.conversationHistory.length === 0) {
      return "这是新对话的开始。";
    }

    const recentMessages = this.conversationHistory.slice(-6);
    const summary = recentMessages
      .map((msg) => `${msg.role === "user" ? "用户" : "助手"}: ${msg.content.slice(0, 100)}...`)
      .join("\n");

    return `最近对话:\n${summary}`;
  }

  /**
   * 构建增强的提示词（包含上下文）
   */
  buildContextualPrompt(userInput: string): string {
    const context = this.getContextSummary();
    const currentIntent = this.currentIntent ? `当前理解的意图: ${this.currentIntent.type}` : "";

    return `
${context}

${currentIntent}

用户新消息: ${userInput}
    `.trim();
  }

  /**
   * 构建包含工作簿上下文的智能提示词
   */
  async buildSmartPrompt(userInput: string): Promise<string> {
    const conversationContext = this.getContextSummary();
    const currentIntent = this.currentIntent ? `当前理解的意图: ${this.currentIntent.type}` : "";

    // 获取工作簿上下文
    let workbookContext = "";
    if (this.config.enableWorkbookContext) {
      workbookContext = await this.getWorkbookContextSummary();
    }

    // 构建带有完整上下文的提示词
    const sections: string[] = [];

    // 1. 工作簿上下文（如果可用）
    if (workbookContext) {
      sections.push("# 当前工作簿状态");
      sections.push(workbookContext);
      sections.push("");
    }

    // 2. 对话历史
    sections.push("# 对话上下文");
    sections.push(conversationContext);
    sections.push("");

    // 3. 当前意图（如果已识别）
    if (currentIntent) {
      sections.push("# 已识别意图");
      sections.push(currentIntent);
      sections.push("");
    }

    // 4. 用户新输入
    sections.push("# 用户请求");
    sections.push(userInput);

    return sections.join("\n");
  }

  /**
   * 记录操作变更到工作簿上下文
   */
  recordOperationChange(operation: ExcelOperation): void {
    if (this.workbookContext) {
      const params = operation.parameters as Record<string, any> | undefined;
      this.workbookContext.recordChange({
        type: this.mapOperationTypeToChangeType(operation.type),
        sheetName: params?.sheetName || "未知",
        range: params?.range || params?.cellAddress || "未知",
        description: this.getOperationDescription(operation),
      });
      // 操作后使缓存失效
      this.workbookContext.invalidateCache();
    }
  }

  private mapOperationTypeToChangeType(
    opType: string
  ): "value" | "format" | "structure" | "selection" {
    const valueOps = ["set_cell_value", "set_range_values", "clear_range", "insert_data"];
    const formatOps = ["format_cells", "add_conditional_format", "set_borders"];
    const structureOps = [
      "insert_rows",
      "insert_columns",
      "delete_rows",
      "delete_columns",
      "merge_cells",
    ];

    if (valueOps.some((op) => opType.includes(op))) return "value";
    if (formatOps.some((op) => opType.includes(op))) return "format";
    if (structureOps.some((op) => opType.includes(op))) return "structure";
    return "selection";
  }

  private getOperationDescription(operation: ExcelOperation): string {
    const typeMap: Record<string, string> = {
      select_range: "选择范围",
      set_cell_value: "设置单元格值",
      set_range_values: "设置范围值",
      clear_range: "清除范围",
      format_cells: "格式化单元格",
      create_chart: "创建图表",
      insert_formula: "插入公式",
    };
    return typeMap[operation.type] || operation.type;
  }

  /**
   * 获取缓存的工作簿数据
   */
  getCachedWorkbookData(): WorkbookContextData | null {
    return this.cachedWorkbookData;
  }

  /**
   * 获取当前选区信息（快速方法）
   */
  async getCurrentSelectionInfo(): Promise<{
    address: string;
    sheetName: string;
    rowCount: number;
    columnCount: number;
    hasData: boolean;
    preview: any[][];
  } | null> {
    if (!this.workbookContext) return null;

    try {
      const selection = await this.workbookContext.getCurrentSelection();
      if (!selection) return null;

      return {
        address: selection.range.address,
        sheetName: selection.sheetName,
        rowCount: selection.range.rowCount,
        columnCount: selection.range.columnCount,
        hasData: selection.range.hasValues,
        preview: selection.values.slice(0, 5).map((row) => row.slice(0, 5)),
      };
    } catch {
      return null;
    }
  }
}

/**
 * Agent响应接口
 */
export interface AgentResponse {
  success: boolean;
  message: string;
  data?: any;
  error?: string;
  requiresConfirmation: boolean;
}
