/**
 * ProactiveAgent - 主动洞察型 Agent
 *
 * 不再被动等待指令，而是：
 * 1. 主动观察和分析工作表
 * 2. 形成判断和洞察
 * 3. 提供建议
 * 4. 最后才询问确认
 *
 * 像一个有经验的数据分析师一样工作
 *
 * @module agent/proactive/ProactiveAgent
 */

import {
  WorksheetAnalyzer,
  WorksheetAnalysis,
  createWorksheetAnalyzer,
} from "./WorksheetAnalyzer";
import {
  InsightGenerator,
  InsightReport,
  Suggestion,
  createInsightGenerator,
} from "./InsightGenerator";
import { ToolRegistry } from "../registry";
import { AgentTracer } from "../tracing";

// ========== 类型定义 ==========

/**
 * Agent 状态
 */
export type ProactiveAgentState =
  | "idle"           // 空闲
  | "observing"      // 正在观察
  | "analyzing"      // 正在分析
  | "presenting"     // 展示洞察
  | "awaiting"       // 等待用户确认
  | "executing"      // 执行中
  | "completed";     // 完成

/**
 * Agent 消息
 */
export interface AgentMessage {
  id: string;
  type: "observation" | "insight" | "suggestion" | "question" | "action" | "result" | "error";
  content: string;
  timestamp: Date;
  metadata?: Record<string, unknown>;
}

/**
 * 对话上下文
 */
export interface ConversationContext {
  messages: AgentMessage[];
  currentAnalysis: WorksheetAnalysis | null;
  currentInsights: InsightReport | null;
  pendingSuggestions: Suggestion[];
  userPreferences: UserPreferences;
}

/**
 * 用户偏好
 */
export interface UserPreferences {
  verbosity: "brief" | "normal" | "detailed";
  autoExecute: boolean;
  confirmBeforeAction: boolean;
}

/**
 * Agent 配置
 */
export interface ProactiveAgentConfig {
  autoAnalyzeOnStart: boolean;
  autoAnalyzeOnSheetChange: boolean;
  minRowsForAutoAnalysis: number;
  verboseLogging: boolean;
}

/**
 * Agent 事件
 */
export type AgentEventType =
  | "state:change"
  | "message:new"
  | "analysis:complete"
  | "insight:ready"
  | "action:start"
  | "action:complete"
  | "error";

export type AgentEventHandler = (event: AgentEventType, data: unknown) => void;

/**
 * 用户意图类型
 */
interface Intent {
  type: "formula_direct" | "formula_natural" | "calculation" | "analyze" | "format" | "data_operation" | "unknown";
  cell?: string;
  formula?: string;
  expression?: string;
  operation?: "sum" | "average" | "max" | "min";
  targetCell?: string;
  sourceRange?: string;
  scope?: "current" | "all";
  rawInput?: string;
}

// ========== 默认配置 ==========

const DEFAULT_CONFIG: ProactiveAgentConfig = {
  autoAnalyzeOnStart: true,
  autoAnalyzeOnSheetChange: true,
  minRowsForAutoAnalysis: 3,
  verboseLogging: true,
};

const DEFAULT_PREFERENCES: UserPreferences = {
  verbosity: "normal",
  autoExecute: false,
  confirmBeforeAction: true,
};

// ========== ProactiveAgent 类 ==========

export class ProactiveAgent {
  private analyzer: WorksheetAnalyzer;
  private insightGenerator: InsightGenerator;
  private toolRegistry: ToolRegistry;
  private tracer: AgentTracer;

  private state: ProactiveAgentState = "idle";
  private config: ProactiveAgentConfig;
  private context: ConversationContext;

  private eventHandlers: AgentEventHandler[] = [];
  private messageIdCounter = 0;

  constructor(
    toolRegistry: ToolRegistry,
    options?: {
      config?: Partial<ProactiveAgentConfig>;
      preferences?: Partial<UserPreferences>;
    }
  ) {
    this.toolRegistry = toolRegistry;
    this.config = { ...DEFAULT_CONFIG, ...options?.config };

    this.analyzer = createWorksheetAnalyzer();
    this.insightGenerator = createInsightGenerator();
    this.tracer = new AgentTracer({ enabled: this.config.verboseLogging });

    this.context = {
      messages: [],
      currentAnalysis: null,
      currentInsights: null,
      pendingSuggestions: [],
      userPreferences: { ...DEFAULT_PREFERENCES, ...options?.preferences },
    };
  }

  // ========== 公共 API ==========

  /**
   * 启动 Agent（主动开始观察）
   */
  async start(): Promise<void> {
    this.log("info", "ProactiveAgent 启动");

    if (this.config.autoAnalyzeOnStart) {
      await this.observeAndAnalyze();
    }
  }

  /**
   * 主动观察和分析当前工作表
   */
  async observeAndAnalyze(sheetName?: string): Promise<InsightReport | null> {
    const span = this.tracer.startSpan("observe_and_analyze");

    try {
      this.setState("observing");
      this.addMessage("observation", "正在观察工作表...");

      // 静默扫描
      const analysis = await this.analyzer.analyze(sheetName);
      this.context.currentAnalysis = analysis;

      this.log("info", `分析完成: ${analysis.totalRows} 行 × ${analysis.totalColumns} 列`);

      // 如果数据太少，不生成洞察
      if (analysis.totalRows < this.config.minRowsForAutoAnalysis) {
        this.setState("idle");
        this.addMessage("observation", "工作表数据较少，有什么我可以帮你的吗？");
        return null;
      }

      this.setState("analyzing");

      // 生成洞察
      const insights = this.insightGenerator.generate(analysis);
      this.context.currentInsights = insights;
      this.context.pendingSuggestions = insights.suggestions;

      this.emit("analysis:complete", { analysis, insights });

      // 展示洞察
      this.setState("presenting");
      this.presentInsights(insights);

      return insights;
    } catch (error) {
      this.handleError(error);
      return null;
    } finally {
      this.tracer.endSpan("success");
    }
  }

  /**
   * 处理用户输入
   */
  async handleUserInput(input: string): Promise<string> {
    const span = this.tracer.startSpan("handle_user_input");

    try {
      const normalizedInput = input.trim().toLowerCase();

      // 快速响应模式
      if (this.isQuickResponse(normalizedInput)) {
        return await this.handleQuickResponse(normalizedInput);
      }

      // 如果没有当前分析，先分析
      if (!this.context.currentAnalysis) {
        await this.observeAndAnalyze();
        if (!this.context.currentInsights) {
          return "我先看看这个表格...";
        }
      }

      // 解析用户意图
      const userIntent = this.parseUserIntent(normalizedInput);

      switch (userIntent.type) {
        case "confirm_all":
          return await this.executeAllSuggestions();

        case "confirm_specific":
          return await this.executeSuggestion(userIntent.suggestionId!);

        case "reject":
          return this.handleRejection();

        case "ask_detail":
          return this.provideDetail(userIntent.topic);

        case "new_request":
          return await this.handleNewRequest(input);

        default:
          return this.handleUnknownIntent(input);
      }
    } catch (error) {
      this.handleError(error);
      return "抱歉，处理过程中出了点问题。";
    } finally {
      this.tracer.endSpan("success");
    }
  }

  /**
   * 执行所有建议
   */
  async executeAllSuggestions(): Promise<string> {
    const suggestions = this.context.pendingSuggestions.filter((s) => s.autoExecutable);

    if (suggestions.length === 0) {
      return "没有可自动执行的建议。";
    }

    this.setState("executing");
    this.addMessage("action", `开始执行 ${suggestions.length} 项优化...`);

    const results: string[] = [];
    let successCount = 0;

    for (const suggestion of suggestions) {
      try {
        const result = await this.executeSuggestionActions(suggestion);
        if (result.success) {
          successCount++;
          results.push(`✅ ${suggestion.title}`);
        } else {
          results.push(`❌ ${suggestion.title}: ${result.error}`);
        }
      } catch (error) {
        results.push(`❌ ${suggestion.title}: ${error instanceof Error ? error.message : "执行失败"}`);
      }
    }

    this.setState("completed");

    const summary = successCount === suggestions.length
      ? `全部完成！${successCount} 项优化已应用。`
      : `完成 ${successCount}/${suggestions.length} 项。`;

    this.addMessage("result", summary + "\n" + results.join("\n"));

    return summary + "\n\n" + results.join("\n");
  }

  /**
   * 执行单个建议
   */
  async executeSuggestion(suggestionId: string): Promise<string> {
    const suggestion = this.context.pendingSuggestions.find((s) => s.id === suggestionId);

    if (!suggestion) {
      return `找不到建议: ${suggestionId}`;
    }

    this.setState("executing");
    this.addMessage("action", `执行: ${suggestion.title}...`);

    try {
      const result = await this.executeSuggestionActions(suggestion);

      this.setState("completed");

      if (result.success) {
        const msg = `✅ ${suggestion.title} - 完成`;
        this.addMessage("result", msg);
        return msg;
      } else {
        const msg = `❌ ${suggestion.title} - ${result.error}`;
        this.addMessage("error", msg);
        return msg;
      }
    } catch (error) {
      this.setState("completed");
      const msg = `执行失败: ${error instanceof Error ? error.message : "未知错误"}`;
      this.addMessage("error", msg);
      return msg;
    }
  }

  /**
   * 获取当前状态
   */
  getState(): ProactiveAgentState {
    return this.state;
  }

  /**
   * 获取对话历史
   */
  getMessages(): AgentMessage[] {
    return this.context.messages;
  }

  /**
   * 获取当前洞察
   */
  getCurrentInsights(): InsightReport | null {
    return this.context.currentInsights;
  }

  /**
   * 订阅事件
   */
  on(handler: AgentEventHandler): () => void {
    this.eventHandlers.push(handler);
    return () => {
      this.eventHandlers = this.eventHandlers.filter((h) => h !== handler);
    };
  }

  /**
   * 重置
   */
  reset(): void {
    this.state = "idle";
    this.context = {
      messages: [],
      currentAnalysis: null,
      currentInsights: null,
      pendingSuggestions: [],
      userPreferences: this.context.userPreferences,
    };
  }

  // ========== 私有方法 ==========

  private setState(newState: ProactiveAgentState): void {
    const oldState = this.state;
    this.state = newState;
    this.emit("state:change", { from: oldState, to: newState });
  }

  private addMessage(type: AgentMessage["type"], content: string, metadata?: Record<string, unknown>): void {
    const message: AgentMessage = {
      id: `msg_${++this.messageIdCounter}`,
      type,
      content,
      timestamp: new Date(),
      metadata,
    };
    this.context.messages.push(message);
    this.emit("message:new", message);
  }

  private presentInsights(insights: InsightReport): void {
    // 先展示叙述性描述
    this.addMessage("insight", insights.narrativeDescription);

    // 然后展示对话提示
    setTimeout(() => {
      this.addMessage("question", insights.conversationPrompt);
      this.setState("awaiting");
    }, 500);
  }

  private isQuickResponse(input: string): boolean {
    const quickPatterns = [
      "好", "好的", "可以", "行", "做吧", "全部", "都做",
      "yes", "ok", "sure", "do it", "all",
      "不", "不要", "算了", "取消", "no", "cancel",
    ];
    return quickPatterns.some((p) => input === p || input.startsWith(p));
  }

  private async handleQuickResponse(input: string): Promise<string> {
    const positivePatterns = ["好", "好的", "可以", "行", "做吧", "全部", "都做", "yes", "ok", "sure", "do it", "all"];
    const negativePatterns = ["不", "不要", "算了", "取消", "no", "cancel"];

    if (positivePatterns.some((p) => input.includes(p))) {
      return await this.executeAllSuggestions();
    }

    if (negativePatterns.some((p) => input.includes(p))) {
      return this.handleRejection();
    }

    return "我不太确定你的意思，可以说得更具体一些吗？";
  }

  private parseUserIntent(input: string): {
    type: "confirm_all" | "confirm_specific" | "reject" | "ask_detail" | "new_request" | "unknown";
    suggestionId?: string;
    topic?: string;
  } {
    // 全部确认
    if (/全部|都|一起|all|both/i.test(input)) {
      return { type: "confirm_all" };
    }

    // 拒绝
    if (/不|算了|取消|暂时|以后|no|cancel|later/i.test(input)) {
      return { type: "reject" };
    }

    // 具体建议
    const suggestions = this.context.pendingSuggestions;
    for (const s of suggestions) {
      if (input.includes(s.title) || input.includes(s.id)) {
        return { type: "confirm_specific", suggestionId: s.id };
      }
    }

    // 询问详情
    if (/什么|为什么|怎么|why|what|how/i.test(input)) {
      return { type: "ask_detail", topic: input };
    }

    // 新请求
    return { type: "new_request" };
  }

  private handleRejection(): string {
    this.setState("idle");
    return "好的，有需要的时候随时叫我。";
  }

  private provideDetail(topic?: string): string {
    const insights = this.context.currentInsights;
    if (!insights) {
      return "我还没有分析这个表格，要我先看看吗？";
    }

    // 提供更详细的解释
    const details: string[] = [];
    for (const insight of insights.insights) {
      details.push(`**${insight.title}**`);
      details.push(insight.description);
      details.push("");
    }

    return details.join("\n");
  }

  private async handleNewRequest(input: string): Promise<string> {
    // 智能意图解析和执行
    this.setState("executing");

    try {
      // 解析用户意图
      const intent = this.parseAdvancedIntent(input);
      
      this.log("info", `解析意图: ${JSON.stringify(intent)}`);

      switch (intent.type) {
        case "formula_direct":
          return await this.executeFormula(intent.cell!, intent.formula!);
        
        case "formula_natural":
          return await this.executeFormulaFromNatural(intent);
        
        case "calculation":
          return await this.executeCalculation(intent);
        
        case "analyze":
          return await this.executeAnalysis(intent);
        
        case "format":
          return await this.executeFormat(intent);
        
        case "data_operation":
          return await this.executeDataOperation(intent);
        
        default:
          return this.handleUnknownRequest(input);
      }
    } catch (error) {
      this.setState("idle");
      this.addMessage("error", error instanceof Error ? error.message : "执行失败");
      return `❌ ${error instanceof Error ? error.message : "执行失败"}`;
    }
  }

  /**
   * 高级意图解析 - 理解多种表达方式
   */
  private parseAdvancedIntent(input: string): Intent {
    const normalized = input.trim();

    // 1. 直接公式格式：K2=E2-F2-J2-J2
    const directFormulaMatch = normalized.match(/^([A-Z]+\d+)\s*=\s*(.+)$/i);
    if (directFormulaMatch) {
      return {
        type: "formula_direct",
        cell: directFormulaMatch[1],
        formula: directFormulaMatch[2],
      };
    }

    // 2. 自然语言公式：帮我在 K2 设置 E2-F2-J2-J2
    const naturalFormulaPatterns = [
      /(?:帮我)?(?:在|给)?([A-Z]+\d+)(?:设置|写入|填入|计算)?(.+)/i,
      /(?:让|把)?([A-Z]+\d+)(?:等于|=)(.+)/i,
      /(?:计算|求)?([A-Z]+\d+)(?:的值)?[:：]?(.+)/i,
    ];
    
    for (const pattern of naturalFormulaPatterns) {
      const match = normalized.match(pattern);
      if (match) {
        const [, cell, expression] = match;
        if (/^[A-Z]+\d+$/i.test(cell)) {
          return {
            type: "formula_natural",
            cell,
            expression: expression.trim(),
          };
        }
      }
    }

    // 3. 计算指令：求和、平均、最大值等
    if (/求和|总和|sum/i.test(normalized)) {
      const cellMatch = normalized.match(/([A-Z]+\d+)/i);
      const rangeMatch = normalized.match(/([A-Z]+\d+:[A-Z]+\d+)/i);
      return {
        type: "calculation",
        operation: "sum",
        targetCell: cellMatch?.[1],
        sourceRange: rangeMatch?.[1],
        rawInput: normalized,
      };
    }

    if (/平均|平均值|average/i.test(normalized)) {
      const cellMatch = normalized.match(/([A-Z]+\d+)/i);
      const rangeMatch = normalized.match(/([A-Z]+\d+:[A-Z]+\d+)/i);
      return {
        type: "calculation",
        operation: "average",
        targetCell: cellMatch?.[1],
        sourceRange: rangeMatch?.[1],
        rawInput: normalized,
      };
    }

    if (/最大|max/i.test(normalized)) {
      return {
        type: "calculation",
        operation: "max",
        rawInput: normalized,
      };
    }

    if (/最小|min/i.test(normalized)) {
      return {
        type: "calculation",
        operation: "min",
        rawInput: normalized,
      };
    }

    // 4. 分析指令
    if (/分析|统计|查看|检查/i.test(normalized)) {
      return {
        type: "analyze",
        scope: normalized.includes("全部") || normalized.includes("所有") ? "all" : "current",
      };
    }

    // 5. 格式化指令
    if (/格式|样式|颜色|加粗|斜体/i.test(normalized)) {
      return {
        type: "format",
        rawInput: normalized,
      };
    }

    // 6. 数据操作
    if (/删除|清除|复制|粘贴|移动/i.test(normalized)) {
      return {
        type: "data_operation",
        rawInput: normalized,
      };
    }

    // 未知意图
    return {
      type: "unknown",
      rawInput: normalized,
    };
  }

  /**
   * 执行直接公式
   */
  private async executeFormula(cell: string, formula: string): Promise<string> {
    const tool = this.toolRegistry.get("excel_set_formula");
    if (!tool) {
      return `❌ 工具不可用：excel_set_formula`;
    }

    this.addMessage("action", `正在设置 ${cell} 的公式...`);

    const result = await tool.execute({
      address: cell,
      formula: formula.startsWith("=") ? formula : `=${formula}`,
    });

    this.setState("completed");

    if (result.success) {
      this.addMessage("result", result.output);
      return result.output;
    } else {
      this.addMessage("error", result.error || "执行失败");
      return `❌ ${result.error || "执行失败"}`;
    }
  }

  /**
   * 从自然语言解析并执行公式
   */
  private async executeFormulaFromNatural(intent: Intent): Promise<string> {
    const { cell, expression } = intent;
    
    if (!cell || !expression) {
      return "❌ 无法解析单元格地址或表达式";
    }

    // 尝试理解表达式
    let formula = expression;
    
    // 处理常见的自然语言模式
    // "E2 减 F2 再减 J2" -> "E2-F2-J2"
    formula = formula
      .replace(/\s*减去?\s*/g, "-")
      .replace(/\s*加上?\s*/g, "+")
      .replace(/\s*乘以?\s*/g, "*")
      .replace(/\s*除以?\s*/g, "/")
      .replace(/的和|求和/g, "SUM")
      .replace(/平均/g, "AVERAGE")
      .replace(/最大/g, "MAX")
      .replace(/最小/g, "MIN");

    this.addMessage("action", `理解为：${cell} = ${formula}`);
    
    return await this.executeFormula(cell, formula);
  }

  /**
   * 执行计算操作
   */
  private async executeCalculation(intent: Intent): Promise<string> {
    const { operation, targetCell, sourceRange, rawInput } = intent;

    // 尝试从输入中提取信息
    let cell = targetCell;
    let range = sourceRange;

    if (!cell || !range) {
      // 尝试智能推断
      const suggestion = this.suggestCalculation(operation!, rawInput!);
      return suggestion;
    }

    // 构建公式
    let formula = "";
    switch (operation) {
      case "sum":
        formula = `SUM(${range})`;
        break;
      case "average":
        formula = `AVERAGE(${range})`;
        break;
      case "max":
        formula = `MAX(${range})`;
        break;
      case "min":
        formula = `MIN(${range})`;
        break;
      default:
        return `❌ 不支持的操作：${operation}`;
    }

    this.addMessage("action", `理解为：在 ${cell} 计算 ${formula}`);
    return await this.executeFormula(cell, formula);
  }

  /**
   * 智能建议计算公式
   */
  private suggestCalculation(operation: string, input: string): string {
    const opName = {
      sum: "求和",
      average: "平均值",
      max: "最大值",
      min: "最小值",
    }[operation] || operation;

    return `我理解你想要计算${opName}。

请告诉我：
1. 要把结果放在哪个单元格？（如 K2）
2. 要计算哪个范围的数据？（如 E2:J2）

例如：
• 在 K2 求和 E2:J2
• K2=SUM(E2:J2)`;
  }

  /**
   * 执行分析
   */
  private async executeAnalysis(intent: Intent): Promise<string> {
    this.addMessage("action", "重新分析工作表...");
    await this.observeAndAnalyze();
    this.setState("completed");
    return "✅ 已完成分析，请查看上面的洞察和建议。";
  }

  /**
   * 执行格式化
   */
  private async executeFormat(intent: Intent): Promise<string> {
    this.setState("idle");
    return `格式化功能正在开发中。

你可以尝试：
• 选中单元格后点击 Excel 的格式化按钮
• 或者告诉我具体要设置什么格式`;
  }

  /**
   * 执行数据操作
   */
  private async executeDataOperation(intent: Intent): Promise<string> {
    this.setState("idle");
    return `数据操作功能正在开发中。

你可以尝试：
• 使用 Excel 的编辑功能
• 或者告诉我具体要做什么操作`;
  }

  /**
   * 处理未知请求
   */
  private handleUnknownRequest(input: string): string {
    this.setState("idle");
    
    return `我还在学习理解 "${input}"。

目前我可以帮你：
• 设置公式：K2=E2-F2-J2-J2
• 自然语言：帮我在 K2 计算 E2 减 F2
• 求和：在 K2 求和 E2:J2
• 平均值：K2 的平均值是 E2:J2
• 分析：重新分析这个表格

你想做什么？`;
  }

  private handleUnknownIntent(input: string): string {
    return `我不太确定你的意思。你可以说：
• "全部做" - 执行所有建议
• "只做格式" - 只执行格式相关的建议
• "取消" - 不做任何操作
• 或者直接告诉我你想做什么`;
  }

  private async executeSuggestionActions(suggestion: Suggestion): Promise<{ success: boolean; error?: string }> {
    // 这里需要对接实际的工具执行
    // 目前返回模拟结果
    for (const action of suggestion.actions) {
      const tool = this.toolRegistry.get(action.type);
      if (tool) {
        try {
          const result = await tool.execute(action.parameters || {});
          if (!result.success) {
            return { success: false, error: result.error };
          }
        } catch (error) {
          return { success: false, error: error instanceof Error ? error.message : "执行失败" };
        }
      } else {
        // 工具不存在，跳过但记录
        this.log("warn", `工具不存在: ${action.type}`);
      }
    }

    return { success: true };
  }

  private handleError(error: unknown): void {
    const message = error instanceof Error ? error.message : String(error);
    this.log("error", `Agent 错误: ${message}`);
    this.addMessage("error", `出错了: ${message}`);
    this.setState("idle");
    this.emit("error", { error });
  }

  private log(level: "info" | "warn" | "error", message: string): void {
    if (this.config.verboseLogging) {
      this.tracer.log(level, `[ProactiveAgent] ${message}`);
    }
  }

  private emit(event: AgentEventType, data: unknown): void {
    for (const handler of this.eventHandlers) {
      try {
        handler(event, data);
      } catch (error) {
        console.error(`Event handler error for ${event}:`, error);
      }
    }
  }
}

// ========== 导出工厂函数 ==========

export function createProactiveAgent(
  toolRegistry: ToolRegistry,
  options?: {
    config?: Partial<ProactiveAgentConfig>;
    preferences?: Partial<UserPreferences>;
  }
): ProactiveAgent {
  return new ProactiveAgent(toolRegistry, options);
}
