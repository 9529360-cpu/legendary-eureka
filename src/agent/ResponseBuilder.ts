/**
 * ResponseBuilder.ts - 响应构建器（P3 响应结构）
 *
 * v2.9.59: LLM 原话必须保留，模板是补充不是替代
 *
 * 核心原则：
 * ┌─────────────────────────────────────────────────────┐
 * │  mainMessage = LLM 原话                            │
 * │  templateMessage = 模板补充（可选）                 │
 * │  最终展示 = mainMessage + templateMessage          │
 * └─────────────────────────────────────────────────────┘
 */

import { AgentReply, AgentReplyDebug, Signal, StepDecision } from "./AgentProtocol";
import ResponseGenerator, { ResponseContext, ExecutionState } from "./ResponseTemplates";
import ApiService from "../services/ApiService";

// ========== 配置 ==========

export interface ResponseBuilderConfig {
  /** 是否使用模板 */
  useTemplate: boolean;
  /** 是否添加建议 */
  addSuggestions: boolean;
  /** 是否包含调试信息 */
  includeDebug: boolean;
  /** LLM 生成失败时回退到纯模板 */
  fallbackToTemplate: boolean;
}

export const DEFAULT_RESPONSE_CONFIG: ResponseBuilderConfig = {
  useTemplate: true,
  addSuggestions: true,
  includeDebug: false,
  fallbackToTemplate: true,
};

// ========== 构建上下文 ==========

export interface BuildContext {
  /** 用户原始请求 */
  userRequest: string;
  /** 执行状态 */
  executionState: ExecutionState;
  /** 执行摘要（给 LLM 用） */
  executionSummary: string;
  /** 响应模板上下文 */
  templateContext?: ResponseContext;
  /** 信号列表 */
  signals?: Signal[];
  /** 最终决策 */
  decision?: StepDecision;
  /** 步骤 ID */
  stepId?: string;
}

// ========== 核心类 ==========

/**
 * 响应构建器
 *
 * 负责组装最终回复：mainMessage + templateMessage + suggestion
 */
export class ResponseBuilder {
  private config: ResponseBuilderConfig;

  constructor(config?: Partial<ResponseBuilderConfig>) {
    this.config = { ...DEFAULT_RESPONSE_CONFIG, ...config };
  }

  /**
   * 构建完整回复
   */
  async build(context: BuildContext): Promise<AgentReply> {
    // 1. 先让 LLM 生成 mainMessage
    let mainMessage: string;
    try {
      mainMessage = await this.generateMainMessage(context);
    } catch (error) {
      console.warn("[ResponseBuilder] LLM 生成失败:", error);
      if (this.config.fallbackToTemplate && context.templateContext) {
        mainMessage = ResponseGenerator.generate(context.templateContext);
      } else {
        mainMessage = this.generateFallbackMessage(context);
      }
    }

    // 2. 生成 templateMessage（如果启用）
    let templateMessage: string | undefined;
    if (this.config.useTemplate && context.templateContext) {
      templateMessage = this.generateTemplateMessage(context, mainMessage);
    }

    // 3. 生成 suggestionMessage（如果启用）
    let suggestionMessage: string | undefined;
    if (this.config.addSuggestions) {
      suggestionMessage = this.generateSuggestion(context);
    }

    // 4. 组装调试信息
    let debug: AgentReplyDebug | undefined;
    if (this.config.includeDebug) {
      debug = {
        signals: context.signals,
        decision: context.decision,
        stepId: context.stepId,
        executionState: context.executionState,
      };
    }

    return {
      mainMessage,
      templateMessage,
      suggestionMessage,
      debug,
    };
  }

  /**
   * 快速构建（不调用 LLM，只用模板）
   */
  buildSync(context: BuildContext): AgentReply {
    let mainMessage: string;
    if (context.templateContext) {
      mainMessage = ResponseGenerator.generate(context.templateContext);
    } else {
      mainMessage = this.generateFallbackMessage(context);
    }

    return {
      mainMessage,
      templateMessage: undefined,
      suggestionMessage: this.config.addSuggestions ? this.generateSuggestion(context) : undefined,
      debug: this.config.includeDebug
        ? {
            signals: context.signals,
            decision: context.decision,
            stepId: context.stepId,
            executionState: context.executionState,
          }
        : undefined,
    };
  }

  // ========== LLM 生成 ==========

  /**
   * 调用 LLM 生成 mainMessage
   */
  private async generateMainMessage(context: BuildContext): Promise<string> {
    const systemPrompt = this.buildLLMSystemPrompt(context.executionState);
    const userPrompt = this.buildLLMUserPrompt(context);

    const response = await ApiService.sendChatRequest({
      message: userPrompt,
      systemPrompt,
      responseFormat: "text",
    });

    if (response.success && response.message) {
      const cleaned = this.cleanLLMResponse(response.message, context.executionState);
      return cleaned;
    }

    throw new Error(response.message || "LLM 响应为空");
  }

  /**
   * 构建 LLM 系统提示词
   */
  private buildLLMSystemPrompt(state: ExecutionState): string {
    const stateConstraint = this.getStateConstraint(state);

    return `你是 Excel 智能助手。用户刚刚请求了一个操作，你需要用自然、友好的方式告诉用户结果。

## 硬性约束（必须遵守）

${stateConstraint}

## 风格要求

- 简洁：一般 1-3 句话
- 自然：像人说话，不要生硬
- 具体：如果有具体数据（行数、范围等），可以提及
- 情感适度：成功时可以表达肯定，失败时要诚恳`;
  }

  /**
   * 构建 LLM 用户提示词
   */
  private buildLLMUserPrompt(context: BuildContext): string {
    let prompt = `用户请求：${context.userRequest}\n\n`;
    prompt += `执行结果：${context.executionSummary}\n\n`;

    if (context.signals && context.signals.length > 0) {
      prompt += `注意事项：\n`;
      for (const signal of context.signals.slice(0, 3)) {
        prompt += `- ${signal.message}\n`;
      }
      prompt += "\n";
    }

    prompt += "请用 1-3 句话自然地告诉用户结果。";

    return prompt;
  }

  /**
   * 获取状态约束
   */
  private getStateConstraint(state: ExecutionState): string {
    const constraints: Record<ExecutionState, string> = {
      planned: '【状态：规划中】你只能说"我打算..."、"计划..."，不能说"已完成"、"已执行"。',
      preview: '【状态：预览中】你只能说"准备..."、"即将..."，不能说"已完成"、"已执行"。',
      executing: '【状态：执行中】你只能说"正在..."、"处理中..."，不能说"已完成"。',
      executed: '【状态：已执行】操作已完成，你可以说"已完成"、"搞定了"等。',
      failed: '【状态：失败】操作失败了，你必须说明失败，不能说"已完成"。要诚恳道歉。',
      partial: "【状态：部分成功】有些操作成功，有些失败。你要说明哪些成功哪些失败。",
      rolled_back: '【状态：已回滚】操作被撤销了，你要说明"已撤销"或"已回滚"。',
    };
    return constraints[state] || "状态未知，请谨慎表达。";
  }

  /**
   * 清理 LLM 响应（验证状态约束）
   */
  private cleanLLMResponse(response: string, state: ExecutionState): string {
    const trimmed = response.trim();

    // 检查是否违反状态约束
    if (this.violatesStateConstraint(trimmed, state)) {
      console.warn("[ResponseBuilder] LLM 响应违反状态约束，需要修正");
      return this.generateFallbackMessage({ executionState: state } as BuildContext);
    }

    return trimmed;
  }

  /**
   * 检查是否违反状态约束
   */
  private violatesStateConstraint(response: string, state: ExecutionState): boolean {
    const completionWords = ["已完成", "完成了", "搞定", "成功", "Done", "Completed"];

    const nonCompleteStates: ExecutionState[] = ["planned", "preview", "executing", "failed"];

    if (nonCompleteStates.includes(state)) {
      for (const word of completionWords) {
        if (response.includes(word)) {
          return true;
        }
      }
    }

    // 失败状态必须有失败语气
    if (state === "failed") {
      const failureWords = ["失败", "抱歉", "出错", "没能", "无法", "sorry", "fail"];
      const hasFailureWord = failureWords.some((w) => response.toLowerCase().includes(w));
      if (!hasFailureWord) {
        return true;
      }
    }

    return false;
  }

  // ========== 模板生成 ==========

  /**
   * 生成模板补充消息
   *
   * 只在 LLM 没说清楚时补充
   */
  private generateTemplateMessage(context: BuildContext, mainMessage: string): string | undefined {
    if (!context.templateContext) return undefined;

    // 如果 mainMessage 已经够清楚，不补充
    if (this.isMessageComplete(mainMessage, context)) {
      return undefined;
    }

    // 生成模板消息
    const templateMsg = ResponseGenerator.generate(context.templateContext);

    // 如果模板和 main 太像，不重复
    if (this.isSimilar(mainMessage, templateMsg)) {
      return undefined;
    }

    // 返回简化版（只要关键信息）
    return this.extractKeyInfo(templateMsg);
  }

  /**
   * 判断消息是否足够完整
   */
  private isMessageComplete(message: string, context: BuildContext): boolean {
    // 如果提到了具体范围或数量，认为足够
    if (/[A-Z]+\d+/.test(message)) return true;
    if (/\d+\s*(行|列|个|条)/.test(message)) return true;

    // 如果是失败且说了原因
    if (context.executionState === "failed" && message.length > 20) return true;

    return false;
  }

  /**
   * 判断两个消息是否相似
   */
  private isSimilar(a: string, b: string): boolean {
    const normalize = (s: string) => s.replace(/[^\u4e00-\u9fa5a-zA-Z0-9]/g, "").toLowerCase();
    const na = normalize(a);
    const nb = normalize(b);

    // 简单的 Jaccard 相似度
    const setA = new Set(na.split(""));
    const setB = new Set(nb.split(""));
    const intersection = new Set([...setA].filter((x) => setB.has(x)));
    const union = new Set([...setA, ...setB]);

    return intersection.size / union.size > 0.7;
  }

  /**
   * 提取关键信息
   */
  private extractKeyInfo(templateMsg: string): string {
    // 提取范围信息
    const rangeMatch = templateMsg.match(/[A-Z]+\d+:[A-Z]+\d+/);
    const countMatch = templateMsg.match(/\d+\s*(行|列|个|条|格)/);

    const parts: string[] = [];
    if (rangeMatch) parts.push(`范围: ${rangeMatch[0]}`);
    if (countMatch) parts.push(countMatch[0]);

    return parts.length > 0 ? `(${parts.join(", ")})` : "";
  }

  // ========== 建议生成 ==========

  /**
   * 生成建议消息
   */
  private generateSuggestion(context: BuildContext): string | undefined {
    // 根据状态和信号生成建议
    if (context.executionState === "executed") {
      return undefined; // 成功就不建议了
    }

    if (context.executionState === "failed") {
      return "💡 你可以尝试缩小操作范围，或者检查数据格式后重试。";
    }

    if (context.signals?.some((s) => s.level === "warning")) {
      return "⚠️ 有一些警告，建议检查后再继续。";
    }

    return undefined;
  }

  // ========== 兜底生成 ==========

  /**
   * 生成兜底消息
   */
  private generateFallbackMessage(context: BuildContext): string {
    switch (context.executionState) {
      case "planned":
        return "我已经规划好了，需要我执行吗？";
      case "preview":
        return "准备就绪，点击确认后开始执行。";
      case "executing":
        return "正在处理中...";
      case "executed":
        return "操作完成。";
      case "failed":
        return "抱歉，操作失败了。";
      case "partial":
        return "部分操作完成，有些步骤失败了。";
      case "rolled_back":
        return "操作已撤销。";
      default:
        return "处理完成。";
    }
  }
}

// ========== 单例导出 ==========

export const responseBuilder = new ResponseBuilder();

// ========== 便捷函数 ==========

/**
 * 构建回复
 */
export async function buildReply(context: BuildContext): Promise<AgentReply> {
  return responseBuilder.build(context);
}

/**
 * 同步构建回复（不调用 LLM）
 */
export function buildReplySync(context: BuildContext): AgentReply {
  return responseBuilder.buildSync(context);
}

/**
 * 格式化回复为字符串
 */
export function formatReply(reply: AgentReply): string {
  let result = reply.mainMessage;

  if (reply.templateMessage) {
    result += " " + reply.templateMessage;
  }

  if (reply.suggestionMessage) {
    result += "\n\n" + reply.suggestionMessage;
  }

  return result;
}
