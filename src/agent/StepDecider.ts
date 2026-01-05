/**
 * StepDecider.ts - 步骤决策器（P0 每步反思，协议版）
 *
 * v2.9.59: 使用统一的 StepDecision 和 Signal 类型
 *
 * 与 StepReflector 的区别：
 * - 输入：统一的 Signal[] 而非自定义结构
 * - 输出：StepDecision 5 选 1 而非 ReflectionAction
 * - 逻辑：规则优先 + LLM 兜底
 */

import {
  StepDecision,
  StepFix,
  Signal,
  ClarifyQuestion,
  inferRecommendedAction,
  RecommendedAction,
} from "./AgentProtocol";
import { PlanStep, ExecutionPlan } from "./TaskPlanner";
import { ToolResult } from "./AgentCore";
import ApiService from "../services/ApiService";

// ========== 类型定义 ==========

/**
 * 决策上下文
 */
export interface DecisionContext {
  /** 用户原始请求 */
  userRequest: string;
  /** 执行计划 */
  plan: ExecutionPlan;
  /** 当前步骤 */
  currentStep: PlanStep;
  /** 步骤执行结果 */
  toolResult: ToolResult;
  /** 收集到的信号（来自 P1 validator） */
  signals: Signal[];
  /** 步骤索引 */
  stepIndex: number;
  /** 总步骤数 */
  totalSteps: number;
  /** 已完成步骤的结果摘要 */
  previousResults?: string[];
}

/**
 * 决策配置
 */
export interface DeciderConfig {
  /** 是否启用 LLM 决策（默认 true） */
  useLLM: boolean;
  /** LLM 决策超时（毫秒） */
  llmTimeout: number;
  /** 是否优先使用规则（默认 true） */
  rulesFirst: boolean;
  /** 最大重试次数（fix_and_retry） */
  maxRetries: number;
}

export const DEFAULT_DECIDER_CONFIG: DeciderConfig = {
  useLLM: true,
  llmTimeout: 5000,
  rulesFirst: true,
  maxRetries: 3,
};

// ========== StepDecider 类 ==========

/**
 * 步骤决策器
 *
 * 职责：根据 toolResult + signals 决定下一步动作
 */
export class StepDecider {
  private config: DeciderConfig;
  private retryCount: Map<string, number> = new Map();

  constructor(config?: Partial<DeciderConfig>) {
    this.config = { ...DEFAULT_DECIDER_CONFIG, ...config };
  }

  /**
   * 做出决策
   */
  async decide(context: DecisionContext): Promise<StepDecision> {
    const { currentStep, toolResult: _toolResult, signals, stepIndex, totalSteps } = context;

    console.log(
      `[StepDecider] 🤔 决策步骤 ${stepIndex + 1}/${totalSteps}: ${currentStep.description}`
    );
    console.log(`[StepDecider] 📊 信号数量: ${signals.length}`);

    // ========== 第一层：硬规则（立即返回） ==========
    if (this.config.rulesFirst) {
      const ruleDecision = this.applyRules(context);
      if (ruleDecision) {
        console.log(`[StepDecider] 📋 规则决策: ${ruleDecision.action}`);
        return ruleDecision;
      }
    }

    // ========== 第二层：信号推断 ==========
    const signalDecision = this.inferFromSignals(context);
    if (signalDecision) {
      console.log(`[StepDecider] 📡 信号推断决策: ${signalDecision.action}`);
      return signalDecision;
    }

    // ========== 第三层：LLM 决策 ==========
    if (this.config.useLLM) {
      try {
        const llmDecision = await Promise.race([this.askLLM(context), this.timeoutPromise()]);
        if (llmDecision) {
          console.log(`[StepDecider] 🤖 LLM 决策: ${llmDecision.action}`);
          return llmDecision;
        }
      } catch (error) {
        console.warn("[StepDecider] LLM 决策失败:", error);
      }
    }

    // ========== 兜底：继续 ==========
    console.log("[StepDecider] ✅ 默认决策: continue");
    return { action: "continue" };
  }

  /**
   * 快速决策（不调用 LLM）
   */
  decideSync(context: DecisionContext): StepDecision {
    // 只用规则和信号
    const ruleDecision = this.applyRules(context);
    if (ruleDecision) return ruleDecision;

    const signalDecision = this.inferFromSignals(context);
    if (signalDecision) return signalDecision;

    return { action: "continue" };
  }

  // ========== 规则层 ==========

  /**
   * 应用硬编码规则
   */
  private applyRules(context: DecisionContext): StepDecision | null {
    const { currentStep, toolResult, signals, stepIndex, totalSteps } = context;

    // 规则 1：工具执行失败
    if (!toolResult.success) {
      const stepKey = currentStep.id || `step_${stepIndex}`;
      const retryCount = this.retryCount.get(stepKey) || 0;

      if (retryCount < this.config.maxRetries) {
        this.retryCount.set(stepKey, retryCount + 1);
        return {
          action: "fix_and_retry",
          fix: this.inferFix(currentStep, toolResult, signals),
        };
      } else {
        return {
          action: "rollback_and_replan",
          reason: `步骤 ${currentStep.description} 重试 ${this.config.maxRetries} 次仍失败`,
        };
      }
    }

    // 规则 2：有 critical 信号
    const criticalSignal = signals.find((s) => s.level === "critical");
    if (criticalSignal) {
      return {
        action: "abort",
        reason: criticalSignal.message,
      };
    }

    // 规则 3：有 error 信号且推荐 rollback
    const rollbackSignal = signals.find(
      (s) => s.level === "error" && s.recommended === "rollback_and_replan"
    );
    if (rollbackSignal) {
      return {
        action: "rollback_and_replan",
        reason: rollbackSignal.message,
      };
    }

    // 规则 4：最后一步成功
    if (stepIndex === totalSteps - 1 && toolResult.success) {
      // 继续（让上层处理完成逻辑）
      return { action: "continue" };
    }

    return null;
  }

  // ========== 信号推断层 ==========

  /**
   * 从信号推断决策
   */
  private inferFromSignals(context: DecisionContext): StepDecision | null {
    const { signals } = context;

    if (signals.length === 0) {
      return null;
    }

    // 按信号的 recommended 决定
    const recommended = inferRecommendedAction(signals);

    return this.recommendedToDecision(recommended, signals);
  }

  /**
   * 将推荐动作转换为决策
   */
  private recommendedToDecision(
    recommended: RecommendedAction,
    signals: Signal[]
  ): StepDecision | null {
    switch (recommended) {
      case "continue":
        return { action: "continue" };

      case "fix_and_retry":
        return {
          action: "fix_and_retry",
          fix: undefined, // 由上层提供具体修复
        };

      case "rollback_and_replan":
        return {
          action: "rollback_and_replan",
          reason: signals.find((s) => s.level === "error")?.message || "需要重新规划",
        };

      case "ask_user":
        return {
          action: "ask_user",
          questions: this.signalsToQuestions(signals),
        };

      case "abort":
        return {
          action: "abort",
          reason: signals.find((s) => s.level === "critical")?.message || "发现严重问题",
        };

      default:
        return null;
    }
  }

  /**
   * 将信号转换为澄清问题
   */
  private signalsToQuestions(signals: Signal[]): ClarifyQuestion[] {
    const questions: ClarifyQuestion[] = [];

    for (const signal of signals.filter((s) => s.recommended === "ask_user")) {
      questions.push({
        id: signal.code,
        question: signal.message,
        required: signal.level === "error",
      });
    }

    if (questions.length === 0) {
      questions.push({
        id: "clarify",
        question: "执行过程中遇到一些问题，你希望如何处理？",
        options: ["继续执行", "重试", "停止"],
        required: true,
      });
    }

    return questions;
  }

  // ========== LLM 层 ==========

  /**
   * 询问 LLM
   */
  private async askLLM(context: DecisionContext): Promise<StepDecision | null> {
    const prompt = this.buildLLMPrompt(context);

    const response = await ApiService.sendAgentRequest({
      message: prompt,
      systemPrompt: `你是步骤决策器。根据执行结果和信号，决定下一步动作。
只能返回以下 5 个动作之一（JSON 格式）：
1. {"action": "continue"} - 继续执行下一步
2. {"action": "fix_and_retry", "fix": {"type": "adjust_parameters", "description": "..."}} - 修复后重试
3. {"action": "rollback_and_replan", "reason": "..."} - 回滚并重新规划
4. {"action": "ask_user", "questions": [{"id": "q1", "question": "..."}]} - 询问用户
5. {"action": "abort", "reason": "..."} - 中止执行

只返回 JSON，不要其他内容。`,
      responseFormat: "json",
    });

    if (response.success && response.message) {
      try {
        const parsed = JSON.parse(response.message);
        if (this.isValidDecision(parsed)) {
          return parsed as StepDecision;
        }
      } catch {
        console.warn("[StepDecider] LLM 响应解析失败");
      }
    }

    return null;
  }

  /**
   * 构建 LLM 提示
   */
  private buildLLMPrompt(context: DecisionContext): string {
    const { currentStep, toolResult, signals, userRequest, stepIndex, totalSteps } = context;

    let prompt = `## 当前任务
${userRequest}

## 刚完成的步骤 (${stepIndex + 1}/${totalSteps})
${currentStep.description}
动作: ${currentStep.action}

## 执行结果
成功: ${toolResult.success}
输出: ${typeof toolResult.output === "string" ? toolResult.output.slice(0, 500) : JSON.stringify(toolResult.output).slice(0, 500)}
`;

    if (signals.length > 0) {
      prompt += `\n## 收到的信号\n`;
      for (const signal of signals) {
        prompt += `- [${signal.level}] ${signal.code}: ${signal.message}\n`;
        if (signal.recommended) {
          prompt += `  推荐: ${signal.recommended}\n`;
        }
      }
    }

    prompt += `\n请决定下一步动作。`;

    return prompt;
  }

  /**
   * 验证决策格式
   */
  private isValidDecision(obj: unknown): boolean {
    if (typeof obj !== "object" || obj === null) return false;

    const decision = obj as { action?: string };
    const validActions = ["continue", "fix_and_retry", "rollback_and_replan", "ask_user", "abort"];

    return typeof decision.action === "string" && validActions.includes(decision.action);
  }

  // ========== 辅助方法 ==========

  /**
   * 推断修复方案
   */
  private inferFix(
    step: PlanStep,
    toolResult: ToolResult,
    _signals: Signal[]
  ): StepFix | undefined {
    // 根据错误类型推断修复
    const errorMsg = String(toolResult.error || toolResult.output || "");

    // 范围错误 → 缩小范围
    if (errorMsg.includes("范围") || errorMsg.includes("range")) {
      return {
        type: "shrink_range",
        description: "尝试缩小操作范围",
      };
    }

    // 公式错误 → 调整公式
    if (errorMsg.includes("公式") || errorMsg.includes("formula")) {
      return {
        type: "adjust_formula",
        description: "调整公式参数",
      };
    }

    // 通用 → 调整参数
    return {
      type: "adjust_parameters",
      description: "调整步骤参数后重试",
    };
  }

  /**
   * 超时 Promise
   */
  private timeoutPromise(): Promise<null> {
    return new Promise((resolve) => {
      setTimeout(() => resolve(null), this.config.llmTimeout);
    });
  }

  /**
   * 重置重试计数
   */
  resetRetryCount(): void {
    this.retryCount.clear();
  }
}

// ========== 单例导出 ==========

export const stepDecider = new StepDecider();

/**
 * 便捷函数：做出决策
 */
export async function makeDecision(context: DecisionContext): Promise<StepDecision> {
  return stepDecider.decide(context);
}

/**
 * 便捷函数：同步决策
 */
export function makeDecisionSync(context: DecisionContext): StepDecision {
  return stepDecider.decideSync(context);
}
