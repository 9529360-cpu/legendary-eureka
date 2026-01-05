/**
 * StepReflector - 步骤反思器 v2.9.58
 *
 * P0 核心组件：让 LLM 参与每一步的评估，而非仅在开始时规划
 *
 * 核心职责：
 * 1. 每步执行后调用 LLM 评估结果
 * 2. 判断是否需要调整后续计划
 * 3. 发现用户可能遗漏但应该做的事情
 * 4. 决定是继续、调整、还是询问用户
 *
 * 设计理念：
 * - 把"真正的智能"放回执行过程
 * - 不是机械执行，而是边做边想
 * - Agent 应该像人一样：做完一步看看效果，再决定下一步
 */

import ApiService from "../services/ApiService";
import { PlanStep, ExecutionPlan } from "./TaskPlanner";
import { ToolResult } from "./AgentCore";

// ========== 类型定义 ==========

/**
 * 反思结果
 */
export interface ReflectionResult {
  /** 反思行为建议 */
  action: ReflectionAction;
  /** 置信度 (0-1) */
  confidence: number;
  /** 反思分析 */
  analysis: string;
  /** 发现的问题（如果有） */
  issues?: ReflectionIssue[];
  /** 建议的调整（如果需要） */
  adjustments?: PlanAdjustment[];
  /** 发现的额外机会（用户没说但可能需要的） */
  opportunities?: Opportunity[];
  /** 需要询问用户的问题（如果 action 是 ask_user） */
  questionForUser?: string;
}

/**
 * 反思后的行为建议
 */
export type ReflectionAction =
  | "continue" // 继续执行下一步
  | "adjust_plan" // 调整后续计划
  | "ask_user" // 暂停，询问用户
  | "abort" // 中止执行（发现严重问题）
  | "skip_remaining"; // 跳过剩余步骤（已达成目标）

/**
 * 反思发现的问题
 */
export interface ReflectionIssue {
  /** 问题类型 */
  type: "semantic_mismatch" | "data_anomaly" | "unexpected_result" | "partial_success" | "warning";
  /** 问题描述 */
  description: string;
  /** 严重程度 */
  severity: "info" | "warning" | "error";
  /** 相关的步骤 */
  relatedStepId?: string;
}

/**
 * 计划调整建议
 */
export interface PlanAdjustment {
  /** 调整类型 */
  type: "modify_step" | "add_step" | "remove_step" | "reorder";
  /** 目标步骤 ID */
  targetStepId?: string;
  /** 调整描述 */
  description: string;
  /** 新的参数（如果是 modify_step） */
  newParameters?: Record<string, unknown>;
  /** 新步骤（如果是 add_step） */
  newStep?: Partial<PlanStep>;
}

/**
 * 发现的额外机会
 */
export interface Opportunity {
  /** 机会描述 */
  description: string;
  /** 建议的操作 */
  suggestedAction: string;
  /** 优先级 */
  priority: "low" | "medium" | "high";
  /** 是否需要用户确认 */
  requiresConfirmation: boolean;
}

/**
 * 反思上下文
 */
export interface ReflectionContext {
  /** 原始用户请求 */
  userRequest: string;
  /** 当前执行计划 */
  plan: ExecutionPlan;
  /** 刚完成的步骤 */
  completedStep: PlanStep;
  /** 步骤执行结果 */
  stepResult: ToolResult;
  /** 已完成的步骤列表 */
  completedSteps: PlanStep[];
  /** 剩余的步骤列表 */
  remainingSteps: PlanStep[];
  /** 累积的执行结果 */
  accumulatedResults: string[];
  /** 当前步骤索引 */
  stepIndex: number;
  /** 总步骤数 */
  totalSteps: number;
}

/**
 * 反思配置
 */
export interface ReflectionConfig {
  /** 是否启用反思（默认 true） */
  enabled: boolean;
  /** 反思频率：每 N 步反思一次（默认 1，即每步都反思） */
  frequency: number;
  /** 置信度阈值：低于此值时触发询问用户（默认 0.6） */
  confidenceThreshold: number;
  /** 是否允许调整计划（默认 true） */
  allowPlanAdjustment: boolean;
  /** 是否发现额外机会（默认 true） */
  discoverOpportunities: boolean;
  /** 最大反思时间（毫秒，默认 5000） */
  maxReflectionTime: number;
  /** 只对写操作反思（默认 false） */
  onlyReflectOnWrites: boolean;
}

/**
 * 默认反思配置
 */
export const DEFAULT_REFLECTION_CONFIG: ReflectionConfig = {
  enabled: true,
  frequency: 1, // 每步都反思
  confidenceThreshold: 0.6,
  allowPlanAdjustment: true,
  discoverOpportunities: true,
  maxReflectionTime: 5000,
  onlyReflectOnWrites: false,
};

// ========== StepReflector 类 ==========

/**
 * 步骤反思器
 */
export class StepReflector {
  private config: ReflectionConfig;
  private reflectionCount: number = 0;

  constructor(config: Partial<ReflectionConfig> = {}) {
    this.config = { ...DEFAULT_REFLECTION_CONFIG, ...config };
  }

  /**
   * 对一个步骤的执行结果进行反思
   */
  async reflect(context: ReflectionContext): Promise<ReflectionResult> {
    this.reflectionCount++;

    // 检查是否应该跳过反思
    if (!this.shouldReflect(context)) {
      return this.createContinueResult("跳过反思（不满足反思条件）");
    }

    console.log(
      `[StepReflector] 🤔 反思步骤 ${context.stepIndex + 1}/${context.totalSteps}: ${context.completedStep.description}`
    );

    try {
      // 构建反思 prompt
      const prompt = this.buildReflectionPrompt(context);

      // 调用 LLM 进行反思
      const response = await Promise.race([
        this.callLLMForReflection(prompt),
        this.timeoutPromise(this.config.maxReflectionTime),
      ]);

      if (!response) {
        console.warn("[StepReflector] ⏱️ 反思超时，继续执行");
        return this.createContinueResult("反思超时");
      }

      // 解析 LLM 响应
      const result = this.parseReflectionResponse(response, context);

      console.log(
        `[StepReflector] 💡 反思结果: ${result.action} (置信度: ${result.confidence.toFixed(2)})`
      );

      // 如果置信度太低，建议询问用户
      if (result.confidence < this.config.confidenceThreshold && result.action === "continue") {
        console.log("[StepReflector] ⚠️ 置信度过低，建议询问用户");
        return {
          ...result,
          action: "ask_user",
          questionForUser: this.generateUserQuestion(context, result),
        };
      }

      return result;
    } catch (error) {
      console.error("[StepReflector] ❌ 反思失败:", error);
      // 反思失败不应阻断执行
      return this.createContinueResult(`反思失败: ${String(error)}`);
    }
  }

  /**
   * 检查是否应该进行反思
   */
  private shouldReflect(context: ReflectionContext): boolean {
    if (!this.config.enabled) {
      return false;
    }

    // 检查频率
    if (this.reflectionCount % this.config.frequency !== 0) {
      return false;
    }

    // 如果只对写操作反思
    if (this.config.onlyReflectOnWrites && !context.completedStep.isWriteOperation) {
      return false;
    }

    // 最后一步总是需要反思
    if (context.stepIndex === context.totalSteps - 1) {
      return true;
    }

    return true;
  }

  /**
   * 构建反思 prompt
   */
  private buildReflectionPrompt(context: ReflectionContext): string {
    const {
      userRequest,
      completedStep,
      stepResult,
      completedSteps,
      remainingSteps,
      stepIndex,
      totalSteps,
    } = context;

    // 构建已完成步骤的摘要
    const completedSummary = completedSteps
      .map((s, i) => `${i + 1}. ${s.description} → ${s.status === "completed" ? "✓" : "✗"}`)
      .join("\n");

    // 构建剩余步骤
    const remainingSummary =
      remainingSteps.length > 0
        ? remainingSteps.map((s, i) => `${stepIndex + 2 + i}. ${s.description}`).join("\n")
        : "（无）";

    // 步骤结果
    const resultSummary = stepResult.success
      ? `成功: ${String(stepResult.output).substring(0, 500)}`
      : `失败: ${stepResult.error || "未知错误"}`;

    return `你是一个智能助手的"反思模块"。你的任务是评估刚才执行的操作，判断是否符合用户意图。

## 用户原始请求
${userRequest}

## 执行进度
当前: 步骤 ${stepIndex + 1}/${totalSteps}

## 已完成的步骤
${completedSummary}

## 刚完成的步骤
操作: ${completedStep.action}
描述: ${completedStep.description}
参数: ${JSON.stringify(completedStep.parameters, null, 2)}
结果: ${resultSummary}

## 剩余步骤
${remainingSummary}

## 你需要做的
1. 评估刚才的步骤是否正确完成了它应该做的事
2. 评估结果是否符合用户的真实意图
3. 检查是否有异常或意外情况
4. 判断后续计划是否需要调整
5. 发现用户可能遗漏但应该做的事情

## 输出格式（JSON）
{
  "action": "continue" | "adjust_plan" | "ask_user" | "abort" | "skip_remaining",
  "confidence": 0.0-1.0,
  "analysis": "你的分析（1-2句话）",
  "issues": [
    {
      "type": "semantic_mismatch" | "data_anomaly" | "unexpected_result" | "partial_success" | "warning",
      "description": "问题描述",
      "severity": "info" | "warning" | "error"
    }
  ],
  "adjustments": [
    {
      "type": "modify_step" | "add_step" | "remove_step",
      "description": "调整描述"
    }
  ],
  "opportunities": [
    {
      "description": "发现的额外机会",
      "suggestedAction": "建议的操作",
      "priority": "low" | "medium" | "high",
      "requiresConfirmation": true | false
    }
  ],
  "questionForUser": "如果需要询问用户，问题是什么"
}

请只输出 JSON，不要其他内容。`;
  }

  /**
   * 调用 LLM 进行反思
   */
  private async callLLMForReflection(prompt: string): Promise<string | null> {
    try {
      const response = await ApiService.sendAgentRequest({
        message: prompt,
        systemPrompt: "你是一个智能助手的反思模块。评估执行结果，给出简洁的 JSON 格式反馈。",
        responseFormat: "json",
      });

      return response.message || null;
    } catch (error) {
      console.error("[StepReflector] LLM 调用失败:", error);
      return null;
    }
  }

  /**
   * 解析 LLM 响应
   */
  private parseReflectionResponse(response: string, _context: ReflectionContext): ReflectionResult {
    try {
      // 尝试提取 JSON
      let jsonStr = response;
      const jsonMatch = response.match(/\{[\s\S]*\}/);
      if (jsonMatch) {
        jsonStr = jsonMatch[0];
      }

      const parsed = JSON.parse(jsonStr);

      // 验证必要字段
      const action = this.validateAction(parsed.action);
      const confidence = this.validateConfidence(parsed.confidence);

      return {
        action,
        confidence,
        analysis: parsed.analysis || "未提供分析",
        issues: this.validateIssues(parsed.issues),
        adjustments: this.validateAdjustments(parsed.adjustments),
        opportunities: this.validateOpportunities(parsed.opportunities),
        questionForUser: parsed.questionForUser,
      };
    } catch (error) {
      console.warn("[StepReflector] 解析反思响应失败:", error);
      return this.createContinueResult("响应解析失败，默认继续");
    }
  }

  /**
   * 验证 action 字段
   */
  private validateAction(action: unknown): ReflectionAction {
    const validActions: ReflectionAction[] = [
      "continue",
      "adjust_plan",
      "ask_user",
      "abort",
      "skip_remaining",
    ];
    if (typeof action === "string" && validActions.includes(action as ReflectionAction)) {
      return action as ReflectionAction;
    }
    return "continue";
  }

  /**
   * 验证置信度
   */
  private validateConfidence(confidence: unknown): number {
    if (typeof confidence === "number" && confidence >= 0 && confidence <= 1) {
      return confidence;
    }
    return 0.8; // 默认较高置信度
  }

  /**
   * 验证问题列表
   */
  private validateIssues(issues: unknown): ReflectionIssue[] | undefined {
    if (!Array.isArray(issues)) {
      return undefined;
    }
    return issues
      .filter(
        (issue): issue is ReflectionIssue =>
          typeof issue === "object" &&
          issue !== null &&
          typeof (issue as Record<string, unknown>).description === "string"
      )
      .slice(0, 5); // 最多 5 个问题
  }

  /**
   * 验证调整建议
   */
  private validateAdjustments(adjustments: unknown): PlanAdjustment[] | undefined {
    if (!Array.isArray(adjustments)) {
      return undefined;
    }
    return adjustments
      .filter(
        (adj): adj is PlanAdjustment =>
          typeof adj === "object" &&
          adj !== null &&
          typeof (adj as Record<string, unknown>).description === "string"
      )
      .slice(0, 3); // 最多 3 个调整
  }

  /**
   * 验证机会列表
   */
  private validateOpportunities(opportunities: unknown): Opportunity[] | undefined {
    if (!Array.isArray(opportunities)) {
      return undefined;
    }
    return opportunities
      .filter(
        (opp): opp is Opportunity =>
          typeof opp === "object" &&
          opp !== null &&
          typeof (opp as Record<string, unknown>).description === "string"
      )
      .slice(0, 3); // 最多 3 个机会
  }

  /**
   * 创建默认的"继续"结果
   */
  private createContinueResult(reason: string): ReflectionResult {
    return {
      action: "continue",
      confidence: 0.9,
      analysis: reason,
    };
  }

  /**
   * 生成询问用户的问题
   */
  private generateUserQuestion(context: ReflectionContext, result: ReflectionResult): string {
    const parts: string[] = [];

    parts.push("🤔 我刚完成了以下操作，想确认一下：");
    parts.push(`• ${context.completedStep.description}`);
    parts.push("");

    if (result.issues && result.issues.length > 0) {
      parts.push("我注意到一些情况：");
      for (const issue of result.issues.slice(0, 2)) {
        parts.push(`• ${issue.description}`);
      }
      parts.push("");
    }

    if (context.remainingSteps.length > 0) {
      parts.push(`接下来还有 ${context.remainingSteps.length} 个步骤要执行。`);
    }

    parts.push("请问：继续执行吗？还是需要调整？");

    return parts.join("\n");
  }

  /**
   * 超时 Promise
   */
  private timeoutPromise(ms: number): Promise<null> {
    return new Promise((resolve) => setTimeout(() => resolve(null), ms));
  }

  /**
   * 快速反思（不调用 LLM，基于规则）
   */
  quickReflect(context: ReflectionContext): ReflectionResult {
    const { stepResult, completedStep, remainingSteps } = context;

    // 规则 1: 步骤失败
    if (!stepResult.success) {
      return {
        action: "ask_user",
        confidence: 0.9,
        analysis: "步骤执行失败",
        issues: [
          {
            type: "unexpected_result",
            description: stepResult.error || "执行失败",
            severity: "error",
          },
        ],
        questionForUser: `操作 "${completedStep.description}" 失败了：${stepResult.error}\n\n要重试吗？还是跳过这一步继续？`,
      };
    }

    // 规则 2: 写操作返回空结果
    if (completedStep.isWriteOperation && (!stepResult.output || stepResult.output === "")) {
      return {
        action: "ask_user",
        confidence: 0.6,
        analysis: "写操作没有返回预期确认",
        issues: [
          {
            type: "partial_success",
            description: "操作可能未完全成功",
            severity: "warning",
          },
        ],
        questionForUser: `操作 "${completedStep.description}" 已执行，但没有返回确认信息。请检查结果是否正确？`,
      };
    }

    // 规则 3: 没有剩余步骤，任务完成
    if (remainingSteps.length === 0) {
      return {
        action: "skip_remaining",
        confidence: 0.95,
        analysis: "所有步骤已完成",
      };
    }

    // 默认: 继续
    return {
      action: "continue",
      confidence: 0.85,
      analysis: "步骤成功，继续下一步",
    };
  }

  /**
   * 重置反思计数
   */
  reset(): void {
    this.reflectionCount = 0;
  }

  /**
   * 更新配置
   */
  updateConfig(config: Partial<ReflectionConfig>): void {
    this.config = { ...this.config, ...config };
  }
}

// ========== 单例导出 ==========

export const stepReflector = new StepReflector();

export default StepReflector;
