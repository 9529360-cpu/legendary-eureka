/**
 * ValidationSignal - 验证信号系统 v2.9.58
 *
 * P1 核心组件：将验证结果作为信号而非硬中断
 *
 * 核心职责：
 * 1. 封装验证失败为信号而非立即中止
 * 2. 让 Agent 决定如何处理验证问题
 * 3. 支持多种处理策略：回滚、修复、询问、忽略
 * 4. 追踪验证问题的历史和解决状态
 *
 * 设计理念：
 * - 验证是反馈，不是惩罚
 * - Agent 应该有权决定如何处理验证问题
 * - 有些验证失败是可以接受的（用户知情下）
 * - 把选择权交给 Agent 和用户，而非代码硬编码
 */

import { ValidationCheckResult } from "./AgentCore";

// ========== 类型定义 ==========

/**
 * 验证信号 - 验证失败后产生的信号
 */
export interface ValidationSignal {
  /** 信号 ID */
  id: string;
  /** 信号类型 */
  type: ValidationSignalType;
  /** 来源规则 */
  sourceRule: {
    id: string;
    name: string;
    severity: "block" | "warn";
  };
  /** 验证结果 */
  checkResult: ValidationCheckResult;
  /** 上下文信息 */
  context: SignalContext;
  /** 创建时间 */
  createdAt: Date;
  /** 建议的处理方式 */
  suggestedActions: SuggestedAction[];
  /** 处理状态 */
  status: SignalStatus;
  /** 处理结果 */
  resolution?: SignalResolution;
}

/**
 * 信号类型
 */
export type ValidationSignalType =
  | "data_integrity" // 数据完整性问题（如覆盖了公式）
  | "semantic_error" // 语义错误（如硬编码汇总值）
  | "structural_issue" // 结构问题（如破坏了表结构）
  | "reference_error" // 引用错误（如引用了不存在的单元格）
  | "quality_warning"; // 质量警告（不阻断，但需注意）

/**
 * 信号上下文
 */
export interface SignalContext {
  /** 触发的工具 */
  toolName: string;
  /** 工具输入 */
  toolInput: Record<string, unknown>;
  /** 工具输出 */
  toolOutput?: string;
  /** 影响的范围 */
  affectedRange?: string;
  /** 影响的工作表 */
  affectedSheet?: string;
  /** 操作前的数据（用于回滚） */
  previousData?: unknown;
  /** 操作前的公式（用于回滚） */
  previousFormulas?: string[][];
}

/**
 * 建议的处理方式
 */
export interface SuggestedAction {
  /** 动作类型 */
  type: ActionType;
  /** 动作描述 */
  description: string;
  /** 是否需要用户确认 */
  requiresConfirmation: boolean;
  /** 预估影响 */
  impact: "none" | "low" | "medium" | "high";
  /** 动作参数（如果需要） */
  parameters?: Record<string, unknown>;
}

/**
 * 动作类型
 */
export type ActionType =
  | "rollback" // 回滚操作
  | "fix_and_retry" // 修复后重试
  | "ask_user" // 询问用户
  | "ignore_once" // 本次忽略
  | "ignore_rule" // 永久忽略该规则
  | "abort_task"; // 中止任务

/**
 * 信号状态
 */
export type SignalStatus =
  | "pending" // 等待处理
  | "processing" // 正在处理
  | "resolved" // 已解决
  | "ignored" // 已忽略
  | "escalated"; // 已上报用户

/**
 * 信号解决结果
 */
export interface SignalResolution {
  /** 采取的动作 */
  action: ActionType;
  /** 是否成功 */
  success: boolean;
  /** 处理说明 */
  description: string;
  /** 处理时间 */
  resolvedAt: Date;
  /** 是否需要用户确认 */
  userConfirmed?: boolean;
}

/**
 * 信号处理决策
 */
export interface SignalDecision {
  /** 决定采取的动作 */
  action: ActionType;
  /** 决策理由 */
  reasoning: string;
  /** 置信度 (0-1) */
  confidence: number;
  /** 是否需要用户确认 */
  needsUserConfirmation: boolean;
  /** 给用户的消息（如果需要确认） */
  userMessage?: string;
}

/**
 * 验证信号配置
 */
export interface ValidationSignalConfig {
  /** 是否启用信号模式（默认 true） */
  enabled: boolean;
  /** 自动回滚阈值：连续失败多少次后自动回滚（默认 3） */
  autoRollbackThreshold: number;
  /** 是否允许忽略阻断级别的验证（默认 false） */
  allowIgnoreBlocking: boolean;
  /** 用户确认超时（毫秒，默认 30000） */
  userConfirmationTimeout: number;
  /** 是否记录被忽略的验证（用于审计）（默认 true） */
  logIgnoredValidations: boolean;
}

/**
 * 默认配置
 */
export const DEFAULT_SIGNAL_CONFIG: ValidationSignalConfig = {
  enabled: true,
  autoRollbackThreshold: 3,
  allowIgnoreBlocking: false,
  userConfirmationTimeout: 30000,
  logIgnoredValidations: true,
};

// ========== ValidationSignalHandler 类 ==========

/**
 * 验证信号处理器
 */
export class ValidationSignalHandler {
  private config: ValidationSignalConfig;
  private pendingSignals: Map<string, ValidationSignal> = new Map();
  private signalHistory: ValidationSignal[] = [];
  private ignoredRules: Set<string> = new Set();

  constructor(config: Partial<ValidationSignalConfig> = {}) {
    this.config = { ...DEFAULT_SIGNAL_CONFIG, ...config };
  }

  /**
   * 从验证结果创建信号
   */
  createSignal(
    rule: { id: string; name: string; severity: "block" | "warn" },
    checkResult: ValidationCheckResult,
    context: SignalContext
  ): ValidationSignal {
    const signalId = `sig_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
    const signalType = this.classifySignalType(rule, checkResult);

    const signal: ValidationSignal = {
      id: signalId,
      type: signalType,
      sourceRule: rule,
      checkResult,
      context,
      createdAt: new Date(),
      suggestedActions: this.generateSuggestedActions(rule, checkResult, signalType),
      status: "pending",
    };

    this.pendingSignals.set(signalId, signal);
    return signal;
  }

  /**
   * 分类信号类型
   */
  private classifySignalType(
    rule: { id: string; name: string },
    checkResult: ValidationCheckResult
  ): ValidationSignalType {
    const message = checkResult.message.toLowerCase();

    if (message.includes("公式") || message.includes("formula")) {
      return "data_integrity";
    }
    if (message.includes("硬编码") || message.includes("hardcode")) {
      return "semantic_error";
    }
    if (message.includes("结构") || message.includes("表头")) {
      return "structural_issue";
    }
    if (message.includes("引用") || message.includes("不存在")) {
      return "reference_error";
    }
    return "quality_warning";
  }

  /**
   * 生成建议的处理方式
   */
  private generateSuggestedActions(
    rule: { id: string; severity: "block" | "warn" },
    checkResult: ValidationCheckResult,
    signalType: ValidationSignalType
  ): SuggestedAction[] {
    const actions: SuggestedAction[] = [];

    // 1. 回滚总是一个选项（对于写操作）
    actions.push({
      type: "rollback",
      description: "回滚此操作，恢复原始数据",
      requiresConfirmation: false,
      impact: "low",
    });

    // 2. 根据类型建议修复
    if (signalType === "data_integrity" || signalType === "semantic_error") {
      actions.push({
        type: "fix_and_retry",
        description: checkResult.suggestedFix || "尝试用正确的方式重新执行",
        requiresConfirmation: false,
        impact: "medium",
      });
    }

    // 3. 询问用户
    actions.push({
      type: "ask_user",
      description: "将问题报告给用户，请求指导",
      requiresConfirmation: true,
      impact: "none",
    });

    // 4. 忽略（仅对警告级别）
    if (rule.severity === "warn") {
      actions.push({
        type: "ignore_once",
        description: "本次忽略此警告，继续执行",
        requiresConfirmation: false,
        impact: "low",
      });
    }

    // 5. 中止（仅对严重问题）
    if (rule.severity === "block" && signalType !== "quality_warning") {
      actions.push({
        type: "abort_task",
        description: "中止任务，避免造成更多问题",
        requiresConfirmation: true,
        impact: "high",
      });
    }

    return actions;
  }

  /**
   * 自动决策：根据规则和历史决定如何处理信号
   */
  autoDecide(signal: ValidationSignal): SignalDecision {
    // 检查是否是被忽略的规则
    if (this.ignoredRules.has(signal.sourceRule.id)) {
      return {
        action: "ignore_once",
        reasoning: "该规则已被用户标记为忽略",
        confidence: 0.95,
        needsUserConfirmation: false,
      };
    }

    // 检查历史中是否有相同问题被成功修复
    const similarSignals = this.signalHistory.filter(
      (s) =>
        s.sourceRule.id === signal.sourceRule.id &&
        s.resolution?.action === "fix_and_retry" &&
        s.resolution?.success
    );

    if (similarSignals.length > 0) {
      // 之前成功修复过，建议再次尝试修复
      return {
        action: "fix_and_retry",
        reasoning: `之前 ${similarSignals.length} 次成功修复过类似问题`,
        confidence: 0.8,
        needsUserConfirmation: false,
      };
    }

    // 检查连续失败次数
    const recentSameRuleSignals = this.pendingSignals.size;
    if (recentSameRuleSignals >= this.config.autoRollbackThreshold) {
      return {
        action: "rollback",
        reasoning: `连续 ${recentSameRuleSignals} 次验证失败，自动回滚`,
        confidence: 0.9,
        needsUserConfirmation: false,
      };
    }

    // 对于阻断级别的信号，建议询问用户
    if (signal.sourceRule.severity === "block") {
      return {
        action: "ask_user",
        reasoning: "这是一个阻断级别的验证问题，需要用户决定",
        confidence: 0.7,
        needsUserConfirmation: true,
        userMessage: this.formatUserMessage(signal),
      };
    }

    // 对于警告级别，可以尝试继续
    return {
      action: "ignore_once",
      reasoning: "这是一个警告级别的问题，可以继续执行",
      confidence: 0.6,
      needsUserConfirmation: false,
    };
  }

  /**
   * 让 LLM 决策如何处理信号
   */
  async llmDecide(signal: ValidationSignal, taskContext: string): Promise<SignalDecision> {
    // 构建决策 prompt
    const prompt = this.buildDecisionPrompt(signal, taskContext);

    // 这里应该调用 LLM，但为了保持简单，先返回自动决策
    // 实际使用时可以调用 ApiService.sendAgentRequest
    console.log("[ValidationSignalHandler] LLM 决策 prompt:", prompt.substring(0, 200));

    // 暂时返回自动决策
    return this.autoDecide(signal);
  }

  /**
   * 构建 LLM 决策 prompt
   */
  private buildDecisionPrompt(signal: ValidationSignal, taskContext: string): string {
    return `你是一个智能助手的决策模块。以下是一个验证问题，请决定如何处理。

## 任务上下文
${taskContext}

## 验证问题
- 规则: ${signal.sourceRule.name}
- 级别: ${signal.sourceRule.severity === "block" ? "阻断" : "警告"}
- 问题: ${signal.checkResult.message}
- 建议: ${signal.checkResult.suggestedFix || "无"}

## 可选处理方式
${signal.suggestedActions.map((a, i) => `${i + 1}. ${a.type}: ${a.description}`).join("\n")}

## 你需要决定
选择一个处理方式，并说明理由。如果需要询问用户，请提供友好的问题。

输出 JSON：
{
  "action": "rollback" | "fix_and_retry" | "ask_user" | "ignore_once" | "abort_task",
  "reasoning": "理由",
  "confidence": 0.0-1.0,
  "needsUserConfirmation": true | false,
  "userMessage": "给用户的消息（如果需要确认）"
}`;
  }

  /**
   * 格式化给用户的消息
   */
  formatUserMessage(signal: ValidationSignal): string {
    const parts: string[] = [];

    parts.push("⚠️ 执行过程中发现一个问题：");
    parts.push("");
    parts.push(`**问题**: ${signal.checkResult.message}`);

    if (signal.checkResult.suggestedFix) {
      parts.push(`**建议**: ${signal.checkResult.suggestedFix}`);
    }

    parts.push("");
    parts.push("请选择如何处理：");
    parts.push("1. 回滚 - 撤销这个操作");
    parts.push("2. 忽略 - 继续执行（可能有风险）");
    parts.push("3. 中止 - 停止整个任务");

    return parts.join("\n");
  }

  /**
   * 解决信号
   */
  resolveSignal(
    signalId: string,
    action: ActionType,
    success: boolean,
    description: string,
    userConfirmed?: boolean
  ): void {
    const signal = this.pendingSignals.get(signalId);
    if (!signal) {
      console.warn(`[ValidationSignalHandler] 信号不存在: ${signalId}`);
      return;
    }

    signal.status = "resolved";
    signal.resolution = {
      action,
      success,
      description,
      resolvedAt: new Date(),
      userConfirmed,
    };

    // 移动到历史
    this.signalHistory.push(signal);
    this.pendingSignals.delete(signalId);

    // 记录日志
    console.log(
      `[ValidationSignalHandler] 信号已解决: ${signalId} -> ${action} (${success ? "成功" : "失败"})`
    );
  }

  /**
   * 忽略规则
   */
  ignoreRule(ruleId: string): void {
    this.ignoredRules.add(ruleId);
    console.log(`[ValidationSignalHandler] 规则已忽略: ${ruleId}`);
  }

  /**
   * 取消忽略规则
   */
  unignoreRule(ruleId: string): void {
    this.ignoredRules.delete(ruleId);
  }

  /**
   * 获取待处理信号数量
   */
  getPendingCount(): number {
    return this.pendingSignals.size;
  }

  /**
   * 获取所有待处理信号
   */
  getPendingSignals(): ValidationSignal[] {
    return Array.from(this.pendingSignals.values());
  }

  /**
   * 清空待处理信号
   */
  clearPending(): void {
    this.pendingSignals.clear();
  }

  /**
   * 重置状态
   */
  reset(): void {
    this.pendingSignals.clear();
    this.signalHistory = [];
    // 注意：不清除 ignoredRules，保持用户偏好
  }
}

// ========== 单例导出 ==========

export const validationSignalHandler = new ValidationSignalHandler();

export default ValidationSignalHandler;
