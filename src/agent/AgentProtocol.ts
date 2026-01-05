/**
 * AgentProtocol.ts - Agent 系统协议定义（SSOT）
 *
 * v2.9.59: 单一事实来源，避免循环引用
 *
 * 所有模块必须使用这里定义的类型，不要自己发明
 */

import type { ExecutionPlan } from "./TaskPlanner";

// ========== Signal 系统 ==========

/**
 * 信号级别
 */
export type SignalLevel = "info" | "warning" | "error" | "critical";

/**
 * 统一信号类型
 *
 * 所有 validator、reflector、gate 都用这个结构
 */
export type Signal = {
  /** 信号级别 */
  level: SignalLevel;
  /** 信号代码，例如 "VALIDATION_RANGE_EMPTY", "FORMULA_REF_ERROR" */
  code: string;
  /** 人类可读消息 */
  message: string;
  /** 可选证据（range、采样结果、错误对象等） */
  evidence?: unknown;
  /** 推荐的下一步动作 */
  recommended?: RecommendedAction;
};

/**
 * 推荐动作类型
 */
export type RecommendedAction =
  | "continue"
  | "fix_and_retry"
  | "rollback_and_replan"
  | "ask_user"
  | "abort";

/**
 * 验证输出（统一格式）
 *
 * validator 永远不 throw，只返回这个结构
 */
export type ValidationOutput = {
  /** 是否可以继续（ok=false 不代表要 throw） */
  ok: boolean;
  /** 所有信号（即使 ok=true 也可能有 info/warning） */
  signals: Signal[];
};

// ========== NextAction 系统（P2 澄清门） ==========

/**
 * LLM 每次思考的输出必须是这 3 种之一
 */
export type NextAction =
  | { kind: "clarify"; questions: ClarifyQuestion[]; signals?: Signal[] }
  | { kind: "plan"; plan: ExecutionPlan; signals?: Signal[] }
  | { kind: "execute"; plan: ExecutionPlan; signals?: Signal[] };

/**
 * 澄清问题
 */
export type ClarifyQuestion = {
  /** 问题唯一 ID */
  id: string;
  /** 问用户的话 */
  question: string;
  /** 可选选项（有枚举就给） */
  options?: string[];
  /** 是否必答 */
  required?: boolean;
  /** 默认值（用户不答就用） */
  defaultValue?: string;
};

/**
 * 风险摘要
 */
export type RiskSummary = {
  /** 是否有写操作 */
  hasWriteOps: boolean;
  /** 写操作目标（sheet/range） */
  writeTargets?: string[];
  /** 是否可能覆盖数据 */
  mayOverwrite?: boolean;
};

// ========== StepDecision 系统（P0 每步反思） ==========

/**
 * 每步执行后的决策
 *
 * 必须是这 5 种之一，不能自己发明
 */
export type StepDecision =
  | { action: "continue" }
  | { action: "fix_and_retry"; fix?: StepFix }
  | { action: "rollback_and_replan"; reason: string }
  | { action: "ask_user"; questions: ClarifyQuestion[] }
  | { action: "abort"; reason: string };

/**
 * 步骤修复信息
 */
export type StepFix = {
  /** 修复类型 */
  type: "adjust_parameters" | "adjust_formula" | "shrink_range" | "change_action";
  /** 修正后的参数 */
  patchedParameters?: Record<string, unknown>;
  /** 修正说明 */
  description?: string;
};

// ========== AgentReply 系统（P3 响应结构） ==========

/**
 * Agent 最终回复结构
 *
 * LLM 原话必须保留，模板是补充不是替代
 */
export type AgentReply = {
  /** LLM 原话（必须保留） */
  mainMessage: string;
  /** 模板补充消息（可选） */
  templateMessage?: string;
  /** 主动建议（可选） */
  suggestionMessage?: string;
  /** 调试信息 */
  debug?: AgentReplyDebug;
};

/**
 * 回复调试信息
 */
export type AgentReplyDebug = {
  /** 收集到的信号 */
  signals?: Signal[];
  /** 最终决策 */
  decision?: StepDecision;
  /** 执行步骤 ID */
  stepId?: string;
  /** 执行状态 */
  executionState?: string;
  /** 其他调试数据 */
  extra?: unknown;
};

// ========== 工具类型 ==========

/**
 * 从旧格式转换到 Signal 的辅助函数签名
 */
export type SignalMapper<T> = (oldResult: T) => Signal[];

/**
 * 创建信号的工厂函数
 */
export function createSignal(
  level: SignalLevel,
  code: string,
  message: string,
  options?: {
    evidence?: unknown;
    recommended?: RecommendedAction;
  }
): Signal {
  return {
    level,
    code,
    message,
    evidence: options?.evidence,
    recommended: options?.recommended,
  };
}

/**
 * 创建成功的 ValidationOutput
 */
export function validationOk(signals: Signal[] = []): ValidationOutput {
  return { ok: true, signals };
}

/**
 * 创建失败的 ValidationOutput
 */
export function validationFail(signals: Signal[]): ValidationOutput {
  return { ok: false, signals };
}

/**
 * 判断信号列表中是否有严重问题
 */
export function hasBlockingSignals(signals: Signal[]): boolean {
  return signals.some((s) => s.level === "critical" || s.level === "error");
}

/**
 * 从信号列表推断推荐动作
 */
export function inferRecommendedAction(signals: Signal[]): RecommendedAction {
  // 按严重程度排序
  const critical = signals.find((s) => s.level === "critical");
  if (critical?.recommended) return critical.recommended;
  if (critical) return "abort";

  const error = signals.find((s) => s.level === "error");
  if (error?.recommended) return error.recommended;
  if (error) return "rollback_and_replan";

  const warning = signals.find((s) => s.level === "warning");
  if (warning?.recommended) return warning.recommended;

  return "continue";
}

// ========== 常用信号代码 ==========

export const SignalCodes = {
  // Validator 相关
  PLAN_VALIDATOR_THROW: "PLAN_VALIDATOR_THROW",
  DATA_VALIDATOR_THROW: "DATA_VALIDATOR_THROW",
  VALIDATION_RANGE_EMPTY: "VALIDATION_RANGE_EMPTY",
  VALIDATION_FORMULA_ERROR: "VALIDATION_FORMULA_ERROR",
  VALIDATION_ALL_SAME_VALUE: "VALIDATION_ALL_SAME_VALUE",

  // 执行相关
  TOOL_EXECUTION_FAILED: "TOOL_EXECUTION_FAILED",
  TOOL_RESULT_UNEXPECTED: "TOOL_RESULT_UNEXPECTED",
  STEP_TIMEOUT: "STEP_TIMEOUT",

  // 数据相关
  DATA_MISMATCH: "DATA_MISMATCH",
  FORMULA_REF_ERROR: "FORMULA_REF_ERROR",
  RANGE_NOT_FOUND: "RANGE_NOT_FOUND",

  // 澄清相关
  MISSING_TARGET_RANGE: "MISSING_TARGET_RANGE",
  MISSING_SHEET_NAME: "MISSING_SHEET_NAME",
  AMBIGUOUS_REFERENCE: "AMBIGUOUS_REFERENCE",
  OVERWRITE_CONFIRMATION_NEEDED: "OVERWRITE_CONFIRMATION_NEEDED",
} as const;
