/**
 * 验证相关类型定义
 *
 * 从 AgentCore.ts 抽取，用于定义硬逻辑校验接口
 */

/**
 * v2.9.2: 硬逻辑校验规则 - 支持异步
 */
export interface HardValidationRule {
  id: string;
  name: string;
  description: string;
  type: "pre_execution" | "post_execution" | "data_quality";
  // v2.9.2: 改成异步！才能读取 Excel 做真正的验证
  check: (context: ValidationContext, excelReader?: ExcelReader) => Promise<ValidationCheckResult>;
  severity: "block" | "warn"; // block = 必须通过，warn = 仅警告
  // v2.9.6: 规则是否启用（默认 true）
  enabled?: boolean;
}

/**
 * v2.9.2: Excel 读取器接口 - 用于硬校验时读取 Excel 数据
 */
export interface ExcelReader {
  readRange: (
    sheet: string,
    range: string
  ) => Promise<{ values: unknown[][]; formulas: string[][] }>;
  sampleRows: (sheet: string, count: number) => Promise<unknown[][]>;
  getColumnFormulas: (sheet: string, column: string) => Promise<string[]>;
}

/**
 * 校验上下文
 */
export interface ValidationContext {
  toolName: string;
  toolInput: Record<string, unknown>;
  toolOutput?: string;
  currentSheet?: string;
  affectedRange?: string;
  previousData?: unknown;
  // v2.9.2: 新增，用于真正的 Excel 验证
  previousFormulas?: string[][];
}

/**
 * 校验检查结果
 */
export interface ValidationCheckResult {
  passed: boolean;
  message: string;
  details?: string[];
  suggestedFix?: string;
}

/**
 * v2.8.7 发现的问题 - Agent 发现的必须解决的问题
 */
export interface DiscoveredIssue {
  id: string;
  type:
    | "hardcoded"
    | "structural"
    | "formula_error"
    | "data_quality"
    | "missing_reference"
    | "other";
  severity: "critical" | "warning";
  description: string;
  location?: string; // 在哪里发现的
  discoveredAt: Date;
  resolved: boolean;
  resolvedAt?: Date;
  resolution?: string; // 如何解决的
}

/**
 * v2.9.0 操作历史 - 用于回滚
 */
export interface OperationRecord {
  id: string;
  timestamp: Date;
  toolName: string;
  toolInput: Record<string, unknown>;
  result: "success" | "failed" | "rolled_back";
  rollbackData?: {
    previousState?: unknown;
    previousFormulas?: unknown;
    rollbackAction?: string;
    rollbackParams?: Record<string, unknown>;
  };
}

/**
 * v2.9.18: 质量问题
 */
export interface QualityIssue {
  severity: "error" | "warning" | "info";
  type:
    | "hardcoded"
    | "missing_formula"
    | "format"
    | "naming"
    | "empty_cell"
    | "error_value"
    | "duplicate"
    | "inconsistent";
  location: string;
  message: string;
  autoFixable: boolean;
  fixAction?: string;
}

/**
 * v2.9.18: 质量报告
 */
export interface QualityReport {
  score: number; // 0-100
  issues: QualityIssue[];
  suggestions: string[];
  passedChecks: string[];
  autoFixedCount: number;
}

/**
 * 关键错误检测结果 - v2.7 硬约束
 */
export interface CriticalErrorResult {
  hasCriticalError: boolean;
  errors: import("../FormulaValidator").ExecutionError[];
  reason: string;
  suggestion: string;
}

/**
 * v2.9.18: 错误恢复策略
 */
export type ErrorRecoveryStrategy =
  | "retry"
  | "retry_with_params"
  | "fallback"
  | "ask_user"
  | "rollback"
  | "skip";

/**
 * v2.9.18: 错误恢复结果
 */
export interface ErrorRecoveryResult {
  strategy: ErrorRecoveryStrategy;
  succeeded: boolean;
  originalError: string;
  recoveryAction: string;
  result?: string;
}

/**
 * v2.9.22: 错误根因分析
 */
export interface ErrorRootCauseAnalysis {
  /** 原始错误 */
  originalError: string;
  /** 根本原因 */
  rootCause: string;
  /** 原因类型 */
  causeType: "user_input" | "data_issue" | "tool_bug" | "api_limit" | "permission" | "unknown";
  /** 影响范围 */
  impactScope: "current_step" | "current_task" | "session" | "persistent";
  /** 修复建议 */
  fixSuggestions: string[];
  /** 预防建议 */
  preventionTips: string[];
  /** 置信度 */
  confidence: number;
}

/**
 * v2.9.22: 自动重试策略
 */
export interface RetryStrategy {
  /** 策略ID */
  id: string;
  /** 最大重试次数 */
  maxRetries: number;
  /** 退避策略 */
  backoffType: "fixed" | "linear" | "exponential";
  /** 初始延迟(ms) */
  initialDelayMs: number;
  /** 最大延迟(ms) */
  maxDelayMs: number;
  /** 可重试的错误类型 */
  retryableErrors: string[];
  /** 每次重试前的变换 */
  transformBeforeRetry?: "simplify" | "decompose" | "rephrase";
}

/**
 * v2.9.22: 自愈动作
 */
export interface SelfHealingAction {
  /** 动作ID */
  id: string;
  /** 触发条件 */
  triggerCondition: string;
  /** 自愈操作 */
  healingAction: "retry" | "rollback" | "skip" | "alternative" | "ask_user";
  /** 替代方案 */
  alternative?: string;
  /** 成功率 */
  successRate: number;
}

/**
 * v2.9.22: 假设验证
 */
export interface HypothesisValidation {
  /** 假设ID */
  id: string;
  /** 假设内容 */
  hypothesis: string;
  /** 验证方法 */
  validationMethod: "data_check" | "execution" | "user_confirm" | "inference";
  /** 验证结果 */
  result: "confirmed" | "rejected" | "inconclusive" | "pending";
  /** 证据 */
  evidence: string[];
  /** 置信度 */
  confidence: number;
}

/**
 * v2.9.22: 不确定性量化
 */
export interface UncertaintyQuantification {
  /** 整体不确定性 (0-100) */
  overallUncertainty: number;
  /** 各维度不确定性 */
  dimensions: {
    intentUnderstanding: number;
    dataAvailability: number;
    toolReliability: number;
    contextClarity: number;
  };
  /** 主要不确定来源 */
  primarySource: string;
  /** 降低不确定性的建议 */
  reductionSuggestions: string[];
}

/**
 * v2.9.22: 反事实推理
 */
export interface CounterfactualReasoning {
  /** 原始场景 */
  originalScenario: string;
  /** 反事实场景 */
  counterfactualScenario: string;
  /** 预测的不同结果 */
  predictedDifference: string;
  /** 置信度 */
  confidence: number;
  /** 推理依据 */
  reasoning: string;
}

/**
 * v2.9.18: 反思结果 - 每步执行后的验证结果
 *
 * @deprecated 使用 StepReflector.ReflectionResult 代替
 * 保留用于向后兼容旧的 reflect() 方法
 */
export interface LegacyReflectionResult {
  stepId: string;
  succeeded: boolean;
  expectedOutcome: string;
  actualOutcome: string;
  gap: string | null;
  action: "continue" | "retry" | "fix" | "replan" | "ask_user";
  fixPlan?: string;
  confidence: number; // 0-100
}
