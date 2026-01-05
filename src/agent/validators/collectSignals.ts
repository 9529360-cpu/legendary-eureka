/**
 * collectSignals.ts - Validator 信号收集器（P1 信号化）
 *
 * v2.9.59: 让 validator 永远不 throw，统一返回 signals
 *
 * 核心原则：
 * ┌─────────────────────────────────────────────────────┐
 * │  Validator 可以发现问题，但不能决定怎么处理         │
 * │  处理权交给 AgentCore 的 reflect/decide 循环       │
 * └─────────────────────────────────────────────────────┘
 */

import {
  Signal,
  SignalLevel,
  ValidationOutput,
  createSignal,
  validationOk,
  validationFail,
  SignalCodes,
} from "../AgentProtocol";
import { PlanValidationResult, PlanValidationError, PlanValidationWarning } from "../PlanValidator";
import { DataValidationIssue } from "../DataValidator";
import type { ExecutionPlan } from "../TaskPlanner";

// ========== 核心包装函数 ==========

/**
 * 安全执行验证函数，捕获所有异常转为 signals
 */
export async function safeValidate<T>(
  fn: () => Promise<T>,
  onThrowCode: string
): Promise<ValidationOutput> {
  try {
    const result = await fn();
    return mapToValidationOutput(result, onThrowCode);
  } catch (e: unknown) {
    const error = e as Error;
    return validationFail([
      createSignal("critical", onThrowCode, error?.message || String(e), {
        evidence: { raw: e, stack: error?.stack },
        recommended: "rollback_and_replan",
      }),
    ]);
  }
}

/**
 * 同步版本的 safeValidate
 */
export function safeValidateSync<T>(fn: () => T, onThrowCode: string): ValidationOutput {
  try {
    const result = fn();
    return mapToValidationOutput(result, onThrowCode);
  } catch (e: unknown) {
    const error = e as Error;
    return validationFail([
      createSignal("critical", onThrowCode, error?.message || String(e), {
        evidence: { raw: e, stack: error?.stack },
        recommended: "rollback_and_replan",
      }),
    ]);
  }
}

// ========== 结果映射 ==========

/**
 * 将各种 validator 结果统一映射成 ValidationOutput
 */
function mapToValidationOutput<T>(result: T, fallbackCode: string): ValidationOutput {
  if (result === null || result === undefined) {
    return validationOk();
  }

  // 如果已经是 ValidationOutput 格式
  if (isValidationOutput(result)) {
    return result;
  }

  // PlanValidationResult 格式
  if (isPlanValidationResult(result)) {
    return mapPlanValidationResult(result);
  }

  // DataValidationIssue 格式（单个 issue）
  if (isDataValidationIssue(result)) {
    return mapDataValidationIssue(result);
  }

  // DataValidationIssue 数组
  if (Array.isArray(result) && result.length > 0 && isDataValidationIssue(result[0])) {
    return mapDataValidationIssueArray(result as DataValidationIssue[]);
  }

  // 简单的 { passed: boolean } 格式
  if (typeof result === "object" && "passed" in result) {
    const passed = (result as { passed: boolean }).passed;
    if (passed) {
      return validationOk();
    }
    return validationFail([
      createSignal("error", fallbackCode, "验证未通过", {
        evidence: result,
        recommended: "fix_and_retry",
      }),
    ]);
  }

  // 简单的 boolean
  if (typeof result === "boolean") {
    return result
      ? validationOk()
      : validationFail([createSignal("error", fallbackCode, "验证返回 false")]);
  }

  // 无法识别，当作成功
  return validationOk();
}

// ========== 类型守卫 ==========

function isValidationOutput(obj: unknown): obj is ValidationOutput {
  return (
    typeof obj === "object" &&
    obj !== null &&
    "ok" in obj &&
    "signals" in obj &&
    Array.isArray((obj as ValidationOutput).signals)
  );
}

function isPlanValidationResult(obj: unknown): obj is PlanValidationResult {
  return (
    typeof obj === "object" &&
    obj !== null &&
    "passed" in obj &&
    "errors" in obj &&
    "warnings" in obj &&
    Array.isArray((obj as PlanValidationResult).errors)
  );
}

function isDataValidationIssue(obj: unknown): obj is DataValidationIssue {
  return (
    typeof obj === "object" &&
    obj !== null &&
    "ruleId" in obj &&
    "ruleName" in obj &&
    "severity" in obj &&
    "message" in obj
  );
}

// ========== 具体映射函数 ==========

/**
 * PlanValidationResult → ValidationOutput
 */
function mapPlanValidationResult(result: PlanValidationResult): ValidationOutput {
  const signals: Signal[] = [];

  // 映射 errors
  for (const err of result.errors) {
    signals.push(mapPlanErrorToSignal(err));
  }

  // 映射 warnings
  for (const warn of result.warnings) {
    signals.push(mapPlanWarningToSignal(warn));
  }

  return {
    ok: result.passed,
    signals,
  };
}

/**
 * PlanValidationError → Signal
 */
function mapPlanErrorToSignal(err: PlanValidationError): Signal {
  return createSignal("error", `PLAN_${err.ruleId}`, err.message, {
    evidence: { affectedSteps: err.affectedSteps, details: err.details },
    recommended: mapPlanErrorToRecommendation(err.ruleId),
  });
}

/**
 * PlanValidationWarning → Signal
 */
function mapPlanWarningToSignal(warn: PlanValidationWarning): Signal {
  return createSignal("warning", `PLAN_${warn.ruleId}`, warn.message, {
    evidence: { affectedSteps: warn.affectedSteps, details: warn.details },
    recommended: "continue",
  });
}

/**
 * DataValidationIssue → ValidationOutput
 */
function mapDataValidationIssue(issue: DataValidationIssue): ValidationOutput {
  const signal = mapDataIssueToSignal(issue);
  const ok = issue.severity !== "block";
  return { ok, signals: [signal] };
}

/**
 * DataValidationIssue[] → ValidationOutput
 */
function mapDataValidationIssueArray(issues: DataValidationIssue[]): ValidationOutput {
  const signals = issues.map(mapDataIssueToSignal);
  const hasBlock = issues.some((i) => i.severity === "block");
  return { ok: !hasBlock, signals };
}

/**
 * DataValidationIssue → Signal
 */
function mapDataIssueToSignal(issue: DataValidationIssue): Signal {
  const level: SignalLevel = issue.severity === "block" ? "error" : "warning";

  return createSignal(level, `DATA_${issue.ruleId}`, issue.message, {
    evidence: {
      affectedRange: issue.affectedRange,
      affectedCells: issue.affectedCells,
      confidence: issue.confidence,
      evidence: issue.evidence,
    },
    recommended: mapDataIssueToRecommendation(issue),
  });
}

// ========== 推荐动作映射 ==========

function mapPlanErrorToRecommendation(
  ruleId: string
): "continue" | "fix_and_retry" | "rollback_and_replan" | "ask_user" | "abort" {
  const id = ruleId.toUpperCase();

  if (id.includes("MISSING") || id.includes("NOT_FOUND")) {
    return "ask_user";
  }
  if (id.includes("CIRCULAR") || id.includes("INVALID")) {
    return "rollback_and_replan";
  }
  if (id.includes("DEPENDENCY")) {
    return "fix_and_retry";
  }

  return "fix_and_retry";
}

function mapDataIssueToRecommendation(
  issue: DataValidationIssue
): "continue" | "fix_and_retry" | "rollback_and_replan" | "ask_user" | "abort" {
  // 根据 suggestedFixPlan 推断
  if (issue.suggestedFixPlan && issue.suggestedFixPlan.length > 0) {
    return "fix_and_retry";
  }

  // 根据严重程度
  if (issue.severity === "block") {
    return "rollback_and_replan";
  }

  return "continue";
}

// ========== 便捷收集函数 ==========

/**
 * 收集一个步骤执行后的所有信号
 */
export async function collectStepSignals(
  step: { action: string; parameters?: Record<string, unknown> },
  toolResult: { success: boolean; output?: unknown; error?: string },
  validators: {
    planValidator?: () => Promise<PlanValidationResult>;
    dataValidator?: () => Promise<DataValidationIssue[]>;
  }
): Promise<Signal[]> {
  const signals: Signal[] = [];

  // 工具执行结果信号
  if (!toolResult.success) {
    signals.push(
      createSignal("error", SignalCodes.TOOL_EXECUTION_FAILED, toolResult.error || "工具执行失败", {
        evidence: { action: step.action, output: toolResult.output },
        recommended: "fix_and_retry",
      })
    );
  }

  // 收集 validator 信号
  if (validators.planValidator) {
    const planOut = await safeValidate(validators.planValidator, SignalCodes.PLAN_VALIDATOR_THROW);
    signals.push(...planOut.signals);
  }

  if (validators.dataValidator) {
    const dataOut = await safeValidate(validators.dataValidator, SignalCodes.DATA_VALIDATOR_THROW);
    signals.push(...dataOut.signals);
  }

  return signals;
}

/**
 * 收集整个计划的验证信号
 */
export async function collectPlanSignals(
  plan: ExecutionPlan,
  validators: {
    planValidator: (plan: ExecutionPlan) => Promise<PlanValidationResult>;
    dataValidator?: (plan: ExecutionPlan) => Promise<DataValidationIssue[]>;
  }
): Promise<ValidationOutput> {
  const allSignals: Signal[] = [];

  // Plan validation
  const planOut = await safeValidate(
    () => validators.planValidator(plan),
    SignalCodes.PLAN_VALIDATOR_THROW
  );
  allSignals.push(...planOut.signals);

  // Data validation (if provided)
  if (validators.dataValidator) {
    const dataOut = await safeValidate(
      () => validators.dataValidator!(plan),
      SignalCodes.DATA_VALIDATOR_THROW
    );
    allSignals.push(...dataOut.signals);
  }

  // 综合判断 ok
  const hasCritical = allSignals.some((s) => s.level === "critical");
  const hasError = allSignals.some((s) => s.level === "error");

  return {
    ok: !hasCritical && !hasError,
    signals: allSignals,
  };
}
