/**
 * ExecutionContext - 不可变执行上下文 v1.0
 *
 * 借鉴自 Activepieces FlowExecutorContext
 *
 * 核心设计：
 * 1. 所有状态更新返回新实例（不可变）
 * 2. verdict 判决机制控制执行流程
 * 3. 支持暂停/恢复/审批中断
 * 4. 步骤输出历史追踪
 *
 * @see https://github.com/activepieces/activepieces
 */

import type { ToolResult as _ToolResult } from "./AgentCore";
import type { ApprovalRequest } from "./ApprovalManager";

// ==================== 类型定义 ====================

/**
 * 执行判决状态
 */
export type VerdictStatus =
  | "running" // 正在执行
  | "paused" // 已暂停（用户请求）
  | "awaiting_approval" // 等待审批
  | "succeeded" // 执行成功
  | "failed" // 执行失败
  | "cancelled"; // 已取消

/**
 * 执行判决
 */
export type ExecutionVerdict =
  | { status: "running" }
  | { status: "paused"; reason: string }
  | { status: "awaiting_approval"; approvalRequest: ApprovalRequest }
  | { status: "succeeded"; stopResponse?: unknown }
  | { status: "failed"; failedStep: FailedStep }
  | { status: "cancelled"; reason: string };

/**
 * 失败步骤信息
 */
export interface FailedStep {
  name: string;
  displayName: string;
  message: string;
  error?: Error;
}

/**
 * 步骤输出状态
 */
export type StepOutputStatus = "running" | "succeeded" | "failed" | "paused" | "skipped";

/**
 * 步骤输出
 */
export interface StepOutput {
  /** 步骤名称 */
  name: string;
  /** 步骤类型 */
  type: string;
  /** 输出状态 */
  status: StepOutputStatus;
  /** 输入参数 */
  input: Record<string, unknown>;
  /** 输出结果 */
  output: unknown;
  /** 错误信息 */
  errorMessage?: string;
  /** 执行时间（毫秒） */
  duration: number;
  /** 开始时间 */
  startTime: Date;
  /** 结束时间 */
  endTime?: Date;
}

/**
 * 执行上下文配置
 */
export interface ExecutionContextConfig {
  maxSteps: number;
  timeoutMs: number;
  enableRetry: boolean;
  maxRetries: number;
}

/**
 * 默认配置
 */
export const DEFAULT_EXECUTION_CONFIG: ExecutionContextConfig = {
  maxSteps: 30,
  timeoutMs: 5 * 60 * 1000, // 5分钟
  enableRetry: true,
  maxRetries: 3,
};

// ==================== ExecutionContext 类 ====================

/**
 * 不可变执行上下文
 *
 * 所有状态更新方法都返回新实例，确保状态不可变
 */
export class ExecutionContext {
  /** 只读标签 */
  readonly tags: readonly string[];
  /** 只读步骤记录 */
  readonly steps: Readonly<Record<string, StepOutput>>;
  /** 执行判决 */
  readonly verdict: ExecutionVerdict;
  /** 执行时间（毫秒） */
  readonly duration: number;
  /** 步骤计数 */
  readonly stepsCount: number;
  /** 开始时间 */
  readonly startTime: Date;
  /** 任务 ID */
  readonly taskId: string;
  /** 配置 */
  readonly config: ExecutionContextConfig;
  /** 审批请求 ID（用于恢复） */
  readonly pendingApprovalId?: string;

  constructor(init?: Partial<ExecutionContext>) {
    this.tags = init?.tags ?? [];
    this.steps = init?.steps ?? {};
    this.verdict = init?.verdict ?? { status: "running" };
    this.duration = init?.duration ?? -1;
    this.stepsCount = init?.stepsCount ?? 0;
    this.startTime = init?.startTime ?? new Date();
    this.taskId = init?.taskId ?? this.generateTaskId();
    this.config = init?.config ?? DEFAULT_EXECUTION_CONFIG;
    this.pendingApprovalId = init?.pendingApprovalId;
  }

  /**
   * 创建空上下文
   */
  static empty(config?: Partial<ExecutionContextConfig>): ExecutionContext {
    return new ExecutionContext({
      config: { ...DEFAULT_EXECUTION_CONFIG, ...config },
    });
  }

  /**
   * 生成任务 ID
   */
  private generateTaskId(): string {
    const timestamp = Date.now().toString(36);
    const random = Math.random().toString(36).substring(2, 8);
    return `task_${timestamp}_${random}`;
  }

  // ==================== 状态更新方法（返回新实例） ====================

  /**
   * 更新/插入步骤输出
   */
  upsertStep(stepName: string, stepOutput: StepOutput): ExecutionContext {
    const steps = {
      ...this.steps,
      [stepName]: stepOutput,
    };
    return new ExecutionContext({
      ...this,
      steps,
      stepsCount: Object.keys(steps).length,
    });
  }

  /**
   * 设置判决
   */
  setVerdict(verdict: ExecutionVerdict): ExecutionContext {
    return new ExecutionContext({
      ...this,
      verdict,
    });
  }

  /**
   * 设置执行时间
   */
  setDuration(duration: number): ExecutionContext {
    return new ExecutionContext({
      ...this,
      duration,
    });
  }

  /**
   * 添加标签
   */
  addTags(newTags: string[]): ExecutionContext {
    const uniqueTags = [...new Set([...this.tags, ...newTags])];
    return new ExecutionContext({
      ...this,
      tags: uniqueTags,
    });
  }

  /**
   * 设置待审批 ID
   */
  setPendingApprovalId(approvalId: string): ExecutionContext {
    return new ExecutionContext({
      ...this,
      pendingApprovalId: approvalId,
    });
  }

  /**
   * 增加步骤计数
   */
  incrementStepsExecuted(): ExecutionContext {
    return new ExecutionContext({
      ...this,
      stepsCount: this.stepsCount + 1,
    });
  }

  // ==================== 判决快捷方法 ====================

  /**
   * 标记为成功
   */
  markSucceeded(stopResponse?: unknown): ExecutionContext {
    return this.setVerdict({ status: "succeeded", stopResponse });
  }

  /**
   * 标记为失败
   */
  markFailed(failedStep: FailedStep): ExecutionContext {
    return this.setVerdict({ status: "failed", failedStep });
  }

  /**
   * 标记为暂停
   */
  markPaused(reason: string): ExecutionContext {
    return this.setVerdict({ status: "paused", reason });
  }

  /**
   * 标记为等待审批
   */
  markAwaitingApproval(approvalRequest: ApprovalRequest): ExecutionContext {
    return this.setVerdict({ status: "awaiting_approval", approvalRequest }).setPendingApprovalId(
      approvalRequest.approvalId
    );
  }

  /**
   * 标记为已取消
   */
  markCancelled(reason: string): ExecutionContext {
    return this.setVerdict({ status: "cancelled", reason });
  }

  /**
   * 恢复执行（从暂停/审批状态）
   */
  resume(): ExecutionContext {
    return this.setVerdict({ status: "running" });
  }

  // ==================== 状态查询方法 ====================

  /**
   * 是否正在运行
   */
  isRunning(): boolean {
    return this.verdict.status === "running";
  }

  /**
   * 是否已完成（成功或失败）
   */
  isCompleted(): boolean {
    return this.verdict.status === "succeeded" || this.verdict.status === "failed";
  }

  /**
   * 是否等待审批
   */
  isAwaitingApproval(): boolean {
    return this.verdict.status === "awaiting_approval";
  }

  /**
   * 是否暂停
   */
  isPaused(): boolean {
    return this.verdict.status === "paused";
  }

  /**
   * 是否应该中断执行
   */
  shouldBreakExecution(): boolean {
    return this.verdict.status !== "running";
  }

  /**
   * 获取步骤输出
   */
  getStepOutput(stepName: string): StepOutput | undefined {
    return this.steps[stepName];
  }

  /**
   * 检查步骤是否已完成
   */
  isStepCompleted(stepName: string): boolean {
    const step = this.steps[stepName];
    return step ? step.status === "succeeded" || step.status === "failed" : false;
  }

  /**
   * 获取所有步骤输出（用于最终结果）
   */
  getAllStepOutputs(): Record<string, unknown> {
    return Object.fromEntries(
      Object.entries(this.steps).map(([name, step]) => [name, step.output])
    );
  }

  /**
   * 获取失败的步骤
   */
  getFailedSteps(): StepOutput[] {
    return Object.values(this.steps).filter((step) => step.status === "failed");
  }

  /**
   * 计算总执行时间
   */
  getTotalDuration(): number {
    if (this.duration >= 0) return this.duration;
    return Date.now() - this.startTime.getTime();
  }

  // ==================== 序列化方法 ====================

  /**
   * 转换为 JSON（用于持久化/传输）
   */
  toJSON(): object {
    return {
      taskId: this.taskId,
      tags: [...this.tags],
      steps: { ...this.steps },
      verdict: this.verdict,
      duration: this.getTotalDuration(),
      stepsCount: this.stepsCount,
      startTime: this.startTime.toISOString(),
      config: this.config,
      pendingApprovalId: this.pendingApprovalId,
    };
  }

  /**
   * 从 JSON 恢复
   */
  static fromJSON(json: ReturnType<ExecutionContext["toJSON"]>): ExecutionContext {
    const data = json as Record<string, unknown>;
    return new ExecutionContext({
      taskId: data.taskId as string,
      tags: data.tags as string[],
      steps: data.steps as Record<string, StepOutput>,
      verdict: data.verdict as ExecutionVerdict,
      duration: data.duration as number,
      stepsCount: data.stepsCount as number,
      startTime: new Date(data.startTime as string),
      config: data.config as ExecutionContextConfig,
      pendingApprovalId: data.pendingApprovalId as string | undefined,
    });
  }
}

// ==================== 辅助函数 ====================

/**
 * 创建步骤输出
 */
export function createStepOutput(
  name: string,
  type: string,
  input: Record<string, unknown> = {}
): StepOutput {
  return {
    name,
    type,
    status: "running",
    input,
    output: null,
    duration: 0,
    startTime: new Date(),
  };
}

/**
 * 更新步骤输出为成功
 */
export function markStepSucceeded(step: StepOutput, output: unknown): StepOutput {
  const endTime = new Date();
  return {
    ...step,
    status: "succeeded",
    output,
    endTime,
    duration: endTime.getTime() - step.startTime.getTime(),
  };
}

/**
 * 更新步骤输出为失败
 */
export function markStepFailed(step: StepOutput, errorMessage: string): StepOutput {
  const endTime = new Date();
  return {
    ...step,
    status: "failed",
    errorMessage,
    endTime,
    duration: endTime.getTime() - step.startTime.getTime(),
  };
}

// ==================== 导出 ====================

export default ExecutionContext;
