/**
 * 工作流系统相关类型定义
 *
 * 从 AgentCore.ts 抽取，借鉴 LlamaIndex Workflows 设计
 */

import type { ToolResult, ToolCallInfo, ToolCallResultData } from "./tool";

// ========== 工作流事件相关类型 ==========

/**
 * 工作流事件基类 - 所有事件的抽象基类
 *
 * 借鉴 LlamaIndex 的 workflowEvent<T>() 设计
 */
export interface WorkflowEvent<T = unknown> {
  type: string;
  data: T;
  timestamp: Date;
  source?: string;
}

/**
 * Agent 流式输出事件 - 类似 LlamaIndex 的 AgentStream
 */
export interface AgentStreamData {
  delta: string; // 增量内容
  response: string; // 累积响应
  currentAgentName: string; // 当前 Agent 名称
  toolCalls: ToolCallInfo[]; // 当前的工具调用
  thinkingDelta?: string; // 思考过程增量
}

/**
 * Agent 最终输出事件 - 类似 LlamaIndex 的 AgentOutput
 */
export interface AgentOutputData {
  response: string; // 最终响应
  structuredResponse?: Record<string, unknown>; // 结构化响应
  currentAgentName: string; // Agent 名称
  toolCalls: ToolCallInfo[]; // 所有工具调用
  retryMessages: string[]; // 重试消息
}

/**
 * 结构化输出流事件 - 类似 LlamaIndex 的 AgentStreamStructuredOutput
 */
export interface AgentStreamStructuredOutputData {
  output: Record<string, unknown>;
}

/**
 * 工作流状态 - 类似 LlamaIndex 的状态中间件
 *
 * 用于在工作流执行过程中维护状态
 */
export interface WorkflowState {
  // 当前任务信息
  currentTaskId: string | null;
  currentRequest: string | null;

  // 执行状态
  isRunning: boolean;
  isPaused: boolean;

  // 上下文数据
  collectedData: unknown[];
  toolCallHistory: ToolCallInfo[];

  // 计数器
  stepCounter: number;
  maxSteps: number;

  // 用户交互状态
  awaitingConfirmation: boolean;
  awaitingFollowUp: boolean;

  // v2.9.60: 补充缺失的属性（修复遗留类型问题）
  startTime?: Date;
  endTime?: Date;
  toolsCalled: string[];
  currentResponse: string;
  errors: Error[];

  // 自定义状态扩展
  custom: Record<string, unknown>;
}

/**
 * 事件处理器类型定义 (LlamaIndex: workflow.handle([event], handler))
 */
export type WorkflowEventHandler<T = unknown> = (
  context: WorkflowContextInterface,
  data: T
) => Promise<void | { type: string; payload: unknown }>;

/**
 * v2.9.60: 工作流执行器接口
 */
export interface SimpleWorkflow {
  on<T>(
    eventType: string | WorkflowEvent<T>,
    handler: WorkflowEventHandler<T>
  ): SimpleWorkflow;
  run(context: WorkflowContextInterface): Promise<WorkflowEventStreamInterface>;
  getStream(): WorkflowEventStreamInterface;
  getRegistry(): WorkflowEventRegistryInterface;
}

/**
 * 工作流上下文接口
 */
export interface WorkflowContextInterface {
  sendEvent<T>(event: WorkflowEvent<T> | { type: string; payload?: T; data?: T }): void;
  getEvents(): Array<WorkflowEvent<unknown>>;
  consumeEvent(): WorkflowEvent<unknown> | undefined;
  clearEvents(): void;
  getState(): WorkflowState;
  updateState(updates: Partial<WorkflowState>): void;
  setCustom<T>(key: string, value: T): void;
  getCustom<T>(key: string): T | undefined;
  cancel(): void;
  readonly isCancelled: boolean;
  readonly signal: AbortSignal;
}

/**
 * 工作流事件流接口
 */
export interface WorkflowEventStreamInterface {
  push(event: WorkflowEvent<unknown> | { type: string; payload: unknown }): void;
  until<T>(stopEvent: WorkflowEvent<T>): this;
  toArray(): Array<{ type: string; payload: unknown }>;
  last(): { type: string; payload: unknown } | undefined;
  filter<T>(eventType: WorkflowEvent<T>): Array<{ type: string; payload: T }>;
  [Symbol.iterator](): IterableIterator<{ type: string; payload: unknown }>;
}

/**
 * 工作流事件注册表接口
 */
export interface WorkflowEventRegistryInterface {
  handle<T>(event: WorkflowEvent<T>, handler: WorkflowEventHandler<T>): this;
  dispatch<T>(
    eventType: string,
    context: WorkflowContextInterface,
    data: T
  ): Promise<Array<{ type: string; payload: unknown } | void>>;
  getRegisteredEvents(): string[];
}

// ========== 计划执行相关类型 ==========

/**
 * 执行计划步骤
 */
export interface PlanStep {
  /** 步骤索引 */
  index: number;
  /** 步骤描述 */
  description: string;
  /** 要调用的工具名称 */
  toolName: string;
  /** 工具参数 */
  toolParams: Record<string, unknown>;
  /** 预期结果描述 */
  expectedResult?: string;
  /** 步骤依赖（前置步骤索引） */
  dependencies?: number[];
  /** 是否可选 */
  optional?: boolean;
  /** 步骤重要性 */
  importance?: "critical" | "normal" | "low";
}

/**
 * 执行计划
 */
export interface ExecutionPlan {
  /** 计划ID */
  id: string;
  /** 计划名称 */
  name: string;
  /** 计划描述 */
  description: string;
  /** 执行步骤列表 */
  steps: PlanStep[];
  /** 创建时间 */
  createdAt: Date;
  /** 是否需要用户确认 */
  requiresConfirmation: boolean;
  /** 预计执行时间 (ms) */
  estimatedDuration?: number;
  /** 风险等级 */
  riskLevel?: "low" | "medium" | "high";
  /** 计划元数据 */
  metadata?: Record<string, unknown>;
}

/**
 * 计划执行结果
 */
export interface PlanExecutionResult {
  /** 计划ID */
  planId: string;
  /** 是否成功 */
  success: boolean;
  /** 完成的步骤数 */
  completedSteps: number;
  /** 总步骤数 */
  totalSteps: number;
  /** 每个步骤的结果 */
  stepResults: Array<{
    stepIndex: number;
    success: boolean;
    result?: ToolResult;
    error?: string;
    duration: number;
  }>;
  /** 总执行时间 (ms) */
  totalDuration: number;
  /** 最终输出 */
  output?: string;
  /** 错误信息 */
  error?: string;
}

// ========== 进度相关类型 ==========

/**
 * 任务进度信息
 */
export interface TaskProgress {
  /** 任务ID */
  taskId: string;
  /** 当前步骤 */
  currentStep: number;
  /** 总步骤数 */
  totalSteps: number;
  /** 进度百分比 (0-100) */
  percentage: number;
  /** 当前步骤描述 */
  currentStepDescription: string;
  /** 状态 */
  status: "pending" | "running" | "completed" | "failed" | "paused";
  /** 开始时间 */
  startTime: Date;
  /** 预计完成时间 */
  estimatedEndTime?: Date;
  /** 已用时间 (ms) */
  elapsedTime: number;
  /** 剩余时间 (ms) */
  remainingTime?: number;
}

/**
 * 进度步骤信息
 */
export interface ProgressStep {
  /** 步骤索引 */
  index: number;
  /** 步骤名称 */
  name: string;
  /** 步骤描述 */
  description: string;
  /** 步骤状态 */
  status: "pending" | "running" | "completed" | "failed" | "skipped";
  /** 开始时间 */
  startTime?: Date;
  /** 结束时间 */
  endTime?: Date;
  /** 持续时间 (ms) */
  duration?: number;
  /** 结果摘要 */
  resultSummary?: string;
  /** 错误信息 */
  error?: string;
}

// ========== 工作流事件工厂 ==========

/**
 * 工作流事件工厂返回类型
 */
export interface WorkflowEventFactory<T> {
  type: string;
  with: (data: T) => WorkflowEvent<T>;
  is: (event: WorkflowEvent<unknown>) => event is WorkflowEvent<T>;
}

// ========== 预定义工作流事件类型 ==========

/**
 * 任务开始事件数据
 */
export interface TaskStartEventData {
  taskId: string;
  request: string;
}

/**
 * 任务完成事件数据
 */
export interface TaskCompleteEventData {
  taskId: string;
  result: string;
}

/**
 * 任务错误事件数据
 */
export interface TaskErrorEventData {
  taskId: string;
  error: string;
}

/**
 * 任务挂起事件数据
 */
export interface TaskPendingEventData {
  taskId: string;
  reason: string;
}

/**
 * 计划生成事件数据
 */
export interface PlanGeneratedEventData {
  plan: ExecutionPlan;
}

/**
 * 计划步骤开始事件数据
 */
export interface PlanStepStartEventData {
  stepIndex: number;
  step: PlanStep;
}

/**
 * 计划步骤完成事件数据
 */
export interface PlanStepCompleteEventData {
  stepIndex: number;
  result: ToolResult;
}

/**
 * 计划确认请求事件数据
 */
export interface PlanConfirmationRequiredEventData {
  taskId: string;
  preview: string;
}

/**
 * 进度更新事件数据
 */
export interface ProgressUpdateEventData {
  current: number;
  total: number;
  message: string;
}

/**
 * 验证开始事件数据
 */
export interface ValidationStartEventData {
  type: string;
}

/**
 * 验证完成事件数据
 */
export interface ValidationCompleteEventData {
  passed: boolean;
  errors: string[];
}

/**
 * 思考链步骤开始事件数据
 */
export interface CotStepStartEventData {
  step: string;
}

/**
 * 思考链步骤完成事件数据
 */
export interface CotStepCompleteEventData {
  step: string;
  result: string;
}

/**
 * 跟进上下文设置事件数据
 */
export interface FollowUpContextSetEventData {
  originalRequest: string;
  suggestedAction: string;
}

/**
 * 跟进处理事件数据
 */
export interface FollowUpHandledEventData {
  userReply: string;
  action: string;
}
