/**
 * 任务相关类型定义
 *
 * 从 AgentCore.ts 抽取，用于定义 Agent 任务的接口
 */

import type { DataModel } from "../DataModeler";
import type { ExecutionError, SampleValidationResult } from "../FormulaValidator";
import type { ExecutionPlan, ReplanResult } from "../TaskPlanner";
import type {
  ValidationCheckResult,
  DiscoveredIssue,
  OperationRecord,
  QualityReport,
} from "./validation";

/**
 * Agent 决策 - LLM 思考后的输出
 */
export interface AgentDecision {
  thought: string; // 思考过程
  action: "tool" | "respond" | "complete" | "clarify";
  toolName?: string;
  toolInput?: Record<string, unknown>;
  response?: string;
  isComplete: boolean;
  confidence?: number;
}

/**
 * v2.9.29: LLM 生成的执行计划
 */
export interface LLMGeneratedPlan {
  intent: "query" | "operation" | "analysis";
  steps: Array<{
    order: number;
    action: string;
    parameters: Record<string, unknown>;
    description: string;
    isWriteOperation: boolean;
    successCondition?: string;
  }>;
  completionMessage: string;
}

/**
 * v2.9.43: 任务状态类型
 */
export type AgentTaskStatus =
  | "pending"
  | "running"
  | "completed"
  | "failed"
  | "cancelled"
  | "pending_confirmation"
  | "pending_clarification";

/**
 * v2.9.17: 任务复杂度
 */
export type TaskComplexity = "simple" | "medium" | "complex";

/**
 * Agent 步骤 - 执行历史中的一步
 */
export interface AgentStep {
  id: string;
  type: "think" | "act" | "observe" | "respond" | "plan" | "validate" | "error";
  thought?: string;
  toolName?: string;
  toolInput?: Record<string, unknown>;
  observation?: string;
  timestamp: Date;
  duration?: number;
  // v2.7 新增
  phase?: "planning" | "validation" | "execution" | "verification";
  validationErrors?: ExecutionError[];
  // v2.9.39: 添加 error 属性用于记录步骤执行错误
  error?: string;
}

/**
 * v2.9.58: 澄清上下文 - 用于多轮澄清
 */
export interface ClarificationContext {
  originalRequest: string;
  analysisResult: ClarificationCheckResult;
  sessionId?: string;
  collectedInfo?: Record<string, unknown>;
}

/**
 * v2.9.58: 澄清检查结果
 */
export interface ClarificationCheckResult {
  needsClarification: boolean;
  confidence: number;
  clarificationMessage?: string;
  reasons?: string[];
  suggestedOptions?: Array<{ id: string; label: string; description?: string }>;
}

/**
 * Agent 任务 - 用户交给 Agent 的任务
 */
export interface AgentTask {
  id: string;
  request: string; // 用户原始请求
  context?: TaskContext; // 任务上下文
  status: AgentTaskStatus;
  steps: AgentStep[];
  result?: string;
  createdAt: Date;
  completedAt?: Date;
  // v2.9.39: 添加 error 属性
  error?: string;
  // v2.7 新增
  dataModel?: DataModel; // 任务的数据模型
  requirementAnalysis?: import("../DataModeler").DataModelAnalysis; // 需求分析结果
  validationErrors?: ExecutionError[]; // 执行后检测到的错误
  rolledBack?: boolean; // 是否已回滚
  // v2.7.2 Agent 成熟化
  executionPlan?: ExecutionPlan; // 任务执行计划
  goals?: TaskGoal[]; // 任务目标列表
  reflection?: TaskReflection; // 执行后反思
  replanHistory?: ReplanResult[]; // replan 历史
  sampleValidation?: SampleValidationResult; // 抽样校验结果
  // v2.8.7 问题追踪
  discoveredIssues?: DiscoveredIssue[]; // Agent 发现的问题
  resolvedIssues?: string[]; // 已解决的问题ID
  // v2.9.0 操作历史与回滚
  operationHistory?: OperationRecord[]; // 操作历史
  validationResults?: ValidationCheckResult[]; // 硬逻辑校验结果
  // v2.9.18 质量报告
  qualityReport?: QualityReport; // 质量检查报告
  // v2.9.58: P2 澄清机制
  clarificationContext?: ClarificationContext; // 澄清上下文
}

/**
 * 任务目标 - 程序级完成判断的依据
 */
export interface TaskGoal {
  id: string;
  description: string;
  type: "create_sheet" | "write_data" | "set_formula" | "create_chart" | "format" | "custom";
  target?: {
    sheet?: string;
    range?: string;
    expectedValue?: unknown;
  };
  status: "pending" | "achieved" | "failed" | "skipped";
  verifiedAt?: Date;
  verificationResult?: string;
}

/**
 * 任务反思 - Agent 自我评估
 */
export interface TaskReflection {
  overallScore: number; // 0-10
  goalsAchieved: number;
  goalsFailed: number;
  issuesFound: string[];
  lessonsLearned: string[];
  suggestionsForNext: string[];
  timestamp: Date;
}

/**
 * 用户反馈类型 - v2.8.2
 */
export interface UserFeedback {
  isFeedback: boolean;
  feedbackType?: "simplify" | "merge" | "error" | "redo" | "modify";
  urgency: "high" | "medium" | "low";
  suggestedAction?: string;
  needsClarification?: boolean;
}

/**
 * v2.9.17: 计划确认请求 - 复杂任务需要用户确认计划
 */
export interface PlanConfirmationRequest {
  planId: string;
  taskDescription: string;
  complexity: TaskComplexity;
  estimatedSteps: number;
  estimatedTime: string;
  proposedStructure?: {
    tables: Array<{
      name: string;
      columns: Array<{
        name: string;
        type: "text" | "number" | "date" | "formula";
        description: string;
      }>;
      purpose: string;
    }>;
  };
  questions?: string[]; // 需要用户澄清的问题
  canSkipConfirmation: boolean;
}

/**
 * v2.9.43: 计划确认等待错误
 */
export class PlanConfirmationPendingError extends Error {
  public readonly planPreview: string;

  constructor(planPreview: string) {
    super("计划等待用户确认");
    this.name = "PlanConfirmationPendingError";
    this.planPreview = planPreview;
  }
}

/**
 * 任务上下文 - 当前环境信息
 */
export interface TaskContext {
  environment: string; // 'excel' | 'word' | 'browser' | ...
  environmentState?: unknown; // 环境特定的状态
  conversationHistory?: Array<{ role: string; content: string }>;
  userPreferences?: Record<string, unknown>;
  userFeedback?: UserFeedback;
  selectedData?: unknown;
  workbookInfo?: unknown;
  // v2.9.58: P2 澄清机制需要的上下文信息
  activeSheet?: string;
  selectedRange?: string;
  availableSheets?: string[];
  currentDataModel?: DataModel;
  recentOperations?: string[];
  // v3.0.3: 强制感知获取的数据
  perceivedData?: {
    address: string;
    values: unknown;
    output: string;
    timestamp: Date;
  };
}

/**
 * v2.9.20: 任务进度信息
 */
export interface TaskProgress {
  /** 任务ID */
  taskId: string;
  /** 当前步骤索引 (从1开始) */
  currentStep: number;
  /** 总步骤数 */
  totalSteps: number;
  /** 进度百分比 (0-100) */
  percentage: number;
  /** 当前阶段 */
  phase: "planning" | "execution" | "verification" | "reflection";
  /** 当前阶段描述 */
  phaseDescription: string;
  /** 步骤列表 */
  steps: ProgressStep[];
  /** 预估剩余时间 (秒) */
  estimatedTimeRemaining?: number;
  /** 开始时间 */
  startedAt: Date;
}

/**
 * v2.9.20: 进度步骤
 */
export interface ProgressStep {
  /** 步骤索引 */
  index: number;
  /** 步骤描述 */
  description: string;
  /** 步骤状态 */
  status: "pending" | "running" | "completed" | "failed" | "skipped";
  /** 完成时间 */
  completedAt?: Date;
  /** 持续时间 (ms) */
  duration?: number;
}

/**
 * v2.9.21: 任务委派
 */
export interface TaskDelegation {
  /** 原始任务 */
  originalTask: string;
  /** 子任务列表 */
  subTasks: Array<{
    id: string;
    description: string;
    assignedAgent: string;
    status: "pending" | "running" | "completed" | "failed";
    result?: string;
  }>;
  /** 协调策略 */
  coordinationStrategy: "sequential" | "parallel" | "conditional";
  /** 总体状态 */
  overallStatus: "pending" | "running" | "completed" | "failed";
}

/**
 * v2.9.21: 思考链步骤
 */
export interface ChainOfThoughtStep {
  id: string;
  /** 子问题 */
  subQuestion: string;
  /** 推理过程 */
  reasoning: string;
  /** 结论 */
  conclusion: string;
  /** 置信度 (0-100) */
  confidence: number;
  /** 依赖的前置步骤 */
  dependsOn: string[];
  /** 状态 */
  status: "pending" | "thinking" | "completed" | "failed";
}

/**
 * v2.9.21: 思考链结果
 */
export interface ChainOfThoughtResult {
  /** 原始问题 */
  originalQuestion: string;
  /** 思考链步骤 */
  steps: ChainOfThoughtStep[];
  /** 最终结论 */
  finalConclusion: string;
  /** 总体置信度 */
  overallConfidence: number;
  /** 思考时间 (ms) */
  thinkingTime: number;
}

/**
 * v2.9.21: 自我提问
 */
export interface SelfQuestion {
  /** 问题 */
  question: string;
  /** 问题类型 */
  type: "clarification" | "prerequisite" | "verification" | "exploration";
  /** 优先级 */
  priority: "high" | "medium" | "low";
  /** 是否已回答 */
  answered: boolean;
  /** 回答 */
  answer?: string;
}

/**
 * v2.9.21: 数据洞察
 */
export interface DataInsight {
  /** 洞察ID */
  id: string;
  /** 洞察类型 */
  type: "trend" | "outlier" | "pattern" | "correlation" | "missing" | "suggestion";
  /** 标题 */
  title: string;
  /** 描述 */
  description: string;
  /** 置信度 (0-100) */
  confidence: number;
  /** 相关数据位置 */
  location?: string;
  /** 建议的操作 */
  suggestedAction?: string;
  /** 是否已展示给用户 */
  presented: boolean;
  /** 用户是否采纳 */
  accepted?: boolean;
}

/**
 * v2.9.21: 预见性建议
 */
export interface ProactiveSuggestion {
  /** 建议ID */
  id: string;
  /** 建议类型 */
  type: "next_step" | "optimization" | "related_task" | "best_practice";
  /** 建议内容 */
  suggestion: string;
  /** 触发条件 */
  trigger: string;
  /** 置信度 */
  confidence: number;
  /** 是否已展示 */
  presented: boolean;
  /** 用户是否采纳 */
  accepted?: boolean;
}
