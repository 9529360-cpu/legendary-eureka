/**
 * Agent 类型定义统一导出
 *
 * 此模块从 AgentCore.ts 中抽取的所有类型定义
 * 保持向后兼容性，所有类型都可以从 AgentCore.ts 或此模块导入
 *
 * @packageDocumentation
 */

// ========== 工具相关类型 ==========
export type {
  Tool,
  ToolParameter,
  ToolResult,
  ToolChain,
  ToolResultValidation,
  ToolCallInfo,
  ToolCallResultData,
  ToolDependency,
  ToolCapability,
  ToolExecutionOptions,
} from "./tool";

// ========== 任务相关类型 ==========
export type {
  AgentTask,
  AgentStep,
  AgentDecision,
  TaskContext,
  TaskGoal,
  TaskReflection,
  PlanConfirmationRequest,
  ChainOfThoughtStep,
  DataInsight,
  ProactiveSuggestion,
  ClarificationQuestion,
  ConversationContext,
  ExecutionHistoryEntry,
  MessageEntry,
  TaskExecutionSummary,
  TaskMetrics,
  TaskDependency,
  SubTask,
  TaskPriority,
} from "./task";

// ========== 验证相关类型 ==========
export type {
  HardValidationRule,
  ExcelReader,
  ValidationContext,
  ValidationCheckResult,
  DiscoveredIssue,
  OperationRecord,
  QualityReport,
  RetryStrategy,
  SelfHealingAction,
  ErrorRecoveryStrategy,
  ValidationResult,
  ValidationError,
  ValidationWarning,
  ValidationSeverity,
} from "./validation";

// ========== 配置相关类型 ==========
export type {
  AgentConfig,
  InteractionConfig,
  ValidationConfig,
  PersistenceConfig,
  ReflectionConfig,
  FriendlyError,
  ExpertAgentConfig,
  AgentMode,
  LogLevel,
  ExecutionMode,
} from "./config";

// ========== 记忆相关类型 ==========
export type {
  UserPreferences,
  UserProfile,
  CompletedTask,
  TaskPattern,
  CachedWorkbookContext,
  CachedSheetInfo,
  LearnedPreference,
  RecentOperation,
  SemanticMemoryEntry,
  UserFeedbackRecord,
  LearnedPattern,
} from "./memory";

// ========== 工作流相关类型 ==========
export type {
  WorkflowEvent,
  WorkflowState,
  WorkflowEventHandler,
  SimpleWorkflow,
  WorkflowContextInterface,
  WorkflowEventStreamInterface,
  WorkflowEventRegistryInterface,
  WorkflowEventFactory,
  AgentStreamData,
  AgentOutputData,
  AgentStreamStructuredOutputData,
  // 计划执行
  PlanStep,
  ExecutionPlan,
  PlanExecutionResult,
  // 进度
  TaskProgress,
  ProgressStep,
  // 事件数据类型
  TaskStartEventData,
  TaskCompleteEventData,
  TaskErrorEventData,
  TaskPendingEventData,
  PlanGeneratedEventData,
  PlanStepStartEventData,
  PlanStepCompleteEventData,
  PlanConfirmationRequiredEventData,
  ProgressUpdateEventData,
  ValidationStartEventData,
  ValidationCompleteEventData,
  CotStepStartEventData,
  CotStepCompleteEventData,
  FollowUpContextSetEventData,
  FollowUpHandledEventData,
} from "./workflow";
