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
} from "./tool";

// ========== 任务相关类型 ==========
export type {
  AgentTask,
  AgentTaskStatus,
  AgentStep,
  AgentDecision,
  TaskContext,
  TaskGoal,
  TaskReflection,
  TaskComplexity,
  TaskProgress,
  ProgressStep,
  TaskDelegation,
  PlanConfirmationRequest,
  ChainOfThoughtStep,
  ChainOfThoughtResult,
  SelfQuestion,
  DataInsight,
  ProactiveSuggestion,
  LLMGeneratedPlan,
  ClarificationContext,
  ClarificationCheckResult,
  UserFeedback,
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
  QualityIssue,
  RetryStrategy,
  SelfHealingAction,
  ErrorRecoveryStrategy,
  ErrorRecoveryResult,
  ErrorRootCauseAnalysis,
  CriticalErrorResult,
  HypothesisValidation,
  UncertaintyQuantification,
  CounterfactualReasoning,
  LegacyReflectionResult,
} from "./validation";

// ========== 配置相关类型 ==========
export type {
  AgentConfig,
  InteractionConfig,
  ValidationConfig,
  PersistenceConfig,
  ReflectionConfig,
  ValidationSignalConfig,
  ResponseSimplificationConfig,
  ConfirmationConfig,
  FriendlyError,
  ExpertAgentConfig,
  ExpertAgentType,
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
  PlanStep,
  ExecutionPlan,
  PlanExecutionResult,
  // TaskProgress 和 ProgressStep 已从 task.ts 导出
  AgentStreamData,
  AgentOutputData,
  AgentStreamStructuredOutputData,
  TaskStartEventData,
  TaskCompleteEventData,
  TaskErrorEventData,
  TaskPendingEventData,
  PlanGeneratedEventData,
} from "./workflow";

// ========== 意图相关类型 (v4.0) ==========
export type {
  IntentType,
  IntentSpec,
  IntentSpecData,
  CreateTableSpec,
  WriteDataSpec,
  FormatSpec,
  FormulaSpec,
  ChartSpec,
  SheetSpec,
  DataOperationSpec,
  QuerySpec,
  ClarifySpec,
  RespondSpec,
  ColumnDefinition,
} from "./intent";
