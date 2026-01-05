/**
 * Agent 模块入口
 *
 * 架构理念：
 * ┌───────────────────────────────────────────────────────┐
 * │                    Agent 系统                         │
 * │                                                       │
 * │  ┌─────────────────────────────────────────────────┐ │
 * │  │              AgentCore (大脑)                   │ │
 * │  │  - ReAct 循环引擎                               │ │
 * │  │  - 工具注册中心                                 │ │
 * │  │  - 记忆系统                                     │ │
 * │  └─────────────────────────────────────────────────┘ │
 * │                        │                              │
 * │                        ▼                              │
 * │  ┌─────────────────────────────────────────────────┐ │
 * │  │           Tool Adapters (手臂)                  │ │
 * │  │                                                 │ │
 * │  │   ┌───────────┐  ┌───────────┐  ┌───────────┐ │ │
 * │  │   │  Excel    │  │   Word    │  │    API    │ │ │
 * │  │   │  Adapter  │  │  Adapter  │  │  Adapter  │ │ │
 * │  │   └───────────┘  └───────────┘  └───────────┘ │ │
 * │  │                                                 │ │
 * │  │   ┌───────────┐  ┌───────────┐  ┌───────────┐ │ │
 * │  │   │ Browser   │  │   File    │  │    ...    │ │ │
 * │  │   │  Adapter  │  │  Adapter  │  │           │ │ │
 * │  │   └───────────┘  └───────────┘  └───────────┘ │ │
 * │  └─────────────────────────────────────────────────┘ │
 * └───────────────────────────────────────────────────────┘
 *
 * 使用方式：
 *
 * ```typescript
 * import { Agent, createExcelTools } from './agent';
 *
 * // 创建 Agent
 * const agent = new Agent();
 *
 * // 注册 Excel 工具（第一站）
 * agent.registerTools(createExcelTools());
 *
 * // 未来可以注册更多工具
 * // agent.registerTools(createWordTools());
 * // agent.registerTools(createApiTools());
 *
 * // 执行任务
 * const result = await agent.run("分析这个数据并创建图表", {
 *   environment: "excel"
 * });
 * ```
 */

// 核心导出
export { Agent, ToolRegistry, AgentMemory, getAgent, createAgent } from "./AgentCore";

// 类型导出
export type {
  Tool,
  ToolParameter,
  ToolResult,
  AgentDecision,
  AgentStep,
  AgentTask,
  TaskContext,
  AgentConfig,
  CriticalErrorResult,
  TaskGoal,
  TaskReflection,
} from "./AgentCore";

// 适配器导出
export { default as createExcelTools, createExcelReader } from "./ExcelAdapter";

// v2.7 新增: 数据建模和验证导出
export { DataModeler } from "./DataModeler";
export type {
  DataModel,
  TableDefinition,
  FieldDefinition,
  DataModelAnalysis,
  ValidationResult as ModelValidationResult,
} from "./DataModeler";

export { FormulaValidator } from "./FormulaValidator";
export type {
  ExecutionError,
  FormulaValidationResult,
  FixSuggestion,
  SampleValidationResult,
  SampleValue,
  SampleIssue,
} from "./FormulaValidator";

// v2.9.53 新增: 公式编译器
export { FormulaCompiler, formulaCompiler } from "./FormulaCompiler";
export type {
  CompileMode,
  CompileContext,
  CompileResult,
  CompileError,
  CompileWarning,
  FieldMapping,
  TableSchema,
} from "./FormulaCompiler";

export { TaskPlanner, taskPlanner } from "./TaskPlanner";
export type {
  ExecutionPlan,
  PlanStep,
  TaskType,
  ExecutionPhase,
  ReplanContext,
  ReplanResult,
  ReplanStrategy,
  FailureAnalysis,
  FailureType,
} from "./TaskPlanner";

export { ExecutionEngine, executionEngine } from "./ExecutionEngine";
export type {
  ExecutionCallbacks,
  ExecutionOptions,
  ExecutionResult,
  ExcelOperations,
  WorkbookSnapshot,
  InterruptionType,
  Interruption,
  InterruptionHandler,
  InterruptionResponse,
  PreExecutionCheckResult,
} from "./ExecutionEngine";

// v2.9.7 新增: 计划验证器
export { PlanValidator, planValidator } from "./PlanValidator";
export type {
  PlanValidationResult,
  PlanValidationError,
  PlanValidationWarning,
  PlanValidationRule,
  WorkbookContext,
} from "./PlanValidator";

// v2.9.54 数据验证器（重构）
export { DataValidator, dataValidator, ColumnResolver } from "./DataValidator";
export type {
  DataValidationResult,
  DataValidationIssue,
  DataValidationContext,
  DataValidationRule,
  CanonicalColumn,
  ColumnRole,
  ResolvedColumns,
  ConfidenceLevel,
  SamplingStrategy,
  SampleData,
  ValidationEvidence,
  FixAction,
} from "./DataValidator";

// v2.9.58: P2 澄清机制
export { IntentAnalyzer, intentAnalyzer } from "./IntentAnalyzer";
export type {
  IntentAnalysis,
  IntentType,
  ParsedEntities,
  VagueReference,
  ClarificationItem,
  ClarificationType,
  AmbiguityInfo,
  SuggestedClarification,
  ClarificationOption,
  SuggestedPlan,
  AnalysisContext,
} from "./IntentAnalyzer";

export { ClarificationEngine, clarificationEngine } from "./ClarificationEngine";
export type {
  ClarificationSession,
  ClarificationTurn,
  CollectedInfo,
  ClarificationResult,
} from "./ClarificationEngine";

// AgentCore 中的澄清相关类型
export type {
  InteractionConfig,
  ClarificationContext,
  ClarificationCheckResult,
} from "./AgentCore";

// v2.9.58: P0 每步反思机制
export { StepReflector, stepReflector, DEFAULT_REFLECTION_CONFIG } from "./StepReflector";
export type {
  ReflectionResult,
  ReflectionAction,
  ReflectionIssue,
  PlanAdjustment,
  Opportunity,
  ReflectionContext,
  ReflectionConfig as StepReflectionConfig,
} from "./StepReflector";

// v2.9.58: P1 验证信号系统
export {
  ValidationSignalHandler,
  validationSignalHandler,
  DEFAULT_SIGNAL_CONFIG,
} from "./ValidationSignal";
export type {
  ValidationSignal,
  ValidationSignalType,
  SignalContext,
  SuggestedAction,
  ActionType,
  SignalStatus,
  SignalResolution,
  SignalDecision,
  ValidationSignalConfig,
} from "./ValidationSignal";

// ========== v2.9.59: 协议版组件 ==========

// AgentProtocol - 单一事实来源
export {
  createSignal,
  validationOk,
  validationFail,
  hasBlockingSignals,
  inferRecommendedAction,
  SignalCodes,
} from "./AgentProtocol";
export type {
  SignalLevel,
  Signal,
  RecommendedAction,
  ValidationOutput,
  NextAction,
  ClarifyQuestion,
  RiskSummary,
  StepDecision,
  StepFix,
  AgentReply,
  AgentReplyDebug,
} from "./AgentProtocol";

// P1: 信号收集器
export {
  safeValidate,
  safeValidateSync,
  collectStepSignals,
  collectPlanSignals,
} from "./validators/collectSignals";

// P2: 澄清门
export {
  ClarifyGate,
  clarifyGate,
  needClarify,
  getNextAction,
  DEFAULT_CLARIFY_CONFIG,
} from "./ClarifyGate";
export type { ClarifyGateConfig } from "./ClarifyGate";
// 注意：WorkbookContext 在 PlanValidator 中也有定义，这里用别名避免冲突
export type { WorkbookContext as ClarifyWorkbookCtx } from "./ClarifyGate";

// P0: 步骤决策器（协议版）
export {
  StepDecider,
  stepDecider,
  makeDecision,
  makeDecisionSync,
  DEFAULT_DECIDER_CONFIG,
} from "./StepDecider";
export type { DecisionContext, DeciderConfig } from "./StepDecider";

// P3: 响应构建器
export {
  ResponseBuilder,
  responseBuilder,
  buildReply,
  buildReplySync,
  formatReply,
  DEFAULT_RESPONSE_CONFIG,
} from "./ResponseBuilder";
export type { BuildContext, ResponseBuilderConfig } from "./ResponseBuilder";

// ========== v3.0: 审批管理和审计日志 ==========

export {
  ApprovalManager,
  approvalManager,
  DEFAULT_APPROVAL_CONFIG,
  HIGH_RISK_OPERATIONS,
  MEDIUM_RISK_OPERATIONS,
  BATCH_KEYWORDS,
} from "./ApprovalManager";
export type {
  RiskLevel,
  ApprovalStatus,
  RiskAssessment,
  ApprovalRequest,
  ApprovalDecision,
  ApprovalCallback,
  ApprovalManagerConfig,
} from "./ApprovalManager";

export { AuditLogger, auditLogger, DEFAULT_AUDIT_CONFIG } from "./AuditLogger";
export type {
  AuditEntry,
  AuditAction,
  AuditResult,
  AuditVerifyResult,
  AuditLoggerConfig,
} from "./AuditLogger";

// ========== v3.1: 从 Activepieces 借鉴的模式 ==========

// 不可变执行上下文
export {
  ExecutionContext,
  DEFAULT_EXECUTION_CONFIG,
  createStepOutput,
  markStepSucceeded,
  markStepFailed,
} from "./ExecutionContext";
export type {
  VerdictStatus,
  ExecutionVerdict,
  FailedStep,
  StepOutputStatus,
  StepOutput,
  ExecutionContextConfig,
} from "./ExecutionContext";

// 指数退避重试处理器
export {
  runWithExponentialBackoff,
  withRetry,
  continueIfFailureHandler,
  robustExecute,
  API_RETRY_OPTIONS,
  EXCEL_RETRY_OPTIONS,
  NO_RETRY_OPTIONS,
  DEFAULT_RETRY_OPTIONS,
} from "./RetryHandler";
export type {
  RetryStrategy,
  RetryOptions,
  RetryResult,
  ContinueOnFailureOptions,
} from "./RetryHandler";

// 进度服务（防抖 + 互斥锁）
export {
  ProgressService,
  Mutex,
  getProgressService,
  resetProgressService,
  createProgressTracker,
  DEFAULT_PROGRESS_CONFIG,
} from "./ProgressService";
export type { ProgressInfo, ProgressListener, ProgressServiceConfig } from "./ProgressService";

// ========== v3.2: 从 sv-excel-agent 借鉴的模式 ==========

// 统一工具响应格式
export {
  ToolSuccess,
  ToolError,
  ErrorCodes,
  success,
  error,
  isSuccess,
  isError,
  parseResponse,
  hasFormulaErrors,
  cellResult,
} from "./ToolResponse";
export type {
  ErrorCode,
  ResponseStatus,
  BaseResponse,
  SuccessResponse,
  ErrorResponse,
  CellResult,
} from "./ToolResponse";

// 公式翻译器
export {
  columnLetterToNumber,
  numberToColumnLetter,
  translateFormula,
  isArrayFormula,
  parseCellAddress,
  parseRangeAddress,
  buildCellAddress,
  buildRangeAddress,
} from "./FormulaTranslator";

// ========== v3.3: 从 ai-agents-for-beginners 学习的模式 ==========

// Zod 结构化输出验证 (第7课 Planning Design)
export {
  LLMResponseValidator,
  validateLLMResponse,
  validateExecutionPlan,
  createExecutionPlan,
  ToolCallSchema,
  ExecutionStepSchema,
  ExecutionPlanSchema,
  LLMResponseSchema,
  RiskLevelSchema,
  OperationTypeSchema,
  ClarificationRequestSchema,
  RejectionResponseSchema,
  CompletionResponseSchema,
} from "./LLMResponseValidator";
export type {
  ToolCall,
  ExecutionStep as ValidatedStep,
  ExecutionPlan as ValidatedPlan,
  ClarificationRequest,
  RejectionResponse,
  CompletionResponse,
  LLMResponse,
  ValidationResult as ZodValidationResult,
  RiskLevel as ValidatedRiskLevel,
  OperationType,
  RepairRetryConfig,
  RepairResult,
} from "./LLMResponseValidator";

// 上下文压缩器 (第12课 Context Engineering)
export {
  ContextCompressor,
  createCompressor,
  compressMessages,
  estimateTokenCount,
} from "./ContextCompressor";
export type {
  CompressionConfig,
  CompressionResult,
  DualCompressionResult,
} from "./ContextCompressor";

// 动态工具选择器 (第12课 Context Engineering - 防止 Context Confusion)
export {
  ToolSelector,
  createToolSelector,
  selectRelevantTools,
  categorizeTools,
} from "./ToolSelector";
export type {
  ToolCategory,
  ToolMetadata,
  SelectionConfig,
  SelectionResult,
  LLMToolSubset,
} from "./ToolSelector";

// 分层 System Message 构建器 (第6课 Building Trustworthy Agents)
export {
  SystemMessageBuilder,
  createSystemMessageBuilder,
  buildSystemMessage,
  PRESET_CONFIGS,
} from "./SystemMessageBuilder";
export type { SystemMessageLayer, BuilderConfig, ToolDescription } from "./SystemMessageBuilder";

// 情景记忆 (第13课 Agent Memory)
export {
  EpisodicMemory,
  getEpisodicMemory,
  createEpisodicMemory,
  recordSuccess,
  recordFailure,
} from "./EpisodicMemory";
export type {
  Episode,
  EpisodeStep,
  PatternAnalysis,
  EpisodicMemoryConfig,
  ReusableExperience,
  UserPreference,
  FailureReason,
  ValidParameters,
  TaskPattern,
} from "./EpisodicMemory";

// 自我反思机制 (第9课 Metacognition)
export {
  SelfReflection,
  getSelfReflection,
  createSelfReflection,
  validatePlan,
  isPlanExecutable,
} from "./SelfReflection";
export type {
  ReflectionResult as SelfReflectionResult,
  ReflectionIssue as SelfReflectionIssue,
  ReflectionConfig,
  ToolRegistry as ReflectionToolRegistry,
  HardRuleViolation,
  HardRuleValidationResult,
} from "./SelfReflection";

// ========== 便捷工厂函数 ==========

import { Agent } from "./AgentCore";
import createExcelTools, { createExcelReader } from "./ExcelAdapter";

/**
 * 创建一个配置好 Excel 工具的 Agent (v2.9.7)
 *
 * Agent 闭环能力:
 * ┌─────────────────────────────────────────────────────┐
 * │  THINK ──────→ EXECUTE ──────→ OBSERVE             │
 * │    │              │              │                  │
 * │    ▼              ▼              ▼                  │
 * │ 计划验证       数据校验       智能回滚              │
 * │ (5条规则)      (6条规则)     (确定性)               │
 * └─────────────────────────────────────────────────────┘
 */
export function createExcelAgent(config?: { verbose?: boolean }): Agent {
  const agent = new Agent({
    maxIterations: 30, // v2.7: 复杂任务需要更多迭代
    enableMemory: true,
    verboseLogging: config?.verbose ?? false,
  });

  // 注册 Excel 工具
  agent.registerTools(createExcelTools());

  // v2.9.5: 注入 ExcelReader，让硬校验能够读取 Excel 数据
  agent.setExcelReader(createExcelReader());

  return agent;
}

/**
 * 全局 Excel Agent 单例
 */
let excelAgent: Agent | null = null;

export function getExcelAgent(): Agent {
  if (!excelAgent) {
    excelAgent = createExcelAgent();
  }
  return excelAgent;
}
