/**
 * Excel 智能助手 核心类型定义
 * 遵循架构约束：严格区分模块职责，确保类型安全
 */

// ==================== 基础类型 ====================

/**
 * Excel 单元格值类型
 */
export type CellValue = string | number | boolean | Date | null;

/**
 * 消息角色类型
 */
export type MessageRole = "user" | "assistant" | "system";

/**
 * 对话消息
 */
export interface ConversationMessage {
  id: string;
  role: MessageRole;
  content: string;
  timestamp: Date;
  metadata?: {
    operation?: string;
    confidence?: number;
    parameters?: Record<string, unknown>;
    executionResult?: ExecutionResult;
  };
}

/**
 * 用户意图
 */
export interface UserIntent {
  type: IntentType;
  confidence: number;
  parameters: IntentParameters;
  rawInput: string;
}

/**
 * 意图类型
 */
export type IntentType =
  | "create_table"
  | "format_cells"
  | "create_chart"
  | "insert_data"
  | "apply_filter"
  | "insert_formula"
  | "sort_data"
  | "clear_range"
  | "copy_range"
  | "merge_cells"
  | "analyze_data"
  | "generate_summary"
  | "unknown";

/**
 * 意图参数
 */
export interface IntentParameters {
  range?: string;
  data?: CellValue[][];
  format?: CellFormat;
  chartType?: ChartType;
  formula?: string;
  condition?: FilterCondition;
  headers?: string[];
  sortBy?: string;
  ascending?: boolean;
  targetRange?: string;
  [key: string]: unknown;
}

// ==================== Excel 操作类型 ====================

/**
 * Excel 操作定义
 */
export interface ExcelOperation {
  id: string;
  type: OperationType;
  description: string;
  parameters: OperationParameters;
  validationRules: ValidationRule[];
  executable: boolean;
}

/**
 * 操作类型
 */
export type OperationType =
  | "insert_data"
  | "create_table"
  | "format_cells"
  | "format_range"
  | "create_chart"
  | "apply_filter"
  | "insert_formula"
  | "sort_data"
  | "clear_range"
  | "copy_range"
  | "merge_cells"
  | "set_cell_value"
  | "set_range_values"
  | "select_range"
  | "get_range_values";

/**
 * 操作参数
 */
export interface OperationParameters {
  range: string;
  data?: CellValue[][];
  format?: CellFormat;
  chartType?: ChartType;
  formula?: string;
  condition?: FilterCondition;
  sortBy?: string;
  ascending?: boolean;
  targetRange?: string;
}

/**
 * 单元格格式
 */
export interface CellFormat {
  fillColor?: string;
  fontColor?: string;
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  numberFormat?: string;
  horizontalAlignment?: "left" | "center" | "right";
  verticalAlignment?: "top" | "middle" | "bottom";
  borders?: BorderStyle[];
}

/**
 * 图表类型
 */
export type ChartType = "column" | "line" | "pie" | "bar" | "area" | "scatter";

/**
 * 筛选条件
 */
export interface FilterCondition {
  column: number;
  operator: "equals" | "contains" | "greaterThan" | "lessThan" | "between";
  value: string | number | boolean | Date;
  value2?: string | number | boolean | Date;
}

/**
 * 边框样式
 */
export interface BorderStyle {
  position: "top" | "right" | "bottom" | "left";
  style: "thin" | "medium" | "thick" | "dashed" | "dotted";
  color: string;
}

// ==================== 工具调用类型 ====================

/**
 * 工具调用请求
 */
export interface ToolCallRequest {
  toolName: string;
  parameters: Record<string, unknown>;
  context: ToolCallContext;
}

/**
 * 工具调用上下文
 */
export interface ToolCallContext {
  conversationId: string;
  userId: string;
  workbookContext: WorkbookContext;
  previousCalls: ToolCallResult[];
}

/**
 * 工具调用结果
 */
export interface ToolCallResult {
  success: boolean;
  toolName: string;
  result: unknown;
  error?: string;
  executionTime: number;
  timestamp: Date;
}

/**
 * 工作簿上下文
 */
export interface WorkbookContext {
  activeSheet: string;
  selectedRange: string;
  workbookName: string;
  sheetNames: string[];
  usedRange?: {
    address: string;
    rowCount: number;
    columnCount: number;
  };
}

// ==================== 验证类型 ====================

/**
 * 验证条件类型
 */
export interface ValidationCondition {
  min?: number;
  max?: number;
  pattern?: string;
  validator?: (value: unknown) => boolean;
}

/**
 * 验证规则
 */
export interface ValidationRule {
  field: string;
  type: "required" | "range" | "format" | "regex" | "custom";
  condition?: ValidationCondition;
  message: string;
}

/**
 * 验证结果
 */
export interface ValidationResult {
  isValid: boolean;
  errors: ValidationError[];
  warnings: ValidationWarning[];
}

/**
 * 验证错误
 */
export interface ValidationError {
  field: string;
  message: string;
  code: string;
}

/**
 * 验证警告
 */
export interface ValidationWarning {
  field: string;
  message: string;
  severity: "low" | "medium" | "high";
}

// ==================== 执行类型 ====================

/**
 * 执行计划
 */
export interface ExecutionPlan {
  id: string;
  operations: ExcelOperation[];
  dependencies: OperationDependency[];
  estimatedTime: number;
  riskLevel: "low" | "medium" | "high";
  validationResults: ValidationResult[];
}

/**
 * 操作依赖
 */
export interface OperationDependency {
  from: string;
  to: string;
  type: "data" | "range" | "format";
}

/**
 * 执行结果
 */
export interface ExecutionResult {
  success: boolean;
  operationId: string;
  result: unknown;
  error?: string;
  executionTime: number;
  affectedRange?: string;
  warnings?: string[];
}

/**
 * 批量执行结果
 */
export interface BatchExecutionResult {
  total: number;
  successful: number;
  failed: number;
  results: ExecutionResult[];
  totalTime: number;
}

// ==================== AI 相关类型 ====================

/**
 * LLM 请求
 */
export interface LLMRequest {
  messages: ConversationMessage[];
  tools?: ToolDefinition[];
  temperature?: number;
  maxTokens?: number;
}

/**
 * LLM 响应
 */
export interface LLMResponse {
  content: string;
  toolCalls?: ToolCall[];
  reasoning?: string;
  confidence: number;
}

/**
 * 工具类别
 */
export enum ToolCategory {
  EXCEL_OPERATION = "excel_operation",
  WORKSHEET_OPERATION = "worksheet_operation",
  DATA_ANALYSIS = "data_analysis",
  DATA_VISUALIZATION = "data_visualization",
  UTILITY = "utility",
}

/**
 * 参数类型
 */
export enum ParameterType {
  STRING = "string",
  NUMBER = "number",
  BOOLEAN = "boolean",
  ARRAY = "array",
  OBJECT = "object",
  ANY = "any",
}

/**
 * 工具定义（Tool Schema）
 */
export interface ToolDefinition {
  id: string;
  name: string;
  category: ToolCategory;
  description: string;
  parameters: ToolParameter[];
  returns: ToolReturn;
  example?: {
    input: Record<string, any>;
    description: string;
  };
}

/**
 * 参数验证规则
 */
export interface ParameterValidationRule {
  pattern?: string;
  minLength?: number;
  maxLength?: number;
  minItems?: number;
  maxItems?: number;
  minValue?: number;
  maxValue?: number;
  errorMessage?: string;
}

/**
 * 工具参数
 */
export interface ToolParameter {
  name: string;
  type: ParameterType;
  description: string;
  required: boolean;
  enum?: string[];
  default?: string | number | boolean | null;
  validation?: ParameterValidationRule;
  properties?: Record<string, Partial<ToolParameter>>;
}

/**
 * 工具返回类型
 */
export type ToolReturn = string;

/**
 * 工具返回描述
 */
export interface ToolReturnDescription {
  type: ParameterType;
  description: string;
}

/**
 * 工具调用
 */
export interface ToolCall {
  id: string;
  name: string;
  arguments: Record<string, any>;
}

// ==================== 配置类型 ====================

/**
 * 应用配置
 */
export interface AppConfig {
  ai: {
    provider: "deepseek" | "openai" | "azure";
    apiKey: string;
    endpoint: string;
    model: string;
    temperature: number;
    maxTokens: number;
  };
  excel: {
    maxRows: number;
    maxColumns: number;
    defaultFormat: CellFormat;
  };
  security: {
    allowDirectExcelCalls: boolean;
    maxToolCallsPerMinute: number;
    allowedOperations: OperationType[];
  };
  logging: {
    level: "debug" | "info" | "warn" | "error";
    persistConversations: boolean;
    maxLogSize: number;
  };
}

/**
 * 用户配置
 */
export interface UserConfig {
  preferences: {
    language: "zh" | "en";
    theme: "light" | "dark";
    autoSave: boolean;
    confirmBeforeExecute: boolean;
  };
  shortcuts: Record<string, string>;
  recentOperations: string[];
}

// ==================== 错误类型 ====================

/**
 * 应用错误
 */
export interface AppError {
  code: string;
  message: string;
  severity: "error" | "warning" | "info";
  context?: Record<string, any>;
  timestamp: Date;
}

/**
 * 错误代码
 */
export enum ErrorCode {
  // Excel 错误
  EXCEL_NOT_READY = "EXCEL_001",
  INVALID_RANGE = "EXCEL_002",
  OPERATION_FAILED = "EXCEL_003",

  // AI 错误
  AI_SERVICE_UNAVAILABLE = "AI_001",
  INVALID_API_KEY = "AI_002",
  RATE_LIMIT_EXCEEDED = "AI_003",

  // 验证错误
  VALIDATION_FAILED = "VAL_001",
  MISSING_PARAMETER = "VAL_002",
  INVALID_PARAMETER = "VAL_003",

  // 工具错误
  TOOL_NOT_FOUND = "TOOL_001",
  TOOL_EXECUTION_FAILED = "TOOL_002",
  TOOL_VALIDATION_FAILED = "TOOL_003",

  // 执行错误
  EXECUTION_TIMEOUT = "EXEC_001",
  DEPENDENCY_CYCLE = "EXEC_002",
  INSUFFICIENT_PERMISSIONS = "EXEC_003",
}

// ==================== 事件类型 ====================

/**
 * 应用事件
 */
export interface AppEvent {
  type: EventType;
  data: unknown;
  timestamp: Date;
  source: string;
}

/**
 * 事件类型
 */
export type EventType =
  | "conversation_started"
  | "conversation_ended"
  | "tool_called"
  | "tool_executed"
  | "excel_operation_started"
  | "excel_operation_completed"
  | "error_occurred"
  | "config_changed"
  | "user_feedback";

// ==================== 工具函数类型 ====================

/**
 * 工具函数
 */
export type ToolFunction = (
  parameters: Record<string, any>,
  context: ToolCallContext
) => Promise<ToolCallResult>;

/**
 * 工具注册项
 */
export interface ToolRegistryItem {
  definition: ToolDefinition;
  function: ToolFunction;
  validator?: (parameters: Record<string, any>) => ValidationResult;
}

// ==================== Agent Goal & Reflection 类型 ====================

/**
 * 任务目标 - 用于程序级验证任务完成状态
 */
export interface TaskGoal {
  id: string;
  type:
    | "formula_applied"
    | "data_validated"
    | "calculation_verified"
    | "format_applied"
    | "range_populated"
    | "error_free";
  description: string;
  verificationMethod: "cell_check" | "range_check" | "formula_check" | "value_comparison";
  targetRange?: string;
  expectedCondition?: string;
  verified: boolean;
  verificationResult?: string;
}

/**
 * 任务反思 - 执行后自我评估
 */
export interface TaskReflection {
  taskId: string;
  summary: string;
  lessonsLearned: string[];
  improvements: string[];
  selfAssessment: {
    planQuality: number;
    executionAccuracy: number;
    verificationThoroughness: number;
    overallScore: number;
  };
}
