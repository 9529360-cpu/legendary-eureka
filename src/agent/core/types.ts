/**
 * 核心类型定义 - 从 AgentCore.ts 抽取
 *
 * 单一职责：定义 Agent 系统的核心接口和类型
 * 行数上限：400 行
 */

// ========== 工具相关类型 ==========

/**
 * 工具定义
 */
export interface Tool {
  name: string;
  description: string;
  category: string;
  parameters: ToolParameter[];
  execute: (input: Record<string, unknown>) => Promise<ToolResult>;
}

/**
 * 工具参数定义
 */
export interface ToolParameter {
  name: string;
  type: "string" | "number" | "boolean" | "array" | "object";
  description: string;
  required: boolean;
  default?: unknown;
  enum?: unknown[];
}

/**
 * 工具执行结果
 */
export interface ToolResult {
  success: boolean;
  output?: unknown;
  error?: string;
  metadata?: Record<string, unknown>;
}

// ========== Agent 决策相关类型 ==========

/**
 * Agent 决策
 */
export interface AgentDecision {
  type: "tool_call" | "respond" | "clarify" | "plan" | "reflect";
  toolName?: string;
  toolInput?: Record<string, unknown>;
  response?: string;
  clarificationQuestion?: string;
  plan?: AgentStep[];
  confidence?: number;
  reasoning?: string;
}

/**
 * Agent 执行步骤
 */
export interface AgentStep {
  id: string;
  order: number;
  description: string;
  action: string;
  parameters: Record<string, unknown>;
  dependsOn: string[];
  status: "pending" | "running" | "completed" | "failed" | "skipped";
  result?: ToolResult;
  startTime?: number;
  endTime?: number;
}

/**
 * Agent 任务
 */
export interface AgentTask {
  id: string;
  userMessage: string;
  status: "pending" | "planning" | "executing" | "completed" | "failed" | "cancelled";
  steps: AgentStep[];
  createdAt: number;
  completedAt?: number;
  error?: string;
  result?: string;
  context?: TaskContext;
}

/**
 * 任务上下文
 */
export interface TaskContext {
  workbookName?: string;
  activeSheet?: string;
  selectedRange?: string;
  previousSteps?: AgentStep[];
  userPreferences?: UserPreferences;
}

/**
 * 用户偏好
 */
export interface UserPreferences {
  noDragging: boolean; // 不拖拽
  autoExpand: boolean; // 自动扩展
  multiUserSafe: boolean; // 多人协作安全
  longTermStable: boolean; // 长期稳定
  preferArrayFormula: boolean; // 优先数组公式
}

// ========== 语义抽取类型 ==========

/**
 * 语义抽取结果
 */
export interface SemanticExtraction {
  /** 用户意图 */
  intent: IntentType;

  /** 实体信息 */
  entities: ExtractedEntities;

  /** 约束条件 */
  constraints: ExtractedConstraints;

  /** 置信度 */
  confidence: number;

  /** 原始用户输入 */
  rawInput: string;
}

/**
 * 意图类型
 */
export type IntentType =
  | "cross_sheet_summary" // 跨表汇总
  | "auto_fill" // 自动填充
  | "diagnose_zero" // 结果为0排查
  | "fix_circular_ref" // 循环引用修复
  | "dropdown_cascade" // 下拉联动
  | "data_cleanup" // 数据清洗
  | "formula_creation" // 公式创建
  | "format_range" // 格式化范围
  | "chart_creation" // 图表创建
  | "data_analysis" // 数据分析
  | "structure_refactor" // 结构重构
  | "create_formula" // 创建公式
  | "format" // 格式化
  | "clean_data" // 清洗数据
  | "create_chart" // 创建图表
  | "analyze" // 分析数据
  | "diagnose" // 诊断问题
  | "debug" // 调试
  | "troubleshoot" // 排查问题
  | "read" // 读取数据
  | "write" // 写入数据
  | "filter" // 筛选
  | "sort" // 排序
  | "unknown"; // 未知意图

/**
 * 抽取的实体
 */
export interface ExtractedEntities {
  columns?: string[]; // 列名
  sheets?: string[]; // 工作表名
  ranges?: string[]; // 范围
  metrics?: string[]; // 指标名
  numerator?: string; // 分子
  denominator?: string; // 分母
  dateGranularity?: "day" | "week" | "month" | "quarter" | "year";
  values?: unknown[]; // 具体值
}

/**
 * 抽取的约束
 */
export interface ExtractedConstraints {
  noDragging: boolean; // 不拖拽
  noManualCopy: boolean; // 不手动复制
  autoExpand: boolean; // 自动扩展
  crossFile: boolean; // 跨文件
  multiUser: boolean; // 多人协作
  reusable: boolean; // 可复用
  longTermStable: boolean; // 长期稳定
  urgent: boolean; // 紧急
  preserveFormat: boolean; // 保持格式
  readOnly: boolean; // 只读
  noCode: boolean; // 不用代码
}

// ========== 诊断相关类型 ==========

/**
 * 诊断结果
 */
export interface DiagnosticResult {
  /** 可能原因（按可能性排序） */
  possibleCauses: DiagnosticCause[];

  /** 推荐验证步骤 */
  validationSteps: ValidationStep[];

  /** 推荐修复方案 */
  recommendedFix: string;

  /** 风险说明 */
  riskNotes: string[];
}

/**
 * 诊断原因
 */
export interface DiagnosticCause {
  rank: number;
  cause: string;
  probability: number;
  shortestValidation: string;
}

/**
 * 验证步骤
 */
export interface ValidationStep {
  order: number;
  description: string;
  formula?: string;
  expectedResult?: string;
}

// ========== 解决方案类型 ==========

/**
 * 分层解决方案
 */
export interface LayeredSolution {
  /** 方案1: 最低改动（立刻能跑） */
  minimal: SolutionOption;

  /** 方案2: 推荐方案（长期稳定） */
  recommended: SolutionOption;

  /** 方案3: 结构重构（当表设计不合理时） */
  structural?: SolutionOption;
}

/**
 * 方案层级
 */
export type SolutionTier = "minimal" | "recommended" | "structural";

/**
 * 方案选项
 */
export interface SolutionOption {
  tier: SolutionTier;
  emoji: string;
  title: string;
  description: string;
  steps?: string[];
  code?: string;
  pros?: string[];
  cons?: string[];
}

// ========== 澄清相关类型 ==========

/**
 * 澄清请求
 */
export interface ClarificationRequest {
  type: string;
  message: string;
  suggestions: string[];
  context?: Record<string, unknown>;
}

// ========== 事件类型 ==========

/**
 * Agent 事件类型
 */
export type AgentEventType =
  | "intent:parsed"
  | "plan:created"
  | "step:started"
  | "step:completed"
  | "step:failed"
  | "task:completed"
  | "task:failed"
  | "clarification:needed"
  | "diagnostic:result"
  | "solution:ready";

/**
 * Agent 事件
 */
export interface AgentEvent {
  type: string;
  data: Record<string, unknown>;
  timestamp: number;
}

// ========== 导出默认用户偏好 ==========

export const DEFAULT_USER_PREFERENCES: UserPreferences = {
  noDragging: true,
  autoExpand: true,
  multiUserSafe: true,
  longTermStable: true,
  preferArrayFormula: true,
};
