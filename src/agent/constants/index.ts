/**
 * Agent 常量定义
 *
 * 从 AgentCore.ts 抽取的硬编码常量
 *
 * @packageDocumentation
 */

import type {
  FriendlyError,
  RetryStrategy,
  SelfHealingAction,
  InteractionConfig,
} from "../types";
import type { ExpertAgentConfig } from "../types/config";

// ========== 错误映射常量 ==========

/**
 * v2.9.20: 友好错误消息映射
 */
export const FRIENDLY_ERROR_MAP: Record<
  string,
  Omit<FriendlyError, "code" | "originalMessage">
> = {
  RangeNotFound: {
    friendlyMessage: "找不到指定的单元格范围",
    possibleCauses: ["范围地址格式错误", "工作表名称不正确", "指定的单元格不存在"],
    suggestions: ["检查范围地址是否正确（如 A1:B10）", "确认工作表名称无误"],
    autoRecoverable: false,
    severity: "error",
  },
  InvalidFormula: {
    friendlyMessage: "公式有误",
    possibleCauses: ["函数名拼写错误", "参数数量不对", "引用的单元格不存在"],
    suggestions: ["检查函数名称拼写", "确认参数格式正确", "验证引用的单元格存在"],
    autoRecoverable: true,
    severity: "warning",
  },
  PermissionDenied: {
    friendlyMessage: "没有权限执行此操作",
    possibleCauses: ["工作表已保护", "工作簿已锁定", "单元格被锁定"],
    suggestions: ["先解除工作表保护", "检查单元格是否被锁定"],
    autoRecoverable: false,
    severity: "error",
  },
  "#NAME?": {
    friendlyMessage: "公式中的函数名无法识别",
    possibleCauses: ["函数名拼写错误", "使用了当前 Excel 版本不支持的函数"],
    suggestions: ["检查函数名称拼写", "尝试使用兼容的替代函数"],
    autoRecoverable: true,
    severity: "warning",
  },
  "#REF!": {
    friendlyMessage: "公式引用了无效的单元格",
    possibleCauses: ["引用的单元格已被删除", "公式复制时引用超出范围"],
    suggestions: ["检查并修正单元格引用", "使用绝对引用（$A$1）"],
    autoRecoverable: false,
    severity: "error",
  },
  "#VALUE!": {
    friendlyMessage: "公式中的值类型不正确",
    possibleCauses: ["文本与数字混用", "参数类型不匹配"],
    suggestions: ["确保参与计算的都是数字", "使用 VALUE() 函数转换文本"],
    autoRecoverable: true,
    severity: "warning",
  },
  "#DIV/0!": {
    friendlyMessage: "公式尝试除以零",
    possibleCauses: ["除数为零或空单元格"],
    suggestions: ["添加 IF 判断避免除零", "使用 IFERROR 处理错误"],
    autoRecoverable: true,
    severity: "warning",
  },
  "#N/A": {
    friendlyMessage: "查找函数找不到匹配项",
    possibleCauses: ["查找值不存在", "查找范围不正确"],
    suggestions: ["确认查找值存在于数据中", "检查查找范围是否正确"],
    autoRecoverable: false,
    severity: "info",
  },
  Timeout: {
    friendlyMessage: "操作超时",
    possibleCauses: ["数据量太大", "网络连接问题", "Excel 响应慢"],
    suggestions: ["尝试减少数据量", "稍后重试", "检查 Excel 是否正常运行"],
    autoRecoverable: true,
    severity: "warning",
  },
  NetworkError: {
    friendlyMessage: "网络连接出现问题",
    possibleCauses: ["网络不稳定", "服务器暂时不可用"],
    suggestions: ["检查网络连接", "稍后重试"],
    autoRecoverable: true,
    severity: "warning",
  },
};

// ========== 专家 Agent 配置常量 ==========

/**
 * 专家 Agent 类型
 */
export type ExpertAgentType =
  | "data_analyst"
  | "formatter"
  | "formula_expert"
  | "chart_expert"
  | "general";

/**
 * v2.9.21: 专家 Agent 预定义配置
 */
export const EXPERT_AGENTS: Record<ExpertAgentType, ExpertAgentConfig> = {
  data_analyst: {
    type: "data_analyst",
    name: "数据分析专家",
    description: "擅长数据分析、统计和洞察发现",
    specialties: ["数据分析", "统计", "趋势识别", "异常检测"],
    tools: ["sample_rows", "excel_read_range", "analyze_data"],
    systemPromptAddition: "你是数据分析专家，专注于发现数据中的模式和洞察。",
  },
  formatter: {
    type: "formatter",
    name: "格式化专家",
    description: "擅长表格格式化和视觉美化",
    specialties: ["格式化", "样式", "条件格式", "表格美化"],
    tools: ["excel_format_range", "set_conditional_format", "auto_fit_columns"],
    systemPromptAddition: "你是格式化专家，专注于让表格更美观易读。",
  },
  formula_expert: {
    type: "formula_expert",
    name: "公式专家",
    description: "擅长复杂公式和函数设计",
    specialties: ["公式", "函数", "计算逻辑", "数组公式"],
    tools: ["excel_set_formula", "validate_formula", "suggest_formula"],
    systemPromptAddition: "你是公式专家，专注于设计高效准确的计算公式。",
  },
  chart_expert: {
    type: "chart_expert",
    name: "图表专家",
    description: "擅长数据可视化和图表设计",
    specialties: ["图表", "可视化", "数据展示", "仪表盘"],
    tools: ["excel_create_chart", "modify_chart", "add_trendline"],
    systemPromptAddition: "你是图表专家，专注于用合适的图表展示数据。",
  },
  general: {
    type: "general",
    name: "通用助手",
    description: "处理一般性任务",
    specialties: ["通用"],
    tools: [],
    systemPromptAddition: "",
  },
};

// ========== 重试策略常量 ==========

/**
 * v2.9.22: 预定义重试策略
 */
export const RETRY_STRATEGIES: Record<string, RetryStrategy> = {
  default: {
    id: "default",
    maxRetries: 3,
    backoffType: "exponential",
    initialDelayMs: 1000,
    maxDelayMs: 10000,
    retryableErrors: ["timeout", "network", "rate_limit"],
  },
  aggressive: {
    id: "aggressive",
    maxRetries: 5,
    backoffType: "linear",
    initialDelayMs: 500,
    maxDelayMs: 5000,
    retryableErrors: ["timeout", "network", "rate_limit", "api_error"],
    transformBeforeRetry: "simplify",
  },
  conservative: {
    id: "conservative",
    maxRetries: 2,
    backoffType: "fixed",
    initialDelayMs: 2000,
    maxDelayMs: 2000,
    retryableErrors: ["timeout", "network"],
  },
};

// ========== 自愈动作常量 ==========

/**
 * v2.9.22: 预定义自愈动作
 */
export const SELF_HEALING_ACTIONS: SelfHealingAction[] = [
  {
    id: "retry_on_timeout",
    triggerCondition: "timeout",
    healingAction: "retry",
    successRate: 70,
  },
  {
    id: "rollback_on_data_corruption",
    triggerCondition: "data_corruption",
    healingAction: "rollback",
    successRate: 90,
  },
  {
    id: "skip_optional_step",
    triggerCondition: "non_critical_failure",
    healingAction: "skip",
    successRate: 85,
  },
  {
    id: "use_alternative_tool",
    triggerCondition: "tool_unavailable",
    healingAction: "alternative",
    alternative: "使用备选工具",
    successRate: 75,
  },
  {
    id: "ask_user_on_ambiguity",
    triggerCondition: "ambiguous_input",
    healingAction: "ask_user",
    successRate: 95,
  },
];

// ========== 交互配置常量 ==========

/**
 * v2.9.58: 默认交互配置
 */
export const DEFAULT_INTERACTION_CONFIG: InteractionConfig = {
  clarificationThreshold: 0.7,
  confirmDestructiveOps: true,
  offerAlternatives: true,
  allowFreeformResponse: true,
  largeOperationThreshold: 100,
  enableStepReflection: true,
  proactiveSuggestions: true,
};

// ========== 存储键常量 ==========

/**
 * 记忆存储键
 */
export const MEMORY_STORAGE_KEY = "agent_memory_v2";

/**
 * 用户档案存储键
 */
export const USER_PROFILE_STORAGE_KEY = "agent_user_profile_v1";

/**
 * 工作簿缓存存储键
 */
export const WORKBOOK_CACHE_STORAGE_KEY = "agent_workbook_cache_v1";

// ========== 默认值常量 ==========

/**
 * 默认最大迭代次数
 */
export const DEFAULT_MAX_ITERATIONS = 30;

/**
 * 默认超时时间 (ms)
 */
export const DEFAULT_TIMEOUT = 60000;

/**
 * 默认工作簿缓存有效期 (ms)
 */
export const DEFAULT_WORKBOOK_CACHE_TTL = 5 * 60 * 1000; // 5分钟
