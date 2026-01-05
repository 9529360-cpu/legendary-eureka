/**
 * Constants - 配置常量集中管理
 * v1.0.0
 *
 * 消除代码中的魔法数字和硬编码配置
 */

// ========== 超时配置 (毫秒) ==========
export const TIMEOUTS = {
  /** API 默认超时 */
  API_DEFAULT: 30000, // 30s
  /** Agent 请求超时 */
  AGENT_REQUEST: 60000, // 60s
  /** Function Calling 超时 */
  FUNCTION_CALLING: 90000, // 90s
  /** 流式聊天超时 */
  STREAM_CHAT: 120000, // 120s
  /** Excel 操作超时 */
  EXCEL_OPERATION: 60000, // 60s
  /** 健康检查超时 */
  HEALTH_CHECK: 5000, // 5s
  /** 重试延迟基础值 */
  RETRY_BASE_DELAY: 1000, // 1s
  /** 防抖延迟 */
  DEBOUNCE_DELAY: 300, // 300ms
} as const;

// ========== 重试配置 ==========
export const RETRY = {
  /** 最大重试次数 */
  MAX_ATTEMPTS: 3,
  /** 智能重试最大次数 */
  SMART_RETRY_MAX: 2,
  /** Replan 最大次数 */
  REPLAN_MAX: 3,
  /** 指数退避乘数 */
  BACKOFF_MULTIPLIER: 2,
  /** 最大重试延迟 */
  MAX_DELAY: 10000, // 10s
} as const;

// ========== Excel 限制 ==========
export const EXCEL_LIMITS = {
  /** 单次操作最大行数 */
  MAX_ROWS_PER_OPERATION: 10000,
  /** 单次操作最大列数 */
  MAX_COLUMNS_PER_OPERATION: 100,
  /** 批量操作最大单元格数 */
  MAX_CELLS_PER_BATCH: 1000,
  /** 表格最大预留行 */
  TABLE_RESERVE_ROWS: 1000,
  /** 最大读取行数 */
  MAX_READ_ROWS: 5000,
  /** 采样校验行数 */
  SAMPLE_VERIFICATION_ROWS: 10,
} as const;

// ========== Agent 配置 ==========
export const AGENT = {
  /** 最大计划步骤数 */
  MAX_PLAN_STEPS: 20,
  /** ReAct 循环最大次数 */
  MAX_REACT_ITERATIONS: 15,
  /** 对话历史保留条数 */
  MAX_CONVERSATION_HISTORY: 50,
  /** 最近操作记录数 */
  MAX_RECENT_OPERATIONS: 10,
  /** 学习偏好最大数量 */
  MAX_LEARNED_PREFERENCES: 100,
  /** 工作簿上下文缓存时间 (毫秒) */
  WORKBOOK_CACHE_TTL: 30000, // 30s
  /** 上下文窗口最大 token 数 */
  MAX_CONTEXT_TOKENS: 8000,
} as const;

// ========== UI 配置 ==========
export const UI = {
  /** Toast 显示时长 (毫秒) */
  TOAST_DURATION: 3000,
  /** 消息最大显示长度 */
  MESSAGE_MAX_LENGTH: 500,
  /** 数据预览最大行数 */
  PREVIEW_MAX_ROWS: 20,
  /** 动画持续时间 (毫秒) */
  ANIMATION_DURATION: 200,
  /** 进度条更新间隔 (毫秒) */
  PROGRESS_UPDATE_INTERVAL: 100,
} as const;

// ========== 日志配置 ==========
export const LOGGING = {
  /** 日志历史最大条数 */
  MAX_HISTORY: 100,
  /** 错误历史最大条数 */
  MAX_ERROR_HISTORY: 50,
  /** 数据截断长度 */
  DATA_TRUNCATE_LENGTH: 500,
  /** 堆栈截断长度 */
  STACK_TRUNCATE_LENGTH: 1000,
} as const;

// ========== API 端点 ==========
export const API_ENDPOINTS = {
  /** 聊天接口 */
  CHAT: "/chat",
  /** 流式聊天 */
  CHAT_STREAM: "/chat/stream",
  /** 健康检查 */
  HEALTH: "/api/health",
  /** 配置状态 */
  CONFIG_STATUS: "/api/config/status",
  /** 配置密钥 */
  CONFIG_KEY: "/api/config/key",
} as const;

// ========== 工具名称常量 ==========
export const TOOL_NAMES = {
  // Excel 读取操作
  EXCEL_READ_RANGE: "excel_read_range",
  EXCEL_GET_USED_RANGE: "excel_get_used_range",
  EXCEL_GET_SELECTION: "excel_get_selection",

  // Excel 写入操作
  EXCEL_WRITE_RANGE: "excel_write_range",
  EXCEL_WRITE_CELL: "excel_write_cell",
  EXCEL_SET_FORMULA: "excel_set_formula",
  EXCEL_CLEAR_RANGE: "excel_clear_range",

  // Excel 格式化
  EXCEL_FORMAT_RANGE: "excel_format_range",
  EXCEL_SET_NUMBER_FORMAT: "excel_set_number_format",
  EXCEL_CREATE_TABLE: "excel_create_table",
  EXCEL_ADD_CHART: "excel_add_chart",

  // Excel 排序筛选
  EXCEL_SORT: "excel_sort", // v2.9.40: 主工具名
  EXCEL_SORT_RANGE: "excel_sort_range", // v2.9.40: 别名
  EXCEL_FILTER: "excel_filter", // v2.9.40: 主筛选工具
  EXCEL_FILTER_TABLE: "excel_filter_table",
  EXCEL_CLEAR_FILTER: "excel_clear_filter",

  // v2.9.40: 公式工具别名
  EXCEL_SET_FORMULAS: "excel_set_formulas",
  EXCEL_FILL_FORMULA: "excel_fill_formula",

  // 虚拟工具
  RESPOND_TO_USER: "respond_to_user",
} as const;

// ========== 表格样式 ==========
export const TABLE_STYLES = {
  DEFAULT: "TableStyleMedium2",
  LIGHT: "TableStyleLight1",
  MEDIUM: "TableStyleMedium9",
  DARK: "TableStyleDark1",
} as const;

// ========== 图表类型 ==========
export const CHART_TYPES = {
  COLUMN: "ColumnClustered",
  BAR: "BarClustered",
  LINE: "Line",
  PIE: "Pie",
  AREA: "Area",
  SCATTER: "XYScatter",
} as const;

// ========== 风险等级 ==========
export const RISK_LEVELS = {
  LOW: "low",
  MEDIUM: "medium",
  HIGH: "high",
} as const;

export type RiskLevel = (typeof RISK_LEVELS)[keyof typeof RISK_LEVELS];

// ========== 导出统一配置对象 ==========
export const CONSTANTS = {
  TIMEOUTS,
  RETRY,
  EXCEL_LIMITS,
  AGENT,
  UI,
  LOGGING,
  API_ENDPOINTS,
  TOOL_NAMES,
  TABLE_STYLES,
  CHART_TYPES,
  RISK_LEVELS,
} as const;

export default CONSTANTS;
