/**
 * 记忆系统相关类型定义
 *
 * 从 AgentCore.ts 抽取，用于定义 Agent 记忆和学习接口
 */

/**
 * v2.9.19: 用户偏好设置
 */
export interface UserPreferences {
  /** 默认表格样式 (e.g., "TableStyleMedium2") */
  tableStyle: string;
  /** 日期格式偏好 (e.g., "YYYY-MM-DD", "MM/DD/YYYY") */
  dateFormat: string;
  /** 货币符号 (e.g., "", "$", "") */
  currencySymbol: string;
  /** 是否总是使用公式（而非硬编码值） */
  alwaysUseFormulas: boolean;
  /** 删除操作前是否确认 */
  confirmBeforeDelete: boolean;
  /** 回复详细程度 */
  verbosityLevel: "brief" | "normal" | "detailed";
  /** 默认小数位数 */
  decimalPlaces: number;
  /** 是否显示执行计划 */
  showExecutionPlan: boolean;
  /** 首选图表类型 */
  preferredChartType: string;
  /** 默认字体 */
  defaultFont: string;
  /** 默认字号 */
  defaultFontSize: number;
}

/**
 * v2.9.19: 用户档案
 */
export interface UserProfile {
  id: string;
  /** 用户偏好设置 */
  preferences: UserPreferences;
  /** 最近创建的表 */
  recentTables: string[];
  /** 常用列名 */
  commonColumns: string[];
  /** 常用公式模式 */
  commonFormulas: string[];
  /** 上次活动时间 */
  lastSeen: Date;
  /** 创建时间 */
  createdAt: Date;
  /** 用户统计 */
  stats: {
    totalTasks: number;
    successfulTasks: number;
    failedTasks: number;
    tablesCreated: number;
    chartsCreated: number;
    formulasWritten: number;
  };
}

/**
 * v2.9.19: 已完成任务（增强版）
 */
export interface CompletedTask {
  id: string;
  /** 用户原始请求 */
  request: string;
  /** 任务结果摘要 */
  result: string;
  /** 涉及的表格 */
  tables: string[];
  /** 使用的公式 */
  formulas: string[];
  /** 创建的列名 */
  columns: string[];
  /** 任务时间戳 */
  timestamp: Date;
  /** 是否成功 */
  success: boolean;
  /** 执行步骤数 */
  stepCount: number;
  /** 执行时长 (ms) */
  duration: number;
  /** 用户反馈 */
  userFeedback?: "positive" | "negative" | "neutral";
  /** 任务标签（用于分类） */
  tags: string[];
  /** 质量分数 */
  qualityScore?: number;
}

/**
 * v2.9.19: 任务模式（从历史中学习）
 */
export interface TaskPattern {
  /** 模式关键词 */
  keywords: string[];
  /** 匹配的任务类型 */
  taskType: string;
  /** 出现次数 */
  frequency: number;
  /** 平均成功率 */
  successRate: number;
  /** 典型执行步骤 */
  typicalSteps: string[];
  /** 推荐的处理方式 */
  recommendation?: string;
  /** 使用次数 */
  usageCount?: number;
}

/**
 * v2.9.19: 工作簿上下文缓存
 */
export interface CachedWorkbookContext {
  /** 工作簿名称 */
  workbookName: string;
  /** 工作表列表 */
  sheets: CachedSheetInfo[];
  /** 命名范围 */
  namedRanges: string[];
  /** 表格列表 */
  tables: string[];
  /** 缓存时间 */
  cachedAt: Date;
  /** 缓存有效期 (ms) */
  ttl: number;
  /** 是否过期 */
  isExpired: boolean;
}

/**
 * v2.9.19: 缓存的工作表信息
 */
export interface CachedSheetInfo {
  name: string;
  /** 已使用范围 */
  usedRange: string;
  /** 行数 */
  rowCount: number;
  /** 列数 */
  columnCount: number;
  /** 列标题 */
  headers: string[];
  /** 数据类型摘要 */
  dataTypes: Record<string, string>;
  /** 是否包含表格 */
  hasTables: boolean;
  /** 是否包含图表 */
  hasCharts: boolean;
}

/**
 * v2.9.19: 学习到的偏好（从用户行为中推断）
 */
export interface LearnedPreference {
  /** 偏好类型 */
  type: "tableStyle" | "dateFormat" | "columnName" | "formula" | "chartType" | "other";
  /** 偏好值 */
  value: string;
  /** 观察到的次数 */
  observedCount: number;
  /** 首次观察时间 */
  firstSeen: Date;
  /** 最近观察时间 */
  lastSeen: Date;
  /** 置信度 (0-100) */
  confidence: number;
}

/**
 * v2.9.39: 最近操作记录（用于上下文理解）
 */
export interface RecentOperation {
  /** 操作ID */
  id: string;
  /** 操作动作 */
  action: string;
  /** 目标范围 */
  targetRange?: string;
  /** 工作表名 */
  sheetName?: string;
  /** 操作描述 */
  description: string;
  /** 操作是否成功 */
  success: boolean;
  /** 时间戳 */
  timestamp: Date;
  /** 额外元数据 */
  metadata?: {
    /** 是否是表格创建操作 */
    isTableCreation?: boolean;
    /** 列标题 */
    headers?: string[];
    /** 行数 */
    rowCount?: number;
    /** 列数 */
    colCount?: number;
    /** 公式（如果是公式操作）*/
    formula?: string;
    /** 格式化类型 */
    formatType?: string;
  };
}

/**
 * v2.9.22: 语义记忆条目
 */
export interface SemanticMemoryEntry {
  /** 记忆ID */
  id: string;
  /** 内容 */
  content: string;
  /** 向量嵌入（简化为关键词） */
  keywords: string[];
  /** 相关性分数 */
  relevanceScore: number;
  /** 来源 */
  source: "task" | "user" | "system" | "learned";
  /** 创建时间 */
  createdAt: Date;
  /** 最后访问时间 */
  lastAccessedAt: Date;
  /** 访问次数 */
  accessCount: number;
}

/**
 * 用户反馈记录
 */
export interface UserFeedbackRecord {
  /** 反馈ID */
  id: string;
  /** 任务ID */
  taskId: string;
  /** 反馈类型 */
  type: "positive" | "negative" | "neutral" | "correction";
  /** 反馈内容 */
  content?: string;
  /** 时间戳 */
  timestamp: Date;
  /** 是否已处理 */
  processed: boolean;
  /** 处理结果 */
  outcome?: string;
}

/**
 * 学习到的模式
 */
export interface LearnedPattern {
  /** 模式ID */
  id: string;
  /** 模式描述 */
  description: string;
  /** 触发条件 */
  triggers: string[];
  /** 推荐动作 */
  recommendedActions: string[];
  /** 置信度 */
  confidence: number;
  /** 来源任务 */
  sourceTasks: string[];
  /** 创建时间 */
  createdAt: Date;
  /** 使用次数 */
  usageCount: number;
}
