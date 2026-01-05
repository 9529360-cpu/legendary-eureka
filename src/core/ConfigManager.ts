/**
 * ConfigManager - 统一配置管理器
 * v1.0.0
 *
 * 功能：
 * 1. 集中管理所有配置
 * 2. 环境感知配置加载
 * 3. 运行时配置更新
 * 4. 配置验证与默认值
 * 5. 配置变更通知
 *
 * 解决的问题：
 * - 配置管理分散
 * - 部分常量硬编码在各自模块
 * - 环境切换困难
 */

import { Logger } from "../utils/Logger";

// ========== 类型定义 ==========

/**
 * 环境类型
 */
export type Environment = "development" | "staging" | "production" | "test";

/**
 * API 配置
 */
export interface ApiConfig {
  /** 后端 URL */
  backendUrl: string;
  /** API 超时（毫秒） */
  timeout: number;
  /** 最大重试次数 */
  maxRetries: number;
  /** 重试延迟（毫秒） */
  retryDelay: number;
  /** 健康检查间隔（毫秒） */
  healthCheckInterval: number;
  /** 是否启用流式响应 */
  enableStreaming: boolean;
}

/**
 * Excel 配置
 */
export interface ExcelConfig {
  /** 单次操作最大行数 */
  maxRowsPerOperation: number;
  /** 单次操作最大列数 */
  maxColumnsPerOperation: number;
  /** 批量操作最大单元格数 */
  maxCellsPerBatch: number;
  /** Excel 操作超时（毫秒） */
  operationTimeout: number;
  /** 是否启用撤销支持 */
  enableUndo: boolean;
  /** 最大撤销历史 */
  maxUndoHistory: number;
  /** 默认表格样式 */
  defaultTableStyle: string;
  /** 默认日期格式 */
  defaultDateFormat: string;
  /** 默认数字格式 */
  defaultNumberFormat: string;
}

/**
 * Agent 配置
 */
export interface AgentConfig {
  /** 最大计划步骤数 */
  maxPlanSteps: number;
  /** ReAct 循环最大次数 */
  maxReactIterations: number;
  /** 对话历史保留条数 */
  maxConversationHistory: number;
  /** 最近操作记录数 */
  maxRecentOperations: number;
  /** 学习偏好最大数量 */
  maxLearnedPreferences: number;
  /** 工作簿上下文缓存时间（毫秒） */
  workbookCacheTtl: number;
  /** 上下文窗口最大 token 数 */
  maxContextTokens: number;
  /** 是否启用记忆 */
  enableMemory: boolean;
  /** 是否启用详细日志 */
  verboseLogging: boolean;
  /** 是否需要确认危险操作 */
  requireConfirmation: boolean;
}

/**
 * UI 配置
 */
export interface UiConfig {
  /** Toast 显示时长（毫秒） */
  toastDuration: number;
  /** 消息最大显示长度 */
  messageMaxLength: number;
  /** 数据预览最大行数 */
  previewMaxRows: number;
  /** 动画持续时间（毫秒） */
  animationDuration: number;
  /** 进度条更新间隔（毫秒） */
  progressUpdateInterval: number;
  /** 主题 */
  theme: "light" | "dark" | "auto";
  /** 语言 */
  locale: string;
}

/**
 * 日志配置
 */
export interface LoggingConfig {
  /** 日志级别 */
  level: "debug" | "info" | "warn" | "error" | "none";
  /** 日志历史最大条数 */
  maxHistory: number;
  /** 错误历史最大条数 */
  maxErrorHistory: number;
  /** 数据截断长度 */
  dataTruncateLength: number;
  /** 是否启用控制台输出 */
  enableConsole: boolean;
  /** 是否启用追踪 */
  enableTracing: boolean;
  /** 追踪采样率 */
  tracingSampleRate: number;
}

/**
 * 安全配置
 */
export interface SecurityConfig {
  /** 敏感字段列表 */
  sensitiveFields: string[];
  /** 是否启用权限检查 */
  enablePermissionCheck: boolean;
  /** 操作确认阈值（高于此风险级别需确认） */
  confirmationThreshold: "low" | "medium" | "high";
  /** 是否记录操作审计 */
  enableAuditLog: boolean;
  /** API 密钥存储位置 */
  apiKeyStorage: "memory" | "localStorage" | "sessionStorage";
}

/**
 * 功能开关
 */
export interface FeatureFlags {
  /** AI 集成 */
  aiIntegration: boolean;
  /** Excel 操作 */
  excelOperations: boolean;
  /** 实时更新 */
  realTimeUpdates: boolean;
  /** 错误日志 */
  errorLogging: boolean;
  /** 调试模式 */
  debugMode: boolean;
  /** 性能监控 */
  performanceMonitoring: boolean;
  /** 工具追踪 */
  toolTracing: boolean;
  /** 智能建议 */
  smartSuggestions: boolean;
  /** 多轮对话 */
  multiTurnConversation: boolean;
}

/**
 * 完整配置
 */
export interface AppConfig {
  /** 环境 */
  environment: Environment;
  /** 版本 */
  version: string;
  /** API 配置 */
  api: ApiConfig;
  /** Excel 配置 */
  excel: ExcelConfig;
  /** Agent 配置 */
  agent: AgentConfig;
  /** UI 配置 */
  ui: UiConfig;
  /** 日志配置 */
  logging: LoggingConfig;
  /** 安全配置 */
  security: SecurityConfig;
  /** 功能开关 */
  features: FeatureFlags;
}

/**
 * 配置变更事件
 */
export interface ConfigChangeEvent {
  path: string;
  oldValue: unknown;
  newValue: unknown;
  timestamp: Date;
}

// ========== 默认配置 ==========

const DEFAULT_CONFIG: AppConfig = {
  environment: "development",
  version: "1.0.0",

  api: {
    backendUrl: "http://localhost:3001",
    timeout: 30000,
    maxRetries: 3,
    retryDelay: 1000,
    healthCheckInterval: 30000,
    enableStreaming: true,
  },

  excel: {
    maxRowsPerOperation: 10000,
    maxColumnsPerOperation: 100,
    maxCellsPerBatch: 1000,
    operationTimeout: 60000,
    enableUndo: true,
    maxUndoHistory: 50,
    defaultTableStyle: "TableStyleMedium2",
    defaultDateFormat: "YYYY-MM-DD",
    defaultNumberFormat: "#,##0.00",
  },

  agent: {
    maxPlanSteps: 20,
    maxReactIterations: 30,
    maxConversationHistory: 50,
    maxRecentOperations: 10,
    maxLearnedPreferences: 100,
    workbookCacheTtl: 30000,
    maxContextTokens: 8000,
    enableMemory: true,
    verboseLogging: false,
    requireConfirmation: true,
  },

  ui: {
    toastDuration: 3000,
    messageMaxLength: 500,
    previewMaxRows: 20,
    animationDuration: 200,
    progressUpdateInterval: 100,
    theme: "auto",
    locale: "zh-CN",
  },

  logging: {
    level: "info",
    maxHistory: 100,
    maxErrorHistory: 50,
    dataTruncateLength: 500,
    enableConsole: true,
    enableTracing: true,
    tracingSampleRate: 1.0,
  },

  security: {
    sensitiveFields: [
      "apiKey",
      "api_key",
      "password",
      "token",
      "secret",
      "authorization",
      "credential",
    ],
    enablePermissionCheck: false,
    confirmationThreshold: "high",
    enableAuditLog: false,
    apiKeyStorage: "memory",
  },

  features: {
    aiIntegration: true,
    excelOperations: true,
    realTimeUpdates: true,
    errorLogging: true,
    debugMode: false,
    performanceMonitoring: true,
    toolTracing: true,
    smartSuggestions: true,
    multiTurnConversation: true,
  },
};

// ========== 环境特定配置 ==========

const ENVIRONMENT_OVERRIDES: Record<Environment, Partial<DeepPartial<AppConfig>>> = {
  development: {
    features: {
      debugMode: true,
    },
    logging: {
      level: "debug",
      enableTracing: true,
    },
    agent: {
      verboseLogging: true,
    },
  },

  staging: {
    features: {
      debugMode: false,
    },
    logging: {
      level: "info",
    },
    api: {
      backendUrl: "https://staging-api.example.com",
    },
  },

  production: {
    features: {
      debugMode: false,
    },
    logging: {
      level: "warn",
      tracingSampleRate: 0.1,
    },
    security: {
      enablePermissionCheck: true,
      enableAuditLog: true,
    },
    api: {
      backendUrl: "https://api.example.com",
    },
  },

  test: {
    features: {
      debugMode: true,
    },
    logging: {
      level: "debug",
      enableConsole: false,
    },
    api: {
      timeout: 5000,
      maxRetries: 1,
    },
  },
};

// ========== 类型辅助 ==========

type DeepPartial<T> = {
  [P in keyof T]?: T[P] extends object ? DeepPartial<T[P]> : T[P];
};

// ========== 配置管理器实现 ==========

/**
 * 配置管理器类
 */
class ConfigManagerClass {
  private config: AppConfig;
  private changeListeners: ((event: ConfigChangeEvent) => void)[] = [];
  private initialized: boolean = false;

  constructor() {
    this.config = this.deepClone(DEFAULT_CONFIG);
  }

  /**
   * 初始化配置管理器
   */
  initialize(environment?: Environment): void {
    if (this.initialized) {
      Logger.warn("ConfigManager", "配置管理器已初始化");
      return;
    }

    // 检测环境
    const env = environment || this.detectEnvironment();
    this.config.environment = env;

    // 应用环境特定配置
    const overrides = ENVIRONMENT_OVERRIDES[env];
    if (overrides) {
      this.config = this.deepMerge(this.config, overrides as Partial<AppConfig>);
    }

    // 从 localStorage 加载用户配置（如果有）
    this.loadUserConfig();

    this.initialized = true;
    Logger.info("ConfigManager", `配置管理器已初始化 (环境: ${env})`);
  }

  /**
   * 获取完整配置
   */
  getConfig(): Readonly<AppConfig> {
    return this.config;
  }

  /**
   * 获取指定路径的配置值
   */
  get<T = unknown>(path: string): T | undefined {
    const keys = path.split(".");
    let value: unknown = this.config;

    for (const key of keys) {
      if (value && typeof value === "object" && key in value) {
        value = (value as Record<string, unknown>)[key];
      } else {
        return undefined;
      }
    }

    return value as T;
  }

  /**
   * 设置配置值
   */
  set(path: string, value: unknown): void {
    const keys = path.split(".");
    const lastKey = keys.pop()!;

    let target: Record<string, unknown> = this.config as unknown as Record<string, unknown>;

    for (const key of keys) {
      if (!(key in target)) {
        target[key] = {};
      }
      target = target[key] as Record<string, unknown>;
    }

    const oldValue = target[lastKey];
    target[lastKey] = value;

    // 触发变更事件
    this.notifyChange({
      path,
      oldValue,
      newValue: value,
      timestamp: new Date(),
    });

    // 持久化到 localStorage
    this.saveUserConfig();

    Logger.debug("ConfigManager", `配置已更新: ${path}`, { oldValue, newValue: value });
  }

  /**
   * 批量更新配置
   */
  update(updates: DeepPartial<AppConfig>): void {
    this.config = this.deepMerge(this.config, updates as Partial<AppConfig>);
    this.saveUserConfig();
    Logger.info("ConfigManager", "配置已批量更新");
  }

  /**
   * 重置为默认配置
   */
  reset(): void {
    const env = this.config.environment;
    this.config = this.deepClone(DEFAULT_CONFIG);
    this.config.environment = env;

    // 应用环境特定配置
    const overrides = ENVIRONMENT_OVERRIDES[env];
    if (overrides) {
      this.config = this.deepMerge(this.config, overrides as Partial<AppConfig>);
    }

    // 清除用户配置
    try {
      localStorage.removeItem("excel-copilot-config");
    } catch {
      // 忽略 localStorage 错误
    }

    Logger.info("ConfigManager", "配置已重置为默认值");
  }

  /**
   * 添加配置变更监听器
   */
  addChangeListener(listener: (event: ConfigChangeEvent) => void): () => void {
    this.changeListeners.push(listener);
    return () => {
      this.changeListeners = this.changeListeners.filter((l) => l !== listener);
    };
  }

  /**
   * 验证配置
   */
  validate(): { valid: boolean; errors: string[] } {
    const errors: string[] = [];

    // API 配置验证
    if (this.config.api.timeout <= 0) {
      errors.push("api.timeout 必须大于 0");
    }
    if (this.config.api.maxRetries < 0) {
      errors.push("api.maxRetries 不能为负数");
    }

    // Excel 配置验证
    if (this.config.excel.maxRowsPerOperation <= 0) {
      errors.push("excel.maxRowsPerOperation 必须大于 0");
    }
    if (this.config.excel.maxColumnsPerOperation <= 0) {
      errors.push("excel.maxColumnsPerOperation 必须大于 0");
    }

    // Agent 配置验证
    if (this.config.agent.maxPlanSteps <= 0) {
      errors.push("agent.maxPlanSteps 必须大于 0");
    }
    if (this.config.agent.maxReactIterations <= 0) {
      errors.push("agent.maxReactIterations 必须大于 0");
    }

    // 日志配置验证
    if (this.config.logging.tracingSampleRate < 0 || this.config.logging.tracingSampleRate > 1) {
      errors.push("logging.tracingSampleRate 必须在 0-1 之间");
    }

    return {
      valid: errors.length === 0,
      errors,
    };
  }

  /**
   * 导出配置
   */
  export(): string {
    return JSON.stringify(this.config, null, 2);
  }

  /**
   * 导入配置
   */
  import(configJson: string): boolean {
    try {
      const imported = JSON.parse(configJson) as Partial<AppConfig>;
      this.config = this.deepMerge(this.config, imported);
      const validation = this.validate();

      if (!validation.valid) {
        Logger.error("ConfigManager", "导入的配置无效", { errors: validation.errors });
        return false;
      }

      this.saveUserConfig();
      Logger.info("ConfigManager", "配置已导入");
      return true;
    } catch (error) {
      Logger.error("ConfigManager", "配置导入失败", { error });
      return false;
    }
  }

  // ========== 便捷访问方法 ==========

  /** 获取 API 配置 */
  get api(): Readonly<ApiConfig> {
    return this.config.api;
  }

  /** 获取 Excel 配置 */
  get excel(): Readonly<ExcelConfig> {
    return this.config.excel;
  }

  /** 获取 Agent 配置 */
  get agent(): Readonly<AgentConfig> {
    return this.config.agent;
  }

  /** 获取 UI 配置 */
  get ui(): Readonly<UiConfig> {
    return this.config.ui;
  }

  /** 获取日志配置 */
  get logging(): Readonly<LoggingConfig> {
    return this.config.logging;
  }

  /** 获取安全配置 */
  get security(): Readonly<SecurityConfig> {
    return this.config.security;
  }

  /** 获取功能开关 */
  get features(): Readonly<FeatureFlags> {
    return this.config.features;
  }

  /** 获取当前环境 */
  get environment(): Environment {
    return this.config.environment;
  }

  /** 是否为开发环境 */
  get isDevelopment(): boolean {
    return this.config.environment === "development";
  }

  /** 是否为生产环境 */
  get isProduction(): boolean {
    return this.config.environment === "production";
  }

  /** 是否启用调试模式 */
  get isDebug(): boolean {
    return this.config.features.debugMode;
  }

  // ========== 私有方法 ==========

  /**
   * 检测环境
   */
  private detectEnvironment(): Environment {
    // 检查 process.env
    if (typeof process !== "undefined" && process.env) {
      const nodeEnv = process.env.NODE_ENV;
      if (nodeEnv === "production") return "production";
      if (nodeEnv === "test") return "test";
      if (nodeEnv === "staging") return "staging";
    }

    // 检查 URL
    if (typeof window !== "undefined" && window.location) {
      const hostname = window.location.hostname;
      if (hostname === "localhost" || hostname === "127.0.0.1") {
        return "development";
      }
      if (hostname.includes("staging")) {
        return "staging";
      }
    }

    return "development";
  }

  /**
   * 从 localStorage 加载用户配置
   */
  private loadUserConfig(): void {
    try {
      const saved = localStorage.getItem("excel-copilot-config");
      if (saved) {
        const userConfig = JSON.parse(saved) as Partial<AppConfig>;
        this.config = this.deepMerge(this.config, userConfig);
        Logger.debug("ConfigManager", "已加载用户配置");
      }
    } catch {
      // 忽略 localStorage 错误
    }
  }

  /**
   * 保存用户配置到 localStorage
   */
  private saveUserConfig(): void {
    try {
      // 只保存与默认值不同的配置
      const diff = this.getConfigDiff();
      if (Object.keys(diff).length > 0) {
        localStorage.setItem("excel-copilot-config", JSON.stringify(diff));
      }
    } catch {
      // 忽略 localStorage 错误
    }
  }

  /**
   * 获取与默认配置的差异
   */
  private getConfigDiff(): Partial<AppConfig> {
    // 简化实现：保存整个配置
    return this.config;
  }

  /**
   * 通知变更
   */
  private notifyChange(event: ConfigChangeEvent): void {
    this.changeListeners.forEach((listener) => {
      try {
        listener(event);
      } catch (error) {
        Logger.error("ConfigManager", "配置变更监听器执行失败", { error });
      }
    });
  }

  /**
   * 深度克隆
   */
  private deepClone<T>(obj: T): T {
    return JSON.parse(JSON.stringify(obj));
  }

  /**
   * 深度合并
   */
  private deepMerge<T extends object>(target: T, source: Partial<T>): T {
    const result = { ...target };

    for (const key in source) {
      if (Object.prototype.hasOwnProperty.call(source, key)) {
        const targetValue = result[key];
        const sourceValue = source[key];

        if (
          sourceValue !== undefined &&
          typeof sourceValue === "object" &&
          sourceValue !== null &&
          !Array.isArray(sourceValue) &&
          typeof targetValue === "object" &&
          targetValue !== null &&
          !Array.isArray(targetValue)
        ) {
          (result as Record<string, unknown>)[key] = this.deepMerge(
            targetValue as object,
            sourceValue as object
          );
        } else if (sourceValue !== undefined) {
          (result as Record<string, unknown>)[key] = sourceValue;
        }
      }
    }

    return result;
  }
}

// 导出单例
export const ConfigManager = new ConfigManagerClass();

// 自动初始化
ConfigManager.initialize();

// 便捷方法导出
export const config = {
  get: <T = unknown>(path: string) => ConfigManager.get<T>(path),
  set: (path: string, value: unknown) => ConfigManager.set(path, value),
  getConfig: () => ConfigManager.getConfig(),
  api: () => ConfigManager.api,
  excel: () => ConfigManager.excel,
  agent: () => ConfigManager.agent,
  ui: () => ConfigManager.ui,
  logging: () => ConfigManager.logging,
  security: () => ConfigManager.security,
  features: () => ConfigManager.features,
  isDebug: () => ConfigManager.isDebug,
  isDev: () => ConfigManager.isDevelopment,
  isProd: () => ConfigManager.isProduction,
};

export default ConfigManager;
