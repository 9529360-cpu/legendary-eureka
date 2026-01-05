/**
 * Logger - 统一日志工具类
 * v1.0.0
 *
 * 功能：
 * - 日志级别控制 (DEBUG, INFO, WARN, ERROR)
 * - 生产环境自动关闭 DEBUG 日志
 * - 敏感信息脱敏
 * - 结构化日志输出
 * - 性能计时器
 */

import { CONFIG } from "../config";

/** 日志级别枚举 */
export enum LogLevel {
  DEBUG = 0,
  INFO = 1,
  WARN = 2,
  ERROR = 3,
  NONE = 4, // 禁用所有日志
}

/** 日志条目接口 */
export interface LogEntry {
  level: LogLevel;
  module: string;
  message: string;
  timestamp: Date;
  data?: unknown;
  duration?: number;
}

/** 日志配置 */
export interface LoggerConfig {
  level: LogLevel;
  enableConsole: boolean;
  enableTimestamp: boolean;
  sensitiveFields: string[];
  maxDataLength: number;
}

/** 敏感字段默认列表 */
const DEFAULT_SENSITIVE_FIELDS = [
  "apiKey",
  "api_key",
  "password",
  "token",
  "secret",
  "authorization",
  "credential",
  "key",
];

/**
 * Logger 单例类
 */
class LoggerClass {
  private config: LoggerConfig;
  private timers: Map<string, number> = new Map();
  private logHistory: LogEntry[] = [];
  private maxHistorySize = 100;

  constructor() {
    // 根据环境自动设置日志级别
    const isProduction = !CONFIG.debug;

    this.config = {
      level: isProduction ? LogLevel.WARN : LogLevel.DEBUG,
      enableConsole: true,
      enableTimestamp: true,
      sensitiveFields: DEFAULT_SENSITIVE_FIELDS,
      maxDataLength: 500, // 生产环境限制数据长度
    };
  }

  /**
   * 配置日志器
   */
  configure(config: Partial<LoggerConfig>): void {
    this.config = { ...this.config, ...config };
  }

  /**
   * 设置日志级别
   */
  setLevel(level: LogLevel): void {
    this.config.level = level;
  }

  /**
   * 获取当前日志级别
   */
  getLevel(): LogLevel {
    return this.config.level;
  }

  /**
   * DEBUG 级别日志
   */
  debug(module: string, message: string, data?: unknown): void {
    this.log(LogLevel.DEBUG, module, message, data);
  }

  /**
   * INFO 级别日志
   */
  info(module: string, message: string, data?: unknown): void {
    this.log(LogLevel.INFO, module, message, data);
  }

  /**
   * WARN 级别日志
   */
  warn(module: string, message: string, data?: unknown): void {
    this.log(LogLevel.WARN, module, message, data);
  }

  /**
   * ERROR 级别日志
   */
  error(module: string, message: string, data?: unknown): void {
    this.log(LogLevel.ERROR, module, message, data);
  }

  /**
   * 开始计时
   */
  time(label: string): void {
    this.timers.set(label, Date.now());
  }

  /**
   * 结束计时并输出
   */
  timeEnd(label: string, module?: string): number {
    const start = this.timers.get(label);
    if (!start) {
      this.warn("Logger", `Timer '${label}' does not exist`);
      return 0;
    }

    const duration = Date.now() - start;
    this.timers.delete(label);

    if (module) {
      this.debug(module, `⏱️ ${label}: ${duration}ms`);
    }

    return duration;
  }

  /**
   * 分组日志开始
   */
  group(label: string): void {
    if (this.config.level <= LogLevel.DEBUG && this.config.enableConsole) {
      console.group(label);
    }
  }

  /**
   * 分组日志结束
   */
  groupEnd(): void {
    if (this.config.level <= LogLevel.DEBUG && this.config.enableConsole) {
      console.groupEnd();
    }
  }

  /**
   * 获取日志历史（用于诊断）
   */
  getHistory(): LogEntry[] {
    return [...this.logHistory];
  }

  /**
   * 清除日志历史
   */
  clearHistory(): void {
    this.logHistory = [];
  }

  /**
   * 核心日志方法
   */
  private log(level: LogLevel, module: string, message: string, data?: unknown): void {
    // 检查日志级别
    if (level < this.config.level) {
      return;
    }

    // 创建日志条目
    const entry: LogEntry = {
      level,
      module,
      message,
      timestamp: new Date(),
      data: data ? this.sanitizeData(data) : undefined,
    };

    // 保存到历史
    this.addToHistory(entry);

    // 输出到控制台
    if (this.config.enableConsole) {
      this.consoleOutput(entry);
    }
  }

  /**
   * 敏感信息脱敏
   */
  private sanitizeData(data: unknown): unknown {
    if (data === null || data === undefined) {
      return data;
    }

    // 字符串处理
    if (typeof data === "string") {
      return this.truncateString(data);
    }

    // 数组处理
    if (Array.isArray(data)) {
      return data.slice(0, 10).map((item) => this.sanitizeData(item));
    }

    // 对象处理
    if (typeof data === "object") {
      const sanitized: Record<string, unknown> = {};
      const obj = data as Record<string, unknown>;

      for (const key of Object.keys(obj)) {
        if (this.isSensitiveField(key)) {
          sanitized[key] = "***REDACTED***";
        } else {
          sanitized[key] = this.sanitizeData(obj[key]);
        }
      }

      return sanitized;
    }

    return data;
  }

  /**
   * 检查是否为敏感字段
   */
  private isSensitiveField(fieldName: string): boolean {
    const lowerName = fieldName.toLowerCase();
    return this.config.sensitiveFields.some((sensitive) =>
      lowerName.includes(sensitive.toLowerCase())
    );
  }

  /**
   * 截断过长字符串
   */
  private truncateString(str: string): string {
    if (str.length <= this.config.maxDataLength) {
      return str;
    }
    return str.substring(0, this.config.maxDataLength) + "...[TRUNCATED]";
  }

  /**
   * 添加到历史记录
   */
  private addToHistory(entry: LogEntry): void {
    this.logHistory.push(entry);
    if (this.logHistory.length > this.maxHistorySize) {
      this.logHistory.shift();
    }
  }

  /**
   * 控制台输出
   */
  private consoleOutput(entry: LogEntry): void {
    const levelIcons: Record<LogLevel, string> = {
      [LogLevel.DEBUG]: "🔍",
      [LogLevel.INFO]: "ℹ️",
      [LogLevel.WARN]: "⚠️",
      [LogLevel.ERROR]: "❌",
      [LogLevel.NONE]: "",
    };

    const levelColors: Record<LogLevel, string> = {
      [LogLevel.DEBUG]: "color: #888",
      [LogLevel.INFO]: "color: #0066cc",
      [LogLevel.WARN]: "color: #cc6600",
      [LogLevel.ERROR]: "color: #cc0000",
      [LogLevel.NONE]: "",
    };

    const icon = levelIcons[entry.level];
    const timestamp = this.config.enableTimestamp
      ? `[${entry.timestamp.toLocaleTimeString()}]`
      : "";
    const prefix = `${icon} ${timestamp}[${entry.module}]`;

    const logMethod = this.getConsoleMethod(entry.level);

    if (entry.data !== undefined) {
      logMethod(`%c${prefix} ${entry.message}`, levelColors[entry.level], entry.data);
    } else {
      logMethod(`%c${prefix} ${entry.message}`, levelColors[entry.level]);
    }
  }

  /**
   * 获取对应的 console 方法
   */
  private getConsoleMethod(level: LogLevel): (...args: unknown[]) => void {
    switch (level) {
      case LogLevel.DEBUG:
        return console.debug.bind(console);
      case LogLevel.INFO:
        return console.info.bind(console);
      case LogLevel.WARN:
        return console.warn.bind(console);
      case LogLevel.ERROR:
        return console.error.bind(console);
      default:
        return console.log.bind(console);
    }
  }
}

/** 导出单例实例 */
export const Logger = new LoggerClass();

/** 快捷方法导出 */
export const log = {
  debug: (module: string, message: string, data?: unknown) => Logger.debug(module, message, data),
  info: (module: string, message: string, data?: unknown) => Logger.info(module, message, data),
  warn: (module: string, message: string, data?: unknown) => Logger.warn(module, message, data),
  error: (module: string, message: string, data?: unknown) => Logger.error(module, message, data),
  time: (label: string) => Logger.time(label),
  timeEnd: (label: string, module?: string) => Logger.timeEnd(label, module),
  group: (label: string) => Logger.group(label),
  groupEnd: () => Logger.groupEnd(),
};

export default Logger;
