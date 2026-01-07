/**
 * AgentTracer - Agent 追踪与可观测性 v4.1
 *
 * 提供结构化日志和追踪能力：
 * 1. Span 追踪（类似 OpenTelemetry）
 * 2. 结构化日志
 * 3. 性能指标
 * 4. 事件回放
 *
 * @module agent/tracing/AgentTracer
 */

// ========== 类型定义 ==========

/**
 * 日志级别
 */
export type LogLevel = "debug" | "info" | "warn" | "error";

/**
 * Span 状态
 */
export type SpanStatus = "running" | "success" | "error" | "cancelled";

/**
 * 日志条目
 */
export interface LogEntry {
  /** 时间戳 */
  timestamp: number;

  /** 日志级别 */
  level: LogLevel;

  /** 消息 */
  message: string;

  /** 关联的 Span ID */
  spanId?: string;

  /** 额外数据 */
  data?: unknown;

  /** 来源模块 */
  source?: string;
}

/**
 * Span 事件
 */
export interface SpanEvent {
  /** 事件名称 */
  name: string;

  /** 时间戳 */
  timestamp: number;

  /** 属性 */
  attributes?: Record<string, unknown>;
}

/**
 * Span 定义
 */
export interface Span {
  /** Span ID */
  id: string;

  /** Span 名称 */
  name: string;

  /** 父 Span ID */
  parentId?: string;

  /** 开始时间 */
  startTime: number;

  /** 结束时间 */
  endTime?: number;

  /** 持续时间 (ms) */
  duration?: number;

  /** 状态 */
  status: SpanStatus;

  /** 属性 */
  attributes: Record<string, unknown>;

  /** 事件列表 */
  events: SpanEvent[];

  /** 错误信息 */
  error?: string;
}

/**
 * 追踪数据（可导出）
 */
export interface TraceData {
  /** Trace ID */
  traceId: string;

  /** 所有 Spans */
  spans: Span[];

  /** 所有日志 */
  logs: LogEntry[];

  /** 开始时间 */
  startTime: number;

  /** 结束时间 */
  endTime?: number;

  /** 元数据 */
  metadata: Record<string, unknown>;
}

/**
 * 追踪配置
 */
export interface TracerConfig {
  /** 是否启用 */
  enabled: boolean;

  /** 最大 Span 数 */
  maxSpans: number;

  /** 最大日志数 */
  maxLogs: number;

  /** 日志级别（只记录此级别及以上） */
  logLevel: LogLevel;

  /** 是否输出到控制台 */
  consoleOutput: boolean;
}

/**
 * 默认配置
 */
export const DEFAULT_TRACER_CONFIG: TracerConfig = {
  enabled: true,
  maxSpans: 100,
  maxLogs: 1000,
  logLevel: "info",
  consoleOutput: true,
};

// ========== AgentTracer 类 ==========

/**
 * Agent 追踪器
 */
export class AgentTracer {
  private traceId: string;
  private spans: Span[] = [];
  private logs: LogEntry[] = [];
  private startTime: number;
  private currentSpanStack: Span[] = [];
  private config: TracerConfig;
  private listeners: Array<(entry: LogEntry) => void> = [];

  constructor(config: Partial<TracerConfig> = {}) {
    this.config = { ...DEFAULT_TRACER_CONFIG, ...config };
    this.traceId = this.generateId("trace");
    this.startTime = Date.now();
  }

  /**
   * 生成 ID
   */
  private generateId(prefix: string): string {
    return `${prefix}_${Date.now().toString(36)}_${Math.random().toString(36).substring(2, 8)}`;
  }

  /**
   * 获取当前 Span
   */
  get currentSpan(): Span | undefined {
    return this.currentSpanStack[this.currentSpanStack.length - 1];
  }

  /**
   * 开始一个 Span
   */
  startSpan(name: string, attributes?: Record<string, unknown>): Span {
    if (!this.config.enabled) {
      return this.createDummySpan(name);
    }

    const span: Span = {
      id: this.generateId("span"),
      name,
      parentId: this.currentSpan?.id,
      startTime: Date.now(),
      status: "running",
      attributes: attributes || {},
      events: [],
    };

    this.spans.push(span);
    this.currentSpanStack.push(span);

    // 限制数量
    if (this.spans.length > this.config.maxSpans) {
      this.spans.shift();
    }

    this.log("debug", `[Span:Start] ${name}`, { spanId: span.id, attributes });

    return span;
  }

  /**
   * 结束当前 Span
   */
  endSpan(status: SpanStatus = "success", error?: string): void {
    const span = this.currentSpanStack.pop();
    if (!span) return;

    span.endTime = Date.now();
    span.duration = span.endTime - span.startTime;
    span.status = status;
    if (error) {
      span.error = error;
    }

    this.log("debug", `[Span:End] ${span.name} (${span.duration}ms)`, {
      spanId: span.id,
      status,
      duration: span.duration,
    });
  }

  /**
   * 添加 Span 事件
   */
  addSpanEvent(name: string, attributes?: Record<string, unknown>): void {
    const span = this.currentSpan;
    if (!span) return;

    span.events.push({
      name,
      timestamp: Date.now(),
      attributes,
    });
  }

  /**
   * 设置 Span 属性
   */
  setSpanAttribute(key: string, value: unknown): void {
    const span = this.currentSpan;
    if (!span) return;

    span.attributes[key] = value;
  }

  /**
   * 记录日志
   */
  log(level: LogLevel, message: string, data?: unknown, source?: string): void {
    if (!this.config.enabled) return;
    if (!this.shouldLog(level)) return;

    const entry: LogEntry = {
      timestamp: Date.now(),
      level,
      message,
      spanId: this.currentSpan?.id,
      data,
      source,
    };

    this.logs.push(entry);

    // 限制数量
    if (this.logs.length > this.config.maxLogs) {
      this.logs.shift();
    }

    // 输出到控制台
    if (this.config.consoleOutput) {
      this.outputToConsole(entry);
    }

    // 通知监听器
    this.listeners.forEach((l) => l(entry));
  }

  /**
   * 便捷日志方法
   */
  debug(message: string, data?: unknown, source?: string): void {
    this.log("debug", message, data, source);
  }

  info(message: string, data?: unknown, source?: string): void {
    this.log("info", message, data, source);
  }

  warn(message: string, data?: unknown, source?: string): void {
    this.log("warn", message, data, source);
  }

  error(message: string, data?: unknown, source?: string): void {
    this.log("error", message, data, source);
  }

  /**
   * 检查是否应该记录此级别
   */
  private shouldLog(level: LogLevel): boolean {
    const levels: LogLevel[] = ["debug", "info", "warn", "error"];
    return levels.indexOf(level) >= levels.indexOf(this.config.logLevel);
  }

  /**
   * 输出到控制台
   */
  private outputToConsole(entry: LogEntry): void {
    const prefix = entry.source ? `[${entry.source}]` : "";
    const spanInfo = entry.spanId ? `(span:${entry.spanId.substring(0, 8)})` : "";
    const msg = `${prefix}${spanInfo} ${entry.message}`;

    switch (entry.level) {
      case "debug":
        console.debug(msg, entry.data || "");
        break;
      case "info":
        console.info(msg, entry.data || "");
        break;
      case "warn":
        console.warn(msg, entry.data || "");
        break;
      case "error":
        console.error(msg, entry.data || "");
        break;
    }
  }

  /**
   * 添加日志监听器
   */
  addListener(listener: (entry: LogEntry) => void): () => void {
    this.listeners.push(listener);
    return () => {
      const idx = this.listeners.indexOf(listener);
      if (idx >= 0) {
        this.listeners.splice(idx, 1);
      }
    };
  }

  /**
   * 导出追踪数据
   */
  export(): TraceData {
    return {
      traceId: this.traceId,
      spans: [...this.spans],
      logs: [...this.logs],
      startTime: this.startTime,
      endTime: Date.now(),
      metadata: {
        version: "4.1",
        exportedAt: new Date().toISOString(),
      },
    };
  }

  /**
   * 获取性能摘要
   */
  getPerformanceSummary(): {
    totalSpans: number;
    totalLogs: number;
    averageSpanDuration: number;
    longestSpan: { name: string; duration: number } | null;
    errorCount: number;
    warningCount: number;
  } {
    const completedSpans = this.spans.filter((s) => s.duration !== undefined);
    const durations = completedSpans.map((s) => s.duration!);
    const avgDuration = durations.length > 0 ? durations.reduce((a, b) => a + b, 0) / durations.length : 0;

    let longestSpan: { name: string; duration: number } | null = null;
    for (const span of completedSpans) {
      if (!longestSpan || span.duration! > longestSpan.duration) {
        longestSpan = { name: span.name, duration: span.duration! };
      }
    }

    return {
      totalSpans: this.spans.length,
      totalLogs: this.logs.length,
      averageSpanDuration: Math.round(avgDuration),
      longestSpan,
      errorCount: this.logs.filter((l) => l.level === "error").length,
      warningCount: this.logs.filter((l) => l.level === "warn").length,
    };
  }

  /**
   * 清空数据
   */
  clear(): void {
    this.spans = [];
    this.logs = [];
    this.currentSpanStack = [];
    this.traceId = this.generateId("trace");
    this.startTime = Date.now();
  }

  /**
   * 创建虚拟 Span（禁用时）
   */
  private createDummySpan(name: string): Span {
    return {
      id: "dummy",
      name,
      startTime: Date.now(),
      status: "success",
      attributes: {},
      events: [],
    };
  }

  /**
   * 执行带追踪的函数
   */
  async trace<T>(
    name: string,
    fn: () => Promise<T>,
    attributes?: Record<string, unknown>
  ): Promise<T> {
    this.startSpan(name, attributes);
    try {
      const result = await fn();
      this.endSpan("success");
      return result;
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      this.endSpan("error", msg);
      throw error;
    }
  }

  /**
   * 同步执行带追踪的函数
   */
  traceSync<T>(name: string, fn: () => T, attributes?: Record<string, unknown>): T {
    this.startSpan(name, attributes);
    try {
      const result = fn();
      this.endSpan("success");
      return result;
    } catch (error) {
      const msg = error instanceof Error ? error.message : String(error);
      this.endSpan("error", msg);
      throw error;
    }
  }
}

// ========== 全局实例 ==========

let globalTracer: AgentTracer | null = null;

/**
 * 获取全局追踪器
 */
export function getTracer(): AgentTracer {
  if (!globalTracer) {
    globalTracer = new AgentTracer();
  }
  return globalTracer;
}

/**
 * 创建新的追踪器
 */
export function createTracer(config?: Partial<TracerConfig>): AgentTracer {
  return new AgentTracer(config);
}

/**
 * 重置全局追踪器
 */
export function resetTracer(): void {
  globalTracer?.clear();
  globalTracer = null;
}

export default AgentTracer;
