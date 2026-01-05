/**
 * TraceContext - 全链路追踪上下文
 * v1.0.0
 *
 * 功能：
 * 1. 请求级别的全链路追踪
 * 2. 上下文传播
 * 3. 性能指标收集
 * 4. 错误链追踪
 * 5. 分布式追踪支持
 *
 * 解决的问题：
 * - Logger仅覆盖部分模块
 * - 缺少全链路追踪
 * - 难以定位复杂问题
 */

import { Logger } from "../utils/Logger";

// ========== 类型定义 ==========

/**
 * 追踪级别
 */
export enum TraceLevel {
  /** 详细 - 所有信息 */
  VERBOSE = 0,
  /** 调试 - 开发信息 */
  DEBUG = 1,
  /** 信息 - 关键节点 */
  INFO = 2,
  /** 警告 - 潜在问题 */
  WARN = 3,
  /** 错误 - 错误信息 */
  ERROR = 4,
}

/**
 * Span 类型
 */
export enum SpanType {
  /** HTTP 请求 */
  HTTP = "http",
  /** 数据库操作 */
  DB = "db",
  /** Excel 操作 */
  EXCEL = "excel",
  /** AI 调用 */
  AI = "ai",
  /** 工具执行 */
  TOOL = "tool",
  /** 内部处理 */
  INTERNAL = "internal",
  /** 用户交互 */
  USER = "user",
}

/**
 * Span 状态
 */
export enum SpanStatus {
  /** 未开始 */
  UNSET = "unset",
  /** 运行中 */
  RUNNING = "running",
  /** 成功 */
  OK = "ok",
  /** 失败 */
  ERROR = "error",
  /** 已取消 */
  CANCELLED = "cancelled",
}

/**
 * Span 属性
 */
export interface SpanAttributes {
  [key: string]: string | number | boolean | string[] | number[] | undefined;
}

/**
 * Span 事件
 */
export interface SpanEvent {
  name: string;
  timestamp: Date;
  attributes?: SpanAttributes;
}

/**
 * Span 定义
 */
export interface Span {
  /** Span ID */
  spanId: string;
  /** 父 Span ID */
  parentSpanId?: string;
  /** Trace ID */
  traceId: string;
  /** 操作名称 */
  operationName: string;
  /** Span 类型 */
  type: SpanType;
  /** 状态 */
  status: SpanStatus;
  /** 开始时间 */
  startTime: Date;
  /** 结束时间 */
  endTime?: Date;
  /** 持续时间（毫秒） */
  duration?: number;
  /** 属性 */
  attributes: SpanAttributes;
  /** 事件 */
  events: SpanEvent[];
  /** 错误信息 */
  error?: {
    message: string;
    stack?: string;
    code?: string;
  };
  /** 子 Span */
  children: Span[];
}

/**
 * Trace 定义
 */
export interface Trace {
  /** Trace ID */
  traceId: string;
  /** 根 Span */
  rootSpan: Span;
  /** 开始时间 */
  startTime: Date;
  /** 结束时间 */
  endTime?: Date;
  /** 总持续时间 */
  totalDuration?: number;
  /** 请求信息 */
  request?: {
    type: string;
    content: string;
  };
  /** 响应信息 */
  response?: {
    success: boolean;
    content?: string;
    error?: string;
  };
  /** 元数据 */
  metadata: Record<string, unknown>;
}

/**
 * 追踪配置
 */
export interface TraceConfig {
  /** 是否启用追踪 */
  enabled: boolean;
  /** 追踪级别 */
  level: TraceLevel;
  /** 最大 Span 数量 */
  maxSpans: number;
  /** 最大 Trace 历史 */
  maxTraceHistory: number;
  /** 是否记录属性 */
  recordAttributes: boolean;
  /** 是否记录事件 */
  recordEvents: boolean;
  /** 采样率（0-1） */
  samplingRate: number;
}

// ========== 追踪上下文实现 ==========

/**
 * 追踪上下文管理器
 */
class TraceContextClass {
  private config: TraceConfig;
  private activeTraces: Map<string, Trace> = new Map();
  private traceHistory: Trace[] = [];
  private currentTraceId: string | null = null;
  private currentSpanId: string | null = null;
  private spanStack: Span[] = [];

  constructor() {
    this.config = {
      enabled: true,
      level: TraceLevel.INFO,
      maxSpans: 1000,
      maxTraceHistory: 50,
      recordAttributes: true,
      recordEvents: true,
      samplingRate: 1.0,
    };
  }

  /**
   * 配置追踪器
   */
  configure(config: Partial<TraceConfig>): void {
    this.config = { ...this.config, ...config };
  }

  /**
   * 生成唯一 ID
   */
  private generateId(): string {
    return `${Date.now().toString(36)}-${Math.random().toString(36).substring(2, 9)}`;
  }

  // ========== Trace 管理 ==========

  /**
   * 开始新的 Trace
   */
  startTrace(operationName: string, metadata?: Record<string, unknown>): Trace {
    if (!this.config.enabled) {
      return this.createEmptyTrace();
    }

    // 采样检查
    if (Math.random() > this.config.samplingRate) {
      return this.createEmptyTrace();
    }

    const traceId = this.generateId();
    const rootSpan = this.createSpan(operationName, SpanType.INTERNAL, traceId);

    const trace: Trace = {
      traceId,
      rootSpan,
      startTime: new Date(),
      metadata: metadata || {},
    };

    this.activeTraces.set(traceId, trace);
    this.currentTraceId = traceId;
    this.currentSpanId = rootSpan.spanId;
    this.spanStack = [rootSpan];

    Logger.debug("TraceContext", `Trace 开始: ${traceId}`, { operationName });

    return trace;
  }

  /**
   * 结束 Trace
   */
  endTrace(
    traceId?: string,
    response?: { success: boolean; content?: string; error?: string }
  ): Trace | undefined {
    const id = traceId || this.currentTraceId;
    if (!id) return;

    const trace = this.activeTraces.get(id);
    if (!trace) return;

    // 结束根 Span
    this.endSpan(trace.rootSpan.spanId);

    trace.endTime = new Date();
    trace.totalDuration = trace.endTime.getTime() - trace.startTime.getTime();
    trace.response = response;

    // 移动到历史
    this.activeTraces.delete(id);
    this.traceHistory.push(trace);
    this.enforceHistoryLimit();

    // 清理当前上下文
    if (this.currentTraceId === id) {
      this.currentTraceId = null;
      this.currentSpanId = null;
      this.spanStack = [];
    }

    Logger.debug("TraceContext", `Trace 结束: ${id}`, {
      duration: trace.totalDuration,
      success: response?.success,
    });

    return trace;
  }

  /**
   * 获取当前 Trace
   */
  getCurrentTrace(): Trace | undefined {
    if (!this.currentTraceId) return;
    return this.activeTraces.get(this.currentTraceId);
  }

  // ========== Span 管理 ==========

  /**
   * 创建 Span
   */
  private createSpan(
    operationName: string,
    type: SpanType,
    traceId: string,
    parentSpanId?: string
  ): Span {
    return {
      spanId: this.generateId(),
      parentSpanId,
      traceId,
      operationName,
      type,
      status: SpanStatus.RUNNING,
      startTime: new Date(),
      attributes: {},
      events: [],
      children: [],
    };
  }

  /**
   * 开始新的 Span
   */
  startSpan(operationName: string, type: SpanType = SpanType.INTERNAL): Span | null {
    if (!this.config.enabled || !this.currentTraceId) {
      return null;
    }

    const parentSpan = this.spanStack[this.spanStack.length - 1];
    const span = this.createSpan(operationName, type, this.currentTraceId, parentSpan?.spanId);

    if (parentSpan) {
      parentSpan.children.push(span);
    }

    this.spanStack.push(span);
    this.currentSpanId = span.spanId;

    if (this.config.level <= TraceLevel.DEBUG) {
      Logger.debug("TraceContext", `Span 开始: ${operationName}`, { spanId: span.spanId });
    }

    return span;
  }

  /**
   * 结束 Span
   */
  endSpan(spanId?: string, status: SpanStatus = SpanStatus.OK): Span | null {
    if (!this.config.enabled) return null;

    const id = spanId || this.currentSpanId;
    if (!id) return null;

    // 查找 Span
    const span = this.findSpan(id);
    if (!span) return null;

    span.endTime = new Date();
    span.duration = span.endTime.getTime() - span.startTime.getTime();
    span.status = status;

    // 更新堆栈
    const index = this.spanStack.findIndex((s) => s.spanId === id);
    if (index !== -1) {
      this.spanStack = this.spanStack.slice(0, index);
      this.currentSpanId =
        this.spanStack.length > 0 ? this.spanStack[this.spanStack.length - 1].spanId : null;
    }

    return span;
  }

  /**
   * 添加 Span 属性
   */
  setSpanAttribute(key: string, value: string | number | boolean, spanId?: string): void {
    if (!this.config.enabled || !this.config.recordAttributes) return;

    const span = spanId ? this.findSpan(spanId) : this.getCurrentSpan();
    if (span) {
      span.attributes[key] = value;
    }
  }

  /**
   * 批量添加属性
   */
  setSpanAttributes(attributes: SpanAttributes, spanId?: string): void {
    if (!this.config.enabled || !this.config.recordAttributes) return;

    const span = spanId ? this.findSpan(spanId) : this.getCurrentSpan();
    if (span) {
      Object.assign(span.attributes, attributes);
    }
  }

  /**
   * 添加 Span 事件
   */
  addSpanEvent(name: string, attributes?: SpanAttributes, spanId?: string): void {
    if (!this.config.enabled || !this.config.recordEvents) return;

    const span = spanId ? this.findSpan(spanId) : this.getCurrentSpan();
    if (span) {
      span.events.push({
        name,
        timestamp: new Date(),
        attributes,
      });
    }
  }

  /**
   * 记录 Span 错误
   */
  recordSpanError(error: Error | string, spanId?: string): void {
    const span = spanId ? this.findSpan(spanId) : this.getCurrentSpan();
    if (span) {
      span.status = SpanStatus.ERROR;
      span.error =
        typeof error === "string"
          ? { message: error }
          : { message: error.message, stack: error.stack };
    }
  }

  /**
   * 获取当前 Span
   */
  getCurrentSpan(): Span | null {
    if (this.spanStack.length === 0) return null;
    return this.spanStack[this.spanStack.length - 1];
  }

  // ========== 便捷方法 ==========

  /**
   * 追踪异步操作
   */
  async traceAsync<T>(
    operationName: string,
    type: SpanType,
    operation: () => Promise<T>
  ): Promise<T> {
    const span = this.startSpan(operationName, type);

    try {
      const result = await operation();
      if (span) {
        this.endSpan(span.spanId, SpanStatus.OK);
      }
      return result;
    } catch (error) {
      if (span) {
        this.recordSpanError(error as Error, span.spanId);
        this.endSpan(span.spanId, SpanStatus.ERROR);
      }
      throw error;
    }
  }

  /**
   * 追踪同步操作
   */
  trace<T>(operationName: string, type: SpanType, operation: () => T): T {
    const span = this.startSpan(operationName, type);

    try {
      const result = operation();
      if (span) {
        this.endSpan(span.spanId, SpanStatus.OK);
      }
      return result;
    } catch (error) {
      if (span) {
        this.recordSpanError(error as Error, span.spanId);
        this.endSpan(span.spanId, SpanStatus.ERROR);
      }
      throw error;
    }
  }

  // ========== 查询与分析 ==========

  /**
   * 查找 Span
   */
  private findSpan(spanId: string): Span | null {
    const trace = this.getCurrentTrace();
    if (!trace) return null;

    const search = (span: Span): Span | null => {
      if (span.spanId === spanId) return span;
      for (const child of span.children) {
        const found = search(child);
        if (found) return found;
      }
      return null;
    };

    return search(trace.rootSpan);
  }

  /**
   * 获取 Trace 历史
   */
  getTraceHistory(): Trace[] {
    return [...this.traceHistory];
  }

  /**
   * 获取 Trace 摘要
   */
  getTraceSummary(traceId: string): {
    traceId: string;
    operationName: string;
    duration: number;
    spanCount: number;
    errorCount: number;
    success: boolean;
  } | null {
    const trace =
      this.traceHistory.find((t) => t.traceId === traceId) || this.activeTraces.get(traceId);

    if (!trace) return null;

    let spanCount = 0;
    let errorCount = 0;

    const countSpans = (span: Span) => {
      spanCount++;
      if (span.status === SpanStatus.ERROR) errorCount++;
      span.children.forEach(countSpans);
    };

    countSpans(trace.rootSpan);

    return {
      traceId: trace.traceId,
      operationName: trace.rootSpan.operationName,
      duration: trace.totalDuration || 0,
      spanCount,
      errorCount,
      success: trace.response?.success ?? errorCount === 0,
    };
  }

  /**
   * 获取性能统计
   */
  getPerformanceStats(): {
    avgDuration: number;
    maxDuration: number;
    minDuration: number;
    totalTraces: number;
    successRate: number;
    spanTypeDistribution: Record<string, number>;
  } {
    const traces = this.traceHistory;
    if (traces.length === 0) {
      return {
        avgDuration: 0,
        maxDuration: 0,
        minDuration: 0,
        totalTraces: 0,
        successRate: 0,
        spanTypeDistribution: {},
      };
    }

    const durations = traces.map((t) => t.totalDuration || 0).filter((d) => d > 0);

    const successCount = traces.filter((t) => t.response?.success).length;
    const spanTypes: Record<string, number> = {};

    traces.forEach((trace) => {
      const countTypes = (span: Span) => {
        spanTypes[span.type] = (spanTypes[span.type] || 0) + 1;
        span.children.forEach(countTypes);
      };
      countTypes(trace.rootSpan);
    });

    return {
      avgDuration:
        durations.length > 0 ? durations.reduce((a, b) => a + b, 0) / durations.length : 0,
      maxDuration: durations.length > 0 ? Math.max(...durations) : 0,
      minDuration: durations.length > 0 ? Math.min(...durations) : 0,
      totalTraces: traces.length,
      successRate: traces.length > 0 ? successCount / traces.length : 0,
      spanTypeDistribution: spanTypes,
    };
  }

  /**
   * 导出 Trace 为可视化格式
   */
  exportTrace(traceId: string): string | null {
    const trace =
      this.traceHistory.find((t) => t.traceId === traceId) || this.activeTraces.get(traceId);

    if (!trace) return null;

    const formatSpan = (span: Span, indent: number = 0): string => {
      const prefix = "  ".repeat(indent);
      const status =
        span.status === SpanStatus.OK
          ? "✓"
          : span.status === SpanStatus.ERROR
            ? "✗"
            : span.status === SpanStatus.RUNNING
              ? "…"
              : "○";
      const duration = span.duration ? `${span.duration}ms` : "running";

      let result = `${prefix}${status} ${span.operationName} [${span.type}] (${duration})`;

      if (span.error) {
        result += `\n${prefix}  └─ Error: ${span.error.message}`;
      }

      span.children.forEach((child) => {
        result += "\n" + formatSpan(child, indent + 1);
      });

      return result;
    };

    return formatSpan(trace.rootSpan);
  }

  // ========== 私有方法 ==========

  private createEmptyTrace(): Trace {
    return {
      traceId: "disabled",
      rootSpan: this.createSpan("disabled", SpanType.INTERNAL, "disabled"),
      startTime: new Date(),
      metadata: {},
    };
  }

  private enforceHistoryLimit(): void {
    while (this.traceHistory.length > this.config.maxTraceHistory) {
      this.traceHistory.shift();
    }
  }

  /**
   * 清空历史
   */
  clearHistory(): void {
    this.traceHistory = [];
  }

  /**
   * 重置（用于测试）
   */
  reset(): void {
    this.activeTraces.clear();
    this.traceHistory = [];
    this.currentTraceId = null;
    this.currentSpanId = null;
    this.spanStack = [];
  }
}

// 导出单例
export const TraceContext = new TraceContextClass();

// 便捷方法导出
export const trace = {
  start: (name: string, metadata?: Record<string, unknown>) =>
    TraceContext.startTrace(name, metadata),
  end: (traceId?: string, response?: { success: boolean; content?: string; error?: string }) =>
    TraceContext.endTrace(traceId, response),
  startSpan: (name: string, type?: SpanType) => TraceContext.startSpan(name, type),
  endSpan: (spanId?: string, status?: SpanStatus) => TraceContext.endSpan(spanId, status),
  setAttr: (key: string, value: string | number | boolean, spanId?: string) =>
    TraceContext.setSpanAttribute(key, value, spanId),
  addEvent: (name: string, attributes?: SpanAttributes, spanId?: string) =>
    TraceContext.addSpanEvent(name, attributes, spanId),
  error: (error: Error | string, spanId?: string) => TraceContext.recordSpanError(error, spanId),
  async: <T>(name: string, type: SpanType, op: () => Promise<T>) =>
    TraceContext.traceAsync(name, type, op),
  sync: <T>(name: string, type: SpanType, op: () => T) => TraceContext.trace(name, type, op),
  getCurrent: () => TraceContext.getCurrentTrace(),
  getStats: () => TraceContext.getPerformanceStats(),
};

export default TraceContext;
