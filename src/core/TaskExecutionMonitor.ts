/**
 * TaskExecutionMonitor - 任务执行监控器
 * v1.0.0
 *
 * 功能：
 * 1. 全链路任务追踪
 * 2. 工具注册与实现一致性检查
 * 3. 任务分发兜底策略
 * 4. 详细日志记录与告警
 * 5. 执行指标统计
 *
 * 解决的问题：
 * - 任务链路断裂难以定位
 * - 工具未注册时无告警
 * - 缺乏全链路追踪能力
 */

import { Logger } from "../utils/Logger";
import { TOOL_NAMES } from "../config/constants";

// ========== 类型定义 ==========

/**
 * 任务执行阶段
 */
export enum TaskPhase {
  /** 接收请求 */
  RECEIVED = "received",
  /** 意图解析 */
  INTENT_PARSING = "intent_parsing",
  /** 计划生成 */
  PLANNING = "planning",
  /** 工具查找 */
  TOOL_LOOKUP = "tool_lookup",
  /** 参数验证 */
  PARAM_VALIDATION = "param_validation",
  /** 工具执行 */
  TOOL_EXECUTION = "tool_execution",
  /** 结果验证 */
  RESULT_VALIDATION = "result_validation",
  /** 响应生成 */
  RESPONSE_GENERATION = "response_generation",
  /** 完成 */
  COMPLETED = "completed",
  /** 失败 */
  FAILED = "failed",
}

/**
 * 任务执行记录
 */
export interface TaskExecutionRecord {
  taskId: string;
  request: string;
  startTime: Date;
  endTime?: Date;
  phases: PhaseRecord[];
  toolCalls: ToolCallRecord[];
  status: "running" | "completed" | "failed" | "fallback";
  result?: string;
  error?: string;
  metrics: TaskMetrics;
}

/**
 * 阶段记录
 */
export interface PhaseRecord {
  phase: TaskPhase;
  startTime: Date;
  endTime?: Date;
  duration?: number;
  status: "running" | "completed" | "failed" | "skipped";
  details?: Record<string, unknown>;
  error?: string;
}

/**
 * 工具调用记录
 */
export interface ToolCallRecord {
  toolName: string;
  phase: TaskPhase;
  startTime: Date;
  endTime?: Date;
  duration?: number;
  input: Record<string, unknown>;
  output?: unknown;
  status: "pending" | "running" | "success" | "failed" | "not_found";
  error?: string;
  fallbackUsed?: string;
}

/**
 * 任务指标
 */
export interface TaskMetrics {
  totalDuration?: number;
  toolCallCount: number;
  successfulToolCalls: number;
  failedToolCalls: number;
  fallbackCount: number;
  retryCount: number;
}

/**
 * 告警级别
 */
export enum AlertLevel {
  INFO = "info",
  WARNING = "warning",
  ERROR = "error",
  CRITICAL = "critical",
}

/**
 * 告警记录
 */
export interface AlertRecord {
  level: AlertLevel;
  code: string;
  message: string;
  taskId?: string;
  toolName?: string;
  timestamp: Date;
  acknowledged: boolean;
  details?: Record<string, unknown>;
}

/**
 * 监控配置
 */
export interface MonitorConfig {
  enableDetailedLogging: boolean;
  alertOnToolNotFound: boolean;
  alertOnFallback: boolean;
  alertOnSlowExecution: boolean;
  slowExecutionThresholdMs: number;
  maxAlertHistory: number;
  maxTaskHistory: number;
}

// ========== 监控器实现 ==========

/**
 * 任务执行监控器
 */
class TaskExecutionMonitorClass {
  private config: MonitorConfig;
  private taskHistory: Map<string, TaskExecutionRecord> = new Map();
  private alertHistory: AlertRecord[] = [];
  private registeredTools: Set<string> = new Set();
  private alertListeners: ((alert: AlertRecord) => void)[] = [];

  constructor() {
    this.config = {
      enableDetailedLogging: true,
      alertOnToolNotFound: true,
      alertOnFallback: true,
      alertOnSlowExecution: true,
      slowExecutionThresholdMs: 5000,
      maxAlertHistory: 100,
      maxTaskHistory: 50,
    };

    // 注册已知工具名称
    this.initializeRegisteredTools();
  }

  /**
   * 配置监控器
   */
  configure(config: Partial<MonitorConfig>): void {
    this.config = { ...this.config, ...config };
  }

  /**
   * 注册工具（用于一致性检查）
   */
  registerTool(toolName: string): void {
    this.registeredTools.add(toolName);
    Logger.debug("TaskMonitor", `工具已注册: ${toolName}`);
  }

  /**
   * 批量注册工具
   */
  registerTools(toolNames: string[]): void {
    toolNames.forEach((name) => this.registeredTools.add(name));
    Logger.info("TaskMonitor", `批量注册 ${toolNames.length} 个工具`);
  }

  /**
   * 检查工具是否已注册
   */
  isToolRegistered(toolName: string): boolean {
    return this.registeredTools.has(toolName);
  }

  /**
   * 获取所有已注册工具
   */
  getRegisteredTools(): string[] {
    return Array.from(this.registeredTools);
  }

  /**
   * 添加告警监听器
   */
  addAlertListener(listener: (alert: AlertRecord) => void): () => void {
    this.alertListeners.push(listener);
    return () => {
      this.alertListeners = this.alertListeners.filter((l) => l !== listener);
    };
  }

  // ========== 任务生命周期 ==========

  /**
   * 开始任务追踪
   */
  startTask(taskId: string, request: string): TaskExecutionRecord {
    const record: TaskExecutionRecord = {
      taskId,
      request,
      startTime: new Date(),
      phases: [],
      toolCalls: [],
      status: "running",
      metrics: {
        toolCallCount: 0,
        successfulToolCalls: 0,
        failedToolCalls: 0,
        fallbackCount: 0,
        retryCount: 0,
      },
    };

    this.taskHistory.set(taskId, record);
    this.enforceHistoryLimit();

    if (this.config.enableDetailedLogging) {
      Logger.info("TaskMonitor", `📋 任务开始: ${taskId}`, { request: request.substring(0, 100) });
    }

    return record;
  }

  /**
   * 开始阶段
   */
  startPhase(taskId: string, phase: TaskPhase, details?: Record<string, unknown>): void {
    const record = this.taskHistory.get(taskId);
    if (!record) {
      Logger.warn("TaskMonitor", `任务不存在: ${taskId}`);
      return;
    }

    const phaseRecord: PhaseRecord = {
      phase,
      startTime: new Date(),
      status: "running",
      details,
    };

    record.phases.push(phaseRecord);

    if (this.config.enableDetailedLogging) {
      Logger.debug("TaskMonitor", `  → 阶段开始: ${phase}`, details);
    }
  }

  /**
   * 完成阶段
   */
  completePhase(taskId: string, phase: TaskPhase, details?: Record<string, unknown>): void {
    const record = this.taskHistory.get(taskId);
    if (!record) return;

    const phaseRecord = record.phases.find((p) => p.phase === phase && p.status === "running");
    if (phaseRecord) {
      phaseRecord.endTime = new Date();
      phaseRecord.duration = phaseRecord.endTime.getTime() - phaseRecord.startTime.getTime();
      phaseRecord.status = "completed";
      if (details) {
        phaseRecord.details = { ...phaseRecord.details, ...details };
      }

      if (this.config.enableDetailedLogging) {
        Logger.debug("TaskMonitor", `  ✓ 阶段完成: ${phase} (${phaseRecord.duration}ms)`);
      }
    }
  }

  /**
   * 阶段失败
   */
  failPhase(taskId: string, phase: TaskPhase, error: string): void {
    const record = this.taskHistory.get(taskId);
    if (!record) return;

    const phaseRecord = record.phases.find((p) => p.phase === phase && p.status === "running");
    if (phaseRecord) {
      phaseRecord.endTime = new Date();
      phaseRecord.duration = phaseRecord.endTime.getTime() - phaseRecord.startTime.getTime();
      phaseRecord.status = "failed";
      phaseRecord.error = error;

      Logger.error("TaskMonitor", `  ✗ 阶段失败: ${phase}`, { error });
    }
  }

  // ========== 工具调用追踪 ==========

  /**
   * 开始工具调用
   */
  startToolCall(taskId: string, toolName: string, input: Record<string, unknown>): ToolCallRecord {
    const record = this.taskHistory.get(taskId);

    const toolCall: ToolCallRecord = {
      toolName,
      phase: TaskPhase.TOOL_EXECUTION,
      startTime: new Date(),
      input,
      status: "running",
    };

    // 检查工具是否已注册
    if (!this.isToolRegistered(toolName)) {
      toolCall.status = "not_found";
      this.raiseAlert(AlertLevel.ERROR, "TOOL_NOT_REGISTERED", `工具未注册或未实现: ${toolName}`, {
        taskId,
        toolName,
        input,
      });
    }

    if (record) {
      record.toolCalls.push(toolCall);
      record.metrics.toolCallCount++;
    }

    if (this.config.enableDetailedLogging) {
      Logger.debug("TaskMonitor", `    🔧 工具调用: ${toolName}`, { input });
    }

    return toolCall;
  }

  /**
   * 完成工具调用
   */
  completeToolCall(
    taskId: string,
    toolName: string,
    output: unknown,
    success: boolean = true
  ): void {
    const record = this.taskHistory.get(taskId);
    if (!record) return;

    const toolCall = record.toolCalls.find(
      (tc) => tc.toolName === toolName && tc.status === "running"
    );

    if (toolCall) {
      toolCall.endTime = new Date();
      toolCall.duration = toolCall.endTime.getTime() - toolCall.startTime.getTime();
      toolCall.output = output;
      toolCall.status = success ? "success" : "failed";

      if (success) {
        record.metrics.successfulToolCalls++;
      } else {
        record.metrics.failedToolCalls++;
      }

      // 检查慢执行
      if (
        this.config.alertOnSlowExecution &&
        toolCall.duration > this.config.slowExecutionThresholdMs
      ) {
        this.raiseAlert(
          AlertLevel.WARNING,
          "SLOW_TOOL_EXECUTION",
          `工具执行时间过长: ${toolName} (${toolCall.duration}ms)`,
          { taskId, toolName, duration: toolCall.duration }
        );
      }

      if (this.config.enableDetailedLogging) {
        const icon = success ? "✓" : "✗";
        Logger.debug("TaskMonitor", `    ${icon} 工具完成: ${toolName} (${toolCall.duration}ms)`);
      }
    }
  }

  /**
   * 工具调用失败
   */
  failToolCall(taskId: string, toolName: string, error: string): void {
    const record = this.taskHistory.get(taskId);
    if (!record) return;

    const toolCall = record.toolCalls.find(
      (tc) => tc.toolName === toolName && tc.status === "running"
    );

    if (toolCall) {
      toolCall.endTime = new Date();
      toolCall.duration = toolCall.endTime.getTime() - toolCall.startTime.getTime();
      toolCall.status = "failed";
      toolCall.error = error;
      record.metrics.failedToolCalls++;

      this.raiseAlert(
        AlertLevel.ERROR,
        "TOOL_EXECUTION_FAILED",
        `工具执行失败: ${toolName} - ${error}`,
        { taskId, toolName, error }
      );
    }
  }

  /**
   * 记录兜底操作
   */
  recordFallback(taskId: string, originalTool: string, fallbackTool: string, reason: string): void {
    const record = this.taskHistory.get(taskId);
    if (!record) return;

    record.metrics.fallbackCount++;

    const toolCall = record.toolCalls.find((tc) => tc.toolName === originalTool);
    if (toolCall) {
      toolCall.fallbackUsed = fallbackTool;
    }

    if (this.config.alertOnFallback) {
      this.raiseAlert(
        AlertLevel.WARNING,
        "FALLBACK_USED",
        `使用兜底策略: ${originalTool} → ${fallbackTool}`,
        { taskId, originalTool, fallbackTool, reason }
      );
    }

    Logger.warn("TaskMonitor", `兜底策略: ${originalTool} → ${fallbackTool}`, { reason });
  }

  // ========== 任务完成 ==========

  /**
   * 完成任务
   */
  completeTask(taskId: string, result: string): TaskExecutionRecord | undefined {
    const record = this.taskHistory.get(taskId);
    if (!record) return;

    record.endTime = new Date();
    record.status = "completed";
    record.result = result;
    record.metrics.totalDuration = record.endTime.getTime() - record.startTime.getTime();

    if (this.config.enableDetailedLogging) {
      Logger.info("TaskMonitor", `✅ 任务完成: ${taskId}`, {
        duration: `${record.metrics.totalDuration}ms`,
        toolCalls: record.metrics.toolCallCount,
        success: record.metrics.successfulToolCalls,
        failed: record.metrics.failedToolCalls,
      });
    }

    return record;
  }

  /**
   * 任务失败
   */
  failTask(taskId: string, error: string): TaskExecutionRecord | undefined {
    const record = this.taskHistory.get(taskId);
    if (!record) return;

    record.endTime = new Date();
    record.status = "failed";
    record.error = error;
    record.metrics.totalDuration = record.endTime.getTime() - record.startTime.getTime();

    this.raiseAlert(AlertLevel.ERROR, "TASK_FAILED", `任务执行失败: ${taskId} - ${error}`, {
      taskId,
      error,
      metrics: record.metrics,
    });

    Logger.error("TaskMonitor", `❌ 任务失败: ${taskId}`, { error });

    return record;
  }

  // ========== 告警管理 ==========

  /**
   * 触发告警
   */
  raiseAlert(
    level: AlertLevel,
    code: string,
    message: string,
    details?: Record<string, unknown>
  ): AlertRecord {
    const alert: AlertRecord = {
      level,
      code,
      message,
      taskId: details?.taskId as string,
      toolName: details?.toolName as string,
      timestamp: new Date(),
      acknowledged: false,
      details,
    };

    this.alertHistory.push(alert);
    this.enforceAlertLimit();

    // 根据级别记录日志
    switch (level) {
      case AlertLevel.CRITICAL:
      case AlertLevel.ERROR:
        Logger.error("TaskMonitor", `🚨 [${code}] ${message}`, details);
        break;
      case AlertLevel.WARNING:
        Logger.warn("TaskMonitor", `⚠️ [${code}] ${message}`, details);
        break;
      default:
        Logger.info("TaskMonitor", `ℹ️ [${code}] ${message}`, details);
    }

    // 通知监听器
    this.alertListeners.forEach((listener) => listener(alert));

    return alert;
  }

  /**
   * 确认告警
   */
  acknowledgeAlert(index: number): void {
    if (index >= 0 && index < this.alertHistory.length) {
      this.alertHistory[index].acknowledged = true;
    }
  }

  /**
   * 获取未确认的告警
   */
  getUnacknowledgedAlerts(): AlertRecord[] {
    return this.alertHistory.filter((a) => !a.acknowledged);
  }

  /**
   * 获取所有告警
   */
  getAlertHistory(): AlertRecord[] {
    return [...this.alertHistory];
  }

  // ========== 统计与分析 ==========

  /**
   * 获取任务记录
   */
  getTaskRecord(taskId: string): TaskExecutionRecord | undefined {
    return this.taskHistory.get(taskId);
  }

  /**
   * 获取所有任务记录
   */
  getAllTaskRecords(): TaskExecutionRecord[] {
    return Array.from(this.taskHistory.values());
  }

  /**
   * 获取执行统计
   */
  getStatistics(): {
    totalTasks: number;
    completedTasks: number;
    failedTasks: number;
    averageDuration: number;
    toolUsageStats: Record<string, { calls: number; failures: number; avgDuration: number }>;
    unregisteredToolCalls: string[];
  } {
    const tasks = this.getAllTaskRecords();
    const completedTasks = tasks.filter((t) => t.status === "completed");
    const failedTasks = tasks.filter((t) => t.status === "failed");

    const avgDuration =
      completedTasks.length > 0
        ? completedTasks.reduce((sum, t) => sum + (t.metrics.totalDuration || 0), 0) /
          completedTasks.length
        : 0;

    // 工具使用统计
    const toolStats: Record<string, { calls: number; failures: number; durations: number[] }> = {};
    const unregisteredTools = new Set<string>();

    tasks.forEach((task) => {
      task.toolCalls.forEach((tc) => {
        if (!toolStats[tc.toolName]) {
          toolStats[tc.toolName] = { calls: 0, failures: 0, durations: [] };
        }
        toolStats[tc.toolName].calls++;
        if (tc.status === "failed") {
          toolStats[tc.toolName].failures++;
        }
        if (tc.duration) {
          toolStats[tc.toolName].durations.push(tc.duration);
        }
        if (tc.status === "not_found") {
          unregisteredTools.add(tc.toolName);
        }
      });
    });

    const toolUsageStats: Record<string, { calls: number; failures: number; avgDuration: number }> =
      {};
    Object.entries(toolStats).forEach(([name, stats]) => {
      toolUsageStats[name] = {
        calls: stats.calls,
        failures: stats.failures,
        avgDuration:
          stats.durations.length > 0
            ? stats.durations.reduce((a, b) => a + b, 0) / stats.durations.length
            : 0,
      };
    });

    return {
      totalTasks: tasks.length,
      completedTasks: completedTasks.length,
      failedTasks: failedTasks.length,
      averageDuration: avgDuration,
      toolUsageStats,
      unregisteredToolCalls: Array.from(unregisteredTools),
    };
  }

  /**
   * 检查工具注册一致性
   */
  checkToolConsistency(): {
    registered: string[];
    usedButNotRegistered: string[];
    registeredButNeverUsed: string[];
  } {
    const usedTools = new Set<string>();
    this.getAllTaskRecords().forEach((task) => {
      task.toolCalls.forEach((tc) => usedTools.add(tc.toolName));
    });

    const registered = this.getRegisteredTools();
    const usedButNotRegistered = Array.from(usedTools).filter((t) => !this.isToolRegistered(t));
    const registeredButNeverUsed = registered.filter((t) => !usedTools.has(t));

    return {
      registered,
      usedButNotRegistered,
      registeredButNeverUsed,
    };
  }

  // ========== 私有方法 ==========

  private initializeRegisteredTools(): void {
    // 从常量中注册已知工具
    Object.values(TOOL_NAMES).forEach((name) => {
      this.registeredTools.add(name);
    });
  }

  private enforceHistoryLimit(): void {
    if (this.taskHistory.size > this.config.maxTaskHistory) {
      const oldestKey = this.taskHistory.keys().next().value;
      if (oldestKey) {
        this.taskHistory.delete(oldestKey);
      }
    }
  }

  private enforceAlertLimit(): void {
    while (this.alertHistory.length > this.config.maxAlertHistory) {
      this.alertHistory.shift();
    }
  }

  /**
   * 重置监控器（用于测试）
   */
  reset(): void {
    this.taskHistory.clear();
    this.alertHistory = [];
    Logger.info("TaskMonitor", "监控器已重置");
  }
}

// 导出单例
export const TaskExecutionMonitor = new TaskExecutionMonitorClass();

// 便捷方法导出
export const monitor = {
  startTask: (taskId: string, request: string) => TaskExecutionMonitor.startTask(taskId, request),
  startPhase: (taskId: string, phase: TaskPhase, details?: Record<string, unknown>) =>
    TaskExecutionMonitor.startPhase(taskId, phase, details),
  completePhase: (taskId: string, phase: TaskPhase, details?: Record<string, unknown>) =>
    TaskExecutionMonitor.completePhase(taskId, phase, details),
  failPhase: (taskId: string, phase: TaskPhase, error: string) =>
    TaskExecutionMonitor.failPhase(taskId, phase, error),
  startToolCall: (taskId: string, toolName: string, input: Record<string, unknown>) =>
    TaskExecutionMonitor.startToolCall(taskId, toolName, input),
  completeToolCall: (taskId: string, toolName: string, output: unknown, success?: boolean) =>
    TaskExecutionMonitor.completeToolCall(taskId, toolName, output, success),
  failToolCall: (taskId: string, toolName: string, error: string) =>
    TaskExecutionMonitor.failToolCall(taskId, toolName, error),
  recordFallback: (taskId: string, original: string, fallback: string, reason: string) =>
    TaskExecutionMonitor.recordFallback(taskId, original, fallback, reason),
  completeTask: (taskId: string, result: string) =>
    TaskExecutionMonitor.completeTask(taskId, result),
  failTask: (taskId: string, error: string) => TaskExecutionMonitor.failTask(taskId, error),
  raiseAlert: (
    level: AlertLevel,
    code: string,
    message: string,
    details?: Record<string, unknown>
  ) => TaskExecutionMonitor.raiseAlert(level, code, message, details),
  getStatistics: () => TaskExecutionMonitor.getStatistics(),
  registerTool: (name: string) => TaskExecutionMonitor.registerTool(name),
  registerTools: (names: string[]) => TaskExecutionMonitor.registerTools(names),
  isToolRegistered: (name: string) => TaskExecutionMonitor.isToolRegistered(name),
};

export default TaskExecutionMonitor;
