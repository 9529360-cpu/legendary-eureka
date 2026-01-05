/**
 * AuditLogger - 审计日志模块 v1.0
 *
 * 核心职责：
 * 1. 记录所有高风险操作的审批过程
 * 2. 记录操作执行结果
 * 3. 提供日志查询和导出
 *
 * 审计字段（上线必备）：
 * - requestId / approvalId
 * - userId
 * - intent
 * - preflightSummary
 * - approvedText
 * - toolName + args
 * - result
 * - verifyResult
 * - timestamp
 */

// ==================== 类型定义 ====================

/**
 * 审计日志条目
 */
export interface AuditEntry {
  /** 唯一标识 */
  id: string;
  /** 时间戳 */
  timestamp: Date;
  /** 操作类型 */
  action: AuditAction;
  /** 审批 ID */
  approvalId?: string;
  /** 请求 ID */
  requestId?: string;
  /** 用户 ID */
  userId?: string;
  /** 会话 ID */
  sessionId?: string;
  /** 操作名称 */
  operationName?: string;
  /** 操作类型 */
  operationType?: string;
  /** 操作参数 */
  parameters?: Record<string, unknown>;
  /** 风险等级 */
  riskLevel?: string;
  /** 用户原始输入 */
  userIntent?: string;
  /** 预检摘要 */
  preflightSummary?: string;
  /** 用户确认文本 */
  approvedText?: string;
  /** 决定者 */
  decidedBy?: string;
  /** 决定原因 */
  reason?: string;
  /** 执行结果 */
  result?: AuditResult;
  /** 验证结果 */
  verifyResult?: AuditVerifyResult;
  /** 额外元数据 */
  metadata?: Record<string, unknown>;
}

/**
 * 审计操作类型
 */
export type AuditAction =
  | "approval_requested" // 请求审批
  | "approval_granted" // 审批通过
  | "approval_rejected" // 审批拒绝
  | "approval_expired" // 审批过期
  | "operation_started" // 操作开始
  | "operation_completed" // 操作完成
  | "operation_failed" // 操作失败
  | "operation_rolled_back" // 操作回滚
  | "preflight_completed" // 预检完成
  | "verify_completed" // 验证完成
  | "user_input" // 用户输入
  | "agent_response"; // Agent 响应

/**
 * 审计结果
 */
export interface AuditResult {
  success: boolean;
  message?: string;
  affectedRange?: string;
  affectedRows?: number;
  affectedCells?: number;
  executionTime?: number;
  error?: string;
}

/**
 * 审计验证结果
 */
export interface AuditVerifyResult {
  passed: boolean;
  checks: {
    name: string;
    passed: boolean;
    message?: string;
  }[];
}

/**
 * 审计日志配置
 */
export interface AuditLoggerConfig {
  /** 最大保留条目数 */
  maxEntries: number;
  /** 是否持久化存储 */
  persist: boolean;
  /** 存储键名 */
  storageKey: string;
  /** 是否输出到控制台 */
  consoleOutput: boolean;
  /** 日志级别 */
  logLevel: "debug" | "info" | "warn" | "error";
}

/**
 * 默认配置
 */
export const DEFAULT_AUDIT_CONFIG: AuditLoggerConfig = {
  maxEntries: 1000,
  persist: true,
  storageKey: "excel_agent_audit_logs",
  consoleOutput: false,
  logLevel: "info",
};

// ==================== AuditLogger 类 ====================

/**
 * 审计日志记录器
 */
export class AuditLogger {
  private config: AuditLoggerConfig;
  private logs: AuditEntry[] = [];
  private logCounter: number = 0;

  constructor(config: Partial<AuditLoggerConfig> = {}) {
    this.config = { ...DEFAULT_AUDIT_CONFIG, ...config };
    this.loadFromStorage();
  }

  /**
   * 生成日志 ID
   */
  private generateLogId(): string {
    const now = new Date();
    const timestamp = now
      .toISOString()
      .replace(/[-:T.Z]/g, "")
      .slice(0, 14);
    this.logCounter++;
    return `LOG-${timestamp}-${String(this.logCounter).padStart(4, "0")}`;
  }

  /**
   * 记录日志
   */
  log(entry: Omit<AuditEntry, "id" | "timestamp">): AuditEntry {
    const fullEntry: AuditEntry = {
      ...entry,
      id: this.generateLogId(),
      timestamp: new Date(),
    };

    this.logs.push(fullEntry);

    // 控制台输出
    if (this.config.consoleOutput) {
      this.consoleLog(fullEntry);
    }

    // 限制条目数量
    if (this.logs.length > this.config.maxEntries) {
      this.logs = this.logs.slice(-this.config.maxEntries);
    }

    // 持久化
    if (this.config.persist) {
      this.saveToStorage();
    }

    return fullEntry;
  }

  /**
   * 记录审批请求
   */
  logApprovalRequest(
    approvalId: string,
    operationName: string,
    parameters: Record<string, unknown>,
    riskLevel: string,
    options?: { userId?: string; sessionId?: string; userIntent?: string }
  ): AuditEntry {
    return this.log({
      action: "approval_requested",
      approvalId,
      operationName,
      parameters,
      riskLevel,
      ...options,
    });
  }

  /**
   * 记录审批决定
   */
  logApprovalDecision(
    approvalId: string,
    approved: boolean,
    decidedBy?: string,
    approvedText?: string
  ): AuditEntry {
    return this.log({
      action: approved ? "approval_granted" : "approval_rejected",
      approvalId,
      decidedBy,
      approvedText,
    });
  }

  /**
   * 记录操作执行
   */
  logOperationExecution(
    approvalId: string | undefined,
    operationName: string,
    parameters: Record<string, unknown>,
    result: AuditResult
  ): AuditEntry {
    return this.log({
      action: result.success ? "operation_completed" : "operation_failed",
      approvalId,
      operationName,
      parameters,
      result,
    });
  }

  /**
   * 记录预检结果
   */
  logPreflight(
    operationName: string,
    preflightSummary: string,
    metadata?: Record<string, unknown>
  ): AuditEntry {
    return this.log({
      action: "preflight_completed",
      operationName,
      preflightSummary,
      metadata,
    });
  }

  /**
   * 记录验证结果
   */
  logVerification(
    approvalId: string | undefined,
    operationName: string,
    verifyResult: AuditVerifyResult
  ): AuditEntry {
    return this.log({
      action: "verify_completed",
      approvalId,
      operationName,
      verifyResult,
    });
  }

  /**
   * 获取所有日志
   */
  getLogs(): AuditEntry[] {
    return [...this.logs];
  }

  /**
   * 按条件查询日志
   */
  query(filter: {
    action?: AuditAction | AuditAction[];
    approvalId?: string;
    userId?: string;
    sessionId?: string;
    startTime?: Date;
    endTime?: Date;
    riskLevel?: string;
  }): AuditEntry[] {
    return this.logs.filter((entry) => {
      if (filter.action) {
        const actions = Array.isArray(filter.action) ? filter.action : [filter.action];
        if (!actions.includes(entry.action)) return false;
      }
      if (filter.approvalId && entry.approvalId !== filter.approvalId) return false;
      if (filter.userId && entry.userId !== filter.userId) return false;
      if (filter.sessionId && entry.sessionId !== filter.sessionId) return false;
      if (filter.startTime && entry.timestamp < filter.startTime) return false;
      if (filter.endTime && entry.timestamp > filter.endTime) return false;
      if (filter.riskLevel && entry.riskLevel !== filter.riskLevel) return false;
      return true;
    });
  }

  /**
   * 获取审批相关的完整链路
   */
  getApprovalChain(approvalId: string): AuditEntry[] {
    return this.logs
      .filter((entry) => entry.approvalId === approvalId)
      .sort((a, b) => a.timestamp.getTime() - b.timestamp.getTime());
  }

  /**
   * 导出日志为 JSON
   */
  exportAsJson(): string {
    return JSON.stringify(this.logs, null, 2);
  }

  /**
   * 导出日志为 CSV
   */
  exportAsCsv(): string {
    if (this.logs.length === 0) return "";

    const headers = [
      "id",
      "timestamp",
      "action",
      "approvalId",
      "userId",
      "sessionId",
      "operationName",
      "riskLevel",
      "result_success",
      "result_message",
    ];

    const rows = this.logs.map((entry) =>
      [
        entry.id,
        entry.timestamp.toISOString(),
        entry.action,
        entry.approvalId || "",
        entry.userId || "",
        entry.sessionId || "",
        entry.operationName || "",
        entry.riskLevel || "",
        entry.result?.success?.toString() || "",
        entry.result?.message || "",
      ]
        .map((v) => `"${String(v).replace(/"/g, '""')}"`)
        .join(",")
    );

    return [headers.join(","), ...rows].join("\n");
  }

  /**
   * 清空日志
   */
  clear(): void {
    this.logs = [];
    if (this.config.persist) {
      this.saveToStorage();
    }
  }

  /**
   * 控制台输出
   */
  private consoleLog(entry: AuditEntry): void {
    const prefix = `[AUDIT ${entry.id}]`;
    const actionEmoji: Record<AuditAction, string> = {
      approval_requested: "📝",
      approval_granted: "✅",
      approval_rejected: "❌",
      approval_expired: "⏰",
      operation_started: "🚀",
      operation_completed: "✔️",
      operation_failed: "💥",
      operation_rolled_back: "↩️",
      preflight_completed: "🔍",
      verify_completed: "🔬",
      user_input: "👤",
      agent_response: "🤖",
    };

    const emoji = actionEmoji[entry.action] || "📋";
    console.log(`${prefix} ${emoji} ${entry.action}`, {
      approvalId: entry.approvalId,
      operation: entry.operationName,
      riskLevel: entry.riskLevel,
    });
  }

  /**
   * 从存储加载日志
   */
  private loadFromStorage(): void {
    if (!this.config.persist) return;

    try {
      const stored = localStorage.getItem(this.config.storageKey);
      if (stored) {
        const parsed = JSON.parse(stored);
        this.logs = parsed.map((entry: any) => ({
          ...entry,
          timestamp: new Date(entry.timestamp),
        }));
      }
    } catch (error) {
      console.warn("[AuditLogger] Failed to load from storage:", error);
    }
  }

  /**
   * 保存日志到存储
   */
  private saveToStorage(): void {
    if (!this.config.persist) return;

    try {
      localStorage.setItem(this.config.storageKey, JSON.stringify(this.logs));
    } catch (error) {
      console.warn("[AuditLogger] Failed to save to storage:", error);
    }
  }
}

// ==================== 导出单例 ====================

export const auditLogger = new AuditLogger();

export default AuditLogger;
