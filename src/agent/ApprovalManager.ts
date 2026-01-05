/**
 * ApprovalManager - Agent 层审批管理器 v1.0
 *
 * 核心职责：
 * 1. 判定哪些操作需要用户审批（needsApproval）
 * 2. 生成审批 ID（approvalId）
 * 3. 管理审批中断/恢复流程
 * 4. 提供审批状态查询
 *
 * 设计原则：
 * - **Agent 层决定是否需要确认，不是 LLM**
 * - 高风险操作必须走审批闸门
 * - 支持动态风险评估（基于操作类型 + 参数）
 */

import { AuditLogger, AuditEntry } from "./AuditLogger";

// ==================== 类型定义 ====================

/**
 * 风险等级
 */
export type RiskLevel = "low" | "medium" | "high" | "critical";

/**
 * 审批状态
 */
export type ApprovalStatus = "pending" | "approved" | "rejected" | "expired";

/**
 * 高风险操作列表（必须确认）
 */
export const HIGH_RISK_OPERATIONS = [
  // 删除操作（不可逆）
  "delete_rows",
  "delete_columns",
  "delete_column",
  "delete_row",
  "delete_sheet",
  "remove_duplicates",

  // 清空操作（可能毁表）
  "clear_range",
  "clear_all",
  "clear_formats",
  "clear_contents",

  // 批量修改操作
  "batch_update",
  "batch_write",
  "batch_formula",
  "fill_formula",
  "fill_range",

  // 保护/权限相关
  "protect_sheet",
  "unprotect_sheet",
  "lock_cells",
  "unlock_cells",

  // 宏/脚本执行
  "run_macro",
  "run_script",
  "execute_vba",
] as const;

/**
 * 中风险操作列表（建议确认）
 */
export const MEDIUM_RISK_OPERATIONS = [
  // 覆盖写入
  "write_range",
  "set_range_values",
  "overwrite_range",

  // 公式相关（可能覆盖公式）
  "set_formula",
  "set_array_formula",

  // 结构变更
  "insert_rows",
  "insert_columns",
  "merge_cells",
  "unmerge_cells",

  // 排序/筛选（可能打乱数据）
  "sort_range",
  "apply_filter",
] as const;

/**
 * 批量操作关键词
 */
export const BATCH_KEYWORDS = [
  "全部",
  "所有",
  "整列",
  "整表",
  "批量",
  "全列",
  "all",
  "entire",
  "whole",
] as const;

/**
 * 操作风险评估结果
 */
export interface RiskAssessment {
  needsApproval: boolean;
  riskLevel: RiskLevel;
  reason: string;
  impactDescription: string;
  reversible: boolean;
  estimatedImpact?: {
    cellCount?: number;
    rowCount?: number;
    columnCount?: number;
    sheetCount?: number;
  };
}

/**
 * 审批请求
 */
export interface ApprovalRequest {
  approvalId: string;
  operationName: string;
  operationType: string;
  parameters: Record<string, unknown>;
  riskAssessment: RiskAssessment;
  requestTime: Date;
  expiresAt: Date;
  status: ApprovalStatus;
  userId?: string;
  sessionId?: string;
}

/**
 * 审批决定
 */
export interface ApprovalDecision {
  approvalId: string;
  approved: boolean;
  decidedAt: Date;
  decidedBy?: string;
  reason?: string;
}

/**
 * 审批回调
 */
export type ApprovalCallback = (request: ApprovalRequest) => Promise<boolean>;

/**
 * 审批管理器配置
 */
export interface ApprovalManagerConfig {
  /** 审批请求超时时间（毫秒） */
  approvalTimeout: number;
  /** 是否启用批量操作自动确认阈值 */
  batchThreshold: number;
  /** 是否启用审计日志 */
  enableAudit: boolean;
  /** 用户偏好：高风险操作是否需要确认 */
  confirmHighRisk: boolean;
  /** 用户偏好：中风险操作是否需要确认 */
  confirmMediumRisk: boolean;
}

/**
 * 默认配置
 */
export const DEFAULT_APPROVAL_CONFIG: ApprovalManagerConfig = {
  approvalTimeout: 5 * 60 * 1000, // 5分钟
  batchThreshold: 200, // 超过200个单元格需要确认
  enableAudit: true,
  confirmHighRisk: true,
  confirmMediumRisk: false,
};

// ==================== ApprovalManager 类 ====================

/**
 * 审批管理器
 *
 * Agent 层的核心组件，负责判定操作风险并管理审批流程
 */
export class ApprovalManager {
  private config: ApprovalManagerConfig;
  private pendingApprovals: Map<string, ApprovalRequest> = new Map();
  private approvalHistory: ApprovalDecision[] = [];
  private auditLogger: AuditLogger;
  private approvalCounter: number = 0;

  constructor(config: Partial<ApprovalManagerConfig> = {}) {
    this.config = { ...DEFAULT_APPROVAL_CONFIG, ...config };
    this.auditLogger = new AuditLogger();
  }

  /**
   * 生成审批 ID
   * 格式: APP-YYYYMMDD-NNN
   */
  generateApprovalId(): string {
    const now = new Date();
    const dateStr = now.toISOString().slice(0, 10).replace(/-/g, "");
    this.approvalCounter++;
    const seq = String(this.approvalCounter).padStart(3, "0");
    return `APP-${dateStr}-${seq}`;
  }

  /**
   * 评估操作风险
   *
   * 这是 Agent 层的核心判定逻辑：
   * - 基于操作类型判定基础风险
   * - 基于参数动态调整风险等级
   */
  assessRisk(
    operationName: string,
    parameters: Record<string, unknown>,
    context?: { userInput?: string; estimatedRows?: number }
  ): RiskAssessment {
    let riskLevel: RiskLevel = "low";
    let needsApproval = false;
    let reason = "";
    let impactDescription = "";
    let reversible = true;

    const userInput = context?.userInput || "";
    const estimatedRows = context?.estimatedRows || 0;

    // 1. 检查是否是高风险操作
    if (HIGH_RISK_OPERATIONS.includes(operationName as any)) {
      riskLevel = "high";
      needsApproval = this.config.confirmHighRisk;
      reversible = false;
      reason = `操作 "${operationName}" 属于高风险操作`;

      // 具体描述影响
      switch (operationName) {
        case "delete_rows":
        case "delete_row":
          impactDescription = `将删除指定行，此操作不可撤销`;
          break;
        case "delete_columns":
        case "delete_column":
          impactDescription = `将删除指定列，此操作不可撤销`;
          break;
        case "delete_sheet":
          impactDescription = `将删除整个工作表及其所有数据，此操作不可撤销`;
          riskLevel = "critical";
          break;
        case "clear_range":
          impactDescription = `将清空指定区域的所有内容`;
          reversible = true;
          break;
        case "remove_duplicates":
          impactDescription = `将删除重复行，被删除的数据无法恢复`;
          break;
        case "protect_sheet":
        case "unprotect_sheet":
          impactDescription = `将改变工作表的保护状态`;
          reversible = true;
          break;
        default:
          impactDescription = `高风险操作，请确认后执行`;
      }
    }

    // 2. 检查是否是中风险操作
    else if (MEDIUM_RISK_OPERATIONS.includes(operationName as any)) {
      riskLevel = "medium";
      needsApproval = this.config.confirmMediumRisk;
      reason = `操作 "${operationName}" 属于中风险操作`;
      impactDescription = `可能覆盖现有数据`;
    }

    // 3. 检查批量操作关键词
    const hasBatchKeyword = BATCH_KEYWORDS.some(
      (kw) => userInput.includes(kw) || JSON.stringify(parameters).includes(kw)
    );
    if (hasBatchKeyword) {
      if (riskLevel === "low") riskLevel = "medium";
      if (riskLevel === "medium") riskLevel = "high";
      needsApproval = true;
      reason += (reason ? "；" : "") + "检测到批量操作关键词";
    }

    // 4. 检查影响范围（超过阈值需要确认）
    if (estimatedRows > this.config.batchThreshold) {
      if (riskLevel === "low") riskLevel = "medium";
      needsApproval = true;
      reason +=
        (reason ? "；" : "") + `影响行数(${estimatedRows})超过阈值(${this.config.batchThreshold})`;
    }

    // 5. 检查参数中的范围大小
    const range = (parameters.range as string) || (parameters.address as string) || "";
    if (this.isLargeRange(range)) {
      if (riskLevel === "low") riskLevel = "medium";
      needsApproval = this.config.confirmMediumRisk || riskLevel === "high";
      reason += (reason ? "；" : "") + "操作范围较大";
    }

    // 6. 检查特殊参数
    if (parameters.scope === "all" || parameters.applyToAll === true) {
      riskLevel = "high";
      needsApproval = true;
      reason += (reason ? "；" : "") + "操作将应用到全部数据";
    }

    return {
      needsApproval,
      riskLevel,
      reason: reason || "常规操作",
      impactDescription: impactDescription || "标准操作",
      reversible,
      estimatedImpact: {
        rowCount: estimatedRows,
      },
    };
  }

  /**
   * 判断是否是大范围
   */
  private isLargeRange(range: string): boolean {
    if (!range) return false;

    // 检查整列/整行标记
    if (/^\d+:\d+$/.test(range) || /^[A-Z]+:[A-Z]+$/.test(range)) {
      return true;
    }

    // 解析范围大小
    const match = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/i);
    if (match) {
      const startCol = this.columnToNumber(match[1]);
      const startRow = parseInt(match[2]);
      const endCol = this.columnToNumber(match[3]);
      const endRow = parseInt(match[4]);

      const cellCount = (endCol - startCol + 1) * (endRow - startRow + 1);
      return cellCount > this.config.batchThreshold;
    }

    return false;
  }

  /**
   * 列字母转数字
   */
  private columnToNumber(col: string): number {
    let result = 0;
    for (let i = 0; i < col.length; i++) {
      result = result * 26 + (col.charCodeAt(i) - 64);
    }
    return result;
  }

  /**
   * 创建审批请求
   */
  createApprovalRequest(
    operationName: string,
    operationType: string,
    parameters: Record<string, unknown>,
    riskAssessment: RiskAssessment,
    options?: { userId?: string; sessionId?: string }
  ): ApprovalRequest {
    const approvalId = this.generateApprovalId();
    const now = new Date();

    const request: ApprovalRequest = {
      approvalId,
      operationName,
      operationType,
      parameters,
      riskAssessment,
      requestTime: now,
      expiresAt: new Date(now.getTime() + this.config.approvalTimeout),
      status: "pending",
      userId: options?.userId,
      sessionId: options?.sessionId,
    };

    this.pendingApprovals.set(approvalId, request);

    // 记录审计日志
    if (this.config.enableAudit) {
      this.auditLogger.log({
        action: "approval_requested",
        approvalId,
        operationName,
        operationType,
        parameters,
        riskLevel: riskAssessment.riskLevel,
        userId: options?.userId,
        sessionId: options?.sessionId,
      });
    }

    return request;
  }

  /**
   * 处理用户审批决定
   *
   * @param approvalId 审批 ID
   * @param approved 是否批准
   * @param decidedBy 决定者
   * @param reason 原因
   */
  handleApprovalDecision(
    approvalId: string,
    approved: boolean,
    decidedBy?: string,
    reason?: string
  ): { success: boolean; request?: ApprovalRequest; error?: string } {
    const request = this.pendingApprovals.get(approvalId);

    if (!request) {
      return { success: false, error: `未找到审批请求: ${approvalId}` };
    }

    // 检查是否过期
    if (new Date() > request.expiresAt) {
      request.status = "expired";
      this.pendingApprovals.delete(approvalId);
      return { success: false, error: `审批请求已过期: ${approvalId}` };
    }

    // 更新状态
    request.status = approved ? "approved" : "rejected";
    this.pendingApprovals.delete(approvalId);

    // 记录决定
    const decision: ApprovalDecision = {
      approvalId,
      approved,
      decidedAt: new Date(),
      decidedBy,
      reason,
    };
    this.approvalHistory.push(decision);

    // 记录审计日志
    if (this.config.enableAudit) {
      this.auditLogger.log({
        action: approved ? "approval_granted" : "approval_rejected",
        approvalId,
        operationName: request.operationName,
        operationType: request.operationType,
        parameters: request.parameters,
        riskLevel: request.riskAssessment.riskLevel,
        decidedBy,
        reason,
        userId: request.userId,
        sessionId: request.sessionId,
      });
    }

    return { success: true, request };
  }

  /**
   * 验证用户确认文本
   *
   * 防止误触/注入：要求用户回复精确的短语
   * 格式: "确认执行 APP-XXXXXXXX-XXX"
   */
  validateConfirmationText(text: string, approvalId: string): boolean {
    const expectedText = `确认执行 ${approvalId}`;
    return text.trim() === expectedText;
  }

  /**
   * 获取待审批请求
   */
  getPendingApproval(approvalId: string): ApprovalRequest | undefined {
    return this.pendingApprovals.get(approvalId);
  }

  /**
   * 获取所有待审批请求
   */
  getAllPendingApprovals(): ApprovalRequest[] {
    return Array.from(this.pendingApprovals.values());
  }

  /**
   * 清理过期的审批请求
   */
  cleanupExpiredApprovals(): number {
    const now = new Date();
    let cleanedCount = 0;

    for (const [id, request] of this.pendingApprovals) {
      if (now > request.expiresAt) {
        request.status = "expired";
        this.pendingApprovals.delete(id);
        cleanedCount++;

        if (this.config.enableAudit) {
          this.auditLogger.log({
            action: "approval_expired",
            approvalId: id,
            operationName: request.operationName,
            operationType: request.operationType,
          });
        }
      }
    }

    return cleanedCount;
  }

  /**
   * 生成确认弹窗文案
   */
  generateConfirmationMessage(request: ApprovalRequest): string {
    const { operationName, parameters, riskAssessment, approvalId } = request;

    const lines = [
      `【${riskAssessment.riskLevel === "critical" ? "严重" : "高"}风险操作待确认｜${approvalId}】`,
      "",
      `📌 将执行：${this.getOperationDisplayName(operationName)}`,
      `📊 影响范围：${parameters.range || parameters.address || "当前选区"}`,
    ];

    if (riskAssessment.estimatedImpact?.rowCount) {
      lines.push(`📈 影响行数：约 ${riskAssessment.estimatedImpact.rowCount} 行`);
    }

    lines.push(`⚠️ 风险说明：${riskAssessment.impactDescription}`);
    lines.push(`🔄 可撤销：${riskAssessment.reversible ? "是" : "否"}`);
    lines.push("");
    lines.push(`请点击「确认执行」继续，或「取消」放弃操作。`);

    return lines.join("\n");
  }

  /**
   * 获取操作显示名称
   */
  private getOperationDisplayName(operationName: string): string {
    const displayNames: Record<string, string> = {
      delete_rows: "删除行",
      delete_row: "删除行",
      delete_columns: "删除列",
      delete_column: "删除列",
      delete_sheet: "删除工作表",
      clear_range: "清空区域",
      clear_all: "清空全部",
      batch_update: "批量更新",
      batch_write: "批量写入",
      batch_formula: "批量公式",
      fill_formula: "填充公式",
      remove_duplicates: "删除重复项",
      protect_sheet: "保护工作表",
      unprotect_sheet: "取消保护",
      write_range: "写入数据",
      set_formula: "设置公式",
      sort_range: "排序数据",
    };
    return displayNames[operationName] || operationName;
  }

  /**
   * 获取审计日志
   */
  getAuditLogs(): AuditEntry[] {
    return this.auditLogger.getLogs();
  }

  /**
   * 更新配置
   */
  updateConfig(newConfig: Partial<ApprovalManagerConfig>): void {
    this.config = { ...this.config, ...newConfig };
  }
}

// ==================== 导出单例 ====================

export const approvalManager = new ApprovalManager();

export default ApprovalManager;
