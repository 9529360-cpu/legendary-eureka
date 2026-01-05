/**
 * ClarificationEngine - 澄清引擎 v2.9.58
 *
 * P2 核心组件：处理澄清交互的完整流程
 *
 * 核心职责：
 * 1. 根据 IntentAnalysis 生成用户友好的澄清消息
 * 2. 解析用户对澄清问题的回复
 * 3. 更新任务上下文
 * 4. 支持多轮澄清
 *
 * 设计理念：
 * - 澄清消息要像人在问，不像机器在报错
 * - 提供选项降低用户认知负担
 * - 支持模糊回复的智能理解
 */

import {
  IntentAnalysis,
  SuggestedClarification,
  ClarificationItem,
  IntentAnalyzer,
  intentAnalyzer,
  AnalysisContext,
} from "./IntentAnalyzer";
import { DataModel } from "./DataModeler";

// ========== 类型定义 ==========

/**
 * 澄清会话状态
 */
export interface ClarificationSession {
  /** 会话 ID */
  sessionId: string;
  /** 原始用户请求 */
  originalRequest: string;
  /** 意图分析结果 */
  intentAnalysis: IntentAnalysis;
  /** 澄清历史 */
  history: ClarificationTurn[];
  /** 已收集的信息 */
  collectedInfo: CollectedInfo;
  /** 会话状态 */
  status: "pending" | "resolved" | "abandoned";
  /** 创建时间 */
  createdAt: Date;
  /** 最后更新时间 */
  updatedAt: Date;
}

/**
 * 一轮澄清交互
 */
export interface ClarificationTurn {
  /** 轮次 */
  turn: number;
  /** Agent 提出的问题 */
  question: SuggestedClarification;
  /** 用户的回复 */
  userResponse?: string;
  /** 解析后的信息 */
  parsedInfo?: Partial<CollectedInfo>;
  /** 时间戳 */
  timestamp: Date;
}

/**
 * 收集到的信息
 */
export interface CollectedInfo {
  /** 目标工作表 */
  targetSheet?: string;
  /** 目标范围 */
  targetRange?: string;
  /** 目标列 */
  targetColumns?: string[];
  /** 操作类型确认 */
  confirmedIntent?: string;
  /** 用户确认（对于风险操作） */
  userConfirmation?: boolean;
  /** 选择的方案 ID */
  selectedPlanId?: string;
  /** 其他自由文本补充 */
  additionalInfo?: string;
}

/**
 * 澄清结果
 */
export interface ClarificationResult {
  /** 是否解决（可以继续执行） */
  resolved: boolean;
  /** 如果未解决，下一个问题 */
  nextQuestion?: SuggestedClarification;
  /** 如果已解决，更新后的请求 */
  enhancedRequest?: string;
  /** 收集到的所有信息 */
  collectedInfo: CollectedInfo;
  /** 给用户的消息 */
  message: string;
  /** 消息类型 */
  messageType: "question" | "confirmation" | "info" | "ready";
}

/**
 * 用户回复解析结果
 */
interface ParsedUserResponse {
  /** 选择的选项 ID */
  selectedOptionId?: string;
  /** 自由文本内容 */
  freeformText?: string;
  /** 是否确认（yes/no） */
  isConfirmation?: boolean;
  /** 提取的实体 */
  extractedEntities?: {
    sheets?: string[];
    ranges?: string[];
    columns?: string[];
  };
}

// ========== ClarificationEngine 类 ==========

/**
 * 澄清引擎
 */
export class ClarificationEngine {
  private sessions: Map<string, ClarificationSession> = new Map();
  private analyzer: IntentAnalyzer;

  constructor(analyzer?: IntentAnalyzer) {
    this.analyzer = analyzer || intentAnalyzer;
  }

  /**
   * 开始澄清会话
   *
   * @param originalRequest - 用户原始请求
   * @param intentAnalysis - 意图分析结果
   * @returns 澄清结果（包含第一个问题或直接可执行）
   */
  startSession(originalRequest: string, intentAnalysis: IntentAnalysis): ClarificationResult {
    // 如果可以直接执行，不需要澄清
    if (intentAnalysis.canProceed) {
      return {
        resolved: true,
        enhancedRequest: originalRequest,
        collectedInfo: {},
        message: "",
        messageType: "ready",
      };
    }

    // 创建会话
    const sessionId = this.generateSessionId();
    const session: ClarificationSession = {
      sessionId,
      originalRequest,
      intentAnalysis,
      history: [],
      collectedInfo: {},
      status: "pending",
      createdAt: new Date(),
      updatedAt: new Date(),
    };

    // 生成第一个问题
    const firstQuestion = intentAnalysis.suggestedClarification;
    if (!firstQuestion) {
      // 兜底：生成通用问题
      return {
        resolved: false,
        nextQuestion: {
          mainQuestion: "请提供更多细节，帮助我理解您的需求。",
          allowFreeform: true,
        },
        collectedInfo: {},
        message: this.formatClarificationMessage({
          mainQuestion: "请提供更多细节，帮助我理解您的需求。",
          allowFreeform: true,
        }),
        messageType: "question",
      };
    }

    // 记录第一轮
    session.history.push({
      turn: 1,
      question: firstQuestion,
      timestamp: new Date(),
    });

    this.sessions.set(sessionId, session);

    return {
      resolved: false,
      nextQuestion: firstQuestion,
      collectedInfo: {},
      message: this.formatClarificationMessage(firstQuestion),
      messageType: "question",
    };
  }

  /**
   * 处理用户回复
   *
   * @param sessionId - 会话 ID
   * @param userResponse - 用户回复
   * @param context - 分析上下文（可选，用于重新分析）
   */
  handleResponse(
    sessionId: string,
    userResponse: string,
    _context?: AnalysisContext
  ): ClarificationResult {
    const session = this.sessions.get(sessionId);

    if (!session) {
      // 会话不存在，作为新请求处理
      return {
        resolved: false,
        collectedInfo: {},
        message: "抱歉，会话已过期。请重新描述您的需求。",
        messageType: "info",
      };
    }

    // 解析用户回复
    const lastTurn = session.history[session.history.length - 1];
    const parsed = this.parseUserResponse(userResponse, lastTurn?.question);

    // 更新最后一轮的用户回复
    if (lastTurn) {
      lastTurn.userResponse = userResponse;
      lastTurn.parsedInfo = this.extractInfoFromParsed(parsed);
    }

    // 合并收集的信息
    this.mergeCollectedInfo(session.collectedInfo, lastTurn?.parsedInfo);

    session.updatedAt = new Date();

    // 检查是否需要继续澄清
    const remainingNeeds = this.checkRemainingNeeds(session);

    if (remainingNeeds.length === 0) {
      // 澄清完成
      session.status = "resolved";
      const enhancedRequest = this.buildEnhancedRequest(session);

      return {
        resolved: true,
        enhancedRequest,
        collectedInfo: session.collectedInfo,
        message: "好的，我明白了。",
        messageType: "ready",
      };
    }

    // 还需要继续澄清
    const nextQuestion = this.generateNextQuestion(remainingNeeds, session);

    session.history.push({
      turn: session.history.length + 1,
      question: nextQuestion,
      timestamp: new Date(),
    });

    return {
      resolved: false,
      nextQuestion,
      collectedInfo: session.collectedInfo,
      message: this.formatClarificationMessage(nextQuestion),
      messageType: "question",
    };
  }

  /**
   * 获取会话状态
   */
  getSession(sessionId: string): ClarificationSession | undefined {
    return this.sessions.get(sessionId);
  }

  /**
   * 放弃会话
   */
  abandonSession(sessionId: string): void {
    const session = this.sessions.get(sessionId);
    if (session) {
      session.status = "abandoned";
    }
  }

  /**
   * 格式化澄清消息（用户友好）
   */
  formatClarificationMessage(clarification: SuggestedClarification): string {
    const parts: string[] = [];

    // 主问题
    parts.push(clarification.mainQuestion);

    // 上下文说明
    if (clarification.context) {
      parts.push(`\n_${clarification.context}_`);
    }

    // 选项
    if (clarification.options && clarification.options.length > 0) {
      parts.push("\n");
      for (let i = 0; i < clarification.options.length; i++) {
        const opt = clarification.options[i];
        const prefix = opt.recommended ? "👉 " : "• ";
        parts.push(`${prefix}**${opt.label}**${opt.description ? ` - ${opt.description}` : ""}`);
      }
    }

    // 如果允许自由回答
    if (clarification.allowFreeform && clarification.options?.length) {
      parts.push("\n_您也可以直接告诉我具体要求_");
    }

    return parts.join("\n");
  }

  /**
   * 快速澄清检查（不创建会话）
   *
   * 用于快速判断请求是否需要澄清
   */
  quickCheck(
    request: string,
    dataModel?: DataModel,
    currentSelection?: string,
    activeSheet?: string,
    clarificationThreshold: number = 0.7
  ): {
    needsClarification: boolean;
    confidence: number;
    reason?: string;
    suggestedQuestion?: string;
  } {
    const analysis = this.analyzer.analyze({
      userRequest: request,
      dataModel,
      currentSelection,
      activeSheet,
      clarificationThreshold,
    });

    if (analysis.canProceed) {
      return {
        needsClarification: false,
        confidence: analysis.confidence,
      };
    }

    return {
      needsClarification: true,
      confidence: analysis.confidence,
      reason: analysis.clarificationNeeded[0]?.reason,
      suggestedQuestion: analysis.suggestedClarification?.mainQuestion,
    };
  }

  // ========== 私有方法 ==========

  private generateSessionId(): string {
    return `clarify_${Date.now()}_${Math.random().toString(36).substring(2, 8)}`;
  }

  private parseUserResponse(
    response: string,
    question?: SuggestedClarification
  ): ParsedUserResponse {
    const result: ParsedUserResponse = {};

    // 检查是否匹配选项
    if (question?.options) {
      const lowerResponse = response.toLowerCase().trim();

      for (const opt of question.options) {
        // 匹配选项 ID 或标签
        if (
          lowerResponse === opt.id.toLowerCase() ||
          lowerResponse === opt.label.toLowerCase() ||
          response.includes(opt.label)
        ) {
          result.selectedOptionId = opt.id;
          break;
        }
      }

      // 数字选择
      const numMatch = response.match(/^(\d+)$/);
      if (numMatch) {
        const index = parseInt(numMatch[1]) - 1;
        if (index >= 0 && index < question.options.length) {
          result.selectedOptionId = question.options[index].id;
        }
      }
    }

    // 检查确认意图
    const yesPatterns = /^(是|对|好|确认|ok|yes|确定|同意|行|可以|嗯)/i;
    const noPatterns = /^(否|不|取消|no|算了|不要|别)/i;

    if (yesPatterns.test(response.trim())) {
      result.isConfirmation = true;
    } else if (noPatterns.test(response.trim())) {
      result.isConfirmation = false;
    }

    // 提取实体
    result.extractedEntities = {
      sheets: this.extractSheetNames(response),
      ranges: this.extractRanges(response),
      columns: this.extractColumnNames(response),
    };

    // 自由文本
    result.freeformText = response;

    return result;
  }

  private extractSheetNames(text: string): string[] {
    const patterns = [
      /'([^']+)'/g, // 'Sheet Name'
      /"([^"]+)"/g, // "Sheet Name"
      /(?:工作表|表)\s*[""']?([^""'\s,，]+)[""']?/gi,
    ];

    const names: string[] = [];
    for (const pattern of patterns) {
      let match;
      while ((match = pattern.exec(text)) !== null) {
        if (match[1]) names.push(match[1]);
      }
    }
    return names;
  }

  private extractRanges(text: string): string[] {
    const patterns = [/([A-Z]+\d+:[A-Z]+\d+)/gi, /([A-Z]+\d+)/gi, /([A-Z]+)列/gi];

    const ranges: string[] = [];
    for (const pattern of patterns) {
      let match;
      while ((match = pattern.exec(text)) !== null) {
        if (match[1]) ranges.push(match[1]);
      }
    }
    return ranges;
  }

  private extractColumnNames(text: string): string[] {
    // 这里简化处理，实际应该结合 dataModel
    const patterns = [/([A-Z]+)列/gi, /(\S+?)(?:列|字段|栏)/gi];

    const columns: string[] = [];
    for (const pattern of patterns) {
      let match;
      while ((match = pattern.exec(text)) !== null) {
        if (match[1]) columns.push(match[1]);
      }
    }
    return columns;
  }

  private extractInfoFromParsed(parsed: ParsedUserResponse): Partial<CollectedInfo> {
    const info: Partial<CollectedInfo> = {};

    if (parsed.selectedOptionId) {
      // 根据选项类型分类
      if (parsed.selectedOptionId.startsWith("sheet_")) {
        info.targetSheet = parsed.selectedOptionId.replace("sheet_", "");
      } else if (parsed.selectedOptionId === "confirm") {
        info.userConfirmation = true;
      } else if (parsed.selectedOptionId === "cancel") {
        info.userConfirmation = false;
      } else {
        info.selectedPlanId = parsed.selectedOptionId;
      }
    }

    if (parsed.isConfirmation !== undefined) {
      info.userConfirmation = parsed.isConfirmation;
    }

    if (parsed.extractedEntities) {
      if (parsed.extractedEntities.sheets?.length) {
        info.targetSheet = parsed.extractedEntities.sheets[0];
      }
      if (parsed.extractedEntities.ranges?.length) {
        info.targetRange = parsed.extractedEntities.ranges[0];
      }
      if (parsed.extractedEntities.columns?.length) {
        info.targetColumns = parsed.extractedEntities.columns;
      }
    }

    if (parsed.freeformText) {
      info.additionalInfo = parsed.freeformText;
    }

    return info;
  }

  private mergeCollectedInfo(target: CollectedInfo, source?: Partial<CollectedInfo>): void {
    if (!source) return;

    if (source.targetSheet) target.targetSheet = source.targetSheet;
    if (source.targetRange) target.targetRange = source.targetRange;
    if (source.targetColumns) target.targetColumns = source.targetColumns;
    if (source.confirmedIntent) target.confirmedIntent = source.confirmedIntent;
    if (source.userConfirmation !== undefined) {
      target.userConfirmation = source.userConfirmation;
    }
    if (source.selectedPlanId) target.selectedPlanId = source.selectedPlanId;
    if (source.additionalInfo) {
      target.additionalInfo = target.additionalInfo
        ? `${target.additionalInfo}; ${source.additionalInfo}`
        : source.additionalInfo;
    }
  }

  private checkRemainingNeeds(session: ClarificationSession): ClarificationItem[] {
    const { intentAnalysis, collectedInfo } = session;

    return intentAnalysis.clarificationNeeded.filter((need) => {
      switch (need.type) {
        case "missing_sheet":
          return !collectedInfo.targetSheet;
        case "missing_range":
          return !collectedInfo.targetRange && !collectedInfo.targetColumns;
        case "ambiguous_intent":
          return !collectedInfo.confirmedIntent;
        case "risky_operation":
          return collectedInfo.userConfirmation === undefined;
        case "vague_reference":
          return !collectedInfo.targetRange;
        default:
          return false;
      }
    });
  }

  private generateNextQuestion(
    remainingNeeds: ClarificationItem[],
    _session: ClarificationSession
  ): SuggestedClarification {
    const primaryNeed = remainingNeeds[0];

    if (!primaryNeed) {
      return {
        mainQuestion: "还有什么需要补充的吗？",
        allowFreeform: true,
      };
    }

    // 复用 IntentAnalyzer 的问题生成逻辑
    // 这里简化处理
    return {
      mainQuestion: `请告诉我${primaryNeed.missing}`,
      context: primaryNeed.reason,
      options: primaryNeed.options?.map((opt, i) => ({
        id: `opt_${i}`,
        label: opt,
      })),
      allowFreeform: true,
    };
  }

  private buildEnhancedRequest(session: ClarificationSession): string {
    const { originalRequest, collectedInfo } = session;
    const parts = [originalRequest];

    if (collectedInfo.targetSheet) {
      parts.push(`工作表: ${collectedInfo.targetSheet}`);
    }
    if (collectedInfo.targetRange) {
      parts.push(`范围: ${collectedInfo.targetRange}`);
    }
    if (collectedInfo.targetColumns?.length) {
      parts.push(`列: ${collectedInfo.targetColumns.join(", ")}`);
    }
    if (collectedInfo.additionalInfo) {
      parts.push(collectedInfo.additionalInfo);
    }

    return parts.join(" | ");
  }
}

// ========== 单例导出 ==========

export const clarificationEngine = new ClarificationEngine();

export default ClarificationEngine;
