/**
 * ConversationMemory - 多轮对话上下文记忆管理
 *
 * 功能：
 * - 多轮对话上下文管理
 * - 意图理解与追踪
 * - 上下文压缩与摘要
 * - 引用历史能力
 * - 会话持久化
 *
 * @version 1.0.0
 */

import { Logger } from "../utils/Logger";

/**
 * 消息类型
 */
export type MessageRole = "user" | "assistant" | "system" | "tool";

/**
 * 对话消息
 */
export interface ConversationMessage {
  /** 消息唯一ID */
  id: string;
  /** 角色 */
  role: MessageRole;
  /** 内容 */
  content: string;
  /** 时间戳 */
  timestamp: number;
  /** 元数据 */
  metadata?: MessageMetadata;
  /** 关联的工具调用 */
  toolCalls?: ToolCallRecord[];
  /** 是否已压缩 */
  compressed?: boolean;
}

/**
 * 消息元数据
 */
export interface MessageMetadata {
  /** 意图分类 */
  intent?: string;
  /** 意图置信度 */
  intentConfidence?: number;
  /** 提取的实体 */
  entities?: ExtractedEntity[];
  /** 情感分析 */
  sentiment?: "positive" | "neutral" | "negative";
  /** Token 数量估算 */
  tokenCount?: number;
  /** 是否为关键消息（不应被压缩） */
  isKeyMessage?: boolean;
}

/**
 * 提取的实体
 */
export interface ExtractedEntity {
  type: string;
  value: string;
  position?: [number, number];
}

/**
 * 工具调用记录
 */
export interface ToolCallRecord {
  toolName: string;
  parameters: Record<string, unknown>;
  result?: unknown;
  success: boolean;
  duration?: number;
}

/**
 * 对话上下文
 */
export interface ConversationContext {
  /** 会话ID */
  sessionId: string;
  /** 会话标题 */
  title?: string;
  /** 创建时间 */
  createdAt: number;
  /** 最后活动时间 */
  lastActivityAt: number;
  /** 消息列表 */
  messages: ConversationMessage[];
  /** 会话摘要 */
  summary?: string;
  /** 会话状态 */
  status: "active" | "paused" | "completed";
  /** 用户偏好（从对话中学习） */
  userPreferences?: UserPreferences;
  /** 当前任务上下文 */
  currentTask?: TaskContext;
  /** v2.9.45: 对话跟踪器 */
  tracker?: ConversationTracker;
  /** v2.9.45: 工作簿快照 */
  workbookSnapshot?: WorkbookSnapshot;
  /** v2.9.45: 操作历史 */
  operationHistory?: OperationRecord[];
}

/**
 * v2.9.45: 工作簿快照（用于上下文理解）
 */
export interface WorkbookSnapshot {
  activeSheet: string;
  sheets: string[];
  selectedRange?: string;
  lastModifiedCell?: string;
  dataRanges?: Array<{ name: string; address: string; type: string }>;
  timestamp: number;
}

/**
 * v2.9.45: 操作记录
 */
export interface OperationRecord {
  id: string;
  type: string;
  target: string;
  params: Record<string, unknown>;
  success: boolean;
  timestamp: number;
  duration?: number;
  undoable: boolean;
}

/**
 * 用户偏好
 */
export interface UserPreferences {
  /** 偏好的格式风格 */
  preferredFormatStyle?: string;
  /** 常用颜色 */
  commonColors?: string[];
  /** 常用字体 */
  preferredFont?: string;
  /** 常用单位格式 */
  numberFormat?: string;
  /** 语言偏好 */
  language?: string;
}

/**
 * 任务上下文
 */
export interface TaskContext {
  /** 任务类型 */
  type: string;
  /** 任务目标 */
  goal: string;
  /** 当前步骤 */
  currentStep: number;
  /** 总步骤数 */
  totalSteps: number;
  /** 中间结果 */
  intermediateResults?: unknown[];
  /** v2.9.45: 任务优先级 */
  priority?: "high" | "medium" | "low";
  /** v2.9.45: 关联的数据范围 */
  targetRanges?: string[];
  /** v2.9.45: 任务创建时间 */
  createdAt?: number;
  /** v2.9.45: 预计完成时间 */
  estimatedDuration?: number;
  /** v2.9.45: 任务依赖 */
  dependencies?: string[];
  /** v2.9.45: 子任务列表 */
  subTasks?: SubTask[];
}

/**
 * v2.9.45: 子任务定义
 */
export interface SubTask {
  id: string;
  name: string;
  status: "pending" | "in-progress" | "completed" | "failed";
  result?: unknown;
  error?: string;
}

/**
 * v2.9.45: 多轮对话跟踪
 */
export interface ConversationTracker {
  /** 当前主题 */
  currentTopic?: string;
  /** 话题变更历史 */
  topicHistory: Array<{ topic: string; timestamp: number }>;
  /** 未解决的问题 */
  pendingQuestions: string[];
  /** 用户确认等待 */
  awaitingConfirmation?: {
    action: string;
    options: string[];
    timeout: number;
  };
  /** 上下文实体缓存 */
  contextEntities: Map<string, unknown>;
  /** 引用消息 */
  referencedMessages: string[];
}

/**
 * 引用信息
 */
export interface Reference {
  /** 引用的消息ID */
  messageId: string;
  /** 引用的内容片段 */
  snippet: string;
  /** 相关度分数 */
  relevanceScore: number;
}

/**
 * 上下文窗口配置
 */
export interface ContextWindowConfig {
  /** 最大消息数 */
  maxMessages: number;
  /** 最大 Token 数 */
  maxTokens: number;
  /** 保留最近N条消息（不压缩） */
  preserveRecentCount: number;
  /** 压缩阈值 */
  compressionThreshold: number;
}

// ============ 意图分类 ============

/**
 * 意图类型
 */
export enum IntentType {
  // 数据操作
  READ_DATA = "read_data",
  WRITE_DATA = "write_data",
  FORMAT_DATA = "format_data",
  DELETE_DATA = "delete_data",

  // 分析类
  ANALYZE_DATA = "analyze_data",
  CREATE_CHART = "create_chart",
  CALCULATE = "calculate",

  // v2.9.45: 新增高级分析意图
  TREND_ANALYSIS = "trend_analysis",
  ANOMALY_DETECTION = "anomaly_detection",
  PREDICTIVE_ANALYSIS = "predictive_analysis",
  STATISTICAL_ANALYSIS = "statistical_analysis",
  DATA_INSIGHTS = "data_insights",

  // 格式化
  BEAUTIFY_TABLE = "beautify_table",
  APPLY_STYLE = "apply_style",
  CONDITIONAL_FORMAT = "conditional_format",

  // 查询类
  QUERY_INFO = "query_info",
  EXPLAIN = "explain",

  // 其他
  UNDO = "undo",
  REDO = "redo",
  HELP = "help",
  UNKNOWN = "unknown",

  // v2.9.45: 新增
  FOLLOW_UP = "follow_up", // 续接对话
  CLARIFICATION = "clarification", // 澄清请求
  CONFIRMATION = "confirmation", // 确认操作
}

/**
 * 意图分析结果
 */
export interface IntentAnalysis {
  primaryIntent: IntentType;
  confidence: number;
  secondaryIntents?: IntentType[];
  entities: ExtractedEntity[];
  suggestedActions?: string[];
}

// ============ ConversationMemory 类 ============

class ConversationMemoryImpl {
  private sessions: Map<string, ConversationContext> = new Map();
  private currentSessionId: string | null = null;
  private config: ContextWindowConfig = {
    maxMessages: 50,
    maxTokens: 8000,
    preserveRecentCount: 10,
    compressionThreshold: 0.7,
  };

  /**
   * 创建新会话
   */
  createSession(title?: string): ConversationContext {
    const sessionId = this.generateSessionId();
    const session: ConversationContext = {
      sessionId,
      title: title || `会话 ${new Date().toLocaleString()}`,
      createdAt: Date.now(),
      lastActivityAt: Date.now(),
      messages: [],
      status: "active",
      tracker: {
        topicHistory: [],
        pendingQuestions: [],
        contextEntities: new Map(),
        referencedMessages: [],
      },
      operationHistory: [],
    };

    this.sessions.set(sessionId, session);
    this.currentSessionId = sessionId;

    Logger.info("[ConversationMemory] 创建新会话", { sessionId });

    return session;
  }

  /**
   * 获取当前会话
   */
  getCurrentSession(): ConversationContext | null {
    if (!this.currentSessionId) {
      return null;
    }
    return this.sessions.get(this.currentSessionId) || null;
  }

  /**
   * v2.9.45: 更新工作簿快照
   */
  updateWorkbookSnapshot(snapshot: Partial<WorkbookSnapshot>): void {
    const session = this.getCurrentSession();
    if (!session) return;

    session.workbookSnapshot = {
      ...session.workbookSnapshot,
      ...snapshot,
      timestamp: Date.now(),
    } as WorkbookSnapshot;
  }

  /**
   * v2.9.45: 获取工作簿快照
   */
  getWorkbookSnapshot(): WorkbookSnapshot | undefined {
    return this.getCurrentSession()?.workbookSnapshot;
  }

  /**
   * v2.9.45: 记录操作
   */
  recordOperation(operation: Omit<OperationRecord, "id" | "timestamp">): void {
    const session = this.getCurrentSession();
    if (!session) return;

    if (!session.operationHistory) {
      session.operationHistory = [];
    }

    session.operationHistory.push({
      ...operation,
      id: `op_${Date.now()}_${Math.random().toString(36).substring(2, 6)}`,
      timestamp: Date.now(),
    });

    // 只保留最近100条操作记录
    if (session.operationHistory.length > 100) {
      session.operationHistory = session.operationHistory.slice(-100);
    }
  }

  /**
   * v2.9.45: 获取最近的操作
   */
  getRecentOperations(count: number = 10): OperationRecord[] {
    const session = this.getCurrentSession();
    if (!session?.operationHistory) return [];
    return session.operationHistory.slice(-count);
  }

  /**
   * v2.9.45: 解析模糊引用（如"这里"、"那个表"、"刚才的范围"等）
   */
  resolveAmbiguousReference(reference: string): { type: string; value: unknown } | null {
    const session = this.getCurrentSession();
    if (!session) return null;

    const lowerRef = reference.toLowerCase();

    // 解析"这里"、"当前"等指向当前选区
    if (/这里|当前|现在|此处|here|current/.test(lowerRef)) {
      const snapshot = session.workbookSnapshot;
      if (snapshot?.selectedRange) {
        return { type: "range", value: snapshot.selectedRange };
      }
    }

    // 解析"刚才"、"之前"等指向最近操作
    if (/刚才|之前|上次|刚刚|just|previous|last/.test(lowerRef)) {
      const lastOp = session.operationHistory?.slice(-1)[0];
      if (lastOp) {
        return { type: "operation", value: lastOp };
      }
    }

    // 解析"那个表"、"那个工作表"等
    if (/那个表|那张表|那个工作表|that sheet/.test(lowerRef)) {
      const tracker = session.tracker;
      if (tracker?.contextEntities.has("lastSheet")) {
        return { type: "sheet", value: tracker.contextEntities.get("lastSheet") };
      }
    }

    // 解析"同样的"、"一样的"等
    if (/同样|一样|相同|same/.test(lowerRef)) {
      const lastOp = session.operationHistory?.slice(-1)[0];
      if (lastOp) {
        return { type: "repeat", value: lastOp };
      }
    }

    // 从对话历史中查找最近提到的范围
    const recentMessages = session.messages.slice(-5);
    for (const msg of recentMessages.reverse()) {
      const rangeMatch = msg.content.match(/([A-Z]+\d+(?::[A-Z]+\d+)?)/i);
      if (rangeMatch) {
        return { type: "range", value: rangeMatch[1].toUpperCase() };
      }
    }

    return null;
  }

  /**
   * v2.9.45: 更新话题
   */
  updateTopic(topic: string): void {
    const session = this.getCurrentSession();
    if (!session?.tracker) return;

    if (session.tracker.currentTopic !== topic) {
      session.tracker.topicHistory.push({
        topic: session.tracker.currentTopic || "开始",
        timestamp: Date.now(),
      });
      session.tracker.currentTopic = topic;
    }
  }

  /**
   * v2.9.45: 添加待确认操作
   */
  setAwaitingConfirmation(action: string, options: string[], timeoutMs: number = 60000): void {
    const session = this.getCurrentSession();
    if (!session?.tracker) return;

    session.tracker.awaitingConfirmation = {
      action,
      options,
      timeout: Date.now() + timeoutMs,
    };
  }

  /**
   * v2.9.45: 检查是否有待确认操作
   */
  getAwaitingConfirmation(): { action: string; options: string[] } | null {
    const session = this.getCurrentSession();
    if (!session?.tracker?.awaitingConfirmation) return null;

    const confirmation = session.tracker.awaitingConfirmation;
    if (Date.now() > confirmation.timeout) {
      session.tracker.awaitingConfirmation = undefined;
      return null;
    }

    return { action: confirmation.action, options: confirmation.options };
  }

  /**
   * v2.9.45: 清除待确认操作
   */
  clearAwaitingConfirmation(): void {
    const session = this.getCurrentSession();
    if (session?.tracker) {
      session.tracker.awaitingConfirmation = undefined;
    }
  }

  /**
   * v2.9.45: 缓存上下文实体
   */
  cacheContextEntity(key: string, value: unknown): void {
    const session = this.getCurrentSession();
    if (!session?.tracker) return;
    session.tracker.contextEntities.set(key, value);
  }

  /**
   * v2.9.45: 获取缓存的上下文实体
   */
  getContextEntity(key: string): unknown {
    const session = this.getCurrentSession();
    return session?.tracker?.contextEntities.get(key);
  }

  /**
   * v2.9.45: 智能理解用户输入，考虑上下文
   */
  understandWithContext(userInput: string): {
    resolvedInput: string;
    detectedEntities: Record<string, unknown>;
    suggestedAction?: string;
  } {
    const session = this.getCurrentSession();
    const detectedEntities: Record<string, unknown> = {};
    let resolvedInput = userInput;

    // 解析模糊引用
    const ambiguousPatterns = [
      { pattern: /这里|当前选区/g, type: "currentRange" },
      { pattern: /那个表|那张表/g, type: "lastSheet" },
      { pattern: /刚才的|之前的/g, type: "lastOperation" },
    ];

    for (const { pattern, type } of ambiguousPatterns) {
      if (pattern.test(userInput)) {
        const resolved = this.resolveAmbiguousReference(userInput);
        if (resolved) {
          detectedEntities[type] = resolved.value;
        }
      }
    }

    // 检测是否是确认回复
    const confirmation = this.getAwaitingConfirmation();
    if (confirmation) {
      const lowerInput = userInput.toLowerCase();
      if (/好|是|确认|ok|yes|确定|执行/.test(lowerInput)) {
        return {
          resolvedInput: userInput,
          detectedEntities: { confirmation: true, action: confirmation.action },
          suggestedAction: confirmation.action,
        };
      }
      if (/不|否|取消|no|cancel|算了/.test(lowerInput)) {
        this.clearAwaitingConfirmation();
        return {
          resolvedInput: userInput,
          detectedEntities: { confirmation: false },
          suggestedAction: "cancel",
        };
      }
    }

    // 检测续接对话
    if (/继续|接着|然后|还有|另外/.test(userInput) && session?.messages.length) {
      const lastAssistantMsg = [...session.messages].reverse().find((m) => m.role === "assistant");
      if (lastAssistantMsg) {
        detectedEntities.isFollowUp = true;
        detectedEntities.previousContext = lastAssistantMsg.content.substring(0, 200);
      }
    }

    return { resolvedInput, detectedEntities };
  }

  /**
   * 切换会话
   */
  switchSession(sessionId: string): boolean {
    if (!this.sessions.has(sessionId)) {
      Logger.warn("[ConversationMemory] 会话不存在", { sessionId });
      return false;
    }

    this.currentSessionId = sessionId;
    const session = this.sessions.get(sessionId)!;
    session.lastActivityAt = Date.now();
    session.status = "active";

    Logger.info("[ConversationMemory] 切换会话", { sessionId });
    return true;
  }

  /**
   * 添加消息
   */
  addMessage(
    role: MessageRole,
    content: string,
    metadata?: Partial<MessageMetadata>,
    toolCalls?: ToolCallRecord[]
  ): ConversationMessage {
    let session = this.getCurrentSession();
    if (!session) {
      session = this.createSession();
    }

    const message: ConversationMessage = {
      id: this.generateMessageId(),
      role,
      content,
      timestamp: Date.now(),
      metadata: {
        tokenCount: this.estimateTokenCount(content),
        ...metadata,
      },
      toolCalls,
    };

    // 如果是用户消息，分析意图
    if (role === "user") {
      const intentAnalysis = this.analyzeIntent(content, session);
      message.metadata!.intent = intentAnalysis.primaryIntent;
      message.metadata!.intentConfidence = intentAnalysis.confidence;
      message.metadata!.entities = intentAnalysis.entities;
    }

    session.messages.push(message);
    session.lastActivityAt = Date.now();

    // 检查是否需要压缩
    this.checkAndCompress(session);

    Logger.debug("[ConversationMemory] 添加消息", {
      sessionId: session.sessionId,
      role,
      messageId: message.id,
    });

    return message;
  }

  /**
   * 分析意图
   */
  analyzeIntent(content: string, context?: ConversationContext): IntentAnalysis {
    const lowerContent = content.toLowerCase();
    const entities: ExtractedEntity[] = [];
    let primaryIntent = IntentType.UNKNOWN;
    let confidence = 0.5;

    // 关键词匹配规则
    const intentPatterns: Array<{
      patterns: RegExp[];
      intent: IntentType;
      weight: number;
    }> = [
      // v2.9.45: 高级分析意图（优先匹配）
      {
        patterns: [/趋势|增长|下降|走势|变化趋势/, /trend|growth|decline/i],
        intent: IntentType.TREND_ANALYSIS,
        weight: 0.95,
      },
      {
        patterns: [/异常|离群|outlier|异常值|不正常/, /anomal|outlier|unusual/i],
        intent: IntentType.ANOMALY_DETECTION,
        weight: 0.95,
      },
      {
        patterns: [/预测|预估|未来|forecast|predict/, /forecast|predict|future/i],
        intent: IntentType.PREDICTIVE_ANALYSIS,
        weight: 0.95,
      },
      {
        patterns: [
          /统计|方差|标准差|相关性|分布/,
          /statistic|variance|std|correlation|distribution/i,
        ],
        intent: IntentType.STATISTICAL_ANALYSIS,
        weight: 0.9,
      },
      {
        patterns: [/洞察|发现|建议|问题|改进/, /insight|discover|suggest|improve/i],
        intent: IntentType.DATA_INSIGHTS,
        weight: 0.9,
      },
      // 格式化意图
      {
        patterns: [/美化|格式化|样式|风格|漂亮|好看/, /beautify|format|style/i],
        intent: IntentType.BEAUTIFY_TABLE,
        weight: 0.9,
      },
      // 图表意图
      {
        patterns: [/图表|柱状图|折线图|饼图|chart/i, /可视化|visualiz/i],
        intent: IntentType.CREATE_CHART,
        weight: 0.9,
      },
      // 分析意图
      {
        patterns: [/分析|汇总|平均|总计/, /analyz|sum|average/i],
        intent: IntentType.ANALYZE_DATA,
        weight: 0.85,
      },
      // 读取意图
      {
        patterns: [/读取|获取|查看|显示|是什么/, /read|get|show|display|what/i],
        intent: IntentType.READ_DATA,
        weight: 0.8,
      },
      // 写入意图
      {
        patterns: [/写入|设置|填充|输入|修改/, /write|set|fill|input|modify/i],
        intent: IntentType.WRITE_DATA,
        weight: 0.8,
      },
      // 条件格式
      {
        patterns: [/条件格式|高亮|标记|颜色编码/, /conditional|highlight|mark/i],
        intent: IntentType.CONDITIONAL_FORMAT,
        weight: 0.9,
      },
      // 计算意图
      {
        patterns: [/计算|求和|公式|函数/, /calculate|formula|function/i],
        intent: IntentType.CALCULATE,
        weight: 0.85,
      },
      // 删除意图
      {
        patterns: [/删除|清空|移除|清除/, /delete|clear|remove/i],
        intent: IntentType.DELETE_DATA,
        weight: 0.85,
      },
      // 撤销/重做
      {
        patterns: [/撤销|undo/i],
        intent: IntentType.UNDO,
        weight: 0.95,
      },
      {
        patterns: [/重做|redo/i],
        intent: IntentType.REDO,
        weight: 0.95,
      },
      // 帮助
      {
        patterns: [/帮助|怎么|如何|教我|help|how/i],
        intent: IntentType.HELP,
        weight: 0.7,
      },
      // v2.9.45: 续接对话
      {
        patterns: [/继续|接着|然后|还有|另外/],
        intent: IntentType.FOLLOW_UP,
        weight: 0.85,
      },
      // v2.9.45: 确认
      {
        patterns: [/好的|是的|确认|确定|执行吧|可以/],
        intent: IntentType.CONFIRMATION,
        weight: 0.9,
      },
    ];

    // 匹配意图
    for (const { patterns, intent, weight } of intentPatterns) {
      for (const pattern of patterns) {
        if (pattern.test(lowerContent)) {
          if (weight > confidence) {
            primaryIntent = intent;
            confidence = weight;
          }
          break;
        }
      }
    }

    // 提取实体
    this.extractEntities(content, entities);

    // 考虑上下文
    if (context && context.messages.length > 0) {
      const lastUserMessage = [...context.messages]
        .reverse()
        .find((m) => m.role === "user" && m.id !== this.generateMessageId());

      if (lastUserMessage?.metadata?.intent) {
        // 如果当前意图不明确，参考上一条消息
        if (confidence < 0.6) {
          const prevIntent = lastUserMessage.metadata.intent as IntentType;
          // 检查是否是续接对话
          if (/继续|接着|还有|另外/.test(content)) {
            primaryIntent = prevIntent;
            confidence = 0.7;
          }
        }
      }
    }

    return {
      primaryIntent,
      confidence,
      entities,
    };
  }

  /**
   * 提取实体
   */
  private extractEntities(content: string, entities: ExtractedEntity[]): void {
    // 提取单元格范围
    const rangePattern = /([A-Z]+\d+(?::[A-Z]+\d+)?)/gi;
    let match;
    while ((match = rangePattern.exec(content)) !== null) {
      entities.push({
        type: "cell_range",
        value: match[1].toUpperCase(),
        position: [match.index, match.index + match[1].length],
      });
    }

    // 提取颜色
    const colorPattern =
      /(#[0-9a-fA-F]{6}|红色?|蓝色?|绿色?|黄色?|橙色?|紫色?|黑色?|白色?|灰色?)/gi;
    while ((match = colorPattern.exec(content)) !== null) {
      entities.push({
        type: "color",
        value: match[1],
        position: [match.index, match.index + match[1].length],
      });
    }

    // 提取数字
    const numberPattern = /\b(\d+(?:\.\d+)?)\b/g;
    while ((match = numberPattern.exec(content)) !== null) {
      entities.push({
        type: "number",
        value: match[1],
        position: [match.index, match.index + match[1].length],
      });
    }

    // 提取工作表名
    const sheetPattern = /(?:工作表|表|sheet)\s*["']?([^"'\s,，]+)["']?/gi;
    while ((match = sheetPattern.exec(content)) !== null) {
      entities.push({
        type: "sheet_name",
        value: match[1],
        position: [match.index, match.index + match[0].length],
      });
    }
  }

  /**
   * 获取上下文窗口（用于发送给AI）
   */
  getContextWindow(): ConversationMessage[] {
    const session = this.getCurrentSession();
    if (!session) {
      return [];
    }

    const messages = session.messages;
    const { maxMessages, maxTokens, preserveRecentCount } = this.config;

    // 始终保留最近的N条消息
    const recentMessages = messages.slice(-preserveRecentCount);
    const olderMessages = messages.slice(0, -preserveRecentCount);

    // 计算Token总数
    let totalTokens = recentMessages.reduce((sum, m) => sum + (m.metadata?.tokenCount || 0), 0);

    // 从老消息中选择，直到达到Token限制
    const selectedOlder: ConversationMessage[] = [];
    for (let i = olderMessages.length - 1; i >= 0; i--) {
      const msg = olderMessages[i];
      const tokens = msg.metadata?.tokenCount || 0;

      // 优先保留关键消息
      if (msg.metadata?.isKeyMessage) {
        if (totalTokens + tokens <= maxTokens) {
          selectedOlder.unshift(msg);
          totalTokens += tokens;
        }
        continue;
      }

      // 使用压缩版本
      if (msg.compressed) {
        if (totalTokens + tokens <= maxTokens) {
          selectedOlder.unshift(msg);
          totalTokens += tokens;
        }
      } else if (totalTokens + tokens <= maxTokens) {
        selectedOlder.unshift(msg);
        totalTokens += tokens;
      }

      // 检查消息数限制
      if (selectedOlder.length + recentMessages.length >= maxMessages) {
        break;
      }
    }

    return [...selectedOlder, ...recentMessages];
  }

  /**
   * 查找相关引用
   */
  findReferences(query: string, topK: number = 3): Reference[] {
    const session = this.getCurrentSession();
    if (!session) {
      return [];
    }

    const references: Reference[] = [];
    const queryLower = query.toLowerCase();
    const queryWords = queryLower.split(/\s+/);

    for (const message of session.messages) {
      if (message.role !== "user" && message.role !== "assistant") {
        continue;
      }

      const contentLower = message.content.toLowerCase();

      // 计算简单的相关度分数
      let matchCount = 0;
      for (const word of queryWords) {
        if (word.length > 2 && contentLower.includes(word)) {
          matchCount++;
        }
      }

      if (matchCount > 0) {
        const relevanceScore = matchCount / queryWords.length;
        references.push({
          messageId: message.id,
          snippet: message.content.substring(0, 200),
          relevanceScore,
        });
      }
    }

    // 排序并返回Top K
    return references.sort((a, b) => b.relevanceScore - a.relevanceScore).slice(0, topK);
  }

  /**
   * 生成会话摘要
   */
  generateSummary(): string {
    const session = this.getCurrentSession();
    if (!session || session.messages.length === 0) {
      return "";
    }

    const userMessages = session.messages.filter((m) => m.role === "user");
    const toolCalls = session.messages.flatMap((m) => m.toolCalls || []).filter((tc) => tc.success);

    // 提取主要意图
    const intents = userMessages.map((m) => m.metadata?.intent).filter(Boolean) as string[];
    const intentCounts = intents.reduce(
      (acc, intent) => {
        acc[intent] = (acc[intent] || 0) + 1;
        return acc;
      },
      {} as Record<string, number>
    );

    const topIntent = Object.entries(intentCounts).sort((a, b) => b[1] - a[1])[0]?.[0] || "unknown";

    // 生成摘要
    const summary =
      `会话包含 ${userMessages.length} 条用户消息，主要意图：${topIntent}。` +
      `成功执行 ${toolCalls.length} 次工具调用。`;

    session.summary = summary;
    return summary;
  }

  /**
   * 更新用户偏好
   */
  updatePreferences(preferences: Partial<UserPreferences>): void {
    const session = this.getCurrentSession();
    if (!session) {
      return;
    }

    session.userPreferences = {
      ...session.userPreferences,
      ...preferences,
    };

    Logger.info("[ConversationMemory] 更新用户偏好", preferences);
  }

  /**
   * 获取用户偏好
   */
  getPreferences(): UserPreferences | undefined {
    return this.getCurrentSession()?.userPreferences;
  }

  /**
   * 设置当前任务上下文
   */
  setTaskContext(task: TaskContext): void {
    const session = this.getCurrentSession();
    if (session) {
      session.currentTask = task;
    }
  }

  /**
   * 获取当前任务上下文
   */
  getTaskContext(): TaskContext | undefined {
    return this.getCurrentSession()?.currentTask;
  }

  /**
   * 更新任务进度
   */
  updateTaskProgress(step: number, intermediateResult?: unknown): void {
    const session = this.getCurrentSession();
    if (!session?.currentTask) {
      return;
    }

    session.currentTask.currentStep = step;
    if (intermediateResult !== undefined) {
      session.currentTask.intermediateResults = session.currentTask.intermediateResults || [];
      session.currentTask.intermediateResults.push(intermediateResult);
    }
  }

  /**
   * 完成当前任务
   */
  completeTask(): void {
    const session = this.getCurrentSession();
    if (session) {
      session.currentTask = undefined;
    }
  }

  /**
   * 列出所有会话
   */
  listSessions(): Array<{
    sessionId: string;
    title: string;
    messageCount: number;
    lastActivityAt: number;
    status: string;
  }> {
    return Array.from(this.sessions.values()).map((session) => ({
      sessionId: session.sessionId,
      title: session.title || "",
      messageCount: session.messages.length,
      lastActivityAt: session.lastActivityAt,
      status: session.status,
    }));
  }

  /**
   * 删除会话
   */
  deleteSession(sessionId: string): boolean {
    if (!this.sessions.has(sessionId)) {
      return false;
    }

    this.sessions.delete(sessionId);
    if (this.currentSessionId === sessionId) {
      this.currentSessionId = null;
    }

    Logger.info("[ConversationMemory] 删除会话", { sessionId });
    return true;
  }

  /**
   * 清除所有会话
   */
  clearAll(): void {
    this.sessions.clear();
    this.currentSessionId = null;
    Logger.info("[ConversationMemory] 清除所有会话");
  }

  /**
   * 导出会话
   */
  exportSession(sessionId?: string): string {
    const id = sessionId || this.currentSessionId;
    if (!id) {
      return "";
    }

    const session = this.sessions.get(id);
    if (!session) {
      return "";
    }

    return JSON.stringify(session, null, 2);
  }

  /**
   * 导入会话
   */
  importSession(json: string): boolean {
    try {
      const session = JSON.parse(json) as ConversationContext;

      // 验证基本结构
      if (!session.sessionId || !Array.isArray(session.messages)) {
        Logger.error("[ConversationMemory] 导入会话格式无效");
        return false;
      }

      // 生成新的sessionId避免冲突
      const newSessionId = this.generateSessionId();
      session.sessionId = newSessionId;

      this.sessions.set(newSessionId, session);
      Logger.info("[ConversationMemory] 导入会话成功", { sessionId: newSessionId });

      return true;
    } catch (error) {
      Logger.error("[ConversationMemory] 导入会话失败", error);
      return false;
    }
  }

  /**
   * 保存到存储
   */
  saveToStorage(): void {
    try {
      const data = {
        currentSessionId: this.currentSessionId,
        sessions: Array.from(this.sessions.entries()),
      };
      localStorage.setItem("excel-copilot-conversations", JSON.stringify(data));
      Logger.info("[ConversationMemory] 保存到存储成功");
    } catch (error) {
      Logger.error("[ConversationMemory] 保存到存储失败", error);
    }
  }

  /**
   * 从存储加载
   */
  loadFromStorage(): void {
    try {
      const json = localStorage.getItem("excel-copilot-conversations");
      if (!json) {
        return;
      }

      const data = JSON.parse(json);
      if (data.sessions && Array.isArray(data.sessions)) {
        this.sessions = new Map(data.sessions);
        this.currentSessionId = data.currentSessionId || null;
        Logger.info("[ConversationMemory] 从存储加载成功");
      }
    } catch (error) {
      Logger.error("[ConversationMemory] 从存储加载失败", error);
    }
  }

  /**
   * 设置配置
   */
  setConfig(config: Partial<ContextWindowConfig>): void {
    this.config = { ...this.config, ...config };
  }

  /**
   * 获取配置
   */
  getConfig(): ContextWindowConfig {
    return { ...this.config };
  }

  /**
   * 重置
   */
  reset(): void {
    this.sessions.clear();
    this.currentSessionId = null;
  }

  // ============ 私有方法 ============

  private generateSessionId(): string {
    return `session_${Date.now()}_${Math.random().toString(36).substring(2, 9)}`;
  }

  private generateMessageId(): string {
    return `msg_${Date.now()}_${Math.random().toString(36).substring(2, 9)}`;
  }

  private estimateTokenCount(text: string): number {
    // 粗略估算：中文字符约1.5个token，英文约0.25个token/字符
    const chineseChars = (text.match(/[\u4e00-\u9fa5]/g) || []).length;
    const otherChars = text.length - chineseChars;
    return Math.ceil(chineseChars * 1.5 + otherChars * 0.25);
  }

  private checkAndCompress(session: ConversationContext): void {
    const totalTokens = session.messages.reduce((sum, m) => sum + (m.metadata?.tokenCount || 0), 0);

    if (totalTokens > this.config.maxTokens * this.config.compressionThreshold) {
      this.compressOldMessages(session);
    }
  }

  private compressOldMessages(session: ConversationContext): void {
    const messagesToCompress = session.messages.slice(0, -this.config.preserveRecentCount);

    for (const message of messagesToCompress) {
      if (message.compressed || message.metadata?.isKeyMessage) {
        continue;
      }

      // 简单压缩：截断长消息
      if (message.content.length > 500) {
        message.content = message.content.substring(0, 200) + "...（已压缩）";
        message.metadata = message.metadata || {};
        message.metadata.tokenCount = this.estimateTokenCount(message.content);
        message.compressed = true;
      }
    }

    Logger.debug("[ConversationMemory] 压缩旧消息完成");
  }
}

// ============ 单例导出 ============

export const ConversationMemory = new ConversationMemoryImpl();

// 便捷方法导出
export const memory = {
  createSession: (title?: string) => ConversationMemory.createSession(title),
  getCurrentSession: () => ConversationMemory.getCurrentSession(),
  addMessage: (
    role: MessageRole,
    content: string,
    metadata?: Partial<MessageMetadata>,
    toolCalls?: ToolCallRecord[]
  ) => ConversationMemory.addMessage(role, content, metadata, toolCalls),
  getContextWindow: () => ConversationMemory.getContextWindow(),
  analyzeIntent: (content: string) => ConversationMemory.analyzeIntent(content),
  findReferences: (query: string, topK?: number) => ConversationMemory.findReferences(query, topK),
  generateSummary: () => ConversationMemory.generateSummary(),
  setTaskContext: (task: TaskContext) => ConversationMemory.setTaskContext(task),
  getTaskContext: () => ConversationMemory.getTaskContext(),
  exportSession: (sessionId?: string) => ConversationMemory.exportSession(sessionId),
};
