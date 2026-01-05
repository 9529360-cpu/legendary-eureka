/**
 * ContextCompressor - 上下文压缩管理器
 *
 * 基于 ai-agents-for-beginners 第12课 Context Engineering 的学习
 * 实现上下文压缩策略，防止 Context Distraction 问题
 *
 * 主要功能:
 * 1. 对话历史自动摘要压缩
 * 2. 智能选择保留的消息
 * 3. Token 使用估算
 * 4. 关键信息提取
 *
 * @version 1.0.0
 * @see docs/AI_AGENTS_FOR_BEGINNERS_LEARNING.md
 */

import { Logger } from "../utils/Logger";
import type { ConversationMessage } from "../core/ConversationMemory";

// ============================================================================
// 类型定义
// ============================================================================

/**
 * 压缩配置
 */
export interface CompressionConfig {
  /** 最大 Token 数 */
  maxTokens: number;
  /** 触发压缩的阈值（占 maxTokens 的比例） */
  compressionThreshold: number;
  /** 保留最近消息数量 */
  keepRecentCount: number;
  /** 是否保留系统消息 */
  keepSystemMessages: boolean;
  /** 是否保留关键消息 */
  keepKeyMessages: boolean;
  /** 摘要最大 Token 数 */
  summaryMaxTokens: number;
}

/**
 * 压缩结果
 */
export interface CompressionResult {
  /** 压缩后的消息列表 */
  messages: ConversationMessage[];
  /** 压缩摘要（如果生成） */
  summary?: string;
  /** 原始 Token 数 */
  originalTokens: number;
  /** 压缩后 Token 数 */
  compressedTokens: number;
  /** 压缩比例 */
  compressionRatio: number;
  /** 被移除的消息数 */
  removedCount: number;
}

/**
 * 双份输出压缩结果
 *
 * short_history: 压缩后的历史，用于喂给 LLM
 * raw_history: 原始历史，用于审计留存
 */
export interface DualCompressionResult {
  /** 压缩后的消息（喂 LLM） */
  shortHistory: ConversationMessage[];
  /** 原始消息（审计留存） */
  rawHistory: ConversationMessage[];
  /** 压缩摘要 */
  summary?: string;
  /** 压缩统计 */
  stats: {
    originalTokens: number;
    compressedTokens: number;
    compressionRatio: number;
    removedCount: number;
    timestamp: number;
  };
}

/**
 * 消息重要性评分
 */
interface MessageImportance {
  message: ConversationMessage;
  score: number;
  reasons: string[];
}

// ============================================================================
// 默认配置
// ============================================================================

const DEFAULT_CONFIG: CompressionConfig = {
  maxTokens: 8000,
  compressionThreshold: 0.7, // 70% 时触发压缩
  keepRecentCount: 5,
  keepSystemMessages: true,
  keepKeyMessages: true,
  summaryMaxTokens: 500,
};

// ============================================================================
// 上下文压缩器
// ============================================================================

/**
 * 上下文压缩器
 *
 * 实现多种压缩策略以管理上下文窗口
 */
export class ContextCompressor {
  private config: CompressionConfig;
  private readonly MODULE_NAME = "ContextCompressor";

  constructor(config: Partial<CompressionConfig> = {}) {
    this.config = { ...DEFAULT_CONFIG, ...config };
  }

  /**
   * 检查是否需要压缩
   */
  needsCompression(messages: ConversationMessage[]): boolean {
    const currentTokens = this.estimateTotalTokens(messages);
    const threshold = this.config.maxTokens * this.config.compressionThreshold;
    return currentTokens > threshold;
  }

  /**
   * 压缩消息列表
   */
  compress(messages: ConversationMessage[]): CompressionResult {
    const originalTokens = this.estimateTotalTokens(messages);

    if (!this.needsCompression(messages)) {
      return {
        messages,
        originalTokens,
        compressedTokens: originalTokens,
        compressionRatio: 1,
        removedCount: 0,
      };
    }

    Logger.info(this.MODULE_NAME, "Starting compression", {
      messageCount: messages.length,
      estimatedTokens: originalTokens,
    });

    // 步骤1: 分类消息
    const { systemMessages, recentMessages, oldMessages, keyMessages } =
      this.categorizeMessages(messages);

    // 步骤2: 评估旧消息重要性
    const importanceScores = this.evaluateImportance(oldMessages);

    // 步骤3: 生成摘要
    const summary = this.generateLocalSummary(
      importanceScores
        .sort((a, b) => b.score - a.score)
        .slice(0, 10)
        .map((i) => i.message)
    );

    // 步骤4: 构建压缩后的消息列表
    const compressedMessages = this.buildCompressedMessages(
      systemMessages,
      summary,
      keyMessages,
      recentMessages
    );

    const compressedTokens = this.estimateTotalTokens(compressedMessages);

    Logger.info(this.MODULE_NAME, "Compression complete", {
      originalMessages: messages.length,
      compressedMessages: compressedMessages.length,
      originalTokens,
      compressedTokens,
      ratio: (compressedTokens / originalTokens).toFixed(2),
    });

    return {
      messages: compressedMessages,
      summary,
      originalTokens,
      compressedTokens,
      compressionRatio: compressedTokens / originalTokens,
      removedCount: messages.length - compressedMessages.length,
    };
  }

  /**
   * 双份输出压缩 - 核心方法
   *
   * 产出两份历史:
   * - shortHistory: 压缩后的消息，用于喂给 LLM（减少 token 消耗）
   * - rawHistory: 原始消息的深拷贝，用于审计留存
   *
   * @param messages 原始消息列表
   * @returns 双份输出结果
   */
  compressDual(messages: ConversationMessage[]): DualCompressionResult {
    // 深拷贝原始消息用于审计
    const rawHistory = messages.map((msg) => ({
      ...msg,
      metadata: { ...msg.metadata },
    }));

    // 执行压缩
    const compressionResult = this.compress(messages);

    Logger.info(this.MODULE_NAME, "Dual compression complete", {
      rawHistoryCount: rawHistory.length,
      shortHistoryCount: compressionResult.messages.length,
      compressionRatio: compressionResult.compressionRatio.toFixed(2),
    });

    return {
      shortHistory: compressionResult.messages,
      rawHistory,
      summary: compressionResult.summary,
      stats: {
        originalTokens: compressionResult.originalTokens,
        compressedTokens: compressionResult.compressedTokens,
        compressionRatio: compressionResult.compressionRatio,
        removedCount: compressionResult.removedCount,
        timestamp: Date.now(),
      },
    };
  }

  /**
   * 滑动窗口双份输出压缩
   *
   * @param messages 原始消息列表
   * @param windowSize 保留窗口大小
   * @returns 双份输出结果
   */
  slidingWindowCompressDual(
    messages: ConversationMessage[],
    windowSize: number = this.config.keepRecentCount
  ): DualCompressionResult {
    // 深拷贝原始消息用于审计
    const rawHistory = messages.map((msg) => ({
      ...msg,
      metadata: { ...msg.metadata },
    }));

    // 执行滑动窗口压缩
    const compressionResult = this.slidingWindowCompress(messages, windowSize);

    return {
      shortHistory: compressionResult.messages,
      rawHistory,
      summary: compressionResult.summary,
      stats: {
        originalTokens: compressionResult.originalTokens,
        compressedTokens: compressionResult.compressedTokens,
        compressionRatio: compressionResult.compressionRatio,
        removedCount: compressionResult.removedCount,
        timestamp: Date.now(),
      },
    };
  }

  /**
   * 滑动窗口压缩
   * 保留最近的 N 条消息，生成之前消息的摘要
   */
  slidingWindowCompress(
    messages: ConversationMessage[],
    windowSize: number = this.config.keepRecentCount
  ): CompressionResult {
    if (messages.length <= windowSize) {
      const tokens = this.estimateTotalTokens(messages);
      return {
        messages,
        originalTokens: tokens,
        compressedTokens: tokens,
        compressionRatio: 1,
        removedCount: 0,
      };
    }

    const originalTokens = this.estimateTotalTokens(messages);

    // 分离系统消息和对话消息
    const systemMessages = messages.filter((m) => m.role === "system");
    const conversationMessages = messages.filter((m) => m.role !== "system");

    // 保留最近的窗口
    const recentMessages = conversationMessages.slice(-windowSize);
    const oldMessages = conversationMessages.slice(0, -windowSize);

    // 生成摘要
    const summary = this.generateLocalSummary(oldMessages);

    // 构建结果
    const compressedMessages: ConversationMessage[] = [...systemMessages];

    if (summary) {
      compressedMessages.push({
        id: `summary-${Date.now()}`,
        role: "system",
        content: `[Previous conversation summary: ${summary}]`,
        timestamp: Date.now(),
        compressed: true,
      });
    }

    compressedMessages.push(...recentMessages);

    const compressedTokens = this.estimateTotalTokens(compressedMessages);

    return {
      messages: compressedMessages,
      summary,
      originalTokens,
      compressedTokens,
      compressionRatio: compressedTokens / originalTokens,
      removedCount: oldMessages.length,
    };
  }

  /**
   * 增量压缩
   * 用于实时对话，每次添加消息时检查是否需要压缩
   */
  incrementalCompress(
    messages: ConversationMessage[],
    newMessage: ConversationMessage
  ): CompressionResult {
    const allMessages = [...messages, newMessage];

    if (!this.needsCompression(allMessages)) {
      const tokens = this.estimateTotalTokens(allMessages);
      return {
        messages: allMessages,
        originalTokens: tokens,
        compressedTokens: tokens,
        compressionRatio: 1,
        removedCount: 0,
      };
    }

    // 使用滑动窗口压缩
    return this.slidingWindowCompress(allMessages);
  }

  // ============================================================================
  // 私有方法
  // ============================================================================

  /**
   * 估算单条消息的 Token 数
   * 使用简单的估算：平均每 4 个字符 = 1 个 token
   */
  private estimateTokens(content: string): number {
    // 中文字符计数更高
    const chineseChars = (content.match(/[\u4e00-\u9fa5]/g) || []).length;
    const otherChars = content.length - chineseChars;

    // 中文约 1.5 token/字，英文约 0.25 token/字符
    return Math.ceil(chineseChars * 1.5 + otherChars * 0.25);
  }

  /**
   * 估算消息列表的总 Token 数
   */
  private estimateTotalTokens(messages: ConversationMessage[]): number {
    return messages.reduce((sum, m) => sum + this.estimateTokens(m.content), 0);
  }

  /**
   * 分类消息
   */
  private categorizeMessages(messages: ConversationMessage[]): {
    systemMessages: ConversationMessage[];
    recentMessages: ConversationMessage[];
    oldMessages: ConversationMessage[];
    keyMessages: ConversationMessage[];
  } {
    const systemMessages: ConversationMessage[] = [];
    const keyMessages: ConversationMessage[] = [];
    const otherMessages: ConversationMessage[] = [];

    for (const msg of messages) {
      if (msg.role === "system") {
        systemMessages.push(msg);
      } else if (this.isKeyMessage(msg)) {
        keyMessages.push(msg);
      } else {
        otherMessages.push(msg);
      }
    }

    const keepCount = this.config.keepRecentCount;
    const recentMessages = otherMessages.slice(-keepCount);
    const oldMessages = otherMessages.slice(0, -keepCount);

    return { systemMessages, recentMessages, oldMessages, keyMessages };
  }

  /**
   * 判断是否为关键消息
   */
  private isKeyMessage(message: ConversationMessage): boolean {
    if (!this.config.keepKeyMessages) return false;

    // 检查元数据标记
    if (message.metadata?.isKeyMessage) return true;

    // 检查是否包含工具调用
    if (message.toolCalls && message.toolCalls.length > 0) return true;

    // 检查是否包含关键词
    const keyPatterns = [
      /重要/,
      /记住/,
      /始终/,
      /always/i,
      /important/i,
      /remember/i,
      /preference/i,
      /偏好/,
      /设置/,
    ];

    return keyPatterns.some((p) => p.test(message.content));
  }

  /**
   * 评估消息重要性
   */
  private evaluateImportance(messages: ConversationMessage[]): MessageImportance[] {
    return messages.map((message) => {
      let score = 0;
      const reasons: string[] = [];

      // 1. 包含实体信息
      if (message.metadata?.entities && message.metadata.entities.length > 0) {
        score += 20 * message.metadata.entities.length;
        reasons.push("Contains entities");
      }

      // 2. 包含工具调用
      if (message.toolCalls && message.toolCalls.length > 0) {
        score += 30;
        reasons.push("Contains tool calls");
      }

      // 3. 成功的工具调用
      if (message.toolCalls?.some((tc) => tc.success)) {
        score += 20;
        reasons.push("Successful tool execution");
      }

      // 4. 用户消息权重更高
      if (message.role === "user") {
        score += 15;
        reasons.push("User message");
      }

      // 5. 消息长度（太短可能不重要）
      const contentLength = message.content.length;
      if (contentLength > 100) {
        score += 10;
        reasons.push("Substantial content");
      }

      // 6. 包含数字或公式
      if (/[=+\-*/]|\d+/.test(message.content)) {
        score += 10;
        reasons.push("Contains calculations");
      }

      // 7. 时间衰减（越旧分数越低）
      const ageHours = (Date.now() - message.timestamp) / (1000 * 60 * 60);
      const decayFactor = Math.exp(-ageHours / 24); // 24小时衰减
      score *= decayFactor;

      return { message, score, reasons };
    });
  }

  /**
   * 生成本地摘要（不调用 LLM）
   */
  private generateLocalSummary(messages: ConversationMessage[]): string {
    if (messages.length === 0) return "";

    const parts: string[] = [];

    // 提取用户请求摘要
    const userMessages = messages.filter((m) => m.role === "user");
    if (userMessages.length > 0) {
      const requests = userMessages.map((m) => this.extractKeyPhrase(m.content)).filter(Boolean);
      if (requests.length > 0) {
        parts.push(`User requests: ${requests.join("; ")}`);
      }
    }

    // 提取工具调用摘要
    const toolCalls = messages.flatMap((m) => m.toolCalls || []);
    if (toolCalls.length > 0) {
      const successCount = toolCalls.filter((tc) => tc.success).length;
      const toolNames = [...new Set(toolCalls.map((tc) => tc.toolName))];
      parts.push(
        `Tools used: ${toolNames.join(", ")} (${successCount}/${toolCalls.length} successful)`
      );
    }

    // 提取实体
    const entities = messages
      .flatMap((m) => m.metadata?.entities || [])
      .filter((e, i, arr) => arr.findIndex((x) => x.value === e.value) === i);
    if (entities.length > 0) {
      const entityStr = entities
        .slice(0, 5)
        .map((e) => `${e.type}:${e.value}`)
        .join(", ");
      parts.push(`Entities: ${entityStr}`);
    }

    return parts.join(". ");
  }

  /**
   * 提取关键短语
   */
  private extractKeyPhrase(content: string): string {
    // 简单实现：取前 50 个字符
    const trimmed = content.trim();
    if (trimmed.length <= 50) return trimmed;

    // 尝试在标点处截断
    const cutoff = trimmed.substring(0, 50);
    const lastPunct = Math.max(
      cutoff.lastIndexOf("。"),
      cutoff.lastIndexOf("，"),
      cutoff.lastIndexOf("."),
      cutoff.lastIndexOf(",")
    );

    if (lastPunct > 20) {
      return trimmed.substring(0, lastPunct + 1);
    }

    return cutoff + "...";
  }

  /**
   * 构建压缩后的消息列表
   */
  private buildCompressedMessages(
    systemMessages: ConversationMessage[],
    summary: string,
    keyMessages: ConversationMessage[],
    recentMessages: ConversationMessage[]
  ): ConversationMessage[] {
    const result: ConversationMessage[] = [];

    // 1. 添加系统消息
    if (this.config.keepSystemMessages) {
      result.push(...systemMessages);
    }

    // 2. 添加摘要（作为系统消息）
    if (summary) {
      result.push({
        id: `summary-${Date.now()}`,
        role: "system",
        content: `[Conversation history summary: ${summary}]`,
        timestamp: Date.now(),
        compressed: true,
        metadata: {
          tokenCount: this.estimateTokens(summary),
        },
      });
    }

    // 3. 添加关键消息
    if (this.config.keepKeyMessages) {
      result.push(...keyMessages);
    }

    // 4. 添加最近消息
    result.push(...recentMessages);

    return result;
  }
}

// ============================================================================
// 便捷函数
// ============================================================================

/**
 * 创建默认压缩器
 */
export function createCompressor(config?: Partial<CompressionConfig>): ContextCompressor {
  return new ContextCompressor(config);
}

/**
 * 快速压缩消息
 */
export function compressMessages(
  messages: ConversationMessage[],
  maxTokens: number = 8000
): CompressionResult {
  const compressor = new ContextCompressor({ maxTokens });
  return compressor.compress(messages);
}

/**
 * 估算消息 Token 数
 */
export function estimateTokenCount(content: string): number {
  const chineseChars = (content.match(/[\u4e00-\u9fa5]/g) || []).length;
  const otherChars = content.length - chineseChars;
  return Math.ceil(chineseChars * 1.5 + otherChars * 0.25);
}
