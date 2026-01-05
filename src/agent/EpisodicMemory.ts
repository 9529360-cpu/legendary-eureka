/**
 * EpisodicMemory - 情景记忆系统
 *
 * 基于 ai-agents-for-beginners 第13课 Agent Memory 的学习
 * 实现 Episodic Memory，记录操作序列的成功/失败，用于改进未来执行
 *
 * 主要功能:
 * 1. 记录完整的操作序列（Episode）
 * 2. 分析成功/失败模式
 * 3. 提供历史经验参考
 * 4. 支持相似场景检索
 *
 * @version 1.0.0
 * @see docs/AI_AGENTS_FOR_BEGINNERS_LEARNING.md
 */

import { Logger } from "../utils/Logger";

// ============================================================================
// 类型定义
// ============================================================================

/**
 * 单个操作步骤
 */
export interface EpisodeStep {
  /** 步骤序号 */
  stepNumber: number;
  /** 工具名称 */
  toolName: string;
  /** 工具参数 */
  parameters: Record<string, unknown>;
  /** 执行结果 */
  result: "success" | "failure" | "skipped";
  /** 错误信息 */
  error?: string;
  /** 执行时长 (ms) */
  duration: number;
  /** 输出摘要 */
  outputSummary?: string;
}

/**
 * 完整的操作情景
 */
export interface Episode {
  /** 情景ID */
  id: string;
  /** 用户原始请求 */
  userRequest: string;
  /** 请求的关键词/标签 */
  tags: string[];
  /** 步骤列表 */
  steps: EpisodeStep[];
  /** 整体结果 */
  outcome: "success" | "partial" | "failure";
  /** 开始时间 */
  startTime: number;
  /** 结束时间 */
  endTime: number;
  /** 总时长(ms) */
  totalDuration: number;
  /** 成功率 */
  successRate: number;
  /** 失败原因（如果失败） */
  failureReason?: string;
  /** 学习要点 */
  learnings?: string[];
  /** 相关工作表/范围 */
  context?: {
    sheetName?: string;
    range?: string;
    dataSize?: number;
  };
}

/**
 * 模式分析结果
 */
export interface PatternAnalysis {
  /** 最常见的成功模式 */
  successPatterns: Array<{
    pattern: string;
    frequency: number;
    averageDuration: number;
  }>;
  /** 最常见的失败原因 */
  failurePatterns: Array<{
    reason: string;
    frequency: number;
    affectedTools: string[];
  }>;
  /** 工具使用统计 */
  toolStats: Map<
    string,
    {
      usageCount: number;
      successRate: number;
      averageDuration: number;
    }
  >;
  /** 建议 */
  recommendations: string[];
}

/**
 * 可复用经验 - 只存储有价值的知识，不是整段对话
 *
 * 设计原则：
 * - 用户偏好：用户明确表达的偏好或选择
 * - 失败原因：工具失败的原因和解决方案
 * - 有效参数：工具成功调用时的参数模式
 * - 任务模式：成功完成某类任务的步骤序列
 */
export interface ReusableExperience {
  /** 经验ID */
  id: string;
  /** 经验类型 */
  type: "user_preference" | "failure_reason" | "valid_parameters" | "task_pattern";
  /** 创建时间 */
  createdAt: number;
  /** 最后使用时间 */
  lastUsedAt: number;
  /** 使用次数 */
  usageCount: number;
  /** 经验内容 */
  content: UserPreference | FailureReason | ValidParameters | TaskPattern;
}

/**
 * 用户偏好
 */
export interface UserPreference {
  /** 偏好类型 */
  preferenceType: string;
  /** 偏好值 */
  value: unknown;
  /** 来源请求 */
  sourceRequest: string;
  /** 置信度 */
  confidence: number;
}

/**
 * 失败原因
 */
export interface FailureReason {
  /** 工具名称 */
  toolName: string;
  /** 错误类型 */
  errorType: string;
  /** 错误消息 */
  errorMessage: string;
  /** 触发条件 */
  triggerCondition: string;
  /** 解决方案 */
  solution?: string;
  /** 发生次数 */
  occurrenceCount: number;
}

/**
 * 有效参数模式
 */
export interface ValidParameters {
  /** 工具名称 */
  toolName: string;
  /** 任务类型 */
  taskType: string;
  /** 参数模板 */
  parameters: Record<string, unknown>;
  /** 成功率 */
  successRate: number;
  /** 使用次数 */
  usageCount: number;
}

/**
 * 任务模式 - 成功完成某类任务的步骤序列
 */
export interface TaskPattern {
  /** 任务类型 */
  taskType: string;
  /** 关键词 */
  keywords: string[];
  /** 步骤序列 */
  steps: Array<{
    toolName: string;
    parameterHints: Record<string, string>;
  }>;
  /** 成功率 */
  successRate: number;
  /** 平均时长 */
  averageDuration: number;
  /** 使用次数 */
  usageCount?: number;
}

/**
 * 存储配置
 */
export interface EpisodicMemoryConfig {
  /** 最大存储情景数 */
  maxEpisodes: number;
  /** 情景过期时间 (ms) */
  expirationTime: number;
  /** 是否启用持久化 */
  enablePersistence: boolean;
  /** 存储键名 */
  storageKey: string;
}

// ============================================================================
// 默认配置
// ============================================================================

const DEFAULT_CONFIG: EpisodicMemoryConfig = {
  maxEpisodes: 100,
  expirationTime: 7 * 24 * 60 * 60 * 1000, // 7天
  enablePersistence: true,
  storageKey: "excel-copilot-episodic-memory",
};

// ============================================================================
// 情景记忆器
// ============================================================================

/**
 * 情景记忆系统
 *
 * 记录和学习操作历史
 */
export class EpisodicMemory {
  private episodes: Episode[] = [];
  private config: EpisodicMemoryConfig;
  private readonly MODULE_NAME = "EpisodicMemory";
  private currentEpisode: Partial<Episode> | null = null;
  private currentSteps: EpisodeStep[] = [];

  constructor(config: Partial<EpisodicMemoryConfig> = {}) {
    this.config = { ...DEFAULT_CONFIG, ...config };
    this.loadFromStorage();
  }

  // ============================================================================
  // 记录 API
  // ============================================================================

  /**
   * 开始新的情景
   */
  startEpisode(userRequest: string, context?: Episode["context"]): string {
    const id = this.generateId();

    this.currentEpisode = {
      id,
      userRequest,
      tags: this.extractTags(userRequest),
      startTime: Date.now(),
      context,
    };
    this.currentSteps = [];

    Logger.debug(this.MODULE_NAME, "Episode started", { id, request: userRequest.substring(0, 50) });

    return id;
  }

  /**
   * 记录步骤
   */
  recordStep(step: Omit<EpisodeStep, "stepNumber">): void {
    if (!this.currentEpisode) {
      Logger.warn(this.MODULE_NAME, "No active episode, step not recorded");
      return;
    }

    this.currentSteps.push({
      ...step,
      stepNumber: this.currentSteps.length + 1,
    });

    Logger.debug(this.MODULE_NAME, "Step recorded", {
      episodeId: this.currentEpisode.id,
      tool: step.toolName,
      result: step.result,
    });
  }

  /**
   * 结束情景
   */
  endEpisode(learnings?: string[]): Episode | null {
    if (!this.currentEpisode) {
      Logger.warn(this.MODULE_NAME, "No active episode to end");
      return null;
    }

    const endTime = Date.now();
    const successfulSteps = this.currentSteps.filter((s) => s.result === "success").length;
    const totalSteps = this.currentSteps.length;

    const outcome: Episode["outcome"] =
      successfulSteps === totalSteps ? "success" : successfulSteps > 0 ? "partial" : "failure";

    const episode: Episode = {
      ...(this.currentEpisode as Episode),
      steps: this.currentSteps,
      outcome,
      endTime,
      totalDuration: endTime - (this.currentEpisode.startTime || endTime),
      successRate: totalSteps > 0 ? successfulSteps / totalSteps : 0,
      learnings,
    };

    // 如果失败，记录失败原因
    if (outcome === "failure" || outcome === "partial") {
      const failedStep = this.currentSteps.find((s) => s.result === "failure");
      episode.failureReason = failedStep?.error || "Unknown failure";
    }

    this.addEpisode(episode);
    this.currentEpisode = null;
    this.currentSteps = [];

    Logger.info(this.MODULE_NAME, "Episode ended", {
      id: episode.id,
      outcome,
      successRate: episode.successRate.toFixed(2),
    });

    return episode;
  }

  /**
   * 放弃当前情景
   */
  abandonEpisode(): void {
    if (this.currentEpisode) {
      Logger.debug(this.MODULE_NAME, "Episode abandoned", { id: this.currentEpisode.id });
    }
    this.currentEpisode = null;
    this.currentSteps = [];
  }

  // ============================================================================
  // 查询 API
  // ============================================================================

  /**
   * 查找相似情景
   */
  findSimilar(request: string, limit: number = 5): Episode[] {
    const tags = this.extractTags(request);

    // 计算相似度分数
    const scored = this.episodes.map((episode) => {
      let score = 0;

      // 标签匹配
      for (const tag of tags) {
        if (episode.tags.includes(tag)) {
          score += 10;
        }
      }

      // 成功情景加分
      if (episode.outcome === "success") {
        score += 5;
      }

      // 新近情景加分
      const ageHours = (Date.now() - episode.endTime) / (1000 * 60 * 60);
      score += Math.max(0, 10 - ageHours / 24); // 24小时内加满分

      return { episode, score };
    });

    return scored
      .sort((a, b) => b.score - a.score)
      .slice(0, limit)
      .map((s) => s.episode);
  }

  /**
   * 获取工具的历史使用情况
   */
  getToolHistory(toolName: string): {
    usageCount: number;
    successRate: number;
    averageDuration: number;
    recentErrors: string[];
  } {
    let usageCount = 0;
    let successCount = 0;
    let totalDuration = 0;
    const recentErrors: string[] = [];

    for (const episode of this.episodes) {
      for (const step of episode.steps) {
        if (step.toolName === toolName) {
          usageCount++;
          totalDuration += step.duration;

          if (step.result === "success") {
            successCount++;
          } else if (step.error) {
            recentErrors.push(step.error);
          }
        }
      }
    }

    return {
      usageCount,
      successRate: usageCount > 0 ? successCount / usageCount : 0,
      averageDuration: usageCount > 0 ? totalDuration / usageCount : 0,
      recentErrors: recentErrors.slice(-5), // 最后5个错误
    };
  }

  /**
   * 分析模式
   */
  analyzePatterns(): PatternAnalysis {
    const toolStats = new Map<
      string,
      {
        usageCount: number;
        successCount: number;
        totalDuration: number;
      }
    >();

    const failureReasons = new Map<
      string,
      {
        count: number;
        tools: Set<string>;
      }
    >();

    const successSequences = new Map<string, number>();

    // 遍历所有情景
    for (const episode of this.episodes) {
      // 统计工具使用
      for (const step of episode.steps) {
        const stats = toolStats.get(step.toolName) || {
          usageCount: 0,
          successCount: 0,
          totalDuration: 0,
        };

        stats.usageCount++;
        stats.totalDuration += step.duration;
        if (step.result === "success") {
          stats.successCount++;
        }

        toolStats.set(step.toolName, stats);

        // 统计失败原因
        if (step.result === "failure" && step.error) {
          const reason = this.normalizeError(step.error);
          const existing = failureReasons.get(reason) || { count: 0, tools: new Set() };
          existing.count++;
          existing.tools.add(step.toolName);
          failureReasons.set(reason, existing);
        }
      }

      // 统计成功序列
      if (episode.outcome === "success") {
        const sequence = episode.steps.map((s) => s.toolName).join(" -> ");
        successSequences.set(sequence, (successSequences.get(sequence) || 0) + 1);
      }
    }

    // 构建结果
    const successPatterns = [...successSequences.entries()]
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5)
      .map(([pattern, frequency]) => ({
        pattern,
        frequency,
        averageDuration: 0, // 简化实现
      }));

    const failurePatterns = [...failureReasons.entries()]
      .sort((a, b) => b[1].count - a[1].count)
      .slice(0, 5)
      .map(([reason, data]) => ({
        reason,
        frequency: data.count,
        affectedTools: [...data.tools],
      }));

    const toolStatsResult = new Map<
      string,
      {
        usageCount: number;
        successRate: number;
        averageDuration: number;
      }
    >();

    for (const [name, stats] of toolStats) {
      toolStatsResult.set(name, {
        usageCount: stats.usageCount,
        successRate: stats.successCount / stats.usageCount,
        averageDuration: stats.totalDuration / stats.usageCount,
      });
    }

    // 生成建议
    const recommendations = this.generateRecommendations(failurePatterns, toolStatsResult);

    return {
      successPatterns,
      failurePatterns,
      toolStats: toolStatsResult,
      recommendations,
    };
  }

  // ============================================================================
  // 可复用经验 API
  // ============================================================================

  /** 可复用经验存储 */
  private experiences: ReusableExperience[] = [];

  /**
   * 从 Episode 提取可复用经验
   *
   * 只提取有价值的知识，不是整段对话
   */
  extractReusableExperience(episode: Episode): ReusableExperience[] {
    const extracted: ReusableExperience[] = [];
    const now = Date.now();

    // 1. 提取失败原因
    for (const step of episode.steps) {
      if (step.result === "failure" && step.error) {
        const failureExp = this.createFailureExperience(step, now);
        if (failureExp) {
          extracted.push(failureExp);
        }
      }
    }

    // 2. 提取有效参数模式（仅从成功步骤）
    for (const step of episode.steps) {
      if (step.result === "success") {
        const paramExp = this.createValidParametersExperience(step, episode, now);
        if (paramExp) {
          extracted.push(paramExp);
        }
      }
    }

    // 3. 提取任务模式（仅从完全成功的 Episode）
    if (episode.outcome === "success" && episode.steps.length >= 2) {
      const patternExp = this.createTaskPatternExperience(episode, now);
      if (patternExp) {
        extracted.push(patternExp);
      }
    }

    // 合并或更新已有经验
    for (const exp of extracted) {
      this.mergeExperience(exp);
    }

    Logger.debug(this.MODULE_NAME, "Extracted reusable experiences", {
      episodeId: episode.id,
      extractedCount: extracted.length,
    });

    return extracted;
  }

  /**
   * 创建失败原因经验
   */
  private createFailureExperience(step: EpisodeStep, now: number): ReusableExperience | null {
    if (!step.error) return null;

    const errorType = this.categorizeError(step.error);

    return {
      id: `failure_${step.toolName}_${errorType}_${now}`,
      type: "failure_reason",
      createdAt: now,
      lastUsedAt: now,
      usageCount: 1,
      content: {
        toolName: step.toolName,
        errorType,
        errorMessage: step.error,
        triggerCondition: JSON.stringify(step.parameters),
        occurrenceCount: 1,
      } as FailureReason,
    };
  }

  /**
   * 创建有效参数经验
   */
  private createValidParametersExperience(
    step: EpisodeStep,
    episode: Episode,
    now: number
  ): ReusableExperience | null {
    // 只有参数非空时才记录
    if (!step.parameters || Object.keys(step.parameters).length === 0) {
      return null;
    }

    // 提取任务类型
    const taskType = this.inferTaskType(episode.userRequest);

    return {
      id: `params_${step.toolName}_${taskType}_${now}`,
      type: "valid_parameters",
      createdAt: now,
      lastUsedAt: now,
      usageCount: 1,
      content: {
        toolName: step.toolName,
        taskType,
        parameters: this.anonymizeParameters(step.parameters),
        successRate: 1,
        usageCount: 1,
      } as ValidParameters,
    };
  }

  /**
   * 创建任务模式经验
   */
  private createTaskPatternExperience(episode: Episode, now: number): ReusableExperience | null {
    const taskType = this.inferTaskType(episode.userRequest);

    // 提取步骤序列（只保留工具名和参数提示）
    const steps = episode.steps.map((s) => ({
      toolName: s.toolName,
      parameterHints: this.extractParameterHints(s.parameters),
    }));

    return {
      id: `pattern_${taskType}_${now}`,
      type: "task_pattern",
      createdAt: now,
      lastUsedAt: now,
      usageCount: 1,
      content: {
        taskType,
        keywords: episode.tags,
        steps,
        successRate: 1,
        averageDuration: episode.totalDuration,
      } as TaskPattern,
    };
  }

  /**
   * 合并或更新经验
   */
  private mergeExperience(newExp: ReusableExperience): void {
    // 查找相似经验
    const existingIndex = this.experiences.findIndex(
      (exp) => exp.type === newExp.type && this.isSimilarExperience(exp, newExp)
    );

    if (existingIndex >= 0) {
      // 更新已有经验
      const existing = this.experiences[existingIndex];
      existing.lastUsedAt = Date.now();
      existing.usageCount++;

      // 根据类型更新内容
      if (newExp.type === "failure_reason") {
        (existing.content as FailureReason).occurrenceCount++;
      } else if (newExp.type === "valid_parameters") {
        const content = existing.content as ValidParameters;
        content.usageCount++;
        // 更新成功率（简化实现）
      } else if (newExp.type === "task_pattern") {
        const content = existing.content as TaskPattern;
        content.usageCount = (content.usageCount || 0) + 1;
      }
    } else {
      // 添加新经验
      this.experiences.push(newExp);
    }
  }

  /**
   * 判断两个经验是否相似
   */
  private isSimilarExperience(exp1: ReusableExperience, exp2: ReusableExperience): boolean {
    if (exp1.type !== exp2.type) return false;

    switch (exp1.type) {
      case "failure_reason": {
        const c1 = exp1.content as FailureReason;
        const c2 = exp2.content as FailureReason;
        return c1.toolName === c2.toolName && c1.errorType === c2.errorType;
      }
      case "valid_parameters": {
        const c1 = exp1.content as ValidParameters;
        const c2 = exp2.content as ValidParameters;
        return c1.toolName === c2.toolName && c1.taskType === c2.taskType;
      }
      case "task_pattern": {
        const c1 = exp1.content as TaskPattern;
        const c2 = exp2.content as TaskPattern;
        return c1.taskType === c2.taskType;
      }
      default:
        return false;
    }
  }

  /**
   * 分类错误类型
   */
  private categorizeError(error: string): string {
    const lowerError = error.toLowerCase();
    if (lowerError.includes("not found") || lowerError.includes("不存在")) {
      return "not_found";
    }
    if (lowerError.includes("permission") || lowerError.includes("权限")) {
      return "permission";
    }
    if (lowerError.includes("invalid") || lowerError.includes("无效")) {
      return "invalid_input";
    }
    if (lowerError.includes("timeout") || lowerError.includes("超时")) {
      return "timeout";
    }
    return "unknown";
  }

  /**
   * 推断任务类型
   */
  private inferTaskType(request: string): string {
    const lowerRequest = request.toLowerCase();
    if (lowerRequest.includes("格式") || lowerRequest.includes("format")) {
      return "formatting";
    }
    if (lowerRequest.includes("公式") || lowerRequest.includes("formula")) {
      return "formula";
    }
    if (lowerRequest.includes("图表") || lowerRequest.includes("chart")) {
      return "chart";
    }
    if (lowerRequest.includes("排序") || lowerRequest.includes("sort")) {
      return "sort";
    }
    if (lowerRequest.includes("筛选") || lowerRequest.includes("filter")) {
      return "filter";
    }
    if (lowerRequest.includes("删除") || lowerRequest.includes("delete")) {
      return "delete";
    }
    if (lowerRequest.includes("复制") || lowerRequest.includes("copy")) {
      return "copy";
    }
    return "general";
  }

  /**
   * 匿名化参数（去除具体值，保留结构）
   */
  private anonymizeParameters(params: Record<string, unknown>): Record<string, unknown> {
    const anonymized: Record<string, unknown> = {};
    for (const [key, value] of Object.entries(params)) {
      if (typeof value === "string") {
        // 保留范围格式但匿名化
        if (/^[A-Z]+\d+:[A-Z]+\d+$/.test(value)) {
          anonymized[key] = "<range>";
        } else if (/^[A-Z]+\d+$/.test(value)) {
          anonymized[key] = "<cell>";
        } else {
          anonymized[key] = "<string>";
        }
      } else if (typeof value === "number") {
        anonymized[key] = "<number>";
      } else if (typeof value === "boolean") {
        anonymized[key] = value; // 保留布尔值
      } else if (Array.isArray(value)) {
        anonymized[key] = "<array>";
      } else {
        anonymized[key] = "<object>";
      }
    }
    return anonymized;
  }

  /**
   * 提取参数提示
   */
  private extractParameterHints(params: Record<string, unknown>): Record<string, string> {
    const hints: Record<string, string> = {};
    for (const [key, value] of Object.entries(params)) {
      if (typeof value === "string") {
        if (/^[A-Z]+\d+:[A-Z]+\d+$/.test(value)) {
          hints[key] = "范围，如 A1:B10";
        } else if (/^[A-Z]+\d+$/.test(value)) {
          hints[key] = "单元格，如 A1";
        } else if (value.startsWith("=")) {
          hints[key] = "公式";
        } else {
          hints[key] = "文本值";
        }
      } else if (typeof value === "number") {
        hints[key] = "数值";
      } else if (typeof value === "boolean") {
        hints[key] = "是/否";
      }
    }
    return hints;
  }

  /**
   * 获取相关经验
   */
  getRelevantExperiences(request: string, toolName?: string): ReusableExperience[] {
    const taskType = this.inferTaskType(request);

    return this.experiences
      .filter((exp) => {
        // 按任务类型过滤
        if (exp.type === "task_pattern") {
          const content = exp.content as TaskPattern;
          return content.taskType === taskType;
        }

        // 按工具名过滤
        if (toolName) {
          if (exp.type === "failure_reason") {
            return (exp.content as FailureReason).toolName === toolName;
          }
          if (exp.type === "valid_parameters") {
            return (exp.content as ValidParameters).toolName === toolName;
          }
        }

        return false;
      })
      .sort((a, b) => b.usageCount - a.usageCount);
  }

  /**
   * 获取统计摘要
   */
  getSummary(): {
    totalEpisodes: number;
    successRate: number;
    averageDuration: number;
    topTools: Array<{ name: string; count: number }>;
  } {
    if (this.episodes.length === 0) {
      return {
        totalEpisodes: 0,
        successRate: 0,
        averageDuration: 0,
        topTools: [],
      };
    }

    const successCount = this.episodes.filter((e) => e.outcome === "success").length;
    const totalDuration = this.episodes.reduce((sum, e) => sum + e.totalDuration, 0);

    const toolCounts = new Map<string, number>();
    for (const episode of this.episodes) {
      for (const step of episode.steps) {
        toolCounts.set(step.toolName, (toolCounts.get(step.toolName) || 0) + 1);
      }
    }

    const topTools = [...toolCounts.entries()]
      .sort((a, b) => b[1] - a[1])
      .slice(0, 5)
      .map(([name, count]) => ({ name, count }));

    return {
      totalEpisodes: this.episodes.length,
      successRate: successCount / this.episodes.length,
      averageDuration: totalDuration / this.episodes.length,
      topTools,
    };
  }

  // ============================================================================
  // 私有方法
  // ============================================================================

  /**
   * 生成唯一 ID
   */
  private generateId(): string {
    return `ep_${Date.now()}_${Math.random().toString(36).substring(2, 8)}`;
  }

  /**
   * 提取标签
   */
  private extractTags(text: string): string[] {
    const tags: string[] = [];

    // 关键词映射
    const keywords: Record<string, string> = {
      写入: "write",
      读取: "read",
      格式: "format",
      图表: "chart",
      公式: "formula",
      排序: "sort",
      筛选: "filter",
      删除: "delete",
      合并: "merge",
      拆分: "split",
      颜色: "color",
      字体: "font",
      边框: "border",
      求和: "sum",
      平均: "average",
      统计: "stats",
    };

    const textLower = text.toLowerCase();

    for (const [cn, en] of Object.entries(keywords)) {
      if (textLower.includes(cn) || textLower.includes(en)) {
        tags.push(en);
      }
    }

    return [...new Set(tags)];
  }

  /**
   * 添加情景
   */
  private addEpisode(episode: Episode): void {
    this.episodes.push(episode);

    // 限制数量
    if (this.episodes.length > this.config.maxEpisodes) {
      this.episodes = this.episodes.slice(-this.config.maxEpisodes);
    }

    // 清理过期
    this.cleanupExpired();

    // 持久化
    if (this.config.enablePersistence) {
      this.saveToStorage();
    }
  }

  /**
   * 清理过期情景
   */
  private cleanupExpired(): void {
    const now = Date.now();
    this.episodes = this.episodes.filter((e) => now - e.endTime < this.config.expirationTime);
  }

  /**
   * 规范化错误信息
   */
  private normalizeError(error: string): string {
    // 简化错误信息以便分析
    return error
      .replace(/\d+/g, "N") // 替换数字
      .replace(/'[^']+'/g, "'X'") // 替换引号内容
      .replace(/"[^"]+"/g, '"X"')
      .substring(0, 100);
  }

  /**
   * 生成建议
   */
  private generateRecommendations(
    failurePatterns: PatternAnalysis["failurePatterns"],
    toolStats: PatternAnalysis["toolStats"]
  ): string[] {
    const recommendations: string[] = [];

    // 基于失败模式的建议
    for (const pattern of failurePatterns.slice(0, 3)) {
      if (pattern.frequency >= 3) {
        recommendations.push(
          `注意：'${pattern.reason}' 错误频繁发生 (${pattern.frequency}次)，` +
            `涉及工具: ${pattern.affectedTools.join(", ")}`
        );
      }
    }

    // 基于工具统计的建议
    for (const [name, stats] of toolStats) {
      if (stats.usageCount >= 5 && stats.successRate < 0.5) {
        recommendations.push(
          `工具 '${name}' 成功率较低 (${(stats.successRate * 100).toFixed(0)}%)，` +
            `建议检查参数或使用替代方案`
        );
      }
    }

    return recommendations;
  }

  /**
   * 从存储加载
   */
  private loadFromStorage(): void {
    if (!this.config.enablePersistence) return;

    try {
      const stored = localStorage.getItem(this.config.storageKey);
      if (stored) {
        this.episodes = JSON.parse(stored);
        this.cleanupExpired();
        Logger.debug(this.MODULE_NAME, "Loaded episodes from storage", {
          count: this.episodes.length,
        });
      }
    } catch (error) {
      Logger.warn(this.MODULE_NAME, "Failed to load from storage", { error });
    }
  }

  /**
   * 保存到存储
   */
  private saveToStorage(): void {
    if (!this.config.enablePersistence) return;

    try {
      localStorage.setItem(this.config.storageKey, JSON.stringify(this.episodes));
    } catch (error) {
      Logger.warn(this.MODULE_NAME, "Failed to save to storage", { error });
    }
  }
}

// ============================================================================
// 便捷函数
// ============================================================================

let globalMemory: EpisodicMemory | null = null;

/**
 * 获取全局情景记忆实例
 */
export function getEpisodicMemory(): EpisodicMemory {
  if (!globalMemory) {
    globalMemory = new EpisodicMemory();
  }
  return globalMemory;
}

/**
 * 创建新的情景记忆实例
 */
export function createEpisodicMemory(config?: Partial<EpisodicMemoryConfig>): EpisodicMemory {
  return new EpisodicMemory(config);
}

/**
 * 快速记录成功操作
 */
export function recordSuccess(
  userRequest: string,
  toolName: string,
  parameters: Record<string, unknown>,
  duration: number
): void {
  const memory = getEpisodicMemory();
  memory.startEpisode(userRequest);
  memory.recordStep({
    toolName,
    parameters,
    result: "success",
    duration,
  });
  memory.endEpisode();
}

/**
 * 快速记录失败操作
 */
export function recordFailure(
  userRequest: string,
  toolName: string,
  parameters: Record<string, unknown>,
  error: string,
  duration: number
): void {
  const memory = getEpisodicMemory();
  memory.startEpisode(userRequest);
  memory.recordStep({
    toolName,
    parameters,
    result: "failure",
    error,
    duration,
  });
  memory.endEpisode([`失败原因: ${error}`]);
}
