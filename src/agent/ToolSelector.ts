/**
 * ToolSelector - 动态工具选择器
 *
 * 基于 ai-agents-for-beginners 第2章 Context Engineering 的学习
 * 实现动态工具选择，限制工具数量(<25个)，防止 Context Confusion
 *
 * 主要功能:
 * 1. 基于查询相关性选择工具
 * 2. 工具分类与分组
 * 3. 工具使用频率追踪
 * 4. 智能工具推荐
 *
 * @version 1.0.0
 * @see docs/AI_AGENTS_FOR_BEGINNERS_LEARNING.md
 */

import { Logger } from "../utils/Logger";
import type { Tool } from "./AgentCore";

// ============================================================================
// 类型定义
// ============================================================================

/**
 * 工具分类
 */
export type ToolCategory =
  | "read" // 读取操作
  | "write" // 写入操作
  | "format" // 格式化操作
  | "calculate" // 计算操作
  | "chart" // 图表操作
  | "data" // 数据处理
  | "navigate" // 导航操作
  | "utility" // 工具类
  | "advanced"; // 高级操作

/**
 * 工具元数据
 */
export interface ToolMetadata {
  /** 工具名称 */
  name: string;
  /** 分类 */
  category: ToolCategory;
  /** 关键词列表 */
  keywords: string[];
  /** 使用频率 */
  usageCount: number;
  /** 最后使用时间 */
  lastUsed?: number;
  /** 成功率 */
  successRate: number;
  /** 复杂度(1-5) */
  complexity: number;
  /** 是否高风险 */
  isHighRisk: boolean;
  /** 依赖的其他工具 */
  dependencies?: string[];
}

/**
 * 选择配置
 */
export interface SelectionConfig {
  /** 最大工具数 */
  maxTools: number;
  /** 是否包含依赖工具 */
  includeDependencies: boolean;
  /** 是否优先高频工具 */
  preferFrequent: boolean;
  /** 是否过滤高风险工具 */
  filterHighRisk: boolean;
  /** 必须包含的工具 */
  requiredTools?: string[];
}

/**
 * 选择结果
 */
export interface SelectionResult {
  /** 选中的工具 */
  tools: Tool[];
  /** 选择原因 */
  reasons: Map<string, string[]>;
  /** 被排除的工具 */
  excluded: string[];
  /** 相关性分数 */
  scores: Map<string, number>;
}

/**
 * LLM 工具子集 - 用于给 LLM 的精简工具表
 */
export interface LLMToolSubset {
  /** 选中的工具 */
  tools: Tool[];
  /** 精简的工具描述（减少 token） */
  toolDescriptions: Array<{
    name: string;
    description: string;
    parameters?: Array<{
      name: string;
      type: string;
      required: boolean;
      description?: string;
    }>;
  }>;
  /** 选择摘要 */
  summary: string;
  /** 统计信息 */
  stats: {
    selectedCount: number;
    totalAvailable: number;
    reductionRatio: number;
    prepareTime: number;
  };
}

// ============================================================================
// 默认配置
// ============================================================================

const DEFAULT_CONFIG: SelectionConfig = {
  maxTools: 25,
  includeDependencies: true,
  preferFrequent: true,
  filterHighRisk: false,
};

/**
 * 关键词到分类的映射
 */
const KEYWORD_CATEGORIES: Record<string, ToolCategory[]> = {
  // 读取类关键词
  读取: ["read"],
  获取: ["read"],
  查看: ["read"],
  显示: ["read"],
  read: ["read"],
  get: ["read"],
  show: ["read"],

  // 写入类关键词
  写入: ["write"],
  设置: ["write"],
  填充: ["write"],
  修改: ["write"],
  write: ["write"],
  set: ["write"],
  fill: ["write"],
  update: ["write"],

  // 格式类关键词
  格式: ["format"],
  颜色: ["format"],
  字体: ["format"],
  边框: ["format"],
  样式: ["format"],
  format: ["format"],
  color: ["format"],
  font: ["format"],
  style: ["format"],

  // 计算类关键词
  计算: ["calculate"],
  公式: ["calculate"],
  求和: ["calculate"],
  平均: ["calculate"],
  统计: ["calculate"],
  calculate: ["calculate"],
  formula: ["calculate"],
  sum: ["calculate"],
  average: ["calculate"],

  // 图表类关键词
  图表: ["chart"],
  图形: ["chart"],
  可视化: ["chart"],
  chart: ["chart"],
  graph: ["chart"],
  visualize: ["chart"],

  // 数据类关键词
  数据: ["data", "read", "write"],
  排序: ["data"],
  筛选: ["data"],
  过滤: ["data"],
  去重: ["data"],
  data: ["data"],
  sort: ["data"],
  filter: ["data"],
  unique: ["data"],

  // 导航类关键词
  跳转: ["navigate"],
  选择: ["navigate", "read"],
  定位: ["navigate"],
  工作表: ["navigate"],
  navigate: ["navigate"],
  select: ["navigate"],
  goto: ["navigate"],
  sheet: ["navigate"],

  // 高级类关键词
  批量: ["advanced", "write"],
  高级: ["advanced"],
  复杂: ["advanced"],
  batch: ["advanced"],
  advanced: ["advanced"],
};

// ============================================================================
// 工具选择器
// ============================================================================

/**
 * 动态工具选择器
 *
 * 根据用户查询智能选择最相关的工具子集
 */
export class ToolSelector {
  private tools: Map<string, Tool> = new Map();
  private metadata: Map<string, ToolMetadata> = new Map();
  private readonly MODULE_NAME = "ToolSelector";

  constructor() {
    this.initializeDefaultMetadata();
  }

  /**
   * 注册工具
   */
  registerTool(tool: Tool, meta?: Partial<ToolMetadata>): void {
    this.tools.set(tool.name, tool);

    // 自动推断元数据
    const inferredMeta = this.inferMetadata(tool);
    this.metadata.set(tool.name, {
      ...inferredMeta,
      ...meta,
      name: tool.name,
    });

    Logger.debug(this.MODULE_NAME, "Tool registered", { name: tool.name, category: inferredMeta.category });
  }

  /**
   * 批量注册工具
   */
  registerTools(tools: Tool[]): void {
    for (const tool of tools) {
      this.registerTool(tool);
    }
    Logger.info(this.MODULE_NAME, "Tools registered", { count: tools.length });
  }

  /**
   * 根据查询选择工具
   */
  selectTools(query: string, config: Partial<SelectionConfig> = {}): SelectionResult {
    const mergedConfig = { ...DEFAULT_CONFIG, ...config };
    const scores = new Map<string, number>();
    const reasons = new Map<string, string[]>();

    // 1. 计算每个工具的相关性分数
    for (const [name, tool] of this.tools) {
      const { score, matchReasons } = this.calculateRelevance(query, tool, name);
      scores.set(name, score);
      reasons.set(name, matchReasons);
    }

    // 2. 添加必需工具
    const selectedNames = new Set<string>(mergedConfig.requiredTools || []);

    // 3. 按分数排序并选择
    const sortedTools = [...scores.entries()]
      .sort((a, b) => b[1] - a[1])
      .filter(([name]) => !selectedNames.has(name));

    // 4. 选择高分工具
    for (const [name, score] of sortedTools) {
      if (selectedNames.size >= mergedConfig.maxTools) break;
      if (score <= 0) break;

      const meta = this.metadata.get(name);

      // 过滤高风险工具
      if (mergedConfig.filterHighRisk && meta?.isHighRisk) {
        continue;
      }

      selectedNames.add(name);

      // 添加依赖工具
      if (mergedConfig.includeDependencies && meta?.dependencies) {
        for (const dep of meta.dependencies) {
          if (selectedNames.size < mergedConfig.maxTools) {
            selectedNames.add(dep);
          }
        }
      }
    }

    // 5. 构建结果
    const selectedTools = [...selectedNames]
      .map((name) => this.tools.get(name))
      .filter((t): t is Tool => t !== undefined);

    const excluded = [...this.tools.keys()].filter((name) => !selectedNames.has(name));

    Logger.info(this.MODULE_NAME, "Tools selected", {
      query: query.substring(0, 50),
      selected: selectedTools.length,
      excluded: excluded.length,
    });

    return {
      tools: selectedTools,
      reasons,
      excluded,
      scores,
    };
  }

  /**
   * 根据分类选择工具
   */
  selectByCategory(categories: ToolCategory[], maxPerCategory: number = 5): Tool[] {
    const result: Tool[] = [];
    const categoryCount = new Map<ToolCategory, number>();

    for (const [name, meta] of this.metadata) {
      if (!categories.includes(meta.category)) continue;

      const count = categoryCount.get(meta.category) || 0;
      if (count >= maxPerCategory) continue;

      const tool = this.tools.get(name);
      if (tool) {
        result.push(tool);
        categoryCount.set(meta.category, count + 1);
      }
    }

    return result;
  }

  /**
   * 获取常用工具
   */
  getFrequentTools(limit: number = 10): Tool[] {
    return [...this.metadata.entries()]
      .sort((a, b) => b[1].usageCount - a[1].usageCount)
      .slice(0, limit)
      .map(([name]) => this.tools.get(name))
      .filter((t): t is Tool => t !== undefined);
  }

  /**
   * 记录工具使用
   */
  recordUsage(toolName: string, success: boolean): void {
    const meta = this.metadata.get(toolName);
    if (!meta) return;

    meta.usageCount++;
    meta.lastUsed = Date.now();

    // 更新成功率
    const totalUses = meta.usageCount;
    const previousSuccesses = meta.successRate * (totalUses - 1);
    meta.successRate = (previousSuccesses + (success ? 1 : 0)) / totalUses;

    this.metadata.set(toolName, meta);
  }

  /**
   * 获取工具统计
   */
  getStats(): {
    totalTools: number;
    byCategory: Map<ToolCategory, number>;
    topUsed: Array<{ name: string; count: number }>;
  } {
    const byCategory = new Map<ToolCategory, number>();
    const usageCounts: Array<{ name: string; count: number }> = [];

    for (const [name, meta] of this.metadata) {
      const current = byCategory.get(meta.category) || 0;
      byCategory.set(meta.category, current + 1);
      usageCounts.push({ name, count: meta.usageCount });
    }

    return {
      totalTools: this.tools.size,
      byCategory,
      topUsed: usageCounts.sort((a, b) => b.count - a.count).slice(0, 10),
    };
  }

  // ============================================================================
  // LLM 工具子集 API
  // ============================================================================

  /**
   * 为 LLM 准备工具候选子集
   *
   * 设计原则：
   * - 不要每次把全工具给 LLM
   * - 只给候选子集（减少幻觉工具调用 & token）
   * - 动态选择最相关的工具
   *
   * @param userRequest 用户请求
   * @param maxTools 最大工具数（默认15，建议不超过25）
   * @returns 工具子集及选择原因
   */
  prepareToolsForLLM(userRequest: string, maxTools: number = 15): LLMToolSubset {
    const startTime = Date.now();

    // 1. 基于请求选择相关工具
    const selectionResult = this.selectTools(userRequest, {
      maxTools,
      includeDependencies: true,
      preferFrequent: true,
      filterHighRisk: false, // 高风险工具也要给 LLM 选择权
    });

    // 2. 生成工具描述（精简版，减少 token）
    const toolDescriptions = selectionResult.tools.map((tool) => ({
      name: tool.name,
      description: this.compactDescription(tool.description),
      parameters: tool.parameters?.map((p) => ({
        name: p.name,
        type: p.type,
        required: p.required,
        description: p.description?.substring(0, 50), // 截断描述
      })),
    }));

    // 3. 生成选择摘要
    const summary = this.generateSelectionSummary(selectionResult, userRequest);

    Logger.info(this.MODULE_NAME, "Tools prepared for LLM", {
      request: userRequest.substring(0, 50),
      selectedCount: selectionResult.tools.length,
      totalAvailable: this.tools.size,
      prepareTime: Date.now() - startTime,
    });

    return {
      tools: selectionResult.tools,
      toolDescriptions,
      summary,
      stats: {
        selectedCount: selectionResult.tools.length,
        totalAvailable: this.tools.size,
        reductionRatio: 1 - selectionResult.tools.length / this.tools.size,
        prepareTime: Date.now() - startTime,
      },
    };
  }

  /**
   * 精简工具描述
   */
  private compactDescription(description: string): string {
    // 取第一句话
    const firstSentence = description.split(/[。！？]/)[0];
    return firstSentence.length > 100 ? firstSentence.substring(0, 100) + "..." : firstSentence;
  }

  /**
   * 生成选择摘要
   */
  private generateSelectionSummary(result: SelectionResult, request: string): string {
    const categories = new Set<ToolCategory>();
    for (const tool of result.tools) {
      const meta = this.metadata.get(tool.name);
      if (meta) categories.add(meta.category);
    }

    const categoryNames = [...categories].join(", ");
    return `基于请求"${request.substring(0, 30)}..."选择了 ${result.tools.length} 个工具，涵盖分类: ${categoryNames}`;
  }

  // ============================================================================
  // 私有方法
  // ============================================================================

  /**
   * 初始化默认元数据
   */
  private initializeDefaultMetadata(): void {
    // 预定义一些常用工具的元数据
    // 实际元数据会在注册工具时自动推断
  }

  /**
   * 推断工具元数据
   */
  private inferMetadata(tool: Tool): ToolMetadata {
    const name = tool.name.toLowerCase();
    const desc = tool.description.toLowerCase();

    // 推断分类
    let category: ToolCategory = "utility";
    if (name.includes("read") || name.includes("get") || desc.includes("读取")) {
      category = "read";
    } else if (name.includes("write") || name.includes("set") || desc.includes("写入")) {
      category = "write";
    } else if (name.includes("format") || desc.includes("格式")) {
      category = "format";
    } else if (name.includes("chart") || desc.includes("图表")) {
      category = "chart";
    } else if (name.includes("formula") || name.includes("calc") || desc.includes("计算")) {
      category = "calculate";
    } else if (name.includes("sort") || name.includes("filter") || desc.includes("排序")) {
      category = "data";
    } else if (name.includes("sheet") || name.includes("navigate")) {
      category = "navigate";
    }

    // 提取关键词
    const keywords = this.extractKeywords(tool.description);

    // 判断是否高风险
    const isHighRisk =
      name.includes("delete") ||
      name.includes("clear") ||
      name.includes("remove") ||
      desc.includes("删除") ||
      desc.includes("清空");

    // 估算复杂度
    const complexity = Math.min(5, Math.ceil((tool.parameters?.length || 0) / 2) + 1);

    return {
      name: tool.name,
      category,
      keywords,
      usageCount: 0,
      successRate: 1,
      complexity,
      isHighRisk,
    };
  }

  /**
   * 提取关键词
   */
  private extractKeywords(text: string): string[] {
    const words = text
      .toLowerCase()
      .replace(/[^\u4e00-\u9fa5a-z\s]/g, " ")
      .split(/\s+/)
      .filter((w) => w.length >= 2);

    return [...new Set(words)];
  }

  /**
   * 计算相关性分数
   */
  private calculateRelevance(
    query: string,
    tool: Tool,
    toolName: string
  ): { score: number; matchReasons: string[] } {
    let score = 0;
    const matchReasons: string[] = [];
    const queryLower = query.toLowerCase();
    const meta = this.metadata.get(toolName);

    // 1. 名称匹配 (高权重)
    const nameParts = toolName.split("_");
    for (const part of nameParts) {
      if (queryLower.includes(part.toLowerCase())) {
        score += 30;
        matchReasons.push(`Name match: ${part}`);
      }
    }

    // 2. 描述匹配
    const descWords = tool.description.toLowerCase().split(/\s+/);
    for (const word of descWords) {
      if (word.length >= 2 && queryLower.includes(word)) {
        score += 10;
        if (matchReasons.length < 5) {
          matchReasons.push(`Desc match: ${word}`);
        }
      }
    }

    // 3. 关键词分类匹配
    for (const [keyword, categories] of Object.entries(KEYWORD_CATEGORIES)) {
      if (queryLower.includes(keyword)) {
        if (meta && categories.includes(meta.category)) {
          score += 25;
          matchReasons.push(`Category match: ${keyword} -> ${meta.category}`);
        }
      }
    }

    // 4. 元数据关键词匹配
    if (meta) {
      for (const kw of meta.keywords) {
        if (queryLower.includes(kw)) {
          score += 15;
          if (matchReasons.length < 5) {
            matchReasons.push(`Keyword match: ${kw}`);
          }
        }
      }

      // 5. 使用频率加成
      if (meta.usageCount > 0) {
        score += Math.min(10, meta.usageCount);
        if (meta.usageCount > 5) {
          matchReasons.push(`Frequently used (${meta.usageCount})`);
        }
      }

      // 6. 成功率加成
      if (meta.successRate > 0.8) {
        score += 5;
      }
    }

    return { score, matchReasons };
  }
}

// ============================================================================
// 便捷函数
// ============================================================================

/**
 * 创建工具选择器
 */
export function createToolSelector(): ToolSelector {
  return new ToolSelector();
}

/**
 * 快速选择工具
 */
export function selectRelevantTools(
  query: string,
  allTools: Tool[],
  maxTools: number = 25
): Tool[] {
  const selector = new ToolSelector();
  selector.registerTools(allTools);
  const result = selector.selectTools(query, { maxTools });
  return result.tools;
}

/**
 * 获取工具分类
 */
export function categorizeTools(tools: Tool[]): Map<ToolCategory, Tool[]> {
  const result = new Map<ToolCategory, Tool[]>();
  const selector = new ToolSelector();

  for (const tool of tools) {
    selector.registerTool(tool);
  }

  const stats = selector.getStats();
  for (const [category] of stats.byCategory) {
    result.set(category, selector.selectByCategory([category], 100));
  }

  return result;
}
