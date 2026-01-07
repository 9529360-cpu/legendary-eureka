/**
 * ToolDiscovery - 工具动态发现器 v4.1
 *
 * 基于语义标签匹配自动发现合适的工具
 *
 * 核心特性：
 * 1. 语义标签索引
 * 2. 意图匹配算法
 * 3. 工具推荐排序
 * 4. 缓存优化
 *
 * @module agent/ToolDiscovery
 */

import { Tool } from "./types/tool";
import { ToolRegistry } from "./registry";

// ========== 类型定义 ==========

/**
 * 语义标签
 */
export interface SemanticTag {
  /** 标签名 */
  name: string;

  /** 权重 (0-1) */
  weight: number;

  /** 类别 */
  category?: "action" | "entity" | "modifier" | "domain";
}

/**
 * 工具元数据（扩展）
 */
export interface ToolMetadata {
  /** 工具名 */
  name: string;

  /** 描述 */
  description: string;

  /** 分类 */
  category: string;

  /** 语义标签 */
  semanticTags: SemanticTag[];

  /** 使用频率 */
  usageCount?: number;

  /** 成功率 */
  successRate?: number;

  /** 平均耗时 (ms) */
  avgDuration?: number;
}

/**
 * 意图原子（从用户意图提取）
 */
export interface IntentAtom {
  /** 动作词 */
  action?: string;

  /** 目标实体 */
  entity?: string;

  /** 修饰词 */
  modifiers?: string[];

  /** 领域 */
  domain?: string;

  /** 原始文本 */
  rawText?: string;
}

/**
 * 匹配结果
 */
export interface ToolMatch {
  /** 工具 */
  tool: Tool;

  /** 匹配分数 (0-1) */
  score: number;

  /** 匹配的标签 */
  matchedTags: string[];

  /** 推荐理由 */
  reason: string;
}

/**
 * 发现选项
 */
export interface DiscoveryOptions {
  /** 最大返回数量 */
  limit?: number;

  /** 最低匹配分数 */
  minScore?: number;

  /** 是否考虑使用统计 */
  useStats?: boolean;

  /** 限定分类 */
  categories?: string[];
}

// ========== 预定义语义标签库 ==========

/**
 * 动作词映射
 */
const ACTION_SYNONYMS: Record<string, string[]> = {
  read: ["读取", "获取", "查看", "读", "取", "看", "get", "fetch", "load"],
  write: ["写入", "填写", "输入", "写", "填", "set", "put", "input"],
  create: ["创建", "新建", "添加", "生成", "建", "create", "add", "new", "generate"],
  delete: ["删除", "移除", "清除", "清空", "删", "remove", "clear", "erase"],
  update: ["更新", "修改", "编辑", "改", "update", "modify", "edit", "change"],
  format: ["格式化", "美化", "样式", "格式", "format", "style", "beautify"],
  calculate: ["计算", "求和", "统计", "算", "calculate", "compute", "sum", "count"],
  analyze: ["分析", "检查", "诊断", "analyze", "check", "diagnose", "inspect"],
  filter: ["筛选", "过滤", "搜索", "查找", "filter", "search", "find", "query"],
  sort: ["排序", "排列", "sort", "order", "arrange"],
  merge: ["合并", "组合", "连接", "merge", "combine", "join", "concat"],
  split: ["拆分", "分割", "分离", "split", "separate", "divide"],
  copy: ["复制", "拷贝", "copy", "duplicate", "clone"],
  move: ["移动", "转移", "move", "transfer", "relocate"],
  chart: ["图表", "可视化", "绘制", "chart", "visualize", "plot", "draw"],
};

/**
 * 实体词映射
 */
const ENTITY_SYNONYMS: Record<string, string[]> = {
  cell: ["单元格", "格子", "cell", "cells"],
  range: ["区域", "范围", "range", "area", "region"],
  row: ["行", "row", "rows", "line"],
  column: ["列", "column", "columns", "col"],
  sheet: ["工作表", "表格", "sheet", "worksheet", "tab"],
  workbook: ["工作簿", "文件", "workbook", "file", "document"],
  formula: ["公式", "函数", "formula", "function"],
  value: ["值", "数据", "内容", "value", "data", "content"],
  format: ["格式", "样式", "format", "style"],
  chart: ["图表", "图", "chart", "graph", "diagram"],
  table: ["表", "表格", "table"],
  filter: ["筛选器", "过滤器", "filter"],
  sort: ["排序", "sort", "order"],
  color: ["颜色", "背景色", "color", "background"],
  border: ["边框", "border"],
  font: ["字体", "font", "text"],
};

// ========== ToolDiscovery 类 ==========

/**
 * 工具发现器
 */
export class ToolDiscovery {
  private toolRegistry: ToolRegistry;
  private metadataCache: Map<string, ToolMetadata>;
  private tagIndex: Map<string, Set<string>>; // tag -> tool names
  private lastBuildTime: number;

  constructor(toolRegistry: ToolRegistry) {
    this.toolRegistry = toolRegistry;
    this.metadataCache = new Map();
    this.tagIndex = new Map();
    this.lastBuildTime = 0;
  }

  /**
   * 初始化索引
   */
  async initialize(): Promise<void> {
    await this.buildIndex();
  }

  /**
   * 构建语义索引
   */
  async buildIndex(): Promise<void> {
    const tools = this.toolRegistry.getAll();

    for (const tool of tools) {
      // 从工具描述提取语义标签
      const semanticTags = this.extractSemanticTags(tool);

      const metadata: ToolMetadata = {
        name: tool.name,
        description: tool.description,
        category: tool.category,
        semanticTags,
      };

      this.metadataCache.set(tool.name, metadata);

      // 构建反向索引
      for (const tag of semanticTags) {
        const normalizedTag = this.normalizeTag(tag.name);
        if (!this.tagIndex.has(normalizedTag)) {
          this.tagIndex.set(normalizedTag, new Set());
        }
        this.tagIndex.get(normalizedTag)!.add(tool.name);
      }
    }

    this.lastBuildTime = Date.now();
    console.log(`[ToolDiscovery] 索引构建完成: ${tools.length} 工具, ${this.tagIndex.size} 标签`);
  }

  /**
   * 从工具描述提取语义标签
   */
  private extractSemanticTags(tool: Tool): SemanticTag[] {
    const tags: SemanticTag[] = [];
    const description = tool.description.toLowerCase();
    const name = tool.name.toLowerCase();

    // 提取动作标签
    for (const [action, synonyms] of Object.entries(ACTION_SYNONYMS)) {
      for (const synonym of synonyms) {
        if (description.includes(synonym.toLowerCase()) || name.includes(synonym.toLowerCase())) {
          tags.push({
            name: action,
            weight: 0.8,
            category: "action",
          });
          break;
        }
      }
    }

    // 提取实体标签
    for (const [entity, synonyms] of Object.entries(ENTITY_SYNONYMS)) {
      for (const synonym of synonyms) {
        if (description.includes(synonym.toLowerCase()) || name.includes(synonym.toLowerCase())) {
          tags.push({
            name: entity,
            weight: 0.7,
            category: "entity",
          });
          break;
        }
      }
    }

    // 添加分类标签
    if (tool.category) {
      tags.push({
        name: tool.category,
        weight: 0.5,
        category: "domain",
      });
    }

    return tags;
  }

  /**
   * 发现匹配的工具
   */
  discover(intent: IntentAtom, options: DiscoveryOptions = {}): ToolMatch[] {
    const { limit = 5, minScore = 0.3, useStats = true, categories } = options;

    const matches: ToolMatch[] = [];
    const intentTags = this.extractIntentTags(intent);

    for (const [toolName, metadata] of this.metadataCache) {
      // 分类过滤
      if (categories && categories.length > 0) {
        if (!categories.includes(metadata.category)) {
          continue;
        }
      }

      // 计算匹配分数
      const { score, matchedTags } = this.calculateMatchScore(intentTags, metadata);

      if (score >= minScore) {
        const tool = this.toolRegistry.get(toolName);
        if (tool) {
          // 统计加权
          let finalScore = score;
          if (useStats && metadata.successRate !== undefined) {
            finalScore = score * 0.7 + metadata.successRate * 0.3;
          }

          matches.push({
            tool,
            score: finalScore,
            matchedTags,
            reason: this.generateReason(matchedTags, intent),
          });
        }
      }
    }

    // 排序
    matches.sort((a, b) => b.score - a.score);

    return matches.slice(0, limit);
  }

  /**
   * 从意图提取标签
   */
  private extractIntentTags(intent: IntentAtom): SemanticTag[] {
    const tags: SemanticTag[] = [];

    // 动作标签
    if (intent.action) {
      const normalizedAction = this.normalizeAction(intent.action);
      if (normalizedAction) {
        tags.push({
          name: normalizedAction,
          weight: 1.0,
          category: "action",
        });
      }
    }

    // 实体标签
    if (intent.entity) {
      const normalizedEntity = this.normalizeEntity(intent.entity);
      if (normalizedEntity) {
        tags.push({
          name: normalizedEntity,
          weight: 0.9,
          category: "entity",
        });
      }
    }

    // 修饰词
    if (intent.modifiers) {
      for (const mod of intent.modifiers) {
        tags.push({
          name: mod,
          weight: 0.5,
          category: "modifier",
        });
      }
    }

    // 领域
    if (intent.domain) {
      tags.push({
        name: intent.domain,
        weight: 0.6,
        category: "domain",
      });
    }

    // 从原始文本提取
    if (intent.rawText) {
      const extractedTags = this.extractFromRawText(intent.rawText);
      tags.push(...extractedTags);
    }

    return tags;
  }

  /**
   * 从原始文本提取标签
   */
  private extractFromRawText(text: string): SemanticTag[] {
    const tags: SemanticTag[] = [];
    const lowerText = text.toLowerCase();

    // 检查动作词
    for (const [action, synonyms] of Object.entries(ACTION_SYNONYMS)) {
      for (const synonym of synonyms) {
        if (lowerText.includes(synonym.toLowerCase())) {
          tags.push({
            name: action,
            weight: 0.6,
            category: "action",
          });
          break;
        }
      }
    }

    // 检查实体词
    for (const [entity, synonyms] of Object.entries(ENTITY_SYNONYMS)) {
      for (const synonym of synonyms) {
        if (lowerText.includes(synonym.toLowerCase())) {
          tags.push({
            name: entity,
            weight: 0.5,
            category: "entity",
          });
          break;
        }
      }
    }

    return tags;
  }

  /**
   * 规范化动作词
   */
  private normalizeAction(action: string): string | null {
    const lowerAction = action.toLowerCase();
    for (const [normalized, synonyms] of Object.entries(ACTION_SYNONYMS)) {
      if (normalized === lowerAction || synonyms.some((s) => s.toLowerCase() === lowerAction)) {
        return normalized;
      }
    }
    return action;
  }

  /**
   * 规范化实体词
   */
  private normalizeEntity(entity: string): string | null {
    const lowerEntity = entity.toLowerCase();
    for (const [normalized, synonyms] of Object.entries(ENTITY_SYNONYMS)) {
      if (normalized === lowerEntity || synonyms.some((s) => s.toLowerCase() === lowerEntity)) {
        return normalized;
      }
    }
    return entity;
  }

  /**
   * 规范化标签
   */
  private normalizeTag(tag: string): string {
    return tag.toLowerCase().trim();
  }

  /**
   * 计算匹配分数
   */
  private calculateMatchScore(
    intentTags: SemanticTag[],
    metadata: ToolMetadata
  ): { score: number; matchedTags: string[] } {
    if (intentTags.length === 0) {
      return { score: 0, matchedTags: [] };
    }

    const matchedTags: string[] = [];
    let totalWeight = 0;
    let matchedWeight = 0;

    for (const intentTag of intentTags) {
      totalWeight += intentTag.weight;

      const toolTag = metadata.semanticTags.find(
        (t) => this.normalizeTag(t.name) === this.normalizeTag(intentTag.name)
      );

      if (toolTag) {
        matchedTags.push(intentTag.name);
        matchedWeight += intentTag.weight * toolTag.weight;
      }
    }

    const score = totalWeight > 0 ? matchedWeight / totalWeight : 0;
    return { score, matchedTags };
  }

  /**
   * 生成推荐理由
   */
  private generateReason(matchedTags: string[], intent: IntentAtom): string {
    if (matchedTags.length === 0) {
      return "通用匹配";
    }

    const tagNames = matchedTags.slice(0, 3).join(", ");
    return `匹配: ${tagNames}`;
  }

  /**
   * 更新工具使用统计
   */
  updateStats(toolName: string, success: boolean, duration: number): void {
    const metadata = this.metadataCache.get(toolName);
    if (!metadata) return;

    // 更新使用次数
    metadata.usageCount = (metadata.usageCount || 0) + 1;

    // 更新成功率（指数平滑）
    const alpha = 0.2;
    const currentSuccessRate = metadata.successRate ?? 0.5;
    metadata.successRate = alpha * (success ? 1 : 0) + (1 - alpha) * currentSuccessRate;

    // 更新平均耗时（指数平滑）
    const currentAvgDuration = metadata.avgDuration ?? duration;
    metadata.avgDuration = alpha * duration + (1 - alpha) * currentAvgDuration;
  }

  /**
   * 搜索工具（基于关键词）
   */
  search(query: string, options: DiscoveryOptions = {}): ToolMatch[] {
    const intent: IntentAtom = {
      rawText: query,
    };
    return this.discover(intent, options);
  }

  /**
   * 获取分类下的所有工具
   */
  getByCategory(category: string): Tool[] {
    const tools: Tool[] = [];

    for (const [toolName, metadata] of this.metadataCache) {
      if (metadata.category === category) {
        const tool = this.toolRegistry.get(toolName);
        if (tool) {
          tools.push(tool);
        }
      }
    }

    return tools;
  }

  /**
   * 获取热门工具
   */
  getPopular(limit: number = 10): ToolMetadata[] {
    const sorted = Array.from(this.metadataCache.values())
      .filter((m) => m.usageCount !== undefined && m.usageCount > 0)
      .sort((a, b) => (b.usageCount || 0) - (a.usageCount || 0));

    return sorted.slice(0, limit);
  }

  /**
   * 获取统计信息
   */
  getStats(): {
    totalTools: number;
    totalTags: number;
    avgTagsPerTool: number;
    categories: string[];
  } {
    const totalTools = this.metadataCache.size;
    const totalTags = this.tagIndex.size;
    const categories = new Set<string>();
    let totalTagCount = 0;

    for (const metadata of this.metadataCache.values()) {
      categories.add(metadata.category);
      totalTagCount += metadata.semanticTags.length;
    }

    return {
      totalTools,
      totalTags,
      avgTagsPerTool: totalTools > 0 ? totalTagCount / totalTools : 0,
      categories: Array.from(categories),
    };
  }
}

// ========== 工厂函数 ==========

/**
 * 创建工具发现器
 */
export async function createToolDiscovery(toolRegistry: ToolRegistry): Promise<ToolDiscovery> {
  const discovery = new ToolDiscovery(toolRegistry);
  await discovery.initialize();
  return discovery;
}

export default ToolDiscovery;
