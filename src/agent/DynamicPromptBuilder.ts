/**
 * DynamicPromptBuilder - 动态提示构建器 v4.2
 *
 * 根据上下文动态生成最优 System Prompt
 *
 * 核心特性：
 * 1. 工具相关性过滤
 * 2. 上下文压缩
 * 3. 经验注入
 * 4. Token 预算管理
 *
 * @module agent/DynamicPromptBuilder
 */

import { Tool } from "./types/tool";
import { ToolRegistry } from "./registry";
import { ToolDiscovery, IntentAtom, ToolMatch } from "./ToolDiscovery";
import { PersistentMemory, StoredEpisode } from "./memory";

// ========== 类型定义 ==========

/**
 * 提示构建配置
 */
export interface PromptBuilderConfig {
  /** 最大 Token 预算 */
  maxTokens?: number;

  /** 工具描述最大数量 */
  maxToolDescriptions?: number;

  /** 经验示例最大数量 */
  maxExamples?: number;

  /** 是否包含工具描述 */
  includeToolDescriptions?: boolean;

  /** 是否包含经验示例 */
  includeExamples?: boolean;

  /** 是否包含安全规则 */
  includeSafetyRules?: boolean;

  /** 语言 */
  language?: "zh" | "en";
}

/**
 * 构建上下文
 */
export interface PromptBuildContext {
  /** 用户消息 */
  userMessage: string;

  /** 意图原子 */
  intentAtom?: IntentAtom;

  /** 工作簿上下文 */
  workbookContext?: {
    sheets: string[];
    activeSheet: string;
    usedRanges: string[];
    dataPreview?: string;
  };

  /** 对话历史 */
  conversationHistory?: Array<{ role: "user" | "assistant"; content: string }>;
}

/**
 * 构建结果
 */
export interface PromptBuildResult {
  /** 系统提示 */
  systemPrompt: string;

  /** 预估 Token 数 */
  estimatedTokens: number;

  /** 包含的工具数 */
  includedTools: number;

  /** 包含的示例数 */
  includedExamples: number;

  /** 压缩比 */
  compressionRatio: number;
}

// ========== 默认配置 ==========

const DEFAULT_CONFIG: Required<PromptBuilderConfig> = {
  maxTokens: 4000,
  maxToolDescriptions: 15,
  maxExamples: 3,
  includeToolDescriptions: true,
  includeExamples: true,
  includeSafetyRules: true,
  language: "zh",
};

// ========== 提示模板 ==========

const SYSTEM_PROMPT_TEMPLATES = {
  zh: {
    base: `你是一个专业的 Excel 智能助手，帮助用户完成各种 Excel 操作任务。

## 你的能力
- 理解用户的自然语言指令
- 将指令转换为具体的 Excel 操作
- 提供清晰的执行反馈

## 工作原则
1. 准确理解用户意图
2. 选择最合适的操作方式
3. 确保数据安全和准确性
4. 提供友好的交互体验`,

    safety: `
## 安全规则
- 不执行可能导致数据丢失的危险操作（除非用户明确确认）
- 大批量操作前先确认影响范围
- 保持操作的可追溯性
- 遇到异常及时报告`,

    toolsHeader: `
## 可用工具`,

    examplesHeader: `
## 成功经验`,

    contextHeader: `
## 当前上下文`,
  },
  en: {
    base: `You are a professional Excel AI assistant, helping users complete various Excel tasks.

## Your Capabilities
- Understand natural language instructions
- Convert instructions to specific Excel operations
- Provide clear execution feedback

## Working Principles
1. Accurately understand user intent
2. Choose the most appropriate operation
3. Ensure data safety and accuracy
4. Provide friendly interaction experience`,

    safety: `
## Safety Rules
- Do not execute dangerous operations that may cause data loss (unless user explicitly confirms)
- Confirm impact scope before batch operations
- Maintain operation traceability
- Report exceptions promptly`,

    toolsHeader: `
## Available Tools`,

    examplesHeader: `
## Successful Experiences`,

    contextHeader: `
## Current Context`,
  },
};

// ========== DynamicPromptBuilder 类 ==========

/**
 * 动态提示构建器
 */
export class DynamicPromptBuilder {
  private toolRegistry: ToolRegistry;
  private toolDiscovery: ToolDiscovery | null = null;
  private memory: PersistentMemory | null = null;
  private config: Required<PromptBuilderConfig>;

  constructor(
    toolRegistry: ToolRegistry,
    config: PromptBuilderConfig = {}
  ) {
    this.toolRegistry = toolRegistry;
    this.config = { ...DEFAULT_CONFIG, ...config };
  }

  /**
   * 设置工具发现器
   */
  setToolDiscovery(discovery: ToolDiscovery): void {
    this.toolDiscovery = discovery;
  }

  /**
   * 设置持久化内存
   */
  setMemory(memory: PersistentMemory): void {
    this.memory = memory;
  }

  /**
   * 构建动态提示
   */
  async build(context: PromptBuildContext): Promise<PromptBuildResult> {
    const templates = SYSTEM_PROMPT_TEMPLATES[this.config.language];
    const sections: string[] = [];
    let estimatedTokens = 0;
    let includedTools = 0;
    let includedExamples = 0;

    // 1. 基础提示
    sections.push(templates.base);
    estimatedTokens += this.estimateTokens(templates.base);

    // 2. 安全规则
    if (this.config.includeSafetyRules) {
      sections.push(templates.safety);
      estimatedTokens += this.estimateTokens(templates.safety);
    }

    // 3. 工具描述（动态选择）
    if (this.config.includeToolDescriptions) {
      const toolSection = await this.buildToolSection(context, templates);
      if (toolSection.content) {
        sections.push(toolSection.content);
        estimatedTokens += toolSection.tokens;
        includedTools = toolSection.count;
      }
    }

    // 4. 经验示例
    if (this.config.includeExamples && this.memory) {
      const exampleSection = await this.buildExampleSection(context, templates);
      if (exampleSection.content) {
        sections.push(exampleSection.content);
        estimatedTokens += exampleSection.tokens;
        includedExamples = exampleSection.count;
      }
    }

    // 5. 当前上下文
    const contextSection = this.buildContextSection(context, templates);
    if (contextSection.content) {
      sections.push(contextSection.content);
      estimatedTokens += contextSection.tokens;
    }

    const systemPrompt = sections.join("\n");
    const fullTokens = this.estimateTokens(
      this.buildFullPrompt(this.toolRegistry.getAll())
    );
    const compressionRatio = fullTokens > 0 ? estimatedTokens / fullTokens : 1;

    return {
      systemPrompt,
      estimatedTokens,
      includedTools,
      includedExamples,
      compressionRatio,
    };
  }

  /**
   * 构建工具描述部分
   */
  private async buildToolSection(
    context: PromptBuildContext,
    templates: typeof SYSTEM_PROMPT_TEMPLATES["zh"]
  ): Promise<{ content: string; tokens: number; count: number }> {
    let relevantTools: Tool[];

    if (this.toolDiscovery && context.intentAtom) {
      // 使用工具发现器获取相关工具
      const matches = this.toolDiscovery.discover(context.intentAtom, {
        limit: this.config.maxToolDescriptions,
        minScore: 0.1,
      });
      relevantTools = matches.map((m) => m.tool);
    } else if (context.userMessage) {
      // 根据关键词简单过滤
      relevantTools = this.filterToolsByKeywords(context.userMessage);
    } else {
      // 返回所有工具（受限）
      relevantTools = this.toolRegistry.getAll().slice(0, this.config.maxToolDescriptions);
    }

    if (relevantTools.length === 0) {
      return { content: "", tokens: 0, count: 0 };
    }

    const lines = [templates.toolsHeader];
    for (const tool of relevantTools) {
      lines.push(`- **${tool.name}**: ${tool.description}`);
    }

    const content = lines.join("\n");
    return {
      content,
      tokens: this.estimateTokens(content),
      count: relevantTools.length,
    };
  }

  /**
   * 构建经验示例部分
   */
  private async buildExampleSection(
    context: PromptBuildContext,
    templates: typeof SYSTEM_PROMPT_TEMPLATES["zh"]
  ): Promise<{ content: string; tokens: number; count: number }> {
    if (!this.memory) {
      return { content: "", tokens: 0, count: 0 };
    }

    try {
      const episodes = await this.memory.getSimilarEpisodes(
        context.userMessage,
        this.config.maxExamples
      );

      if (episodes.length === 0) {
        return { content: "", tokens: 0, count: 0 };
      }

      const lines = [templates.examplesHeader];
      for (const ep of episodes) {
        if (ep.result === "success") {
          lines.push(`- 用户: "${ep.intent}" → 使用工具: ${ep.toolsUsed.join(", ")}`);
        }
      }

      const content = lines.join("\n");
      return {
        content,
        tokens: this.estimateTokens(content),
        count: episodes.length,
      };
    } catch (error) {
      console.warn("[DynamicPromptBuilder] 获取经验示例失败", error);
      return { content: "", tokens: 0, count: 0 };
    }
  }

  /**
   * 构建上下文部分
   */
  private buildContextSection(
    context: PromptBuildContext,
    templates: typeof SYSTEM_PROMPT_TEMPLATES["zh"]
  ): { content: string; tokens: number } {
    if (!context.workbookContext) {
      return { content: "", tokens: 0 };
    }

    const wb = context.workbookContext;
    const lines = [templates.contextHeader];

    lines.push(`- 当前工作表: ${wb.activeSheet}`);
    if (wb.sheets.length > 0) {
      lines.push(`- 所有工作表: ${wb.sheets.join(", ")}`);
    }
    if (wb.usedRanges.length > 0) {
      lines.push(`- 使用区域: ${wb.usedRanges.join(", ")}`);
    }
    if (wb.dataPreview) {
      lines.push(`- 数据预览:\n\`\`\`\n${wb.dataPreview}\n\`\`\``);
    }

    const content = lines.join("\n");
    return {
      content,
      tokens: this.estimateTokens(content),
    };
  }

  /**
   * 根据关键词过滤工具
   */
  private filterToolsByKeywords(userMessage: string): Tool[] {
    const keywords = userMessage.toLowerCase().split(/\s+/);
    const allTools = this.toolRegistry.getAll();

    // 计算每个工具的相关性分数
    const scored = allTools.map((tool) => {
      const toolText = `${tool.name} ${tool.description}`.toLowerCase();
      let score = 0;
      for (const kw of keywords) {
        if (toolText.includes(kw)) {
          score++;
        }
      }
      return { tool, score };
    });

    // 按分数排序，取前 N 个
    return scored
      .filter((s) => s.score > 0)
      .sort((a, b) => b.score - a.score)
      .slice(0, this.config.maxToolDescriptions)
      .map((s) => s.tool);
  }

  /**
   * 构建完整提示（用于计算压缩比）
   */
  private buildFullPrompt(tools: Tool[]): string {
    const lines = [SYSTEM_PROMPT_TEMPLATES[this.config.language].base];

    for (const tool of tools) {
      lines.push(`- **${tool.name}**: ${tool.description}`);
      if (tool.parameters.length > 0) {
        for (const param of tool.parameters) {
          lines.push(`  - ${param.name}: ${param.description}`);
        }
      }
    }

    return lines.join("\n");
  }

  /**
   * 估算 Token 数
   */
  private estimateTokens(text: string): number {
    // 简单估算：中文约 2 字符/token，英文约 4 字符/token
    const chineseChars = (text.match(/[\u4e00-\u9fa5]/g) || []).length;
    const otherChars = text.length - chineseChars;
    return Math.ceil(chineseChars / 1.5 + otherChars / 4);
  }

  /**
   * 压缩对话历史
   */
  compressHistory(
    history: Array<{ role: "user" | "assistant"; content: string }>,
    maxTokens: number
  ): Array<{ role: "user" | "assistant"; content: string }> {
    if (history.length === 0) return [];

    const result: Array<{ role: "user" | "assistant"; content: string }> = [];
    let totalTokens = 0;

    // 从最近的消息开始，向前添加
    for (let i = history.length - 1; i >= 0; i--) {
      const msg = history[i];
      const msgTokens = this.estimateTokens(msg.content);

      if (totalTokens + msgTokens > maxTokens) {
        // 如果是最后一条消息（最新），截断但保留
        if (result.length === 0) {
          const truncated = this.truncateText(msg.content, maxTokens);
          result.unshift({ ...msg, content: truncated });
        }
        break;
      }

      result.unshift(msg);
      totalTokens += msgTokens;
    }

    return result;
  }

  /**
   * 截断文本
   */
  private truncateText(text: string, maxTokens: number): string {
    const maxChars = maxTokens * 3; // 估算
    if (text.length <= maxChars) return text;
    return text.substring(0, maxChars - 3) + "...";
  }
}

// ========== 工厂函数 ==========

/**
 * 创建动态提示构建器
 */
export function createDynamicPromptBuilder(
  toolRegistry: ToolRegistry,
  config?: PromptBuilderConfig
): DynamicPromptBuilder {
  return new DynamicPromptBuilder(toolRegistry, config);
}

export default DynamicPromptBuilder;
