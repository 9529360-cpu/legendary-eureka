/**
 * SystemMessageBuilder - 分层 System Message 构建器
 *
 * 基于 ai-agents-for-beginners 第6课 Building Trustworthy Agents 的学习:
 * 实现 Meta + Basic + Optimization 三层结构的系统提示
 *
 * 架构:
 * ┌─────────────────────────────────────────┐
 * │           Meta Prompt                    │
 * │  (公司政策、安全规则、角色定义)          │
 * ├─────────────────────────────────────────┤
 * │           Basic Prompt                   │
 * │  (任务指令、输出格式、约束条件)          │
 * ├─────────────────────────────────────────┤
 * │        LLM Optimization                  │
 * │  (CoT、Few-shot、结构化输出)            │
 * └─────────────────────────────────────────┘
 *
 * @version 1.0.0
 * @see docs/AI_AGENTS_FOR_BEGINNERS_LEARNING.md
 */

import type { Tool } from "./AgentCore";

// ============================================================================
// 类型定义
// ============================================================================

/**
 * 系统消息层
 */
export interface SystemMessageLayer {
  /** 层名称 */
  name: string;
  /** 优先级 (越小越优先) */
  priority: number;
  /** 内容 */
  content: string;
  /** 是否启用 */
  enabled: boolean;
  /** 条件 (可选) */
  condition?: () => boolean;
}

/**
 * 构建配置
 */
export interface BuilderConfig {
  /** 角色名称 */
  roleName: string;
  /** 公司/产品名称 */
  productName: string;
  /** 主要职责 */
  responsibilities: string[];
  /** 安全规则 */
  securityRules: string[];
  /** 输出语言 */
  outputLanguage: "zh-CN" | "en-US";
  /** 是否启用 Chain-of-Thought */
  enableCoT: boolean;
  /** 是否启用结构化输出 */
  enableStructuredOutput: boolean;
  /** 最大工具数 */
  maxToolsInPrompt: number;
}

/**
 * 工具描述格式
 */
export interface ToolDescription {
  name: string;
  description: string;
  parameters: string;
  example?: string;
}

// ============================================================================
// 默认配置
// ============================================================================

const DEFAULT_CONFIG: BuilderConfig = {
  roleName: "Excel 智能助手",
  productName: "Excel Copilot",
  responsibilities: [
    "帮助用户处理 Excel 电子表格任务",
    "理解自然语言指令并转换为 Excel 操作",
    "提供数据分析和可视化建议",
  ],
  securityRules: [
    "不执行可能导致数据丢失的操作，除非用户明确确认",
    "不访问或修改用户未授权的数据",
    "对于大规模修改操作，必须先获得用户确认",
    "保护用户隐私，不记录敏感数据",
  ],
  outputLanguage: "zh-CN",
  enableCoT: true,
  enableStructuredOutput: true,
  maxToolsInPrompt: 25,
};

// ============================================================================
// System Message 构建器
// ============================================================================

/**
 * 系统消息构建器
 *
 * 构建结构化、分层的系统提示
 */
export class SystemMessageBuilder {
  private config: BuilderConfig;
  private layers: SystemMessageLayer[] = [];
  private readonly MODULE_NAME = "SystemMessageBuilder";

  constructor(config: Partial<BuilderConfig> = {}) {
    this.config = { ...DEFAULT_CONFIG, ...config };
    this.initializeDefaultLayers();
  }

  /**
   * 构建完整的系统消息
   */
  build(tools?: Tool[], context?: Record<string, unknown>): string {
    const activeLayers = this.getActiveLayers();
    const sections: string[] = [];

    for (const layer of activeLayers) {
      let content = layer.content;

      // 替换模板变量
      content = this.replaceTemplates(content, context);

      sections.push(content);
    }

    // 添加工具描述
    if (tools && tools.length > 0) {
      sections.push(this.buildToolSection(tools));
    }

    return sections.join("\n\n");
  }

  /**
   * 构建精简版系统消息（用于 token 受限场景）
   */
  buildCompact(tools?: Tool[]): string {
    const essential = this.layers
      .filter((l) => l.priority <= 2 && l.enabled)
      .sort((a, b) => a.priority - b.priority);

    const sections = essential.map((l) => l.content);

    // 只添加最常用的工具
    if (tools) {
      const limitedTools = tools.slice(0, 15);
      sections.push(this.buildCompactToolSection(limitedTools));
    }

    return sections.join("\n\n");
  }

  /**
   * 添加自定义层
   */
  addLayer(layer: SystemMessageLayer): this {
    this.layers.push(layer);
    this.layers.sort((a, b) => a.priority - b.priority);
    return this;
  }

  /**
   * 更新配置
   */
  updateConfig(config: Partial<BuilderConfig>): this {
    this.config = { ...this.config, ...config };
    this.initializeDefaultLayers(); // 重新生成默认层
    return this;
  }

  /**
   * 启用/禁用层
   */
  toggleLayer(name: string, enabled: boolean): this {
    const layer = this.layers.find((l) => l.name === name);
    if (layer) {
      layer.enabled = enabled;
    }
    return this;
  }

  /**
   * 获取当前层列表
   */
  getLayers(): SystemMessageLayer[] {
    return [...this.layers];
  }

  // ============================================================================
  // 私有方法
  // ============================================================================

  /**
   * 初始化默认层
   */
  private initializeDefaultLayers(): void {
    this.layers = [
      // 层 1: Meta Prompt (最高优先级)
      {
        name: "meta",
        priority: 1,
        enabled: true,
        content: this.buildMetaPrompt(),
      },
      // 层 2: Safety Rules
      {
        name: "safety",
        priority: 2,
        enabled: true,
        content: this.buildSafetyRules(),
      },
      // 层 3: Basic Prompt
      {
        name: "basic",
        priority: 3,
        enabled: true,
        content: this.buildBasicPrompt(),
      },
      // 层 4: Output Format
      {
        name: "output",
        priority: 4,
        enabled: this.config.enableStructuredOutput,
        content: this.buildOutputFormat(),
      },
      // 层 5: Chain-of-Thought
      {
        name: "cot",
        priority: 5,
        enabled: this.config.enableCoT,
        content: this.buildCoTInstructions(),
      },
      // 层 6: Adaptive Behavior
      {
        name: "adaptive",
        priority: 6,
        enabled: true,
        content: this.buildAdaptiveBehavior(),
      },
    ];
  }

  /**
   * 构建 Meta Prompt
   */
  private buildMetaPrompt(): string {
    const isZh = this.config.outputLanguage === "zh-CN";

    if (isZh) {
      return `# 角色定义

你是 **${this.config.roleName}**，${this.config.productName} 的核心 AI 助手。

## 核心职责
${this.config.responsibilities.map((r) => `- ${r}`).join("\n")}

## 行为准则
- 始终以用户利益为先
- 保持专业、友好、高效的沟通风格
- 诚实承认能力边界，不做超出能力范围的承诺
- 使用中文回复，专业术语可保留英文`;
    }

    return `# Role Definition

You are **${this.config.roleName}**, the core AI assistant of ${this.config.productName}.

## Core Responsibilities
${this.config.responsibilities.map((r) => `- ${r}`).join("\n")}

## Behavioral Guidelines
- Always prioritize user interests
- Maintain professional, friendly, and efficient communication
- Honestly acknowledge capability boundaries
- Respond in a clear and concise manner`;
  }

  /**
   * 构建安全规则
   */
  private buildSafetyRules(): string {
    const isZh = this.config.outputLanguage === "zh-CN";

    if (isZh) {
      return `## 安全规则 (必须遵守)

${this.config.securityRules.map((r, i) => `${i + 1}. ${r}`).join("\n")}

### 高风险操作识别
以下操作需要用户确认：
- 删除数据 (delete, clear, remove)
- 批量修改 (超过 100 个单元格)
- 格式重置或清空
- 工作表结构变更

### 错误处理
- 遇到错误时，清晰解释问题
- 提供可能的解决方案
- 不要尝试未经验证的修复`;
    }

    return `## Security Rules (MUST follow)

${this.config.securityRules.map((r, i) => `${i + 1}. ${r}`).join("\n")}

### High-Risk Operation Detection
The following operations require user confirmation:
- Data deletion (delete, clear, remove)
- Bulk modifications (>100 cells)
- Format reset or clearing
- Worksheet structure changes

### Error Handling
- Clearly explain problems when errors occur
- Provide possible solutions
- Do not attempt unverified fixes`;
  }

  /**
   * 构建 Basic Prompt
   */
  private buildBasicPrompt(): string {
    const isZh = this.config.outputLanguage === "zh-CN";

    if (isZh) {
      return `## 任务处理流程

### 1. 理解阶段
- 仔细分析用户请求
- 识别关键意图和参数
- 如有歧义，提出澄清问题

### 2. 规划阶段
- 将复杂任务分解为子步骤
- 选择最合适的工具
- 评估操作风险

### 3. 执行阶段
- 按计划执行每个步骤
- 验证每步结果
- 必要时调整计划

### 4. 反馈阶段
- 总结完成的操作
- 报告遇到的问题
- 提供后续建议`;
    }

    return `## Task Processing Flow

### 1. Understanding Phase
- Carefully analyze user request
- Identify key intent and parameters
- Ask clarifying questions if ambiguous

### 2. Planning Phase
- Break complex tasks into subtasks
- Select the most appropriate tools
- Assess operation risks

### 3. Execution Phase
- Execute each step as planned
- Verify results of each step
- Adjust plan when necessary

### 4. Feedback Phase
- Summarize completed operations
- Report any issues encountered
- Provide follow-up suggestions`;
  }

  /**
   * 构建输出格式指令
   */
  private buildOutputFormat(): string {
    const isZh = this.config.outputLanguage === "zh-CN";

    if (isZh) {
      return `## 输出格式

你的响应必须是有效的 JSON，遵循以下结构：

\`\`\`json
{
  "operation": "execute" | "confirm" | "clarify" | "reject" | "complete",
  "mainTask": "任务描述",
  "steps": [
    {
      "stepNumber": 1,
      "description": "步骤描述",
      "toolCall": {
        "name": "工具名称",
        "parameters": { ... }
      },
      "rationale": "为什么需要这个步骤"
    }
  ],
  "riskLevel": "low" | "medium" | "high" | "critical",
  "userMessage": "给用户的消息（可选）"
}
\`\`\`

### 操作类型说明
- \`execute\`: 直接执行，低风险操作
- \`confirm\`: 需要用户确认，中高风险操作
- \`clarify\`: 需要更多信息
- \`reject\`: 无法执行或不安全
- \`complete\`: 任务已完成`;
    }

    return `## Output Format

Your response must be valid JSON following this structure:

\`\`\`json
{
  "operation": "execute" | "confirm" | "clarify" | "reject" | "complete",
  "mainTask": "Task description",
  "steps": [
    {
      "stepNumber": 1,
      "description": "Step description",
      "toolCall": {
        "name": "tool_name",
        "parameters": { ... }
      },
      "rationale": "Why this step is needed"
    }
  ],
  "riskLevel": "low" | "medium" | "high" | "critical",
  "userMessage": "Message for user (optional)"
}
\`\`\``;
  }

  /**
   * 构建 Chain-of-Thought 指令
   */
  private buildCoTInstructions(): string {
    const isZh = this.config.outputLanguage === "zh-CN";

    if (isZh) {
      return `## 思考过程 (Chain-of-Thought)

在生成响应前，请按以下步骤思考：

1. **任务分析**: 用户想要完成什么？
2. **可行性评估**: 现有工具能否完成？需要什么信息？
3. **风险评估**: 这个操作有什么风险？需要确认吗？
4. **步骤规划**: 最优的执行顺序是什么？
5. **结果预测**: 预期的结果是什么？如何验证？

将思考过程融入你的 rationale 字段中。`;
    }

    return `## Chain-of-Thought

Before generating a response, think through:

1. **Task Analysis**: What does the user want to accomplish?
2. **Feasibility Assessment**: Can existing tools complete this? What info is needed?
3. **Risk Assessment**: What are the risks? Is confirmation needed?
4. **Step Planning**: What is the optimal execution order?
5. **Result Prediction**: What is the expected outcome? How to verify?

Include your reasoning in the rationale field.`;
  }

  /**
   * 构建自适应行为指令
   */
  private buildAdaptiveBehavior(): string {
    const isZh = this.config.outputLanguage === "zh-CN";

    if (isZh) {
      return `## 自适应行为

### 记住用户偏好
- 记录用户的操作习惯和偏好
- 后续请求自动应用已知偏好
- 明确说明 "根据您之前的偏好..."

### 从错误中学习
- 记录失败的操作及其原因
- 避免重复相同的错误
- 主动提供改进建议

### 反思验证
- 生成计划后，验证每个步骤是否合理
- 检查参数是否完整和正确
- 如果发现问题，主动修正`;
    }

    return `## Adaptive Behavior

### Remember User Preferences
- Record user operation habits and preferences
- Automatically apply known preferences in subsequent requests
- Explicitly state "Based on your previous preference..."

### Learn from Errors
- Record failed operations and their reasons
- Avoid repeating the same mistakes
- Proactively provide improvement suggestions

### Reflective Verification
- After generating a plan, verify each step is reasonable
- Check if parameters are complete and correct
- If issues are found, proactively correct them`;
  }

  /**
   * 构建工具描述部分
   */
  private buildToolSection(tools: Tool[]): string {
    const isZh = this.config.outputLanguage === "zh-CN";
    const limitedTools = tools.slice(0, this.config.maxToolsInPrompt);

    const toolDescriptions = limitedTools.map((tool) => {
      const params = tool.parameters
        ? tool.parameters
            .map((p) => `  - ${p.name}: ${p.description} (${p.required ? "必需" : "可选"})`)
            .join("\n")
        : "  无参数";

      return `### ${tool.name}
${tool.description}
参数:
${params}`;
    });

    const header = isZh
      ? `## 可用工具 (${limitedTools.length}/${tools.length})`
      : `## Available Tools (${limitedTools.length}/${tools.length})`;

    return `${header}

${toolDescriptions.join("\n\n")}`;
  }

  /**
   * 构建精简工具描述
   */
  private buildCompactToolSection(tools: Tool[]): string {
    const descriptions = tools.map((t) => `- ${t.name}: ${t.description.substring(0, 50)}...`);

    return `## 可用工具

${descriptions.join("\n")}`;
  }

  /**
   * 获取活动层
   */
  private getActiveLayers(): SystemMessageLayer[] {
    return this.layers
      .filter((l) => l.enabled && (!l.condition || l.condition()))
      .sort((a, b) => a.priority - b.priority);
  }

  /**
   * 替换模板变量
   */
  private replaceTemplates(content: string, context?: Record<string, unknown>): string {
    let result = content;

    // 替换配置变量
    result = result.replace(/\{\{roleName\}\}/g, this.config.roleName);
    result = result.replace(/\{\{productName\}\}/g, this.config.productName);

    // 替换上下文变量
    if (context) {
      for (const [key, value] of Object.entries(context)) {
        result = result.replace(new RegExp(`\\{\\{${key}\\}\\}`, "g"), String(value));
      }
    }

    return result;
  }
}

// ============================================================================
// 便捷函数
// ============================================================================

/**
 * 创建默认构建器
 */
export function createSystemMessageBuilder(config?: Partial<BuilderConfig>): SystemMessageBuilder {
  return new SystemMessageBuilder(config);
}

/**
 * 快速构建系统消息
 */
export function buildSystemMessage(tools?: Tool[], config?: Partial<BuilderConfig>): string {
  const builder = new SystemMessageBuilder(config);
  return builder.build(tools);
}

/**
 * 预定义的角色配置
 */
export const PRESET_CONFIGS = {
  /** Excel 助手 (默认) */
  excelAssistant: DEFAULT_CONFIG,

  /** 数据分析师 */
  dataAnalyst: {
    ...DEFAULT_CONFIG,
    roleName: "数据分析助手",
    responsibilities: [
      "帮助用户分析 Excel 数据",
      "生成统计报告和可视化图表",
      "识别数据模式和异常",
      "提供数据驱动的洞察和建议",
    ],
  },

  /** 格式专家 */
  formatExpert: {
    ...DEFAULT_CONFIG,
    roleName: "Excel 格式专家",
    responsibilities: [
      "优化表格布局和格式",
      "创建专业的报表样式",
      "应用条件格式和数据验证",
      "确保视觉一致性和可读性",
    ],
  },

  /** 公式专家 */
  formulaExpert: {
    ...DEFAULT_CONFIG,
    roleName: "Excel 公式专家",
    responsibilities: [
      "设计和优化 Excel 公式",
      "解释复杂公式的工作原理",
      "修复公式错误",
      "推荐最佳公式解决方案",
    ],
  },
} as const;
