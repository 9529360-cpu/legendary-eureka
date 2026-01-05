/**
 * 工具相关类型定义
 *
 * 从 AgentCore.ts 抽取，用于定义 Agent 可使用的工具接口
 */

/**
 * 工具定义 - 任何外部能力都是一个 Tool
 */
export interface Tool {
  name: string;
  description: string;
  category: string; // 'excel' | 'word' | 'filesystem' | 'api' | 'browser' | ...
  parameters: ToolParameter[];
  execute: (input: Record<string, unknown>) => Promise<ToolResult>;
}

/**
 * 工具参数定义
 */
export interface ToolParameter {
  name: string;
  type: "string" | "number" | "boolean" | "array" | "object";
  description: string;
  required: boolean;
  default?: unknown;
}

/**
 * 工具执行结果
 */
export interface ToolResult {
  success: boolean;
  output: string;
  data?: unknown;
  error?: string;
}

/**
 * v2.9.21: 工具链 - 预定义的工具调用序列
 */
export interface ToolChain {
  /** 工具链ID */
  id: string;
  /** 工具链名称 */
  name: string;
  /** 工具调用序列 */
  steps: Array<{
    toolName: string;
    purpose: string;
    dependsOn: string[];
    outputMapping?: Record<string, string>;
  }>;
  /** 适用场景 */
  applicablePatterns: string[];
  /** 成功率 */
  successRate: number;
  /** 使用次数 */
  usageCount: number;
}

/**
 * v2.9.22: 工具调用结果验证
 */
export interface ToolResultValidation {
  /** 是否有效 */
  isValid: boolean;
  /** 验证类型 */
  validationType: "type_check" | "range_check" | "semantic_check" | "custom";
  /** 验证详情 */
  details: string;
  /** 建议的修复 */
  suggestedFix?: string;
  /** 是否可自动修复 */
  autoFixable: boolean;
}

/**
 * 工具调用事件信息 - 类似 LlamaIndex 的 ToolCall
 */
export interface ToolCallInfo {
  toolName: string;
  toolKwargs: Record<string, unknown>;
  toolId: string;
}

/**
 * 工具调用结果事件数据 - 类似 LlamaIndex 的 ToolCallResult
 */
export interface ToolCallResultData {
  toolName: string;
  toolKwargs: Record<string, unknown>;
  toolId: string;
  toolOutput: ToolResult;
  returnDirect: boolean; // 是否直接返回给用户
}
