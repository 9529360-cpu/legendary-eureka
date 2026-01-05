/**
 * ClaudeAdapter - Claude 模型适配器
 *
 * 单一职责：适配 Anthropic Claude API
 * 行数上限：300 行
 */

import {
  BaseAdapter,
  ModelConfig,
  ChatMessage,
  ModelResponse,
  ToolCallRequest,
} from "./BaseAdapter";
import { ToolContract } from "../contracts/ToolContract";

// ========== Claude 配置 ==========

const CLAUDE_API_BASE = "https://api.anthropic.com/v1";
const CLAUDE_MESSAGES_ENDPOINT = "/messages";
const ANTHROPIC_VERSION = "2023-06-01";

/**
 * Claude 特定配置
 */
export interface ClaudeConfig extends ModelConfig {
  model: "claude-3-opus-20240229" | "claude-3-sonnet-20240229" | "claude-3-haiku-20240307" | string;
}

// ========== ClaudeAdapter 类 ==========

/**
 * Claude 模型适配器
 */
export class ClaudeAdapter extends BaseAdapter {
  get name(): string {
    return "claude";
  }

  constructor(config: ClaudeConfig) {
    super({
      ...config,
      baseUrl: config.baseUrl || CLAUDE_API_BASE,
    });
  }

  /**
   * 发送聊天消息
   */
  async chat(messages: ChatMessage[], tools?: ToolContract[]): Promise<ModelResponse> {
    // 提取 system 消息
    const systemMessage = messages.find((m) => m.role === "system");
    const otherMessages = messages.filter((m) => m.role !== "system");

    const body: Record<string, unknown> = {
      model: this.config.model,
      messages: this.formatMessages(otherMessages),
      max_tokens: this.config.maxTokens ?? 4096,
    };

    if (systemMessage) {
      body.system = systemMessage.content;
    }

    if (tools && tools.length > 0) {
      body.tools = this.formatToolsForApi(tools);
    }

    try {
      const response = await this.withTimeout(this.fetchClaude(CLAUDE_MESSAGES_ENDPOINT, body));
      return this.parseApiResponse(response);
    } catch (error) {
      return {
        content: `Claude API 错误: ${error instanceof Error ? error.message : String(error)}`,
        finishReason: "error",
      };
    }
  }

  /**
   * 生成文本
   */
  async generate(prompt: string): Promise<string> {
    const messages: ChatMessage[] = [{ role: "user", content: prompt }];
    const response = await this.chat(messages);
    return response.content;
  }

  /**
   * 测试连接
   */
  async testConnection(): Promise<boolean> {
    try {
      const response = await this.chat([{ role: "user", content: '你好，请回复"连接成功"' }]);
      return response.finishReason !== "error";
    } catch {
      return false;
    }
  }

  /**
   * 获取模型信息
   */
  getModelInfo(): { name: string; version?: string; maxTokens?: number } {
    return {
      name: this.config.model,
      maxTokens: this.config.maxTokens,
    };
  }

  /**
   * Claude 专用请求
   */
  private async fetchClaude(endpoint: string, body: Record<string, unknown>): Promise<unknown> {
    const url = `${this.config.baseUrl}${endpoint}`;

    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": this.config.apiKey,
        "anthropic-version": ANTHROPIC_VERSION,
      },
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      const error = await response.text();
      throw new Error(`API Error (${response.status}): ${error}`);
    }

    return response.json();
  }

  /**
   * 格式化消息 (Claude 格式)
   */
  private formatMessages(messages: ChatMessage[]): unknown[] {
    return messages.map((msg) => ({
      role: msg.role === "assistant" ? "assistant" : "user",
      content: msg.content,
    }));
  }

  /**
   * 格式化工具为 Claude API 格式
   */
  protected formatToolsForApi(tools: ToolContract[]): unknown[] {
    return tools.map((tool) => ({
      name: tool.name,
      description: tool.description,
      input_schema: tool.inputSchema,
    }));
  }

  /**
   * 解析 Claude API 响应
   */
  protected parseApiResponse(response: unknown): ModelResponse {
    const resp = response as {
      content?: Array<{
        type: "text" | "tool_use";
        text?: string;
        id?: string;
        name?: string;
        input?: Record<string, unknown>;
      }>;
      stop_reason?: string;
      usage?: {
        input_tokens?: number;
        output_tokens?: number;
      };
    };

    let content = "";
    const toolCalls: ToolCallRequest[] = [];

    if (resp.content) {
      for (const block of resp.content) {
        if (block.type === "text" && block.text) {
          content += block.text;
        } else if (block.type === "tool_use" && block.id && block.name) {
          toolCalls.push({
            id: block.id,
            name: block.name,
            arguments: block.input || {},
          });
        }
      }
    }

    return {
      content,
      toolCalls: toolCalls.length > 0 ? toolCalls : undefined,
      finishReason: this.mapFinishReason(resp.stop_reason),
      usage: resp.usage
        ? {
            promptTokens: resp.usage.input_tokens || 0,
            completionTokens: resp.usage.output_tokens || 0,
            totalTokens: (resp.usage.input_tokens || 0) + (resp.usage.output_tokens || 0),
          }
        : undefined,
    };
  }

  /**
   * 映射完成原因
   */
  private mapFinishReason(reason?: string): ModelResponse["finishReason"] {
    switch (reason) {
      case "end_turn":
        return "stop";
      case "tool_use":
        return "tool_calls";
      case "max_tokens":
        return "length";
      default:
        return "stop";
    }
  }
}

/**
 * 创建 Claude 适配器
 */
export function createClaudeAdapter(apiKey: string, model?: string): ClaudeAdapter {
  return new ClaudeAdapter({
    apiKey,
    model: model || "claude-3-sonnet-20240229",
  });
}

export default ClaudeAdapter;
