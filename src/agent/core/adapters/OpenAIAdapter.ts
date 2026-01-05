/**
 * OpenAIAdapter - OpenAI 模型适配器
 *
 * 单一职责：适配 OpenAI API (GPT-4, GPT-3.5 等)
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

// ========== OpenAI 配置 ==========

const OPENAI_API_BASE = "https://api.openai.com/v1";
const OPENAI_CHAT_ENDPOINT = "/chat/completions";

/**
 * OpenAI 特定配置
 */
export interface OpenAIConfig extends ModelConfig {
  model: "gpt-4" | "gpt-4-turbo" | "gpt-3.5-turbo" | string;
  organization?: string;
}

// ========== OpenAIAdapter 类 ==========

/**
 * OpenAI 模型适配器
 */
export class OpenAIAdapter extends BaseAdapter {
  private organization?: string;

  get name(): string {
    return "openai";
  }

  constructor(config: OpenAIConfig) {
    super({
      ...config,
      baseUrl: config.baseUrl || OPENAI_API_BASE,
    });
    this.organization = config.organization;
  }

  /**
   * 发送聊天消息
   */
  async chat(messages: ChatMessage[], tools?: ToolContract[]): Promise<ModelResponse> {
    const body: Record<string, unknown> = {
      model: this.config.model,
      messages: this.formatMessages(messages),
      temperature: this.config.temperature ?? 0.7,
      max_tokens: this.config.maxTokens ?? 4096,
    };

    if (tools && tools.length > 0) {
      body.tools = this.formatToolsForApi(tools);
      body.tool_choice = "auto";
    }

    try {
      const response = await this.withTimeout(this.fetchWithOrg(OPENAI_CHAT_ENDPOINT, body));
      return this.parseApiResponse(response);
    } catch (error) {
      return {
        content: `OpenAI API 错误: ${error instanceof Error ? error.message : String(error)}`,
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
      const response = await this.chat([{ role: "user", content: 'Hello, respond with "OK"' }]);
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
   * 带 Organization 的请求
   */
  private async fetchWithOrg(endpoint: string, body: Record<string, unknown>): Promise<unknown> {
    const url = `${this.config.baseUrl}${endpoint}`;

    const headers: Record<string, string> = {
      "Content-Type": "application/json",
      Authorization: `Bearer ${this.config.apiKey}`,
    };

    if (this.organization) {
      headers["OpenAI-Organization"] = this.organization;
    }

    const response = await fetch(url, {
      method: "POST",
      headers,
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      const error = await response.text();
      throw new Error(`API Error (${response.status}): ${error}`);
    }

    return response.json();
  }

  /**
   * 格式化消息
   */
  private formatMessages(messages: ChatMessage[]): unknown[] {
    return messages.map((msg) => ({
      role: msg.role,
      content: msg.content,
      ...(msg.name && { name: msg.name }),
      ...(msg.toolCallId && { tool_call_id: msg.toolCallId }),
    }));
  }

  /**
   * 格式化工具为 API 格式
   */
  protected formatToolsForApi(tools: ToolContract[]): unknown[] {
    return tools.map((tool) => ({
      type: "function",
      function: {
        name: tool.name,
        description: tool.description,
        parameters: tool.inputSchema,
      },
    }));
  }

  /**
   * 解析 API 响应
   */
  protected parseApiResponse(response: unknown): ModelResponse {
    const resp = response as {
      choices?: Array<{
        message?: {
          content?: string;
          tool_calls?: Array<{
            id: string;
            function: {
              name: string;
              arguments: string;
            };
          }>;
        };
        finish_reason?: string;
      }>;
      usage?: {
        prompt_tokens?: number;
        completion_tokens?: number;
        total_tokens?: number;
      };
    };

    const choice = resp.choices?.[0];
    const message = choice?.message;

    let toolCalls: ToolCallRequest[] | undefined;
    if (message?.tool_calls && message.tool_calls.length > 0) {
      toolCalls = message.tool_calls.map((tc) => ({
        id: tc.id,
        name: tc.function.name,
        arguments: JSON.parse(tc.function.arguments || "{}"),
      }));
    }

    return {
      content: message?.content || "",
      toolCalls,
      finishReason: this.mapFinishReason(choice?.finish_reason),
      usage: resp.usage
        ? {
            promptTokens: resp.usage.prompt_tokens || 0,
            completionTokens: resp.usage.completion_tokens || 0,
            totalTokens: resp.usage.total_tokens || 0,
          }
        : undefined,
    };
  }

  /**
   * 映射完成原因
   */
  private mapFinishReason(reason?: string): ModelResponse["finishReason"] {
    switch (reason) {
      case "stop":
        return "stop";
      case "tool_calls":
        return "tool_calls";
      case "length":
        return "length";
      default:
        return "stop";
    }
  }
}

/**
 * 创建 OpenAI 适配器
 */
export function createOpenAIAdapter(apiKey: string, model?: string): OpenAIAdapter {
  return new OpenAIAdapter({
    apiKey,
    model: model || "gpt-4-turbo",
  });
}

export default OpenAIAdapter;
