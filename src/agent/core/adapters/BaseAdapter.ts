/**
 * BaseAdapter - 基础模型适配器
 *
 * 单一职责：定义模型适配器的抽象接口
 * 行数上限：200 行
 *
 * 所有模型适配器必须实现此接口，确保模型无关性
 */

import { ToolContract } from "../contracts/ToolContract";

// ========== 适配器接口 ==========

/**
 * 模型配置
 */
export interface ModelConfig {
  apiKey: string;
  baseUrl?: string;
  model: string;
  temperature?: number;
  maxTokens?: number;
  timeout?: number;
}

/**
 * 消息角色
 */
export type MessageRole = "system" | "user" | "assistant" | "tool";

/**
 * 聊天消息
 */
export interface ChatMessage {
  role: MessageRole;
  content: string;
  name?: string;
  toolCallId?: string;
}

/**
 * 工具调用请求
 */
export interface ToolCallRequest {
  id: string;
  name: string;
  arguments: Record<string, unknown>;
}

/**
 * 模型响应
 */
export interface ModelResponse {
  content: string;
  toolCalls?: ToolCallRequest[];
  finishReason: "stop" | "tool_calls" | "length" | "error";
  usage?: {
    promptTokens: number;
    completionTokens: number;
    totalTokens: number;
  };
}

/**
 * 适配器接口
 */
export interface IModelAdapter {
  /**
   * 适配器名称
   */
  readonly name: string;

  /**
   * 发送消息
   */
  chat(messages: ChatMessage[], tools?: ToolContract[]): Promise<ModelResponse>;

  /**
   * 生成文本（无工具调用）
   */
  generate(prompt: string): Promise<string>;

  /**
   * 检查连接
   */
  testConnection(): Promise<boolean>;

  /**
   * 获取模型信息
   */
  getModelInfo(): { name: string; version?: string; maxTokens?: number };
}

// ========== 抽象基类 ==========

/**
 * 基础适配器
 */
export abstract class BaseAdapter implements IModelAdapter {
  protected config: ModelConfig;

  constructor(config: ModelConfig) {
    this.config = config;
  }

  abstract get name(): string;

  abstract chat(messages: ChatMessage[], tools?: ToolContract[]): Promise<ModelResponse>;

  abstract generate(prompt: string): Promise<string>;

  abstract testConnection(): Promise<boolean>;

  abstract getModelInfo(): { name: string; version?: string; maxTokens?: number };

  /**
   * 将 ToolContract 转换为 API 格式
   */
  protected abstract formatToolsForApi(tools: ToolContract[]): unknown;

  /**
   * 解析 API 响应
   */
  protected abstract parseApiResponse(response: unknown): ModelResponse;

  /**
   * 发送 HTTP 请求
   */
  protected async fetchApi(endpoint: string, body: Record<string, unknown>): Promise<unknown> {
    const url = this.config.baseUrl ? `${this.config.baseUrl}${endpoint}` : endpoint;

    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${this.config.apiKey}`,
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
   * 超时控制
   */
  protected withTimeout<T>(promise: Promise<T>, timeoutMs?: number): Promise<T> {
    const timeout = timeoutMs || this.config.timeout || 30000;

    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        reject(new Error(`Request timeout after ${timeout}ms`));
      }, timeout);

      promise
        .then((result) => {
          clearTimeout(timer);
          resolve(result);
        })
        .catch((error) => {
          clearTimeout(timer);
          reject(error);
        });
    });
  }
}

export default BaseAdapter;
