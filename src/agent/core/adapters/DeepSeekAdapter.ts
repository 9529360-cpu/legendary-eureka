/**
 * DeepSeekAdapter - DeepSeek 模型适配器
 * 
 * 单一职责：适配 DeepSeek API
 * 行数上限：300 行
 */

import { BaseAdapter, ModelConfig, ChatMessage, ModelResponse, ToolCallRequest } from './BaseAdapter';
import { ToolContract } from '../contracts/ToolContract';

// ========== DeepSeek 配置 ==========

const DEEPSEEK_API_BASE = 'https://api.deepseek.com';
const DEEPSEEK_CHAT_ENDPOINT = '/chat/completions';

/**
 * DeepSeek 特定配置
 */
export interface DeepSeekConfig extends ModelConfig {
  model: 'deepseek-chat' | 'deepseek-coder' | string;
}

// ========== DeepSeekAdapter 类 ==========

/**
 * DeepSeek 模型适配器
 */
export class DeepSeekAdapter extends BaseAdapter {
  get name(): string {
    return 'deepseek';
  }

  constructor(config: DeepSeekConfig) {
    super({
      ...config,
      baseUrl: config.baseUrl || DEEPSEEK_API_BASE,
    });
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
      body.tool_choice = 'auto';
    }

    try {
      const response = await this.withTimeout(
        this.fetchApi(DEEPSEEK_CHAT_ENDPOINT, body)
      );
      return this.parseApiResponse(response);
    } catch (error) {
      return {
        content: `DeepSeek API 错误: ${error instanceof Error ? error.message : String(error)}`,
        finishReason: 'error',
      };
    }
  }

  /**
   * 生成文本
   */
  async generate(prompt: string): Promise<string> {
    const messages: ChatMessage[] = [{ role: 'user', content: prompt }];
    const response = await this.chat(messages);
    return response.content;
  }

  /**
   * 测试连接
   */
  async testConnection(): Promise<boolean> {
    try {
      const response = await this.chat([
        { role: 'user', content: '你好，请回复"连接成功"' },
      ]);
      return response.finishReason !== 'error';
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
   * 格式化消息
   */
  private formatMessages(messages: ChatMessage[]): unknown[] {
    return messages.map(msg => ({
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
    return tools.map(tool => ({
      type: 'function',
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
      toolCalls = message.tool_calls.map(tc => ({
        id: tc.id,
        name: tc.function.name,
        arguments: JSON.parse(tc.function.arguments || '{}'),
      }));
    }

    return {
      content: message?.content || '',
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
  private mapFinishReason(reason?: string): ModelResponse['finishReason'] {
    switch (reason) {
      case 'stop':
        return 'stop';
      case 'tool_calls':
        return 'tool_calls';
      case 'length':
        return 'length';
      default:
        return 'stop';
    }
  }
}

/**
 * 创建 DeepSeek 适配器
 */
export function createDeepSeekAdapter(apiKey: string, model?: string): DeepSeekAdapter {
  return new DeepSeekAdapter({
    apiKey,
    model: model || 'deepseek-chat',
  });
}

export default DeepSeekAdapter;
