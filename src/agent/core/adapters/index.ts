/**
 * 适配器模块索引文件
 */

export {
  BaseAdapter,
  IModelAdapter,
  ModelConfig,
  ChatMessage,
  MessageRole,
  ModelResponse,
  ToolCallRequest,
} from './BaseAdapter';

export { DeepSeekAdapter, DeepSeekConfig, createDeepSeekAdapter } from './DeepSeekAdapter';
export { OpenAIAdapter, OpenAIConfig, createOpenAIAdapter } from './OpenAIAdapter';
export { ClaudeAdapter, ClaudeConfig, createClaudeAdapter } from './ClaudeAdapter';

// ========== 适配器工厂 ==========

import { IModelAdapter, ModelConfig } from './BaseAdapter';
import { DeepSeekAdapter } from './DeepSeekAdapter';
import { OpenAIAdapter } from './OpenAIAdapter';
import { ClaudeAdapter } from './ClaudeAdapter';

/**
 * 支持的模型提供商
 */
export type ModelProvider = 'deepseek' | 'openai' | 'claude';

/**
 * 创建适配器
 */
export function createAdapter(
  provider: ModelProvider,
  config: ModelConfig
): IModelAdapter {
  switch (provider) {
    case 'deepseek':
      return new DeepSeekAdapter(config as import('./DeepSeekAdapter').DeepSeekConfig);
    case 'openai':
      return new OpenAIAdapter(config as import('./OpenAIAdapter').OpenAIConfig);
    case 'claude':
      return new ClaudeAdapter(config as import('./ClaudeAdapter').ClaudeConfig);
    default:
      throw new Error(`Unsupported model provider: ${provider}`);
  }
}

/**
 * 根据模型名称推断提供商
 */
export function inferProvider(model: string): ModelProvider {
  const lowerModel = model.toLowerCase();
  
  if (lowerModel.includes('gpt') || lowerModel.includes('openai')) {
    return 'openai';
  }
  if (lowerModel.includes('claude')) {
    return 'claude';
  }
  if (lowerModel.includes('deepseek')) {
    return 'deepseek';
  }
  
  // 默认使用 DeepSeek
  return 'deepseek';
}
