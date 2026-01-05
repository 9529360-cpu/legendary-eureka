/**
 * 配置相关类型定义
 *
 * 从 AgentCore.ts 抽取，用于定义 Agent 配置接口
 */

import type { HardValidationRule } from "./validation";

/**
 * Agent 配置
 */
export interface AgentConfig {
  maxIterations: number;
  defaultTimeout: number;
  systemPrompt?: string;
  enableMemory: boolean;
  verboseLogging: boolean;
  // v2.9.6: 校验规则配置
  validation?: ValidationConfig;
  // v2.9.6: 操作历史持久化配置
  persistence?: PersistenceConfig;
  // v2.9.58: 交互策略配置（P2: 澄清机制）
  interaction?: InteractionConfig;
  // v2.9.58: 反思配置（P0: 每步反思）
  reflection?: ReflectionConfig;
  // v2.9.58: 验证信号配置（P1: 验证作为信号）
  validationSignal?: ValidationSignalConfig;
  // v2.9.59: 协议版组件配置
  clarifyGate?: Partial<import("../ClarifyGate").ClarifyGateConfig>;
  stepDecider?: Partial<import("../StepDecider").DeciderConfig>;
  responseBuilder?: Partial<import("../ResponseBuilder").ResponseBuilderConfig>;
}

/**
 * v2.9.58: 交互策略配置
 */
export interface InteractionConfig {
  /**
   * 意图置信度阈值 (0-1)
   * 低于此值时必须向用户澄清，不直接执行
   * 默认 0.7
   */
  clarificationThreshold: number;

  /**
   * 破坏性操作确认
   * 删除/覆盖/清空等操作前是否需要用户确认
   * 默认 true
   */
  confirmDestructiveOps: boolean;

  /**
   * 提供备选方案
   * 当意图不够明确时，是否提供多个可选方案让用户选择
   * 默认 true
   */
  offerAlternatives: boolean;

  /**
   * 允许自由表达
   * 允许 LLM 自由生成响应，而非强制使用模板
   * 默认 true
   */
  allowFreeformResponse: boolean;

  /**
   * 大范围操作确认阈值
   * 影响超过此数量单元格的操作需要确认
   * 默认 100
   */
  largeOperationThreshold: number;

  /**
   * 每步反思（P0）
   * 每个执行步骤后让 LLM 反思结果，决定是否调整后续计划
   * 默认 true
   */
  enableStepReflection: boolean;

  /**
   * 主动建议
   * 任务完成后主动提供相关建议
   * 默认 true
   */
  proactiveSuggestions: boolean;
}

/**
 * v2.9.6: 校验规则配置 - 可启用/禁用特定规则
 */
export interface ValidationConfig {
  /** 是否启用硬校验（默认 true） */
  enabled: boolean;
  /** 要禁用的规则 ID 列表 */
  disabledRules?: string[];
  /** 将特定规则的严重性从 block 降级为 warn */
  downgradeToWarn?: string[];
  /** 自定义规则（可外部注入） */
  customRules?: HardValidationRule[];
}

/**
 * v2.9.6: 操作历史持久化配置
 */
export interface PersistenceConfig {
  /** 是否启用持久化（默认 false） */
  enabled: boolean;
  /** 存储键名前缀 */
  storageKeyPrefix?: string;
  /** 最大保存的操作数量 */
  maxOperations?: number;
  /** 保留时间（小时） */
  retentionHours?: number;
}

/**
 * v2.9.58: 反思配置
 */
export interface ReflectionConfig {
  /** 是否启用每步反思 */
  enabled: boolean;
  /** 反思详细程度 */
  verbosity?: "minimal" | "normal" | "detailed";
  /** 置信度阈值，低于此值需要更详细的反思 */
  confidenceThreshold?: number;
}

/**
 * v2.9.58: 验证信号配置
 */
export interface ValidationSignalConfig {
  /** 是否启用验证信号 */
  enabled: boolean;
  /** 严重级别阈值 */
  severityThreshold?: "block" | "warn" | "info";
}

/**
 * v2.9.20: 回复简化配置
 */
export interface ResponseSimplificationConfig {
  /** 是否隐藏技术细节 */
  hideTechnicalDetails: boolean;
  /** 最大回复长度（字符数） */
  maxLength: number;
  /** 是否显示步骤进度 */
  showProgress: boolean;
  /** 是否显示思考过程 */
  showThinking: boolean;
  /** 是否显示工具调用 */
  showToolCalls: boolean;
  /** 详细程度 */
  verbosity: "minimal" | "normal" | "detailed";
}

/**
 * v2.9.20: 智能确认配置
 */
export interface ConfirmationConfig {
  /** 操作风险等级 */
  riskLevel: "low" | "medium" | "high" | "critical";
  /** 操作类型 */
  operationType: string;
  /** 是否需要确认 */
  requiresConfirmation: boolean;
  /** 确认消息 */
  confirmationMessage: string;
  /** 影响范围描述 */
  impactDescription: string;
  /** 可撤销性 */
  reversible: boolean;
}

/**
 * v2.9.20: 友好错误信息
 */
export interface FriendlyError {
  /** 原始错误代码 */
  code: string;
  /** 原始错误消息 */
  originalMessage: string;
  /** 友好错误消息 */
  friendlyMessage: string;
  /** 可能的原因 */
  possibleCauses: string[];
  /** 建议的解决方案 */
  suggestions: string[];
  /** 是否可自动恢复 */
  autoRecoverable: boolean;
  /** 严重程度 */
  severity: "info" | "warning" | "error" | "critical";
}

/**
 * v2.9.21: 专家 Agent 类型
 */
export type ExpertAgentType =
  | "data_analyst"
  | "formula_expert"
  | "chart_specialist"
  | "format_designer"
  | "data_cleaner";

/**
 * v2.9.21: 专家 Agent 配置
 */
export interface ExpertAgentConfig {
  /** Agent 类型标识 */
  type: string;
  /** Agent 名称 */
  name: string;
  /** Agent 描述 */
  description: string;
  /** 擅长的任务类型 */
  specialties: string[];
  /** 可使用的工具 */
  tools: string[];
  /** 系统提示词补充 */
  systemPromptAddition: string;
}
