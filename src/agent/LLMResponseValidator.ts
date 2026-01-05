/**
 * LLMResponseValidator - LLM 响应结构化验证器
 *
 * 基于 ai-agents-for-beginners 第7课 Planning Design 的学习
 * 使用 Zod schema 强制验证 LLM 输出格式，确保类型安全
 *
 * 主要功能:
 * 1. 定义 LLM 响应的 Schema
 * 2. 验证 LLM 输出是否符合预期格式
 * 3. 提供类型安全的响应解析
 * 4. 错误恢复与重试建议
 *
 * @version 1.0.0
 * @see docs/AI_AGENTS_FOR_BEGINNERS_LEARNING.md
 */

import { z } from "zod";
import { Logger } from "../utils/Logger";

// ============================================================================
// Schema 定义
// ============================================================================

/**
 * 工具调用参数 Schema
 */
export const ToolCallSchema = z.object({
  /** 工具名称 */
  name: z.string().describe("Tool name to execute"),
  /** 工具参数 */
  parameters: z.record(z.string(), z.unknown()).describe("Tool parameters as key-value pairs"),
});

/**
 * 执行步骤 Schema
 */
export const ExecutionStepSchema = z.object({
  /** 步骤序号 */
  stepNumber: z.number().int().positive().describe("Step sequence number"),
  /** 步骤描述 */
  description: z.string().describe("Human-readable description of this step"),
  /** 工具调用 */
  toolCall: ToolCallSchema.describe("The tool to execute in this step"),
  /** 步骤原因 */
  rationale: z.string().optional().describe("Why this step is necessary"),
  /** 预期结果 */
  expectedOutcome: z.string().optional().describe("What we expect after this step"),
});

/**
 * 风险等级 Schema
 */
export const RiskLevelSchema = z.enum(["low", "medium", "high", "critical"]);

/**
 * Agent 操作类型 Schema
 */
export const OperationTypeSchema = z.enum([
  "execute", // 直接执行
  "confirm", // 需要用户确认
  "clarify", // 需要澄清
  "reject", // 拒绝执行
  "complete", // 任务完成
]);

/**
 * 执行计划 Schema
 */
export const ExecutionPlanSchema = z.object({
  /** 操作类型 */
  operation: OperationTypeSchema.describe("Type of operation to perform"),

  /** 主任务描述 */
  mainTask: z.string().describe("Overall task description"),

  /** 执行步骤列表 */
  steps: z.array(ExecutionStepSchema).describe("List of execution steps"),

  /** 风险评估 */
  riskLevel: RiskLevelSchema.describe("Estimated risk level of this plan"),

  /** 风险原因 */
  riskReason: z.string().optional().describe("Why this risk level was assigned"),

  /** 用户消息 */
  userMessage: z.string().optional().describe("Message to show to user"),

  /** 需要确认的原因 */
  confirmationReason: z.string().optional().describe("Why confirmation is needed"),
});

/**
 * 澄清请求 Schema
 */
export const ClarificationRequestSchema = z.object({
  operation: z.literal("clarify"),

  /** 需要澄清的问题 */
  question: z.string().describe("Question to ask user for clarification"),

  /** 缺失的信息 */
  missingInfo: z.array(z.string()).describe("List of missing information"),

  /** 建议选项 */
  suggestions: z.array(z.string()).optional().describe("Suggested options for user"),
});

/**
 * 拒绝响应 Schema
 */
export const RejectionResponseSchema = z.object({
  operation: z.literal("reject"),

  /** 拒绝原因 */
  reason: z.string().describe("Reason for rejection"),

  /** 建议替代方案 */
  alternatives: z.array(z.string()).optional().describe("Alternative suggestions"),
});

/**
 * 完成响应 Schema
 */
export const CompletionResponseSchema = z.object({
  operation: z.literal("complete"),

  /** 完成摘要 */
  summary: z.string().describe("Summary of what was accomplished"),

  /** 执行结果 */
  results: z
    .array(
      z.object({
        step: z.string(),
        success: z.boolean(),
        output: z.unknown().optional(),
      })
    )
    .optional(),
});

/**
 * 统一的 LLM 响应 Schema
 */
export const LLMResponseSchema = z.discriminatedUnion("operation", [
  ExecutionPlanSchema.extend({ operation: z.literal("execute") }),
  ExecutionPlanSchema.extend({ operation: z.literal("confirm") }),
  ClarificationRequestSchema,
  RejectionResponseSchema,
  CompletionResponseSchema,
]);

// ============================================================================
// 类型导出
// ============================================================================

export type ToolCall = z.infer<typeof ToolCallSchema>;
export type ExecutionStep = z.infer<typeof ExecutionStepSchema>;
export type RiskLevel = z.infer<typeof RiskLevelSchema>;
export type OperationType = z.infer<typeof OperationTypeSchema>;
export type ExecutionPlan = z.infer<typeof ExecutionPlanSchema>;
export type ClarificationRequest = z.infer<typeof ClarificationRequestSchema>;
export type RejectionResponse = z.infer<typeof RejectionResponseSchema>;
export type CompletionResponse = z.infer<typeof CompletionResponseSchema>;
export type LLMResponse = z.infer<typeof LLMResponseSchema>;

// ============================================================================
// 验证器类
// ============================================================================

/**
 * 验证结果
 */
export interface ValidationResult<T> {
  success: boolean;
  data?: T;
  error?: {
    message: string;
    issues: z.ZodIssue[];
    rawInput: unknown;
  };
  /** 恢复建议 */
  recovery?: {
    canRetry: boolean;
    suggestion: string;
    fixedData?: Partial<T>;
  };
}

/**
 * Repair Retry 配置
 */
export interface RepairRetryConfig {
  /** 最大重试次数 */
  maxRetries: number;
  /** LLM 调用函数 - 用于请求修复 */
  llmCall: (prompt: string) => Promise<string>;
  /** 是否启用自动修复 */
  autoRepair: boolean;
}

/**
 * Repair Retry 结果
 */
export interface RepairResult<T> {
  success: boolean;
  data?: T;
  attempts: number;
  repairHistory: Array<{
    attempt: number;
    error: string;
    repairPrompt: string;
    repairResponse?: string;
  }>;
  finalError?: string;
}

/**
 * LLM 响应验证器
 */
export class LLMResponseValidator {
  private static readonly MODULE_NAME = "LLMResponseValidator";

  /**
   * 验证 LLM 响应
   */
  static validate(input: unknown): ValidationResult<LLMResponse> {
    try {
      // 如果输入是字符串，尝试解析 JSON
      const data = typeof input === "string" ? this.parseJSON(input) : input;

      if (data === null) {
        return this.createError("Invalid JSON input", [], input);
      }

      // 使用 Zod 验证
      const result = LLMResponseSchema.safeParse(data);

      if (result.success) {
        Logger.debug(this.MODULE_NAME, "Validation successful", {
          operation: result.data.operation,
        });
        return { success: true, data: result.data };
      }

      // 验证失败，尝试恢复
      return this.handleValidationError(result.error, data);
    } catch (error) {
      Logger.error(this.MODULE_NAME, "Validation exception", { error });
      return this.createError(error instanceof Error ? error.message : "Unknown error", [], input);
    }
  }

  /**
   * 验证执行计划
   */
  static validateExecutionPlan(input: unknown): ValidationResult<ExecutionPlan> {
    try {
      const data = typeof input === "string" ? this.parseJSON(input) : input;

      if (data === null) {
        return this.createError("Invalid JSON input", [], input);
      }

      const result = ExecutionPlanSchema.safeParse(data);

      if (result.success) {
        return { success: true, data: result.data };
      }

      return this.createError(
        "Invalid execution plan format",
        result.error.issues,
        input,
        this.suggestPlanRecovery(data, result.error)
      );
    } catch (error) {
      return this.createError(error instanceof Error ? error.message : "Unknown error", [], input);
    }
  }

  /**
   * 带 Repair Retry 的验证 - 核心方法
   * 验证失败时自动调用 LLM 进行修复，直到成功或达到最大重试次数
   */
  static async validateWithRepair(
    input: unknown,
    config: RepairRetryConfig
  ): Promise<RepairResult<LLMResponse>> {
    const repairHistory: RepairResult<LLMResponse>["repairHistory"] = [];
    let currentInput = input;
    let attempts = 0;

    while (attempts <= config.maxRetries) {
      attempts++;

      // 验证当前输入
      const result = this.validate(currentInput);

      if (result.success) {
        Logger.info(this.MODULE_NAME, "Validation successful", { attempts });
        return {
          success: true,
          data: result.data,
          attempts,
          repairHistory,
        };
      }

      // 如果达到最大重试次数或不允许自动修复，返回失败
      if (attempts > config.maxRetries || !config.autoRepair) {
        return {
          success: false,
          attempts,
          repairHistory,
          finalError: result.error?.message ?? "Validation failed",
        };
      }

      // 构建修复提示
      const repairPrompt = this.buildRepairPrompt(currentInput, result);

      Logger.info(this.MODULE_NAME, "Attempting repair", {
        attempt: attempts,
        error: result.error?.message,
      });

      try {
        // 调用 LLM 进行修复
        const repairResponse = await config.llmCall(repairPrompt);

        repairHistory.push({
          attempt: attempts,
          error: result.error?.message ?? "Unknown error",
          repairPrompt,
          repairResponse,
        });

        // 使用修复后的响应作为下一轮输入
        currentInput = repairResponse;
      } catch (llmError) {
        Logger.error(this.MODULE_NAME, "Repair LLM call failed", { error: llmError });
        repairHistory.push({
          attempt: attempts,
          error: result.error?.message ?? "Unknown error",
          repairPrompt,
          repairResponse: `LLM error: ${llmError instanceof Error ? llmError.message : "Unknown"}`,
        });

        return {
          success: false,
          attempts,
          repairHistory,
          finalError: `Repair LLM call failed: ${llmError instanceof Error ? llmError.message : "Unknown"}`,
        };
      }
    }

    return {
      success: false,
      attempts,
      repairHistory,
      finalError: "Max retries exceeded",
    };
  }

  /**
   * 构建修复提示
   */
  private static buildRepairPrompt(input: unknown, result: ValidationResult<LLMResponse>): string {
    const issues = result.error?.issues ?? [];
    const issueDescriptions = issues
      .map((issue) => {
        const path = issue.path.join(".");
        return `- 字段 "${path || "root"}": ${issue.message} (期望: ${issue.code})`;
      })
      .join("\n");

    return `你之前的响应格式不正确，请修复并重新生成。

## 原始响应
\`\`\`json
${typeof input === "string" ? input : JSON.stringify(input, null, 2)}
\`\`\`

## 验证错误
${result.error?.message ?? "Unknown error"}

## 具体问题
${issueDescriptions || "格式不符合预期"}

## 修复建议
${result.recovery?.suggestion ?? "请确保响应是有效的 JSON，且包含所有必需字段"}

## 正确格式示例
\`\`\`json
{
  "operation": "execute",
  "mainTask": "任务描述",
  "steps": [
    {
      "stepNumber": 1,
      "description": "步骤描述",
      "toolCall": {
        "name": "tool_name",
        "parameters": {}
      }
    }
  ],
  "riskLevel": "low"
}
\`\`\`

请直接返回修复后的 JSON，不要添加任何解释文字。`;
  }

  /**
   * 验证单个工具调用
   */
  static validateToolCall(input: unknown): ValidationResult<ToolCall> {
    try {
      const data = typeof input === "string" ? this.parseJSON(input) : input;

      if (data === null) {
        return this.createError("Invalid JSON input", [], input);
      }

      const result = ToolCallSchema.safeParse(data);

      if (result.success) {
        return { success: true, data: result.data };
      }

      return this.createError("Invalid tool call format", result.error.issues, input);
    } catch (error) {
      return this.createError(error instanceof Error ? error.message : "Unknown error", [], input);
    }
  }

  /**
   * 从自由文本中提取工具调用
   * 用于处理 LLM 没有严格按照 JSON 格式输出的情况
   */
  static extractToolCallsFromText(text: string): ToolCall[] {
    const toolCalls: ToolCall[] = [];

    // 尝试提取 JSON 块
    const jsonBlocks = text.match(/```json\s*([\s\S]*?)```/g) || [];
    for (const block of jsonBlocks) {
      const json = block
        .replace(/```json\s*/, "")
        .replace(/```/, "")
        .trim();
      try {
        const parsed = JSON.parse(json);
        if (parsed.name && parsed.parameters) {
          const result = ToolCallSchema.safeParse(parsed);
          if (result.success) {
            toolCalls.push(result.data);
          }
        }
      } catch {
        // 继续尝试其他块
      }
    }

    // 尝试提取内联 JSON 对象
    const inlineJson = text.match(/\{[^{}]*"name"\s*:\s*"[^"]+"\s*[^{}]*\}/g) || [];
    for (const json of inlineJson) {
      try {
        const parsed = JSON.parse(json);
        const result = ToolCallSchema.safeParse(parsed);
        if (result.success) {
          // 避免重复
          if (!toolCalls.some((tc) => tc.name === result.data.name)) {
            toolCalls.push(result.data);
          }
        }
      } catch {
        // 忽略解析错误
      }
    }

    return toolCalls;
  }

  // ============================================================================
  // 私有方法
  // ============================================================================

  /**
   * 解析 JSON
   */
  private static parseJSON(input: string): unknown | null {
    try {
      // 清理常见问题
      let cleaned = input.trim();

      // 移除 markdown 代码块标记
      if (cleaned.startsWith("```json")) {
        cleaned = cleaned.replace(/^```json\s*/, "").replace(/```$/, "");
      } else if (cleaned.startsWith("```")) {
        cleaned = cleaned.replace(/^```\s*/, "").replace(/```$/, "");
      }

      return JSON.parse(cleaned);
    } catch {
      // 尝试修复常见的 JSON 问题
      try {
        // 替换单引号为双引号
        const fixed = input.replace(/'/g, '"');
        return JSON.parse(fixed);
      } catch {
        return null;
      }
    }
  }

  /**
   * 处理验证错误
   */
  private static handleValidationError(
    error: z.ZodError,
    data: unknown
  ): ValidationResult<LLMResponse> {
    const issues = error.issues;

    // 尝试自动修复常见问题
    const recovery = this.attemptRecovery(data, issues);

    if (recovery.fixed) {
      const retryResult = LLMResponseSchema.safeParse(recovery.fixed);
      if (retryResult.success) {
        Logger.info(this.MODULE_NAME, "Auto-recovered from validation error", {
          originalIssues: issues.length,
        });
        return {
          success: true,
          data: retryResult.data,
          recovery: {
            canRetry: false,
            suggestion: "Auto-fixed",
            fixedData: recovery.fixed as Partial<LLMResponse>,
          },
        };
      }
    }

    return this.createError(
      "LLM response validation failed",
      issues,
      data,
      recovery.suggestion ? { canRetry: true, suggestion: recovery.suggestion } : undefined
    );
  }

  /**
   * 尝试恢复数据
   */
  private static attemptRecovery(
    data: unknown,
    issues: z.ZodIssue[]
  ): { fixed?: unknown; suggestion?: string } {
    if (typeof data !== "object" || data === null) {
      return { suggestion: "Input must be a valid JSON object" };
    }

    const obj = data as Record<string, unknown>;
    const fixed = { ...obj };

    // 修复缺失的 operation
    if (!obj.operation) {
      // 根据其他字段推断 operation
      if (obj.steps && Array.isArray(obj.steps)) {
        fixed.operation = "execute";
      } else if (obj.question) {
        fixed.operation = "clarify";
      } else if (obj.reason && obj.alternatives) {
        fixed.operation = "reject";
      } else if (obj.summary) {
        fixed.operation = "complete";
      }
    }

    // 修复 riskLevel 大小写
    if (typeof obj.riskLevel === "string") {
      fixed.riskLevel = obj.riskLevel.toLowerCase();
    }

    // 修复空 steps 数组
    if (obj.operation === "execute" && !obj.steps) {
      fixed.steps = [];
    }

    // 确保 mainTask 存在
    if (!obj.mainTask && typeof obj.task === "string") {
      fixed.mainTask = obj.task;
    }

    return { fixed, suggestion: "Check operation type and required fields" };
  }

  /**
   * 建议执行计划恢复
   */
  private static suggestPlanRecovery(
    _data: unknown,
    error: z.ZodError
  ): ValidationResult<ExecutionPlan>["recovery"] {
    const missingFields = error.issues
      .filter(
        (i) => i.code === "invalid_type" && (i as { received?: string }).received === "undefined"
      )
      .map((i) => i.path.join("."));

    if (missingFields.length > 0) {
      return {
        canRetry: true,
        suggestion: `Missing required fields: ${missingFields.join(", ")}`,
      };
    }

    return {
      canRetry: true,
      suggestion: "Check the execution plan format",
    };
  }

  /**
   * 创建错误结果
   */
  private static createError<T>(
    message: string,
    issues: z.ZodIssue[],
    rawInput: unknown,
    recovery?: ValidationResult<T>["recovery"]
  ): ValidationResult<T> {
    Logger.warn(this.MODULE_NAME, "Validation failed", {
      message,
      issueCount: issues.length,
      firstIssue: issues[0],
    });

    return {
      success: false,
      error: { message, issues, rawInput },
      recovery,
    };
  }
}

// ============================================================================
// 便捷函数
// ============================================================================

/**
 * 验证并解析 LLM 响应
 */
export function validateLLMResponse(input: unknown): LLMResponse | null {
  const result = LLMResponseValidator.validate(input);
  return result.success ? result.data! : null;
}

/**
 * 验证并解析执行计划
 */
export function validateExecutionPlan(input: unknown): ExecutionPlan | null {
  const result = LLMResponseValidator.validateExecutionPlan(input);
  return result.success ? result.data! : null;
}

/**
 * 创建一个类型安全的执行计划
 */
export function createExecutionPlan(
  mainTask: string,
  steps: Array<{
    description: string;
    toolName: string;
    parameters: Record<string, unknown>;
    rationale?: string;
  }>,
  options?: {
    operation?: OperationType;
    riskLevel?: RiskLevel;
    userMessage?: string;
  }
): ExecutionPlan {
  return {
    operation: options?.operation ?? "execute",
    mainTask,
    steps: steps.map((step, index) => ({
      stepNumber: index + 1,
      description: step.description,
      toolCall: {
        name: step.toolName,
        parameters: step.parameters,
      },
      rationale: step.rationale,
    })),
    riskLevel: options?.riskLevel ?? "low",
    userMessage: options?.userMessage,
  };
}
