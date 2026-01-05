/**
 * ToolExecutor - 工具执行器（带兜底策略）
 * v1.0.0
 *
 * 功能：
 * 1. 统一的工具执行入口
 * 2. 工具查找与兜底策略
 * 3. 参数验证与转换
 * 4. 执行监控集成
 * 5. 错误恢复与重试
 *
 * 解决的问题：
 * - 工具未注册时仅返回错误信息，无详细处理
 * - 缺乏统一的工具执行管理
 * - 工具调用链路不可追踪
 */

import { Tool, ToolResult } from "../agent/AgentCore";
import { TaskExecutionMonitor } from "./TaskExecutionMonitor";
import { Logger } from "../utils/Logger";
import { RETRY, TIMEOUTS } from "../config/constants";

// ========== 类型定义 ==========

/**
 * 工具查找结果
 */
export interface ToolLookupResult {
  found: boolean;
  tool?: Tool;
  alternatives?: Tool[];
  suggestion?: string;
}

/**
 * 工具执行选项
 */
export interface ToolExecutionOptions {
  taskId?: string;
  timeout?: number;
  retryOnFailure?: boolean;
  maxRetries?: number;
  retryDelay?: number;
  fallbackTools?: string[];
  validateParams?: boolean;
  skipMonitoring?: boolean;
}

/**
 * 工具执行结果（增强版）
 */
export interface EnhancedToolResult extends ToolResult {
  toolName: string;
  executionTime: number;
  retryCount: number;
  fallbackUsed?: string;
  warnings?: string[];
}

/**
 * 工具映射类型
 */
export type ToolMap = Map<string, Tool>;

/**
 * 兜底策略配置
 */
export interface FallbackConfig {
  /** 工具名称 -> 备选工具列表 */
  toolFallbacks: Record<string, string[]>;
  /** 分类 -> 默认兜底工具 */
  categoryFallbacks: Record<string, string>;
  /** 全局默认兜底工具 */
  globalFallback?: string;
}

// ========== 默认兜底配置 ==========

const DEFAULT_FALLBACK_CONFIG: FallbackConfig = {
  toolFallbacks: {
    // Excel 格式化相关
    excel_format_range: ["excel_format_cells", "excel_set_style"],
    excel_create_table: ["excel_format_as_table", "excel_set_range_values"],
    excel_add_chart: ["excel_create_chart", "excel_insert_chart"],
    // 读取相关
    excel_read_range: ["excel_get_range_values", "excel_read_selection"],
    excel_get_selection: ["excel_read_selection"],
    // 写入相关
    excel_write_range: ["excel_set_range_values", "excel_write_cell"],
    excel_set_formula: ["excel_write_cell", "excel_set_cell_formula"],
  },
  categoryFallbacks: {
    excel: "respond_to_user",
    data_analysis: "respond_to_user",
    formatting: "respond_to_user",
  },
  globalFallback: "respond_to_user",
};

// ========== 工具执行器 ==========

/**
 * 工具执行器类
 */
class ToolExecutorClass {
  private tools: ToolMap = new Map();
  private fallbackConfig: FallbackConfig = DEFAULT_FALLBACK_CONFIG;
  private executionStats: Map<string, { calls: number; successes: number; failures: number }> =
    new Map();

  /**
   * 注册单个工具
   */
  registerTool(tool: Tool): void {
    this.tools.set(tool.name, tool);
    TaskExecutionMonitor.registerTool(tool.name);
    Logger.debug("ToolExecutor", `工具已注册: ${tool.name}`);
  }

  /**
   * 批量注册工具
   */
  registerTools(tools: Tool[]): void {
    tools.forEach((tool) => this.registerTool(tool));
    Logger.info("ToolExecutor", `批量注册 ${tools.length} 个工具`);
  }

  /**
   * 取消注册工具
   */
  unregisterTool(toolName: string): boolean {
    const result = this.tools.delete(toolName);
    if (result) {
      Logger.debug("ToolExecutor", `工具已取消注册: ${toolName}`);
    }
    return result;
  }

  /**
   * 查找工具
   */
  lookupTool(toolName: string): ToolLookupResult {
    const tool = this.tools.get(toolName);

    if (tool) {
      return { found: true, tool };
    }

    // 查找备选工具
    const alternatives = this.findAlternatives(toolName);
    const suggestion = this.generateSuggestion(toolName, alternatives);

    return {
      found: false,
      alternatives,
      suggestion,
    };
  }

  /**
   * 获取工具
   */
  getTool(toolName: string): Tool | undefined {
    return this.tools.get(toolName);
  }

  /**
   * 获取所有工具
   */
  getAllTools(): Tool[] {
    return Array.from(this.tools.values());
  }

  /**
   * 检查工具是否存在
   */
  hasTool(toolName: string): boolean {
    return this.tools.has(toolName);
  }

  /**
   * 配置兜底策略
   */
  configureFallback(config: Partial<FallbackConfig>): void {
    this.fallbackConfig = {
      ...this.fallbackConfig,
      ...config,
      toolFallbacks: {
        ...this.fallbackConfig.toolFallbacks,
        ...config.toolFallbacks,
      },
      categoryFallbacks: {
        ...this.fallbackConfig.categoryFallbacks,
        ...config.categoryFallbacks,
      },
    };
  }

  /**
   * 执行工具（核心方法）
   */
  async execute(
    toolName: string,
    input: Record<string, unknown>,
    options: ToolExecutionOptions = {}
  ): Promise<EnhancedToolResult> {
    const startTime = Date.now();
    const taskId = options.taskId || `auto_${Date.now()}`;
    let retryCount = 0;
    let fallbackUsed: string | undefined;
    const warnings: string[] = [];

    // 开始监控
    if (!options.skipMonitoring) {
      TaskExecutionMonitor.startToolCall(taskId, toolName, input);
    }

    try {
      // 查找工具
      const lookup = this.lookupTool(toolName);

      if (!lookup.found) {
        // 尝试使用兜底工具
        const fallbackResult = await this.tryFallback(toolName, input, taskId, options);

        if (fallbackResult) {
          fallbackUsed = fallbackResult.toolName;
          warnings.push(`原工具 ${toolName} 未找到，使用兜底工具 ${fallbackUsed}`);

          return {
            ...fallbackResult,
            toolName,
            executionTime: Date.now() - startTime,
            retryCount,
            fallbackUsed,
            warnings,
          };
        }

        // 无可用兜底
        const errorResult: EnhancedToolResult = {
          success: false,
          output: `工具 "${toolName}" 未注册且无可用兜底。${lookup.suggestion || ""}`,
          error: `TOOL_NOT_FOUND: ${toolName}`,
          toolName,
          executionTime: Date.now() - startTime,
          retryCount,
          warnings,
        };

        if (!options.skipMonitoring) {
          TaskExecutionMonitor.failToolCall(taskId, toolName, errorResult.error!);
        }

        return errorResult;
      }

      // 参数验证
      if (options.validateParams !== false) {
        const validationResult = this.validateParameters(lookup.tool!, input);
        if (!validationResult.valid) {
          warnings.push(...validationResult.warnings);
          if (validationResult.errors.length > 0) {
            return {
              success: false,
              output: `参数验证失败: ${validationResult.errors.join("; ")}`,
              error: "PARAM_VALIDATION_FAILED",
              toolName,
              executionTime: Date.now() - startTime,
              retryCount,
              warnings,
            };
          }
        }
      }

      // 执行工具（带重试）
      const maxRetries = options.retryOnFailure ? options.maxRetries || RETRY.MAX_ATTEMPTS : 0;
      let lastError: Error | undefined;

      while (retryCount <= maxRetries) {
        try {
          const result = await this.executeWithTimeout(
            lookup.tool!,
            input,
            options.timeout || TIMEOUTS.EXCEL_OPERATION
          );

          // 更新统计
          this.updateStats(toolName, true);

          if (!options.skipMonitoring) {
            TaskExecutionMonitor.completeToolCall(taskId, toolName, result.output, result.success);
          }

          return {
            ...result,
            toolName,
            executionTime: Date.now() - startTime,
            retryCount,
            fallbackUsed,
            warnings,
          };
        } catch (error) {
          lastError = error as Error;

          // 仅在确实要重试时增加 retryCount，避免多计数
          if (retryCount < maxRetries) {
            retryCount++;
            const delay =
              (options.retryDelay || TIMEOUTS.RETRY_BASE_DELAY) * Math.pow(2, retryCount - 1);
            Logger.warn(
              "ToolExecutor",
              `工具执行失败，${delay}ms 后重试 (${retryCount}/${maxRetries})`,
              {
                toolName,
                error: lastError.message,
              }
            );
            await this.delay(delay);
            // 继续下一次重试
            continue;
          }
          // 不再重试，跳出循环
          break;
        }
      }

      // 所有重试失败，尝试兜底
      if (options.fallbackTools && options.fallbackTools.length > 0) {
        const fallbackResult = await this.tryFallback(toolName, input, taskId, {
          ...options,
          fallbackTools: options.fallbackTools,
        });

        if (fallbackResult) {
          fallbackUsed = fallbackResult.toolName;
          warnings.push(`工具 ${toolName} 执行失败，使用兜底工具 ${fallbackUsed}`);

          return {
            ...fallbackResult,
            toolName,
            executionTime: Date.now() - startTime,
            retryCount,
            fallbackUsed,
            warnings,
          };
        }
      }

      // 更新统计
      this.updateStats(toolName, false);

      if (!options.skipMonitoring) {
        TaskExecutionMonitor.failToolCall(taskId, toolName, lastError?.message || "Unknown error");
      }

      return {
        success: false,
        output: `工具执行失败: ${lastError?.message || "Unknown error"}`,
        error: lastError?.message,
        toolName,
        executionTime: Date.now() - startTime,
        retryCount,
        warnings,
      };
    } catch (error) {
      const errorMessage = (error as Error).message || "Unknown error";

      this.updateStats(toolName, false);

      if (!options.skipMonitoring) {
        TaskExecutionMonitor.failToolCall(taskId, toolName, errorMessage);
      }

      return {
        success: false,
        output: `工具执行异常: ${errorMessage}`,
        error: errorMessage,
        toolName,
        executionTime: Date.now() - startTime,
        retryCount,
        warnings,
      };
    }
  }

  /**
   * 批量执行工具
   */
  async executeBatch(
    calls: Array<{ toolName: string; input: Record<string, unknown> }>,
    options: ToolExecutionOptions = {}
  ): Promise<EnhancedToolResult[]> {
    const results: EnhancedToolResult[] = [];

    for (const call of calls) {
      const result = await this.execute(call.toolName, call.input, options);
      results.push(result);

      // 如果某个调用失败且不是可恢复的，可以选择中止
      if (!result.success && !options.retryOnFailure) {
        Logger.warn("ToolExecutor", `批量执行中止: ${call.toolName} 失败`);
        break;
      }
    }

    return results;
  }

  /**
   * 获取执行统计
   */
  getExecutionStats(): Record<
    string,
    { calls: number; successes: number; failures: number; successRate: number }
  > {
    const result: Record<
      string,
      { calls: number; successes: number; failures: number; successRate: number }
    > = {};

    this.executionStats.forEach((stats, toolName) => {
      result[toolName] = {
        ...stats,
        successRate: stats.calls > 0 ? stats.successes / stats.calls : 0,
      };
    });

    return result;
  }

  /**
   * 重置统计
   */
  resetStats(): void {
    this.executionStats.clear();
  }

  // ========== 私有方法 ==========

  /**
   * 查找备选工具
   */
  private findAlternatives(toolName: string): Tool[] {
    const alternatives: Tool[] = [];

    // 1. 检查配置的备选工具
    const configuredFallbacks = this.fallbackConfig.toolFallbacks[toolName];
    if (configuredFallbacks) {
      configuredFallbacks.forEach((name) => {
        const tool = this.tools.get(name);
        if (tool) {
          alternatives.push(tool);
        }
      });
    }

    // 2. 基于名称相似性查找
    const normalizedName = toolName.toLowerCase().replace(/[-_]/g, "");
    this.tools.forEach((tool, name) => {
      const normalizedToolName = name.toLowerCase().replace(/[-_]/g, "");
      if (
        normalizedToolName.includes(normalizedName) ||
        normalizedName.includes(normalizedToolName)
      ) {
        if (!alternatives.includes(tool)) {
          alternatives.push(tool);
        }
      }
    });

    return alternatives;
  }

  /**
   * 生成建议信息
   */
  private generateSuggestion(toolName: string, alternatives: Tool[]): string {
    if (alternatives.length === 0) {
      return `可用工具: ${Array.from(this.tools.keys()).slice(0, 10).join(", ")}${this.tools.size > 10 ? "..." : ""}`;
    }
    return `建议使用: ${alternatives.map((t) => t.name).join(", ")}`;
  }

  /**
   * 尝试使用兜底工具
   */
  private async tryFallback(
    originalTool: string,
    input: Record<string, unknown>,
    taskId: string,
    options: ToolExecutionOptions
  ): Promise<EnhancedToolResult | null> {
    // 获取可能的兜底工具列表
    const fallbackList: string[] = [];

    // 1. 选项中指定的兜底
    if (options.fallbackTools) {
      fallbackList.push(...options.fallbackTools);
    }

    // 2. 配置的工具兜底
    const configuredFallbacks = this.fallbackConfig.toolFallbacks[originalTool];
    if (configuredFallbacks) {
      fallbackList.push(...configuredFallbacks);
    }

    // 3. 全局兜底
    if (this.fallbackConfig.globalFallback) {
      fallbackList.push(this.fallbackConfig.globalFallback);
    }

    // 去重
    const uniqueFallbacks = [...new Set(fallbackList)];

    // 尝试每个兜底工具
    for (const fallbackName of uniqueFallbacks) {
      const fallbackTool = this.tools.get(fallbackName);
      if (!fallbackTool) continue;

      try {
        // 记录使用兜底
        TaskExecutionMonitor.recordFallback(
          taskId,
          originalTool,
          fallbackName,
          `原工具 ${originalTool} 不可用`
        );

        const result = await this.executeWithTimeout(
          fallbackTool,
          input,
          options.timeout || TIMEOUTS.EXCEL_OPERATION
        );

        if (result.success) {
          return {
            ...result,
            toolName: fallbackName,
            executionTime: 0, // 将由调用者计算
            retryCount: 0,
            fallbackUsed: fallbackName,
          };
        }
      } catch (error) {
        Logger.warn("ToolExecutor", `兜底工具 ${fallbackName} 执行失败`, {
          error: (error as Error).message,
        });
      }
    }

    return null;
  }

  /**
   * 带超时的工具执行
   */
  private async executeWithTimeout(
    tool: Tool,
    input: Record<string, unknown>,
    timeout: number
  ): Promise<ToolResult> {
    return new Promise<ToolResult>((resolve, reject) => {
      const timeoutId = setTimeout(() => {
        reject(new Error(`工具执行超时 (${timeout}ms)`));
      }, timeout);

      tool
        .execute(input)
        .then((result) => {
          clearTimeout(timeoutId);
          resolve(result);
        })
        .catch((error) => {
          clearTimeout(timeoutId);
          reject(error);
        });
    });
  }

  /**
   * 参数验证
   */
  private validateParameters(
    tool: Tool,
    input: Record<string, unknown>
  ): { valid: boolean; errors: string[]; warnings: string[] } {
    const errors: string[] = [];
    const warnings: string[] = [];

    // 检查必需参数
    tool.parameters.forEach((param) => {
      if (param.required && !(param.name in input)) {
        errors.push(`缺少必需参数: ${param.name}`);
      }
    });

    // 检查未知参数
    Object.keys(input).forEach((key) => {
      const paramDef = tool.parameters.find((p) => p.name === key);
      if (!paramDef) {
        warnings.push(`未知参数: ${key}`);
      }
    });

    // 类型检查
    tool.parameters.forEach((param) => {
      const value = input[param.name];
      if (value !== undefined && value !== null) {
        const expectedType = param.type;
        const actualType = Array.isArray(value) ? "array" : typeof value;

        const exp = String(expectedType);
        const act = String(actualType);
        if (exp !== act && exp !== "any") {
          // 尝试类型转换
          if (expectedType === "string" && actualType !== "string") {
            warnings.push(`参数 ${param.name} 类型不匹配，已自动转换为字符串`);
          } else if (expectedType === "number" && actualType === "string") {
            const num = Number(value);
            if (isNaN(num)) {
              errors.push(`参数 ${param.name} 无法转换为数字`);
            } else {
              warnings.push(`参数 ${param.name} 已自动转换为数字`);
            }
          } else {
            errors.push(`参数 ${param.name} 类型错误: 期望 ${expectedType}，实际 ${actualType}`);
          }
        }
      }
    });

    return {
      valid: errors.length === 0,
      errors,
      warnings,
    };
  }

  /**
   * 更新统计
   */
  private updateStats(toolName: string, success: boolean): void {
    if (!this.executionStats.has(toolName)) {
      this.executionStats.set(toolName, { calls: 0, successes: 0, failures: 0 });
    }

    const stats = this.executionStats.get(toolName)!;
    stats.calls++;
    if (success) {
      stats.successes++;
    } else {
      stats.failures++;
    }
  }

  /**
   * 延迟
   */
  private delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  /**
   * 重置执行器（用于测试）
   */
  reset(): void {
    this.tools.clear();
    this.executionStats.clear();
    this.fallbackConfig = DEFAULT_FALLBACK_CONFIG;
  }
}

// 导出单例
export const ToolExecutor = new ToolExecutorClass();

// 便捷方法导出
export const executor = {
  register: (tool: Tool) => ToolExecutor.registerTool(tool),
  registerAll: (tools: Tool[]) => ToolExecutor.registerTools(tools),
  execute: (name: string, input: Record<string, unknown>, options?: ToolExecutionOptions) =>
    ToolExecutor.execute(name, input, options),
  lookup: (name: string) => ToolExecutor.lookupTool(name),
  has: (name: string) => ToolExecutor.hasTool(name),
  get: (name: string) => ToolExecutor.getTool(name),
  getAll: () => ToolExecutor.getAllTools(),
  stats: () => ToolExecutor.getExecutionStats(),
};

export default ToolExecutor;
