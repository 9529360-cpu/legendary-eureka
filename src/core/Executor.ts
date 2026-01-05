/**
 * Executor - 负责执行、校验、错误处理的核心模块
 *
 * 设计原则：
 * 1. 执行器是唯一直接调用ExcelService的模块
 * 2. 所有工具调用必须经过参数验证、权限检查、执行监控
 * 3. 支持批量执行和依赖管理
 * 4. 提供详细的执行日志和错误恢复机制
 */

import {
  ToolCall,
  ToolCallResult,
  ExecutionResult,
  BatchExecutionResult,
  ValidationResult,
  ValidationError,
  ValidationWarning,
  OperationDependency,
  ErrorCode,
  AppError,
  ToolCallContext,
} from "../types";
import { ExcelService } from "./ExcelService";
import { validateToolParameters, getToolById } from "./ToolRegistry";

/**
 * 执行器配置
 */
export interface ExecutorConfig {
  maxExecutionTime: number; // 最大执行时间（毫秒）
  maxRetries: number; // 最大重试次数
  enableValidation: boolean; // 是否启用参数验证
  enableLogging: boolean; // 是否启用执行日志
  allowParallelExecution: boolean; // 是否允许并行执行
  timeoutBetweenCalls: number; // 调用之间的超时时间（毫秒）
}

/**
 * 执行器状态
 */
export enum ExecutorState {
  IDLE = "idle",
  VALIDATING = "validating",
  EXECUTING = "executing",
  ROLLING_BACK = "rolling_back",
  ERROR = "error",
  COMPLETED = "completed",
}

/**
 * 执行上下文
 */
export interface ExecutionContext {
  toolCall: ToolCall;
  context: ToolCallContext;
  startTime: Date;
  endTime?: Date;
  result?: ToolCallResult;
  error?: AppError;
  retryCount: number;
}

/**
 * 执行器类
 */
export class Executor {
  private state: ExecutorState = ExecutorState.IDLE;
  private excelService: ExcelService;
  private config: ExecutorConfig;
  private executionQueue: ExecutionContext[] = [];
  private executionHistory: ExecutionContext[] = [];
  private currentExecution: ExecutionContext | null = null;

  constructor(excelService: ExcelService, config?: Partial<ExecutorConfig>) {
    this.excelService = excelService;

    this.config = {
      maxExecutionTime: 30000, // 30秒
      maxRetries: 3,
      enableValidation: true,
      enableLogging: true,
      allowParallelExecution: false,
      timeoutBetweenCalls: 100, // 100毫秒
      ...config,
    };
  }

  /**
   * 执行单个工具调用
   */
  async executeToolCall(toolCall: ToolCall, context: ToolCallContext): Promise<ToolCallResult> {
    const executionContext: ExecutionContext = {
      toolCall,
      context,
      startTime: new Date(),
      retryCount: 0,
    };

    try {
      // 1. 验证工具调用
      this.setState(ExecutorState.VALIDATING);
      const validation = await this.validateToolCall(toolCall);
      if (!validation.isValid) {
        throw new Error(`工具调用验证失败: ${validation.errors.map((e) => e.message).join(", ")}`);
      }

      // 2. 执行工具
      this.setState(ExecutorState.EXECUTING);
      const result = await this.executeWithRetry(toolCall, context);

      // 3. 记录执行结果
      executionContext.endTime = new Date();
      executionContext.result = result;
      this.executionHistory.push(executionContext);

      // 4. 更新状态
      this.setState(ExecutorState.COMPLETED);

      return result;
    } catch (error) {
      // 处理执行错误
      executionContext.endTime = new Date();
      executionContext.error = this.createAppError(
        ErrorCode.TOOL_EXECUTION_FAILED,
        error instanceof Error ? error.message : String(error),
        { toolCall, context }
      );

      this.executionHistory.push(executionContext);
      this.setState(ExecutorState.ERROR);

      throw error;
    }
  }

  /**
   * 批量执行工具调用
   */
  async executeBatch(
    toolCalls: ToolCall[],
    context: ToolCallContext,
    dependencies?: OperationDependency[]
  ): Promise<BatchExecutionResult> {
    const startTime = Date.now();
    const results: ExecutionResult[] = [];

    try {
      // 1. 验证所有工具调用
      this.setState(ExecutorState.VALIDATING);
      for (const toolCall of toolCalls) {
        const validation = await this.validateToolCall(toolCall);
        if (!validation.isValid) {
          throw new Error(`工具调用 ${toolCall.name} 验证失败`);
        }
      }

      // 2. 检查依赖关系
      if (dependencies && dependencies.length > 0) {
        const dependencyValidation = this.validateDependencies(toolCalls, dependencies);
        if (!dependencyValidation.isValid) {
          throw new Error(
            `依赖关系验证失败: ${dependencyValidation.errors.map((e) => e.message).join(", ")}`
          );
        }
      }

      // 3. 按顺序执行工具调用
      this.setState(ExecutorState.EXECUTING);
      for (let i = 0; i < toolCalls.length; i++) {
        const toolCall = toolCalls[i];

        try {
          const result = await this.executeToolCall(toolCall, context);

          results.push({
            success: result.success,
            operationId: toolCall.id,
            result: result.result,
            error: result.error,
            executionTime: result.executionTime,
            affectedRange: (result.result as any)?.affectedRange,
          });

          // 添加调用之间的延迟（如果需要）
          if (i < toolCalls.length - 1 && this.config.timeoutBetweenCalls > 0) {
            await this.delay(this.config.timeoutBetweenCalls);
          }
        } catch (error) {
          results.push({
            success: false,
            operationId: toolCall.id,
            result: null,
            error: error instanceof Error ? error.message : String(error),
            executionTime: 0,
          });

          // 如果配置了失败时停止，则中断执行
          if (this.config.maxRetries === 0) {
            break;
          }
        }
      }

      // 4. 计算统计信息
      const successful = results.filter((r) => r.success).length;
      const failed = results.filter((r) => !r.success).length;
      const totalTime = Date.now() - startTime;

      this.setState(ExecutorState.COMPLETED);

      return {
        total: toolCalls.length,
        successful,
        failed,
        results,
        totalTime,
      };
    } catch (error) {
      this.setState(ExecutorState.ERROR);
      throw error;
    }
  }

  /**
   * 验证工具调用
   */
  private async validateToolCall(toolCall: ToolCall): Promise<ValidationResult> {
    const errors: ValidationError[] = [];
    const warnings: ValidationWarning[] = [];

    // 1. 检查工具是否存在
    const tool = getToolById(toolCall.name);
    if (!tool) {
      errors.push({
        field: "toolName",
        message: `工具 ${toolCall.name} 未注册`,
        code: ErrorCode.TOOL_NOT_FOUND,
      });
      return { isValid: false, errors, warnings };
    }

    // 2. 验证参数
    if (this.config.enableValidation) {
      const validation = validateToolParameters(toolCall.name, toolCall.arguments);
      if (!validation.isValid) {
        validation.errors.forEach((error) => {
          errors.push({
            field: "parameters",
            message: error,
            code: ErrorCode.TOOL_VALIDATION_FAILED,
          });
        });
      }
    }

    // 3. 检查权限（这里可以扩展为更复杂的权限检查）
    if (tool.category === "worksheet_operation" && toolCall.name.includes("delete")) {
      warnings.push({
        field: "permission",
        message: "删除工作表是高危操作，请确认",
        severity: "high",
      });
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings,
    };
  }

  /**
   * 验证依赖关系
   */
  private validateDependencies(
    toolCalls: ToolCall[],
    dependencies: OperationDependency[]
  ): ValidationResult {
    const errors: ValidationError[] = [];
    const warnings: ValidationWarning[] = [];

    // 检查循环依赖
    const dependencyGraph = this.buildDependencyGraph(toolCalls, dependencies);
    const hasCycle = this.detectCycle(dependencyGraph);

    if (hasCycle) {
      errors.push({
        field: "dependencies",
        message: "检测到循环依赖",
        code: ErrorCode.DEPENDENCY_CYCLE,
      });
    }

    // 检查缺失的依赖
    for (const dependency of dependencies) {
      const fromExists = toolCalls.some((tc) => tc.id === dependency.from);
      const toExists = toolCalls.some((tc) => tc.id === dependency.to);

      if (!fromExists || !toExists) {
        errors.push({
          field: "dependencies",
          message: `依赖关系引用不存在的工具调用: ${dependency.from} -> ${dependency.to}`,
          code: ErrorCode.VALIDATION_FAILED,
        });
      }
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings,
    };
  }

  /**
   * 构建依赖图
   */
  private buildDependencyGraph(
    toolCalls: ToolCall[],
    dependencies: OperationDependency[]
  ): Map<string, string[]> {
    const graph = new Map<string, string[]>();

    // 初始化所有节点
    for (const toolCall of toolCalls) {
      graph.set(toolCall.id, []);
    }

    // 添加边
    for (const dependency of dependencies) {
      const neighbors = graph.get(dependency.from) || [];
      neighbors.push(dependency.to);
      graph.set(dependency.from, neighbors);
    }

    return graph;
  }

  /**
   * 检测循环依赖
   */
  private detectCycle(graph: Map<string, string[]>): boolean {
    const visited = new Set<string>();
    const recursionStack = new Set<string>();

    const dfs = (node: string): boolean => {
      if (recursionStack.has(node)) {
        return true; // 发现循环
      }

      if (visited.has(node)) {
        return false;
      }

      visited.add(node);
      recursionStack.add(node);

      const neighbors = graph.get(node) || [];
      for (const neighbor of neighbors) {
        if (dfs(neighbor)) {
          return true;
        }
      }

      recursionStack.delete(node);
      return false;
    };

    for (const node of graph.keys()) {
      if (!visited.has(node)) {
        if (dfs(node)) {
          return true;
        }
      }
    }

    return false;
  }

  /**
   * 带重试的执行
   */
  private async executeWithRetry(
    toolCall: ToolCall,
    _context: ToolCallContext
  ): Promise<ToolCallResult> {
    let lastError: Error | null = null;

    for (let attempt = 0; attempt <= this.config.maxRetries; attempt++) {
      try {
        const startTime = Date.now();
        const result = await this.excelService.executeTool(toolCall.name, toolCall.arguments);
        const executionTime = Date.now() - startTime;

        return {
          success: result.success,
          toolName: toolCall.name,
          result: result.data,
          error: result.error,
          executionTime,
          timestamp: new Date(),
        };
      } catch (error) {
        lastError = error instanceof Error ? error : new Error(String(error));

        // 如果不是最后一次尝试，等待后重试
        if (attempt < this.config.maxRetries) {
          const delayTime = Math.pow(2, attempt) * 1000; // 指数退避
          await this.delay(delayTime);
        }
      }
    }

    throw lastError || new Error("执行失败");
  }

  /**
   * 延迟函数
   */
  private delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  /**
   * 创建应用错误
   */
  private createAppError(
    code: ErrorCode,
    message: string,
    context?: Record<string, any>
  ): AppError {
    return {
      code,
      message,
      severity: "error",
      context,
      timestamp: new Date(),
    };
  }

  /**
   * 设置执行器状态
   */
  private setState(state: ExecutorState): void {
    this.state = state;

    // 记录状态变化（如果启用了日志）
    if (this.config.enableLogging) {
      console.log(`Executor状态变化: ${this.state}`);
    }
  }

  /**
   * 获取当前状态
   */
  getState(): ExecutorState {
    return this.state;
  }

  /**
   * 获取执行历史
   */
  getExecutionHistory(): ExecutionContext[] {
    return [...this.executionHistory];
  }

  /**
   * 清除执行历史
   */
  clearHistory(): void {
    this.executionHistory = [];
  }

  /**
   * 重置执行器
   */
  reset(): void {
    this.state = ExecutorState.IDLE;
    this.executionQueue = [];
    this.currentExecution = null;
  }
}
