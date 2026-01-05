/**
 * ErrorHandler - 统一的错误处理机制
 *
 * 设计原则：
 * 1. 集中管理所有错误处理逻辑
 * 2. 提供用户友好的错误消息
 * 3. 支持错误分类和严重性分级
 * 4. 集成日志记录和错误报告
 * 5. 支持错误恢复和重试机制
 */

import { ErrorCode, AppError, ValidationError, ValidationWarning } from "../types";
import { ERROR_CONFIG } from "../config";

/**
 * 错误严重性级别
 */
export enum ErrorSeverity {
  INFO = "info",
  WARNING = "warning",
  ERROR = "error",
  CRITICAL = "critical",
}

/**
 * 错误类别
 */
export enum ErrorCategory {
  EXCEL = "excel",
  AI = "ai",
  VALIDATION = "validation",
  NETWORK = "network",
  SECURITY = "security",
  SYSTEM = "system",
  USER = "user",
}

/**
 * 错误处理选项
 */
export interface ErrorHandlingOptions {
  showToUser?: boolean;
  logToConsole?: boolean;
  throwOriginal?: boolean;
  retryable?: boolean;
  maxRetries?: number;
  retryDelay?: number;
  fallbackAction?: () => Promise<any>;
  userFriendlyMessage?: string;
}

/**
 * 错误上下文
 */
export interface ErrorContext {
  operation?: string;
  parameters?: Record<string, any>;
  userId?: string;
  conversationId?: string;
  timestamp?: Date;
  stackTrace?: string;
  additionalInfo?: Record<string, any>;
}

/**
 * 错误处理结果
 */
export interface ErrorHandlingResult {
  handled: boolean;
  recovered: boolean;
  error?: AppError;
  result?: any;
  userMessage?: string;
  shouldRetry?: boolean;
  retryAfter?: number;
}

/**
 * 错误处理器类
 */
export class ErrorHandler {
  private static instance: ErrorHandler;
  private errorHistory: AppError[] = [];
  private maxHistorySize = 1000;

  private constructor() {
    // 私有构造函数，确保单例模式
  }

  /**
   * 获取错误处理器实例
   */
  public static getInstance(): ErrorHandler {
    if (!ErrorHandler.instance) {
      ErrorHandler.instance = new ErrorHandler();
    }
    return ErrorHandler.instance;
  }

  /**
   * 处理错误
   */
  public async handleError(
    error: unknown,
    context: ErrorContext = {},
    options: ErrorHandlingOptions = {}
  ): Promise<ErrorHandlingResult> {
    const appError = this.normalizeError(error, context);
    this.addToHistory(appError);

    // 应用配置选项
    const finalOptions = this.mergeWithConfig(options);

    // 记录错误
    if (finalOptions.logToConsole) {
      this.logError(appError, context);
    }

    // 检查是否为可抑制的错误
    if (this.shouldSuppressError(appError)) {
      return {
        handled: true,
        recovered: false,
        error: appError,
        userMessage: "操作已取消或忽略",
      };
    }

    // 尝试恢复
    let recovered = false;
    let recoveryResult: any = undefined;

    if (finalOptions.fallbackAction) {
      try {
        recoveryResult = await finalOptions.fallbackAction();
        recovered = true;
      } catch (fallbackError) {
        // 回退操作也失败
        const fallbackAppError = this.normalizeError(fallbackError, {
          ...context,
          operation: "fallback_action",
        });
        this.addToHistory(fallbackAppError);
      }
    }

    // 生成用户消息
    const userMessage = this.generateUserMessage(appError, finalOptions, recovered);

    // 确定是否需要重试
    const shouldRetry = finalOptions.retryable && this.isRetryableError(appError);
    const retryAfter = shouldRetry ? finalOptions.retryDelay || 1000 : undefined;

    // 是否抛出原始错误
    if (finalOptions.throwOriginal && !recovered) {
      throw error;
    }

    return {
      handled: true,
      recovered,
      error: appError,
      result: recoveryResult,
      userMessage,
      shouldRetry,
      retryAfter,
    };
  }

  /**
   * 处理验证错误
   */
  public handleValidationErrors(
    errors: ValidationError[],
    warnings: ValidationWarning[] = [],
    context: ErrorContext = {}
  ): ErrorHandlingResult {
    if (errors.length === 0) {
      return {
        handled: true,
        recovered: true,
        userMessage: "验证通过",
      };
    }

    const validationError = this.createAppError(
      ErrorCode.VALIDATION_FAILED,
      `验证失败: ${errors.map((e) => e.message).join(", ")}`,
      {
        ...context,
        validationErrors: errors,
        validationWarnings: warnings,
      }
    );

    this.addToHistory(validationError);

    const userMessage = this.generateValidationUserMessage(errors, warnings);

    return {
      handled: true,
      recovered: false,
      error: validationError,
      userMessage,
    };
  }

  /**
   * 创建应用错误
   */
  public createAppError(
    code: ErrorCode,
    message: string,
    context?: Record<string, any>,
    severity: ErrorSeverity = ErrorSeverity.ERROR
  ): AppError {
    return {
      code,
      message: this.truncateErrorMessage(message),
      severity: severity === ErrorSeverity.CRITICAL ? "error" : severity,
      context,
      timestamp: new Date(),
    };
  }

  /**
   * 获取错误历史
   */
  public getErrorHistory(filter?: {
    category?: ErrorCategory;
    severity?: ErrorSeverity;
    startDate?: Date;
    endDate?: Date;
  }): AppError[] {
    let filtered = [...this.errorHistory];

    if (filter?.category) {
      filtered = filtered.filter((error) => this.getErrorCategory(error) === filter.category);
    }

    if (filter?.severity) {
      filtered = filtered.filter((error) => error.severity === filter.severity);
    }

    if (filter?.startDate) {
      filtered = filtered.filter((error) => error.timestamp >= filter.startDate!);
    }

    if (filter?.endDate) {
      filtered = filtered.filter((error) => error.timestamp <= filter.endDate!);
    }

    return filtered;
  }

  /**
   * 清除错误历史
   */
  public clearHistory(): void {
    this.errorHistory = [];
  }

  /**
   * 获取错误统计信息
   */
  public getErrorStatistics(): {
    total: number;
    byCategory: Record<ErrorCategory, number>;
    bySeverity: Record<ErrorSeverity, number>;
    last24Hours: number;
  } {
    const now = new Date();
    const twentyFourHoursAgo = new Date(now.getTime() - 24 * 60 * 60 * 1000);

    const byCategory: Record<ErrorCategory, number> = {
      [ErrorCategory.EXCEL]: 0,
      [ErrorCategory.AI]: 0,
      [ErrorCategory.VALIDATION]: 0,
      [ErrorCategory.NETWORK]: 0,
      [ErrorCategory.SECURITY]: 0,
      [ErrorCategory.SYSTEM]: 0,
      [ErrorCategory.USER]: 0,
    };

    const bySeverity: Record<ErrorSeverity, number> = {
      [ErrorSeverity.INFO]: 0,
      [ErrorSeverity.WARNING]: 0,
      [ErrorSeverity.ERROR]: 0,
      [ErrorSeverity.CRITICAL]: 0,
    };

    let last24Hours = 0;

    for (const error of this.errorHistory) {
      // 分类统计
      const category = this.getErrorCategory(error);
      byCategory[category] = (byCategory[category] || 0) + 1;

      // 严重性统计
      const severity = this.getErrorSeverity(error);
      bySeverity[severity] = (bySeverity[severity] || 0) + 1;

      // 24小时内统计
      if (error.timestamp >= twentyFourHoursAgo) {
        last24Hours++;
      }
    }

    return {
      total: this.errorHistory.length,
      byCategory,
      bySeverity,
      last24Hours,
    };
  }

  /**
   * 规范化错误
   */
  private normalizeError(error: unknown, context: ErrorContext): AppError {
    if (this.isAppError(error)) {
      return error;
    }

    let message: string;
    let code: ErrorCode;

    if (error instanceof Error) {
      message = error.message;
      code = this.determineErrorCode(error);
    } else if (typeof error === "string") {
      message = error;
      code = ErrorCode.OPERATION_FAILED;
    } else {
      message = "未知错误";
      code = ErrorCode.OPERATION_FAILED;
    }

    return this.createAppError(
      code,
      message,
      {
        ...context,
        originalError: error,
        stackTrace: error instanceof Error ? error.stack : undefined,
      },
      this.determineErrorSeverity(code)
    );
  }

  /**
   * 确定错误代码
   */
  private determineErrorCode(error: Error): ErrorCode {
    const message = error.message.toLowerCase();

    if (message.includes("excel") || message.includes("office")) {
      return ErrorCode.EXCEL_NOT_READY;
    } else if (message.includes("network") || message.includes("fetch")) {
      return ErrorCode.AI_SERVICE_UNAVAILABLE;
    } else if (message.includes("api key") || message.includes("authentication")) {
      return ErrorCode.INVALID_API_KEY;
    } else if (message.includes("rate limit") || message.includes("too many requests")) {
      return ErrorCode.RATE_LIMIT_EXCEEDED;
    } else if (message.includes("validation") || message.includes("invalid")) {
      return ErrorCode.VALIDATION_FAILED;
    } else if (message.includes("timeout") || message.includes("timed out")) {
      return ErrorCode.EXECUTION_TIMEOUT;
    } else {
      return ErrorCode.OPERATION_FAILED;
    }
  }

  /**
   * 确定错误严重性
   */
  private determineErrorSeverity(code: ErrorCode): ErrorSeverity {
    switch (code) {
      case ErrorCode.EXCEL_NOT_READY:
      case ErrorCode.INVALID_API_KEY:
        return ErrorSeverity.ERROR;
      case ErrorCode.AI_SERVICE_UNAVAILABLE:
      case ErrorCode.RATE_LIMIT_EXCEEDED:
        return ErrorSeverity.WARNING;
      case ErrorCode.VALIDATION_FAILED:
      case ErrorCode.MISSING_PARAMETER:
      case ErrorCode.INVALID_PARAMETER:
        return ErrorSeverity.INFO;
      case ErrorCode.TOOL_EXECUTION_FAILED:
      case ErrorCode.EXECUTION_TIMEOUT:
      case ErrorCode.DEPENDENCY_CYCLE:
      case ErrorCode.INSUFFICIENT_PERMISSIONS:
        return ErrorSeverity.CRITICAL;
      default:
        return ErrorSeverity.ERROR;
    }
  }

  /**
   * 获取错误类别
   */
  private getErrorCategory(error: AppError): ErrorCategory {
    const code = error.code;

    if (code.startsWith("EXCEL")) {
      return ErrorCategory.EXCEL;
    } else if (code.startsWith("AI")) {
      return ErrorCategory.AI;
    } else if (code.startsWith("VAL")) {
      return ErrorCategory.VALIDATION;
    } else if (code.startsWith("TOOL")) {
      return ErrorCategory.SYSTEM;
    } else if (code.startsWith("EXEC")) {
      return ErrorCategory.SYSTEM;
    } else if (error.message.includes("network") || error.message.includes("CORS")) {
      return ErrorCategory.NETWORK;
    } else if (error.message.includes("permission") || error.message.includes("security")) {
      return ErrorCategory.SECURITY;
    } else {
      return ErrorCategory.USER;
    }
  }

  /**
   * 获取错误严重性
   */
  private getErrorSeverity(error: AppError): ErrorSeverity {
    switch (error.severity) {
      case "error":
        return ErrorSeverity.ERROR;
      case "warning":
        return ErrorSeverity.WARNING;
      case "info":
        return ErrorSeverity.INFO;
      default:
        return ErrorSeverity.ERROR;
    }
  }

  /**
   * 检查是否为AppError
   */
  private isAppError(error: any): error is AppError {
    return (
      error &&
      typeof error === "object" &&
      "code" in error &&
      "message" in error &&
      "severity" in error &&
      "timestamp" in error
    );
  }

  /**
   * 合并配置选项
   */
  private mergeWithConfig(options: ErrorHandlingOptions): ErrorHandlingOptions {
    return {
      showToUser: ERROR_CONFIG.SHOW_USER_FRIENDLY_MESSAGES,
      logToConsole: ERROR_CONFIG.LOG_TO_CONSOLE,
      throwOriginal: false,
      retryable: true,
      maxRetries: 3,
      retryDelay: 1000,
      ...options,
    };
  }

  /**
   * 记录错误
   */
  private logError(error: AppError, context: ErrorContext): void {
    const logEntry = {
      timestamp: error.timestamp.toISOString(),
      code: error.code,
      message: error.message,
      severity: error.severity,
      category: this.getErrorCategory(error),
      context: {
        ...context,
        ...error.context,
      },
    };

    console.error("[ErrorHandler]", logEntry);

    // 这里可以扩展为发送到日志服务
    if (ERROR_CONFIG.LOG_TO_CONSOLE) {
      console.error(JSON.stringify(logEntry, null, 2));
    }
  }

  /**
   * 生成用户消息
   */
  private generateUserMessage(
    error: AppError,
    options: ErrorHandlingOptions,
    recovered: boolean
  ): string {
    if (!options.showToUser) {
      return "";
    }

    if (recovered) {
      return options.userFriendlyMessage || "操作已恢复完成";
    }

    // 使用用户友好的消息或生成默认消息
    if (options.userFriendlyMessage) {
      return options.userFriendlyMessage;
    }

    // 根据错误代码生成用户友好的消息
    switch (error.code) {
      case ErrorCode.EXCEL_NOT_READY:
        return "Excel未准备好，请确保Excel已打开并加载了工作簿";
      case ErrorCode.INVALID_RANGE:
        return "指定的单元格范围无效，请检查范围格式（如A1:B10）";
      case ErrorCode.AI_SERVICE_UNAVAILABLE:
        return "AI服务暂时不可用，请稍后重试";
      case ErrorCode.INVALID_API_KEY:
        return "API密钥无效，请检查配置";
      case ErrorCode.RATE_LIMIT_EXCEEDED:
        return "请求过于频繁，请稍后重试";
      case ErrorCode.VALIDATION_FAILED:
        return "输入验证失败，请检查输入参数";
      case ErrorCode.TOOL_EXECUTION_FAILED:
        return "工具执行失败，请检查操作参数";
      case ErrorCode.EXECUTION_TIMEOUT:
        return "操作超时，请稍后重试或简化操作";
      default:
        return `操作失败: ${error.message.substring(0, 100)}${error.message.length > 100 ? "..." : ""}`;
    }
  }

  /**
   * 生成验证用户消息
   */
  private generateValidationUserMessage(
    errors: ValidationError[],
    warnings: ValidationWarning[]
  ): string {
    const errorMessages = errors.map((e) => e.message).join("；");
    const warningMessages = warnings.map((w) => w.message).join("；");

    let message = "";

    if (errors.length > 0) {
      message += `验证错误：${errorMessages}`;
    }

    if (warnings.length > 0) {
      if (message) message += "\n";
      message += `警告：${warningMessages}`;
    }

    return message || "验证通过";
  }

  /**
   * 检查是否应该抑制错误
   */
  private shouldSuppressError(error: AppError): boolean {
    if (!ERROR_CONFIG.SUPPRESSED_ERRORS || ERROR_CONFIG.SUPPRESSED_ERRORS.length === 0) {
      return false;
    }

    const errorMessage = error.message.toLowerCase();
    return ERROR_CONFIG.SUPPRESSED_ERRORS.some((suppressedError) =>
      errorMessage.includes(suppressedError.toLowerCase())
    );
  }

  /**
   * 检查是否为可重试错误
   */
  private isRetryableError(error: AppError): boolean {
    const nonRetryableCodes = [
      ErrorCode.INVALID_API_KEY,
      ErrorCode.INVALID_RANGE,
      ErrorCode.VALIDATION_FAILED,
      ErrorCode.MISSING_PARAMETER,
      ErrorCode.INVALID_PARAMETER,
    ];

    return !nonRetryableCodes.includes(error.code as ErrorCode);
  }

  /**
   * 截断错误消息
   */
  private truncateErrorMessage(message: string): string {
    const maxLength = ERROR_CONFIG.MAX_ERROR_LENGTH;
    if (message.length <= maxLength) {
      return message;
    }
    return message.substring(0, maxLength) + "...";
  }

  /**
   * 添加到历史记录
   */
  private addToHistory(error: AppError): void {
    this.errorHistory.unshift(error); // 添加到开头

    // 限制历史记录大小
    if (this.errorHistory.length > this.maxHistorySize) {
      this.errorHistory = this.errorHistory.slice(0, this.maxHistorySize);
    }
  }
}

/**
 * 全局错误处理函数
 */
export function handleGlobalError(
  error: unknown,
  context: ErrorContext = {},
  options: ErrorHandlingOptions = {}
): Promise<ErrorHandlingResult> {
  return ErrorHandler.getInstance().handleError(error, context, options);
}

/**
 * 处理验证错误的便捷函数
 */
export function handleValidation(
  errors: ValidationError[],
  warnings: ValidationWarning[] = [],
  context: ErrorContext = {}
): ErrorHandlingResult {
  return ErrorHandler.getInstance().handleValidationErrors(errors, warnings, context);
}
