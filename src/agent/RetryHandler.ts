/**
 * RetryHandler - 指数退避重试处理器 v1.0
 *
 * 借鉴自 Activepieces error-handling.ts
 *
 * 功能特性：
 * 1. 指数退避重试策略
 * 2. 可定制的错误判断逻辑
 * 3. 继续执行失败处理
 * 4. 可选的随机抖动（jitter）
 *
 * @see https://github.com/activepieces/activepieces
 */

import { Logger } from "../utils/Logger";

const MODULE_NAME = "RetryHandler";

// ==================== 类型定义 ====================

/**
 * 重试策略类型
 */
export type RetryStrategy =
  | "exponential" // 指数退避
  | "linear" // 线性增长
  | "constant"; // 固定间隔

/**
 * 重试配置
 */
export interface RetryOptions {
  /** 最大重试次数 */
  maxRetries: number;
  /** 初始延迟（毫秒） */
  initialDelayMs: number;
  /** 最大延迟（毫秒） */
  maxDelayMs: number;
  /** 退避倍数（仅用于 exponential） */
  multiplier: number;
  /** 是否添加随机抖动 */
  jitter: boolean;
  /** 抖动比例（0-1） */
  jitterFactor: number;
  /** 重试策略 */
  strategy: RetryStrategy;
  /** 可重试的错误类型 */
  retryableErrors?: Array<new (...args: unknown[]) => Error>;
  /** 自定义可重试判断 */
  isRetryable?: (error: Error) => boolean;
  /** 重试回调（用于日志记录） */
  onRetry?: (attempt: number, error: Error, delay: number) => void;
}

/**
 * 默认重试配置
 */
export const DEFAULT_RETRY_OPTIONS: RetryOptions = {
  maxRetries: 3,
  initialDelayMs: 1000,
  maxDelayMs: 30000,
  multiplier: 2,
  jitter: true,
  jitterFactor: 0.2,
  strategy: "exponential",
};

/**
 * 重试结果
 */
export interface RetryResult<T> {
  success: boolean;
  result?: T;
  error?: Error;
  attempts: number;
  totalDuration: number;
}

/**
 * 继续执行配置
 */
export interface ContinueOnFailureOptions {
  /** 失败时继续执行 */
  continueOnFailure: boolean;
  /** 默认返回值 */
  defaultValue?: unknown;
  /** 记录错误 */
  logError?: boolean;
}

// ==================== 工具函数 ====================

/**
 * 延迟执行
 */
function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * 计算延迟时间
 */
function calculateDelay(attempt: number, options: RetryOptions): number {
  let delay: number;

  switch (options.strategy) {
    case "exponential":
      delay = options.initialDelayMs * Math.pow(options.multiplier, attempt);
      break;
    case "linear":
      delay = options.initialDelayMs * (attempt + 1);
      break;
    case "constant":
    default:
      delay = options.initialDelayMs;
  }

  // 应用最大延迟限制
  delay = Math.min(delay, options.maxDelayMs);

  // 添加抖动
  if (options.jitter) {
    const jitterRange = delay * options.jitterFactor;
    const jitter = (Math.random() - 0.5) * 2 * jitterRange;
    delay = Math.max(0, delay + jitter);
  }

  return Math.round(delay);
}

/**
 * 判断错误是否可重试
 */
function isErrorRetryable(error: Error, options: RetryOptions): boolean {
  // 自定义判断优先
  if (options.isRetryable) {
    return options.isRetryable(error);
  }

  // 检查指定的可重试错误类型
  if (options.retryableErrors && options.retryableErrors.length > 0) {
    return options.retryableErrors.some((ErrorClass) => error instanceof ErrorClass);
  }

  // 默认: 网络错误、超时、临时性错误可重试
  const retryablePatterns = [
    /network/i,
    /timeout/i,
    /econnreset/i,
    /econnrefused/i,
    /enotfound/i,
    /temporarily/i,
    /service unavailable/i,
    /429/, // Too Many Requests
    /502/, // Bad Gateway
    /503/, // Service Unavailable
    /504/, // Gateway Timeout
  ];

  const errorMessage = error.message || "";
  return retryablePatterns.some((pattern) => pattern.test(errorMessage));
}

// ==================== 主要函数 ====================

/**
 * 带指数退避的重试执行
 *
 * @example
 * ```typescript
 * const result = await runWithExponentialBackoff(
 *   () => fetchData(),
 *   { maxRetries: 3, initialDelayMs: 1000 }
 * );
 *
 * if (result.success) {
 *   console.log('成功:', result.result);
 * } else {
 *   console.error('失败:', result.error);
 * }
 * ```
 */
export async function runWithExponentialBackoff<T>(
  fn: () => Promise<T>,
  options: Partial<RetryOptions> = {}
): Promise<RetryResult<T>> {
  const opts: RetryOptions = { ...DEFAULT_RETRY_OPTIONS, ...options };
  const startTime = Date.now();
  let lastError: Error | undefined;
  let attempts = 0;

  for (let attempt = 0; attempt <= opts.maxRetries; attempt++) {
    attempts = attempt + 1;

    try {
      const result = await fn();
      return {
        success: true,
        result,
        attempts,
        totalDuration: Date.now() - startTime,
      };
    } catch (error) {
      lastError = error instanceof Error ? error : new Error(String(error));

      // 检查是否可重试
      if (attempt < opts.maxRetries && isErrorRetryable(lastError, opts)) {
        const delay = calculateDelay(attempt, opts);

        Logger.warn(MODULE_NAME, "重试执行", {
          attempt: attempt + 1,
          maxRetries: opts.maxRetries,
          delay,
          error: lastError.message,
        });

        // 调用重试回调
        opts.onRetry?.(attempt + 1, lastError, delay);

        await sleep(delay);
      } else {
        // 不可重试或已达到最大次数
        break;
      }
    }
  }

  return {
    success: false,
    error: lastError,
    attempts,
    totalDuration: Date.now() - startTime,
  };
}

/**
 * 带重试的函数包装器
 *
 * @example
 * ```typescript
 * const fetchWithRetry = withRetry(fetchData, { maxRetries: 3 });
 * const data = await fetchWithRetry();
 * ```
 */
export function withRetry<T, Args extends unknown[]>(
  fn: (...args: Args) => Promise<T>,
  options: Partial<RetryOptions> = {}
): (...args: Args) => Promise<T> {
  return async (...args: Args): Promise<T> => {
    const result = await runWithExponentialBackoff(() => fn(...args), options);

    if (result.success) {
      return result.result!;
    }

    throw result.error;
  };
}

/**
 * 继续执行失败处理器
 *
 * 当 continueOnFailure 为 true 时，返回默认值而不是抛出异常
 *
 * @example
 * ```typescript
 * const result = await continueIfFailureHandler(
 *   () => riskyOperation(),
 *   { continueOnFailure: true, defaultValue: null }
 * );
 * ```
 */
export async function continueIfFailureHandler<T>(
  fn: () => Promise<T>,
  options: ContinueOnFailureOptions
): Promise<T | undefined> {
  try {
    return await fn();
  } catch (error) {
    if (options.logError !== false) {
      Logger.warn(MODULE_NAME, "操作失败，继续执行", {
        continueOnFailure: options.continueOnFailure,
        error: error instanceof Error ? error.message : String(error),
      });
    }

    if (options.continueOnFailure) {
      return options.defaultValue as T | undefined;
    }

    throw error;
  }
}

/**
 * 带重试和失败继续的完整处理器
 *
 * @example
 * ```typescript
 * const result = await robustExecute(
 *   () => unreliableOperation(),
 *   { maxRetries: 3 },
 *   { continueOnFailure: true }
 * );
 * ```
 */
export async function robustExecute<T>(
  fn: () => Promise<T>,
  retryOptions: Partial<RetryOptions> = {},
  continueOptions: Partial<ContinueOnFailureOptions> = {}
): Promise<RetryResult<T>> {
  const fullContinueOptions: ContinueOnFailureOptions = {
    continueOnFailure: false,
    logError: true,
    ...continueOptions,
  };

  const result = await runWithExponentialBackoff(fn, retryOptions);

  if (!result.success && fullContinueOptions.continueOnFailure) {
    return {
      ...result,
      success: true,
      result: fullContinueOptions.defaultValue as T,
    };
  }

  return result;
}

// ==================== 预设配置 ====================

/**
 * API 调用重试配置
 */
export const API_RETRY_OPTIONS: Partial<RetryOptions> = {
  maxRetries: 3,
  initialDelayMs: 1000,
  maxDelayMs: 10000,
  multiplier: 2,
  strategy: "exponential",
  jitter: true,
};

/**
 * Excel 操作重试配置
 */
export const EXCEL_RETRY_OPTIONS: Partial<RetryOptions> = {
  maxRetries: 2,
  initialDelayMs: 500,
  maxDelayMs: 3000,
  multiplier: 2,
  strategy: "exponential",
  jitter: false,
  isRetryable: (error) => {
    const message = error.message || "";
    // Excel 特定的可重试错误
    return (
      message.includes("GeneralException") ||
      message.includes("ItemNotFound") ||
      message.includes("InvalidArgument") === false // 参数错误不可重试
    );
  },
};

/**
 * 快速失败配置（不重试）
 */
export const NO_RETRY_OPTIONS: Partial<RetryOptions> = {
  maxRetries: 0,
};

// ==================== 导出 ====================

export default {
  runWithExponentialBackoff,
  withRetry,
  continueIfFailureHandler,
  robustExecute,
  API_RETRY_OPTIONS,
  EXCEL_RETRY_OPTIONS,
  NO_RETRY_OPTIONS,
};
