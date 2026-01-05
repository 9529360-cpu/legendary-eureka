/**
 * GlobalErrorHandler - 全局错误处理器
 * v1.0.0
 *
 * 功能：
 * - 捕获未处理的 Promise rejection
 * - 捕获全局 JavaScript 错误
 * - 统一错误上报
 * - 与 Logger 集成
 */

import { Logger } from "./Logger";

/** 全局错误类型 */
export type GlobalErrorType = "unhandled-rejection" | "runtime-error" | "network-error";

/** 全局错误信息 */
export interface GlobalError {
  type: GlobalErrorType;
  message: string;
  stack?: string;
  timestamp: Date;
  url?: string;
  line?: number;
  column?: number;
}

/** 错误监听器类型 */
export type ErrorListener = (error: GlobalError) => void;

/**
 * 全局错误处理器
 */
class GlobalErrorHandlerClass {
  private listeners: ErrorListener[] = [];
  private errorHistory: GlobalError[] = [];
  private maxHistorySize = 50;
  private isInitialized = false;

  /**
   * 初始化全局错误处理
   */
  initialize(): void {
    if (this.isInitialized) {
      Logger.warn("GlobalErrorHandler", "已经初始化过，跳过重复初始化");
      return;
    }

    // 监听未处理的 Promise rejection
    window.addEventListener("unhandledrejection", this.handleUnhandledRejection);

    // 监听全局错误
    window.addEventListener("error", this.handleGlobalError);

    this.isInitialized = true;
    Logger.info("GlobalErrorHandler", "全局错误处理器已初始化");
  }

  /**
   * 清理（销毁）
   */
  destroy(): void {
    window.removeEventListener("unhandledrejection", this.handleUnhandledRejection);
    window.removeEventListener("error", this.handleGlobalError);
    this.isInitialized = false;
    Logger.info("GlobalErrorHandler", "全局错误处理器已销毁");
  }

  /**
   * 添加错误监听器
   */
  addListener(listener: ErrorListener): () => void {
    this.listeners.push(listener);
    return () => {
      this.listeners = this.listeners.filter((l) => l !== listener);
    };
  }

  /**
   * 获取错误历史
   */
  getHistory(): GlobalError[] {
    return [...this.errorHistory];
  }

  /**
   * 清除错误历史
   */
  clearHistory(): void {
    this.errorHistory = [];
  }

  /**
   * 处理未处理的 Promise rejection
   */
  private handleUnhandledRejection = (event: PromiseRejectionEvent): void => {
    const error: GlobalError = {
      type: "unhandled-rejection",
      message: this.extractErrorMessage(event.reason),
      stack: event.reason?.stack,
      timestamp: new Date(),
    };

    Logger.error("GlobalErrorHandler", "未捕获的 Promise rejection", {
      message: error.message,
      stack: error.stack?.substring(0, 500),
    });

    this.recordError(error);

    // 阻止默认行为（控制台显示）
    event.preventDefault();
  };

  /**
   * 处理全局 JavaScript 错误
   */
  private handleGlobalError = (event: ErrorEvent): void => {
    const error: GlobalError = {
      type: "runtime-error",
      message: event.message,
      stack: event.error?.stack,
      timestamp: new Date(),
      url: event.filename,
      line: event.lineno,
      column: event.colno,
    };

    Logger.error("GlobalErrorHandler", "全局运行时错误", {
      message: error.message,
      url: error.url,
      line: error.line,
      column: error.column,
    });

    this.recordError(error);
  };

  /**
   * 从各种格式中提取错误消息
   */
  private extractErrorMessage(reason: unknown): string {
    if (reason instanceof Error) {
      return reason.message;
    }
    if (typeof reason === "string") {
      return reason;
    }
    if (reason && typeof reason === "object" && "message" in reason) {
      return String((reason as { message: unknown }).message);
    }
    return String(reason);
  }

  /**
   * 记录错误并通知监听器
   */
  private recordError(error: GlobalError): void {
    // 添加到历史
    this.errorHistory.push(error);
    if (this.errorHistory.length > this.maxHistorySize) {
      this.errorHistory.shift();
    }

    // 通知所有监听器
    for (const listener of this.listeners) {
      try {
        listener(error);
      } catch (e) {
        Logger.error("GlobalErrorHandler", "错误监听器执行失败", e);
      }
    }
  }
}

/** 导出单例 */
export const GlobalErrorHandler = new GlobalErrorHandlerClass();

export default GlobalErrorHandler;
