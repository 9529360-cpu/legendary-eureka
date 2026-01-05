/**
 * ErrorBoundary - React 错误边界组件
 * v1.0.0
 *
 * 功能：
 * - 捕获子组件渲染错误
 * - 显示用户友好的错误界面
 * - 提供重试机制
 * - 错误上报（可扩展）
 */

import React, { Component, ErrorInfo, ReactNode } from "react";
import { Logger } from "../utils/Logger";

/** 错误边界状态 */
interface ErrorBoundaryState {
  hasError: boolean;
  error: Error | null;
  errorInfo: ErrorInfo | null;
}

/** 错误边界属性 */
interface ErrorBoundaryProps {
  children: ReactNode;
  /** 自定义错误回退UI */
  fallback?: ReactNode;
  /** 错误回调 */
  onError?: (error: Error, errorInfo: ErrorInfo) => void;
  /** 模块名称（用于日志） */
  moduleName?: string;
}

/**
 * React 错误边界组件
 */
export class ErrorBoundary extends Component<ErrorBoundaryProps, ErrorBoundaryState> {
  constructor(props: ErrorBoundaryProps) {
    super(props);
    this.state = {
      hasError: false,
      error: null,
      errorInfo: null,
    };
  }

  static getDerivedStateFromError(error: Error): Partial<ErrorBoundaryState> {
    return { hasError: true, error };
  }

  componentDidCatch(error: Error, errorInfo: ErrorInfo): void {
    const moduleName = this.props.moduleName || "ErrorBoundary";
    
    // 记录错误日志
    Logger.error(moduleName, "React 组件渲染错误", {
      message: error.message,
      stack: error.stack,
      componentStack: errorInfo.componentStack,
    });

    // 更新状态
    this.setState({ errorInfo });

    // 调用错误回调
    if (this.props.onError) {
      this.props.onError(error, errorInfo);
    }
  }

  /**
   * 重置错误状态（重试）
   */
  handleRetry = (): void => {
    this.setState({
      hasError: false,
      error: null,
      errorInfo: null,
    });
  };

  render(): ReactNode {
    if (this.state.hasError) {
      // 使用自定义回退UI
      if (this.props.fallback) {
        return this.props.fallback;
      }

      // 默认错误UI
      return (
        <div style={styles.container}>
          <div style={styles.icon}>⚠️</div>
          <h2 style={styles.title}>出现了一些问题</h2>
          <p style={styles.message}>
            {this.state.error?.message || "未知错误"}
          </p>
          <button style={styles.button} onClick={this.handleRetry}>
            重试
          </button>
          <details style={styles.details}>
            <summary style={styles.summary}>查看详细信息</summary>
            <pre style={styles.stack}>
              {this.state.error?.stack}
              {this.state.errorInfo?.componentStack}
            </pre>
          </details>
        </div>
      );
    }

    return this.props.children;
  }
}

/** 样式定义 */
const styles: Record<string, React.CSSProperties> = {
  container: {
    padding: "24px",
    textAlign: "center",
    backgroundColor: "#fff8f8",
    border: "1px solid #ffcccc",
    borderRadius: "8px",
    margin: "16px",
  },
  icon: {
    fontSize: "48px",
    marginBottom: "16px",
  },
  title: {
    color: "#cc0000",
    fontSize: "18px",
    margin: "0 0 8px 0",
  },
  message: {
    color: "#666",
    fontSize: "14px",
    margin: "0 0 16px 0",
  },
  button: {
    backgroundColor: "#0078d4",
    color: "white",
    border: "none",
    padding: "8px 24px",
    borderRadius: "4px",
    cursor: "pointer",
    fontSize: "14px",
  },
  details: {
    marginTop: "16px",
    textAlign: "left",
  },
  summary: {
    cursor: "pointer",
    color: "#666",
    fontSize: "12px",
  },
  stack: {
    backgroundColor: "#f5f5f5",
    padding: "12px",
    borderRadius: "4px",
    fontSize: "11px",
    overflow: "auto",
    maxHeight: "200px",
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
  },
};

export default ErrorBoundary;
