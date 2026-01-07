/**
 * useStreamingAgent - 流式 Agent Hook v4.1
 *
 * 提供流式输出能力的 React Hook，用户发送消息后立即看到反馈
 *
 * 特性：
 * 1. 实时流式输出
 * 2. 支持取消操作
 * 3. 进度追踪
 * 4. 错误恢复反馈
 *
 * @module hooks/useStreamingAgent
 */

import * as React from "react";
import {
  StreamingAgentExecutor,
  createStreamingExecutor,
  StreamChunk,
  StreamExecutionResult,
} from "../../agent";
import { ParseContext } from "../../agent/IntentParser";

// ========== 类型定义 ==========

/**
 * 流式消息
 */
export interface StreamMessage {
  /** 消息 ID */
  id: string;

  /** 角色 */
  role: "user" | "assistant";

  /** 内容（逐步累积） */
  content: string;

  /** 状态 */
  status: "streaming" | "complete" | "error" | "cancelled";

  /** 进度 (0-100) */
  progress: number;

  /** 时间戳 */
  timestamp: Date;

  /** 步骤信息 */
  steps?: StreamStepInfo[];
}

/**
 * 步骤信息
 */
export interface StreamStepInfo {
  id: string;
  description: string;
  status: "pending" | "running" | "done" | "error" | "skipped";
  output?: string;
  error?: string;
}

/**
 * Hook 状态
 */
export interface StreamingAgentState {
  /** 是否正在运行 */
  isStreaming: boolean;

  /** 当前消息列表 */
  messages: StreamMessage[];

  /** 当前进度 */
  progress: number;

  /** 当前阶段 */
  phase: "idle" | "thinking" | "executing" | "complete" | "error";

  /** 错误信息 */
  error: string | null;
}

/**
 * Hook 选项
 */
export interface UseStreamingAgentOptions {
  /** 是否启用错误恢复 */
  enableRecovery?: boolean;

  /** 进度回调 */
  onProgress?: (progress: number) => void;

  /** 消息回调 */
  onMessage?: (message: StreamMessage) => void;

  /** 完成回调 */
  onComplete?: (result: StreamExecutionResult) => void;

  /** 错误回调 */
  onError?: (error: Error) => void;
}

/**
 * Hook 返回值
 */
export interface UseStreamingAgentReturn {
  /** 状态 */
  state: StreamingAgentState;

  /** 发送消息 */
  sendMessage: (message: string, context?: Partial<ParseContext>) => Promise<void>;

  /** 取消执行 */
  cancel: () => void;

  /** 清空消息 */
  clearMessages: () => void;

  /** 是否可以取消 */
  canCancel: boolean;
}

// ========== Hook 实现 ==========

/**
 * 流式 Agent Hook
 */
export function useStreamingAgent(
  options: UseStreamingAgentOptions = {}
): UseStreamingAgentReturn {
  const { enableRecovery = true, onProgress, onMessage, onComplete, onError } = options;

  // 执行器实例
  const executorRef = React.useRef<StreamingAgentExecutor | null>(null);

  // 取消控制器
  const abortControllerRef = React.useRef<AbortController | null>(null);

  // 状态
  const [state, setState] = React.useState<StreamingAgentState>({
    isStreaming: false,
    messages: [],
    progress: 0,
    phase: "idle",
    error: null,
  });

  // 回调 refs
  const callbackRefs = React.useRef({ onProgress, onMessage, onComplete, onError });
  callbackRefs.current = { onProgress, onMessage, onComplete, onError };

  // 初始化执行器
  React.useEffect(() => {
    if (!executorRef.current) {
      try {
        executorRef.current = createStreamingExecutor();
        console.log("[useStreamingAgent] 流式执行器已创建");
      } catch (error) {
        console.error("[useStreamingAgent] 创建执行器失败:", error);
      }
    }
  }, []);

  // 生成消息 ID
  const generateId = (): string => {
    return `msg_${Date.now()}_${Math.random().toString(36).substring(2, 8)}`;
  };

  // 发送消息
  const sendMessage = React.useCallback(
    async (message: string, context: Partial<ParseContext> = {}) => {
      if (!executorRef.current) {
        console.error("[useStreamingAgent] 执行器未初始化");
        return;
      }

      // 创建取消控制器
      abortControllerRef.current = new AbortController();

      // 添加用户消息
      const userMessage: StreamMessage = {
        id: generateId(),
        role: "user",
        content: message,
        status: "complete",
        progress: 100,
        timestamp: new Date(),
      };

      // 创建助手消息（流式）
      const assistantMessage: StreamMessage = {
        id: generateId(),
        role: "assistant",
        content: "",
        status: "streaming",
        progress: 0,
        timestamp: new Date(),
        steps: [],
      };

      setState((prev) => ({
        ...prev,
        isStreaming: true,
        messages: [...prev.messages, userMessage, assistantMessage],
        progress: 0,
        phase: "thinking",
        error: null,
      }));

      try {
        // 构建完整上下文
        const fullContext: ParseContext = {
          userMessage: message,
          ...context,
        };

        // 流式执行
        const stream = executorRef.current.executeStream(fullContext, {
          enableRecovery,
          signal: abortControllerRef.current.signal,
        });

        let contentBuffer = "";
        const steps: StreamStepInfo[] = [];

        for await (const chunk of stream) {
          // 处理每个 chunk
          const updatedMessage = processChunk(chunk, assistantMessage, contentBuffer, steps);
          contentBuffer = updatedMessage.content;

          // 更新状态
          setState((prev) => {
            const newMessages = [...prev.messages];
            const lastIndex = newMessages.length - 1;
            newMessages[lastIndex] = {
              ...updatedMessage,
              steps: [...steps],
            };

            return {
              ...prev,
              messages: newMessages,
              progress: chunk.progress || prev.progress,
              phase: getPhaseFromChunk(chunk),
            };
          });

          // 回调
          callbackRefs.current.onProgress?.(chunk.progress || 0);
        }

        // 获取最终结果
        const finalResult: StreamExecutionResult = {
          success: true,
          message: contentBuffer,
          stepsExecuted: steps.filter((s) => s.status === "done").length,
          duration: Date.now() - assistantMessage.timestamp.getTime(),
        };

        // 标记完成
        setState((prev) => {
          const newMessages = [...prev.messages];
          const lastIndex = newMessages.length - 1;
          newMessages[lastIndex] = {
            ...newMessages[lastIndex],
            status: "complete",
            progress: 100,
          };

          return {
            ...prev,
            isStreaming: false,
            messages: newMessages,
            progress: 100,
            phase: "complete",
          };
        });

        callbackRefs.current.onComplete?.(finalResult);
      } catch (error) {
        const errorMsg = error instanceof Error ? error.message : String(error);

        // 检查是否是取消
        if (abortControllerRef.current?.signal.aborted) {
          setState((prev) => {
            const newMessages = [...prev.messages];
            const lastIndex = newMessages.length - 1;
            newMessages[lastIndex] = {
              ...newMessages[lastIndex],
              content: prev.messages[lastIndex].content || "操作已取消",
              status: "cancelled",
            };

            return {
              ...prev,
              isStreaming: false,
              messages: newMessages,
              phase: "idle",
            };
          });
        } else {
          setState((prev) => {
            const newMessages = [...prev.messages];
            const lastIndex = newMessages.length - 1;
            newMessages[lastIndex] = {
              ...newMessages[lastIndex],
              content: `错误: ${errorMsg}`,
              status: "error",
            };

            return {
              ...prev,
              isStreaming: false,
              messages: newMessages,
              phase: "error",
              error: errorMsg,
            };
          });

          callbackRefs.current.onError?.(error instanceof Error ? error : new Error(errorMsg));
        }
      } finally {
        abortControllerRef.current = null;
      }
    },
    [enableRecovery]
  );

  // 取消执行
  const cancel = React.useCallback(() => {
    if (abortControllerRef.current) {
      abortControllerRef.current.abort();
      console.log("[useStreamingAgent] 执行已取消");
    }
  }, []);

  // 清空消息
  const clearMessages = React.useCallback(() => {
    setState((prev) => ({
      ...prev,
      messages: [],
      progress: 0,
      phase: "idle",
      error: null,
    }));
  }, []);

  return {
    state,
    sendMessage,
    cancel,
    clearMessages,
    canCancel: state.isStreaming,
  };
}

// ========== 辅助函数 ==========

/**
 * 处理流式 chunk
 */
function processChunk(
  chunk: StreamChunk,
  message: StreamMessage,
  currentContent: string,
  steps: StreamStepInfo[]
): StreamMessage {
  let newContent = currentContent;

  switch (chunk.type) {
    case "status":
    case "thinking":
      // 状态更新，不改变内容
      break;

    case "intent":
      newContent = `${chunk.content}\n`;
      break;

    case "plan":
      newContent += `\n📋 ${chunk.content}\n`;
      // 初始化步骤
      const planData = chunk.data as { steps?: Array<{ id: string; description: string }> };
      if (planData?.steps) {
        steps.length = 0;
        planData.steps.forEach((s) => {
          steps.push({
            id: s.id,
            description: s.description,
            status: "pending",
          });
        });
      }
      break;

    case "step:start":
      const startData = chunk.data as { stepIndex?: number; stepId?: string };
      if (startData?.stepId) {
        const step = steps.find((s) => s.id === startData.stepId);
        if (step) {
          step.status = "running";
        }
      }
      break;

    case "step:done":
      newContent += `  ${chunk.content}\n`;
      const doneData = chunk.data as { stepIndex?: number; stepId?: string; output?: string };
      if (doneData?.stepId) {
        const step = steps.find((s) => s.id === doneData.stepId);
        if (step) {
          step.status = "done";
          step.output = doneData.output;
        }
      }
      break;

    case "step:error":
      newContent += `  ${chunk.content}\n`;
      const errorData = chunk.data as { stepIndex?: number; stepId?: string; error?: string };
      if (errorData?.stepId) {
        const step = steps.find((s) => s.id === errorData.stepId);
        if (step) {
          step.status = "error";
          step.error = errorData.error;
        }
      }
      break;

    case "step:recovery":
      newContent += `  ${chunk.content}\n`;
      break;

    case "message":
    case "complete":
      newContent = chunk.content;
      break;

    case "error":
      newContent = `❌ ${chunk.content}`;
      break;

    case "cancelled":
      newContent = "⊘ 操作已取消";
      break;
  }

  return {
    ...message,
    content: newContent,
    progress: chunk.progress || message.progress,
  };
}

/**
 * 从 chunk 类型获取阶段
 */
function getPhaseFromChunk(chunk: StreamChunk): StreamingAgentState["phase"] {
  switch (chunk.type) {
    case "thinking":
    case "intent":
    case "plan":
      return "thinking";
    case "step:start":
    case "step:done":
    case "step:error":
    case "step:recovery":
      return "executing";
    case "complete":
      return "complete";
    case "error":
      return "error";
    default:
      return "thinking";
  }
}

export default useStreamingAgent;
