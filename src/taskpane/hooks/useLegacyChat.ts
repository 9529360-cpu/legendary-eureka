/**
 * useLegacyChat Hook
 * v2.9.8: 封装旧的非 Agent 模式聊天调用
 *
 * 注意：这是过渡性 hook，未来应该完全迁移到 Agent 模式
 * 目前保留是为了支持 shouldUseAgentMode() 返回 false 的简单查询场景
 */

import { useCallback, useRef } from "react";
import ApiService, { ChatRequest, ChatResponse } from "../../services/ApiService";

export interface StreamingCallbacks {
  onStart?: () => void;
  onChunk?: (content: string, accumulated: number) => void;
  onComplete?: (response: ChatResponse) => void;
  onError?: (error: string) => void;
}

export interface UseLegacyChatReturn {
  /** 发送流式聊天请求 */
  sendStreaming: (request: ChatRequest, callbacks: StreamingCallbacks) => Promise<void>;
  /** 发送普通聊天请求 */
  sendSync: (request: ChatRequest) => Promise<ChatResponse>;
  /** 是否正在请求 */
  isRequesting: boolean;
}

export function useLegacyChat(): UseLegacyChatReturn {
  const isRequestingRef = useRef(false);

  const sendStreaming = useCallback(
    async (request: ChatRequest, callbacks: StreamingCallbacks): Promise<void> => {
      if (isRequestingRef.current) return;

      isRequestingRef.current = true;
      try {
        await ApiService.sendStreamingChatRequest(request, {
          onStart: callbacks.onStart,
          onChunk: callbacks.onChunk,
          onComplete: callbacks.onComplete,
          onError: callbacks.onError,
        });
      } finally {
        isRequestingRef.current = false;
      }
    },
    []
  );

  const sendSync = useCallback(async (request: ChatRequest): Promise<ChatResponse> => {
    if (isRequestingRef.current) {
      throw new Error("请求正在进行中");
    }

    isRequestingRef.current = true;
    try {
      return await ApiService.sendChatRequest(request);
    } finally {
      isRequestingRef.current = false;
    }
  }, []);

  return {
    sendStreaming,
    sendSync,
    isRequesting: isRequestingRef.current,
  };
}

export default useLegacyChat;
