/**
 * useMessages - 消息状态管理 Hook
 * @file src/taskpane/hooks/useMessages.ts
 * @description v2.9.8 管理聊天消息的增删改查
 */
import * as React from "react";
import type { ChatMessage, CopilotAction } from "../types/ui.types";

/** 生成唯一 ID */
function uid(): string {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}

export interface UseMessagesReturn {
  /** 消息列表 */
  messages: ChatMessage[];
  /** 是否显示欢迎界面 */
  showWelcome: boolean;
  /** 添加用户消息 */
  addUserMessage: (text: string) => string;
  /** 添加助手消息 */
  addAssistantMessage: (text: string, actions?: CopilotAction[]) => string;
  /** 更新消息文本 */
  updateMessage: (id: string, text: string) => void;
  /** 更新消息（完整替换） */
  updateMessageFull: (id: string, updates: Partial<ChatMessage>) => void;
  /** 追加文本到消息 */
  appendToMessage: (id: string, text: string) => void;
  /** 清除所有消息 */
  clearMessages: () => void;
  /** 获取最近 N 条消息 */
  getRecentMessages: (count: number) => ChatMessage[];
}

/**
 * 消息状态管理 Hook
 */
export function useMessages(): UseMessagesReturn {
  const [messages, setMessages] = React.useState<ChatMessage[]>([
    {
      id: "welcome",
      role: "assistant",
      text: "你好！我是你的 Excel 智能助手。选中一些数据，告诉我你想做什么？",
      timestamp: new Date(),
    },
  ]);

  const showWelcome = messages.length <= 1;

  const addUserMessage = React.useCallback((text: string): string => {
    const id = uid();
    const message: ChatMessage = {
      id,
      role: "user",
      text,
      timestamp: new Date(),
    };
    setMessages((prev) => [...prev, message]);
    return id;
  }, []);

  const addAssistantMessage = React.useCallback(
    (text: string, actions?: CopilotAction[]): string => {
      const id = uid();
      const message: ChatMessage = {
        id,
        role: "assistant",
        text,
        timestamp: new Date(),
        actions,
      };
      setMessages((prev) => [...prev, message]);
      return id;
    },
    []
  );

  const updateMessage = React.useCallback((id: string, text: string): void => {
    setMessages((prev) => prev.map((msg) => (msg.id === id ? { ...msg, text } : msg)));
  }, []);

  const updateMessageFull = React.useCallback((id: string, updates: Partial<ChatMessage>): void => {
    setMessages((prev) => prev.map((msg) => (msg.id === id ? { ...msg, ...updates } : msg)));
  }, []);

  const appendToMessage = React.useCallback((id: string, text: string): void => {
    setMessages((prev) =>
      prev.map((msg) => (msg.id === id ? { ...msg, text: msg.text + text } : msg))
    );
  }, []);

  const clearMessages = React.useCallback((): void => {
    setMessages([
      {
        id: "welcome",
        role: "assistant",
        text: "你好！我是你的 Excel 智能助手。选中一些数据，告诉我你想做什么？",
        timestamp: new Date(),
      },
    ]);
  }, []);

  const getRecentMessages = React.useCallback(
    (count: number): ChatMessage[] => {
      return messages.slice(-count);
    },
    [messages]
  );

  return {
    messages,
    showWelcome,
    addUserMessage,
    addAssistantMessage,
    updateMessage,
    updateMessageFull,
    appendToMessage,
    clearMessages,
    getRecentMessages,
  };
}

export default useMessages;
