/**
 * useProactiveAgent - 主动洞察型 Agent Hook
 *
 * 像一个有经验的分析师：
 * 1. 主动观察工作表
 * 2. 形成洞察和判断
 * 3. 提供建议
 * 4. 等待用户确认
 *
 * @module hooks/useProactiveAgent
 */

import * as React from "react";
import {
  ProactiveAgent,
  createProactiveAgent,
  ProactiveAgentState,
  AgentMessage,
  InsightReport,
  Suggestion,
  AgentEventType,
} from "../../agent/proactive";
import { ToolRegistry, createExcelTools } from "../../agent";

// ========== 类型定义 ==========

export interface UseProactiveAgentOptions {
  /** 启动时自动分析 */
  autoAnalyzeOnStart?: boolean;
  /** 工作表切换时自动分析 */
  autoAnalyzeOnSheetChange?: boolean;
  /** 详细日志 */
  verboseLogging?: boolean;
  /** 事件回调 */
  onEvent?: (event: AgentEventType, data: unknown) => void;
}

export interface UseProactiveAgentReturn {
  // 状态
  state: ProactiveAgentState;
  isAnalyzing: boolean;
  isExecuting: boolean;

  // 消息和洞察
  messages: AgentMessage[];
  insights: InsightReport | null;
  suggestions: Suggestion[];

  // 操作
  startAnalysis: (sheetName?: string) => Promise<void>;
  sendMessage: (message: string) => Promise<string>;
  executeAll: () => Promise<string>;
  executeSuggestion: (suggestionId: string) => Promise<string>;
  reset: () => void;

  // 最新消息（用于显示）
  latestMessage: AgentMessage | null;

  // 快速操作
  quickActions: Array<{ label: string; action: string }>;
}

// ========== Hook 实现 ==========

export function useProactiveAgent(
  options: UseProactiveAgentOptions = {}
): UseProactiveAgentReturn {
  const {
    autoAnalyzeOnStart = true,
    autoAnalyzeOnSheetChange = true,
    verboseLogging = false,
    onEvent,
  } = options;

  // Agent 实例
  const agentRef = React.useRef<ProactiveAgent | null>(null);
  const toolRegistryRef = React.useRef<ToolRegistry | null>(null);

  // 状态
  const [state, setState] = React.useState<ProactiveAgentState>("idle");
  const [messages, setMessages] = React.useState<AgentMessage[]>([]);
  const [insights, setInsights] = React.useState<InsightReport | null>(null);

  // 初始化
  React.useEffect(() => {
    // 创建工具注册表
    if (!toolRegistryRef.current) {
      toolRegistryRef.current = new ToolRegistry();
      const excelTools = createExcelTools();
      for (const tool of excelTools) {
        toolRegistryRef.current.register(tool);
      }
    }

    // 创建 Agent
    if (!agentRef.current) {
      agentRef.current = createProactiveAgent(toolRegistryRef.current, {
        config: {
          autoAnalyzeOnStart,
          autoAnalyzeOnSheetChange,
          verboseLogging,
        },
      });

      // 订阅事件
      agentRef.current.on((event, data) => {
        switch (event) {
          case "state:change":
            const { to } = data as { from: ProactiveAgentState; to: ProactiveAgentState };
            setState(to);
            break;

          case "message:new":
            const message = data as AgentMessage;
            setMessages((prev) => [...prev, message]);
            break;

          case "insight:ready":
            const insightData = data as { insights: InsightReport };
            setInsights(insightData.insights);
            break;

          case "analysis:complete":
            const analysisData = data as { insights: InsightReport };
            setInsights(analysisData.insights);
            break;
        }

        // 转发事件给外部
        onEvent?.(event, data);
      });

      // 自动开始分析
      if (autoAnalyzeOnStart) {
        agentRef.current.start();
      }
    }

    return () => {
      // 清理
    };
  }, [autoAnalyzeOnStart, autoAnalyzeOnSheetChange, verboseLogging, onEvent]);

  // 计算派生状态
  const isAnalyzing = state === "observing" || state === "analyzing";
  const isExecuting = state === "executing";
  const latestMessage = messages.length > 0 ? messages[messages.length - 1] : null;
  const suggestions = insights?.suggestions || [];
  const quickActions = insights?.quickActions || [];

  // 操作函数
  const startAnalysis = React.useCallback(async (sheetName?: string) => {
    if (!agentRef.current) return;
    await agentRef.current.observeAndAnalyze(sheetName);
  }, []);

  const sendMessage = React.useCallback(async (message: string): Promise<string> => {
    if (!agentRef.current) return "Agent 未初始化";
    return await agentRef.current.handleUserInput(message);
  }, []);

  const executeAll = React.useCallback(async (): Promise<string> => {
    if (!agentRef.current) return "Agent 未初始化";
    return await agentRef.current.executeAllSuggestions();
  }, []);

  const executeSuggestion = React.useCallback(async (suggestionId: string): Promise<string> => {
    if (!agentRef.current) return "Agent 未初始化";
    return await agentRef.current.executeSuggestion(suggestionId);
  }, []);

  const reset = React.useCallback(() => {
    if (!agentRef.current) return;
    agentRef.current.reset();
    setMessages([]);
    setInsights(null);
    setState("idle");
  }, []);

  return {
    state,
    isAnalyzing,
    isExecuting,
    messages,
    insights,
    suggestions,
    startAnalysis,
    sendMessage,
    executeAll,
    executeSuggestion,
    reset,
    latestMessage,
    quickActions,
  };
}
