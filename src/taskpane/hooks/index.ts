/**
 * Taskpane Hooks 导出
 * @description v2.9.8 新增 useMessages, useSelection, useWorkbookContext
 * @description v2.9.12 新增 useSelectionListener, useUndoStack
 * @description v4.0 新增 useAgentV4 (新架构)
 * @description v4.0.1 useAgent 默认使用 v4 架构
 */

// v4.0: 新架构 - useAgentV4 作为主要 Agent Hook
export { useAgentV4, useAgentV4 as useAgent } from "./useAgentV4";
export type {
  UseAgentV4Options as UseAgentOptions,
  UseAgentV4Return as UseAgentReturn,
  AgentV4State as AgentState,
  AgentV4Progress as AgentProgress,
  AgentV4Context as AgentContext,
  AgentV4Status as AgentStatus,
} from "./useAgentV4";

// 旧版 Agent Hook（已弃用，保留兼容）
export { useAgent as useLegacyAgent } from "./useAgent";

export * from "./useApiSettings";
export * from "./useLegacyChat";
export * from "./useMessages";
export * from "./useSelection";
export * from "./useWorkbookContext";
export * from "./useSelectionListener";
export * from "./useUndoStack";
