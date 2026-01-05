/**
 * 组件索引文件
 * 统一导出所有可复用组件
 * @description v2.9.8 新增 ChatInputArea, ApiConfigDialog, PreviewConfirmDialog, MessageBubble
 *              HeaderBar, MessageList, WelcomeView, InsightPanel
 */

export { default as ChatMessage } from "./ChatMessage";
export type { ChatMessageData } from "./ChatMessage";

export { default as ChatInput } from "./ChatInput";

export { default as AnalysisPanel } from "./AnalysisPanel";
export type { DataInsight } from "./AnalysisPanel";

export { default as HistorySidebar } from "./HistorySidebar";
export type { HistoryItemData } from "./HistorySidebar";

export { default as SettingsDialog } from "./SettingsDialog";

// v2.9.8 新增组件
export { MessageBubble } from "./MessageBubble";
export type { MessageBubbleProps } from "./MessageBubble";

export { ChatInputArea } from "./ChatInputArea";
export type { ChatInputAreaProps } from "./ChatInputArea";

export { ApiConfigDialog } from "./ApiConfigDialog";
export type { ApiConfigDialogProps } from "./ApiConfigDialog";

export { PreviewConfirmDialog } from "./PreviewConfirmDialog";
export type { PreviewConfirmDialogProps } from "./PreviewConfirmDialog";

// v3.0: Agent 层审批确认弹窗
export { ApprovalDialog } from "./ApprovalDialog";
export type { ApprovalDialogProps } from "./ApprovalDialog";

export { HeaderBar } from "./HeaderBar";
export type { HeaderBarProps, WorkbookSummary } from "./HeaderBar";

export { MessageList } from "./MessageList";
export type { MessageListProps } from "./MessageList";

export { WelcomeView } from "./WelcomeView";
export type { WelcomeViewProps, QuickAction } from "./WelcomeView";

export { InsightPanel } from "./InsightPanel";
export type { InsightPanelProps, ProactiveSuggestion } from "./InsightPanel";
