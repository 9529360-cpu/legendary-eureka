/**
 * Proactive Agent 模块入口
 *
 * 主动洞察型 Agent - 像人一样先观察、理解、再建议
 *
 * @module agent/proactive
 */

// 工作表分析器
export {
  WorksheetAnalyzer,
  createWorksheetAnalyzer,
  type WorksheetAnalysis,
  type ColumnAnalysis,
  type RowAnalysis,
  type DataRegion,
  type AnalysisIssue,
  type IssueType,
  type TableStructure,
  type CellInfo,
  type MergedCellInfo,
} from "./WorksheetAnalyzer";

// 洞察生成器
export {
  InsightGenerator,
  createInsightGenerator,
  type InsightReport,
  type Insight,
  type InsightType,
  type Suggestion,
  type SuggestedAction,
  type QuickAction,
} from "./InsightGenerator";

// 主动型 Agent
export {
  ProactiveAgent,
  createProactiveAgent,
  type ProactiveAgentState,
  type ProactiveAgentConfig,
  type AgentMessage,
  type ConversationContext,
  type UserPreferences,
  type AgentEventType,
  type AgentEventHandler,
} from "./ProactiveAgent";
