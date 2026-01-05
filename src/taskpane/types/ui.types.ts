/**
 * UI 类型定义
 *
 * 所有 UI 层使用的类型定义
 * 注意：不要在这里引入业务逻辑或 Agent 类型
 *
 * @module ui.types
 */

import type { ChatResponse } from "../../services/ApiService";

// ========== 基础类型 ==========

export type Role = "user" | "assistant";

export type CellValue = string | number | boolean | null;

export type ExcelCommand = NonNullable<ChatResponse["excelCommand"]>;

// ========== 操作类型 ==========

export type CopilotAction =
  | { type: "writeRange"; address: string; values: CellValue[][] }
  | { type: "setFormula"; address: string; formula: string }
  | { type: "writeCell"; address: string; value: CellValue }
  | { type: "executeCommand"; command: ExcelCommand; label: string };

// ========== 消息类型 ==========

export interface ChatMessage {
  id: string;
  role: Role;
  text: string;
  actions?: CopilotAction[];
  timestamp: Date;
}

// ========== Excel 相关类型 ==========

export interface SelectionResult {
  address: string;
  values: CellValue[][];
  numberFormat?: string[][];
  rowCount: number;
  columnCount: number;
}

export interface CopilotResponse {
  message: string;
  actions: CopilotAction[];
}

// ========== 数据分析类型 ==========

export interface DataInsight {
  id: string;
  type: "statistic" | "quality" | "trend" | "recommendation";
  title: string;
  description: string;
  value?: string | number;
  severity?: "info" | "warning" | "error" | "success";
  actionable?: boolean;
}

export interface DataSummary {
  rowCount: number;
  columnCount: number;
  dataTypes: string[];
  hasHeaders: boolean;
  numericColumns: number;
  textColumns: number;
  dateColumns: number;
  emptyCount: number;
  qualityScore: number;
}

// ========== 操作历史与撤销 ==========

export interface OperationHistoryItem {
  id: string;
  operation: string;
  timestamp: Date;
  success: boolean;
  details: string;
}

export interface UndoStackItem {
  id: string;
  operation: string;
  timestamp: Date;
  sheetName: string;
  rangeAddress: string;
  previousValues: CellValue[][];
  previousFormulas?: string[][];
}

export interface OperationVerification {
  success: boolean;
  operationType: string;
  targetAddress: string;
  expectedResult?: unknown;
  actualResult?: unknown;
  matchesExpectation: boolean;
  details: string;
  timestamp: Date;
}

// ========== 主动建议 ==========

export interface ProactiveSuggestion {
  id: string;
  icon: "chart" | "formula" | "format" | "clean" | "analyze";
  title: string;
  description: string;
  action: () => Promise<void>;
  confidence: number; // 0-1
}

// ========== 用户偏好 ==========

export interface UserPreferences {
  theme: "light" | "dark";
  autoAnalyze: boolean;
  requireConfirmation: boolean;
  streamingMode: boolean;
  defaultChartType: string;
  favoriteOperations: string[];
  lastUsedOperations: Array<{ operation: string; timestamp: number }>;
}

export const DEFAULT_PREFERENCES: UserPreferences = {
  theme: "dark",
  autoAnalyze: true,
  requireConfirmation: true,
  streamingMode: true,
  defaultChartType: "column",
  favoriteOperations: [],
  lastUsedOperations: [],
};

export const PREFERENCES_STORAGE_KEY = "excel-copilot-preferences";

// ========== Agent UI 展示类型 ==========

/**
 * Agent 步骤（UI 展示用，简化版）
 * 注意：这是展示用的，不要和 AgentCore 的 AgentStep 混淆
 */
export interface AgentStepUI {
  id: string;
  type: "think" | "act" | "observe" | "complete" | "error";
  thought?: string;
  action?: string;
  toolName?: string;
  toolInput?: Record<string, unknown>;
  observation?: string;
  status: "running" | "success" | "failed";
  timestamp: Date;
  duration?: number;
}

export interface AgentThought {
  id: string;
  type: "observation" | "reasoning" | "decision" | "reflection";
  content: string;
  timestamp: Date;
}

export interface AgentPlanUI {
  id: string;
  originalRequest: string;
  steps: AgentStepUI[];
  currentStepIndex: number;
  status: "planning" | "executing" | "completed" | "failed" | "retrying";
  retryCount: number;
  maxRetries: number;
  createdAt: Date;
  completedAt?: Date;
}

// ========== 工作簿上下文 ==========

export interface SheetInfo {
  name: string;
  index: number;
  isActive: boolean;
  usedRangeAddress: string;
  rowCount: number;
  columnCount: number;
  hasData: boolean;
  hasTables: boolean;
  hasCharts: boolean;
  hasPivotTables: boolean;
}

export interface NamedRangeInfo {
  name: string;
  address: string;
  scope: "workbook" | "worksheet";
  comment?: string;
}

export interface TableInfo {
  name: string;
  sheetName: string;
  address: string;
  rowCount: number;
  columnCount: number;
  hasHeaders: boolean;
  columns: string[];
  style?: string;
}

export interface ChartInfo {
  name: string;
  sheetName: string;
  chartType: string;
  dataRange?: string;
  title?: string;
}

export interface PivotTableInfo {
  name: string;
  sheetName: string;
  sourceRange?: string;
}

export interface FormulaDependency {
  cellAddress: string;
  formula: string;
  dependsOn: string[];
  usedBy: string[];
}

export interface DataRelationship {
  sourceTable: string;
  targetTable: string;
  sourceColumn: string;
  targetColumn: string;
  type: "one-to-one" | "one-to-many" | "many-to-many";
}

export interface WorkbookContext {
  lastScanned: Date;
  fileName: string;
  sheets: SheetInfo[];
  namedRanges: NamedRangeInfo[];
  tables: TableInfo[];
  charts: ChartInfo[];
  pivotTables: PivotTableInfo[];
  totalCellsWithData: number;
  totalFormulas: number;
  formulaDependencies: FormulaDependency[];
  dataRelationships: DataRelationship[];
  // 快速查询索引
  sheetByName: Record<string, SheetInfo>;
  tableByName: Record<string, TableInfo>;
  // 整体质量摘要
  overallQualityScore: number;
  issues: Array<{
    type: "warning" | "error" | "suggestion";
    message: string;
    location?: string;
  }>;
}
