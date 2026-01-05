/**
 * Intent 类型定义 - 意图规格层
 *
 * v4.0 架构核心: LLM 只输出意图规格，不知道任何工具名
 *
 * @module agent/types/intent
 */

// ========== 意图类型枚举 ==========

/**
 * 高层意图类型 - LLM 只需要理解这些业务概念
 * 注意: 不包含任何工具名如 excel_xxx
 */
export type IntentType =
  // 表格操作
  | "create_table" // 创建表格结构
  | "write_data" // 写入数据
  | "update_data" // 更新现有数据
  | "delete_data" // 删除数据

  // 格式化
  | "format_range" // 格式化区域
  | "style_table" // 美化表格
  | "conditional_format" // 条件格式

  // 公式计算
  | "create_formula" // 创建公式
  | "batch_formula" // 批量公式
  | "calculate_summary" // 汇总计算

  // 数据分析
  | "analyze_data" // 分析数据
  | "find_pattern" // 发现规律
  | "statistics" // 统计分析

  // 图表
  | "create_chart" // 创建图表
  | "modify_chart" // 修改图表

  // 工作表
  | "create_sheet" // 创建工作表
  | "switch_sheet" // 切换工作表
  | "organize_sheets" // 整理工作表

  // 数据处理
  | "sort_data" // 排序
  | "filter_data" // 筛选
  | "remove_duplicates" // 去重
  | "clean_data" // 清理数据

  // 查询
  | "query_data" // 查询数据
  | "lookup_value" // 查找值

  // 特殊
  | "clarify" // 需要澄清
  | "respond_only"; // 只需回复，不需操作

// ========== 意图规格 ==========

/**
 * 意图规格 - IntentParser 的输出
 *
 * 注意: 不包含任何工具名，只有业务规格
 */
export interface IntentSpec {
  /** 意图类型 */
  intent: IntentType;

  /** 置信度 0-1 */
  confidence: number;

  /** 是否需要澄清 */
  needsClarification: boolean;

  /** 澄清问题（如果 needsClarification 为 true） */
  clarificationQuestion?: string;

  /** 澄清选项 */
  clarificationOptions?: string[];

  /** 业务规格 - 不同意图有不同的规格 */
  spec: IntentSpecData;

  /** LLM 的思考过程 */
  reasoning?: string;
}

/**
 * 业务规格数据 - 联合类型
 */
export type IntentSpecData =
  | CreateTableSpec
  | WriteDataSpec
  | FormatSpec
  | FormulaSpec
  | ChartSpec
  | SheetSpec
  | DataOperationSpec
  | QuerySpec
  | ClarifySpec
  | RespondSpec;

// ========== 具体规格类型 ==========

/**
 * 创建表格规格
 */
export interface CreateTableSpec {
  type: "create_table";
  /** 表格类型 */
  tableType?: "sales" | "inventory" | "employee" | "custom";
  /** 列定义 */
  columns: ColumnDefinition[];
  /** 初始数据行数 */
  initialRows?: number;
  /** 起始位置 */
  startCell?: string;
  /** 目标工作表 */
  targetSheet?: string;
  /** 选项 */
  options?: {
    hasHeader?: boolean;
    hasTotalRow?: boolean;
    autoFormat?: boolean;
  };
}

/**
 * 列定义
 */
export interface ColumnDefinition {
  name: string;
  type: "text" | "number" | "date" | "currency" | "percentage" | "formula";
  formula?: string; // 如果是公式列
  width?: number;
}

/**
 * 写入数据规格
 */
export interface WriteDataSpec {
  type: "write_data";
  /** 目标位置 */
  target: string | { sheet?: string; range?: string };
  /** 要写入的数据 */
  data: unknown[][] | string;
  /** 是否覆盖 */
  overwrite?: boolean;
}

/**
 * 格式化规格
 */
export interface FormatSpec {
  type: "format_range";
  /** 目标范围 */
  range?: string;
  /** 格式选项 */
  format: {
    bold?: boolean;
    italic?: boolean;
    fontSize?: number;
    fontColor?: string;
    backgroundColor?: string;
    alignment?: "left" | "center" | "right";
    borders?: boolean;
    numberFormat?: string;
  };
}

/**
 * 公式规格
 */
export interface FormulaSpec {
  type: "formula";
  /** 目标单元格 */
  targetCell: string;
  /** 公式类型 */
  formulaType: "sum" | "average" | "count" | "if" | "vlookup" | "custom";
  /** 数据源范围 */
  sourceRange?: string;
  /** 自定义公式（如果 formulaType 是 custom） */
  customFormula?: string;
  /** 条件（用于 IF 等） */
  condition?: string;
}

/**
 * 图表规格
 */
export interface ChartSpec {
  type: "chart";
  /** 图表类型 */
  chartType: "line" | "bar" | "pie" | "column" | "area" | "scatter";
  /** 数据范围 */
  dataRange: string;
  /** 标题 */
  title?: string;
  /** 选项 */
  options?: {
    showLegend?: boolean;
    showDataLabels?: boolean;
  };
}

/**
 * 工作表规格
 */
export interface SheetSpec {
  type: "sheet";
  /** 操作类型 */
  operation: "create" | "switch" | "rename" | "delete";
  /** 工作表名称 */
  sheetName: string;
  /** 新名称（重命名时） */
  newName?: string;
}

/**
 * 数据操作规格
 */
export interface DataOperationSpec {
  type: "data_operation";
  /** 操作类型 */
  operation: "sort" | "filter" | "dedupe" | "clean";
  /** 目标范围 */
  range?: string;
  /** 排序列 */
  sortColumn?: string;
  /** 排序方向 */
  sortDirection?: "asc" | "desc";
  /** 筛选条件 */
  filterCondition?: string;
}

/**
 * 查询规格
 */
export interface QuerySpec {
  type: "query";
  /** 查询目标 */
  target: "selection" | "range" | "sheet";
  /** 范围（如果 target 是 range） */
  range?: string;
  /** 查询条件 */
  condition?: string;
}

/**
 * 澄清规格
 */
export interface ClarifySpec {
  type: "clarify";
  /** 需要澄清的问题 */
  question: string;
  /** 选项 */
  options?: string[];
  /** 原因 */
  reason: string;
}

/**
 * 回复规格
 */
export interface RespondSpec {
  type: "respond";
  /** 回复内容 */
  message: string;
}

// ========== 导出 ==========

export default IntentSpec;
