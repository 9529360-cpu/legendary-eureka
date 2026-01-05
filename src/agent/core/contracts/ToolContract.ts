/**
 * ToolContract - 模型无关的工具契约层
 *
 * 单一职责：定义统一的工具接口规范，供不同 LLM 适配器使用
 * 行数上限：300 行
 *
 * 设计原则：
 * 1. 所有工具必须声明：name, description, input_schema, output_schema, failure_modes
 * 2. 与具体 LLM 实现解耦
 * 3. 支持 OpenAI / Claude / Gemini / DeepSeek 等模型
 */

// ========== 工具契约定义 ==========

/**
 * 工具契约 - 模型无关的工具定义
 */
export interface ToolContract {
  /** 工具名称 */
  name: string;

  /** 工具描述 */
  description: string;

  /** 输入参数 Schema */
  inputSchema: JsonSchema;

  /** 输出结果 Schema */
  outputSchema: JsonSchema;

  /** 可能的失败模式 */
  failureModes: FailureMode[];

  /** 工具类别 */
  category: ToolCategory;

  /** 是否需要确认 */
  requiresConfirmation: boolean;

  /** 风险等级 */
  riskLevel: "read_only" | "low" | "medium" | "high";
}

/**
 * JSON Schema 类型
 */
export interface JsonSchema {
  type: "object" | "array" | "string" | "number" | "boolean";
  properties?: Record<string, JsonSchemaProperty>;
  required?: string[];
  items?: JsonSchemaProperty;
}

/**
 * JSON Schema 属性
 */
export interface JsonSchemaProperty {
  type: string;
  description?: string;
  enum?: unknown[];
  default?: unknown;
  items?: JsonSchemaProperty;
}

/**
 * 失败模式
 */
export interface FailureMode {
  code: string;
  description: string;
  recoveryHint: string;
}

/**
 * 工具类别
 */
export type ToolCategory =
  | "read" // 读取类
  | "write" // 写入类
  | "formula" // 公式类
  | "format" // 格式化类
  | "chart" // 图表类
  | "data" // 数据操作类
  | "sheet" // 工作表类
  | "analysis" // 分析类
  | "diagnostic" // 诊断类
  | "refactor"; // 重构建议类

// ========== 预定义工具契约 ==========

/**
 * 表格读取工具
 */
export const READ_SHEET_CONTRACT: ToolContract = {
  name: "read_sheet",
  description: "读取 Excel / Google Sheets 中指定范围的数据",
  inputSchema: {
    type: "object",
    properties: {
      file_id: { type: "string", description: "文件标识符" },
      sheet_name: { type: "string", description: "工作表名称" },
      range: { type: "string", description: "单元格范围，如 A1:D10" },
    },
    required: ["range"],
  },
  outputSchema: {
    type: "object",
    properties: {
      data: { type: "array", description: "二维数据数组" },
      data_type: { type: "string", enum: ["number", "text", "mixed"] },
    },
  },
  failureModes: [
    {
      code: "RANGE_NOT_FOUND",
      description: "指定范围不存在",
      recoveryHint: "检查范围地址是否正确",
    },
    { code: "SHEET_NOT_FOUND", description: "工作表不存在", recoveryHint: "检查工作表名称" },
    { code: "PERMISSION_DENIED", description: "无权限访问", recoveryHint: "请求访问权限" },
  ],
  category: "read",
  requiresConfirmation: false,
  riskLevel: "read_only",
};

/**
 * 表格写入工具
 */
export const WRITE_SHEET_CONTRACT: ToolContract = {
  name: "write_sheet",
  description: "向指定表格范围写入公式或数据",
  inputSchema: {
    type: "object",
    properties: {
      file_id: { type: "string", description: "文件标识符" },
      sheet_name: { type: "string", description: "工作表名称" },
      range: { type: "string", description: "目标范围" },
      value: { type: "string", description: "要写入的值或公式" },
    },
    required: ["range", "value"],
  },
  outputSchema: {
    type: "object",
    properties: {
      status: { type: "string", enum: ["success", "failure"] },
    },
  },
  failureModes: [
    { code: "PROTECTED_RANGE", description: "范围被保护", recoveryHint: "解除保护或选择其他范围" },
    { code: "INVALID_FORMULA", description: "公式语法错误", recoveryHint: "检查公式语法" },
    { code: "CIRCULAR_REFERENCE", description: "循环引用", recoveryHint: "调整公式避免自引用" },
  ],
  category: "write",
  requiresConfirmation: true,
  riskLevel: "medium",
};

/**
 * 公式生成工具
 */
export const GENERATE_FORMULA_CONTRACT: ToolContract = {
  name: "generate_formula",
  description: "根据语义生成 Excel / Google Sheets 公式",
  inputSchema: {
    type: "object",
    properties: {
      platform: { type: "string", enum: ["excel", "google_sheets"] },
      intent: { type: "string", description: "用户意图描述" },
      entities: { type: "object", description: "相关实体（列名、范围等）" },
      constraints: { type: "array", description: "约束条件" },
    },
    required: ["platform", "intent"],
  },
  outputSchema: {
    type: "object",
    properties: {
      formula: { type: "string", description: "生成的公式" },
      explanation: { type: "string", description: "公式说明" },
      risk_notes: { type: "array", description: "风险说明" },
    },
  },
  failureModes: [
    { code: "AMBIGUOUS_INTENT", description: "意图不明确", recoveryHint: "请提供更多上下文" },
    { code: "UNSUPPORTED_FUNCTION", description: "不支持的函数", recoveryHint: "使用替代方案" },
  ],
  category: "formula",
  requiresConfirmation: false,
  riskLevel: "low",
};

/**
 * 诊断工具
 */
export const DIAGNOSE_ISSUE_CONTRACT: ToolContract = {
  name: "diagnose_spreadsheet_issue",
  description: "诊断表格中计算异常的根因",
  inputSchema: {
    type: "object",
    properties: {
      symptom: { type: "string", description: '症状描述（如"结果是0"、"循环引用"）' },
      context: { type: "object", description: "上下文信息（公式、范围、数据类型）" },
    },
    required: ["symptom"],
  },
  outputSchema: {
    type: "object",
    properties: {
      possible_causes: { type: "array", description: "Top3 可能原因" },
      validation_steps: { type: "array", description: "验证步骤" },
      recommended_fix: { type: "string", description: "推荐修复方案" },
    },
  },
  failureModes: [
    { code: "INSUFFICIENT_CONTEXT", description: "上下文不足", recoveryHint: "提供更多表格信息" },
  ],
  category: "diagnostic",
  requiresConfirmation: false,
  riskLevel: "read_only",
};

/**
 * 结构重构建议工具
 */
export const SUGGEST_REFACTOR_CONTRACT: ToolContract = {
  name: "suggest_schema_refactor",
  description: "评估当前表结构是否适合长期使用，并给出重构建议",
  inputSchema: {
    type: "object",
    properties: {
      current_schema: { type: "object", description: "当前表结构描述" },
      usage_pattern: { type: "string", description: '使用模式（如"每日汇总"、"跨部门协作"）' },
    },
    required: ["usage_pattern"],
  },
  outputSchema: {
    type: "object",
    properties: {
      issues: { type: "array", description: "发现的结构问题" },
      recommended_structure: { type: "object", description: "推荐的表结构" },
      migration_steps: { type: "array", description: "迁移步骤" },
    },
  },
  failureModes: [
    { code: "COMPLEX_SCHEMA", description: "表结构过于复杂", recoveryHint: "分阶段重构" },
  ],
  category: "refactor",
  requiresConfirmation: true,
  riskLevel: "high",
};

// ========== 工具契约注册表 ==========

/**
 * 所有工具契约
 */
export const TOOL_CONTRACTS: Record<string, ToolContract> = {
  read_sheet: READ_SHEET_CONTRACT,
  write_sheet: WRITE_SHEET_CONTRACT,
  generate_formula: GENERATE_FORMULA_CONTRACT,
  diagnose_spreadsheet_issue: DIAGNOSE_ISSUE_CONTRACT,
  suggest_schema_refactor: SUGGEST_REFACTOR_CONTRACT,
};

/**
 * 获取工具契约
 */
export function getToolContract(name: string): ToolContract | undefined {
  return TOOL_CONTRACTS[name];
}

/**
 * 获取所有工具契约
 */
export function getAllToolContracts(): ToolContract[] {
  return Object.values(TOOL_CONTRACTS);
}

/**
 * 按类别获取工具契约
 */
export function getToolContractsByCategory(category: ToolCategory): ToolContract[] {
  return Object.values(TOOL_CONTRACTS).filter((c) => c.category === category);
}
