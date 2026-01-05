/**
 * ToolProtocol - 统一工具描述协议
 * v1.0.0
 *
 * 功能：
 * 1. 定义统一的工具Schema标准
 * 2. 工具参数/返回值/异常的规范化描述
 * 3. 支持工具版本管理
 * 4. 工具依赖关系声明
 * 5. 工具能力标签
 *
 * 解决的问题：
 * - 工具参数、返回值、异常处理无统一规范
 * - 调用方与实现方理解不一致
 * - 工具扩展性不足
 */

// ========== 基础类型 ==========

/**
 * 工具能力标签
 */
export enum ToolCapability {
  /** 读取数据 */
  READ = "read",
  /** 写入数据 */
  WRITE = "write",
  /** 格式化 */
  FORMAT = "format",
  /** 创建对象 */
  CREATE = "create",
  /** 删除对象 */
  DELETE = "delete",
  /** 分析数据 */
  ANALYZE = "analyze",
  /** 可视化 */
  VISUALIZE = "visualize",
  /** 查询 */
  QUERY = "query",
  /** 计算 */
  CALCULATE = "calculate",
  /** 批量操作 */
  BATCH = "batch",
  /** 可撤销 */
  UNDOABLE = "undoable",
  /** 需要确认 */
  REQUIRES_CONFIRMATION = "requires_confirmation",
}

/**
 * 工具风险级别
 */
export enum ToolRiskLevel {
  /** 无风险 - 只读操作 */
  NONE = "none",
  /** 低风险 - 可撤销的修改 */
  LOW = "low",
  /** 中等风险 - 影响范围有限的修改 */
  MEDIUM = "medium",
  /** 高风险 - 大范围修改或不可撤销 */
  HIGH = "high",
  /** 危险 - 可能导致数据丢失 */
  CRITICAL = "critical",
}

/**
 * 工具执行环境
 */
export enum ToolEnvironment {
  /** Excel */
  EXCEL = "excel",
  /** Word */
  WORD = "word",
  /** PowerPoint */
  POWERPOINT = "powerpoint",
  /** Outlook */
  OUTLOOK = "outlook",
  /** 通用 */
  UNIVERSAL = "universal",
  /** 服务端 */
  SERVER = "server",
}

/**
 * 参数数据类型
 */
export enum ParamDataType {
  STRING = "string",
  NUMBER = "number",
  BOOLEAN = "boolean",
  ARRAY = "array",
  OBJECT = "object",
  DATE = "date",
  RANGE_ADDRESS = "range_address",
  CELL_ADDRESS = "cell_address",
  FORMULA = "formula",
  COLOR = "color",
  ENUM = "enum",
  ANY = "any",
}

// ========== 参数协议 ==========

/**
 * 参数验证规则
 */
export interface ParamValidation {
  /** 是否必填 */
  required?: boolean;
  /** 最小值（数字/长度/数量） */
  min?: number;
  /** 最大值（数字/长度/数量） */
  max?: number;
  /** 正则表达式 */
  pattern?: string;
  /** 枚举值列表 */
  enum?: string[];
  /** 自定义验证函数名 */
  customValidator?: string;
  /** 错误消息 */
  errorMessage?: string;
}

/**
 * 参数协议定义
 */
export interface ParamProtocol {
  /** 参数名 */
  name: string;
  /** 参数类型 */
  type: ParamDataType;
  /** 描述 */
  description: string;
  /** 是否必填 */
  required: boolean;
  /** 默认值 */
  defaultValue?: unknown;
  /** 验证规则 */
  validation?: ParamValidation;
  /** 子属性（对象类型） */
  properties?: ParamProtocol[];
  /** 数组元素类型 */
  itemType?: ParamDataType;
  /** 示例值 */
  example?: unknown;
  /** 别名（支持多种参数名） */
  aliases?: string[];
}

// ========== 返回值协议 ==========

/**
 * 返回值协议定义
 */
export interface ReturnProtocol {
  /** 返回类型 */
  type: ParamDataType;
  /** 描述 */
  description: string;
  /** 子属性（对象类型） */
  properties?: ParamProtocol[];
  /** 数组元素类型 */
  itemType?: ParamDataType;
  /** 是否可能为null */
  nullable?: boolean;
  /** 示例值 */
  example?: unknown;
}

// ========== 异常协议 ==========

/**
 * 异常类型
 */
export enum ExceptionType {
  /** 参数验证失败 */
  VALIDATION_ERROR = "validation_error",
  /** 资源未找到 */
  NOT_FOUND = "not_found",
  /** 权限不足 */
  PERMISSION_DENIED = "permission_denied",
  /** 超时 */
  TIMEOUT = "timeout",
  /** 资源已存在 */
  ALREADY_EXISTS = "already_exists",
  /** 格式错误 */
  FORMAT_ERROR = "format_error",
  /** 依赖缺失 */
  DEPENDENCY_MISSING = "dependency_missing",
  /** 内部错误 */
  INTERNAL_ERROR = "internal_error",
  /** 用户取消 */
  USER_CANCELLED = "user_cancelled",
}

/**
 * 异常协议定义
 */
export interface ExceptionProtocol {
  /** 异常类型 */
  type: ExceptionType;
  /** 异常码 */
  code: string;
  /** 描述 */
  description: string;
  /** 是否可恢复 */
  recoverable: boolean;
  /** 建议的处理方式 */
  suggestedAction?: string;
  /** 用户友好消息 */
  userMessage?: string;
}

// ========== 工具协议 ==========

/**
 * 工具版本信息
 */
export interface ToolVersion {
  /** 主版本 */
  major: number;
  /** 次版本 */
  minor: number;
  /** 补丁版本 */
  patch: number;
  /** 发布日期 */
  releaseDate?: string;
  /** 变更说明 */
  changelog?: string;
}

/**
 * 工具依赖
 */
export interface ToolDependency {
  /** 依赖的工具名 */
  toolName: string;
  /** 版本要求（semver） */
  version?: string;
  /** 是否可选 */
  optional?: boolean;
  /** 依赖说明 */
  reason?: string;
}

/**
 * 工具示例
 */
export interface ToolExample {
  /** 示例名称 */
  name: string;
  /** 描述 */
  description: string;
  /** 输入参数 */
  input: Record<string, unknown>;
  /** 预期输出 */
  expectedOutput?: unknown;
  /** 前置条件 */
  preconditions?: string[];
}

/**
 * 完整工具协议定义
 */
export interface ToolProtocol {
  // ========== 基础信息 ==========
  /** 工具唯一标识 */
  id: string;
  /** 工具名称（调用时使用） */
  name: string;
  /** 显示名称 */
  displayName: string;
  /** 描述 */
  description: string;
  /** 详细说明 */
  longDescription?: string;
  /** 版本 */
  version: ToolVersion;
  /** 作者 */
  author?: string;
  /** 标签 */
  tags?: string[];

  // ========== 分类与能力 ==========
  /** 工具分类 */
  category: string;
  /** 子分类 */
  subCategory?: string;
  /** 执行环境 */
  environment: ToolEnvironment;
  /** 能力标签 */
  capabilities: ToolCapability[];
  /** 风险级别 */
  riskLevel: ToolRiskLevel;

  // ========== 参数与返回 ==========
  /** 输入参数 */
  parameters: ParamProtocol[];
  /** 返回值 */
  returns: ReturnProtocol;
  /** 可能的异常 */
  exceptions: ExceptionProtocol[];

  // ========== 依赖与关联 ==========
  /** 依赖的其他工具 */
  dependencies?: ToolDependency[];
  /** 相关工具 */
  relatedTools?: string[];
  /** 替代工具（当此工具不可用时） */
  alternatives?: string[];

  // ========== 使用示例 ==========
  /** 使用示例 */
  examples: ToolExample[];

  // ========== 元数据 ==========
  /** 是否启用 */
  enabled: boolean;
  /** 是否已废弃 */
  deprecated?: boolean;
  /** 废弃说明 */
  deprecationNote?: string;
  /** 替代方案 */
  replacement?: string;
  /** 最低支持的Office版本 */
  minOfficeVersion?: string;
  /** 创建时间 */
  createdAt?: string;
  /** 更新时间 */
  updatedAt?: string;
}

// ========== 协议验证 ==========

/**
 * 协议验证结果
 */
export interface ProtocolValidationResult {
  valid: boolean;
  errors: string[];
  warnings: string[];
}

/**
 * 验证工具协议
 */
export function validateToolProtocol(protocol: Partial<ToolProtocol>): ProtocolValidationResult {
  const errors: string[] = [];
  const warnings: string[] = [];

  // 必需字段检查
  if (!protocol.id) errors.push("缺少必需字段: id");
  if (!protocol.name) errors.push("缺少必需字段: name");
  if (!protocol.description) errors.push("缺少必需字段: description");
  if (!protocol.version) errors.push("缺少必需字段: version");
  if (!protocol.category) errors.push("缺少必需字段: category");
  if (!protocol.environment) errors.push("缺少必需字段: environment");
  if (!protocol.capabilities || protocol.capabilities.length === 0) {
    warnings.push("建议添加能力标签 (capabilities)");
  }
  if (!protocol.riskLevel) {
    warnings.push("建议指定风险级别 (riskLevel)");
  }
  if (!protocol.parameters) {
    warnings.push("建议定义参数列表 (parameters)");
  }
  if (!protocol.returns) {
    warnings.push("建议定义返回值 (returns)");
  }
  if (!protocol.examples || protocol.examples.length === 0) {
    warnings.push("建议添加使用示例 (examples)");
  }

  // ID格式检查
  if (protocol.id && !/^[a-z][a-z0-9_]*(\.[a-z][a-z0-9_]*)*$/.test(protocol.id)) {
    warnings.push("ID 建议使用小写字母和下划线，用点分隔命名空间");
  }

  // 参数验证
  if (protocol.parameters) {
    protocol.parameters.forEach((param, index) => {
      if (!param.name) {
        errors.push(`参数 ${index + 1}: 缺少 name`);
      }
      if (!param.type) {
        errors.push(`参数 ${param.name || index + 1}: 缺少 type`);
      }
      if (!param.description) {
        warnings.push(`参数 ${param.name || index + 1}: 建议添加 description`);
      }
    });
  }

  // 异常定义检查
  if (!protocol.exceptions || protocol.exceptions.length === 0) {
    warnings.push("建议定义可能的异常 (exceptions)");
  }

  return {
    valid: errors.length === 0,
    errors,
    warnings,
  };
}

// ========== 协议转换 ==========

/**
 * 将简化的工具定义转换为完整协议
 */
export function createToolProtocol(basic: {
  id: string;
  name: string;
  description: string;
  category: string;
  parameters?: Partial<ParamProtocol>[];
  capabilities?: ToolCapability[];
  riskLevel?: ToolRiskLevel;
  examples?: Partial<ToolExample>[];
}): ToolProtocol {
  return {
    id: basic.id,
    name: basic.name,
    displayName: basic.name.replace(/_/g, " ").replace(/\b\w/g, (c) => c.toUpperCase()),
    description: basic.description,
    version: { major: 1, minor: 0, patch: 0 },
    category: basic.category,
    environment: ToolEnvironment.EXCEL,
    capabilities: basic.capabilities || [ToolCapability.READ],
    riskLevel: basic.riskLevel || ToolRiskLevel.LOW,
    parameters: (basic.parameters || []).map((p) => ({
      name: p.name || "",
      type: p.type || ParamDataType.STRING,
      description: p.description || "",
      required: p.required ?? true,
      ...p,
    })) as ParamProtocol[],
    returns: {
      type: ParamDataType.OBJECT,
      description: "工具执行结果",
    },
    exceptions: [
      {
        type: ExceptionType.VALIDATION_ERROR,
        code: "PARAM_VALIDATION_FAILED",
        description: "参数验证失败",
        recoverable: true,
        userMessage: "请检查输入参数是否正确",
      },
      {
        type: ExceptionType.INTERNAL_ERROR,
        code: "EXECUTION_FAILED",
        description: "执行失败",
        recoverable: false,
        userMessage: "操作执行失败，请稍后重试",
      },
    ],
    examples: (basic.examples || []).map((e) => ({
      name: e.name || "示例",
      description: e.description || "",
      input: e.input || {},
      ...e,
    })) as ToolExample[],
    enabled: true,
  };
}

// ========== 协议注册表 ==========

/**
 * 工具协议注册表
 */
class ToolProtocolRegistryClass {
  private protocols: Map<string, ToolProtocol> = new Map();
  private changeListeners: ((event: {
    type: "add" | "remove" | "update";
    protocol: ToolProtocol;
  }) => void)[] = [];

  /**
   * 注册协议
   */
  register(protocol: ToolProtocol): ProtocolValidationResult {
    const validation = validateToolProtocol(protocol);

    if (validation.valid) {
      this.protocols.set(protocol.id, protocol);
      this.notifyListeners({ type: "add", protocol });
    }

    return validation;
  }

  /**
   * 批量注册
   */
  registerAll(protocols: ToolProtocol[]): { registered: number; failed: number; errors: string[] } {
    let registered = 0;
    let failed = 0;
    const errors: string[] = [];

    protocols.forEach((p) => {
      const result = this.register(p);
      if (result.valid) {
        registered++;
      } else {
        failed++;
        errors.push(`${p.id || "unknown"}: ${result.errors.join(", ")}`);
      }
    });

    return { registered, failed, errors };
  }

  /**
   * 获取协议
   */
  get(id: string): ToolProtocol | undefined {
    return this.protocols.get(id);
  }

  /**
   * 获取所有协议
   */
  getAll(): ToolProtocol[] {
    return Array.from(this.protocols.values());
  }

  /**
   * 按分类获取
   */
  getByCategory(category: string): ToolProtocol[] {
    return this.getAll().filter((p) => p.category === category);
  }

  /**
   * 按能力获取
   */
  getByCapability(capability: ToolCapability): ToolProtocol[] {
    return this.getAll().filter((p) => p.capabilities.includes(capability));
  }

  /**
   * 按环境获取
   */
  getByEnvironment(environment: ToolEnvironment): ToolProtocol[] {
    return this.getAll().filter((p) => p.environment === environment);
  }

  /**
   * 搜索协议
   */
  search(query: string): ToolProtocol[] {
    const lowerQuery = query.toLowerCase();
    return this.getAll().filter(
      (p) =>
        p.name.toLowerCase().includes(lowerQuery) ||
        p.description.toLowerCase().includes(lowerQuery) ||
        p.tags?.some((t) => t.toLowerCase().includes(lowerQuery))
    );
  }

  /**
   * 移除协议
   */
  remove(id: string): boolean {
    const protocol = this.protocols.get(id);
    if (protocol) {
      this.protocols.delete(id);
      this.notifyListeners({ type: "remove", protocol });
      return true;
    }
    return false;
  }

  /**
   * 更新协议
   */
  update(id: string, updates: Partial<ToolProtocol>): boolean {
    const existing = this.protocols.get(id);
    if (existing) {
      const updated = { ...existing, ...updates, id };
      const validation = validateToolProtocol(updated);
      if (validation.valid) {
        this.protocols.set(id, updated as ToolProtocol);
        this.notifyListeners({ type: "update", protocol: updated as ToolProtocol });
        return true;
      }
    }
    return false;
  }

  /**
   * 添加变更监听器
   */
  addChangeListener(
    listener: (event: { type: "add" | "remove" | "update"; protocol: ToolProtocol }) => void
  ): () => void {
    this.changeListeners.push(listener);
    return () => {
      this.changeListeners = this.changeListeners.filter((l) => l !== listener);
    };
  }

  /**
   * 导出所有协议为JSON
   */
  exportToJSON(): string {
    return JSON.stringify(this.getAll(), null, 2);
  }

  /**
   * 从JSON导入
   */
  importFromJSON(json: string): { imported: number; failed: number } {
    try {
      const protocols = JSON.parse(json) as ToolProtocol[];
      const result = this.registerAll(protocols);
      return { imported: result.registered, failed: result.failed };
    } catch {
      return { imported: 0, failed: 1 };
    }
  }

  /**
   * 清空注册表
   */
  clear(): void {
    this.protocols.clear();
  }

  private notifyListeners(event: {
    type: "add" | "remove" | "update";
    protocol: ToolProtocol;
  }): void {
    this.changeListeners.forEach((l) => l(event));
  }
}

// 导出单例
export const ToolProtocolRegistry = new ToolProtocolRegistryClass();

export default ToolProtocolRegistry;
