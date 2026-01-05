/**
 * DataModeler - 数据建模引擎 v2.9.52
 *
 * 职责：
 * 1. 分析任务需求，识别需要的表结构
 * 2. 确定表之间的依赖关系（主表、从表）
 * 3. 规划计算链（源字段 vs 派生字段）
 * 4. 验证公式引用的有效性
 * 5. 编译逻辑公式为 Excel 可执行公式
 *
 * v2.9.52 重构要点：
 * - 分离 LogicalFormula（字段名） 和 ExcelFormula（A1/Table引用）
 * - 统一依赖来源：从公式解析生成，写回模型
 * - 增加稳定定位：headerIndex、tableRef
 * - 升级依赖提取：支持 A1:B2、A:A、$A$1、Table[Col]
 * - 修复 isFieldUsed：基于引用图而非正则
 * - 添加执行成本和风险评估
 */

// ========== 数据建模类型 ==========

/**
 * 字段类型
 */
export type FieldType = "source" | "derived" | "lookup" | "validation";

/**
 * v2.9.52: 引用类型
 */
export type RefType = "cell" | "range" | "column" | "tableColumn" | "namedRange";

/**
 * v2.9.52: 稳定引用（不依赖列字母）
 */
export interface StableReference {
  // 方式1: 表头定位
  headerName?: string;
  headerRowIndex?: number; // 表头所在行（1-based）

  // 方式2: Excel Table 结构化引用
  tableName?: string;
  columnName?: string;

  // 方式3: 范围地址（fallback）
  rangeAddress?: string;
}

/**
 * 字段定义
 */
export interface FieldDefinition {
  name: string;
  column: string; // A, B, C... (编译后填充)
  columnIndex?: number; // 0-based 列索引
  type: FieldType;
  dataType: "text" | "number" | "date" | "currency" | "percentage" | "formula";
  description?: string;

  // v2.9.52: 稳定引用
  stableRef?: StableReference;

  // v2.9.52: 分离逻辑公式和Excel公式
  logicalFormula?: string; // 用字段名表达，如 "=单价*数量"
  excelFormula?: string; // 编译后的Excel公式，如 "=C2*D2" 或 "=Table1[@单价]*Table1[@数量]"

  // 兼容旧接口
  formula?: string;

  // v2.9.52: 统一从公式解析的依赖
  dependencies?: FieldDependency[];

  // 数据验证规则
  validation?: ValidationRule;
}

/**
 * 字段依赖 v2.9.52: 增强结构
 */
export interface FieldDependency {
  sheet: string;
  field: string;
  column?: string;

  // v2.9.52: 详细引用信息
  refType?: RefType;
  ref?: string; // 具体引用：A1 | A:A | Table1[销售额]
  lookupKey?: string;
}

/**
 * 验证规则
 */
export interface ValidationRule {
  type: "list" | "range" | "custom";
  values?: string[];
  min?: number;
  max?: number;
  formula?: string;
  errorMessage?: string;
}

/**
 * 表定义
 */
export interface TableDefinition {
  name: string;
  description: string;
  role: "master" | "transaction" | "summary" | "analysis";
  fields: FieldDefinition[];

  // v2.9.52: 稳定引用
  excelTableName?: string; // 如果创建为 Excel Table
  headerRowIndex?: number; // 表头行（1-based）
  dataStartRow?: number; // 数据起始行

  // 主键字段
  primaryKey?: string;

  // 外键关系
  foreignKeys?: ForeignKeyDefinition[];

  // 依赖的其他表（必须先创建）
  dependsOn?: string[];
}

/**
 * 外键定义
 */
export interface ForeignKeyDefinition {
  field: string;
  referencesTable: string;
  referencesField: string;
}

/**
 * 数据模型
 */
export interface DataModel {
  name: string;
  description: string;
  tables: TableDefinition[];

  // 执行顺序（拓扑排序后的结果）
  executionOrder: string[];

  // 计算链
  calculationChain: CalculationStep[];

  // v2.9.52: 字段索引（用于快速查找和引用图）
  fieldIndex?: Map<string, FieldIndexEntry>;

  // v2.9.52: 执行成本估算
  estimatedCost?: ExecutionCost;

  // v2.9.52: 风险评估
  risk?: RiskAssessment;
}

/**
 * v2.9.52: 字段索引条目
 */
export interface FieldIndexEntry {
  tableName: string;
  fieldName: string;
  fieldKey: string; // "表名.字段名"
  usedBy: string[]; // 被哪些字段引用
  uses: string[]; // 引用了哪些字段
}

/**
 * v2.9.52: 执行成本估算
 */
export interface ExecutionCost {
  readCells: number;
  writeCells: number;
  formulaWrites: number;
  apiCalls: number;
  estimatedTimeMs: number;
}

/**
 * v2.9.52: 风险评估
 */
export interface RiskAssessment {
  level: "low" | "medium" | "high";
  hasDestructiveWrites: boolean;
  hasFormulaWrites: boolean;
  hasLargeRangeOperations: boolean;
  requiresConfirmation: boolean;
  reasons: string[];
}

/**
 * 计算步骤
 */
export interface CalculationStep {
  order: number;
  sheet: string;
  field: string;
  column: string;

  // v2.9.52: 分离逻辑公式和Excel公式
  logicalFormula: string;
  excelFormula?: string;

  // 兼容旧接口
  formula: string;

  dependencies: string[];

  // v2.9.52: 详细依赖信息
  dependencyDetails?: ParsedDependency[];
}

/**
 * v2.9.52: 解析后的依赖
 */
export interface ParsedDependency {
  raw: string; // 原始引用文本
  sheet?: string; // 工作表名
  refType: RefType; // 引用类型
  ref: string; // 具体引用
  isAbsolute?: boolean; // 是否绝对引用
}

/**
 * 验证结果
 */
export interface ValidationResult {
  isValid: boolean;
  errors: ValidationError[];
  warnings: ValidationWarning[];
}

export interface ValidationError {
  type: "missing_dependency" | "circular_reference" | "invalid_formula" | "missing_source";
  sheet: string;
  field: string;
  message: string;
}

export interface ValidationWarning {
  type: "orphan_field" | "unused_lookup" | "potential_issue";
  sheet: string;
  field: string;
  message: string;
}

// ========== 数据建模引擎 ==========

export class DataModeler {
  /**
   * 分析任务需求，生成数据模型
   */
  analyzeRequirement(requirement: string): DataModelAnalysis {
    // 识别需要的表
    const tables = this.identifyTables(requirement);

    // 识别字段和关系
    const fields = this.identifyFields(requirement, tables);

    // 识别计算关系
    const calculations = this.identifyCalculations(requirement);

    return {
      tables,
      fields,
      calculations,
      suggestedModel: this.buildSuggestedModel(tables, fields, calculations),
    };
  }

  /**
   * 识别需要的表
   */
  private identifyTables(requirement: string): string[] {
    const tables: string[] = [];

    // 常见表名模式
    const tablePatterns = [
      { pattern: /订单|order/gi, name: "订单明细" },
      { pattern: /产品|product|目录|catalog/gi, name: "产品目录" },
      { pattern: /客户|customer|client/gi, name: "客户信息" },
      { pattern: /成本|cost|利润|profit/gi, name: "成本与利润" },
      { pattern: /汇总|summary|月度|daily|weekly|monthly/gi, name: "月度汇总" },
      { pattern: /分析|analysis|洞察|insight/gi, name: "经营洞察" },
      { pattern: /库存|inventory|stock/gi, name: "库存管理" },
    ];

    for (const { pattern, name } of tablePatterns) {
      if (pattern.test(requirement) && !tables.includes(name)) {
        tables.push(name);
      }
    }

    return tables;
  }

  /**
   * 识别字段和关系
   */
  private identifyFields(requirement: string, _tables: string[]): Map<string, string[]> {
    const fieldMap = new Map<string, string[]>();

    // 常见字段模式
    if (/订单/.test(requirement)) {
      fieldMap.set("订单明细", [
        "订单ID",
        "订单日期",
        "客户ID",
        "产品ID",
        "数量",
        "单价",
        "销售额",
        "销售渠道",
        "是否复购",
      ]);
    }

    if (/成本|利润/.test(requirement)) {
      fieldMap.set("成本与利润", [
        "订单ID",
        "单件成本",
        "物流成本",
        "渠道成本",
        "总成本",
        "销售额",
        "毛利润",
        "毛利率",
      ]);
    }

    if (/汇总|月度/.test(requirement)) {
      fieldMap.set("月度汇总", ["月份", "订单数", "总销售额", "总成本", "总毛利润", "平均毛利率"]);
    }

    return fieldMap;
  }

  /**
   * 识别计算关系
   */
  private identifyCalculations(requirement: string): CalculationPattern[] {
    const patterns: CalculationPattern[] = [];

    // 识别计算模式
    if (/销售额\s*[=＝]\s*.*[×*]|单价.*数量/.test(requirement)) {
      patterns.push({
        field: "销售额",
        formula: "=单价*数量",
        type: "derived",
      });
    }

    if (/总成本|成本合计/.test(requirement)) {
      patterns.push({
        field: "总成本",
        formula: "=单件成本+物流成本+渠道成本",
        type: "derived",
      });
    }

    if (/毛利润|利润/.test(requirement)) {
      patterns.push({
        field: "毛利润",
        formula: "=销售额-总成本",
        type: "derived",
      });
    }

    if (/毛利率|利润率/.test(requirement)) {
      patterns.push({
        field: "毛利率",
        formula: "=毛利润/销售额",
        type: "derived",
      });
    }

    if (/XLOOKUP|查找|引用/.test(requirement)) {
      patterns.push({
        field: "销售额",
        formula: '=XLOOKUP(订单ID,订单明细!A:A,订单明细!I:I,"")',
        type: "lookup",
      });
    }

    return patterns;
  }

  /**
   * 构建建议的数据模型
   */
  private buildSuggestedModel(
    tables: string[],
    fields: Map<string, string[]>,
    calculations: CalculationPattern[]
  ): DataModel {
    const tableDefinitions: TableDefinition[] = [];

    // 确定表的角色和依赖关系
    for (const tableName of tables) {
      const role = this.determineTableRole(tableName);
      const tableFields = fields.get(tableName) || [];

      const definition: TableDefinition = {
        name: tableName,
        description: this.getTableDescription(tableName),
        role,
        fields: tableFields.map((f) => this.createFieldDefinition(f, tableName, calculations)),
        dependsOn: this.getDependencies(tableName, tables),
      };

      tableDefinitions.push(definition);
    }

    // 拓扑排序确定执行顺序
    const executionOrder = this.topologicalSort(tableDefinitions);

    // 生成计算链
    const calculationChain = this.generateCalculationChain(tableDefinitions);

    return {
      name: "企业经营分析模型",
      description: "包含订单、成本、汇总等多表联动的数据分析模型",
      tables: tableDefinitions,
      executionOrder,
      calculationChain,
    };
  }

  /**
   * 确定表的角色
   */
  private determineTableRole(tableName: string): "master" | "transaction" | "summary" | "analysis" {
    if (/产品|客户|目录/.test(tableName)) return "master";
    if (/订单|交易|销售/.test(tableName)) return "transaction";
    if (/汇总|统计|月度/.test(tableName)) return "summary";
    return "analysis";
  }

  /**
   * 获取表描述
   */
  private getTableDescription(tableName: string): string {
    const descriptions: Record<string, string> = {
      订单明细: "记录每笔订单的详细信息，是业务数据的核心来源",
      成本与利润: "计算每笔订单的成本和利润，依赖订单明细的销售额",
      月度汇总: "按月汇总订单数据，使用 SUMIFS/COUNTIFS 聚合",
      经营洞察: "分析数据，生成经营洞察报告",
      产品目录: "产品主数据，包含产品信息和定价",
      客户信息: "客户主数据，包含客户基本信息",
    };
    return descriptions[tableName] || "数据表";
  }

  /**
   * 创建字段定义
   * v2.9.52: 增加 table 维度匹配，避免同名字段冲突
   */
  private createFieldDefinition(
    fieldName: string,
    tableName: string,
    calculations: CalculationPattern[]
  ): FieldDefinition {
    // v2.9.52: 优先匹配 table+field，避免同名字段冲突
    let calc = calculations.find(
      (c) =>
        c.field === fieldName && (c.table === tableName || c.appliesToTables?.includes(tableName))
    );

    // 如果没有精确匹配，退回到只匹配字段名
    if (!calc) {
      calc = calculations.find((c) => c.field === fieldName && !c.table);
    }

    return {
      name: fieldName,
      column: "", // 将在 ModelCompiler 中分配
      columnIndex: undefined,
      type: calc ? (calc.type as FieldType) : "source",
      dataType: this.inferDataType(fieldName),
      // v2.9.52: 分离逻辑公式和Excel公式
      logicalFormula: calc?.formula,
      formula: calc?.formula, // 兼容旧接口
      // 稳定引用将在表创建时填充
      stableRef: {
        headerName: fieldName,
      },
    };
  }

  /**
   * 推断数据类型
   */
  private inferDataType(fieldName: string): FieldDefinition["dataType"] {
    if (/日期|date|time/i.test(fieldName)) return "date";
    if (/率|ratio|percent/i.test(fieldName)) return "percentage";
    if (/金额|价|成本|利润|额|cost|price|profit|amount/i.test(fieldName)) return "currency";
    if (/数|量|count|number/i.test(fieldName)) return "number";
    return "text";
  }

  /**
   * 获取表的依赖
   */
  private getDependencies(tableName: string, allTables: string[]): string[] {
    const deps: string[] = [];

    // 成本与利润依赖订单明细
    if (tableName === "成本与利润" && allTables.includes("订单明细")) {
      deps.push("订单明细");
    }

    // 月度汇总依赖订单明细和成本与利润
    if (tableName === "月度汇总") {
      if (allTables.includes("订单明细")) deps.push("订单明细");
      if (allTables.includes("成本与利润")) deps.push("成本与利润");
    }

    // 经营洞察依赖其他所有表
    if (tableName === "经营洞察") {
      deps.push(...allTables.filter((t) => t !== tableName));
    }

    return deps;
  }

  /**
   * 拓扑排序 - 确定表的创建顺序
   */
  private topologicalSort(tables: TableDefinition[]): string[] {
    const result: string[] = [];
    const visited = new Set<string>();
    const visiting = new Set<string>();

    const visit = (tableName: string) => {
      if (visited.has(tableName)) return;
      if (visiting.has(tableName)) {
        throw new Error(`循环依赖检测到: ${tableName}`);
      }

      visiting.add(tableName);

      const table = tables.find((t) => t.name === tableName);
      if (table?.dependsOn) {
        for (const dep of table.dependsOn) {
          visit(dep);
        }
      }

      visiting.delete(tableName);
      visited.add(tableName);
      result.push(tableName);
    };

    for (const table of tables) {
      visit(table.name);
    }

    return result;
  }

  /**
   * 生成计算链
   */
  private generateCalculationChain(tables: TableDefinition[]): CalculationStep[] {
    const chain: CalculationStep[] = [];
    let order = 0;

    for (const table of tables) {
      for (const field of table.fields) {
        if (field.type === "derived" || field.type === "lookup") {
          const logicalFormula = field.logicalFormula || field.formula || "";
          const dependencyDetails = this.extractDependenciesV2(logicalFormula);

          chain.push({
            order: order++,
            sheet: table.name,
            field: field.name,
            column: field.column,
            logicalFormula,
            excelFormula: field.excelFormula,
            formula: field.excelFormula || logicalFormula, // 兼容旧接口
            dependencies: dependencyDetails.map((d) => d.raw),
            dependencyDetails,
          });
        }
      }
    }

    return chain;
  }

  /**
   * v2.9.52: 从公式中提取依赖（升级版）
   * 支持: A1, A1:B2, A:A, $A$1, Table[Col], 'Sheet Name'!A1, 命名区域
   */
  private extractDependenciesV2(formula: string): ParsedDependency[] {
    const deps: ParsedDependency[] = [];

    if (!formula) return deps;

    // 1. 匹配 Table 结构化引用: Table1[列名], Table1[@列名], Table1[[#All],[列名]]
    const tableRefPattern = /(\w+)\[(@?\[?#?\w*\]?,?\s*\[?)?([\u4e00-\u9fa5\w]+)\]?\]/g;
    let match;
    while ((match = tableRefPattern.exec(formula)) !== null) {
      deps.push({
        raw: match[0],
        refType: "tableColumn",
        ref: `${match[1]}[${match[3]}]`,
      });
    }

    // 2. 匹配跨表引用: Sheet!A1, 'Sheet Name'!A1:B2, Sheet!A:A
    const crossSheetPattern = /'?([^'!]+)'?!(\$?[A-Z]+\$?\d*(?::\$?[A-Z]+\$?\d*)?)/gi;
    while ((match = crossSheetPattern.exec(formula)) !== null) {
      const ref = match[2];
      const isAbsolute = ref.includes("$");
      let refType: RefType = "cell";

      if (ref.includes(":")) {
        refType = /^[A-Z]+:[A-Z]+$/i.test(ref) ? "column" : "range";
      } else if (/^[A-Z]+$/i.test(ref)) {
        refType = "column";
      }

      deps.push({
        raw: match[0],
        sheet: match[1],
        refType,
        ref,
        isAbsolute,
      });
    }

    // 3. 匹配本表引用: A1, $A$1, A1:B2, A:A, 1:1
    // 排除已匹配的跨表引用和 Table 引用
    const localPattern =
      /(?<![A-Z!'"\]\w])(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?|\$?[A-Z]+:\$?[A-Z]+|\d+:\d+)(?![(\[])/gi;
    while ((match = localPattern.exec(formula)) !== null) {
      const ref = match[1];

      // 排除函数名后的括号内容
      const beforeMatch = formula.substring(0, match.index);
      if (/[A-Z]+$/i.test(beforeMatch)) continue;

      const isAbsolute = ref.includes("$");
      let refType: RefType = "cell";

      if (/^\d+:\d+$/.test(ref)) {
        continue; // 整行引用暂不处理
      } else if (/^[A-Z]+:[A-Z]+$/i.test(ref.replace(/\$/g, ""))) {
        refType = "column";
      } else if (ref.includes(":")) {
        refType = "range";
      }

      deps.push({
        raw: match[0],
        refType,
        ref,
        isAbsolute,
      });
    }

    // 4. 匹配命名区域（简单处理：非函数的单词）
    // 这个需要更复杂的逻辑，暂时跳过

    return deps;
  }

  /**
   * 从公式中提取依赖（旧版本，保持兼容）
   * @deprecated 使用 extractDependenciesV2
   */
  private extractDependencies(formula: string): string[] {
    return this.extractDependenciesV2(formula).map((d) => d.raw);
  }

  /**
   * v2.9.52: 构建字段索引（用于引用图）
   */
  buildFieldIndex(model: DataModel): Map<string, FieldIndexEntry> {
    const index = new Map<string, FieldIndexEntry>();

    // 第一遍：创建所有字段条目
    for (const table of model.tables) {
      for (const field of table.fields) {
        const fieldKey = `${table.name}.${field.name}`;
        index.set(fieldKey, {
          tableName: table.name,
          fieldName: field.name,
          fieldKey,
          usedBy: [],
          uses: [],
        });
      }
    }

    // 第二遍：填充 uses 和 usedBy
    for (const step of model.calculationChain) {
      const currentKey = `${step.sheet}.${step.field}`;
      const currentEntry = index.get(currentKey);

      if (currentEntry && step.dependencyDetails) {
        for (const dep of step.dependencyDetails) {
          // 尝试匹配依赖到具体字段
          const depKey = this.resolveDependencyToFieldKey(dep, step.sheet, model);
          if (depKey) {
            currentEntry.uses.push(depKey);
            const depEntry = index.get(depKey);
            if (depEntry) {
              depEntry.usedBy.push(currentKey);
            }
          }
        }
      }
    }

    return index;
  }

  /**
   * v2.9.52: 解析依赖到字段Key
   */
  private resolveDependencyToFieldKey(
    dep: ParsedDependency,
    currentSheet: string,
    model: DataModel
  ): string | null {
    const sheet = dep.sheet || currentSheet;

    // Table 结构化引用
    if (dep.refType === "tableColumn") {
      // 从 Table1[销售额] 提取表名和列名
      const match = dep.ref.match(/(\w+)\[([\u4e00-\u9fa5\w]+)\]/);
      if (match) {
        const tableName = match[1];
        const columnName = match[2];
        // 找到对应的表
        const table = model.tables.find(
          (t) => t.excelTableName === tableName || t.name === tableName
        );
        if (table) {
          return `${table.name}.${columnName}`;
        }
      }
    }

    // 列引用 - 需要根据列字母映射到字段名
    if (dep.refType === "cell" || dep.refType === "column" || dep.refType === "range") {
      const colMatch = dep.ref.match(/([A-Z]+)/i);
      if (colMatch) {
        const table = model.tables.find((t) => t.name === sheet);
        if (table) {
          const field = table.fields.find((f) => f.column === colMatch[1].toUpperCase());
          if (field) {
            return `${table.name}.${field.name}`;
          }
        }
      }
    }

    return null;
  }

  /**
   * v2.9.52: 从公式丰富字段依赖
   * 统一依赖来源：从公式解析生成，写回 field.dependencies
   */
  enrichDependenciesFromFormulas(model: DataModel): void {
    for (const table of model.tables) {
      for (const field of table.fields) {
        const formula = field.logicalFormula || field.formula;
        if (!formula) continue;

        const parsedDeps = this.extractDependenciesV2(formula);
        field.dependencies = parsedDeps.map((dep) => ({
          sheet: dep.sheet || table.name,
          field: this.extractFieldNameFromRef(dep, table.name, model),
          column: this.extractColumnFromRef(dep),
          refType: dep.refType,
          ref: dep.ref,
        }));
      }
    }

    // 重新生成计算链
    model.calculationChain = this.generateCalculationChain(model.tables);

    // 构建字段索引
    model.fieldIndex = this.buildFieldIndex(model);
  }

  /**
   * 从引用提取字段名
   */
  private extractFieldNameFromRef(
    dep: ParsedDependency,
    currentSheet: string,
    model: DataModel
  ): string {
    if (dep.refType === "tableColumn") {
      const match = dep.ref.match(/\[([\u4e00-\u9fa5\w]+)\]/);
      return match ? match[1] : dep.ref;
    }

    // 尝试从列字母映射
    const colMatch = dep.ref.match(/([A-Z]+)/i);
    if (colMatch) {
      const sheet = dep.sheet || currentSheet;
      const table = model.tables.find((t) => t.name === sheet);
      if (table) {
        const field = table.fields.find((f) => f.column === colMatch[1].toUpperCase());
        if (field) return field.name;
      }
    }

    return dep.ref;
  }

  /**
   * 从引用提取列字母
   */
  private extractColumnFromRef(dep: ParsedDependency): string {
    const match = dep.ref.match(/([A-Z]+)/i);
    return match ? match[1].toUpperCase() : "";
  }

  /**
   * 验证数据模型 v2.9.52: 基于引用图验证
   */
  validateModel(model: DataModel): ValidationResult {
    const errors: ValidationError[] = [];
    const warnings: ValidationWarning[] = [];

    // 确保有字段索引
    if (!model.fieldIndex) {
      model.fieldIndex = this.buildFieldIndex(model);
    }

    // 检查循环依赖
    try {
      this.topologicalSort(model.tables);
    } catch (e) {
      errors.push({
        type: "circular_reference",
        sheet: "",
        field: "",
        message: e instanceof Error ? e.message : "检测到循环依赖",
      });
    }

    // v2.9.52: 基于 calculationChain 检查依赖
    for (const step of model.calculationChain) {
      if (step.dependencyDetails) {
        for (const dep of step.dependencyDetails) {
          const depKey = this.resolveDependencyToFieldKey(dep, step.sheet, model);
          if (depKey && !model.fieldIndex.has(depKey)) {
            errors.push({
              type: "missing_dependency",
              sheet: step.sheet,
              field: step.field,
              message: `依赖的字段 "${depKey}" 不存在`,
            });
          }
        }
      }
    }

    // v2.9.52: 基于引用图检查孤岛字段
    for (const [_fieldKey, entry] of model.fieldIndex) {
      // 源字段没有被任何公式引用
      const table = model.tables.find((t) => t.name === entry.tableName);
      const field = table?.fields.find((f) => f.name === entry.fieldName);

      if (field?.type === "source" && entry.usedBy.length === 0) {
        warnings.push({
          type: "orphan_field",
          sheet: entry.tableName,
          field: entry.fieldName,
          message: `字段 "${entry.fieldName}" 没有被任何公式引用，可能是孤岛数据`,
        });
      }
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings,
    };
  }

  /**
   * 检查字段是否被使用 v2.9.52: 基于引用图
   */
  private isFieldUsed(sheet: string, field: string, model: DataModel): boolean {
    if (!model.fieldIndex) {
      model.fieldIndex = this.buildFieldIndex(model);
    }

    const fieldKey = `${sheet}.${field}`;
    const entry = model.fieldIndex.get(fieldKey);

    return entry ? entry.usedBy.length > 0 : false;
  }

  /**
   * v2.9.52: 估算执行成本
   */
  estimateExecutionCost(model: DataModel): ExecutionCost {
    let readCells = 0;
    let writeCells = 0;
    let formulaWrites = 0;

    for (const table of model.tables) {
      const estimatedRows = 100; // 假设每表100行
      const _columnCount = table.fields.length;

      // 源字段需要读取
      const sourceFields = table.fields.filter((f) => f.type === "source").length;
      readCells += sourceFields * estimatedRows;

      // 派生字段需要写入公式
      const derivedFields = table.fields.filter(
        (f) => f.type === "derived" || f.type === "lookup"
      ).length;
      writeCells += derivedFields * estimatedRows;
      formulaWrites += derivedFields;
    }

    // 每个表一个 API 调用
    const apiCalls = model.tables.length * 2; // 读+写

    return {
      readCells,
      writeCells,
      formulaWrites,
      apiCalls,
      estimatedTimeMs: apiCalls * 100 + writeCells * 0.5,
    };
  }

  /**
   * v2.9.52: 评估风险
   */
  assessRisk(model: DataModel): RiskAssessment {
    const cost = this.estimateExecutionCost(model);
    const reasons: string[] = [];

    const hasDestructiveWrites = model.tables.some(
      (t) => t.role === "master" && t.fields.some((f) => f.type === "derived")
    );
    if (hasDestructiveWrites) {
      reasons.push("将修改主数据表的数据");
    }

    const hasFormulaWrites = cost.formulaWrites > 0;
    if (hasFormulaWrites) {
      reasons.push(`将写入 ${cost.formulaWrites} 个公式`);
    }

    const hasLargeRangeOperations = cost.writeCells > 1000;
    if (hasLargeRangeOperations) {
      reasons.push(`将影响超过 ${cost.writeCells} 个单元格`);
    }

    let level: "low" | "medium" | "high" = "low";
    if (hasDestructiveWrites || hasLargeRangeOperations) {
      level = "high";
    } else if (hasFormulaWrites) {
      level = "medium";
    }

    return {
      level,
      hasDestructiveWrites,
      hasFormulaWrites,
      hasLargeRangeOperations,
      requiresConfirmation: level !== "low",
      reasons,
    };
  }
}

// ========== 辅助类型 ==========

export interface DataModelAnalysis {
  tables: string[];
  fields: Map<string, string[]>;
  calculations: CalculationPattern[];
  suggestedModel: DataModel;
  // v2.7 新增
  identifiedFields?: Array<{
    name: string;
    fieldType: "source" | "derived" | "lookup";
    dataType: string;
  }>;
}

export interface CalculationPattern {
  field: string;
  formula: string;
  type: string;
  // v2.9.52: 增加 table 维度，避免同名字段冲突
  table?: string;
  appliesToTables?: string[];
}

// 导出单例
export const dataModeler = new DataModeler();
