/**
 * TaskPlanner - 任务规划引擎 v2.9.56
 *
 * 职责：
 * 1. 接收用户任务，进行深度分析
 * 2. 生成结构化的执行计划
 * 3. 验证计划的可行性
 * 4. 输出计划供用户确认或直接执行
 *
 * v2.9.56 重大修复：
 * 1. dependsOn 精确依赖 - 建立 fieldStepIdMap，按字段精确匹配
 * 2. translateFormula 结构化引用 - 使用 @[字段名] 或 {row} 模板
 * 3. 去掉硬编码 2:1000 - 规划层只输出列标识，执行层按真实行数写
 * 4. isWriteOperation 标记 - 所有写操作必须标记 + preview
 * 5. 禁用 execute_task - 无模型时返回 needsClarification
 * 6. 修正方法名拼写 - analyzeAndPlan（保留 deprecated alias）
 * 7. successCondition 必填 - 每个 step 和任务级都有成功条件
 * 8. 语义依赖检查 - 不只检查 stepId 存在，还检查逻辑依赖
 *
 * 核心原则：
 * - 规划层输出"可执行 + 可验证"的计划
 * - 公式使用结构化引用，执行层负责行数
 * - 每个步骤必须有明确的成功条件
 */

import { DataModeler, DataModel, ValidationResult } from "./DataModeler";

// ========== 类型定义 ==========

/**
 * 任务类型
 */
export type TaskType =
  | "data_modeling" // 数据建模（创建表结构）
  | "data_entry" // 数据录入
  | "formula_setup" // 公式设置
  | "data_analysis" // 数据分析
  | "formatting" // 格式化
  | "chart_creation" // 图表创建
  | "mixed"; // 混合任务

/**
 * 执行阶段
 */
export type ExecutionPhase =
  | "planning" // 规划中
  | "validation" // 验证中
  | "execution" // 执行中
  | "verification" // 校验中
  | "completed" // 已完成
  | "failed"; // 失败

/**
 * v2.9.56: 公式引用模式
 */
export type FormulaReferenceMode =
  | "structured" // Excel Table 结构化引用: @[字段名]
  | "row_template" // 行模板: {row} 占位符
  | "a1_fixed"; // A1 固定引用（仅用于单格，不推荐用于列）

/**
 * v2.9.56: 公式步骤参数（区别于普通参数）
 */
export interface FormulaStepParameters {
  sheet: string;
  column: string; // 只指定列（如 "D"），不指定行范围
  logicalFormula: string; // 逻辑公式（如 "=@[单价]*@[数量]"）
  referenceMode: FormulaReferenceMode;
  // 执行层根据真实行数决定范围
  // 如果有 Table → 用 Table 结构化引用
  // 如果无 Table → 读 usedRange 行数 N，写 column2:columnN
}

/**
 * v2.9.56: 写操作预览信息
 */
export interface WritePreview {
  affectedRange: string; // 预估影响范围（如 "D:D" 或 "A1:F100"）
  affectedCells: string; // 预估单元格数（如 "~100格" 或 "整列"）
  overwriteExisting: boolean; // 是否会覆盖现有数据
  warningMessage?: string; // 警告信息
}

/**
 * 计划步骤 v2.9.56: 可执行 + 可验证
 */
export interface PlanStep {
  id: string;
  order: number;
  phase:
    | "create_structure"
    | "write_data"
    | "set_formulas"
    | "add_validation"
    | "format"
    | "verify"
    | "read_data"
    | "analyze"
    // v4.0 新增 phase
    | "sensing" // 感知阶段（读取当前状态）
    | "response" // 响应阶段（回复用户）
    | "chart" // 图表阶段
    | "sheet" // 工作表操作阶段
    | "execution"; // 通用执行阶段
  description: string;

  // 具体操作
  action: string;
  parameters: Record<string, unknown>;

  // v2.9.56: 精确依赖（stepId 列表，按 fieldStepIdMap 生成）
  dependsOn: string[];

  // v2.9.56: 成功条件（必填！）
  successCondition: SuccessCondition;

  // v2.9.56: 写操作标记 + 预览
  isWriteOperation: boolean;
  writePreview?: WritePreview;

  // 验证条件（旧版兼容，会被迁移到 successCondition）
  verificationCondition?: string;

  // 状态
  status: "pending" | "running" | "completed" | "failed" | "skipped";
  result?: StepResult;
}

/**
 * v2.9.56: 成功条件定义（更丰富）
 */
export interface SuccessCondition {
  type:
    | "tool_success" // 工具返回 success=true 即可
    | "value_check" // 检查某个单元格值
    | "range_exists" // 检查范围是否有数据
    | "formula_result" // 检查公式计算结果（抽样几格不为错误）
    | "sheet_exists" // 检查工作表存在
    | "headers_match" // 检查表头匹配
    | "no_error_values" // 检查范围内无 #REF! #VALUE! 等
    | "custom"; // 自定义检查函数

  expectedValue?: unknown;
  targetRange?: string;
  targetSheet?: string;
  expectedHeaders?: string[];
  tolerance?: number; // 数值允许误差
  sampleCount?: number; // 抽样检查数量（用于 formula_result）
  customFn?: string; // 自定义检查函数名
}

/**
 * 步骤执行结果
 */
export interface StepResult {
  success: boolean;
  output?: string;
  error?: string;
  duration?: number;
  verificationPassed?: boolean;
  // v2.9.56: 实际影响
  actualAffectedRange?: string;
  actualAffectedCells?: number;
  // v2.9.60: 警告信息（降级执行等场景）
  warning?: string;
}

/**
 * 执行计划 v2.9.56: 可执行 + 可验证
 */
export interface ExecutionPlan {
  id: string;
  taskDescription: string;
  taskType: TaskType;

  // v2.9.60: 补充缺失的 goal 属性
  goal?: string;

  // 数据模型（如果适用）
  dataModel?: DataModel;
  modelValidation?: ValidationResult;

  // 执行步骤
  steps: PlanStep[];

  // v2.9.56: 任务级成功条件（必填！）
  taskSuccessConditions: TaskSuccessCondition[];

  // v2.9.56: 任务完成时的回复模板（基于真实结果生成）
  completionMessage?: string;

  // 依赖检查结果
  dependencyCheck: DependencyCheckResult;

  // v2.9.56: 字段到步骤 ID 的映射（用于精确依赖）
  fieldStepIdMap: FieldStepIdMap;

  // v2.9.56: 需要澄清（无法生成可执行计划时）
  needsClarification?: boolean;
  clarificationMessage?: string;

  // 预估
  estimatedDuration: number;
  estimatedSteps: number;

  // 风险评估
  risks: RiskAssessment[];

  // 状态
  phase: ExecutionPhase;
  currentStep: number;
  completedSteps: number;
  failedSteps: number;
}

/**
 * v2.9.56: 字段到步骤 ID 映射
 * 结构: { "订单明细": { "金额": "step-123", "数量": "step-456" } }
 */
export type FieldStepIdMap = Record<string, Record<string, string>>;

/**
 * v2.9.56: 任务级成功条件
 */
export interface TaskSuccessCondition {
  id: string;
  description: string;
  type:
    | "all_steps_complete" // 所有步骤完成
    | "specific_steps_complete" // 指定步骤完成
    | "final_verify_passed" // 终验通过
    | "value_check" // 值检查
    | "custom"; // 自定义
  stepIds?: string[]; // specific_steps_complete 时使用
  checkConfig?: SuccessCondition;
  priority: number; // 检查优先级
}

/**
 * v2.9.56: 语义依赖检查项
 */
export interface SemanticDependency {
  sourceSheet: string;
  sourceField: string;
  targetSheet: string;
  targetField: string;
  dependencyType: "formula_reference" | "lookup_source" | "data_source";
  isResolved: boolean;
  resolvedStepId?: string;
}

/**
 * 依赖检查结果 v2.9.56: 增加语义检查
 */
export interface DependencyCheckResult {
  passed: boolean;
  missingDependencies: string[];
  circularDependencies: string[];
  warnings: string[];
  // v2.9.56: 语义依赖分析
  semanticDependencies: SemanticDependency[];
  unresolvedSemanticDeps: SemanticDependency[];
}

/**
 * 风险评估
 */
export interface RiskAssessment {
  level: "low" | "medium" | "high";
  description: string;
  mitigation: string;
}

// ========== 任务规划引擎 ==========

export class TaskPlanner {
  private dataModeler: DataModeler;

  constructor() {
    this.dataModeler = new DataModeler();
  }

  /**
   * v2.9.56: 正确拼写的方法名
   * 分析任务并生成执行计划
   */
  async analyzeAndPlan(taskDescription: string): Promise<ExecutionPlan> {
    const planId = this.generateId();

    // 1. 识别任务类型
    const taskType = this.identifyTaskType(taskDescription);

    // 2. 如果是数据建模任务，先生成数据模型
    let dataModel: DataModel | undefined;
    let modelValidation: ValidationResult | undefined;

    if (taskType === "data_modeling" || taskType === "mixed") {
      const analysis = this.dataModeler.analyzeRequirement(taskDescription);
      dataModel = analysis.suggestedModel;
      modelValidation = this.dataModeler.validateModel(dataModel);
    }

    // 3. 建立字段到步骤 ID 的映射（用于精确依赖）
    const fieldStepIdMap: FieldStepIdMap = {};

    // 4. 生成执行步骤（带精确依赖）
    const steps = this.generateSteps(taskDescription, taskType, dataModel, fieldStepIdMap);

    // 5. 如果没有生成有效步骤，返回需要澄清
    if (steps.length === 0 || this.hasUnexecutableSteps(steps)) {
      return this.createClarificationPlan(planId, taskDescription, taskType);
    }

    // 6. 语义依赖检查（不只检查 stepId 存在）
    const dependencyCheck = this.checkDependenciesWithSemantics(steps, dataModel, fieldStepIdMap);

    // 7. 风险评估
    const risks = this.assessRisks(taskDescription, steps, dependencyCheck);

    // 8. 生成任务级成功条件
    const taskSuccessConditions = this.generateTaskSuccessConditions(steps, dataModel);

    // 9. 构建执行计划
    const plan: ExecutionPlan = {
      id: planId,
      taskDescription,
      taskType,
      dataModel,
      modelValidation,
      steps,
      taskSuccessConditions,
      dependencyCheck,
      fieldStepIdMap,
      estimatedDuration: steps.length * 2000,
      estimatedSteps: steps.length,
      risks,
      phase: "planning",
      currentStep: 0,
      completedSteps: 0,
      failedSteps: 0,
    };

    return plan;
  }

  /**
   * v2.9.56: 保留旧方法名作为 deprecated alias
   * @deprecated 使用 analyzeAndPlan 代替
   */
  async analyzAndPlan(taskDescription: string): Promise<ExecutionPlan> {
    console.warn("[TaskPlanner] analyzAndPlan 已废弃，请使用 analyzeAndPlan");
    return this.analyzeAndPlan(taskDescription);
  }

  /**
   * v2.9.56: 检查是否有不可执行的步骤
   */
  private hasUnexecutableSteps(steps: PlanStep[]): boolean {
    const unexecutableActions = ["execute_task"]; // 黑名单
    return steps.some((s) => unexecutableActions.includes(s.action));
  }

  /**
   * v2.9.56: 创建需要澄清的计划
   */
  private createClarificationPlan(
    planId: string,
    taskDescription: string,
    taskType: TaskType
  ): ExecutionPlan {
    return {
      id: planId,
      taskDescription,
      taskType,
      steps: [],
      taskSuccessConditions: [],
      dependencyCheck: {
        passed: false,
        missingDependencies: ["无法理解任务或生成可执行步骤"],
        circularDependencies: [],
        warnings: [],
        semanticDependencies: [],
        unresolvedSemanticDeps: [],
      },
      fieldStepIdMap: {},
      needsClarification: true,
      clarificationMessage: this.generateClarificationMessage(taskDescription),
      estimatedDuration: 0,
      estimatedSteps: 0,
      risks: [],
      phase: "planning",
      currentStep: 0,
      completedSteps: 0,
      failedSteps: 0,
    };
  }

  /**
   * v2.9.56: 生成澄清请求消息
   */
  private generateClarificationMessage(taskDescription: string): string {
    const suggestions: string[] = [];

    // 根据任务描述给出具体建议
    if (taskDescription.length < 20) {
      suggestions.push("请提供更详细的任务描述");
    }
    if (!/(表|工作表|Sheet)/.test(taskDescription)) {
      suggestions.push("请指明要操作的工作表名称");
    }
    if (!/(列|字段|单元格|范围|区域)/.test(taskDescription)) {
      suggestions.push("请指明要操作的列或范围");
    }
    if (/(公式|计算)/.test(taskDescription) && !/(=|求和|乘|加|减)/.test(taskDescription)) {
      suggestions.push("请说明具体的计算逻辑（如：金额 = 单价 * 数量）");
    }

    if (suggestions.length === 0) {
      suggestions.push("请提供更具体的操作要求");
    }

    return `我需要更多信息才能执行这个任务：\n${suggestions.map((s) => `• ${s}`).join("\n")}`;
  }

  /**
   * 识别任务类型
   */
  private identifyTaskType(description: string): TaskType {
    const patterns: Array<{ pattern: RegExp; type: TaskType }> = [
      { pattern: /创建.*表|新建.*工作表|设计.*结构|建立.*模型/i, type: "data_modeling" },
      { pattern: /录入|输入|填写|添加.*数据/i, type: "data_entry" },
      { pattern: /公式|计算|求和|XLOOKUP|SUMIF/i, type: "formula_setup" },
      { pattern: /分析|统计|汇总|趋势|洞察/i, type: "data_analysis" },
      { pattern: /格式|颜色|字体|样式|美化/i, type: "formatting" },
      { pattern: /图表|柱状图|折线图|饼图/i, type: "chart_creation" },
    ];

    const matchedTypes: TaskType[] = [];

    for (const { pattern, type } of patterns) {
      if (pattern.test(description)) {
        matchedTypes.push(type);
      }
    }

    if (matchedTypes.length === 0) {
      return "mixed";
    } else if (matchedTypes.length === 1) {
      return matchedTypes[0];
    } else {
      return "mixed";
    }
  }

  /**
   * v2.9.56: 生成执行步骤（带精确依赖 + 结构化公式 + successCondition）
   */
  private generateSteps(
    description: string,
    taskType: TaskType,
    dataModel: DataModel | undefined,
    fieldStepIdMap: FieldStepIdMap
  ): PlanStep[] {
    const steps: PlanStep[] = [];
    let order = 0;

    // 如果有数据模型，按模型生成步骤
    if (dataModel) {
      // 阶段 1: 创建表结构
      for (const tableName of dataModel.executionOrder) {
        const table = dataModel.tables.find((t) => t.name === tableName);
        if (!table) continue;

        // 初始化该表的字段映射
        if (!fieldStepIdMap[tableName]) {
          fieldStepIdMap[tableName] = {};
        }

        // 创建工作表
        const createSheetStepId = this.generateId();
        steps.push({
          id: createSheetStepId,
          order: order++,
          phase: "create_structure",
          description: `\u521B\u5EFA\u5DE5\u4F5C\u8868: ${tableName}`,
          action: "excel_create_sheet",
          parameters: { name: tableName },
          dependsOn: this.resolveTableDependencies(table.dependsOn || [], fieldStepIdMap),
          isWriteOperation: true,
          writePreview: {
            affectedRange: `\u65B0\u5DE5\u4F5C\u8868 "${tableName}"`,
            affectedCells: "0\u683C\uFF08\u65B0\u5EFA\uFF09",
            overwriteExisting: false,
          },
          successCondition: {
            type: "sheet_exists",
            targetSheet: tableName,
          },
          verificationCondition: `\u5DE5\u4F5C\u8868 "${tableName}" \u5B58\u5728`,
          status: "pending",
        });

        // 写入表头
        const headers = table.fields.map((f) => f.name);
        const writeHeadersStepId = this.generateId();
        steps.push({
          id: writeHeadersStepId,
          order: order++,
          phase: "write_data",
          description: `\u5199\u5165\u8868\u5934: ${tableName}`,
          action: "excel_write_range",
          parameters: {
            sheet: tableName,
            range: `A1:${this.indexToColumn(headers.length)}1`,
            values: [headers],
          },
          dependsOn: [createSheetStepId],
          isWriteOperation: true,
          writePreview: {
            affectedRange: `A1:${this.indexToColumn(headers.length)}1`,
            affectedCells: `${headers.length}\u683C`,
            overwriteExisting: false,
          },
          successCondition: {
            type: "headers_match",
            targetSheet: tableName,
            targetRange: `A1:${this.indexToColumn(headers.length)}1`,
            expectedHeaders: headers,
          },
          status: "pending",
        });

        // 记录每个字段的"存在性"步骤（用于依赖解析）
        for (const field of table.fields) {
          fieldStepIdMap[tableName][field.name] = writeHeadersStepId;
        }
      }

      // 阶段 2: 设置公式（按计算链顺序 + 结构化引用）
      for (const calcStep of dataModel.calculationChain) {
        const table = dataModel.tables.find((t) => t.name === calcStep.sheet);
        const field = table?.fields.find((f) => f.name === calcStep.field);

        if (field && field.formula) {
          const fieldIndex = table!.fields.findIndex((f) => f.name === calcStep.field);
          const column = this.indexToColumn(fieldIndex + 1);

          // v2.9.56: 转换为结构化引用公式
          const structuredFormula = this.translateToStructuredFormula(field.formula, table!.fields);

          const formulaStepId = this.generateId();

          // v2.9.56: 精确依赖 - 按字段解析，不是按 sheet 名模糊匹配
          const preciseDependsOn = this.resolvePreciseDependencies(
            calcStep.dependencies,
            fieldStepIdMap
          );

          steps.push({
            id: formulaStepId,
            order: order++,
            phase: "set_formulas",
            description: `\u8BBE\u7F6E\u516C\u5F0F: ${calcStep.sheet}.${calcStep.field}`,
            action: "excel_set_formula",
            parameters: {
              sheet: calcStep.sheet,
              column: column, // v2.9.56: 只指定列，不指定行
              logicalFormula: structuredFormula, // v2.9.56: 逻辑公式
              referenceMode: "structured" as FormulaReferenceMode, // v2.9.56: 使用结构化引用
              // 执行层会根据真实行数决定范围
            } as FormulaStepParameters as unknown as Record<string, unknown>,
            dependsOn: preciseDependsOn,
            isWriteOperation: true,
            writePreview: {
              affectedRange: `${calcStep.sheet}!${column}:\u6574\u5217`,
              affectedCells: "\u6839\u636E\u5B9E\u9645\u6570\u636E\u884C\u6570",
              overwriteExisting: true,
              warningMessage: "\u5C06\u8986\u76D6\u8BE5\u5217\u73B0\u6709\u516C\u5F0F",
            },
            successCondition: {
              type: "no_error_values",
              targetSheet: calcStep.sheet,
              targetRange: `${column}:\u6574\u5217`, // 执行层解析
              sampleCount: 5, // 抽样 5 个格子检查
            },
            verificationCondition: `${calcStep.sheet}!${column} \u65E0\u9519\u8BEF\u503C`,
            status: "pending",
          });

          // 更新字段映射
          fieldStepIdMap[calcStep.sheet][calcStep.field] = formulaStepId;
        }
      }

      // 阶段 3: 添加数据验证
      for (const table of dataModel.tables) {
        for (const field of table.fields) {
          if (field.validation) {
            const fieldIndex = table.fields.findIndex((f) => f.name === field.name);
            const column = this.indexToColumn(fieldIndex + 1);

            steps.push({
              id: this.generateId(),
              order: order++,
              phase: "add_validation",
              description: `\u6DFB\u52A0\u9A8C\u8BC1: ${table.name}.${field.name}`,
              action: "excel_add_data_validation",
              parameters: {
                sheet: table.name,
                column: column, // v2.9.56: 只指定列
                type: field.validation.type,
                values: field.validation.values,
              },
              dependsOn: fieldStepIdMap[table.name]?.[field.name]
                ? [fieldStepIdMap[table.name][field.name]]
                : [],
              isWriteOperation: true,
              writePreview: {
                affectedRange: `${table.name}!${column}:\u6574\u5217`,
                affectedCells: "\u6839\u636E\u5B9E\u9645\u6570\u636E\u884C\u6570",
                overwriteExisting: false,
              },
              successCondition: {
                type: "tool_success",
              },
              status: "pending",
            });
          }
        }
      }

      // 阶段 4: 验证步骤
      const verifyStepId = this.generateId();
      steps.push({
        id: verifyStepId,
        order: order++,
        phase: "verify",
        description: "\u9A8C\u8BC1\u6240\u6709\u516C\u5F0F\u548C\u6570\u636E",
        action: "verify_execution",
        parameters: {
          sheets: dataModel.tables.map((t) => t.name),
          checkFormulas: true,
          sampleRows: 5,
        },
        dependsOn: steps.filter((s) => s.phase === "set_formulas").map((s) => s.id),
        isWriteOperation: false,
        successCondition: {
          type: "no_error_values",
          sampleCount: 10,
        },
        status: "pending",
      });
    }
    // v2.9.56: 没有数据模型时，不生成 execute_task，返回空步骤（触发澄清）

    return steps;
  }

  /**
   * v2.9.56: 转换为结构化引用公式
   * 输入: "=单价*数量"
   * 输出: "=@[单价]*@[数量]"（Excel Table 结构化引用）
   */
  private translateToStructuredFormula(formula: string, fields: Array<{ name: string }>): string {
    let translated = formula;

    // 按字段名长度降序排序，避免短名称误替换长名称的一部分
    const sortedFields = [...fields].sort((a, b) => b.name.length - a.name.length);

    for (const field of sortedFields) {
      // 替换字段名为结构化引用 @[字段名]
      // 使用负向断言避免重复替换
      const pattern = new RegExp(
        `(?<!@\\[)(?<![A-Za-z\u4e00-\u9fa5])${this.escapeRegex(field.name)}(?![A-Za-z\u4e00-\u9fa5])(?!\\])`,
        "g"
      );
      translated = translated.replace(pattern, `@[${field.name}]`);
    }

    return translated;
  }

  /**
   * v2.9.56: 转义正则特殊字符
   */
  private escapeRegex(str: string): string {
    return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }

  /**
   * v2.9.56: 解析精确依赖（按字段，不是按 sheet 名）
   * 输入: ["订单明细!数量", "订单明细!单价"]
   * 输出: ["step-123", "step-456"]（对应的步骤 ID）
   */
  private resolvePreciseDependencies(
    dependencies: string[],
    fieldStepIdMap: FieldStepIdMap
  ): string[] {
    const resolvedIds: string[] = [];

    for (const dep of dependencies) {
      // 解析 "Sheet!Field" 格式
      const parts = dep.split("!");
      if (parts.length !== 2) continue;

      const [sheetName, fieldName] = parts;
      const stepId = fieldStepIdMap[sheetName]?.[fieldName];

      if (stepId) {
        if (!resolvedIds.includes(stepId)) {
          resolvedIds.push(stepId);
        }
      } else {
        console.warn(`[TaskPlanner] \u65E0\u6CD5\u89E3\u6790\u4F9D\u8D56: ${dep}`);
      }
    }

    return resolvedIds;
  }

  /**
   * v2.9.56: 解析表依赖
   */
  private resolveTableDependencies(tableDeps: string[], fieldStepIdMap: FieldStepIdMap): string[] {
    const resolvedIds: string[] = [];

    for (const tableName of tableDeps) {
      // 获取该表的任意一个字段的步骤 ID
      const tableFields = fieldStepIdMap[tableName];
      if (tableFields) {
        const firstStepId = Object.values(tableFields)[0];
        if (firstStepId && !resolvedIds.includes(firstStepId)) {
          resolvedIds.push(firstStepId);
        }
      }
    }

    return resolvedIds;
  }

  /**
   * v2.9.56: 语义依赖检查（不只检查 stepId 存在）
   */
  private checkDependenciesWithSemantics(
    steps: PlanStep[],
    dataModel: DataModel | undefined,
    fieldStepIdMap: FieldStepIdMap
  ): DependencyCheckResult {
    const missingDependencies: string[] = [];
    const circularDependencies: string[] = [];
    const warnings: string[] = [];
    const semanticDependencies: SemanticDependency[] = [];
    const unresolvedSemanticDeps: SemanticDependency[] = [];

    // 1. 检查步骤间依赖（stepId 存在性）
    const stepIds = new Set(steps.map((s) => s.id));
    for (const step of steps) {
      for (const dep of step.dependsOn) {
        if (dep && !stepIds.has(dep)) {
          missingDependencies.push(
            `\u6B65\u9AA4 "${step.description}" \u4F9D\u8D56\u4E0D\u5B58\u5728\u7684\u6B65\u9AA4: ${dep}`
          );
        }
      }
    }

    // 2. 语义依赖检查（公式引用、lookup 源等）
    if (dataModel) {
      for (const calcStep of dataModel.calculationChain) {
        for (const dep of calcStep.dependencies) {
          const parts = dep.split("!");
          if (parts.length !== 2) continue;

          const [depSheet, depField] = parts;

          const semDep: SemanticDependency = {
            sourceSheet: calcStep.sheet,
            sourceField: calcStep.field,
            targetSheet: depSheet,
            targetField: depField,
            dependencyType: "formula_reference",
            isResolved: false,
          };

          // 检查依赖是否存在于步骤中
          const resolvedStepId = fieldStepIdMap[depSheet]?.[depField];
          if (resolvedStepId) {
            semDep.isResolved = true;
            semDep.resolvedStepId = resolvedStepId;
            semanticDependencies.push(semDep);
          } else {
            // 检查目标表是否存在于模型中
            const targetTable = dataModel.tables.find((t) => t.name === depSheet);
            if (targetTable) {
              const targetField = targetTable.fields.find((f) => f.name === depField);
              if (targetField) {
                // 字段存在但没有对应步骤 - 可能是源数据字段，不需要计算
                semDep.isResolved = true;
                semanticDependencies.push(semDep);
              } else {
                semDep.isResolved = false;
                unresolvedSemanticDeps.push(semDep);
                warnings.push(
                  `\u516C\u5F0F ${calcStep.sheet}.${calcStep.field} \u5F15\u7528\u4E86\u4E0D\u5B58\u5728\u7684\u5B57\u6BB5 ${depSheet}.${depField}`
                );
              }
            } else {
              semDep.isResolved = false;
              unresolvedSemanticDeps.push(semDep);
              warnings.push(
                `\u516C\u5F0F ${calcStep.sheet}.${calcStep.field} \u5F15\u7528\u4E86\u4E0D\u5B58\u5728\u7684\u8868 ${depSheet}`
              );
            }
          }
        }
      }

      // 3. 检查 lookup 依赖（XLOOKUP/VLOOKUP 需要主数据表存在）
      for (const table of dataModel.tables) {
        for (const field of table.fields) {
          if (field.formula && /XLOOKUP|VLOOKUP|INDEX.*MATCH/i.test(field.formula)) {
            // 这是一个 lookup 公式，需要确保查找源存在
            const lookupSource = this.extractLookupSource(field.formula, dataModel);
            if (lookupSource) {
              const semDep: SemanticDependency = {
                sourceSheet: table.name,
                sourceField: field.name,
                targetSheet: lookupSource.sheet,
                targetField: lookupSource.field,
                dependencyType: "lookup_source",
                isResolved: false,
              };

              // 检查 lookup 源是否存在
              const sourceTable = dataModel.tables.find((t) => t.name === lookupSource.sheet);
              if (sourceTable) {
                const sourceField = sourceTable.fields.find((f) => f.name === lookupSource.field);
                if (sourceField) {
                  semDep.isResolved = true;
                  semDep.resolvedStepId = fieldStepIdMap[lookupSource.sheet]?.[lookupSource.field];
                } else {
                  semDep.isResolved = false;
                  unresolvedSemanticDeps.push(semDep);
                  warnings.push(
                    `LOOKUP \u516C\u5F0F ${table.name}.${field.name} \u7684\u67E5\u627E\u6E90\u5B57\u6BB5 ${lookupSource.sheet}.${lookupSource.field} \u4E0D\u5B58\u5728`
                  );
                }
              } else {
                semDep.isResolved = false;
                unresolvedSemanticDeps.push(semDep);
                warnings.push(
                  `LOOKUP \u516C\u5F0F ${table.name}.${field.name} \u7684\u67E5\u627E\u6E90\u8868 ${lookupSource.sheet} \u4E0D\u5B58\u5728`
                );
              }
              semanticDependencies.push(semDep);
            }
          }
        }
      }

      // 4. DataModeler 的验证结果
      const validation = this.dataModeler.validateModel(dataModel);
      for (const error of validation.errors) {
        if (error.type === "missing_dependency") {
          missingDependencies.push(error.message);
        } else if (error.type === "circular_reference") {
          circularDependencies.push(error.message);
        }
      }
      for (const warning of validation.warnings) {
        warnings.push(warning.message);
      }
    }

    return {
      passed:
        missingDependencies.length === 0 &&
        circularDependencies.length === 0 &&
        unresolvedSemanticDeps.length === 0,
      missingDependencies,
      circularDependencies,
      warnings,
      semanticDependencies,
      unresolvedSemanticDeps,
    };
  }

  /**
   * v2.9.56: 从 LOOKUP 公式中提取查找源
   */
  private extractLookupSource(
    formula: string,
    dataModel: DataModel
  ): { sheet: string; field: string } | null {
    // 简化处理：从公式中尝试提取表名和字段名
    // 实际上应该用更复杂的公式解析
    for (const table of dataModel.tables) {
      if (formula.includes(table.name)) {
        for (const field of table.fields) {
          if (formula.includes(field.name)) {
            return { sheet: table.name, field: field.name };
          }
        }
      }
    }
    return null;
  }

  /**
   * v2.9.56: 生成任务级成功条件
   */
  private generateTaskSuccessConditions(
    steps: PlanStep[],
    _dataModel: DataModel | undefined
  ): TaskSuccessCondition[] {
    const conditions: TaskSuccessCondition[] = [];

    // 条件 1: 所有步骤完成
    conditions.push({
      id: this.generateId(),
      description: "\u6240\u6709\u6B65\u9AA4\u5747\u5DF2\u5B8C\u6210",
      type: "all_steps_complete",
      priority: 1,
    });

    // 条件 2: 公式步骤无错误
    const formulaSteps = steps.filter((s) => s.phase === "set_formulas");
    if (formulaSteps.length > 0) {
      conditions.push({
        id: this.generateId(),
        description: "\u6240\u6709\u516C\u5F0F\u6B65\u9AA4\u65E0\u9519\u8BEF\u503C",
        type: "specific_steps_complete",
        stepIds: formulaSteps.map((s) => s.id),
        checkConfig: {
          type: "no_error_values",
          sampleCount: 5,
        },
        priority: 2,
      });
    }

    // 条件 3: 终验通过
    const verifySteps = steps.filter((s) => s.phase === "verify");
    if (verifySteps.length > 0) {
      conditions.push({
        id: this.generateId(),
        description: "\u7EC8\u9A8C\u6B65\u9AA4\u901A\u8FC7",
        type: "final_verify_passed",
        stepIds: verifySteps.map((s) => s.id),
        priority: 3,
      });
    }

    return conditions;
  }

  /**
   * 风险评估
   */
  private assessRisks(
    description: string,
    steps: PlanStep[],
    dependencyCheck: DependencyCheckResult
  ): RiskAssessment[] {
    const risks: RiskAssessment[] = [];

    // 依赖风险
    if (!dependencyCheck.passed) {
      risks.push({
        level: "high",
        description: "\u5B58\u5728\u672A\u89E3\u51B3\u7684\u4F9D\u8D56\u95EE\u9898",
        mitigation: "\u8BF7\u5148\u4FEE\u590D\u4F9D\u8D56\u95EE\u9898\u518D\u6267\u884C",
      });
    }

    // 语义依赖风险
    if (dependencyCheck.unresolvedSemanticDeps.length > 0) {
      risks.push({
        level: "high",
        description: `\u5B58\u5728 ${dependencyCheck.unresolvedSemanticDeps.length} \u4E2A\u672A\u89E3\u6790\u7684\u8BED\u4E49\u4F9D\u8D56`,
        mitigation:
          "\u8BF7\u68C0\u67E5\u516C\u5F0F\u5F15\u7528\u7684\u8868\u548C\u5B57\u6BB5\u662F\u5426\u5B58\u5728",
      });
    }

    // 复杂度风险
    if (steps.length > 20) {
      risks.push({
        level: "medium",
        description:
          "\u4EFB\u52A1\u6B65\u9AA4\u8F83\u591A\uFF0C\u6267\u884C\u65F6\u95F4\u53EF\u80FD\u8F83\u957F",
        mitigation: "\u5EFA\u8BAE\u5206\u6279\u6267\u884C\u6216\u76D1\u63A7\u8FDB\u5EA6",
      });
    }

    // 公式风险
    const formulaSteps = steps.filter((s) => s.phase === "set_formulas");
    if (formulaSteps.length > 5) {
      risks.push({
        level: "medium",
        description:
          "\u5B58\u5728\u591A\u4E2A\u516C\u5F0F\u4F9D\u8D56\uFF0C\u53EF\u80FD\u51FA\u73B0\u8BA1\u7B97\u9519\u8BEF",
        mitigation: "\u6BCF\u4E2A\u516C\u5F0F\u8BBE\u7F6E\u540E\u4F1A\u8FDB\u884C\u9A8C\u8BC1",
      });
    }

    // 跨表引用风险
    if (/XLOOKUP|VLOOKUP|INDIRECT/.test(description)) {
      risks.push({
        level: "medium",
        description:
          "\u4F7F\u7528\u8DE8\u8868\u67E5\u627E\u51FD\u6570\uFF0C\u53EF\u80FD\u56E0\u6570\u636E\u6E90\u95EE\u9898\u51FA\u9519",
        mitigation:
          "\u786E\u4FDD\u6570\u636E\u6E90\u8868\u5DF2\u521B\u5EFA\u4E14\u6709\u6570\u636E",
      });
    }

    // 写操作风险
    const writeSteps = steps.filter((s) => s.isWriteOperation);
    if (writeSteps.length > 0) {
      const overwriteSteps = writeSteps.filter((s) => s.writePreview?.overwriteExisting);
      if (overwriteSteps.length > 0) {
        risks.push({
          level: "medium",
          description: `${overwriteSteps.length} \u4E2A\u6B65\u9AA4\u4F1A\u8986\u76D6\u73B0\u6709\u6570\u636E`,
          mitigation:
            "\u6267\u884C\u524D\u8BF7\u786E\u8BA4\u9884\u89C8\uFF0C\u6216\u5148\u5907\u4EFD",
        });
      }
    }

    return risks;
  }

  /**
   * v2.9.56: 格式化计划输出（包含写操作预览）
   */
  formatPlanForDisplay(plan: ExecutionPlan): string {
    const lines: string[] = [];

    // 如果需要澄清，直接返回澄清消息
    if (plan.needsClarification) {
      lines.push("\u2753 **\u9700\u8981\u66F4\u591A\u4FE1\u606F**");
      lines.push("");
      lines.push(
        plan.clarificationMessage ||
          "\u8BF7\u63D0\u4F9B\u66F4\u5177\u4F53\u7684\u4EFB\u52A1\u63CF\u8FF0"
      );
      return lines.join("\n");
    }

    lines.push("\uD83D\uDCCB **\u6267\u884C\u8BA1\u5212**");
    lines.push("");
    lines.push(`**\u4EFB\u52A1**: ${plan.taskDescription.substring(0, 100)}...`);
    lines.push(`**\u7C7B\u578B**: ${this.translateTaskType(plan.taskType)}`);
    lines.push(`**\u9884\u8BA1\u6B65\u9AA4**: ${plan.estimatedSteps}`);
    lines.push("");

    // 数据模型
    if (plan.dataModel) {
      lines.push("\uD83D\uDCCA **\u6570\u636E\u6A21\u578B**");
      lines.push("");
      lines.push(
        `\u8868\u7ED3\u6784: ${plan.dataModel.tables.map((t) => t.name).join(" \u2192 ")}`
      );
      lines.push("");

      for (const table of plan.dataModel.tables) {
        lines.push(`**${table.name}** (${this.translateTableRole(table.role)})`);
        lines.push(`  \u5B57\u6BB5: ${table.fields.map((f) => f.name).join(", ")}`);
        if (table.dependsOn && table.dependsOn.length > 0) {
          lines.push(`  \u4F9D\u8D56: ${table.dependsOn.join(", ")}`);
        }
        lines.push("");
      }
    }

    // 依赖检查
    if (!plan.dependencyCheck.passed) {
      lines.push("\u26A0\uFE0F **\u4F9D\u8D56\u95EE\u9898**");
      for (const issue of plan.dependencyCheck.missingDependencies) {
        lines.push(`  - ${issue}`);
      }
      for (const issue of plan.dependencyCheck.circularDependencies) {
        lines.push(`  - ${issue}`);
      }
      // v2.9.56: 显示未解析的语义依赖
      for (const dep of plan.dependencyCheck.unresolvedSemanticDeps) {
        lines.push(
          `  - \u672A\u89E3\u6790: ${dep.sourceSheet}.${dep.sourceField} \u2192 ${dep.targetSheet}.${dep.targetField}`
        );
      }
      lines.push("");
    }

    // 风险
    if (plan.risks.length > 0) {
      lines.push("\u26A0\uFE0F **\u98CE\u9669\u8BC4\u4F30**");
      for (const risk of plan.risks) {
        const icon =
          risk.level === "high"
            ? "\uD83D\uDD34"
            : risk.level === "medium"
              ? "\uD83D\uDFE1"
              : "\uD83D\uDFE2";
        lines.push(`  ${icon} ${risk.description}`);
      }
      lines.push("");
    }

    // v2.9.56: 写操作预览
    const writeSteps = plan.steps.filter((s) => s.isWriteOperation && s.writePreview);
    if (writeSteps.length > 0) {
      lines.push("\u270F\uFE0F **\u5199\u64CD\u4F5C\u9884\u89C8**");
      for (const step of writeSteps) {
        const preview = step.writePreview!;
        const overwriteIcon = preview.overwriteExisting ? "\u26A0\uFE0F" : "\u2705";
        lines.push(`  ${overwriteIcon} ${step.description}`);
        lines.push(
          `      \u8303\u56F4: ${preview.affectedRange} (\u7EA6 ${preview.affectedCells})`
        );
        if (preview.warningMessage) {
          lines.push(`      \u26A0\uFE0F ${preview.warningMessage}`);
        }
      }
      lines.push("");
    }

    // 执行步骤
    lines.push("\uD83D\uDCDD **\u6267\u884C\u6B65\u9AA4**");
    lines.push("");

    let currentPhase = "";
    for (const step of plan.steps) {
      if (step.phase !== currentPhase) {
        currentPhase = step.phase;
        lines.push(`**\u9636\u6BB5: ${this.translatePhase(step.phase)}**`);
      }

      const statusIcon =
        step.status === "completed"
          ? "\u2705"
          : step.status === "failed"
            ? "\u274C"
            : step.status === "running"
              ? "\uD83D\uDD04"
              : "\u23F3";
      const writeIcon = step.isWriteOperation ? "\u270F\uFE0F" : "";
      lines.push(`  ${statusIcon}${writeIcon} ${step.order + 1}. ${step.description}`);
    }

    // v2.9.56: 任务成功条件
    if (plan.taskSuccessConditions.length > 0) {
      lines.push("");
      lines.push("\u2705 **\u6210\u529F\u6761\u4EF6**");
      for (const cond of plan.taskSuccessConditions) {
        lines.push(`  - ${cond.description}`);
      }
    }

    return lines.join("\n");
  }

  /**
   * 翻译任务类型
   */
  private translateTaskType(type: TaskType): string {
    const translations: Record<TaskType, string> = {
      data_modeling: "\u6570\u636E\u5EFA\u6A21",
      data_entry: "\u6570\u636E\u5F55\u5165",
      formula_setup: "\u516C\u5F0F\u8BBE\u7F6E",
      data_analysis: "\u6570\u636E\u5206\u6790",
      formatting: "\u683C\u5F0F\u5316",
      chart_creation: "\u56FE\u8868\u521B\u5EFA",
      mixed: "\u7EFC\u5408\u4EFB\u52A1",
    };
    return translations[type] || type;
  }

  /**
   * 翻译表角色
   */
  private translateTableRole(role: string): string {
    const translations: Record<string, string> = {
      master: "\u4E3B\u6570\u636E\u8868",
      transaction: "\u4EA4\u6613\u6570\u636E\u8868",
      summary: "\u6C47\u603B\u8868",
      analysis: "\u5206\u6790\u8868",
    };
    return translations[role] || role;
  }

  /**
   * 翻译阶段
   */
  private translatePhase(phase: string): string {
    const translations: Record<string, string> = {
      create_structure: "\u521B\u5EFA\u7ED3\u6784",
      write_data: "\u5199\u5165\u6570\u636E",
      set_formulas: "\u8BBE\u7F6E\u516C\u5F0F",
      add_validation: "\u6DFB\u52A0\u9A8C\u8BC1",
      format: "\u683C\u5F0F\u5316",
      verify: "\u9A8C\u8BC1\u68C0\u67E5",
      read_data: "\u8BFB\u53D6\u6570\u636E",
      analyze: "\u5206\u6790\u6570\u636E",
    };
    return translations[phase] || phase;
  }

  /**
   * 索引转列字母
   */
  private indexToColumn(index: number): string {
    let column = "";
    while (index > 0) {
      const remainder = (index - 1) % 26;
      column = String.fromCharCode(65 + remainder) + column;
      index = Math.floor((index - 1) / 26);
    }
    return column || "A";
  }

  /**
   * 生成唯一 ID
   */
  private generateId(): string {
    return `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
  }

  // ========== v2.7 硬约束: Replan 能力 ==========

  /**
   * 失败后重新规划
   *
   * 这不是简单的"重试"，而是根据失败原因生成新的执行策略
   */
  replan(
    originalPlan: ExecutionPlan,
    failedStep: PlanStep,
    failureReason: string,
    currentState: ReplanContext
  ): ReplanResult {
    const replanId = this.generateId();

    // 1. 分析失败原因
    const failureAnalysis = this.analyzeFailure(failedStep, failureReason);

    // 2. 决定 replan 策略
    const strategy = this.determineReplanStrategy(failureAnalysis, originalPlan, currentState);

    // 3. 生成新的步骤
    const newSteps = this.generateReplanSteps(originalPlan, failedStep, strategy, currentState);

    // 4. 构建 replan 结果
    const result: ReplanResult = {
      replanId,
      originalPlanId: originalPlan.id,
      failedStepId: failedStep.id,
      failureAnalysis,
      strategy,
      newSteps,
      estimatedAdditionalDuration: newSteps.length * 2000,
      recommendation: this.generateReplanRecommendation(strategy, failureAnalysis),
      canContinue: strategy !== "abort",
      requiresUserConfirmation:
        strategy === "alternative_approach" || strategy === "partial_rollback",
    };

    return result;
  }

  /**
   * 分析失败原因
   */
  private analyzeFailure(failedStep: PlanStep, failureReason: string): FailureAnalysis {
    // 检测失败类型
    let failureType: FailureType = "unknown";
    let rootCause = "";
    const suggestions: string[] = [];

    // #REF! - 引用错误
    if (failureReason.includes("#REF!")) {
      failureType = "reference_error";
      rootCause = "\u5F15\u7528\u7684\u5DE5\u4F5C\u8868\u6216\u8303\u56F4\u4E0D\u5B58\u5728";
      suggestions.push(
        "\u68C0\u67E5\u5F15\u7528\u7684\u5DE5\u4F5C\u8868\u662F\u5426\u5DF2\u521B\u5EFA"
      );
      suggestions.push("\u786E\u8BA4\u6570\u636E\u6E90\u8303\u56F4\u662F\u5426\u6B63\u786E");
    }
    // #VALUE! - 值类型错误
    else if (failureReason.includes("#VALUE!")) {
      failureType = "value_error";
      rootCause = "\u6570\u636E\u7C7B\u578B\u4E0D\u5339\u914D\u6216\u8BA1\u7B97\u65E0\u6548";
      suggestions.push(
        "\u68C0\u67E5\u88AB\u5F15\u7528\u7684\u5355\u5143\u683C\u662F\u5426\u5305\u542B\u6B63\u786E\u7684\u6570\u636E\u7C7B\u578B"
      );
      suggestions.push("\u8003\u8651\u4F7F\u7528 IFERROR \u5305\u88C5\u516C\u5F0F");
    }
    // #NAME? - 函数名错误
    else if (failureReason.includes("#NAME?")) {
      failureType = "name_error";
      rootCause = "\u51FD\u6570\u540D\u4E0D\u5B58\u5728\u6216\u62FC\u5199\u9519\u8BEF";
      suggestions.push("\u68C0\u67E5\u51FD\u6570\u540D\u62FC\u5199");
      suggestions.push(
        "\u786E\u8BA4\u4F7F\u7528\u7684\u51FD\u6570\u5728\u5F53\u524D Excel \u7248\u672C\u4E2D\u53EF\u7528"
      );
    }
    // 工作表不存在
    else if (failureReason.includes("\u4E0D\u5B58\u5728") || failureReason.includes("not found")) {
      failureType = "missing_dependency";
      rootCause = "\u4F9D\u8D56\u7684\u5DE5\u4F5C\u8868\u6216\u6570\u636E\u4E0D\u5B58\u5728";
      suggestions.push("\u5148\u521B\u5EFA\u4F9D\u8D56\u7684\u5DE5\u4F5C\u8868");
      suggestions.push("\u68C0\u67E5\u5DE5\u4F5C\u8868\u540D\u79F0\u62FC\u5199");
    }
    // 执行超时
    else if (failureReason.includes("timeout") || failureReason.includes("\u8D85\u65F6")) {
      failureType = "timeout";
      rootCause = "\u64CD\u4F5C\u6267\u884C\u8D85\u65F6";
      suggestions.push("\u51CF\u5C11\u64CD\u4F5C\u7684\u6570\u636E\u8303\u56F4");
      suggestions.push("\u5206\u6279\u6267\u884C");
    }
    // 通用错误
    else {
      failureType = "execution_error";
      rootCause = failureReason;
      suggestions.push("\u68C0\u67E5\u53C2\u6570\u662F\u5426\u6B63\u786E");
      suggestions.push("\u67E5\u770B\u8BE6\u7EC6\u9519\u8BEF\u65E5\u5FD7");
    }

    return {
      failureType,
      rootCause,
      failedStep: failedStep.description,
      failedAction: failedStep.action,
      suggestions,
      isRecoverable: true, // 所有已识别的类型都是可恢复的
    };
  }

  /**
   * 决定 replan 策略
   */
  private determineReplanStrategy(
    analysis: FailureAnalysis,
    originalPlan: ExecutionPlan,
    currentState: ReplanContext
  ): ReplanStrategy {
    // 如果已经 replan 过多次，放弃
    if (currentState.replanCount >= 3) {
      return "abort";
    }

    // 根据失败类型决定策略
    switch (analysis.failureType) {
      case "missing_dependency":
        // 缺少依赖 -> 尝试先创建依赖
        return "add_prerequisite";

      case "reference_error":
        // 引用错误 -> 检查并修复引用
        if (currentState.replanCount === 0) {
          return "retry_with_fix";
        } else {
          return "alternative_approach";
        }

      case "value_error":
        // 值错误 -> 添加错误处理
        return "retry_with_fix";

      case "name_error":
        // 函数名错误 -> 使用替代函数
        return "alternative_approach";

      case "timeout":
        // 超时 -> 分批执行
        return "split_step";

      case "execution_error":
        // 通用错误 -> 先简单重试
        if (currentState.replanCount === 0) {
          return "simple_retry";
        } else {
          return "partial_rollback";
        }

      default:
        return "abort";
    }
  }

  /**
   * 生成 replan 步骤
   */
  private generateReplanSteps(
    originalPlan: ExecutionPlan,
    failedStep: PlanStep,
    strategy: ReplanStrategy,
    currentState: ReplanContext
  ): PlanStep[] {
    const steps: PlanStep[] = [];
    let order = originalPlan.steps.length;

    switch (strategy) {
      case "simple_retry":
        // 简单重试 - 重新执行失败的步骤
        steps.push({
          ...failedStep,
          id: this.generateId(),
          order: order++,
          status: "pending",
          result: undefined,
        });
        break;

      case "retry_with_fix":
        // 带修复的重试
        if (failedStep.action === "excel_set_formula") {
          // 给公式添加 IFERROR 包装
          const originalFormula = failedStep.parameters.formula as string;
          const wrappedFormula = this.wrapWithIferror(originalFormula);

          steps.push({
            ...failedStep,
            id: this.generateId(),
            order: order++,
            description: `[修复重试] ${failedStep.description}`,
            parameters: {
              ...failedStep.parameters,
              formula: wrappedFormula,
            },
            status: "pending",
            result: undefined,
          });
        } else {
          // 其他操作简单重试
          steps.push({
            ...failedStep,
            id: this.generateId(),
            order: order++,
            description: `[重试] ${failedStep.description}`,
            status: "pending",
            result: undefined,
          });
        }
        break;

      case "add_prerequisite": {
        // 添加前置步骤
        const missingSheet = this.extractMissingSheet(currentState.errorDetails || "");
        if (missingSheet) {
          // 添加创建工作表的步骤
          steps.push({
            id: this.generateId(),
            order: order++,
            phase: "create_structure",
            description: `[\u8865\u5145] \u521B\u5EFA\u7F3A\u5931\u7684\u5DE5\u4F5C\u8868: ${missingSheet}`,
            action: "excel_create_sheet",
            parameters: { name: missingSheet },
            dependsOn: [],
            isWriteOperation: true,
            successCondition: { type: "sheet_exists", targetSheet: missingSheet },
            status: "pending",
          });
        }

        // 重新执行失败的步骤
        steps.push({
          ...failedStep,
          id: this.generateId(),
          order: order++,
          description: `[\u91CD\u8BD5] ${failedStep.description}`,
          dependsOn: steps.length > 0 ? [steps[steps.length - 1].id] : failedStep.dependsOn,
          status: "pending",
          result: undefined,
        });
        break;
      }

      case "split_step": {
        // 分批执行
        if (failedStep.parameters.range && typeof failedStep.parameters.range === "string") {
          const ranges = this.splitRange(failedStep.parameters.range as string);

          for (let i = 0; i < ranges.length; i++) {
            steps.push({
              ...failedStep,
              id: this.generateId(),
              order: order++,
              description: `[分批 ${i + 1}/${ranges.length}] ${failedStep.description}`,
              parameters: {
                ...failedStep.parameters,
                range: ranges[i],
              },
              dependsOn: i === 0 ? failedStep.dependsOn : [steps[steps.length - 1].id],
              status: "pending",
              result: undefined,
            });
          }
        }
        break;
      }

      case "alternative_approach": {
        // 使用替代方案
        if (failedStep.action === "excel_set_formula") {
          const originalFormula = failedStep.parameters.formula as string;
          const alternativeFormula = this.generateAlternativeFormula(originalFormula);

          if (alternativeFormula !== originalFormula) {
            steps.push({
              ...failedStep,
              id: this.generateId(),
              order: order++,
              description: `[替代方案] ${failedStep.description}`,
              parameters: {
                ...failedStep.parameters,
                formula: alternativeFormula,
              },
              status: "pending",
              result: undefined,
            });
          }
        }
        break;
      }

      case "partial_rollback":
        // 部分回滚
        steps.push({
          id: this.generateId(),
          order: order++,
          phase: "verify",
          description: `[\u56DE\u6EDA] \u6E05\u9664\u5931\u8D25\u6B65\u9AA4\u7684\u7ED3\u679C`,
          action: "excel_clear_range",
          parameters: {
            sheet: failedStep.parameters.sheet,
            range: failedStep.parameters.range,
          },
          dependsOn: [],
          isWriteOperation: true,
          successCondition: { type: "tool_success" },
          status: "pending",
        });
        break;

      case "abort":
        // 放弃 - 不添加新步骤
        break;
    }

    return steps;
  }

  /**
   * 生成 replan 建议
   */
  private generateReplanRecommendation(
    strategy: ReplanStrategy,
    analysis: FailureAnalysis
  ): string {
    const recommendations: Record<ReplanStrategy, string> = {
      simple_retry: "\u5C06\u7B80\u5355\u91CD\u8BD5\u5931\u8D25\u7684\u6B65\u9AA4",
      retry_with_fix: "\u5C06\u4F7F\u7528 IFERROR \u5305\u88C5\u516C\u5F0F\u540E\u91CD\u8BD5",
      add_prerequisite:
        "\u5C06\u5148\u521B\u5EFA\u7F3A\u5931\u7684\u4F9D\u8D56\u9879\uFF0C\u7136\u540E\u91CD\u65B0\u6267\u884C",
      split_step:
        "\u5C06\u628A\u5927\u8303\u56F4\u64CD\u4F5C\u5206\u6210\u5C0F\u6279\u6B21\u6267\u884C",
      alternative_approach:
        "\u5C06\u5C1D\u8BD5\u4F7F\u7528\u66FF\u4EE3\u7684\u51FD\u6570\u6216\u65B9\u6CD5",
      partial_rollback:
        "\u5C06\u56DE\u6EDA\u5931\u8D25\u7684\u64CD\u4F5C\uFF0C\u9700\u8981\u624B\u52A8\u786E\u8BA4\u540E\u7EED\u6B65\u9AA4",
      abort: "\u65E0\u6CD5\u81EA\u52A8\u4FEE\u590D\uFF0C\u5EFA\u8BAE\u624B\u52A8\u68C0\u67E5",
    };

    let recommendation = recommendations[strategy] || "\u9700\u8981\u624B\u52A8\u5904\u7406";
    recommendation += `\n\n**\u5931\u8D25\u539F\u56E0**: ${analysis.rootCause}`;

    if (analysis.suggestions.length > 0) {
      recommendation += "\n\n**\u5EFA\u8BAE**:\n";
      recommendation += analysis.suggestions.map((s) => `- ${s}`).join("\n");
    }

    return recommendation;
  }

  /**
   * 给公式添加 IFERROR 包装
   */
  private wrapWithIferror(formula: string): string {
    if (formula.startsWith("=")) {
      const inner = formula.substring(1);
      // 如果已经有 IFERROR，不再包装
      if (inner.toUpperCase().startsWith("IFERROR(")) {
        return formula;
      }
      return `=IFERROR(${inner},"")`;
    }
    return formula;
  }

  /**
   * 从错误信息中提取缺失的工作表名
   */
  private extractMissingSheet(errorDetails: string): string | null {
    // 匹配 "工作表 'xxx' 不存在" 或 "'xxx' not found"
    const patterns = [
      /工作表\s*['"]?([^'"]+)['"]?\s*不存在/,
      /['"]([^'"]+)['"]\s*not found/i,
      /Sheet\s*['"]?([^'"]+)['"]?\s*does not exist/i,
    ];

    for (const pattern of patterns) {
      const match = errorDetails.match(pattern);
      if (match) {
        return match[1];
      }
    }

    return null;
  }

  /**
   * 分割范围（用于分批执行）
   */
  private splitRange(range: string): string[] {
    // 解析范围 A1:A1000 -> [A1:A333, A334:A666, A667:A1000]
    const match = range.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (!match) return [range];

    const startCol = match[1];
    const startRow = parseInt(match[2]);
    const endCol = match[3];
    const endRow = parseInt(match[4]);

    const totalRows = endRow - startRow + 1;
    const batchSize = Math.ceil(totalRows / 3); // 分成 3 批

    const ranges: string[] = [];
    let currentStart = startRow;

    while (currentStart <= endRow) {
      const currentEnd = Math.min(currentStart + batchSize - 1, endRow);
      ranges.push(`${startCol}${currentStart}:${endCol}${currentEnd}`);
      currentStart = currentEnd + 1;
    }

    return ranges;
  }

  /**
   * 生成替代公式
   */
  private generateAlternativeFormula(formula: string): string {
    // XLOOKUP -> VLOOKUP (兼容性更好)
    if (formula.toUpperCase().includes("XLOOKUP")) {
      // 简化处理：提示用户考虑 VLOOKUP
      return formula; // 这里可以实现更复杂的转换
    }

    // FILTER -> 其他方法
    if (formula.toUpperCase().includes("FILTER(")) {
      return formula; // 这里可以实现更复杂的转换
    }

    return formula;
  }
}

// ========== v2.7 Replan 类型定义 ==========

/**
 * Replan 上下文
 */
export interface ReplanContext {
  replanCount: number; // 已经 replan 的次数
  completedSteps: string[]; // 已完成的步骤 ID
  errorDetails?: string; // 详细错误信息
  workbookState?: {
    // 当前工作簿状态
    sheets: string[];
    hasUnsavedChanges: boolean;
  };
}

/**
 * 失败类型
 */
export type FailureType =
  | "reference_error" // #REF!
  | "value_error" // #VALUE!
  | "name_error" // #NAME?
  | "missing_dependency" // 缺少依赖
  | "timeout" // 超时
  | "execution_error" // 通用执行错误
  | "unknown"; // 未知错误

/**
 * Replan 策略
 */
export type ReplanStrategy =
  | "simple_retry" // 简单重试
  | "retry_with_fix" // 带修复的重试
  | "add_prerequisite" // 添加前置步骤
  | "split_step" // 分批执行
  | "alternative_approach" // 使用替代方案
  | "partial_rollback" // 部分回滚
  | "abort"; // 放弃

/**
 * 失败分析结果
 */
export interface FailureAnalysis {
  failureType: FailureType;
  rootCause: string;
  failedStep: string;
  failedAction: string;
  suggestions: string[];
  isRecoverable: boolean;
}

/**
 * Replan 结果
 */
export interface ReplanResult {
  replanId: string;
  originalPlanId: string;
  failedStepId: string;
  failureAnalysis: FailureAnalysis;
  strategy: ReplanStrategy;
  newSteps: PlanStep[];
  estimatedAdditionalDuration: number;
  recommendation: string;
  canContinue: boolean;
  requiresUserConfirmation: boolean;
}

// 导出单例
export const taskPlanner = new TaskPlanner();
