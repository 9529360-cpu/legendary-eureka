/**
 * PlanValidator - 执行计划验证器 v2.9.7
 *
 * 核心原则：验证"会不会必然失败"，不是"能不能执行"
 *
 * 这是 Agent 闭环的第一环：
 * ┌─────────────────────────────────────────────────────┐
 * │  THINK ──────→ EXECUTE ──────→ OBSERVE             │
 * │    │              │              │                  │
 * │    ▼              ▼              ▼                  │
 * │ [计划验证]     数据校验       智能回滚              │
 * │ (拦截必然失败)  (检测已发生错误) (不弄脏Excel)       │
 * └─────────────────────────────────────────────────────┘
 *
 * 5条核心规则：
 * 1. 依赖完整性 - 计划顺序不满足依赖关系
 * 2. 引用存在性 - 引用的表/列还未创建
 * 3. 角色违规 - transaction表写死值、summary表手填
 * 4. 批量行为缺失 - 只写D2但数据行>1
 * 5. 高风险操作未声明 - 覆盖整表、删除sheet
 */

import { ExecutionPlan } from "./TaskPlanner";

// ========== 类型定义 ==========

/**
 * 计划验证结果
 */
export interface PlanValidationResult {
  passed: boolean;
  canProceed: boolean; // 即使有警告，是否可以继续
  errors: PlanValidationError[];
  warnings: PlanValidationWarning[];
  suggestions: string[];
}

/**
 * 验证错误（阻止执行）
 */
export interface PlanValidationError {
  ruleId: string;
  ruleName: string;
  severity: "block";
  message: string;
  details: string[];
  affectedSteps: string[]; // step ids
  suggestedFix?: string;
}

/**
 * 验证警告（允许继续但需注意）
 */
export interface PlanValidationWarning {
  ruleId: string;
  ruleName: string;
  severity: "warn";
  message: string;
  details: string[];
  affectedSteps: string[];
}

/**
 * 工作簿上下文（用于验证引用存在性）
 */
export interface WorkbookContext {
  sheets: string[];
  // 每个 sheet 的列信息
  sheetColumns: Map<string, string[]>;
  // 每个 sheet 的行数
  sheetRowCounts: Map<string, number>;
}

/**
 * 计划验证规则
 */
export interface PlanValidationRule {
  id: string;
  name: string;
  description: string;
  severity: "block" | "warn";
  enabled: boolean;
  check: (
    plan: ExecutionPlan,
    context?: WorkbookContext
  ) => PlanValidationError | PlanValidationWarning | null;
}

// ========== 计划验证器 ==========

export class PlanValidator {
  private rules: PlanValidationRule[] = [];

  constructor() {
    this.registerDefaultRules();
  }

  /**
   * 注册默认的5条核心规则
   */
  private registerDefaultRules(): void {
    // 规则1: 依赖完整性
    this.rules.push({
      id: "dependency_order",
      name: "依赖完整性",
      description: "计划顺序必须满足依赖关系",
      severity: "block",
      enabled: true,
      check: (plan: ExecutionPlan) => {
        const stepMap = new Map<string, number>();
        plan.steps.forEach((step, index) => {
          stepMap.set(step.id, index);
        });

        for (const step of plan.steps) {
          const stepIndex = stepMap.get(step.id)!;

          for (const depId of step.dependsOn) {
            const depIndex = stepMap.get(depId);
            if (depIndex === undefined) {
              return {
                ruleId: "dependency_order",
                ruleName: "依赖完整性",
                severity: "block",
                message: `步骤 "${step.description}" 依赖不存在的步骤 ${depId}`,
                details: [`步骤 ${step.id} 声明依赖 ${depId}，但该步骤不在计划中`],
                affectedSteps: [step.id],
                suggestedFix: "检查计划步骤是否完整，或移除无效依赖",
              };
            }

            if (depIndex >= stepIndex) {
              return {
                ruleId: "dependency_order",
                ruleName: "依赖完整性",
                severity: "block",
                message: `依赖顺序错误: "${step.description}" 在其依赖项之前`,
                details: [
                  `步骤 ${step.id} (顺序 ${stepIndex}) 依赖步骤 ${depId} (顺序 ${depIndex})`,
                  "被依赖的步骤必须先执行",
                ],
                affectedSteps: [step.id, depId],
                suggestedFix: "调整步骤顺序，确保依赖项先执行",
              };
            }
          }
        }

        // 检查隐式依赖：交易表引用产品表，汇总表引用交易表
        const sheetCreationOrder: string[] = [];
        const sheetReferenceOrder: { sheet: string; references: string[] }[] = [];

        for (const step of plan.steps) {
          const params = step.parameters;
          const sheet = (params.sheet as string) || "";

          // 检测创建表的步骤
          if (step.phase === "create_structure" && sheet) {
            sheetCreationOrder.push(sheet);
          }

          // 检测公式中的跨表引用
          if (step.phase === "set_formulas") {
            const formula = (params.formula as string) || "";
            const references = this.extractSheetReferences(formula);
            if (sheet && references.length > 0) {
              sheetReferenceOrder.push({ sheet, references });
            }
          }
        }

        // 验证引用的表是否在引用前创建
        for (const { sheet, references } of sheetReferenceOrder) {
          const sheetIdx = sheetCreationOrder.indexOf(sheet);
          for (const ref of references) {
            const refIdx = sheetCreationOrder.indexOf(ref);
            if (refIdx === -1) {
              // 引用的表不在创建列表中，可能是已存在的表
              continue;
            }
            if (refIdx > sheetIdx) {
              return {
                ruleId: "dependency_order",
                ruleName: "依赖完整性",
                severity: "block",
                message: `隐式依赖错误: "${sheet}" 引用 "${ref}"，但 "${ref}" 还未创建`,
                details: [
                  `工作表 "${sheet}" 的公式引用了 "${ref}"`,
                  `但 "${ref}" 在计划中排在 "${sheet}" 之后`,
                ],
                affectedSteps: [],
                suggestedFix: `调整顺序，先创建/处理 "${ref}"，再处理 "${sheet}"`,
              };
            }
          }
        }

        return null;
      },
    });

    // 规则2: 引用存在性
    this.rules.push({
      id: "reference_exists",
      name: "引用存在性",
      description: "公式引用的表/列必须存在",
      severity: "block",
      enabled: true,
      check: (plan: ExecutionPlan, context?: WorkbookContext) => {
        if (!context) return null; // 没有上下文时跳过

        // 收集计划中将创建的表
        const plannedSheets = new Set<string>();
        for (const step of plan.steps) {
          if (step.phase === "create_structure") {
            const sheet = step.parameters.sheet as string;
            if (sheet) plannedSheets.add(sheet);
          }
        }

        // 检查公式中引用的表是否存在
        for (const step of plan.steps) {
          if (step.phase === "set_formulas") {
            const formula = (step.parameters.formula as string) || "";
            const references = this.extractSheetReferences(formula);

            for (const ref of references) {
              const existsNow = context.sheets.includes(ref);
              const willBeCreated = plannedSheets.has(ref);

              if (!existsNow && !willBeCreated) {
                return {
                  ruleId: "reference_exists",
                  ruleName: "引用存在性",
                  severity: "block",
                  message: `引用的工作表 "${ref}" 不存在`,
                  details: [
                    `公式 "${formula}" 引用了工作表 "${ref}"`,
                    `该工作表既不存在于当前工作簿，也不在创建计划中`,
                  ],
                  affectedSteps: [step.id],
                  suggestedFix: `先创建工作表 "${ref}"，或检查表名是否正确`,
                };
              }
            }
          }
        }

        return null;
      },
    });

    // 规则3: 角色违规 - Excel 语义错误
    this.rules.push({
      id: "role_violation",
      name: "角色违规",
      description: "交易表不能写死值，汇总表不能手填数值",
      severity: "block",
      enabled: true,
      check: (plan: ExecutionPlan) => {
        for (const step of plan.steps) {
          const params = step.parameters;
          const sheet = ((params.sheet as string) || "").toLowerCase();
          const action = step.action.toLowerCase();

          // 检测交易表写死值
          if (/交易|订单|销售|transaction|order/.test(sheet)) {
            // 检查是否直接写入数值到敏感列
            if (action.includes("write") || action.includes("update")) {
              const range = (params.range as string) || "";
              const values = params.values;

              // 检测敏感列: 单价、成本、金额
              if (/[D-G]/i.test(range) && values) {
                // D-G 列通常是数值列
                if (this.containsHardcodedNumbers(values)) {
                  return {
                    ruleId: "role_violation",
                    ruleName: "角色违规",
                    severity: "block",
                    message: `交易表中检测到硬编码数值`,
                    details: [
                      `步骤 "${step.description}" 向交易表写入硬编码数值`,
                      `交易表的单价、成本、金额应该用公式引用主数据表`,
                    ],
                    affectedSteps: [step.id],
                    suggestedFix: "使用 XLOOKUP 从主数据表引用单价/成本",
                  };
                }
              }
            }
          }

          // 检测汇总表手填数值
          if (/汇总|统计|summary|report|月度|年度/.test(sheet)) {
            if (action.includes("write") || action.includes("update")) {
              const values = params.values;
              if (this.containsHardcodedNumbers(values)) {
                return {
                  ruleId: "role_violation",
                  ruleName: "角色违规",
                  severity: "block",
                  message: `汇总表中检测到手填数值`,
                  details: [
                    `步骤 "${step.description}" 向汇总表写入硬编码数值`,
                    `汇总表的数值应该用 SUMIF/SUMIFS 等公式计算`,
                  ],
                  affectedSteps: [step.id],
                  suggestedFix: "使用 SUMIF/SUMIFS 公式从交易表汇总数据",
                };
              }
            }
          }
        }

        return null;
      },
    });

    // 规则4: 批量行为缺失
    this.rules.push({
      id: "batch_behavior_missing",
      name: "批量行为缺失",
      description: "单行公式设置后应填充到所有数据行",
      severity: "warn", // 警告级别，因为可能是有意为之
      enabled: true,
      check: (plan: ExecutionPlan, context?: WorkbookContext) => {
        for (let i = 0; i < plan.steps.length; i++) {
          const step = plan.steps[i];

          // 只检查单个公式设置
          if (step.action !== "excel_set_formula") continue;

          const params = step.parameters;
          const range = (params.range as string) || "";
          const sheet = params.sheet as string;

          // 检测是否只设置单行，如 D2
          const singleCellMatch = range.match(/^([A-Z]+)(\d+)$/i);
          if (!singleCellMatch) continue;

          const row = parseInt(singleCellMatch[2]);
          if (row <= 1) continue; // 表头行跳过

          // v2.9.41: 修复工具名 - fill_formula -> excel_fill_formula 或 excel_batch_formula
          const hasFollowupFill = plan.steps.slice(i + 1).some((s) => {
            return (
              (s.action === "excel_fill_formula" || s.action === "excel_batch_formula") &&
              s.parameters.sheet === sheet &&
              ((s.parameters.sourceRange as string) || "").includes(singleCellMatch[1])
            );
          });

          if (!hasFollowupFill) {
            // 检查数据行数
            const rowCount = context?.sheetRowCounts.get(sheet || "") || 0;
            if (rowCount > 2) {
              return {
                ruleId: "batch_behavior_missing",
                ruleName: "批量行为缺失",
                severity: "warn",
                message: `公式只设置了单行 ${range}，但表有 ${rowCount} 行数据`,
                details: [
                  `步骤 "${step.description}" 只设置了 ${range} 的公式`,
                  `工作表 "${sheet}" 有 ${rowCount} 行数据`,
                  `后续没有 fill_formula 步骤`,
                ],
                affectedSteps: [step.id],
              };
            }
          }
        }

        return null;
      },
    });

    // 规则5: 高风险操作未声明
    this.rules.push({
      id: "high_risk_operation",
      name: "高风险操作",
      description: "覆盖整表、删除sheet等操作需要显式确认",
      severity: "block",
      enabled: true,
      check: (plan: ExecutionPlan) => {
        for (const step of plan.steps) {
          const action = step.action.toLowerCase();
          const params = step.parameters;

          // 检测删除工作表
          if (action.includes("delete") && action.includes("sheet")) {
            return {
              ruleId: "high_risk_operation",
              ruleName: "高风险操作",
              severity: "block",
              message: `计划包含删除工作表操作`,
              details: [
                `步骤 "${step.description}" 将删除工作表`,
                "这是不可逆操作，需要用户显式确认",
              ],
              affectedSteps: [step.id],
              suggestedFix: "在执行前请求用户确认",
            };
          }

          // 检测清空整表
          if (action.includes("clear") || action.includes("delete_all")) {
            const range = params.range as string;
            if (!range || (range.includes(":") && this.isWholeSheetRange(range))) {
              return {
                ruleId: "high_risk_operation",
                ruleName: "高风险操作",
                severity: "block",
                message: `计划包含清空整表操作`,
                details: [`步骤 "${step.description}" 将清空大量数据`, "这可能导致数据丢失"],
                affectedSteps: [step.id],
                suggestedFix: "在执行前请求用户确认，或指定更精确的范围",
              };
            }
          }

          // 检测覆盖已有数据
          if (action === "excel_write_data" || action === "excel_update_cells") {
            const range = params.range as string;
            if (range && this.isLargeRange(range)) {
              return {
                ruleId: "high_risk_operation",
                ruleName: "高风险操作",
                severity: "block",
                message: `计划将覆盖大范围数据: ${range}`,
                details: [`步骤 "${step.description}" 将写入 ${range}`, "这可能覆盖已有数据"],
                affectedSteps: [step.id],
                suggestedFix: "确认目标范围没有重要数据，或先备份",
              };
            }
          }
        }

        return null;
      },
    });
  }

  /**
   * 验证执行计划
   * v2.9.42: 查询类任务跳过严格验证，避免误杀
   */
  validate(plan: ExecutionPlan, context?: WorkbookContext): PlanValidationResult {
    const errors: PlanValidationError[] = [];
    const warnings: PlanValidationWarning[] = [];
    const suggestions: string[] = [];

    // v2.9.42: 识别查询类任务 - 只有读取和回复，不应被严格验证
    const isQueryOnlyPlan = this.isQueryOnlyPlan(plan);
    if (isQueryOnlyPlan) {
      console.log("[PlanValidator] 识别为查询类计划，跳过严格验证");
      return {
        passed: true,
        canProceed: true,
        errors: [],
        warnings: [],
        suggestions: ["查询类任务，直接执行"],
      };
    }

    for (const rule of this.rules) {
      if (!rule.enabled) continue;

      const result = rule.check(plan, context);
      if (result) {
        if (result.severity === "block") {
          errors.push(result as PlanValidationError);
        } else {
          warnings.push(result as PlanValidationWarning);
        }
      }
    }

    // 生成建议
    if (errors.length === 0 && warnings.length > 0) {
      suggestions.push("计划有警告但可以继续执行，建议检查后再执行");
    }
    if (errors.length > 0) {
      suggestions.push("计划存在致命错误，必须修复后才能执行");
      for (const error of errors) {
        if (error.suggestedFix) {
          suggestions.push(`建议: ${error.suggestedFix}`);
        }
      }
    }

    return {
      passed: errors.length === 0,
      canProceed: errors.length === 0, // 有 block 级别错误就不能继续
      errors,
      warnings,
      suggestions,
    };
  }

  /**
   * 快速验证（只检查 block 级别）
   */
  quickValidate(plan: ExecutionPlan, context?: WorkbookContext): boolean {
    for (const rule of this.rules) {
      if (!rule.enabled || rule.severity !== "block") continue;

      const result = rule.check(plan, context);
      if (result && result.severity === "block") {
        return false;
      }
    }
    return true;
  }

  // ========== 工具方法 ==========

  /**
   * v2.9.42: 识别查询类计划
   * 查询类计划特征：只有 read 操作和 respond_to_user，不包含写操作
   */
  private isQueryOnlyPlan(plan: ExecutionPlan): boolean {
    if (!plan.steps || plan.steps.length === 0) return false;

    // 定义只读工具（不修改 Excel）
    const readOnlyTools = new Set([
      "excel_read_range",
      "excel_read_cell",
      "excel_get_sheets",
      "excel_get_selection",
      "excel_get_used_range",
      "excel_get_active_sheet",
      "excel_get_workbook_info",
      "respond_to_user",
    ]);

    // 所有步骤都必须是只读工具
    const allReadOnly = plan.steps.every((step) => readOnlyTools.has(step.action));

    // 至少有一个 read 操作或 respond_to_user
    const hasReadOrRespond = plan.steps.some(
      (step) => step.action.includes("read") || step.action === "respond_to_user"
    );

    return allReadOnly && hasReadOrRespond;
  }

  /**
   * 从公式中提取引用的工作表名
   */
  private extractSheetReferences(formula: string): string[] {
    const references: string[] = [];
    // 匹配 'Sheet Name'! 或 SheetName!
    const pattern = /'([^']+)'!|([A-Za-z_\u4e00-\u9fa5][A-Za-z0-9_\u4e00-\u9fa5]*)!/g;
    let match;
    while ((match = pattern.exec(formula)) !== null) {
      const sheetName = match[1] || match[2];
      if (sheetName && !references.includes(sheetName)) {
        references.push(sheetName);
      }
    }
    return references;
  }

  /**
   * 检查值中是否包含硬编码数字
   */
  private containsHardcodedNumbers(values: unknown): boolean {
    if (!values) return false;

    if (Array.isArray(values)) {
      return values.some((row) => {
        if (Array.isArray(row)) {
          return row.some((cell) => typeof cell === "number" && cell > 0);
        }
        return typeof row === "number" && row > 0;
      });
    }

    return typeof values === "number" && values > 0;
  }

  /**
   * 检查是否是整表范围
   */
  private isWholeSheetRange(range: string): boolean {
    // 如 A:Z 或 1:1000 或 A1:ZZ10000
    if (/^[A-Z]+:[A-Z]+$/i.test(range)) return true;
    if (/^\d+:\d+$/.test(range)) return true;

    const match = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
    if (match) {
      const endRow = parseInt(match[4]);
      if (endRow >= 1000) return true;
    }

    return false;
  }

  /**
   * 检查是否是大范围
   */
  private isLargeRange(range: string): boolean {
    const match = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
    if (!match) return false;

    const startRow = parseInt(match[2]);
    const endRow = parseInt(match[4]);
    const startCol = this.columnToNumber(match[1]);
    const endCol = this.columnToNumber(match[3]);

    const rows = endRow - startRow + 1;
    const cols = endCol - startCol + 1;

    return rows * cols > 500; // 超过500个单元格视为大范围
  }

  /**
   * 列字母转数字
   */
  private columnToNumber(col: string): number {
    let result = 0;
    for (let i = 0; i < col.length; i++) {
      result = result * 26 + (col.charCodeAt(i) - 64);
    }
    return result;
  }
}

// ========== 导出 ==========

export const planValidator = new PlanValidator();
