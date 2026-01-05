/**
 * ExecutionEngine - 执行引擎
 *
 * @deprecated v2.9.5 - 当前未被使用！
 *
 * 说明：
 * - AgentCore 使用内置的 ReAct 循环 + rollbackOperations() 进行执行和回滚
 * - 此模块是早期设计的计划执行引擎，暂时保留供未来参考
 * - 如需启用，需要调用 setExcelOperations() 注入 Excel 操作接口
 *
 * 原始职责：
 * 1. 按执行计划逐步执行
 * 2. 每步执行后立即验证
 * 3. 检测错误并决定是否继续/回滚
 * 4. 提供实时进度反馈
 *
 * 核心理念：
 * - 先验证再执行
 * - 执行后立即检查
 * - 发现错误立即停止
 * - 支持回滚操作
 */

import { FormulaValidator } from "./FormulaValidator";
import { ExecutionPlan, PlanStep, StepResult } from "./TaskPlanner";

// 本地定义 ExcelError 类型
interface ExcelError {
  type: "value" | "reference" | "name" | "div_zero" | "na" | "num" | "null" | "unknown";
  cell: string;
  value: string;
  formula?: string;
  suggestion?: string;
}

// ========== 类型定义 ==========

/**
 * 执行回调
 */
export interface ExecutionCallbacks {
  onStepStart?: (step: PlanStep, plan: ExecutionPlan) => void;
  onStepComplete?: (step: PlanStep, result: StepResult, plan: ExecutionPlan) => void;
  onStepFailed?: (step: PlanStep, error: string, plan: ExecutionPlan) => void;
  onPlanComplete?: (plan: ExecutionPlan) => void;
  onPlanFailed?: (plan: ExecutionPlan, reason: string) => void;
  onValidationError?: (errors: ExcelError[], step: PlanStep) => void;
  onRollbackStart?: (step: PlanStep) => void;
  onRollbackComplete?: (step: PlanStep) => void;
}

/**
 * 执行选项
 */
export interface ExecutionOptions {
  /** 是否在错误时停止 */
  stopOnError: boolean;
  /** 是否在错误时回滚 */
  rollbackOnError: boolean;
  /** 每步执行后验证 */
  verifyAfterEachStep: boolean;
  /** 验证错误阈值（超过多少个错误算失败） */
  errorThreshold: number;
  /** 超时时间（毫秒） */
  timeout: number;
}

/**
 * 执行结果
 */
export interface ExecutionResult {
  success: boolean;
  plan: ExecutionPlan;
  completedSteps: number;
  failedSteps: number;
  errors: Array<{ step: PlanStep; error: string }>;
  validationErrors: ExcelError[];
  duration: number;
  rolledBack: boolean;
}

/**
 * 回滚记录
 */
interface RollbackRecord {
  stepId: string;
  action: string;
  parameters: Record<string, unknown>;
  undoAction: string;
  undoParameters: Record<string, unknown>;
}

// ========== 执行引擎 ==========

export class ExecutionEngine {
  private formulaValidator: FormulaValidator;
  private rollbackStack: RollbackRecord[] = [];
  private defaultOptions: ExecutionOptions = {
    stopOnError: true,
    rollbackOnError: true,
    verifyAfterEachStep: true,
    errorThreshold: 3,
    timeout: 60000,
  };

  // Excel 操作接口（需要外部注入）
  private excelOperations?: ExcelOperations;

  // v2.7 外部中断感知
  private isAborted: boolean = false;
  private abortReason: string = "";
  private networkCheckInterval?: ReturnType<typeof setInterval>;
  private lastWorkbookState?: WorkbookSnapshot;
  private interruptionHandlers: InterruptionHandler[] = [];

  constructor() {
    this.formulaValidator = new FormulaValidator();
  }

  /**
   * 设置 Excel 操作接口
   */
  setExcelOperations(ops: ExcelOperations): void {
    this.excelOperations = ops;
  }

  /**
   * 执行计划
   */
  async executePlan(
    plan: ExecutionPlan,
    callbacks?: ExecutionCallbacks,
    options?: Partial<ExecutionOptions>
  ): Promise<ExecutionResult> {
    const startTime = Date.now();
    const opts = { ...this.defaultOptions, ...options };

    // 重置回滚栈
    this.rollbackStack = [];

    // 验证计划可行性
    if (!plan.dependencyCheck.passed) {
      return {
        success: false,
        plan,
        completedSteps: 0,
        failedSteps: 0,
        errors: [{ step: plan.steps[0], error: "依赖检查未通过，无法执行" }],
        validationErrors: [],
        duration: Date.now() - startTime,
        rolledBack: false,
      };
    }

    const errors: Array<{ step: PlanStep; error: string }> = [];
    const validationErrors: ExcelError[] = [];
    let rolledBack = false;

    // 按顺序执行步骤
    for (let i = 0; i < plan.steps.length; i++) {
      const step = plan.steps[i];

      // 检查超时
      if (Date.now() - startTime > opts.timeout) {
        errors.push({ step, error: "执行超时" });
        break;
      }

      // 通知步骤开始
      step.status = "running";
      plan.currentStep = i;
      plan.phase = "execution";
      callbacks?.onStepStart?.(step, plan);

      try {
        // 执行步骤
        const result = await this.executeStep(step);
        step.result = result;

        if (result.success) {
          step.status = "completed";
          plan.completedSteps++;
          callbacks?.onStepComplete?.(step, result, plan);

          // 执行后验证
          if (opts.verifyAfterEachStep && step.phase === "set_formulas") {
            const verifyResult = await this.verifyStep(step);

            if (verifyResult.errors.length > 0) {
              validationErrors.push(...verifyResult.errors);
              callbacks?.onValidationError?.(verifyResult.errors, step);

              // 检查是否超过错误阈值
              if (validationErrors.length >= opts.errorThreshold) {
                errors.push({ step, error: `验证错误数超过阈值 (${opts.errorThreshold})` });
                step.status = "failed";
                plan.failedSteps++;

                if (opts.rollbackOnError) {
                  rolledBack = await this.rollback(plan, callbacks);
                }

                if (opts.stopOnError) {
                  break;
                }
              }
            }
          }
        } else {
          step.status = "failed";
          plan.failedSteps++;
          errors.push({ step, error: result.error || "执行失败" });
          callbacks?.onStepFailed?.(step, result.error || "执行失败", plan);

          if (opts.rollbackOnError) {
            rolledBack = await this.rollback(plan, callbacks);
          }

          if (opts.stopOnError) {
            break;
          }
        }
      } catch (err) {
        const errorMessage = err instanceof Error ? err.message : String(err);
        step.status = "failed";
        plan.failedSteps++;
        errors.push({ step, error: errorMessage });
        callbacks?.onStepFailed?.(step, errorMessage, plan);

        if (opts.rollbackOnError) {
          rolledBack = await this.rollback(plan, callbacks);
        }

        if (opts.stopOnError) {
          break;
        }
      }
    }

    // 最终验证
    if (errors.length === 0 && validationErrors.length === 0) {
      plan.phase = "verification";
      const finalValidation = await this.finalVerification(plan);
      validationErrors.push(...finalValidation.errors);

      if (finalValidation.errors.length > 0) {
        callbacks?.onValidationError?.(finalValidation.errors, plan.steps[plan.steps.length - 1]);

        if (finalValidation.shouldRollback && opts.rollbackOnError) {
          rolledBack = await this.rollback(plan, callbacks);
        }
      }
    }

    // 确定最终状态
    const success =
      errors.length === 0 && validationErrors.length < opts.errorThreshold && !rolledBack;

    plan.phase = success ? "completed" : "failed";

    if (success) {
      callbacks?.onPlanComplete?.(plan);
    } else {
      callbacks?.onPlanFailed?.(plan, errors.map((e) => e.error).join("; "));
    }

    return {
      success,
      plan,
      completedSteps: plan.completedSteps,
      failedSteps: plan.failedSteps,
      errors,
      validationErrors,
      duration: Date.now() - startTime,
      rolledBack,
    };
  }

  /**
   * 执行单个步骤
   */
  private async executeStep(step: PlanStep): Promise<StepResult> {
    const startTime = Date.now();

    if (!this.excelOperations) {
      return {
        success: false,
        error: "Excel 操作接口未设置",
        duration: Date.now() - startTime,
      };
    }

    try {
      let result: string | undefined;

      switch (step.action) {
        case "excel_create_sheet":
          result = await this.excelOperations.createSheet(step.parameters.name as string);
          // 记录回滚信息
          this.rollbackStack.push({
            stepId: step.id,
            action: step.action,
            parameters: step.parameters,
            undoAction: "excel_delete_sheet",
            undoParameters: { name: step.parameters.name },
          });
          break;

        case "excel_write_range": {
          // 先读取原始值用于回滚
          const originalValues = await this.excelOperations.readRange(
            step.parameters.sheet as string,
            step.parameters.range as string
          );

          result = await this.excelOperations.writeRange(
            step.parameters.sheet as string,
            step.parameters.range as string,
            step.parameters.values as unknown[][]
          );

          this.rollbackStack.push({
            stepId: step.id,
            action: step.action,
            parameters: step.parameters,
            undoAction: "excel_write_range",
            undoParameters: {
              sheet: step.parameters.sheet,
              range: step.parameters.range,
              values: originalValues,
            },
          });
          break;
        }

        case "excel_set_formula":
          result = await this.excelOperations.setFormula(
            step.parameters.sheet as string,
            step.parameters.range as string,
            step.parameters.formula as string
          );

          this.rollbackStack.push({
            stepId: step.id,
            action: step.action,
            parameters: step.parameters,
            undoAction: "excel_clear_range",
            undoParameters: {
              sheet: step.parameters.sheet,
              range: step.parameters.range,
            },
          });
          break;

        case "excel_add_data_validation":
          result = await this.excelOperations.addDataValidation(
            step.parameters.sheet as string,
            step.parameters.range as string,
            step.parameters.type as string,
            step.parameters.values as string[]
          );
          break;

        case "verify_execution":
          // 验证步骤不需要回滚
          result = "verification_requested";
          break;

        default:
          return {
            success: false,
            error: `未知操作: ${step.action}`,
            duration: Date.now() - startTime,
          };
      }

      return {
        success: true,
        output: result,
        duration: Date.now() - startTime,
      };
    } catch (err) {
      return {
        success: false,
        error: err instanceof Error ? err.message : String(err),
        duration: Date.now() - startTime,
      };
    }
  }

  /**
   * 验证步骤执行结果
   */
  private async verifyStep(step: PlanStep): Promise<{
    errors: ExcelError[];
    warnings: string[];
  }> {
    const errors: ExcelError[] = [];
    const warnings: string[] = [];

    if (!this.excelOperations) {
      return { errors, warnings };
    }

    // 对于公式设置步骤，检查是否有错误
    if (step.action === "excel_set_formula") {
      const sheet = step.parameters.sheet as string;
      const range = step.parameters.range as string;

      // 读取设置后的值
      const values = await this.excelOperations.readRange(sheet, range);

      // 检查错误值
      for (let row = 0; row < values.length; row++) {
        for (let col = 0; col < (values[row]?.length || 0); col++) {
          const value = String(values[row][col] || "");

          if (this.isErrorValue(value)) {
            const cellRef = this.getCellReference(range, row, col);
            errors.push({
              type: this.getErrorType(value),
              cell: `${sheet}!${cellRef}`,
              value,
              formula: step.parameters.formula as string,
              suggestion: this.getErrorSuggestion(value),
            });
          }
        }
      }
    }

    return { errors, warnings };
  }

  /**
   * 最终验证
   */
  private async finalVerification(plan: ExecutionPlan): Promise<{
    errors: ExcelError[];
    shouldRollback: boolean;
  }> {
    const errors: ExcelError[] = [];

    if (!this.excelOperations || !plan.dataModel) {
      return { errors, shouldRollback: false };
    }

    // 遍历所有表，检查错误
    for (const table of plan.dataModel.tables) {
      const usedRange = await this.excelOperations.getUsedRange(table.name);

      if (usedRange) {
        const values = await this.excelOperations.readRange(table.name, usedRange);

        for (let row = 0; row < values.length; row++) {
          for (let col = 0; col < (values[row]?.length || 0); col++) {
            const value = String(values[row][col] || "");

            if (this.isErrorValue(value)) {
              const cellRef = this.getCellReference(usedRange, row, col);
              errors.push({
                type: this.getErrorType(value),
                cell: `${table.name}!${cellRef}`,
                value,
              });
            }
          }
        }
      }
    }

    // 判断是否需要回滚
    const criticalErrors = errors.filter(
      (e) => e.type === "reference" || e.type === "name" || e.type === "value"
    );

    return {
      errors,
      shouldRollback: criticalErrors.length > 0,
    };
  }

  /**
   * 回滚操作
   */
  private async rollback(plan: ExecutionPlan, callbacks?: ExecutionCallbacks): Promise<boolean> {
    if (!this.excelOperations) {
      return false;
    }

    console.log(`[ExecutionEngine] 开始回滚 ${this.rollbackStack.length} 个操作...`);

    // 逆序执行回滚
    while (this.rollbackStack.length > 0) {
      const record = this.rollbackStack.pop()!;
      const step = plan.steps.find((s) => s.id === record.stepId);

      if (step) {
        callbacks?.onRollbackStart?.(step);
      }

      try {
        switch (record.undoAction) {
          case "excel_delete_sheet":
            await this.excelOperations.deleteSheet(record.undoParameters.name as string);
            break;

          case "excel_write_range":
            await this.excelOperations.writeRange(
              record.undoParameters.sheet as string,
              record.undoParameters.range as string,
              record.undoParameters.values as unknown[][]
            );
            break;

          case "excel_clear_range":
            await this.excelOperations.clearRange(
              record.undoParameters.sheet as string,
              record.undoParameters.range as string
            );
            break;
        }

        if (step) {
          step.status = "pending";
          callbacks?.onRollbackComplete?.(step);
        }
      } catch (err) {
        console.error(`[ExecutionEngine] 回滚失败: ${err}`);
        // 继续尝试其他回滚
      }
    }

    return true;
  }

  /**
   * 检查是否是错误值
   */
  private isErrorValue(value: string): boolean {
    const errorPatterns = [
      "#VALUE!",
      "#REF!",
      "#NAME?",
      "#DIV/0!",
      "#NULL!",
      "#N/A",
      "#NUM!",
      "#SPILL!",
      "#CALC!",
    ];
    return errorPatterns.some((pattern) => value.includes(pattern));
  }

  /**
   * 获取错误类型
   */
  private getErrorType(value: string): ExcelError["type"] {
    if (value.includes("#VALUE!")) return "value";
    if (value.includes("#REF!")) return "reference";
    if (value.includes("#NAME?")) return "name";
    if (value.includes("#DIV/0!")) return "div_zero";
    if (value.includes("#N/A")) return "na";
    if (value.includes("#NUM!")) return "num";
    if (value.includes("#NULL!")) return "null";
    return "unknown";
  }

  /**
   * 获取错误建议
   */
  private getErrorSuggestion(value: string): string {
    if (value.includes("#VALUE!")) return "数据类型不匹配，检查公式中引用的单元格";
    if (value.includes("#REF!")) return "引用的单元格或工作表不存在";
    if (value.includes("#NAME?")) return "函数名或命名范围不存在";
    if (value.includes("#DIV/0!")) return "除数为零，添加 IFERROR 处理";
    if (value.includes("#N/A")) return "查找函数未找到匹配值";
    if (value.includes("#NUM!")) return "数值无效";
    if (value.includes("#NULL!")) return "引用区域不正确";
    return "检查公式和数据";
  }

  /**
   * 获取单元格引用
   */
  private getCellReference(range: string, row: number, col: number): string {
    // 从范围中解析起始单元格
    const match = range.match(/^([A-Z]+)(\d+)/i);
    if (!match) return `A${row + 1}`;

    const startCol = match[1];
    const startRow = parseInt(match[2]);

    // 计算实际单元格
    const colIndex = this.columnToIndex(startCol) + col;
    const rowNum = startRow + row;

    return `${this.indexToColumn(colIndex)}${rowNum}`;
  }

  /**
   * 列字母转索引
   */
  private columnToIndex(col: string): number {
    let index = 0;
    for (let i = 0; i < col.length; i++) {
      index = index * 26 + (col.charCodeAt(i) - 64);
    }
    return index;
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

  // ========== v2.7 硬约束: 外部中断感知 ==========

  /**
   * 注册中断处理器
   */
  registerInterruptionHandler(handler: InterruptionHandler): void {
    this.interruptionHandlers.push(handler);
    // 按优先级排序
    this.interruptionHandlers.sort((a, b) => b.priority - a.priority);
  }

  /**
   * 移除中断处理器
   */
  removeInterruptionHandler(type: InterruptionType | "all"): void {
    this.interruptionHandlers = this.interruptionHandlers.filter((h) => h.type !== type);
  }

  /**
   * 外部中断 - 用户主动调用
   */
  abort(reason: string = "用户取消"): void {
    this.isAborted = true;
    this.abortReason = reason;
    console.log(`[ExecutionEngine] 执行被中断: ${reason}`);
  }

  /**
   * 重置中断状态
   */
  resetAbort(): void {
    this.isAborted = false;
    this.abortReason = "";
  }

  /**
   * 检查是否被中断
   */
  isInterrupted(): boolean {
    return this.isAborted;
  }

  /**
   * 获取中断原因
   */
  getAbortReason(): string {
    return this.abortReason;
  }

  /**
   * 开始监控外部变化
   */
  startMonitoring(intervalMs: number = 2000): void {
    this.stopMonitoring();

    this.networkCheckInterval = setInterval(async () => {
      await this.checkForInterruptions();
    }, intervalMs);

    console.log(`[ExecutionEngine] 开始监控外部变化 (间隔: ${intervalMs}ms)`);
  }

  /**
   * 停止监控
   */
  stopMonitoring(): void {
    if (this.networkCheckInterval) {
      clearInterval(this.networkCheckInterval);
      this.networkCheckInterval = undefined;
      console.log("[ExecutionEngine] 停止监控外部变化");
    }
  }

  /**
   * 检查外部中断
   */
  private async checkForInterruptions(): Promise<Interruption | null> {
    // 1. 检查网络连接
    const networkCheck = await this.checkNetworkConnection();
    if (networkCheck) {
      await this.handleInterruption(networkCheck);
      return networkCheck;
    }

    // 2. 检查工作簿是否被外部修改
    const workbookCheck = await this.checkWorkbookChanges();
    if (workbookCheck) {
      await this.handleInterruption(workbookCheck);
      return workbookCheck;
    }

    return null;
  }

  /**
   * 检查网络连接
   */
  private async checkNetworkConnection(): Promise<Interruption | null> {
    try {
      // 简单的网络检查 - 检查 navigator.onLine
      if (typeof navigator !== "undefined" && !navigator.onLine) {
        return {
          type: "network_error",
          reason: "网络连接已断开",
          timestamp: new Date(),
          recoverable: true,
          suggestedAction: "请检查网络连接后重试",
        };
      }

      // 可以在这里添加更复杂的网络检查，如 ping API 服务器

      return null;
    } catch (error) {
      return {
        type: "network_error",
        reason: `网络检查失败: ${error instanceof Error ? error.message : String(error)}`,
        timestamp: new Date(),
        recoverable: true,
        suggestedAction: "请检查网络连接",
      };
    }
  }

  /**
   * 检查工作簿是否被外部修改
   */
  private async checkWorkbookChanges(): Promise<Interruption | null> {
    if (!this.excelOperations) {
      return null;
    }

    try {
      const currentSnapshot = await this.takeWorkbookSnapshot();

      if (this.lastWorkbookState) {
        const changes = this.compareSnapshots(this.lastWorkbookState, currentSnapshot);

        if (changes.hasSignificantChanges) {
          return {
            type: "workbook_changed",
            reason: changes.description,
            timestamp: new Date(),
            recoverable: true,
            suggestedAction: "工作簿已被修改，建议重新规划任务",
            details: {
              addedSheets: changes.addedSheets,
              removedSheets: changes.removedSheets,
              modifiedSheets: changes.modifiedSheets,
            },
          };
        }
      }

      // 更新快照
      this.lastWorkbookState = currentSnapshot;

      return null;
    } catch (error) {
      console.error("[ExecutionEngine] 工作簿检查失败:", error);
      return null;
    }
  }

  /**
   * 拍摄工作簿快照
   */
  async takeWorkbookSnapshot(): Promise<WorkbookSnapshot> {
    const sheets: WorkbookSnapshot["sheets"] = [];

    if (this.excelOperations) {
      try {
        await Excel.run(async (context) => {
          const worksheets = context.workbook.worksheets;
          worksheets.load("items/name");
          await context.sync();

          for (const sheet of worksheets.items) {
            const usedRange = sheet.getUsedRangeOrNullObject();
            usedRange.load(["address", "cellCount"]);
            await context.sync();

            sheets.push({
              name: sheet.name,
              usedRange: usedRange.isNullObject ? null : usedRange.address,
              cellCount: usedRange.isNullObject ? 0 : usedRange.cellCount,
            });
          }
        });
      } catch (error) {
        console.error("[ExecutionEngine] 快照失败:", error);
      }
    }

    return {
      timestamp: Date.now(),
      sheets,
    };
  }

  /**
   * 比较两个快照
   */
  private compareSnapshots(
    oldSnapshot: WorkbookSnapshot,
    newSnapshot: WorkbookSnapshot
  ): WorkbookChangeResult {
    const oldSheetNames = new Set(oldSnapshot.sheets.map((s) => s.name));
    const newSheetNames = new Set(newSnapshot.sheets.map((s) => s.name));

    const addedSheets = [...newSheetNames].filter((n) => !oldSheetNames.has(n));
    const removedSheets = [...oldSheetNames].filter((n) => !newSheetNames.has(n));
    const modifiedSheets: string[] = [];

    // 检查现有工作表是否被修改
    for (const newSheet of newSnapshot.sheets) {
      const oldSheet = oldSnapshot.sheets.find((s) => s.name === newSheet.name);
      if (oldSheet) {
        // 检查单元格数量变化
        const cellDiff = Math.abs(newSheet.cellCount - oldSheet.cellCount);
        if (cellDiff > 10) {
          // 超过 10 个单元格变化视为显著修改
          modifiedSheets.push(newSheet.name);
        }
      }
    }

    const hasSignificantChanges =
      addedSheets.length > 0 || removedSheets.length > 0 || modifiedSheets.length > 0;

    let description = "";
    if (addedSheets.length > 0) {
      description += `新增工作表: ${addedSheets.join(", ")}; `;
    }
    if (removedSheets.length > 0) {
      description += `删除工作表: ${removedSheets.join(", ")}; `;
    }
    if (modifiedSheets.length > 0) {
      description += `修改工作表: ${modifiedSheets.join(", ")}`;
    }

    return {
      hasSignificantChanges,
      description: description || "无变化",
      addedSheets,
      removedSheets,
      modifiedSheets,
    };
  }

  /**
   * 处理中断
   */
  private async handleInterruption(interruption: Interruption): Promise<InterruptionResponse> {
    console.log(`[ExecutionEngine] 检测到中断: ${interruption.type} - ${interruption.reason}`);

    // 查找匹配的处理器
    const handlers = this.interruptionHandlers.filter(
      (h) => h.type === interruption.type || h.type === "all"
    );

    for (const handler of handlers) {
      try {
        const response = await handler.handler(interruption);

        if (response.action === "abort") {
          this.abort(interruption.reason);
        }

        return response;
      } catch (error) {
        console.error(`[ExecutionEngine] 中断处理器失败:`, error);
      }
    }

    // 默认响应
    return {
      action: interruption.recoverable ? "pause" : "abort",
      reason: interruption.reason,
    };
  }

  /**
   * 执行前检查 - 确保可以安全执行
   */
  async preExecutionCheck(): Promise<PreExecutionCheckResult> {
    const issues: string[] = [];
    const warnings: string[] = [];

    // 1. 检查是否被中断
    if (this.isAborted) {
      issues.push(`执行已被中断: ${this.abortReason}`);
    }

    // 2. 检查 Excel 操作接口
    if (!this.excelOperations) {
      issues.push("Excel 操作接口未设置");
    }

    // 3. 检查网络连接
    if (typeof navigator !== "undefined" && !navigator.onLine) {
      warnings.push("当前处于离线状态，可能影响某些功能");
    }

    // 4. 拍摄初始快照
    this.lastWorkbookState = await this.takeWorkbookSnapshot();

    return {
      canExecute: issues.length === 0,
      issues,
      warnings,
      workbookSnapshot: this.lastWorkbookState,
    };
  }
}

/**
 * 工作簿变化结果
 */
interface WorkbookChangeResult {
  hasSignificantChanges: boolean;
  description: string;
  addedSheets: string[];
  removedSheets: string[];
  modifiedSheets: string[];
}

/**
 * 执行前检查结果
 */
export interface PreExecutionCheckResult {
  canExecute: boolean;
  issues: string[];
  warnings: string[];
  workbookSnapshot?: WorkbookSnapshot;
}

// ========== Excel 操作接口 ==========

/**
 * Excel 操作接口（由 ExcelAdapter 实现）
 */
export interface ExcelOperations {
  createSheet(name: string): Promise<string>;
  deleteSheet(name: string): Promise<string>;
  writeRange(sheet: string, range: string, values: unknown[][]): Promise<string>;
  readRange(sheet: string, range: string): Promise<unknown[][]>;
  setFormula(sheet: string, range: string, formula: string): Promise<string>;
  clearRange(sheet: string, range: string): Promise<string>;
  addDataValidation(sheet: string, range: string, type: string, values: string[]): Promise<string>;
  getUsedRange(sheet: string): Promise<string | null>;
}

// 导出单例
export const executionEngine = new ExecutionEngine();

// ========== v2.7 外部中断感知类型 ==========

/**
 * 工作簿快照 - 用于检测用户手动修改
 */
export interface WorkbookSnapshot {
  timestamp: number;
  sheets: Array<{
    name: string;
    usedRange: string | null;
    cellCount: number;
  }>;
  checksum?: string;
}

/**
 * 中断类型
 */
export type InterruptionType =
  | "user_abort" // 用户主动中断
  | "network_error" // 网络错误
  | "workbook_changed" // 工作簿被外部修改
  | "excel_error" // Excel 内部错误
  | "timeout" // 超时
  | "resource_limit"; // 资源限制

/**
 * 中断信息
 */
export interface Interruption {
  type: InterruptionType;
  reason: string;
  timestamp: Date;
  recoverable: boolean;
  suggestedAction: string;
  details?: Record<string, unknown>;
}

/**
 * 中断处理器
 */
export interface InterruptionHandler {
  type: InterruptionType | "all";
  handler: (interruption: Interruption) => Promise<InterruptionResponse>;
  priority: number;
}

/**
 * 中断响应
 */
export interface InterruptionResponse {
  action: "continue" | "pause" | "rollback" | "abort";
  reason?: string;
  modifiedPlan?: ExecutionPlan;
}
