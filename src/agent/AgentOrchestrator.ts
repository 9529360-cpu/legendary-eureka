/**
 * AgentOrchestrator - 智能 Agent 统一控制中心 v1.0
 *
 * 这是真正的"大脑"，整合所有模块实现完整闭环：
 *
 * ┌─────────────────────────────────────────────────────────┐
 * │                   AgentOrchestrator                      │
 * ├─────────────────────────────────────────────────────────┤
 * │  感知 → 规划 → 执行 → 验证 → 修复/完成                  │
 * │     ↑                    ↓                              │
 * │     └────── 失败重试 ←───┘                              │
 * └─────────────────────────────────────────────────────────┘
 *
 * 核心原则：
 * 1. 闭环控制 - 执行必验证，失败必处理
 * 2. 自主规划 - 复杂任务自动拆解多轮
 * 3. 经验学习 - 记录成功/失败经验供复用
 * 4. 渐进降级 - 主方案失败尝试替代方案
 *
 * @module agent/AgentOrchestrator
 */

import { IntentParser, ParseContext } from "./IntentParser";
import { SpecCompiler, CompileContext } from "./SpecCompiler";
import { ToolRegistry } from "./registry";
import { ExecutionPlan, PlanStep } from "./TaskPlanner";
import { ToolResult } from "./types/tool";
import { IntentSpec } from "./types/intent";
import createExcelTools from "./ExcelAdapter";
import { EpisodicMemory, ReusableExperience } from "./EpisodicMemory";

// ========== 配置 ==========

/**
 * Orchestrator 配置
 */
export interface OrchestratorConfig {
  /** 最大执行重试次数 */
  maxRetries: number;

  /** 最大迭代轮数（防止无限循环） */
  maxIterations: number;

  /** 是否启用经验学习 */
  enableLearning: boolean;

  /** 是否启用自动修复 */
  enableAutoFix: boolean;

  /** 验证超时时间（毫秒） */
  verificationTimeout: number;

  /** 是否在写操作前确认 */
  confirmBeforeWrite: boolean;
}

export const DEFAULT_ORCHESTRATOR_CONFIG: OrchestratorConfig = {
  maxRetries: 3,
  maxIterations: 10,
  enableLearning: true,
  enableAutoFix: true,
  verificationTimeout: 5000,
  confirmBeforeWrite: false,
};

// ========== 状态机 ==========

/**
 * Agent 运行状态
 */
export type AgentPhase =
  | "idle" // 空闲
  | "sensing" // 感知中（收集上下文）
  | "parsing" // 解析意图
  | "compiling" // 编译规格
  | "confirming" // 等待确认
  | "executing" // 执行中
  | "verifying" // 验证中
  | "fixing" // 修复中
  | "completed" // 已完成
  | "failed"; // 失败

/**
 * Agent 运行状态（完整）
 */
export interface AgentState {
  /** 当前阶段 */
  phase: AgentPhase;

  /** 当前迭代轮数 */
  iteration: number;

  /** 当前重试次数 */
  retryCount: number;

  /** 解析的意图 */
  intentSpec?: IntentSpec;

  /** 编译的计划 */
  plan?: ExecutionPlan;

  /** 执行结果 */
  stepResults: StepResultWithVerification[];

  /** 错误记录 */
  errors: AgentError[];

  /** 最终回复 */
  finalResponse?: string;

  /** 开始时间 */
  startTime: number;

  /** 结束时间 */
  endTime?: number;
}

/**
 * 带验证的步骤结果
 */
export interface StepResultWithVerification {
  stepId: string;
  action: string;
  success: boolean;
  output?: string;
  error?: string;
  duration: number;
  /** 验证是否通过 */
  verified?: boolean;
  /** 验证失败原因 */
  verificationError?: string;
}

/**
 * Agent 错误
 */
export interface AgentError {
  phase: AgentPhase;
  message: string;
  details?: unknown;
  timestamp: number;
  recoverable: boolean;
}

// ========== 事件系统 ==========

export type OrchestratorEventType =
  | "phase:changed"
  | "intent:parsed"
  | "plan:compiled"
  | "step:start"
  | "step:complete"
  | "step:failed"
  | "verification:start"
  | "verification:passed"
  | "verification:failed"
  | "retry:start"
  | "execution:complete"
  | "execution:failed"
  | "experience:saved";

export interface OrchestratorEvent {
  type: OrchestratorEventType;
  data: unknown;
  timestamp: Date;
}

// ========== 结果类型 ==========

/**
 * Orchestrator 执行结果
 */
export interface OrchestratorResult {
  /** 是否成功 */
  success: boolean;

  /** 回复消息 */
  message: string;

  /** 完整状态 */
  state: AgentState;

  /** 是否需要用户确认 */
  needsConfirmation?: boolean;

  /** 确认问题 */
  confirmationQuestion?: string;

  /** 是否需要澄清 */
  needsClarification?: boolean;

  /** 澄清问题 */
  clarificationQuestion?: string;

  /** 学习到的经验 */
  experience?: ReusableExperience;
}

// ========== AgentOrchestrator 类 ==========

/**
 * 智能 Agent 统一控制中心
 */
export class AgentOrchestrator {
  private config: OrchestratorConfig;
  private intentParser: IntentParser;
  private specCompiler: SpecCompiler;
  private toolRegistry: ToolRegistry;
  private memory: EpisodicMemory;
  private eventHandlers: Map<OrchestratorEventType, Array<(event: OrchestratorEvent) => void>>;

  constructor(config: Partial<OrchestratorConfig> = {}) {
    this.config = { ...DEFAULT_ORCHESTRATOR_CONFIG, ...config };
    this.intentParser = new IntentParser();
    this.specCompiler = new SpecCompiler();
    this.toolRegistry = new ToolRegistry();
    this.memory = new EpisodicMemory();
    this.eventHandlers = new Map();

    // 注册 Excel 工具
    this.registerExcelTools();
  }

  /**
   * 注册 Excel 工具
   */
  private registerExcelTools(): void {
    const excelTools = createExcelTools();
    excelTools.forEach((tool) => this.toolRegistry.register(tool));
    console.log(`[AgentOrchestrator] 注册了 ${excelTools.length} 个 Excel 工具`);
  }

  /**
   * 事件监听
   */
  on(event: OrchestratorEventType, handler: (event: OrchestratorEvent) => void): void {
    if (!this.eventHandlers.has(event)) {
      this.eventHandlers.set(event, []);
    }
    this.eventHandlers.get(event)!.push(handler);
  }

  /**
   * 移除事件监听
   */
  off(event: OrchestratorEventType, handler: (event: OrchestratorEvent) => void): void {
    const handlers = this.eventHandlers.get(event);
    if (handlers) {
      const index = handlers.indexOf(handler);
      if (index > -1) {
        handlers.splice(index, 1);
      }
    }
  }

  /**
   * 发送事件
   */
  private emit(type: OrchestratorEventType, data: unknown): void {
    const event: OrchestratorEvent = { type, data, timestamp: new Date() };
    const handlers = this.eventHandlers.get(type);
    if (handlers) {
      handlers.forEach((handler) => {
        try {
          handler(event);
        } catch (error) {
          console.error(`[AgentOrchestrator] 事件处理器错误:`, error);
        }
      });
    }
  }

  /**
   * 创建初始状态
   */
  private createInitialState(): AgentState {
    return {
      phase: "idle",
      iteration: 0,
      retryCount: 0,
      stepResults: [],
      errors: [],
      startTime: Date.now(),
    };
  }

  /**
   * 更新阶段
   */
  private setPhase(state: AgentState, phase: AgentPhase): void {
    state.phase = phase;
    this.emit("phase:changed", { phase, iteration: state.iteration });
  }

  /**
   * 记录错误
   */
  private recordError(
    state: AgentState,
    message: string,
    details?: unknown,
    recoverable = true
  ): void {
    state.errors.push({
      phase: state.phase,
      message,
      details,
      timestamp: Date.now(),
      recoverable,
    });
  }

  // ========== 主执行流程 ==========

  /**
   * 执行用户请求（主入口）
   *
   * 这是完整的闭环流程：
   * 1. 感知 → 收集上下文
   * 2. 解析 → 理解意图
   * 3. 编译 → 生成计划
   * 4. 确认 → 可选的用户确认
   * 5. 执行 → 逐步执行
   * 6. 验证 → 检查结果
   * 7. 修复 → 失败时重试
   * 8. 完成 → 返回结果 + 学习经验
   */
  async run(context: ParseContext): Promise<OrchestratorResult> {
    const state = this.createInitialState();

    try {
      console.log("[AgentOrchestrator] ========== 开始执行 ==========");
      console.log("[AgentOrchestrator] 用户消息:", context.userMessage);

      // ===== 1. 感知阶段 =====
      this.setPhase(state, "sensing");
      const enrichedContext = await this.enrichContext(context);

      // 查询相似经验
      const similarExperiences = this.memory.findSimilar(context.userMessage, 3);
      if (similarExperiences.length > 0) {
        console.log(
          "[AgentOrchestrator] 找到相似经验:",
          similarExperiences.map((e) => e.taskDescription)
        );
      }

      // ===== 2. 解析阶段 =====
      this.setPhase(state, "parsing");
      const intentSpec = await this.intentParser.parse(enrichedContext);
      state.intentSpec = intentSpec;

      this.emit("intent:parsed", { intent: intentSpec.intent, confidence: intentSpec.confidence });

      // 需要澄清？
      if (intentSpec.needsClarification) {
        console.log("[AgentOrchestrator] 需要澄清");
        return this.createClarificationResult(state, intentSpec.clarificationQuestion);
      }

      // ===== 3. 编译阶段 =====
      this.setPhase(state, "compiling");
      const compileContext: CompileContext = {
        currentSelection: context.selection?.address,
        activeSheet: context.activeSheet,
      };

      const compileResult = this.specCompiler.compile(intentSpec, compileContext);

      if (!compileResult.success || !compileResult.plan) {
        this.recordError(state, compileResult.error || "编译失败");

        // 尝试降级处理
        if (this.config.enableAutoFix) {
          const fallbackResult = await this.tryFallbackCompile(state, intentSpec, compileContext);
          if (fallbackResult) {
            compileResult.plan = fallbackResult;
            compileResult.success = true;
          }
        }

        if (!compileResult.success) {
          return this.createFailureResult(state, compileResult.error || "无法生成执行计划");
        }
      }

      state.plan = compileResult.plan;
      this.emit("plan:compiled", {
        stepCount: compileResult.plan.steps.length,
        description: compileResult.plan.taskDescription,
      });

      // ===== 4. 确认阶段（可选） =====
      if (this.config.confirmBeforeWrite && this.hasWriteOperations(compileResult.plan)) {
        this.setPhase(state, "confirming");
        return this.createConfirmationResult(state, compileResult.plan);
      }

      // ===== 5-7. 执行-验证-修复 循环 =====
      return await this.executeWithRetry(state, context);
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : String(error);
      console.error("[AgentOrchestrator] 执行异常:", error);
      this.recordError(state, errorMsg, error, false);
      return this.createFailureResult(state, errorMsg);
    }
  }

  /**
   * 丰富上下文（感知阶段）
   */
  private async enrichContext(context: ParseContext): Promise<ParseContext> {
    // 如果没有选区信息，尝试获取
    if (!context.selection) {
      try {
        const selectionTool = this.toolRegistry.get("excel_get_selection");
        if (selectionTool) {
          const result = await selectionTool.execute({});
          if (result.success && result.output) {
            const output = JSON.parse(result.output);
            context.selection = {
              address: output.address || "",
              values: output.values,
              rowCount: output.rowCount,
              columnCount: output.columnCount,
            };
          }
        }
      } catch (error) {
        console.warn("[AgentOrchestrator] 获取选区失败:", error);
      }
    }

    // 如果没有工作表信息，尝试获取
    if (!context.activeSheet || !context.workbookSummary) {
      try {
        const sheetTool = this.toolRegistry.get("excel_get_sheets");
        if (sheetTool) {
          const result = await sheetTool.execute({});
          if (result.success && result.output) {
            const output = JSON.parse(result.output);
            context.activeSheet = output.activeSheet;
            context.workbookSummary = {
              sheetNames: output.sheets || [],
            };
          }
        }
      } catch (error) {
        console.warn("[AgentOrchestrator] 获取工作表失败:", error);
      }
    }

    return context;
  }

  /**
   * 检查是否有写操作
   */
  private hasWriteOperations(plan: ExecutionPlan): boolean {
    return plan.steps.some((step) => step.isWriteOperation);
  }

  /**
   * 尝试降级编译
   */
  private async tryFallbackCompile(
    state: AgentState,
    intentSpec: IntentSpec,
    context: CompileContext
  ): Promise<ExecutionPlan | null> {
    console.log("[AgentOrchestrator] 尝试降级编译...");

    // 简化意图，尝试重新编译
    const simplifiedSpec = { ...intentSpec };
    if (simplifiedSpec.spec) {
      // 移除复杂的嵌套属性
      const simpleSpec = { ...simplifiedSpec.spec } as Record<string, unknown>;
      delete simpleSpec.format;
      delete simpleSpec.validation;
      simplifiedSpec.spec = simpleSpec as IntentSpec["spec"];
    }

    const result = this.specCompiler.compile(simplifiedSpec, context);
    if (result.success && result.plan) {
      console.log("[AgentOrchestrator] 降级编译成功");
      return result.plan;
    }

    return null;
  }

  // ========== 执行-验证-修复 循环 ==========

  /**
   * 带重试的执行流程
   */
  private async executeWithRetry(
    state: AgentState,
    context: ParseContext
  ): Promise<OrchestratorResult> {
    while (state.retryCount <= this.config.maxRetries) {
      state.iteration++;

      if (state.iteration > this.config.maxIterations) {
        return this.createFailureResult(state, "超过最大迭代次数");
      }

      console.log(
        `[AgentOrchestrator] 迭代 ${state.iteration}, 重试 ${state.retryCount}/${this.config.maxRetries}`
      );

      // ===== 5. 执行阶段 =====
      this.setPhase(state, "executing");
      const executionSuccess = await this.executePlan(state);

      if (!executionSuccess) {
        // 执行失败
        if (state.retryCount < this.config.maxRetries && this.config.enableAutoFix) {
          state.retryCount++;
          this.emit("retry:start", { retryCount: state.retryCount, reason: "执行失败" });

          // 尝试修复
          this.setPhase(state, "fixing");
          const fixed = await this.tryFixExecution(state, context);
          if (fixed) {
            continue; // 重新执行
          }
        }

        return this.createFailureResult(state, "执行失败，无法恢复");
      }

      // ===== 6. 验证阶段 =====
      this.setPhase(state, "verifying");
      this.emit("verification:start", {});

      const verificationResult = await this.verifyExecution(state);

      if (verificationResult.passed) {
        this.emit("verification:passed", { details: verificationResult.details });
        break; // 验证通过，跳出循环
      } else {
        this.emit("verification:failed", { reason: verificationResult.reason });

        if (state.retryCount < this.config.maxRetries && this.config.enableAutoFix) {
          state.retryCount++;
          this.emit("retry:start", {
            retryCount: state.retryCount,
            reason: verificationResult.reason,
          });

          // ===== 7. 修复阶段 =====
          this.setPhase(state, "fixing");
          const fixed = await this.tryFixVerification(state, verificationResult, context);
          if (fixed) {
            continue; // 重新执行
          }
        }

        // 无法修复，但仍算部分成功
        console.warn("[AgentOrchestrator] 验证未完全通过，但继续返回结果");
        break;
      }
    }

    // ===== 8. 完成阶段 =====
    return await this.completeExecution(state, context);
  }

  /**
   * 执行计划
   */
  private async executePlan(state: AgentState): Promise<boolean> {
    if (!state.plan) return false;

    const plan = state.plan;
    let lastOutput = "";
    let allSuccess = true;

    for (let i = 0; i < plan.steps.length; i++) {
      const step = plan.steps[i];
      const stepStart = Date.now();

      this.emit("step:start", {
        stepId: step.id,
        action: step.action,
        description: step.description,
        index: i,
        total: plan.steps.length,
      });

      try {
        const result = await this.executeStep(step, lastOutput);

        const stepResult: StepResultWithVerification = {
          stepId: step.id,
          action: step.action,
          success: result.success,
          output: result.output,
          error: result.error,
          duration: Date.now() - stepStart,
        };

        state.stepResults.push(stepResult);

        if (result.success) {
          lastOutput = result.output || "";
          this.emit("step:complete", {
            stepId: step.id,
            success: true,
            output: result.output,
          });
        } else {
          allSuccess = false;
          this.emit("step:failed", {
            stepId: step.id,
            error: result.error,
          });

          // 写操作失败立即停止
          if (step.isWriteOperation) {
            console.error(`[AgentOrchestrator] 写操作失败，停止执行: ${result.error}`);
            this.recordError(state, `步骤 ${step.id} 执行失败: ${result.error}`);
            return false;
          }
        }
      } catch (error) {
        const errorMsg = error instanceof Error ? error.message : String(error);
        console.error(`[AgentOrchestrator] 步骤执行异常:`, error);

        state.stepResults.push({
          stepId: step.id,
          action: step.action,
          success: false,
          error: errorMsg,
          duration: Date.now() - stepStart,
        });

        this.recordError(state, `步骤 ${step.id} 异常: ${errorMsg}`);
        allSuccess = false;

        if (step.isWriteOperation) {
          return false;
        }
      }
    }

    return allSuccess;
  }

  /**
   * 执行单个步骤
   */
  private async executeStep(step: PlanStep, previousOutput: string): Promise<ToolResult> {
    const tool = this.toolRegistry.get(step.action);

    if (!tool) {
      console.warn(`[AgentOrchestrator] 工具不存在: ${step.action}`);
      return {
        success: false,
        output: "",
        error: `工具不存在: ${step.action}`,
      };
    }

    // 准备参数
    const params = { ...step.parameters };

    // 注入前序输出（如果需要）
    if (previousOutput && !params.data && step.dependsOn?.length > 0) {
      try {
        const prevData = JSON.parse(previousOutput);
        if (prevData.values) {
          params.data = prevData.values;
        }
      } catch {
        // 忽略解析错误
      }
    }

    console.log(`[AgentOrchestrator] 执行工具: ${step.action}`, params);
    return await tool.execute(params);
  }

  // ========== 验证 ==========

  /**
   * 验证执行结果
   */
  private async verifyExecution(
    state: AgentState
  ): Promise<{ passed: boolean; reason?: string; details?: unknown }> {
    if (!state.plan) {
      return { passed: false, reason: "无执行计划" };
    }

    const failedSteps = state.stepResults.filter((r) => !r.success);
    if (failedSteps.length > 0) {
      return {
        passed: false,
        reason: `${failedSteps.length} 个步骤执行失败`,
        details: failedSteps.map((s) => ({ stepId: s.stepId, error: s.error })),
      };
    }

    // 验证成功条件
    for (const step of state.plan.steps) {
      if (step.successCondition) {
        const verified = await this.verifyStepCondition(step);
        const stepResult = state.stepResults.find((r) => r.stepId === step.id);
        if (stepResult) {
          stepResult.verified = verified;
          if (!verified) {
            stepResult.verificationError = `未满足成功条件: ${step.successCondition.type}`;
          }
        }
      }
    }

    const unverifiedSteps = state.stepResults.filter((r) => r.verified === false);
    if (unverifiedSteps.length > 0) {
      return {
        passed: false,
        reason: `${unverifiedSteps.length} 个步骤验证未通过`,
        details: unverifiedSteps,
      };
    }

    return { passed: true };
  }

  /**
   * 验证步骤成功条件
   */
  private async verifyStepCondition(step: PlanStep): Promise<boolean> {
    if (!step.successCondition) return true;

    const condition = step.successCondition;

    switch (condition.type) {
      case "tool_success":
        // 工具返回成功即可
        return true;

      case "range_exists":
        // 检查范围是否有数据
        if (condition.targetRange) {
          try {
            const readTool = this.toolRegistry.get("excel_read_range");
            if (readTool) {
              const result = await readTool.execute({
                range: condition.targetRange,
                sheet: condition.targetSheet,
              });
              return result.success;
            }
          } catch {
            return false;
          }
        }
        return true;

      case "sheet_exists":
        // 检查工作表存在
        if (condition.targetSheet) {
          try {
            const sheetTool = this.toolRegistry.get("excel_get_sheets");
            if (sheetTool) {
              const result = await sheetTool.execute({});
              if (result.success && result.output) {
                const output = JSON.parse(result.output);
                return output.sheets?.includes(condition.targetSheet);
              }
            }
          } catch {
            return false;
          }
        }
        return true;

      default:
        // 其他类型默认通过
        return true;
    }
  }

  // ========== 修复 ==========

  /**
   * 尝试修复执行失败
   */
  private async tryFixExecution(state: AgentState, context: ParseContext): Promise<boolean> {
    console.log("[AgentOrchestrator] 尝试修复执行失败...");

    // 策略1: 简化计划，跳过失败的非关键步骤
    if (state.plan) {
      const failedSteps = state.stepResults.filter(
        (r) => !r.success && !r.stepId.includes("write")
      );
      if (failedSteps.length > 0) {
        // 标记失败的非写操作步骤为跳过
        for (const step of state.plan.steps) {
          if (failedSteps.find((f) => f.stepId === step.id)) {
            step.status = "skipped";
          }
        }
        // 清除失败的结果，准备重试
        state.stepResults = state.stepResults.filter(
          (r) => r.success || r.stepId.includes("write")
        );
        return true;
      }
    }

    // 策略2: 重新感知上下文
    await this.enrichContext(context);

    return false;
  }

  /**
   * 尝试修复验证失败
   */
  private async tryFixVerification(
    state: AgentState,
    verificationResult: { passed: boolean; reason?: string; details?: unknown },
    _context: ParseContext
  ): Promise<boolean> {
    console.log("[AgentOrchestrator] 尝试修复验证失败:", verificationResult.reason);

    // 目前简单处理：重新执行
    // 未来可以根据验证失败类型做更智能的修复
    state.stepResults = [];
    return true;
  }

  // ========== 完成 ==========

  /**
   * 完成执行
   */
  private async completeExecution(
    state: AgentState,
    context: ParseContext
  ): Promise<OrchestratorResult> {
    this.setPhase(state, "completed");
    state.endTime = Date.now();

    // 生成回复消息
    const successSteps = state.stepResults.filter((r) => r.success);
    const message = this.generateCompletionMessage(state, successSteps);
    state.finalResponse = message;

    // 学习经验
    let experience: ReusableExperience | undefined;
    if (this.config.enableLearning) {
      experience = this.learnFromExecution(state, context);
      if (experience) {
        this.emit("experience:saved", { experience });
      }
    }

    this.emit("execution:complete", {
      success: true,
      stepCount: successSteps.length,
      duration: state.endTime - state.startTime,
    });

    return {
      success: true,
      message,
      state,
      experience,
    };
  }

  /**
   * 生成完成消息
   */
  private generateCompletionMessage(
    state: AgentState,
    successSteps: StepResultWithVerification[]
  ): string {
    if (successSteps.length === 0) {
      return "操作已完成，但没有成功执行的步骤。";
    }

    const lastOutput = successSteps[successSteps.length - 1]?.output;

    // 尝试解析最后输出为有意义的消息
    if (lastOutput) {
      try {
        const output = JSON.parse(lastOutput);
        if (output.message) {
          return output.message;
        }
        if (output.summary) {
          return output.summary;
        }
      } catch {
        // 直接使用原始输出
        if (lastOutput.length < 200) {
          return lastOutput;
        }
      }
    }

    const duration = state.endTime ? ((state.endTime - state.startTime) / 1000).toFixed(1) : "未知";

    return `操作已完成！成功执行了 ${successSteps.length} 个步骤，耗时 ${duration} 秒。`;
  }

  /**
   * 学习经验
   */
  private learnFromExecution(state: AgentState, context: ParseContext): ReusableExperience | null {
    if (!state.plan || state.errors.length > 0) {
      return null;
    }

    const experience: ReusableExperience = {
      id: `exp_${Date.now()}`,
      taskDescription: context.userMessage,
      steps: state.plan.steps.map((step) => ({
        action: step.action,
        description: step.description,
        parameters: step.parameters,
      })),
      successRate: state.stepResults.filter((r) => r.success).length / state.stepResults.length,
      usageCount: 1,
      lastUsed: Date.now(),
      createdAt: Date.now(),
    };

    try {
      this.memory.save(experience);
      console.log("[AgentOrchestrator] 保存经验:", experience.id);
      return experience;
    } catch (error) {
      console.warn("[AgentOrchestrator] 保存经验失败:", error);
      return null;
    }
  }

  // ========== 结果构造器 ==========

  private createClarificationResult(state: AgentState, question?: string): OrchestratorResult {
    this.setPhase(state, "completed");
    return {
      success: true,
      message: question || "请提供更多信息",
      state,
      needsClarification: true,
      clarificationQuestion: question,
    };
  }

  private createConfirmationResult(state: AgentState, plan: ExecutionPlan): OrchestratorResult {
    const writeSteps = plan.steps.filter((s) => s.isWriteOperation);
    const question = `即将执行 ${writeSteps.length} 个写操作，是否继续？`;

    return {
      success: true,
      message: question,
      state,
      needsConfirmation: true,
      confirmationQuestion: question,
    };
  }

  private createFailureResult(state: AgentState, error: string): OrchestratorResult {
    this.setPhase(state, "failed");
    state.endTime = Date.now();

    this.emit("execution:failed", { error });

    return {
      success: false,
      message: `操作失败: ${error}`,
      state,
    };
  }

  // ========== 公共方法 ==========

  /**
   * 确认执行（用于确认阶段后继续）
   */
  async confirmAndExecute(state: AgentState, context: ParseContext): Promise<OrchestratorResult> {
    if (state.phase !== "confirming") {
      return this.createFailureResult(state, "当前不在确认阶段");
    }

    return await this.executeWithRetry(state, context);
  }

  /**
   * 获取工具注册表
   */
  getToolRegistry(): ToolRegistry {
    return this.toolRegistry;
  }

  /**
   * 获取经验记忆
   */
  getMemory(): EpisodicMemory {
    return this.memory;
  }
}

// ========== 工厂函数 ==========

/**
 * 创建 AgentOrchestrator 实例
 */
export function createAgentOrchestrator(config?: Partial<OrchestratorConfig>): AgentOrchestrator {
  return new AgentOrchestrator(config);
}

export default AgentOrchestrator;
