/**
 * StreamingAgentExecutor - 流式执行引擎 v4.1
 *
 * 提供流式输出能力，用户发送消息后立即看到反馈
 *
 * 核心特性：
 * 1. AsyncGenerator 流式输出
 * 2. 实时进度反馈
 * 3. 支持取消操作
 * 4. 与 RecoveryManager 集成
 *
 * @module agent/StreamingAgentExecutor
 */

import { IntentParser, ParseContext } from "./IntentParser";
import { SpecCompiler, CompileContext } from "./SpecCompiler";
import { ToolRegistry } from "./registry";
import { ExecutionPlan, PlanStep } from "./TaskPlanner";
import { ToolResult } from "./types/tool";
import { IntentSpec } from "./types/intent";
import createExcelTools from "./ExcelAdapter";
import { RecoveryManager, RecoveryAction, RecoverableStep } from "./RecoveryManager";

// ========== 流式输出类型 ==========

/**
 * 流式输出块类型
 */
export type StreamChunkType =
  | "status" // 状态更新
  | "thinking" // 思考中
  | "intent" // 意图识别
  | "plan" // 执行计划
  | "step:start" // 步骤开始
  | "step:progress" // 步骤进度
  | "step:done" // 步骤完成
  | "step:error" // 步骤错误
  | "step:recovery" // 步骤恢复
  | "message" // 文本消息
  | "complete" // 完成
  | "error" // 错误
  | "cancelled"; // 已取消

/**
 * 流式输出块
 */
export interface StreamChunk {
  /** 块类型 */
  type: StreamChunkType;

  /** 内容 */
  content: string;

  /** 额外数据 */
  data?: unknown;

  /** 时间戳 */
  timestamp: number;

  /** 进度百分比 (0-100) */
  progress?: number;
}

/**
 * 流式执行选项
 */
export interface StreamExecutionOptions {
  /** 是否启用错误恢复 */
  enableRecovery?: boolean;

  /** 取消信号 */
  signal?: AbortSignal;

  /** 进度回调 */
  onProgress?: (progress: number) => void;
}

/**
 * 流式执行结果
 */
export interface StreamExecutionResult {
  /** 是否成功 */
  success: boolean;

  /** 最终消息 */
  message: string;

  /** 执行的步骤数 */
  stepsExecuted: number;

  /** 总耗时 (ms) */
  duration: number;

  /** 是否被取消 */
  cancelled?: boolean;

  /** 恢复动作次数 */
  recoveryCount?: number;
}

// ========== StreamingAgentExecutor 类 ==========

/**
 * 流式执行引擎
 */
export class StreamingAgentExecutor {
  private intentParser: IntentParser;
  private specCompiler: SpecCompiler;
  private toolRegistry: ToolRegistry;
  private recoveryManager: RecoveryManager;

  constructor(toolRegistry?: ToolRegistry) {
    this.intentParser = new IntentParser();
    this.specCompiler = new SpecCompiler();
    this.toolRegistry = toolRegistry ?? this.createDefaultToolRegistry();
    this.recoveryManager = new RecoveryManager();
  }

  /**
   * 创建默认工具注册中心
   */
  private createDefaultToolRegistry(): ToolRegistry {
    const registry = new ToolRegistry();
    const excelTools = createExcelTools();
    registry.registerMany(excelTools);
    console.log(`[StreamingAgentExecutor] 已注册 ${excelTools.length} 个 Excel 工具`);
    return registry;
  }

  /**
   * 流式执行用户请求
   *
   * @param context 解析上下文
   * @param options 执行选项
   * @yields StreamChunk
   */
  async *executeStream(
    context: ParseContext,
    options: StreamExecutionOptions = {}
  ): AsyncGenerator<StreamChunk, StreamExecutionResult, unknown> {
    const startTime = Date.now();
    let stepsExecuted = 0;
    let recoveryCount = 0;

    const { enableRecovery = true, signal } = options;

    // 检查取消
    const checkCancelled = (): boolean => signal?.aborted ?? false;

    try {
      // ===== 阶段 1: 开始处理 =====
      yield this.createChunk("status", "正在理解您的请求...", { phase: "start" }, 5);

      if (checkCancelled()) {
        return this.createCancelledResult(startTime);
      }

      // ===== 阶段 2: 意图解析 =====
      yield this.createChunk("thinking", "分析用户意图...", undefined, 10);

      const intentSpec = await this.intentParser.parse(context);

      yield this.createChunk(
        "intent",
        `识别意图: ${this.getIntentDescription(intentSpec.intent)}`,
        { intent: intentSpec.intent, confidence: intentSpec.confidence },
        25
      );

      // 如果需要澄清
      if (intentSpec.needsClarification) {
        yield this.createChunk(
          "message",
          intentSpec.clarificationQuestion || "请提供更多信息",
          { needsClarification: true },
          100
        );

        return {
          success: true,
          message: intentSpec.clarificationQuestion || "请提供更多信息",
          stepsExecuted: 0,
          duration: Date.now() - startTime,
        };
      }

      if (checkCancelled()) {
        return this.createCancelledResult(startTime);
      }

      // ===== 阶段 3: 规格编译 =====
      yield this.createChunk("thinking", "规划执行步骤...", undefined, 30);

      const compileContext: CompileContext = {
        currentSelection: context.selection?.address,
        activeSheet: context.activeSheet,
      };

      const compileResult = this.specCompiler.compile(intentSpec, compileContext);

      if (!compileResult.success || !compileResult.plan) {
        yield this.createChunk(
          "error",
          compileResult.error || "无法生成执行计划",
          undefined,
          100
        );

        return {
          success: false,
          message: compileResult.error || "无法生成执行计划",
          stepsExecuted: 0,
          duration: Date.now() - startTime,
        };
      }

      const plan = compileResult.plan;
      const totalSteps = plan.steps.length;

      yield this.createChunk(
        "plan",
        `规划完成: ${totalSteps} 个步骤`,
        {
          stepCount: totalSteps,
          steps: plan.steps.map((s) => ({
            id: s.id,
            action: s.action,
            description: s.description,
          })),
        },
        35
      );

      if (checkCancelled()) {
        return this.createCancelledResult(startTime);
      }

      // ===== 阶段 4: 执行步骤 =====
      let lastOutput = "";
      const baseProgress = 35;
      const stepProgressRange = 60; // 35% - 95%

      for (let i = 0; i < plan.steps.length; i++) {
        if (checkCancelled()) {
          return this.createCancelledResult(startTime, stepsExecuted);
        }

        const step = plan.steps[i];
        const stepProgress =
          baseProgress + Math.floor((i / totalSteps) * stepProgressRange);

        yield this.createChunk(
          "step:start",
          step.description || step.action,
          {
            stepIndex: i,
            totalSteps,
            stepId: step.id,
            action: step.action,
          },
          stepProgress
        );

        try {
          const result = await this.executeStep(step, lastOutput);

          if (result.success) {
            lastOutput = result.output;
            stepsExecuted++;

            yield this.createChunk(
              "step:done",
              `✓ ${step.description || step.action}`,
              {
                stepIndex: i,
                totalSteps,
                output: result.output,
              },
              stepProgress + Math.floor(stepProgressRange / totalSteps)
            );
          } else {
            // 尝试错误恢复
            if (enableRecovery) {
              const recovery = await this.attemptRecovery(step, result.error || "未知错误");

              if (recovery) {
                recoveryCount++;
                yield this.createChunk(
                  "step:recovery",
                  `⟳ ${recovery.description}`,
                  { recoveryType: recovery.type, originalError: result.error },
                  stepProgress
                );

                if (recovery.type === "retry") {
                  // 重试
                  await this.delay(recovery.delay || 500);
                  const retryResult = await this.executeStep(step, lastOutput);
                  if (retryResult.success) {
                    lastOutput = retryResult.output;
                    stepsExecuted++;
                    yield this.createChunk(
                      "step:done",
                      `✓ ${step.description || step.action} (重试成功)`,
                      { stepIndex: i, output: retryResult.output },
                      stepProgress + Math.floor(stepProgressRange / totalSteps)
                    );
                    continue;
                  }
                } else if (recovery.type === "substitute" && recovery.alternativeStep) {
                  // 替代方案
                  const altResult = await this.executeStep(recovery.alternativeStep, lastOutput);
                  if (altResult.success) {
                    lastOutput = altResult.output;
                    stepsExecuted++;
                    yield this.createChunk(
                      "step:done",
                      `✓ ${recovery.alternativeStep.description || "替代步骤"} (替代成功)`,
                      { stepIndex: i, output: altResult.output, isSubstitute: true },
                      stepProgress + Math.floor(stepProgressRange / totalSteps)
                    );
                    continue;
                  }
                } else if (recovery.type === "skip") {
                  // 跳过
                  yield this.createChunk(
                    "step:done",
                    `⊘ ${step.description || step.action} (已跳过: ${recovery.reason})`,
                    { stepIndex: i, skipped: true },
                    stepProgress + Math.floor(stepProgressRange / totalSteps)
                  );
                  continue;
                }
              }
            }

            // 恢复失败，报告错误
            yield this.createChunk(
              "step:error",
              `✗ ${step.description || step.action}: ${result.error}`,
              { stepIndex: i, error: result.error },
              stepProgress
            );

            // 写操作失败则停止
            if (step.isWriteOperation) {
              return {
                success: false,
                message: `操作失败: ${result.error}`,
                stepsExecuted,
                duration: Date.now() - startTime,
                recoveryCount,
              };
            }
          }
        } catch (error) {
          const errorMsg = error instanceof Error ? error.message : String(error);

          yield this.createChunk(
            "step:error",
            `✗ ${step.description || step.action}: ${errorMsg}`,
            { stepIndex: i, error: errorMsg },
            stepProgress
          );

          if (step.isWriteOperation) {
            return {
              success: false,
              message: `操作异常: ${errorMsg}`,
              stepsExecuted,
              duration: Date.now() - startTime,
              recoveryCount,
            };
          }
        }
      }

      // ===== 阶段 5: 生成回复 =====
      yield this.createChunk("thinking", "生成回复...", undefined, 95);

      const finalMessage = this.generateFinalMessage(intentSpec, stepsExecuted, lastOutput);

      yield this.createChunk(
        "complete",
        finalMessage,
        {
          success: true,
          stepsExecuted,
          duration: Date.now() - startTime,
        },
        100
      );

      return {
        success: true,
        message: finalMessage,
        stepsExecuted,
        duration: Date.now() - startTime,
        recoveryCount,
      };
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : String(error);

      yield this.createChunk("error", `执行失败: ${errorMsg}`, { error: errorMsg }, 100);

      return {
        success: false,
        message: `执行失败: ${errorMsg}`,
        stepsExecuted,
        duration: Date.now() - startTime,
        recoveryCount,
      };
    }
  }

  /**
   * 执行单个步骤
   */
  private async executeStep(
    step: PlanStep | RecoverableStep,
    previousOutput: string
  ): Promise<{ success: boolean; output: string; error?: string }> {
    const tool = this.toolRegistry.get(step.action);

    if (!tool) {
      return {
        success: false,
        output: "",
        error: `工具不存在: ${step.action}`,
      };
    }

    try {
      // 准备参数（替换动态引用）
      const params = this.resolveParameters(step.parameters || {}, previousOutput);

      // 执行工具
      const result: ToolResult = await tool.execute(params);

      return {
        success: result.success,
        output: typeof result.output === "string" ? result.output : JSON.stringify(result.output),
        error: result.error,
      };
    } catch (error) {
      return {
        success: false,
        output: "",
        error: error instanceof Error ? error.message : String(error),
      };
    }
  }

  /**
   * 解析参数中的动态引用
   */
  private resolveParameters(
    params: Record<string, unknown>,
    previousOutput: string
  ): Record<string, unknown> {
    const resolved: Record<string, unknown> = {};

    for (const [key, value] of Object.entries(params)) {
      if (typeof value === "string" && value.includes("{{previous}}")) {
        resolved[key] = value.replace("{{previous}}", previousOutput);
      } else {
        resolved[key] = value;
      }
    }

    return resolved;
  }

  /**
   * 尝试错误恢复
   */
  private async attemptRecovery(
    step: PlanStep,
    errorMessage: string
  ): Promise<RecoveryAction | null> {
    return this.recoveryManager.recover(step, new Error(errorMessage));
  }

  /**
   * 创建流式输出块
   */
  private createChunk(
    type: StreamChunkType,
    content: string,
    data?: unknown,
    progress?: number
  ): StreamChunk {
    return {
      type,
      content,
      data,
      timestamp: Date.now(),
      progress,
    };
  }

  /**
   * 创建取消结果
   */
  private createCancelledResult(startTime: number, stepsExecuted: number = 0): StreamExecutionResult {
    return {
      success: false,
      message: "操作已取消",
      stepsExecuted,
      duration: Date.now() - startTime,
      cancelled: true,
    };
  }

  /**
   * 获取意图的中文描述
   */
  private getIntentDescription(intent: string): string {
    const descriptions: Record<string, string> = {
      create_table: "创建表格",
      write_data: "写入数据",
      update_data: "更新数据",
      delete_data: "删除数据",
      format_range: "格式化",
      create_formula: "创建公式",
      batch_formula: "批量公式",
      calculate_summary: "计算汇总",
      analyze_data: "分析数据",
      create_chart: "创建图表",
      sort_data: "排序",
      filter_data: "筛选",
      query_data: "查询数据",
      create_sheet: "创建工作表",
      switch_sheet: "切换工作表",
      clarify: "需要澄清",
      respond_only: "回复",
    };

    return descriptions[intent] || intent;
  }

  /**
   * 生成最终回复消息
   */
  private generateFinalMessage(
    intent: IntentSpec,
    stepsExecuted: number,
    lastOutput: string
  ): string {
    const intentDesc = this.getIntentDescription(intent.intent);

    if (intent.intent === "query_data" || intent.intent === "analyze_data") {
      // 查询类意图，返回数据
      return lastOutput || "查询完成，未找到数据";
    }

    if (stepsExecuted === 0) {
      return "操作完成，无需执行步骤";
    }

    return `${intentDesc}完成，共执行 ${stepsExecuted} 个步骤`;
  }

  /**
   * 延迟
   */
  private delay(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }
}

// ========== 工厂函数 ==========

/**
 * 创建流式执行器
 */
export function createStreamingExecutor(toolRegistry?: ToolRegistry): StreamingAgentExecutor {
  return new StreamingAgentExecutor(toolRegistry);
}

export default StreamingAgentExecutor;
