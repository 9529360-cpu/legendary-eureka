/**
 * SmartOrchestrator - 智能编排器 v4.2
 *
 * 集成 Phase 1 & 2 所有组件的统一调度器
 *
 * 核心职责：
 * 1. 意图解析 → 工具发现 → 执行计划
 * 2. 并行执行 + 流式输出
 * 3. 错误恢复 + 追踪
 * 4. 经验记忆 + 持久化
 *
 * @module agent/SmartOrchestrator
 */

import { IntentParser, ParseContext } from "./IntentParser";
import { SpecCompiler, SpecCompileContext, SpecCompileResult } from "./SpecCompiler";
import { ToolRegistry } from "./registry";
import { ToolDiscovery, IntentAtom, ToolMatch } from "./ToolDiscovery";
import { ParallelExecutor, ParallelExecutionResult } from "./ParallelExecutor";
import { StreamingAgentExecutor } from "./StreamingAgentExecutor";
import { RecoveryManager, RecoverableStep } from "./RecoveryManager";
import { AgentTracer, getTracer } from "./tracing";
import { PersistentMemory, StoredEpisode } from "./memory";
import { IntentSpec } from "./types/intent";

// ========== 类型定义 ==========

/**
 * 编排选项
 */
export interface OrchestrationOptions {
  /** 是否启用流式输出 */
  streaming?: boolean;

  /** 是否启用并行执行 */
  parallel?: boolean;

  /** 最大并发数 */
  maxConcurrency?: number;

  /** 是否启用错误恢复 */
  enableRecovery?: boolean;

  /** 是否启用追踪 */
  enableTracing?: boolean;

  /** 是否启用持久化 */
  enablePersistence?: boolean;

  /** 会话 ID */
  sessionId?: string;

  /** 取消信号 */
  signal?: AbortSignal;

  /** 进度回调 */
  onProgress?: (progress: OrchestrationProgress) => void;
}

/**
 * 编排进度
 */
export interface OrchestrationProgress {
  phase: "parsing" | "discovering" | "compiling" | "executing" | "reflecting" | "complete";
  current: number;
  total: number;
  message: string;
}

/**
 * 编排结果
 */
export interface OrchestrationResult {
  /** 是否成功 */
  success: boolean;

  /** 最终回复 */
  reply: string;

  /** 执行统计 */
  stats: {
    parseTime: number;
    discoverTime: number;
    compileTime: number;
    executeTime: number;
    totalTime: number;
    stepsExecuted: number;
    stepsSucceeded: number;
    stepsFailed: number;
    parallelism?: number;
  };

  /** 意图规格 */
  intent?: IntentSpec;

  /** 编译结果 */
  compiled?: SpecCompileResult;

  /** 执行结果 */
  execution?: ParallelExecutionResult;

  /** 发现的工具 */
  discoveredTools?: ToolMatch[];

  /** 追踪 ID */
  traceId?: string;

  /** 错误信息 */
  error?: string;
}

// ========== SmartOrchestrator 类 ==========

/**
 * 智能编排器
 */
export class SmartOrchestrator {
  private intentParser: IntentParser;
  private specCompiler: SpecCompiler;
  private toolRegistry: ToolRegistry;
  private toolDiscovery: ToolDiscovery;
  private parallelExecutor: ParallelExecutor;
  private streamingExecutor: StreamingAgentExecutor;
  private recoveryManager: RecoveryManager;
  private tracer: AgentTracer;
  private memory: PersistentMemory | null = null;

  private sessionId: string;
  private initialized: boolean = false;

  constructor(
    toolRegistry: ToolRegistry,
    options?: {
      intentParser?: IntentParser;
      specCompiler?: SpecCompiler;
      recoveryManager?: RecoveryManager;
    }
  ) {
    this.toolRegistry = toolRegistry;
    this.intentParser = options?.intentParser ?? new IntentParser();
    this.specCompiler = options?.specCompiler ?? new SpecCompiler();
    this.recoveryManager = options?.recoveryManager ?? new RecoveryManager();

    this.toolDiscovery = new ToolDiscovery(toolRegistry);
    this.parallelExecutor = new ParallelExecutor(toolRegistry, this.recoveryManager);
    this.streamingExecutor = new StreamingAgentExecutor(toolRegistry);
    this.tracer = getTracer();

    this.sessionId = `session_${Date.now()}`;
  }

  /**
   * 初始化
   */
  async initialize(enablePersistence: boolean = false): Promise<void> {
    if (this.initialized) return;

    // 初始化工具发现器
    await this.toolDiscovery.initialize();

    // 初始化持久化内存
    if (enablePersistence) {
      try {
        this.memory = new PersistentMemory();
        await this.memory.initialize();
        console.log("[SmartOrchestrator] 持久化内存已初始化");
      } catch (error) {
        console.warn("[SmartOrchestrator] 持久化内存初始化失败，将使用内存模式", error);
        this.memory = null;
      }
    }

    this.initialized = true;
    console.log("[SmartOrchestrator] 初始化完成");
  }

  /**
   * 执行编排
   */
  async orchestrate(
    userMessage: string,
    options: OrchestrationOptions = {}
  ): Promise<OrchestrationResult> {
    const {
      parallel = true,
      maxConcurrency = 5,
      enableRecovery = true,
      enableTracing = true,
      enablePersistence = false,
      sessionId,
      signal,
      onProgress,
    } = options;

    // 确保已初始化
    if (!this.initialized) {
      await this.initialize(enablePersistence);
    }

    // 使用提供的会话 ID 或默认会话
    const currentSessionId = sessionId ?? this.sessionId;

    // 开始追踪
    const rootSpan = enableTracing
      ? this.tracer.startSpan("orchestrate", {
          userMessage: userMessage.substring(0, 100),
          sessionId: currentSessionId,
        })
      : undefined;

    const startTime = Date.now();
    const stats = {
      parseTime: 0,
      discoverTime: 0,
      compileTime: 0,
      executeTime: 0,
      totalTime: 0,
      stepsExecuted: 0,
      stepsSucceeded: 0,
      stepsFailed: 0,
      parallelism: 0,
    };

    try {
      // ===== Phase 1: 意图解析 =====
      onProgress?.({
        phase: "parsing",
        current: 1,
        total: 5,
        message: "正在理解您的意图...",
      });

      const parseStart = Date.now();
      const parseSpan = enableTracing ? this.tracer.startSpan("parse_intent") : undefined;

      const parseContext: ParseContext = {
        userMessage,
        activeSheet: "Sheet1",
        workbookSummary: {
          sheetNames: [],
        },
        conversationHistory: [],
      };

      let intent: IntentSpec;
      try {
        intent = await this.intentParser.parse(parseContext);
      } catch (error) {
        // 解析失败，使用默认意图
        console.warn("[SmartOrchestrator] 意图解析失败，使用默认意图", error);
        intent = {
          intent: "respond_only",
          confidence: 0.5,
          needsClarification: true,
          clarificationQuestion: "无法理解您的意图，请提供更多信息",
          spec: {},
        } as IntentSpec;
      }

      stats.parseTime = Date.now() - parseStart;
      if (parseSpan) this.tracer.endSpan("success");

      this.tracer.log("info", `意图解析完成: ${intent.intent}`, { intent });

      // ===== Phase 2: 工具发现 =====
      onProgress?.({
        phase: "discovering",
        current: 2,
        total: 5,
        message: "正在发现合适的工具...",
      });

      const discoverStart = Date.now();
      const discoverSpan = enableTracing ? this.tracer.startSpan("discover_tools") : undefined;

      const intentAtom: IntentAtom = {
        rawText: userMessage,
        action: this.extractAction(intent),
        entity: this.extractEntity(intent),
      };

      const discoveredTools = this.toolDiscovery.discover(intentAtom, {
        limit: 10,
        minScore: 0.2,
      });

      stats.discoverTime = Date.now() - discoverStart;
      if (discoverSpan) this.tracer.endSpan("success");

      this.tracer.log("info", `发现 ${discoveredTools.length} 个相关工具`, {
        tools: discoveredTools.map((t) => t.tool.name),
      });

      // ===== Phase 3: 规格编译 =====
      onProgress?.({
        phase: "compiling",
        current: 3,
        total: 5,
        message: "正在生成执行计划...",
      });

      const compileStart = Date.now();
      const compileSpan = enableTracing ? this.tracer.startSpan("compile_spec") : undefined;

      const compileContext: SpecCompileContext = {
        activeSheet: parseContext.activeSheet,
        currentSelection: parseContext.selection?.address,
      };

      const compiled = this.specCompiler.compile(intent, compileContext);

      stats.compileTime = Date.now() - compileStart;
      if (compileSpan) this.tracer.endSpan(compiled.success ? "success" : "error");

      if (!compiled.success) {
        this.tracer.log("error", "编译失败", { error: compiled.error });

        if (rootSpan) this.tracer.endSpan("error", compiled.error || undefined);

        return {
          success: false,
          reply: `无法生成执行计划：${compiled.error || "未知错误"}`,
          stats: { ...stats, totalTime: Date.now() - startTime },
          intent,
          compiled,
          discoveredTools,
          error: "编译失败",
        };
      }

      // ===== Phase 4: 执行 =====
      onProgress?.({
        phase: "executing",
        current: 4,
        total: 5,
        message: "正在执行操作...",
      });

      const executeStart = Date.now();
      const executeSpan = enableTracing ? this.tracer.startSpan("execute") : undefined;

      // 将 PlanStep 转换为 RecoverableStep
      const steps: RecoverableStep[] = (compiled.plan?.steps || []).map((step) => ({
        id: step.id,
        action: step.action,
        parameters: step.parameters,
        dependsOn: step.dependsOn,
        critical: true,
      }));

      let execution: ParallelExecutionResult;

      if (parallel && steps.length > 1) {
        // 并行执行
        execution = await this.parallelExecutor.execute(steps, {
          maxConcurrency,
          enableRecovery,
          continueOnFailure: true,
          signal,
        });
      } else {
        // 顺序执行（使用并行执行器但限制并发为 1）
        execution = await this.parallelExecutor.execute(steps, {
          maxConcurrency: 1,
          enableRecovery,
          continueOnFailure: true,
          signal,
        });
      }

      stats.executeTime = Date.now() - executeStart;
      stats.stepsExecuted = execution.totalSteps;
      stats.stepsSucceeded = execution.successCount;
      stats.stepsFailed = execution.failedCount;
      stats.parallelism = execution.parallelism.avgConcurrent;

      if (executeSpan) this.tracer.endSpan(execution.success ? "success" : "error");

      // ===== Phase 5: 反思与记忆 =====
      onProgress?.({
        phase: "reflecting",
        current: 5,
        total: 5,
        message: "正在总结经验...",
      });

      // 保存经验到持久化内存
      if (this.memory) {
        try {
          const episode: Omit<StoredEpisode, "id" | "timestamp"> = {
            sessionId: currentSessionId,
            intent: userMessage,
            actions: steps.map((s) => s.action),
            result: execution.success ? "success" : execution.failedCount > 0 ? "partial" : "failure",
            duration: stats.executeTime,
            toolsUsed: steps.map((s) => s.action),
          };
          await this.memory.saveEpisode(episode);
        } catch (error) {
          console.warn("[SmartOrchestrator] 保存经验失败", error);
        }
      }

      // 更新工具使用统计
      for (const [stepId, result] of execution.stepResults) {
        const step = steps.find((s) => s.id === stepId);
        if (step) {
          this.toolDiscovery.updateStats(step.action, result.success, result.duration);
        }
      }

      // ===== 生成回复 =====
      const reply = this.generateReply(intent, execution, discoveredTools);

      stats.totalTime = Date.now() - startTime;

      onProgress?.({
        phase: "complete",
        current: 5,
        total: 5,
        message: "完成！",
      });

      if (rootSpan) this.tracer.endSpan(execution.success ? "success" : "error");

      return {
        success: execution.success,
        reply,
        stats,
        intent,
        compiled,
        execution,
        discoveredTools,
        traceId: rootSpan?.id,
      };
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : String(error);
      this.tracer.log("error", "编排失败", { error: errorMsg });

      if (rootSpan) this.tracer.endSpan("error", errorMsg);

      stats.totalTime = Date.now() - startTime;

      return {
        success: false,
        reply: `执行出错：${errorMsg}`,
        stats,
        error: errorMsg,
      };
    }
  }

  /**
   * 从意图提取动作
   */
  private extractAction(intent: IntentSpec): string | undefined {
    switch (intent.intent) {
      case "create_table":
        return "create";
      case "write_data":
      case "update_data":
        return "write";
      case "delete_data":
        return "delete";
      case "format_range":
      case "style_table":
      case "conditional_format":
        return "format";
      case "create_formula":
      case "batch_formula":
      case "calculate_summary":
        return "calculate";
      case "create_chart":
      case "modify_chart":
        return "chart";
      case "create_sheet":
      case "switch_sheet":
      case "organize_sheets":
        return "sheet";
      case "sort_data":
      case "filter_data":
      case "remove_duplicates":
      case "clean_data":
        return "data";
      case "query_data":
      case "lookup_value":
        return "read";
      case "analyze_data":
      case "find_pattern":
      case "statistics":
        return "analyze";
      default:
        return undefined;
    }
  }

  /**
   * 从意图提取实体
   */
  private extractEntity(intent: IntentSpec): string | undefined {
    switch (intent.intent) {
      case "create_table":
        return "table";
      case "write_data":
      case "update_data":
      case "delete_data":
        return "cell";
      case "format_range":
      case "style_table":
      case "conditional_format":
        return "format";
      case "create_formula":
      case "batch_formula":
      case "calculate_summary":
        return "formula";
      case "create_chart":
      case "modify_chart":
        return "chart";
      case "create_sheet":
      case "switch_sheet":
      case "organize_sheets":
        return "sheet";
      case "sort_data":
      case "filter_data":
      case "remove_duplicates":
      case "clean_data":
        return "data";
      case "query_data":
      case "lookup_value":
        return "range";
      case "analyze_data":
      case "find_pattern":
      case "statistics":
        return "analysis";
      default:
        return undefined;
    }
  }

  /**
   * 生成用户回复
   */
  private generateReply(
    intent: IntentSpec,
    execution: ParallelExecutionResult,
    discoveredTools: ToolMatch[]
  ): string {
    if (!execution.success && execution.failedCount === execution.totalSteps) {
      return `❌ 操作失败。共 ${execution.totalSteps} 个步骤全部失败。`;
    }

    if (execution.success) {
      const lines: string[] = ["✅ 操作完成！"];

      lines.push(`\n📊 执行统计：`);
      lines.push(`- 成功: ${execution.successCount}/${execution.totalSteps} 步`);

      if (execution.parallelism.maxConcurrent > 1) {
        lines.push(`- 最大并行: ${execution.parallelism.maxConcurrent} 步`);
        lines.push(`- 批次数: ${execution.parallelism.batches}`);
      }

      lines.push(`- 耗时: ${execution.totalDuration}ms`);

      return lines.join("\n");
    } else {
      const lines: string[] = ["⚠️ 部分操作完成"];

      lines.push(`\n📊 执行统计：`);
      lines.push(`- 成功: ${execution.successCount}/${execution.totalSteps} 步`);
      lines.push(`- 失败: ${execution.failedCount} 步`);
      if (execution.skippedCount > 0) {
        lines.push(`- 跳过: ${execution.skippedCount} 步`);
      }

      return lines.join("\n");
    }
  }

  /**
   * 获取追踪数据
   */
  getTraceData(): ReturnType<AgentTracer["export"]> {
    return this.tracer.export();
  }

  /**
   * 获取会话历史
   */
  async getSessionHistory(): Promise<StoredEpisode[]> {
    if (!this.memory) return [];
    const episodes = await this.memory.getSimilarEpisodes("", 100);
    return episodes.filter((e) => e.sessionId === this.sessionId);
  }

  /**
   * 清理资源
   */
  close(): void {
    if (this.memory) {
      this.memory.close();
      this.memory = null;
    }
    this.initialized = false;
  }
}

// ========== 工厂函数 ==========

/**
 * 创建智能编排器
 */
export async function createSmartOrchestrator(
  toolRegistry: ToolRegistry,
  options?: { enablePersistence?: boolean }
): Promise<SmartOrchestrator> {
  const orchestrator = new SmartOrchestrator(toolRegistry);
  await orchestrator.initialize(options?.enablePersistence ?? false);
  return orchestrator;
}

export default SmartOrchestrator;
