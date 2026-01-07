/**
 * useAgentOrchestrator - 智能 Agent 闭环集成 Hook
 *
 * 这是 v5.0 架构的 React 集成层：
 * - 完整的执行-验证-修复闭环
 * - 实时进度反馈
 * - 经验学习与复用
 *
 * @module useAgentOrchestrator
 */

import * as React from "react";
import {
  AgentOrchestrator,
  createAgentOrchestrator,
  OrchestratorConfig,
  OrchestratorResult,
  OrchestratorEvent,
  AgentPhase,
  AgentState,
  StepResultWithVerification,
} from "../../agent/AgentOrchestrator";
import { ParseContext } from "../../agent/IntentParser";
import { ReusableExperience } from "../../agent/EpisodicMemory";

// ========== 类型定义 ==========

/**
 * Hook 配置
 */
export interface UseAgentOrchestratorOptions extends Partial<OrchestratorConfig> {
  /** 初始化回调 */
  onInitialized?: (orchestrator: AgentOrchestrator) => void;

  /** 阶段变化回调 */
  onPhaseChange?: (phase: AgentPhase) => void;

  /** 步骤完成回调 */
  onStepComplete?: (stepId: string, success: boolean, output?: string) => void;

  /** 执行完成回调 */
  onComplete?: (result: OrchestratorResult) => void;

  /** 错误回调 */
  onError?: (error: string) => void;
}

/**
 * 进度信息
 */
export interface ProgressInfo {
  /** 当前阶段 */
  phase: AgentPhase;

  /** 阶段显示名称 */
  phaseName: string;

  /** 当前步骤索引 */
  currentStep: number;

  /** 总步骤数 */
  totalSteps: number;

  /** 进度百分比 (0-100) */
  percentage: number;

  /** 当前操作描述 */
  description: string;

  /** 迭代轮数 */
  iteration: number;

  /** 重试次数 */
  retryCount: number;
}

/**
 * Hook 返回值
 */
export interface UseAgentOrchestratorReturn {
  // ===== 状态 =====
  /** 是否正在执行 */
  isExecuting: boolean;

  /** 当前阶段 */
  phase: AgentPhase;

  /** 进度信息 */
  progress: ProgressInfo | null;

  /** 步骤结果列表 */
  stepResults: StepResultWithVerification[];

  /** 最后结果 */
  lastResult: OrchestratorResult | null;

  /** 错误信息 */
  error: string | null;

  /** 学习到的经验 */
  experiences: ReusableExperience[];

  // ===== 方法 =====
  /** 执行用户请求 */
  execute: (context: ParseContext) => Promise<OrchestratorResult>;

  /** 确认并继续执行 */
  confirm: () => Promise<OrchestratorResult | null>;

  /** 取消执行 */
  cancel: () => void;

  /** 重置状态 */
  reset: () => void;

  /** 获取 Orchestrator 实例 */
  getOrchestrator: () => AgentOrchestrator | null;
}

// ========== 阶段名称映射 ==========

const PHASE_NAMES: Record<AgentPhase, string> = {
  idle: "就绪",
  sensing: "感知中",
  parsing: "理解意图",
  compiling: "生成计划",
  confirming: "等待确认",
  executing: "执行中",
  verifying: "验证结果",
  fixing: "修复中",
  completed: "已完成",
  failed: "执行失败",
};

// ========== Hook 实现 ==========

export function useAgentOrchestrator(
  options: UseAgentOrchestratorOptions = {}
): UseAgentOrchestratorReturn {
  // ===== Refs =====
  const orchestratorRef = React.useRef<AgentOrchestrator | null>(null);
  const pendingStateRef = React.useRef<AgentState | null>(null);
  const pendingContextRef = React.useRef<ParseContext | null>(null);

  // ===== State =====
  const [isExecuting, setIsExecuting] = React.useState(false);
  const [phase, setPhase] = React.useState<AgentPhase>("idle");
  const [progress, setProgress] = React.useState<ProgressInfo | null>(null);
  const [stepResults, setStepResults] = React.useState<StepResultWithVerification[]>([]);
  const [lastResult, setLastResult] = React.useState<OrchestratorResult | null>(null);
  const [error, setError] = React.useState<string | null>(null);
  const [experiences, setExperiences] = React.useState<ReusableExperience[]>([]);

  // ===== 初始化 =====
  React.useEffect(() => {
    const orchestrator = createAgentOrchestrator({
      maxRetries: options.maxRetries ?? 3,
      maxIterations: options.maxIterations ?? 10,
      enableLearning: options.enableLearning ?? true,
      enableAutoFix: options.enableAutoFix ?? true,
      verificationTimeout: options.verificationTimeout ?? 5000,
      confirmBeforeWrite: options.confirmBeforeWrite ?? false,
    });

    // 注册事件监听
    orchestrator.on("phase:changed", (event: OrchestratorEvent) => {
      const data = event.data as { phase: AgentPhase; iteration: number };
      setPhase(data.phase);
      options.onPhaseChange?.(data.phase);

      setProgress((prev) => ({
        ...prev!,
        phase: data.phase,
        phaseName: PHASE_NAMES[data.phase],
        iteration: data.iteration,
      }));
    });

    orchestrator.on("step:start", (event: OrchestratorEvent) => {
      const data = event.data as {
        stepId: string;
        action: string;
        description: string;
        index: number;
        total: number;
      };

      setProgress((prev) => ({
        ...prev!,
        currentStep: data.index + 1,
        totalSteps: data.total,
        percentage: Math.round(((data.index + 1) / data.total) * 100),
        description: data.description || data.action,
      }));
    });

    orchestrator.on("step:complete", (event: OrchestratorEvent) => {
      const data = event.data as { stepId: string; success: boolean; output?: string };
      options.onStepComplete?.(data.stepId, data.success, data.output);
    });

    orchestrator.on("step:failed", (event: OrchestratorEvent) => {
      const data = event.data as { stepId: string; error: string };
      options.onStepComplete?.(data.stepId, false, data.error);
    });

    orchestrator.on("retry:start", (event: OrchestratorEvent) => {
      const data = event.data as { retryCount: number; reason: string };
      setProgress((prev) => ({
        ...prev!,
        retryCount: data.retryCount,
        description: `重试中 (${data.retryCount})...`,
      }));
    });

    orchestrator.on("experience:saved", (event: OrchestratorEvent) => {
      const data = event.data as { experience: ReusableExperience };
      setExperiences((prev) => [...prev, data.experience]);
    });

    orchestratorRef.current = orchestrator;
    options.onInitialized?.(orchestrator);

    console.log("[useAgentOrchestrator] Orchestrator 初始化完成");

    return () => {
      // 清理
      orchestratorRef.current = null;
    };
  }, []); // 只在挂载时初始化

  // ===== 执行方法 =====
  const execute = React.useCallback(
    async (context: ParseContext): Promise<OrchestratorResult> => {
      const orchestrator = orchestratorRef.current;
      if (!orchestrator) {
        const errorResult: OrchestratorResult = {
          success: false,
          message: "Orchestrator 未初始化",
          state: {
            phase: "failed",
            iteration: 0,
            retryCount: 0,
            stepResults: [],
            errors: [],
            startTime: Date.now(),
          },
        };
        return errorResult;
      }

      setIsExecuting(true);
      setError(null);
      setStepResults([]);
      setProgress({
        phase: "sensing",
        phaseName: PHASE_NAMES.sensing,
        currentStep: 0,
        totalSteps: 0,
        percentage: 0,
        description: "正在分析请求...",
        iteration: 1,
        retryCount: 0,
      });

      try {
        console.log("[useAgentOrchestrator] 开始执行:", context.userMessage);

        const result = await orchestrator.run(context);

        setLastResult(result);
        setStepResults(result.state.stepResults);

        if (result.needsConfirmation) {
          // 保存状态用于后续确认
          pendingStateRef.current = result.state;
          pendingContextRef.current = context;
        }

        if (!result.success) {
          setError(result.message);
          options.onError?.(result.message);
        }

        options.onComplete?.(result);
        return result;
      } catch (err) {
        const errorMsg = err instanceof Error ? err.message : String(err);
        console.error("[useAgentOrchestrator] 执行异常:", err);
        setError(errorMsg);
        options.onError?.(errorMsg);

        const errorResult: OrchestratorResult = {
          success: false,
          message: errorMsg,
          state: {
            phase: "failed",
            iteration: 0,
            retryCount: 0,
            stepResults: [],
            errors: [
              {
                phase: "executing",
                message: errorMsg,
                timestamp: Date.now(),
                recoverable: false,
              },
            ],
            startTime: Date.now(),
          },
        };
        return errorResult;
      } finally {
        setIsExecuting(false);
      }
    },
    [options]
  );

  // ===== 确认方法 =====
  const confirm = React.useCallback(async (): Promise<OrchestratorResult | null> => {
    const orchestrator = orchestratorRef.current;
    const pendingState = pendingStateRef.current;
    const pendingContext = pendingContextRef.current;

    if (!orchestrator || !pendingState || !pendingContext) {
      console.warn("[useAgentOrchestrator] 无待确认的执行");
      return null;
    }

    setIsExecuting(true);

    try {
      const result = await orchestrator.confirmAndExecute(pendingState, pendingContext);

      setLastResult(result);
      setStepResults(result.state.stepResults);

      // 清除待确认状态
      pendingStateRef.current = null;
      pendingContextRef.current = null;

      if (!result.success) {
        setError(result.message);
        options.onError?.(result.message);
      }

      options.onComplete?.(result);
      return result;
    } catch (err) {
      const errorMsg = err instanceof Error ? err.message : String(err);
      console.error("[useAgentOrchestrator] 确认执行异常:", err);
      setError(errorMsg);
      options.onError?.(errorMsg);
      return null;
    } finally {
      setIsExecuting(false);
    }
  }, [options]);

  // ===== 取消方法 =====
  const cancel = React.useCallback(() => {
    console.log("[useAgentOrchestrator] 取消执行");
    // 目前只能清除待确认状态
    pendingStateRef.current = null;
    pendingContextRef.current = null;
    setIsExecuting(false);
    setPhase("idle");
  }, []);

  // ===== 重置方法 =====
  const reset = React.useCallback(() => {
    console.log("[useAgentOrchestrator] 重置状态");
    setIsExecuting(false);
    setPhase("idle");
    setProgress(null);
    setStepResults([]);
    setLastResult(null);
    setError(null);
    pendingStateRef.current = null;
    pendingContextRef.current = null;
  }, []);

  // ===== 获取 Orchestrator =====
  const getOrchestrator = React.useCallback(() => {
    return orchestratorRef.current;
  }, []);

  return {
    // 状态
    isExecuting,
    phase,
    progress,
    stepResults,
    lastResult,
    error,
    experiences,

    // 方法
    execute,
    confirm,
    cancel,
    reset,
    getOrchestrator,
  };
}

export default useAgentOrchestrator;
