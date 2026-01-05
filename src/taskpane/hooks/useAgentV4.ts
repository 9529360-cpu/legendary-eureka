/**
 * useAgentV4 - v4.0 架构集成 Hook
 *
 * 新架构特点：
 * - LLM 只理解意图，不知道工具名
 * - SpecCompiler 纯规则编译，零 token 消耗
 * - 更清晰的职责分离
 *
 * 架构流程：
 * User Input → IntentParser (LLM) → IntentSpec
 *           → SpecCompiler (Rules) → ExecutionPlan
 *           → AgentExecutor (Execute) → Result
 *
 * @module useAgentV4
 */

import * as React from "react";
import {
  AgentExecutor,
  createAgentExecutor,
  ExecutionResult,
  ExecutorEventType,
  IntentSpec,
  // 旧架构兼容
  Agent,
  createExcelTools,
  createExcelReader,
  AgentTask,
} from "../../agent";

// ========== Types ==========

export type AgentV4Status =
  | "idle"
  | "parsing" // 意图解析中
  | "compiling" // 规格编译中
  | "executing" // 执行中
  | "completed"
  | "failed"
  | "cancelled"
  | "pending" // 等待用户确认
  | "awaiting_approval"; // 等待审批

export interface AgentV4Progress {
  phase: "parsing" | "compiling" | "executing";
  totalSteps: number;
  completedSteps: number;
  currentStep: string;
  percentage: number;
  // 兼容旧版
  currentPhase?: string;
  iteration?: number;
  maxIterations?: number;
  planSteps?: number;
}

export interface AgentV4State {
  status: AgentV4Status;
  isRunning: boolean;
  progress: AgentV4Progress | null;
  intent: IntentSpec | null;
  result: ExecutionResult | null;
  error: string | null;
  logs: AgentV4Log[];
  // 兼容旧版
  currentSteps: string[];
  lastTask: AgentTask | null;
  pendingApproval: unknown | null;
}

export interface AgentV4Log {
  timestamp: Date;
  type: "info" | "think" | "act" | "observe" | "error";
  message: string;
}

export interface UseAgentV4Options {
  verboseLogging?: boolean;
  onEvent?: (event: ExecutorEventType, data: unknown) => void;
  // 兼容旧版选项
  maxIterations?: number;
  enableMemory?: boolean;
  onStep?: (event: { type: string; text: string; timestamp: Date }) => void;
}

export interface UseAgentV4Return {
  /** 发送请求给 Agent v4 */
  send: (request: string, context?: AgentV4Context) => Promise<ExecutionResult | AgentTask>;
  /** 取消当前执行 */
  cancel: () => void;
  /** 当前状态 */
  state: AgentV4State;
  /** 重置状态 */
  reset: () => void;
  /** 获取执行器实例 */
  executor: AgentExecutor | null;
  // ========== 旧版兼容 API ==========
  /** @deprecated 使用 cancel() */
  pause: () => boolean;
  /** @deprecated 不再需要 */
  resume: () => boolean;
  /** @deprecated 使用 executor */
  agentInstance: Agent | null;
  /** @deprecated 使用 onEvent */
  setStepCallback: (callback: ((step: string) => void) | null) => void;
  /** v3.0 兼容: 批准审批 */
  approve: (approvalId: string) => void;
  /** v3.0 兼容: 拒绝审批 */
  reject: (approvalId: string, reason?: string) => void;
  /** v3.0 兼容: 审批管理器 */
  approvalManager: unknown;
}

export interface AgentV4Context {
  workbookName?: string;
  activeSheet?: string;
  selection?: string;
  // 兼容旧版
  environment?: string;
  environmentState?: Record<string, unknown>;
  conversationHistory?: Array<{ role: string; content: string }>;
}

// ========== Initial State ==========

const initialState: AgentV4State = {
  status: "idle",
  isRunning: false,
  progress: null,
  intent: null,
  result: null,
  error: null,
  logs: [],
  // 兼容旧版
  currentSteps: [],
  lastTask: null,
  pendingApproval: null,
};

// ========== Hook ==========

export function useAgentV4(options: UseAgentV4Options = {}): UseAgentV4Return {
  const { verboseLogging = true, onEvent, maxIterations = 30, enableMemory = true } = options;

  // v4 执行器实例
  const executorRef = React.useRef<AgentExecutor | null>(null);

  // 旧版 Agent 实例（兼容层）
  const legacyAgentRef = React.useRef<Agent | null>(null);

  // 步骤回调 ref（兼容旧版）
  const stepCallbackRef = React.useRef<((step: string) => void) | null>(null);

  // 状态
  const [state, setState] = React.useState<AgentV4State>(initialState);

  // 事件回调 ref
  const onEventRef = React.useRef(onEvent);
  onEventRef.current = onEvent;

  // 添加日志
  const addLog = React.useCallback((type: AgentV4Log["type"], message: string) => {
    setState((prev) => ({
      ...prev,
      logs: [...prev.logs, { timestamp: new Date(), type, message }],
    }));
  }, []);

  // 初始化执行器
  React.useEffect(() => {
    if (executorRef.current) return;

    const executor = createAgentExecutor();

    // 订阅事件
    executor.on("intent:parsed", (data) => {
      const { intent } = (data as unknown) as { intent: IntentSpec };
      if (verboseLogging) {
        console.log("[AgentV4] 意图解析完成:", intent.intent);
      }

      // 去掉可能的语义原子字段，避免 UI 展示内部原子
      const safeIntent = { ...intent } as IntentSpec;
      delete (safeIntent as any).semanticAtoms;
      delete (safeIntent as any).compressedIntent;

      setState((prev) => ({
        ...prev,
        status: "compiling",
        intent: safeIntent,
        progress: {
          phase: "compiling",
          totalSteps: 0,
          completedSteps: 0,
          currentStep: "编译执行计划",
          percentage: 30,
        },
      }));

      addLog("think", `意图识别: ${intent.intent}`);
      onEventRef.current?.("intent:parsed", data);
    });

    executor.on("plan:compiled", (data) => {
      const { plan } = (data as unknown) as { plan: { steps: unknown[] } };
      const stepCount = plan.steps.length;

      if (verboseLogging) {
        console.log("[AgentV4] 执行计划编译完成:", stepCount, "步");
      }

      setState((prev) => ({
        ...prev,
        status: "executing",
        progress: {
          phase: "executing",
          totalSteps: stepCount,
          completedSteps: 0,
          currentStep: "准备执行",
          percentage: 40,
        },
      }));

      addLog("info", `执行计划: ${stepCount} 个步骤`);
      onEventRef.current?.("plan:compiled", data);
    });

    executor.on("step:start", (data) => {
      const { step, index, total } = (data as unknown) as {
        step: { description: string };
        index: number;
        total: number;
      };

      setState((prev) => ({
        ...prev,
        progress: prev.progress
          ? {
              ...prev.progress,
              completedSteps: index,
              currentStep: step.description,
              percentage: 40 + Math.round((index / total) * 50),
            }
          : null,
      }));

      addLog("act", `执行: ${step.description}`);
      onEventRef.current?.("step:start", data);
    });

    executor.on("step:complete", (data) => {
      const { step, result, index, total } = (data as unknown) as {
        step: { description: string };
        result: { success: boolean };
        index: number;
        total: number;
      };

      if (result.success) {
        addLog("observe", `完成: ${step.description}`);
      } else {
        addLog("error", `失败: ${step.description}`);
      }

      setState((prev) => ({
        ...prev,
        progress: prev.progress
          ? {
              ...prev.progress,
              completedSteps: index + 1,
              percentage: 40 + Math.round(((index + 1) / total) * 50),
            }
          : null,
      }));

      onEventRef.current?.("step:complete", data);
    });

    executor.on("execution:complete", (data) => {
      const { result } = (data as unknown) as { result: ExecutionResult };

      if (verboseLogging) {
        console.log("[AgentV4] 执行完成:", result.success ? "成功" : "失败");
      }

      setState((prev) => ({
        ...prev,
        status: result.success ? "completed" : "failed",
        isRunning: false,
        result,
        progress: prev.progress
          ? {
              ...prev.progress,
              percentage: 100,
              currentStep: result.success ? "执行完成" : "执行失败",
            }
          : null,
      }));

      const resultError = (result as any).error ?? (result as any).errors ?? null;
      const resultErrorMessage = Array.isArray(resultError) ? resultError.join("; ") : resultError;
      addLog("info", result.success ? "✅ 任务完成" : `❌ 任务失败: ${resultErrorMessage}`);
      onEventRef.current?.("execution:complete", data);
    });

    executorRef.current = executor;
    console.log("[AgentV4] Executor initialized");
  }, [verboseLogging, addLog]);

  // 发送请求
  const send = React.useCallback(
    async (request: string, context?: AgentV4Context): Promise<ExecutionResult | AgentTask> => {
      if (!executorRef.current) {
        throw new Error("Executor not initialized");
      }

      // 重置状态
      setState({
        ...initialState,
        status: "parsing",
        isRunning: true,
        progress: {
          phase: "parsing",
          totalSteps: 0,
          completedSteps: 0,
          currentStep: "解析用户意图",
          percentage: 10,
          // 兼容旧版
          currentPhase: "解析用户意图",
          iteration: 0,
          maxIterations: maxIterations,
          planSteps: 0,
        },
        logs: [],
        currentSteps: [],
        lastTask: null,
        pendingApproval: null,
      });

      addLog("info", `收到请求: ${request}`);

      // 通知步骤回调
      if (stepCallbackRef.current) {
        stepCallbackRef.current("解析用户意图...");
      }

      try {
        const result = await executorRef.current.execute({
          userMessage: request,
          activeSheet: context?.activeSheet,
          selection: context?.selection
            ? {
                address: context.selection,
              }
            : undefined,
          workbookSummary: context?.workbookName
            ? {
                sheetNames: [],
              }
            : undefined,
        });

        // 构造兼容的 AgentTask 对象
        const compatibleTask: any = {
          id: `task_${Date.now()}`,
          request,
          status: result.success ? "completed" : "failed",
          result: result.message || (result.success ? "任务完成" : result.error || "执行失败"),
          steps: [],
          startTime: new Date(),
          endTime: new Date(),
          context: { environment: context?.environment || "excel" } as any,
        };

        setState((prev) => ({
          ...prev,
          lastTask: compatibleTask,
        }));

        return result as any;
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);

        setState((prev) => ({
          ...prev,
          status: "failed",
          isRunning: false,
          error: errorMessage,
        }));

        addLog("error", `执行错误: ${errorMessage}`);

        return {
          success: false,
          error: errorMessage,
          executedSteps: 0,
          totalSteps: 0,
          stepResults: [],
        } as any;
      }
    },
    [addLog, maxIterations]
  );

  // 取消执行
  const cancel = React.useCallback(() => {
    if (legacyAgentRef.current) {
      try {
        (legacyAgentRef.current as any)?.cancel?.();
      } catch (e) {
        // ignore
      }
    }
    setState((prev) => ({
      ...prev,
      status: "cancelled",
      isRunning: false,
    }));
    addLog("info", "执行已取消");
  }, [addLog]);

  // 重置状态
  const reset = React.useCallback(() => {
    setState(initialState);
  }, []);

  // ========== 旧版兼容 API ==========

  const pause = React.useCallback(() => {
    console.warn("[useAgentV4] pause() is deprecated");
    return false;
  }, []);

  const resume = React.useCallback(() => {
    console.warn("[useAgentV4] resume() is deprecated");
    return false;
  }, []);

  const setStepCallback = React.useCallback((callback: ((step: string) => void) | null) => {
    stepCallbackRef.current = callback;
  }, []);

  const approve = React.useCallback((approvalId: string) => {
    console.log("[useAgentV4] approve:", approvalId);
    // TODO: 实现审批逻辑
  }, []);

  const reject = React.useCallback((approvalId: string, reason?: string) => {
    console.log("[useAgentV4] reject:", approvalId, reason);
    // TODO: 实现拒绝逻辑
  }, []);

  // 初始化旧版 Agent（兼容层）
  React.useEffect(() => {
    if (legacyAgentRef.current) return;

    const agent = new Agent({
      maxIterations,
      enableMemory,
      verboseLogging,
    });

    agent.registerTools(createExcelTools());
    agent.setExcelReader(createExcelReader());

    legacyAgentRef.current = agent;
    console.log("[useAgentV4] Legacy Agent initialized for compatibility");
  }, [maxIterations, enableMemory, verboseLogging]);

  return {
    send,
    cancel,
    state,
    reset,
    executor: executorRef.current,
    // 旧版兼容 API
    pause,
    resume,
    agentInstance: legacyAgentRef.current,
    setStepCallback,
    approve,
    reject,
    approvalManager: null, // TODO: 实现审批管理器
  };
}

export default useAgentV4;
