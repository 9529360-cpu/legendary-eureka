/**
 * Agent Hook
 *
 * 封装 Agent 的调用边界，只暴露：
 * - send(): 发送请求
 * - cancel(): 取消执行
 * - state: 当前状态
 * - approval: 审批相关方法
 *
 * UI 层不应该知道 Agent 内部结构
 *
 * v3.0: 新增审批闭环支持
 * - 高风险操作自动触发审批
 * - approvalRequest 状态暴露给 UI
 * - approve/reject 方法处理用户决定
 *
 * @module useAgent
 */

import * as React from "react";
import {
  Agent,
  createExcelTools,
  createExcelReader,
  AgentTask,
  AgentStep as AgentCoreStep,
  // v3.0: 审批管理
  ApprovalManager,
  approvalManager,
  ApprovalRequest,
} from "../../agent";

// ========== Types ==========

export type AgentStatus =
  | "idle"
  | "running"
  | "completed"
  | "failed"
  | "cancelled"
  | "pending"
  | "awaiting_approval";

// v2.9.17: 进度信息接口
export interface AgentProgress {
  iteration: number;
  maxIterations: number;
  planSteps: number;
  completedSteps: number;
  currentPhase: string;
  percentage: number;
}

export interface AgentState {
  status: AgentStatus;
  isRunning: boolean;
  currentSteps: string[];
  lastTask: AgentTask | null;
  error: string | null;
  // v2.9.17: 进度信息
  progress: AgentProgress | null;
  // v3.0: 审批状态
  pendingApproval: ApprovalRequest | null;
}

export interface AgentStepEvent {
  type: "think" | "act" | "observe" | "plan" | "validate" | "approval";
  text: string;
  timestamp: Date;
}

export interface UseAgentOptions {
  maxIterations?: number;
  enableMemory?: boolean;
  verboseLogging?: boolean;
  onStep?: (event: AgentStepEvent) => void;
}

export interface UseAgentReturn {
  /** 发送请求给 Agent */
  send: (request: string, context?: AgentContext) => Promise<AgentTask>;
  /** 取消当前执行 */
  cancel: () => boolean;
  /** v2.9.17: 暂停执行 */
  pause: () => boolean;
  /** v2.9.17: 恢复执行 */
  resume: () => boolean;
  /** 当前状态 */
  state: AgentState;
  /** 重置状态 */
  reset: () => void;
  /** v3.0: 批准当前待审批操作 */
  approve: (approvalId: string) => void;
  /** v3.0: 拒绝当前待审批操作 */
  reject: (approvalId: string, reason?: string) => void;
  /** v3.0: 获取审批管理器 */
  approvalManager: ApprovalManager;
  /**
   * 获取底层 Agent 实例（过渡期使用）
   * @deprecated 应该通过 send() 方法与 Agent 交互，而不是直接访问实例
   */
  agentInstance: Agent | null;
  /**
   * 设置步骤回调（过渡期使用）
   * @deprecated 应该通过 options.onStep 配置
   */
  setStepCallback: (callback: ((step: string) => void) | null) => void;
}

export interface AgentContext {
  environment?: string;
  selectedData?: unknown;
  workbookInfo?: unknown;
}

// ========== Helper Functions ==========

/**
 * 将工具名转换为用户友好的描述
 */
function getToolDescription(toolName: string): string {
  if (toolName.includes("write_range")) return "写入数据";
  if (toolName.includes("formula")) return "设置公式";
  if (toolName.includes("format")) return "格式化";
  if (toolName.includes("chart")) return "创建图表";
  if (toolName.includes("create_sheet")) return "创建工作表";
  if (toolName.includes("switch_sheet")) return "切换工作表";
  if (toolName.includes("validation")) return "设置数据验证";
  if (toolName.includes("read")) return "读取数据";
  if (toolName.includes("analyze")) return "分析数据";
  return toolName.replace("excel_", "").replace(/_/g, " ");
}

// ========== Hook ==========

export function useAgent(options: UseAgentOptions = {}): UseAgentReturn {
  const { maxIterations = 30, enableMemory = true, verboseLogging = true, onStep } = options;

  // Agent 实例（单例）
  const agentRef = React.useRef<Agent | null>(null);

  // v2.9.28: 执行锁 - 防止重复执行
  const executionLockRef = React.useRef<string | null>(null);
  const isExecutingRef = React.useRef(false);

  // 状态
  const [state, setState] = React.useState<AgentState>({
    status: "idle",
    isRunning: false,
    currentSteps: [],
    lastTask: null,
    error: null,
    progress: null,
    pendingApproval: null,
  });

  // 步骤回调 ref（避免闭包问题）
  const onStepRef = React.useRef(onStep);
  onStepRef.current = onStep;

  // 外部步骤回调（过渡期使用）
  const externalStepCallbackRef = React.useRef<((step: string) => void) | null>(null);

  // 初始化 Agent
  React.useEffect(() => {
    if (agentRef.current) return;

    const agent = new Agent({
      maxIterations,
      enableMemory,
      verboseLogging,
    });

    // 注册 Excel 工具
    agent.registerTools(createExcelTools());

    // 注入 ExcelReader
    agent.setExcelReader(createExcelReader());

    // v2.9.17: 监听迭代进度事件
    agent.on("iteration:start", (data: unknown) => {
      const { progress } = data as {
        iteration: number;
        task: AgentTask;
        progress: {
          iteration: number;
          maxIterations: number;
          planSteps: number;
          completedSteps: number;
          currentPhase: string;
        };
      };

      if (progress) {
        const percentage =
          progress.planSteps > 0
            ? Math.round((progress.completedSteps / progress.planSteps) * 100)
            : Math.round((progress.iteration / progress.maxIterations) * 100);

        setState((prev) => ({
          ...prev,
          progress: {
            ...progress,
            percentage,
          },
        }));
      }
    });

    // 监听事件
    // v2.9.23: 简化输出 - 只显示关键步骤，不显示每个工具调用
    let lastThinkText = "";
    let actionCount = 0;

    agent.on("step:think", (data: unknown) => {
      const { step } = data as { step: AgentCoreStep };
      // 只在思考内容变化时更新，避免重复
      const thought = step.thought || "思考中...";
      if (thought === lastThinkText) return;
      lastThinkText = thought;

      // 简化思考输出 - 只显示第一次
      if (actionCount === 0) {
        const text = `🤔 正在分析任务...`;
        externalStepCallbackRef.current?.(text);
      }

      onStepRef.current?.({
        type: "think",
        text: thought,
        timestamp: new Date(),
      });
    });

    agent.on("step:act", (data: unknown) => {
      const { step } = data as { step: AgentCoreStep };
      actionCount++;

      // v2.9.23: 只在写入操作时通知用户，读取操作静默执行
      const toolName = step.toolName || "";
      const isWriteOperation =
        toolName.includes("write") ||
        toolName.includes("set") ||
        toolName.includes("format") ||
        toolName.includes("create") ||
        toolName.includes("delete");

      if (isWriteOperation) {
        const toolDesc = getToolDescription(toolName);
        const text = `🔧 ${toolDesc}...`;
        externalStepCallbackRef.current?.(text);
      }
      // 读取操作不更新 UI

      onStepRef.current?.({
        type: "act",
        text: `执行: ${step.toolName}`,
        timestamp: new Date(),
      });
    });

    agent.on("step:observe", (data: unknown) => {
      const { step, result } = data as { step: AgentCoreStep; result: { success: boolean } };

      // v2.9.23: 只在失败或重要结果时通知
      if (!result.success) {
        const text = `❌ 操作失败，正在重试...`;
        externalStepCallbackRef.current?.(text);
      }
      // 成功时静默，不刷新 UI

      onStepRef.current?.({
        type: "observe",
        text: step.observation || "完成",
        timestamp: new Date(),
      });
    });

    agent.on("step:plan", (data: unknown) => {
      const { step: _step } = data as { step: AgentCoreStep };
      // v2.9.23: 规划阶段简化输出
      const text = `📋 制定执行计划...`;
      externalStepCallbackRef.current?.(text);

      setState((prev) => ({
        ...prev,
        currentSteps: [...prev.currentSteps, text],
      }));

      onStepRef.current?.({
        type: "plan",
        text,
        timestamp: new Date(),
      });
    });

    agent.on("step:validate", (data: unknown) => {
      const { step } = data as { step: AgentCoreStep };
      const hasErrors = step.validationErrors && step.validationErrors.length > 0;

      // v2.9.50: 修复承诺性措辞，只描述事实不承诺动作
      if (hasErrors) {
        const text = `⚠️ 验证发现 ${step.validationErrors?.length || 1} 个问题`;
        externalStepCallbackRef.current?.(text);

        setState((prev) => ({
          ...prev,
          currentSteps: [...prev.currentSteps, text],
        }));

        onStepRef.current?.({
          type: "validate",
          text,
          timestamp: new Date(),
        });
      }
      // 验证通过时完全静默，不更新任何状态
    });

    // v2.9.41: 订阅写操作预览事件
    agent.on("write:preview", (data: unknown) => {
      const { toolName, description, riskLevel } = data as {
        toolName: string;
        description: string;
        riskLevel: string;
      };
      console.log(`[useAgent] 📝 写操作预览: ${toolName} - ${description} (风险: ${riskLevel})`);

      const text = `📝 准备${description}...`;
      externalStepCallbackRef.current?.(text);
    });

    // v2.9.41: 订阅计划确认事件
    agent.on("plan:confirmation_required", (data: unknown) => {
      const confirmRequest = data as {
        planId: string;
        taskDescription: string;
        estimatedSteps: number;
      };
      console.log(`[useAgent] ⚠️ 计划需要确认: ${confirmRequest.taskDescription}`);

      const text = `⚠️ 发现复杂任务，需要确认执行计划...`;
      externalStepCallbackRef.current?.(text);

      setState((prev) => ({
        ...prev,
        status: "pending",
        currentSteps: [...prev.currentSteps, text],
      }));
    });

    // v3.0: 订阅高风险操作审批事件
    agent.on("approval:required", (data: unknown) => {
      const { approvalRequest } = data as { approvalRequest: ApprovalRequest };
      console.log(`[useAgent] 🔒 需要用户审批: ${approvalRequest.approvalId}`);

      const text = `⚠️ 高风险操作需要确认: ${approvalRequest.operationName}`;
      externalStepCallbackRef.current?.(text);

      onStepRef.current?.({
        type: "approval",
        text,
        timestamp: new Date(),
      });

      setState((prev) => ({
        ...prev,
        status: "awaiting_approval",
        pendingApproval: approvalRequest,
        currentSteps: [...prev.currentSteps, text],
      }));
    });

    agentRef.current = agent;
    console.log("[useAgent] Agent initialized with Excel tools");
  }, [maxIterations, enableMemory, verboseLogging]);

  // 发送请求
  const send = React.useCallback(
    async (request: string, context?: AgentContext): Promise<AgentTask> => {
      if (!agentRef.current) {
        throw new Error("Agent not initialized");
      }

      // v2.9.28: 执行锁检查 - 防止重复执行
      if (isExecutingRef.current) {
        console.warn("[useAgent] ⚠️ 任务正在执行中，忽略重复请求");
        throw new Error("任务正在执行中，请等待完成");
      }

      // 生成唯一执行 ID
      const executionId = `exec_${Date.now()}_${Math.random().toString(36).substring(2, 9)}`;
      executionLockRef.current = executionId;
      isExecutingRef.current = true;
      console.log(`[useAgent] 🔒 获取执行锁: ${executionId}`);

      // 重置状态
      setState({
        status: "running",
        isRunning: true,
        currentSteps: [],
        lastTask: null,
        error: null,
        progress: null,
        pendingApproval: null,
      });

      try {
        // 检查锁是否仍然属于当前执行
        if (executionLockRef.current !== executionId) {
          console.warn(`[useAgent] ⚠️ 执行锁已被覆盖: ${executionId}`);
          throw new Error("执行已被取消");
        }

        const task = await agentRef.current.run(request, {
          environment: context?.environment || "excel",
          selectedData: context?.selectedData,
          workbookInfo: context?.workbookInfo,
        });

        // v2.9.25: 正确处理所有任务状态，包括 pending
        // v2.9.44: 添加 pending_confirmation 状态处理
        const finalStatus: AgentStatus =
          task.status === "completed"
            ? "completed"
            : task.status === "cancelled"
              ? "cancelled"
              : task.status === "pending" || task.status === "pending_confirmation"
                ? "pending" // Agent 等待用户回复或确认
                : "failed";

        // v2.9.44: 如果是待确认状态，不释放执行锁
        const shouldReleaseLock = task.status !== "pending_confirmation";

        setState((prev) => ({
          ...prev,
          status: finalStatus,
          isRunning: task.status === "pending_confirmation", // 待确认时仍显示运行中
          lastTask: task,
        }));

        // v2.9.44: 根据状态决定是否释放锁
        if (!shouldReleaseLock) {
          console.log(`[useAgent] 🔒 任务待确认，保持执行锁`);
        }

        return task;
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);

        setState((prev) => ({
          ...prev,
          status: "failed",
          isRunning: false,
          error: errorMessage,
        }));

        throw error;
      } finally {
        // v2.9.28: 释放执行锁
        // v3.0.6: 简化锁释放逻辑，移除不存在的getCurrentTask调用
        if (executionLockRef.current === executionId) {
          console.log(`[useAgent] 🔓 释放执行锁: ${executionId}`);
          executionLockRef.current = null;
          isExecutingRef.current = false;
        }
      }
    },
    []
  );

  // 取消执行
  const cancel = React.useCallback((): boolean => {
    if (!agentRef.current) {
      return false;
    }

    const result = agentRef.current.cancelCurrentTask();

    if (result) {
      // v2.9.28: 取消时也释放执行锁
      console.log("[useAgent] 🔓 取消任务，释放执行锁");
      executionLockRef.current = null;
      isExecutingRef.current = false;

      setState((prev) => ({
        ...prev,
        status: "cancelled",
        isRunning: false,
      }));
    }

    return result;
  }, []);

  // v2.9.17: 暂停执行
  const pause = React.useCallback((): boolean => {
    if (!agentRef.current) {
      return false;
    }

    const result = agentRef.current.pauseTask();

    if (result) {
      setState((prev) => ({
        ...prev,
        progress: prev.progress ? { ...prev.progress, currentPhase: "已暂停" } : null,
      }));
    }

    return result;
  }, []);

  // v2.9.17: 恢复执行
  const resume = React.useCallback((): boolean => {
    if (!agentRef.current) {
      return false;
    }

    const result = agentRef.current.resumeTask();

    if (result) {
      setState((prev) => ({
        ...prev,
        progress: prev.progress ? { ...prev.progress, currentPhase: "已恢复" } : null,
      }));
    }

    return result;
  }, []);

  // 重置状态
  const reset = React.useCallback(() => {
    setState({
      status: "idle",
      isRunning: false,
      currentSteps: [],
      lastTask: null,
      error: null,
      progress: null,
      pendingApproval: null,
    });
  }, []);

  // v3.0: 批准审批请求
  const approve = React.useCallback((approvalId: string) => {
    const result = approvalManager.handleApprovalDecision(approvalId, true, "user");

    if (result.success) {
      console.log(`[useAgent] ✅ 审批通过: ${approvalId}`);

      // 通知 Agent 继续执行
      if (agentRef.current) {
        agentRef.current.emit("approval:granted", { approvalId, request: result.request });
      }

      setState((prev) => ({
        ...prev,
        status: "running",
        pendingApproval: null,
        currentSteps: [...prev.currentSteps, `✅ 已确认执行 ${approvalId}`],
      }));
    } else {
      console.error(`[useAgent] ❌ 审批处理失败: ${result.error}`);
    }
  }, []);

  // v3.0: 拒绝审批请求
  const reject = React.useCallback((approvalId: string, reason?: string) => {
    const result = approvalManager.handleApprovalDecision(approvalId, false, "user", reason);

    if (result.success) {
      console.log(`[useAgent] ❌ 审批拒绝: ${approvalId}`);

      // 通知 Agent 取消操作
      if (agentRef.current) {
        agentRef.current.emit("approval:rejected", { approvalId, reason });
      }

      setState((prev) => ({
        ...prev,
        status: "completed",
        isRunning: false,
        pendingApproval: null,
        currentSteps: [...prev.currentSteps, `❌ 已取消操作 ${approvalId}`],
      }));
    } else {
      console.error(`[useAgent] ❌ 审批处理失败: ${result.error}`);
    }
  }, []);

  // 设置外部步骤回调（过渡期使用）
  const setStepCallback = React.useCallback((callback: ((step: string) => void) | null) => {
    externalStepCallbackRef.current = callback;
  }, []);

  return {
    send,
    cancel,
    pause,
    resume,
    state,
    reset,
    // v3.0: 审批相关
    approve,
    reject,
    approvalManager,
    // 过渡期 API
    agentInstance: agentRef.current,
    setStepCallback,
  };
}

export default useAgent;
