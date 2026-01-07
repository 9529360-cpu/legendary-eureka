/**
 * ParallelExecutor - 并行执行引擎 v4.1
 *
 * 基于 DAG（有向无环图）分析步骤依赖，并行执行独立步骤
 *
 * 核心特性：
 * 1. 依赖图构建与拓扑排序
 * 2. 独立步骤并行执行
 * 3. 失败传播与隔离
 * 4. 执行进度追踪
 *
 * @module agent/ParallelExecutor
 */

import { ToolRegistry } from "./registry";
import { ToolResult } from "./types/tool";
import { RecoverableStep, RecoveryManager, RecoveryAction } from "./RecoveryManager";

// ========== 类型定义 ==========

/**
 * DAG 节点状态
 */
export type NodeStatus = "pending" | "ready" | "running" | "completed" | "failed" | "skipped";

/**
 * DAG 节点
 */
export interface DAGNode {
  /** 步骤 */
  step: RecoverableStep;

  /** 状态 */
  status: NodeStatus;

  /** 依赖的节点 ID */
  dependencies: string[];

  /** 被依赖的节点 ID（出边） */
  dependents: string[];

  /** 执行结果 */
  result?: StepExecutionResult;

  /** 开始时间 */
  startTime?: number;

  /** 结束时间 */
  endTime?: number;
}

/**
 * 步骤执行结果
 */
export interface StepExecutionResult {
  success: boolean;
  output: string;
  error?: string;
  duration: number;
  recovered?: boolean;
  recoveryAction?: string;
}

/**
 * 并行执行结果
 */
export interface ParallelExecutionResult {
  /** 是否全部成功 */
  success: boolean;

  /** 执行的步骤数 */
  totalSteps: number;

  /** 成功步骤数 */
  successCount: number;

  /** 失败步骤数 */
  failedCount: number;

  /** 跳过步骤数 */
  skippedCount: number;

  /** 各步骤结果 */
  stepResults: Map<string, StepExecutionResult>;

  /** 总耗时 (ms) */
  totalDuration: number;

  /** 并行度统计 */
  parallelism: {
    maxConcurrent: number;
    avgConcurrent: number;
    batches: number;
  };
}

/**
 * 执行事件
 */
export interface ExecutionEvent {
  type: "batch:start" | "step:start" | "step:complete" | "step:error" | "step:skip" | "complete";
  stepId?: string;
  batchIndex?: number;
  batchSize?: number;
  data?: unknown;
  timestamp: number;
}

/**
 * 执行选项
 */
export interface ParallelExecutionOptions {
  /** 最大并发数 */
  maxConcurrency?: number;

  /** 是否启用错误恢复 */
  enableRecovery?: boolean;

  /** 失败时是否继续 */
  continueOnFailure?: boolean;

  /** 取消信号 */
  signal?: AbortSignal;

  /** 事件回调 */
  onEvent?: (event: ExecutionEvent) => void;
}

// ========== DAG 构建器 ==========

/**
 * 构建依赖图
 */
export function buildDAG(steps: RecoverableStep[]): Map<string, DAGNode> {
  const nodes = new Map<string, DAGNode>();

  // 创建所有节点
  for (const step of steps) {
    nodes.set(step.id, {
      step,
      status: "pending",
      dependencies: [...(step.dependsOn || [])],
      dependents: [],
    });
  }

  // 构建反向边（dependents）
  for (const [id, node] of nodes) {
    for (const depId of node.dependencies) {
      const depNode = nodes.get(depId);
      if (depNode) {
        depNode.dependents.push(id);
      }
    }
  }

  // 标记初始可执行节点
  for (const [, node] of nodes) {
    if (node.dependencies.length === 0) {
      node.status = "ready";
    }
  }

  return nodes;
}

/**
 * 检测循环依赖
 */
export function detectCycle(nodes: Map<string, DAGNode>): string[] | null {
  const visited = new Set<string>();
  const recStack = new Set<string>();
  const path: string[] = [];

  function dfs(nodeId: string): boolean {
    visited.add(nodeId);
    recStack.add(nodeId);
    path.push(nodeId);

    const node = nodes.get(nodeId);
    if (node) {
      for (const depId of node.dependents) {
        if (!visited.has(depId)) {
          if (dfs(depId)) return true;
        } else if (recStack.has(depId)) {
          // 找到循环
          const cycleStart = path.indexOf(depId);
          return true;
        }
      }
    }

    path.pop();
    recStack.delete(nodeId);
    return false;
  }

  for (const nodeId of nodes.keys()) {
    if (!visited.has(nodeId)) {
      if (dfs(nodeId)) {
        return path;
      }
    }
  }

  return null;
}

/**
 * 获取可执行的节点（依赖已完成）
 */
export function getReadyNodes(nodes: Map<string, DAGNode>): DAGNode[] {
  const ready: DAGNode[] = [];

  for (const [, node] of nodes) {
    if (node.status === "pending") {
      // 检查所有依赖是否已完成
      const allDepsCompleted = node.dependencies.every((depId) => {
        const dep = nodes.get(depId);
        return dep && (dep.status === "completed" || dep.status === "skipped");
      });

      // 检查是否有失败的依赖
      const hasFailedDep = node.dependencies.some((depId) => {
        const dep = nodes.get(depId);
        return dep && dep.status === "failed";
      });

      if (hasFailedDep) {
        node.status = "skipped";
      } else if (allDepsCompleted) {
        node.status = "ready";
        ready.push(node);
      }
    } else if (node.status === "ready") {
      ready.push(node);
    }
  }

  return ready;
}

// ========== ParallelExecutor 类 ==========

/**
 * 并行执行器
 */
export class ParallelExecutor {
  private toolRegistry: ToolRegistry;
  private recoveryManager: RecoveryManager;

  constructor(toolRegistry: ToolRegistry, recoveryManager?: RecoveryManager) {
    this.toolRegistry = toolRegistry;
    this.recoveryManager = recoveryManager ?? new RecoveryManager();
  }

  /**
   * 并行执行步骤
   */
  async execute(
    steps: RecoverableStep[],
    options: ParallelExecutionOptions = {}
  ): Promise<ParallelExecutionResult> {
    const {
      maxConcurrency = 5,
      enableRecovery = true,
      continueOnFailure = true,
      signal,
      onEvent,
    } = options;

    const startTime = Date.now();
    const stepResults = new Map<string, StepExecutionResult>();
    let batchCount = 0;
    let maxConcurrent = 0;
    let totalConcurrent = 0;
    let concurrentSamples = 0;

    // 构建 DAG
    const nodes = buildDAG(steps);

    // 检测循环
    const cycle = detectCycle(nodes);
    if (cycle) {
      console.error("[ParallelExecutor] 检测到循环依赖:", cycle);
      return {
        success: false,
        totalSteps: steps.length,
        successCount: 0,
        failedCount: steps.length,
        skippedCount: 0,
        stepResults,
        totalDuration: Date.now() - startTime,
        parallelism: { maxConcurrent: 0, avgConcurrent: 0, batches: 0 },
      };
    }

    // 存储步骤输出（供后续步骤引用）
    const outputs = new Map<string, string>();

    // 执行循环
    while (true) {
      // 检查取消
      if (signal?.aborted) {
        console.log("[ParallelExecutor] 执行已取消");
        break;
      }

      // 获取可执行节点
      const ready = getReadyNodes(nodes);

      if (ready.length === 0) {
        // 检查是否全部完成
        const allDone = Array.from(nodes.values()).every(
          (n) => n.status === "completed" || n.status === "failed" || n.status === "skipped"
        );
        if (allDone) break;

        // 如果有节点既不是 ready 也不是 done，说明有问题
        console.warn("[ParallelExecutor] 无可执行节点但未全部完成");
        break;
      }

      // 限制并发数
      const batch = ready.slice(0, maxConcurrency);
      batchCount++;
      maxConcurrent = Math.max(maxConcurrent, batch.length);
      totalConcurrent += batch.length;
      concurrentSamples++;

      onEvent?.({
        type: "batch:start",
        batchIndex: batchCount,
        batchSize: batch.length,
        timestamp: Date.now(),
      });

      console.log(`[ParallelExecutor] 批次 ${batchCount}: 并行执行 ${batch.length} 个步骤`);

      // 标记为运行中
      for (const node of batch) {
        node.status = "running";
        node.startTime = Date.now();
      }

      // 并行执行
      const batchPromises = batch.map(async (node) => {
        const step = node.step;

        onEvent?.({
          type: "step:start",
          stepId: step.id,
          timestamp: Date.now(),
        });

        try {
          const result = await this.executeStep(step, outputs, enableRecovery);

          node.endTime = Date.now();
          node.result = result;
          stepResults.set(step.id, result);

          if (result.success) {
            node.status = "completed";
            outputs.set(step.id, result.output);

            onEvent?.({
              type: "step:complete",
              stepId: step.id,
              data: { output: result.output },
              timestamp: Date.now(),
            });
          } else {
            node.status = "failed";

            onEvent?.({
              type: "step:error",
              stepId: step.id,
              data: { error: result.error },
              timestamp: Date.now(),
            });

            if (!continueOnFailure) {
              // 标记所有依赖此节点的为 skipped
              this.propagateFailure(node, nodes, onEvent);
            }
          }
        } catch (error) {
          const errorMsg = error instanceof Error ? error.message : String(error);
          node.status = "failed";
          node.endTime = Date.now();

          const result: StepExecutionResult = {
            success: false,
            output: "",
            error: errorMsg,
            duration: (node.endTime || Date.now()) - (node.startTime || Date.now()),
          };
          node.result = result;
          stepResults.set(step.id, result);

          onEvent?.({
            type: "step:error",
            stepId: step.id,
            data: { error: errorMsg },
            timestamp: Date.now(),
          });
        }
      });

      // 等待当前批次完成
      await Promise.all(batchPromises);
    }

    // 统计结果
    let successCount = 0;
    let failedCount = 0;
    let skippedCount = 0;

    for (const [, node] of nodes) {
      if (node.status === "completed") successCount++;
      else if (node.status === "failed") failedCount++;
      else if (node.status === "skipped") skippedCount++;
    }

    const totalDuration = Date.now() - startTime;

    onEvent?.({
      type: "complete",
      timestamp: Date.now(),
      data: { successCount, failedCount, skippedCount, totalDuration },
    });

    return {
      success: failedCount === 0,
      totalSteps: steps.length,
      successCount,
      failedCount,
      skippedCount,
      stepResults,
      totalDuration,
      parallelism: {
        maxConcurrent,
        avgConcurrent: concurrentSamples > 0 ? totalConcurrent / concurrentSamples : 0,
        batches: batchCount,
      },
    };
  }

  /**
   * 执行单个步骤
   */
  private async executeStep(
    step: RecoverableStep,
    outputs: Map<string, string>,
    enableRecovery: boolean
  ): Promise<StepExecutionResult> {
    const startTime = Date.now();

    const tool = this.toolRegistry.get(step.action);
    if (!tool) {
      return {
        success: false,
        output: "",
        error: `工具不存在: ${step.action}`,
        duration: Date.now() - startTime,
      };
    }

    try {
      // 解析参数（替换依赖输出引用）
      const params = this.resolveParameters(step.parameters || {}, outputs, step.dependsOn || []);

      // 执行工具
      const result: ToolResult = await tool.execute(params);

      if (result.success) {
        return {
          success: true,
          output: typeof result.output === "string" ? result.output : JSON.stringify(result.output),
          duration: Date.now() - startTime,
        };
      } else {
        // 尝试恢复
        if (enableRecovery) {
          const recovery = await this.recoveryManager.recover(step, new Error(result.error || "执行失败"));
          if (recovery && recovery.type === "retry") {
            // 重试
            await this.delay(recovery.delay || 500);
            const retryResult: ToolResult = await tool.execute(params);
            if (retryResult.success) {
              return {
                success: true,
                output:
                  typeof retryResult.output === "string"
                    ? retryResult.output
                    : JSON.stringify(retryResult.output),
                duration: Date.now() - startTime,
                recovered: true,
                recoveryAction: "retry",
              };
            }
          }
        }

        return {
          success: false,
          output: "",
          error: result.error || "执行失败",
          duration: Date.now() - startTime,
        };
      }
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : String(error);

      // 尝试恢复
      if (enableRecovery) {
        const recovery = await this.recoveryManager.recover(step, error as Error);
        if (recovery && recovery.type === "skip") {
          return {
            success: true,
            output: "",
            error: `已跳过: ${recovery.reason}`,
            duration: Date.now() - startTime,
            recovered: true,
            recoveryAction: "skip",
          };
        }
      }

      return {
        success: false,
        output: "",
        error: errorMsg,
        duration: Date.now() - startTime,
      };
    }
  }

  /**
   * 解析参数（替换依赖引用）
   */
  private resolveParameters(
    params: Record<string, unknown>,
    outputs: Map<string, string>,
    dependsOn: string[]
  ): Record<string, unknown> {
    const resolved: Record<string, unknown> = {};

    for (const [key, value] of Object.entries(params)) {
      if (typeof value === "string") {
        let resolvedValue = value;

        // 替换 {{stepId}} 引用
        for (const depId of dependsOn) {
          const output = outputs.get(depId);
          if (output) {
            resolvedValue = resolvedValue.replace(`{{${depId}}}`, output);
          }
        }

        // 替换 {{previous}} 为最后一个依赖的输出
        if (resolvedValue.includes("{{previous}}") && dependsOn.length > 0) {
          const lastDep = dependsOn[dependsOn.length - 1];
          const lastOutput = outputs.get(lastDep) || "";
          resolvedValue = resolvedValue.replace("{{previous}}", lastOutput);
        }

        resolved[key] = resolvedValue;
      } else {
        resolved[key] = value;
      }
    }

    return resolved;
  }

  /**
   * 传播失败（标记依赖此节点的所有节点为 skipped）
   */
  private propagateFailure(
    failedNode: DAGNode,
    nodes: Map<string, DAGNode>,
    onEvent?: (event: ExecutionEvent) => void
  ): void {
    const toSkip = new Set<string>();
    const queue = [...failedNode.dependents];

    while (queue.length > 0) {
      const depId = queue.shift()!;
      if (toSkip.has(depId)) continue;

      const node = nodes.get(depId);
      if (node && node.status === "pending") {
        node.status = "skipped";
        toSkip.add(depId);
        queue.push(...node.dependents);

        onEvent?.({
          type: "step:skip",
          stepId: depId,
          data: { reason: `依赖 ${failedNode.step.id} 失败` },
          timestamp: Date.now(),
        });
      }
    }
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
 * 创建并行执行器
 */
export function createParallelExecutor(
  toolRegistry: ToolRegistry,
  recoveryManager?: RecoveryManager
): ParallelExecutor {
  return new ParallelExecutor(toolRegistry, recoveryManager);
}

export default ParallelExecutor;
