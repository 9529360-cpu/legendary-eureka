/**
 * 状态机 - Agent 运行状态管理
 *
 * 职责：
 * 1. 管理 Agent 状态转换
 * 2. 强制状态推进规则
 * 3. 禁止非法跳转（如 INIT → DEPLOYED）
 */

import { AgentState, AgentRun, Checklist, Validation, ValidationStatus } from "./types";

// ========== 状态转换规则 ==========

/**
 * 允许的状态转换
 */
const ALLOWED_TRANSITIONS: Map<AgentState, AgentState[]> = new Map([
  [AgentState.INIT, [AgentState.ANALYZED]],
  [AgentState.ANALYZED, [AgentState.DESIGNED, AgentState.INIT]], // 可回退
  [AgentState.DESIGNED, [AgentState.EXECUTED, AgentState.ANALYZED]], // 可回退
  [AgentState.EXECUTED, [AgentState.VERIFIED, AgentState.DESIGNED]], // 可回退
  [AgentState.VERIFIED, [AgentState.DEPLOYED, AgentState.EXECUTED]], // 验证失败回退
  [AgentState.DEPLOYED, []], // 终态，不可转换
]);

/**
 * 状态转换结果
 */
export interface TransitionResult {
  success: boolean;
  previousState: AgentState;
  currentState: AgentState;
  reason?: string;
}

// ========== StateMachine 类 ==========

/**
 * Agent 状态机
 */
export class StateMachine {
  /**
   * 尝试状态转换
   */
  static transition(run: AgentRun, targetState: AgentState): TransitionResult {
    const allowed = ALLOWED_TRANSITIONS.get(run.state) || [];

    if (!allowed.includes(targetState)) {
      return {
        success: false,
        previousState: run.state,
        currentState: run.state,
        reason: `不允许从 ${run.state} 转换到 ${targetState}`,
      };
    }

    const previousState = run.state;
    run.state = targetState;
    run.updatedAt = Date.now();

    return {
      success: true,
      previousState,
      currentState: targetState,
    };
  }

  /**
   * 根据验证结果决定下一个状态
   */
  static nextStateAfterValidation(
    run: AgentRun,
    validations: Validation[],
    checklist: Checklist
  ): AgentState {
    const allValidationsPass = validations.every((v) => v.status === ValidationStatus.PASS);
    const checklistComplete = this.isChecklistComplete(checklist);

    // 全部通过 → 可以部署
    if (allValidationsPass && checklistComplete) {
      return AgentState.DEPLOYED;
    }

    // 没有产物 → 回到 DESIGNED
    if (!checklist.hasExecutableArtifact) {
      return AgentState.DESIGNED;
    }

    // 有产物但验证失败 → 回到 EXECUTED
    return AgentState.EXECUTED;
  }

  /**
   * 根据失败原因决定回退状态
   */
  static nextStateAfterFail(run: AgentRun, checklist: Checklist): AgentState {
    // 没有产物 → 回到 DESIGNED
    if (!checklist.hasExecutableArtifact) {
      return AgentState.DESIGNED;
    }
    // 有产物但验证失败 → 回到 EXECUTED
    return AgentState.EXECUTED;
  }

  /**
   * 检查是否可以结束
   */
  static canFinish(run: AgentRun): boolean {
    return run.state === AgentState.DEPLOYED;
  }

  /**
   * 检查是否超过最大迭代次数
   */
  static isMaxIterationsReached(run: AgentRun): boolean {
    return run.iteration >= run.maxIterations;
  }

  /**
   * 获取状态描述
   */
  static getStateDescription(state: AgentState): string {
    const descriptions: Record<AgentState, string> = {
      [AgentState.INIT]: "初始化 - 等待用户输入",
      [AgentState.ANALYZED]: "已分析 - 语义抽取完成",
      [AgentState.DESIGNED]: "已设计 - 方案设计完成",
      [AgentState.EXECUTED]: "已执行 - 产物生成完成",
      [AgentState.VERIFIED]: "已验证 - 系统验证通过",
      [AgentState.DEPLOYED]: "已部署 - 任务完成",
    };
    return descriptions[state];
  }

  /**
   * 检查 Checklist 是否全部完成
   */
  private static isChecklistComplete(checklist: Checklist): boolean {
    return (
      checklist.hasExecutableArtifact &&
      checklist.hasPlacementInfo &&
      checklist.supportsAutoExpand &&
      checklist.avoidsSelfReference &&
      checklist.has3AcceptanceTests &&
      checklist.hasFallbackPlan &&
      checklist.hasDeployNotes
    );
  }
}

// ========== 导出单例 ==========

export const stateMachine = new StateMachine();
