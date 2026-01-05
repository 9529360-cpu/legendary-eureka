/**
 * 反假完成闭环控制器 - AntiHallucinationController
 *
 * 职责：
 * 1. 整合所有门槛、拦截器、验证器
 * 2. 实现完整的"提交→验证→失败重试→放行"闭环
 * 3. 确保模型无法"假装完成"
 *
 * 核心原则：完成权不在模型，在系统
 */

import { AgentRun, AgentState, Submission, createAgentRun, isChecklistComplete } from "./types";
import { StateMachine } from "./StateMachine";
import { CompletionGate, GateCheckResult } from "./CompletionGate";
import { SubmissionParser, ParseResult } from "./SubmissionParser";
import { ValidationEngine, ValidationReport } from "./ValidationEngine";
import {
  CompletionInterceptor,
  SelfReferenceInterceptor,
  MaxIterationsInterceptor,
  completionInterceptor,
  selfReferenceInterceptor,
  maxIterationsInterceptor,
} from "./Interceptors";

// ========== 处理结果 ==========

/**
 * 回合处理结果
 */
export interface TurnResult {
  /** 是否允许结束 */
  allowFinish: boolean;
  /** 最终状态 */
  state: AgentState;
  /** 需要发送给模型的系统消息（如果需要继续） */
  systemMessage?: string;
  /** 需要返回给用户的消息（如果完成） */
  userMessage?: string;
  /** 验证报告 */
  validationReport?: ValidationReport;
  /** 门槛检查结果 */
  gateCheckResult?: GateCheckResult;
  /** 解析结果 */
  parseResult?: ParseResult;
}

// ========== AntiHallucinationController ==========

/**
 * 反假完成闭环控制器
 */
export class AntiHallucinationController {
  private parser: SubmissionParser;
  private validator: ValidationEngine;
  private completionInterceptor: CompletionInterceptor;
  private selfRefInterceptor: SelfReferenceInterceptor;
  private maxIterInterceptor: MaxIterationsInterceptor;

  constructor() {
    this.parser = new SubmissionParser();
    this.validator = new ValidationEngine();
    this.completionInterceptor = completionInterceptor;
    this.selfRefInterceptor = selfReferenceInterceptor;
    this.maxIterInterceptor = maxIterationsInterceptor;
  }

  /**
   * 创建新的运行实例
   */
  createRun(userId: string, taskId: string): AgentRun {
    return createAgentRun(userId, taskId);
  }

  /**
   * 处理用户输入
   */
  handleUserMessage(run: AgentRun, userMessage: string): void {
    run.history.push({
      role: "user",
      content: userMessage,
      timestamp: Date.now(),
    });
    run.iteration++;
    run.updatedAt = Date.now();
  }

  /**
   * 处理模型输出（核心方法）
   */
  handleModelOutput(run: AgentRun, modelOutput: string): TurnResult {
    // 保存模型输出
    run.lastModelOutput = modelOutput;
    run.history.push({
      role: "assistant",
      content: modelOutput,
      timestamp: Date.now(),
    });

    // 1. 检查最大迭代次数
    const maxIterResult = this.maxIterInterceptor.intercept(run);
    if (maxIterResult.intercepted) {
      return {
        allowFinish: false,
        state: run.state,
        userMessage: maxIterResult.systemMessage,
      };
    }

    // 2. 解析模型输出
    const parseResult = this.parser.parse(modelOutput);

    // 3. 拦截器检查（格式完整性）
    const completionResult = this.completionInterceptor.intercept(parseResult);
    if (completionResult.intercepted) {
      return {
        allowFinish: false,
        state: run.state,
        systemMessage: completionResult.systemMessage,
        parseResult,
      };
    }

    const submission = parseResult.submission!;

    // 4. 验证引擎检查（规则验证）
    const validationReport = this.validator.validate(submission);

    // 5. 自引用拦截器
    const selfRefResult = this.selfRefInterceptor.intercept(validationReport);
    if (selfRefResult.intercepted) {
      return {
        allowFinish: false,
        state: run.state,
        systemMessage: selfRefResult.systemMessage,
        validationReport,
        parseResult,
      };
    }

    // 6. 完成门槛检查
    const gateCheckResult = CompletionGate.check(run, submission);

    // 更新 run 的状态
    run.artifacts = submission.artifacts;
    run.checklist = gateCheckResult.checklist;
    run.validations = gateCheckResult.validations;

    // 7. 判断是否可以放行
    if (gateCheckResult.passed && validationReport.allPassed) {
      // 更新 checklist 中的验证结果
      run.checklist.supportsAutoExpand = true;
      run.checklist.avoidsSelfReference = true;

      // 状态转换到 DEPLOYED
      StateMachine.transition(run, AgentState.VERIFIED);
      StateMachine.transition(run, AgentState.DEPLOYED);

      return {
        allowFinish: true,
        state: AgentState.DEPLOYED,
        userMessage: this.generateSuccessMessage(run, submission),
        validationReport,
        gateCheckResult,
        parseResult,
      };
    }

    // 8. 不能放行，强制继续
    const nextState = StateMachine.nextStateAfterFail(run, run.checklist);
    StateMachine.transition(run, nextState);

    return {
      allowFinish: false,
      state: nextState,
      systemMessage: this.generateForceContinueMessage(gateCheckResult, validationReport),
      validationReport,
      gateCheckResult,
      parseResult,
    };
  }

  /**
   * 生成成功消息
   */
  private generateSuccessMessage(run: AgentRun, submission: Submission): string {
    const lines = [
      "✅ 任务已完成并通过所有验证！",
      "",
      "📋 完成清单：",
      `  ✓ 可执行产物: ${submission.artifacts.length} 个`,
      `  ✓ 验收测试: ${submission.acceptanceTests.length} 条`,
      `  ✓ 回退方案: ${submission.fallback.length} 个`,
      "  ✓ 部署说明: 已提供",
      "",
      "📊 验证结果：",
      "  ✓ 无自引用风险",
      "  ✓ 支持自动扩展",
      "",
      "🚀 可以上线部署。",
    ];
    return lines.join("\n");
  }

  /**
   * 生成强制继续消息
   */
  private generateForceContinueMessage(
    gateResult: GateCheckResult,
    validationReport: ValidationReport
  ): string {
    const lines = ["❌ 未通过上线放行检查。禁止结束。必须修复并重新提交。", ""];

    // 门槛检查失败原因
    if (gateResult.failReasons.length > 0) {
      lines.push("🚫 门槛检查失败：");
      gateResult.failReasons.forEach((r, i) => lines.push(`   ${i + 1}. ${r}`));
      lines.push("");
    }

    // 验证失败原因
    if (validationReport.criticalFails.length > 0) {
      lines.push("🔴 验证失败：");
      validationReport.criticalFails.forEach((v) => {
        lines.push(`   - ${v.name}: ${v.reason}`);
      });
      lines.push("");
    }

    // 警告
    if (validationReport.warnings.length > 0) {
      lines.push("⚠️ 警告：");
      validationReport.warnings.forEach((v) => {
        lines.push(`   - ${v.name}: ${v.reason}`);
      });
      lines.push("");
    }

    // 必须完成的操作
    if (gateResult.requiredActions.length > 0) {
      lines.push("📝 必须完成的操作：");
      gateResult.requiredActions.forEach((a, i) => lines.push(`   ${i + 1}. ${a}`));
      lines.push("");
    }

    lines.push("请修复上述问题后，重新提交完整的提交包（含所有必需块）。");

    return lines.join("\n");
  }

  /**
   * 检查运行是否可以结束
   */
  canFinish(run: AgentRun): boolean {
    return run.state === AgentState.DEPLOYED && isChecklistComplete(run.checklist);
  }

  /**
   * 获取运行状态摘要
   */
  getRunSummary(run: AgentRun): string {
    const checklist = run.checklist;
    const checkItems = [
      `可执行产物: ${checklist.hasExecutableArtifact ? "✓" : "✗"}`,
      `放置位置: ${checklist.hasPlacementInfo ? "✓" : "✗"}`,
      `自动扩展: ${checklist.supportsAutoExpand ? "✓" : "✗"}`,
      `避免自引用: ${checklist.avoidsSelfReference ? "✓" : "✗"}`,
      `验收测试: ${checklist.has3AcceptanceTests ? "✓" : "✗"}`,
      `回退方案: ${checklist.hasFallbackPlan ? "✓" : "✗"}`,
      `部署说明: ${checklist.hasDeployNotes ? "✓" : "✗"}`,
    ];

    return `运行状态: ${run.state}
迭代次数: ${run.iteration}/${run.maxIterations}
产物数量: ${run.artifacts.length}

完成清单:
${checkItems.map((i) => `  ${i}`).join("\n")}`;
  }
}

// ========== 导出单例 ==========

export const antiHallucinationController = new AntiHallucinationController();
