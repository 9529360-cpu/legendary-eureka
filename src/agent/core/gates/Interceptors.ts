/**
 * 拦截器 - Interceptors
 *
 * 职责：
 * 1. CompletionInterceptor: 拦截不完整的模型输出
 * 2. SelfReferenceInterceptor: 拦截自引用公式
 * 3. 强制模型继续修复，不允许偷懒结束
 */

import { ParseResult } from "./SubmissionParser";
import { ValidationReport } from "./ValidationEngine";
import { AgentRun } from "./types";

// ========== 拦截结果 ==========

/**
 * 拦截结果
 */
export interface InterceptResult {
  intercepted: boolean;
  reason?: string;
  systemMessage?: string;
}

// ========== CompletionInterceptor ==========

/**
 * 完成拦截器
 * 如果模型输出缺少必需块，直接拦截并要求补齐
 */
export class CompletionInterceptor {
  /**
   * 检查并拦截
   */
  intercept(parseResult: ParseResult): InterceptResult {
    if (!parseResult.success) {
      const missingBlocks = parseResult.missingBlocks;
      const systemMessage = this.generateRetryMessage(missingBlocks);

      return {
        intercepted: true,
        reason: `输出缺少必需块: ${missingBlocks.join(", ")}`,
        systemMessage,
      };
    }

    // 检查验收测试数量
    const submission = parseResult.submission!;
    if (!submission.acceptanceTests || submission.acceptanceTests.length < 3) {
      return {
        intercepted: true,
        reason: "验收测试不足 3 条",
        systemMessage: this.generateTestsRequiredMessage(submission.acceptanceTests?.length || 0),
      };
    }

    // 检查是否有 NextAction
    if (!submission.nextAction) {
      return {
        intercepted: true,
        reason: "缺少 [NEXT_ACTION] 块",
        systemMessage:
          "请添加 [NEXT_ACTION] 块，说明系统将验证什么、用户需要提供什么、失败时如何处理。",
      };
    }

    return { intercepted: false };
  }

  /**
   * 生成重试消息
   */
  private generateRetryMessage(missingBlocks: string[]): string {
    const lines = [
      "未通过放行检查，请补齐缺失的输出块。",
      "",
      "缺少的块：",
      ...missingBlocks.map((b) => `- ${b}`),
      "",
      "请按以下格式重新输出完整的提交包：",
      "",
      "[STATE]",
      "current_state=EXECUTED",
      "next_state=VERIFIED",
      "",
      "[ARTIFACTS]",
      "- type=FORMULA platform=excel target_sheet=Sheet1 target_range=C2 content==公式",
      "",
      "[ACCEPTANCE_TESTS]",
      "1) 测试1描述",
      "2) 测试2描述",
      "3) 测试3描述",
      "",
      "[FALLBACK]",
      "- if 条件 then 处理方式",
      "",
      "[DEPLOY_NOTES]",
      "- protect_ranges: ...",
      "- naming_conventions: ...",
      "",
      "[NEXT_ACTION]",
      "- system_will_validate: ...",
      "- user_needs_to_provide: ...",
      "- if_fail_agent_will: ...",
    ];
    return lines.join("\n");
  }

  /**
   * 生成测试数量不足的消息
   */
  private generateTestsRequiredMessage(currentCount: number): string {
    return `验收测试不足：需要至少 3 条，当前只有 ${currentCount} 条。
请在 [ACCEPTANCE_TESTS] 块中补充更多测试用例，包括：
1) 正常情况测试
2) 边界情况测试（如新增行、空值）
3) 异常情况测试（如数据格式错误）`;
  }
}

// ========== SelfReferenceInterceptor ==========

/**
 * 自引用拦截器
 * 检测到自引用公式时强制重试
 */
export class SelfReferenceInterceptor {
  /**
   * 检查并拦截
   */
  intercept(validationReport: ValidationReport): InterceptResult {
    const selfRefFails = validationReport.criticalFails.filter((v) =>
      v.ruleId.includes("SELF_REFERENCE")
    );

    if (selfRefFails.length > 0) {
      return {
        intercepted: true,
        reason: "检测到自引用公式",
        systemMessage: this.generateSelfRefFixMessage(selfRefFails),
      };
    }

    return { intercepted: false };
  }

  /**
   * 生成修复消息
   */
  private generateSelfRefFixMessage(
    fails: { name: string; reason?: string; details?: Record<string, unknown> }[]
  ): string {
    const lines = [
      "⚠️ 检测到循环引用风险，公式验证未通过。",
      "",
      "问题详情：",
      ...fails.map((f) => `- ${f.reason || f.name}`),
      "",
      "修复要求：",
      "1. 公式不能引用其输出所在的列",
      "2. 如果需要累计计算，请使用辅助列或不同的方法",
      "3. 重新设计公式并提交",
      "",
      "请修改公式后重新提交完整的 [ARTIFACTS] 块。",
    ];
    return lines.join("\n");
  }
}

// ========== MaxIterationsInterceptor ==========

/**
 * 最大迭代次数拦截器
 */
export class MaxIterationsInterceptor {
  /**
   * 检查是否超过最大迭代次数
   */
  intercept(run: AgentRun): InterceptResult {
    if (run.iteration >= run.maxIterations) {
      return {
        intercepted: true,
        reason: `已达到最大迭代次数 (${run.maxIterations})`,
        systemMessage: this.generateMaxIterMessage(run),
      };
    }
    return { intercepted: false };
  }

  /**
   * 生成超限消息
   */
  private generateMaxIterMessage(run: AgentRun): string {
    // 收集未完成的清单项
    const incomplete: string[] = [];
    if (!run.checklist.hasExecutableArtifact) incomplete.push("可执行产物");
    if (!run.checklist.hasPlacementInfo) incomplete.push("放置位置");
    if (!run.checklist.supportsAutoExpand) incomplete.push("自动扩展支持");
    if (!run.checklist.avoidsSelfReference) incomplete.push("避免自引用");
    if (!run.checklist.has3AcceptanceTests) incomplete.push("3条验收测试");
    if (!run.checklist.hasFallbackPlan) incomplete.push("回退方案");
    if (!run.checklist.hasDeployNotes) incomplete.push("部署说明");

    return `已达到最大迭代次数 (${run.maxIterations})，任务无法自动完成。

未完成的项目：
${incomplete.map((i) => `- ${i}`).join("\n")}

请用户提供更多信息或手动完成剩余步骤。`;
  }
}

// ========== 导出 ==========

export const completionInterceptor = new CompletionInterceptor();
export const selfReferenceInterceptor = new SelfReferenceInterceptor();
export const maxIterationsInterceptor = new MaxIterationsInterceptor();
