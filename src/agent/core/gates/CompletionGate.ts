/**
 * 完成门槛 - CompletionGate
 *
 * 职责：
 * 1. 检查模型输出是否满足完成条件
 * 2. 模型无权自行决定"完成"
 * 3. 只有系统验证通过才放行
 *
 * 核心原则：完成权不在模型，在系统
 */

import {
  AgentRun,
  Checklist,
  Submission,
  Validation,
  ValidationStatus,
  isChecklistComplete,
  Artifact,
  AcceptanceTest,
} from "./types";

// ========== 门槛检查结果 ==========

/**
 * 门槛检查结果
 */
export interface GateCheckResult {
  passed: boolean;
  checklist: Checklist;
  validations: Validation[];
  failReasons: string[];
  requiredActions: string[];
}

// ========== CompletionGate 类 ==========

/**
 * 完成门槛检查器
 */
export class CompletionGate {
  /**
   * 检查是否可以放行（核心方法）
   */
  static check(run: AgentRun, submission: Submission): GateCheckResult {
    const validations: Validation[] = [];
    const failReasons: string[] = [];
    const requiredActions: string[] = [];

    // 1. 计算 Checklist
    const checklist = this.computeChecklist(submission);

    // 2. 检查 Artifacts
    const artifactValidation = this.validateArtifacts(submission.artifacts);
    validations.push(artifactValidation);
    if (artifactValidation.status === ValidationStatus.FAIL) {
      failReasons.push(artifactValidation.reason || "产物验证失败");
      requiredActions.push("提供可执行的公式或步骤，并明确放置位置");
    }

    // 3. 检查验收测试
    const testsValidation = this.validateAcceptanceTests(submission.acceptanceTests);
    validations.push(testsValidation);
    if (testsValidation.status === ValidationStatus.FAIL) {
      failReasons.push(testsValidation.reason || "验收测试不足");
      requiredActions.push("提供至少 3 条验收测试");
    }

    // 4. 检查回退方案
    const fallbackValidation = this.validateFallback(submission.fallback);
    validations.push(fallbackValidation);
    if (fallbackValidation.status === ValidationStatus.FAIL) {
      failReasons.push(fallbackValidation.reason || "缺少回退方案");
      requiredActions.push("提供失败时的回退方案");
    }

    // 5. 检查部署说明
    const deployValidation = this.validateDeployNotes(submission.deployNotes);
    validations.push(deployValidation);
    if (deployValidation.status === ValidationStatus.FAIL) {
      failReasons.push(deployValidation.reason || "缺少部署说明");
      requiredActions.push("提供部署与防错清单");
    }

    // 6. 检查 Checklist 完整性
    const checklistComplete = isChecklistComplete(checklist);

    return {
      passed: checklistComplete && validations.every((v) => v.status !== ValidationStatus.FAIL),
      checklist,
      validations,
      failReasons,
      requiredActions,
    };
  }

  /**
   * 计算 Checklist
   */
  static computeChecklist(submission: Submission): Checklist {
    const artifacts = submission.artifacts || [];
    const tests = submission.acceptanceTests || [];
    const fallback = submission.fallback || [];
    const deployNotes = submission.deployNotes;

    return {
      // 是否有可执行产物
      hasExecutableArtifact: artifacts.length > 0,
      // 是否有位置信息
      hasPlacementInfo: artifacts.every(
        (a) => a.target && (a.target.sheet || a.target.range || a.target.column || a.target.cell)
      ),
      // 是否支持自动扩展（待验证时设为 false，验证通过后置为 true）
      supportsAutoExpand: false,
      // 是否避免自引用（待验证时设为 false，验证通过后置为 true）
      avoidsSelfReference: false,
      // 是否有 3 条以上验收测试
      has3AcceptanceTests: tests.length >= 3,
      // 是否有回退方案
      hasFallbackPlan: fallback.length > 0,
      // 是否有部署说明
      hasDeployNotes: !!deployNotes && Object.keys(deployNotes).length > 0,
    };
  }

  /**
   * 验证产物
   */
  private static validateArtifacts(artifacts: Artifact[]): Validation {
    if (!artifacts || artifacts.length === 0) {
      return {
        name: "产物检查",
        ruleId: "R1_ARTIFACT_REQUIRED",
        status: ValidationStatus.FAIL,
        reason: "未提供可执行产物（公式/步骤/模板）",
      };
    }

    // 检查每个产物是否有位置信息
    const missingTarget = artifacts.filter(
      (a) => !a.target || (!a.target.sheet && !a.target.range && !a.target.column && !a.target.cell)
    );

    if (missingTarget.length > 0) {
      return {
        name: "产物检查",
        ruleId: "R1_PLACEMENT_REQUIRED",
        status: ValidationStatus.FAIL,
        reason: `${missingTarget.length} 个产物缺少放置位置信息`,
      };
    }

    return {
      name: "产物检查",
      ruleId: "R1_ARTIFACT_OK",
      status: ValidationStatus.PASS,
    };
  }

  /**
   * 验证验收测试
   */
  private static validateAcceptanceTests(tests: AcceptanceTest[]): Validation {
    if (!tests || tests.length < 3) {
      return {
        name: "验收测试检查",
        ruleId: "R4_3_TESTS_REQUIRED",
        status: ValidationStatus.FAIL,
        reason: `验收测试不足：需要至少 3 条，当前 ${tests?.length || 0} 条`,
      };
    }

    return {
      name: "验收测试检查",
      ruleId: "R4_TESTS_OK",
      status: ValidationStatus.PASS,
    };
  }

  /**
   * 验证回退方案
   */
  private static validateFallback(fallback: { condition: string; action: string }[]): Validation {
    if (!fallback || fallback.length === 0) {
      return {
        name: "回退方案检查",
        ruleId: "R5_FALLBACK_REQUIRED",
        status: ValidationStatus.FAIL,
        reason: "未提供失败回退方案",
      };
    }

    return {
      name: "回退方案检查",
      ruleId: "R5_FALLBACK_OK",
      status: ValidationStatus.PASS,
    };
  }

  /**
   * 验证部署说明
   */
  private static validateDeployNotes(
    notes:
      | { protectedRanges?: string[]; namingConventions?: string[]; permissions?: string[] }
      | undefined
  ): Validation {
    if (!notes || Object.keys(notes).length === 0) {
      return {
        name: "部署说明检查",
        ruleId: "R6_DEPLOY_NOTES_REQUIRED",
        status: ValidationStatus.FAIL,
        reason: "未提供部署与防错清单",
      };
    }

    return {
      name: "部署说明检查",
      ruleId: "R6_DEPLOY_NOTES_OK",
      status: ValidationStatus.PASS,
    };
  }

  /**
   * 生成强制继续的系统消息
   */
  static generateForceContinueMessage(result: GateCheckResult): string {
    const lines = [
      "你未通过上线放行。禁止结束。必须修复并重新提交。",
      "",
      "失败原因：",
      ...result.failReasons.map((r, i) => `${i + 1}. ${r}`),
      "",
      "必须完成的操作：",
      ...result.requiredActions.map((a, i) => `${i + 1}. ${a}`),
      "",
      "现在进入修复回合：给出修复后的提交包（含：artifact + placement + tests + fallback + deploy_notes）。",
    ];
    return lines.join("\n");
  }
}

// ========== 导出 ==========

export const completionGate = new CompletionGate();
