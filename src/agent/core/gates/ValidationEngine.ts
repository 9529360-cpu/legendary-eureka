/**
 * 验证引擎 - ValidationEngine
 *
 * 职责：
 * 1. 整合所有验证规则
 * 2. 执行系统级验证（非模型自验证）
 * 3. 生成验证报告
 */

import { Submission, Validation, ValidationStatus, Checklist } from "./types";
import { FormulaValidator, formulaValidator } from "./FormulaValidator";

// ========== 验证报告 ==========

/**
 * 验证报告
 */
export interface ValidationReport {
  allPassed: boolean;
  criticalFails: Validation[];
  warnings: Validation[];
  passes: Validation[];
  checklist: Checklist;
  summary: string;
}

// ========== ValidationEngine 类 ==========

/**
 * 验证引擎
 */
export class ValidationEngine {
  private formulaValidator: FormulaValidator;

  constructor() {
    this.formulaValidator = formulaValidator;
  }

  /**
   * 验证提交包（核心方法）
   */
  validate(submission: Submission): ValidationReport {
    const validations: Validation[] = [];

    // 1. 验证产物（公式规则）
    for (const artifact of submission.artifacts) {
      const artifactValidations = this.formulaValidator.validate(artifact);
      validations.push(...artifactValidations);
    }

    // 2. 验证结构完整性
    validations.push(this.validateStructure(submission));

    // 3. 分类结果
    const criticalFails = validations.filter((v) => v.status === ValidationStatus.FAIL);
    const warnings = validations.filter((v) => v.status === ValidationStatus.WARN);
    const passes = validations.filter((v) => v.status === ValidationStatus.PASS);

    // 4. 更新 Checklist
    const checklist = this.computeChecklist(submission, validations);

    // 5. 生成摘要
    const summary = this.generateSummary(criticalFails, warnings, passes);

    return {
      allPassed: criticalFails.length === 0,
      criticalFails,
      warnings,
      passes,
      checklist,
      summary,
    };
  }

  /**
   * 验证结构完整性
   */
  private validateStructure(submission: Submission): Validation {
    const issues: string[] = [];

    // 检查必需字段
    if (!submission.artifacts || submission.artifacts.length === 0) {
      issues.push("缺少可执行产物");
    }

    if (!submission.acceptanceTests || submission.acceptanceTests.length < 3) {
      issues.push("验收测试不足 3 条");
    }

    if (!submission.fallback || submission.fallback.length === 0) {
      issues.push("缺少回退方案");
    }

    if (!submission.deployNotes || Object.keys(submission.deployNotes).length === 0) {
      issues.push("缺少部署说明");
    }

    if (issues.length > 0) {
      return {
        name: "结构完整性检查",
        ruleId: "STRUCTURE_CHECK",
        status: ValidationStatus.FAIL,
        reason: issues.join("；"),
        details: { issues },
      };
    }

    return {
      name: "结构完整性检查",
      ruleId: "STRUCTURE_CHECK",
      status: ValidationStatus.PASS,
    };
  }

  /**
   * 计算 Checklist（基于验证结果）
   */
  private computeChecklist(submission: Submission, validations: Validation[]): Checklist {
    const selfRefCheck = validations.find((v) => v.ruleId === "R2_SELF_REFERENCE");
    const autoExpandCheck = validations.find((v) => v.ruleId === "R3_AUTO_EXPAND");

    return {
      hasExecutableArtifact: submission.artifacts.length > 0,
      hasPlacementInfo: submission.artifacts.every(
        (a) => a.target && (a.target.sheet || a.target.range || a.target.column || a.target.cell)
      ),
      supportsAutoExpand:
        autoExpandCheck?.status === ValidationStatus.PASS ||
        autoExpandCheck?.status === ValidationStatus.WARN,
      avoidsSelfReference: selfRefCheck?.status !== ValidationStatus.FAIL,
      has3AcceptanceTests: (submission.acceptanceTests?.length || 0) >= 3,
      hasFallbackPlan: (submission.fallback?.length || 0) > 0,
      hasDeployNotes: !!submission.deployNotes && Object.keys(submission.deployNotes).length > 0,
    };
  }

  /**
   * 生成验证摘要
   */
  private generateSummary(
    criticalFails: Validation[],
    warnings: Validation[],
    passes: Validation[]
  ): string {
    const lines: string[] = [];

    if (criticalFails.length === 0 && warnings.length === 0) {
      lines.push("✅ 所有验证通过");
    } else {
      if (criticalFails.length > 0) {
        lines.push(`❌ ${criticalFails.length} 个严重问题需要修复：`);
        criticalFails.forEach((v) => lines.push(`   - ${v.name}: ${v.reason}`));
      }

      if (warnings.length > 0) {
        lines.push(`⚠️ ${warnings.length} 个警告：`);
        warnings.forEach((v) => lines.push(`   - ${v.name}: ${v.reason}`));
      }
    }

    lines.push(
      `📊 通过: ${passes.length}, 警告: ${warnings.length}, 失败: ${criticalFails.length}`
    );

    return lines.join("\n");
  }
}

// ========== 导出单例 ==========

export const validationEngine = new ValidationEngine();
