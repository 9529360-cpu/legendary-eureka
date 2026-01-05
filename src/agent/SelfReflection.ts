/**
 * SelfReflection - 自我反思机制
 *
 * 基于 ai-agents-for-beginners 第七课 Metacognition 的学习
 * 实现 Agent 的自我反思能力，验证输出是否合理
 *
 * 主要功能:
 * 1. 验证执行计划的合理性
 * 2. 检查参数完整性和正确性
 * 3. 预测可能的问题
 * 4. 提供改进建议
 *
 * @version 1.0.0
 * @see docs/AI_AGENTS_FOR_BEGINNERS_LEARNING.md
 */

import { Logger } from "../utils/Logger";
import type { Tool } from "./AgentCore";
import type { ExecutionPlan, ExecutionStep } from "./LLMResponseValidator";

// ============================================================================
// 类型定义
// ============================================================================

/**
 * 反思结果
 */
export interface ReflectionResult {
  /** 是否通过验证 */
  isValid: boolean;
  /** 置信度(0-1) */
  confidence: number;
  /** 发现的问题 */
  issues: ReflectionIssue[];
  /** 改进建议 */
  suggestions: string[];
  /** 修正后的计划（如果有）*/
  correctedPlan?: ExecutionPlan;
  /** 反思耗时 (ms) */
  reflectionTime: number;
}

/**
 * 反思发现的问题
 */
export interface ReflectionIssue {
  /** 问题类型 */
  type: "error" | "warning" | "info";
  /** 问题代码 */
  code: string;
  /** 问题描述 */
  message: string;
  /** 相关步骤 */
  stepNumber?: number;
  /** 相关字段 */
  field?: string;
  /** 自动修复建议 */
  autoFix?: {
    description: string;
    apply: () => void;
  };
}

/**
 * 反思配置
 */
export interface ReflectionConfig {
  /** 是否启用严格模式 */
  strictMode: boolean;
  /** 最大步骤数限制 */
  maxSteps: number;
  /** 是否检查工具可用性 */
  checkToolAvailability: boolean;
  /** 是否检查参数类型 */
  checkParameterTypes: boolean;
  /** 是否检查依赖关系 */
  checkDependencies: boolean;
  /** 是否尝试自动修复 */
  autoFix: boolean;
}

/**
 * 硬规则违规
 */
export interface HardRuleViolation {
  /** 规则代码 */
  rule: string;
  /** 严重程度: block = 阻止执行, warn = 建议确认 */
  severity: "block" | "warn";
  /** 违规描述 */
  message: string;
  /** 相关步骤 */
  stepNumber?: number;
  /** 自动修复建议 */
  fix?: {
    setOperation?: "confirm" | "clarify";
    removeStep?: boolean;
  };
}

/**
 * 硬规则验证结果
 */
export interface HardRuleValidationResult {
  /** 是否通过（无阻塞性违规） */
  passed: boolean;
  /** 是否需要用户确认 */
  needsConfirmation: boolean;
  /** 违规列表 */
  violations: HardRuleViolation[];
}

/**
 * 工具注册表接口
 */
export interface ToolRegistry {
  getTool(name: string): Tool | undefined;
  getAllTools(): Tool[];
}

// ============================================================================
// 默认配置
// ============================================================================

const DEFAULT_CONFIG: ReflectionConfig = {
  strictMode: false,
  maxSteps: 20,
  checkToolAvailability: true,
  checkParameterTypes: true,
  checkDependencies: true,
  autoFix: true,
};

// ============================================================================
// 问题代码定义
// ============================================================================

const ISSUE_CODES = {
  // 错误级别
  TOOL_NOT_FOUND: "E001",
  MISSING_REQUIRED_PARAM: "E002",
  INVALID_PARAM_TYPE: "E003",
  CIRCULAR_DEPENDENCY: "E004",
  EMPTY_PLAN: "E005",

  // 警告级别
  TOO_MANY_STEPS: "W001",
  HIGH_RISK_OPERATION: "W002",
  MISSING_OPTIONAL_PARAM: "W003",
  DUPLICATE_STEP: "W004",
  INEFFICIENT_ORDER: "W005",

  // 信息级别
  OPTIMIZABLE_PLAN: "I001",
  ALTERNATIVE_AVAILABLE: "I002",
} as const;

// ============================================================================
// 自我反思类
// ============================================================================

/**
 * 自我反思引擎
 *
 * 验证和改进执行计划
 */
export class SelfReflection {
  private config: ReflectionConfig;
  private toolRegistry: ToolRegistry | null = null;
  private readonly MODULE_NAME = "SelfReflection";

  constructor(config: Partial<ReflectionConfig> = {}) {
    this.config = { ...DEFAULT_CONFIG, ...config };
  }

  /**
   * 设置工具注册表
   */
  setToolRegistry(registry: ToolRegistry): void {
    this.toolRegistry = registry;
  }

  /**
   * 反思执行计划
   */
  reflect(plan: ExecutionPlan): ReflectionResult {
    const startTime = Date.now();
    const issues: ReflectionIssue[] = [];

    Logger.debug(this.MODULE_NAME, "Starting reflection", {
      mainTask: plan.mainTask,
      stepCount: plan.steps.length,
    });

    // 1. 基础验证
    issues.push(...this.validateBasics(plan));

    // 2. 步骤验证
    for (const step of plan.steps) {
      issues.push(...this.validateStep(step, plan));
    }

    // 3. 依赖关系验证
    if (this.config.checkDependencies) {
      issues.push(...this.validateDependencies(plan));
    }

    // 4. 优化检查
    issues.push(...this.checkOptimizations(plan));

    // 5. 计算置信度
    const confidence = this.calculateConfidence(issues);

    // 6. 生成建议
    const suggestions = this.generateSuggestions(issues, plan);

    // 7. 尝试自动修复
    let correctedPlan: ExecutionPlan | undefined;
    if (this.config.autoFix && issues.some((i) => i.autoFix)) {
      correctedPlan = this.applyAutoFixes(plan, issues);
    }

    const result: ReflectionResult = {
      isValid: !issues.some((i) => i.type === "error"),
      confidence,
      issues,
      suggestions,
      correctedPlan,
      reflectionTime: Date.now() - startTime,
    };

    Logger.info(this.MODULE_NAME, "Reflection complete", {
      isValid: result.isValid,
      confidence: result.confidence.toFixed(2),
      issueCount: issues.length,
      errorCount: issues.filter((i) => i.type === "error").length,
    });

    return result;
  }

  /**
   * 快速验证（仅检查关键错误）
   */
  quickValidate(plan: ExecutionPlan): { isValid: boolean; errors: string[] } {
    const errors: string[] = [];

    // 检查空计划
    if (!plan.steps || plan.steps.length === 0) {
      errors.push("执行计划为空");
    }

    // 检查工具可用性
    if (this.toolRegistry) {
      for (const step of plan.steps || []) {
        if (!this.toolRegistry.getTool(step.toolCall.name)) {
          errors.push(`工具 '${step.toolCall.name}' 不存在`);
        }
      }
    }

    // 检查必需参数
    for (const step of plan.steps || []) {
      if (!step.toolCall.name) {
        errors.push(`步骤 ${step.stepNumber} 缺少工具名称`);
      }
    }

    return { isValid: errors.length === 0, errors };
  }

  /**
   * 硬规则验证 - 只检查客观、可判定的规则
   *
   * 设计原则：
   * - 只做安全/可执行性是否需要确认这类硬规则
   * - 不做主观评审（比如"步骤是否足够优雅"）
   * - 返回结果是明确的 pass/fail，不是 confidence 分数
   *
   * @param plan 执行计划
   * @returns 硬规则验证结果
   */
  validateHardRules(plan: ExecutionPlan): HardRuleValidationResult {
    const violations: HardRuleViolation[] = [];

    // ======================================================================
    // 规则 1: 安全性检查
    // ======================================================================

    // 1.1 高风险操作必须有确认
    if (plan.riskLevel === "high" || plan.riskLevel === "critical") {
      if (plan.operation !== "confirm") {
        violations.push({
          rule: "SAFETY_HIGH_RISK_NO_CONFIRM",
          severity: "block",
          message: `高风险操作(${plan.riskLevel}) 必须请求用户确认`,
          fix: { setOperation: "confirm" },
        });
      }
    }

    // 1.2 破坏性操作检查
    const destructiveTools = [
      "excel_delete_sheet",
      "excel_clear_range",
      "excel_delete_rows",
      "excel_delete_columns",
    ];
    for (const step of plan.steps || []) {
      if (destructiveTools.includes(step.toolCall.name)) {
        if (plan.operation !== "confirm") {
          violations.push({
            rule: "SAFETY_DESTRUCTIVE_NO_CONFIRM",
            severity: "block",
            message: `破坏性操作 '${step.toolCall.name}' 必须请求用户确认`,
            stepNumber: step.stepNumber,
            fix: { setOperation: "confirm" },
          });
        }
      }
    }

    // ======================================================================
    // 规则 2: 可执行性检查
    // ======================================================================

    // 2.1 空计划
    if (!plan.steps || plan.steps.length === 0) {
      violations.push({
        rule: "EXEC_EMPTY_PLAN",
        severity: "block",
        message: "执行计划为空",
      });
    }

    // 2.2 工具必须存在
    if (this.toolRegistry) {
      for (const step of plan.steps || []) {
        if (!this.toolRegistry.getTool(step.toolCall.name)) {
          violations.push({
            rule: "EXEC_TOOL_NOT_FOUND",
            severity: "block",
            message: `工具 '${step.toolCall.name}' 不存在`,
            stepNumber: step.stepNumber,
          });
        }
      }
    }

    // 2.3 必需参数必须存在
    if (this.toolRegistry) {
      for (const step of plan.steps || []) {
        const tool = this.toolRegistry.getTool(step.toolCall.name);
        if (tool?.parameters) {
          for (const param of tool.parameters) {
            if (param.required) {
              const value = step.toolCall.parameters?.[param.name];
              if (value === undefined || value === null || value === "") {
                violations.push({
                  rule: "EXEC_MISSING_REQUIRED_PARAM",
                  severity: "block",
                  message: `步骤 ${step.stepNumber}: 缺少必需参数 '${param.name}'`,
                  stepNumber: step.stepNumber,
                });
              }
            }
          }
        }
      }
    }

    // ======================================================================
    // 规则 3: 需要用户确认的场景
    // ======================================================================

    // 3.1 批量操作
    const batchThreshold = 100;
    for (const step of plan.steps || []) {
      const params = step.toolCall.parameters || {};
      // 检查范围是否可能影响大量单元格
      if (params.range && typeof params.range === "string") {
        const rangeParts = params.range.split(":");
        if (rangeParts.length === 2) {
          // 简单判断：如果范围跨越多行多列，建议确认
          const startCol = rangeParts[0].replace(/[0-9]/g, "");
          const endCol = rangeParts[1].replace(/[0-9]/g, "");
          const startRow = parseInt(rangeParts[0].replace(/[A-Z]/gi, ""), 10);
          const endRow = parseInt(rangeParts[1].replace(/[A-Z]/gi, ""), 10);

          if (!isNaN(startRow) && !isNaN(endRow) && endRow - startRow > batchThreshold) {
            violations.push({
              rule: "CONFIRM_BATCH_OPERATION",
              severity: "warn",
              message: `步骤 ${step.stepNumber}: 操作可能影响超过 ${batchThreshold} 行，建议确认`,
              stepNumber: step.stepNumber,
            });
          }

          if (startCol !== endCol && endCol.length > startCol.length) {
            violations.push({
              rule: "CONFIRM_WIDE_RANGE",
              severity: "warn",
              message: `步骤 ${step.stepNumber}: 操作范围较宽，建议确认`,
              stepNumber: step.stepNumber,
            });
          }
        }
      }
    }

    // 计算结果
    const hasBlockingViolation = violations.some((v) => v.severity === "block");
    const needsConfirmation =
      violations.some((v) => v.severity === "warn") && plan.operation !== "confirm";

    Logger.info(this.MODULE_NAME, "Hard rules validation complete", {
      passed: !hasBlockingViolation,
      needsConfirmation,
      violationCount: violations.length,
      blockingCount: violations.filter((v) => v.severity === "block").length,
    });

    return {
      passed: !hasBlockingViolation,
      needsConfirmation,
      violations,
    };
  }

  /**
   * 验证单个步骤
   */
  validateSingleStep(
    step: ExecutionStep,
    context?: { previousSteps?: ExecutionStep[] }
  ): ReflectionIssue[] {
    const issues: ReflectionIssue[] = [];

    // 工具验证
    if (this.toolRegistry && this.config.checkToolAvailability) {
      const tool = this.toolRegistry.getTool(step.toolCall.name);
      if (!tool) {
        issues.push({
          type: "error",
          code: ISSUE_CODES.TOOL_NOT_FOUND,
          message: `工具 '${step.toolCall.name}' 不存在`,
          stepNumber: step.stepNumber,
        });
      } else {
        // 参数验证
        issues.push(...this.validateParameters(step, tool));
      }
    }

    return issues;
  }

  // ============================================================================
  // 私有方法 - 验证
  // ============================================================================

  /**
   * 基础验证
   */
  private validateBasics(plan: ExecutionPlan): ReflectionIssue[] {
    const issues: ReflectionIssue[] = [];

    // 空计划检查
    if (!plan.steps || plan.steps.length === 0) {
      issues.push({
        type: "error",
        code: ISSUE_CODES.EMPTY_PLAN,
        message: "执行计划为空，没有任何步骤",
      });
    }

    // 步骤数量检查
    if (plan.steps && plan.steps.length > this.config.maxSteps) {
      issues.push({
        type: "warning",
        code: ISSUE_CODES.TOO_MANY_STEPS,
        message: `步骤数量 (${plan.steps.length}) 超过建议上限 (${this.config.maxSteps})`,
      });
    }

    // 高风险操作检查
    if (plan.riskLevel === "high" || plan.riskLevel === "critical") {
      issues.push({
        type: "warning",
        code: ISSUE_CODES.HIGH_RISK_OPERATION,
        message: `计划包含${plan.riskLevel === "critical" ? "极" : ""}高风险操作，建议确认后执行`,
      });
    }

    return issues;
  }

  /**
   * 验证步骤
   */
  private validateStep(step: ExecutionStep, _plan: ExecutionPlan): ReflectionIssue[] {
    const issues: ReflectionIssue[] = [];

    // 工具存在性检查
    if (this.toolRegistry && this.config.checkToolAvailability) {
      const tool = this.toolRegistry.getTool(step.toolCall.name);
      if (!tool) {
        issues.push({
          type: "error",
          code: ISSUE_CODES.TOOL_NOT_FOUND,
          message: `步骤 ${step.stepNumber}: 工具 '${step.toolCall.name}' 不存在`,
          stepNumber: step.stepNumber,
          field: "toolCall.name",
        });
      } else {
        // 参数验证
        issues.push(...this.validateParameters(step, tool));
      }
    }

    return issues;
  }

  /**
   * 验证参数
   */
  private validateParameters(step: ExecutionStep, tool: Tool): ReflectionIssue[] {
    const issues: ReflectionIssue[] = [];
    const params = step.toolCall.parameters || {};

    if (!tool.parameters) return issues;

    for (const param of tool.parameters) {
      const value = params[param.name];

      // 必需参数检查
      if (param.required && (value === undefined || value === null)) {
        issues.push({
          type: "error",
          code: ISSUE_CODES.MISSING_REQUIRED_PARAM,
          message: `步骤 ${step.stepNumber}: 缺少必需参数 '${param.name}'`,
          stepNumber: step.stepNumber,
          field: `parameters.${param.name}`,
          autoFix: this.getDefaultValueFix(step, param),
        });
      }

      // 类型检查
      if (this.config.checkParameterTypes && value !== undefined) {
        const typeError = this.checkParameterType(value, param.type);
        if (typeError) {
          issues.push({
            type: "warning",
            code: ISSUE_CODES.INVALID_PARAM_TYPE,
            message: `步骤 ${step.stepNumber}: 参数 '${param.name}' ${typeError}`,
            stepNumber: step.stepNumber,
            field: `parameters.${param.name}`,
          });
        }
      }
    }

    return issues;
  }

  /**
   * 验证依赖关系
   */
  private validateDependencies(plan: ExecutionPlan): ReflectionIssue[] {
    const issues: ReflectionIssue[] = [];

    // 检查步骤顺序是否合理
    const readOps = new Set(["excel_read_cell", "excel_get_range", "excel_get_active_cell"]);
    const writeOps = new Set(["excel_write_cell", "excel_write_range", "excel_set_format"]);

    let lastWriteStep = -1;
    let hasReadAfterWrite = false;

    for (const step of plan.steps) {
      const toolName = step.toolCall.name;

      if (writeOps.has(toolName)) {
        lastWriteStep = step.stepNumber;
      }

      if (readOps.has(toolName) && lastWriteStep > 0 && step.stepNumber > lastWriteStep) {
        // 检查是否在写入后读取同一位置（可能需要）
        hasReadAfterWrite = true;
      }
    }

    if (hasReadAfterWrite) {
      issues.push({
        type: "info",
        code: ISSUE_CODES.OPTIMIZABLE_PLAN,
        message: "检测到写入后读取操作，确保这是必要的验证步骤",
      });
    }

    // 检查重复步骤
    const seen = new Set<string>();
    for (const step of plan.steps) {
      const key = `${step.toolCall.name}:${JSON.stringify(step.toolCall.parameters)}`;
      if (seen.has(key)) {
        issues.push({
          type: "warning",
          code: ISSUE_CODES.DUPLICATE_STEP,
          message: `步骤 ${step.stepNumber}: 可能是重复操作`,
          stepNumber: step.stepNumber,
        });
      }
      seen.add(key);
    }

    return issues;
  }

  /**
   * 检查优化机会
   */
  private checkOptimizations(plan: ExecutionPlan): ReflectionIssue[] {
    const issues: ReflectionIssue[] = [];

    // 检查是否可以合并的操作
    const consecutiveWrites: number[] = [];
    for (let i = 0; i < plan.steps.length - 1; i++) {
      if (
        plan.steps[i].toolCall.name === "excel_write_cell" &&
        plan.steps[i + 1].toolCall.name === "excel_write_cell"
      ) {
        consecutiveWrites.push(i);
      }
    }

    if (consecutiveWrites.length >= 3) {
      issues.push({
        type: "info",
        code: ISSUE_CODES.OPTIMIZABLE_PLAN,
        message: `发现 ${consecutiveWrites.length + 1} 个连续的单元格写入，考虑使用 excel_write_range 批量写入`,
      });
    }

    return issues;
  }

  // ============================================================================
  // 私有方法 - 辅助
  // ============================================================================

  /**
   * 检查参数类型
   */
  private checkParameterType(value: unknown, expectedType?: string): string | null {
    if (!expectedType) return null;

    const actualType = typeof value;

    switch (expectedType.toLowerCase()) {
      case "string":
        if (actualType !== "string") return `期望字符串，实际是 ${actualType}`;
        break;
      case "number":
        if (actualType !== "number" && isNaN(Number(value))) {
          return `期望数字，实际是 ${actualType}`;
        }
        break;
      case "boolean":
        if (actualType !== "boolean") return `期望布尔值，实际是 ${actualType}`;
        break;
      case "array":
        if (!Array.isArray(value)) return `期望数组，实际是 ${actualType}`;
        break;
      case "object":
        if (actualType !== "object" || value === null || Array.isArray(value)) {
          return `期望对象，实际是 ${actualType}`;
        }
        break;
    }

    return null;
  }

  /**
   * 获取默认值修复
   */
  private getDefaultValueFix(
    _step: ExecutionStep,
    param: { name: string; type?: string; default?: unknown }
  ): ReflectionIssue["autoFix"] | undefined {
    if (param.default !== undefined) {
      return {
        description: `使用默认值 '${param.default}'`,
        apply: () => {
          // 修复逻辑在 applyAutoFixes 处理
        },
      };
    }
    return undefined;
  }

  /**
   * 计算置信度
   */
  private calculateConfidence(issues: ReflectionIssue[]): number {
    let confidence = 1.0;

    for (const issue of issues) {
      switch (issue.type) {
        case "error":
          confidence -= 0.3;
          break;
        case "warning":
          confidence -= 0.1;
          break;
        case "info":
          confidence -= 0.02;
          break;
      }
    }

    return Math.max(0, Math.min(1, confidence));
  }

  /**
   * 生成建议
   */
  private generateSuggestions(issues: ReflectionIssue[], _plan: ExecutionPlan): string[] {
    const suggestions: string[] = [];

    // 基于问题生成建议
    const errorCount = issues.filter((i) => i.type === "error").length;
    const warningCount = issues.filter((i) => i.type === "warning").length;

    if (errorCount > 0) {
      suggestions.push(`发现 ${errorCount} 个错误，需要修复后才能执行`);
    }

    if (warningCount > 0) {
      suggestions.push(`发现 ${warningCount} 个警告，建议检查后执行`);
    }

    // 特定问题的建议
    for (const issue of issues) {
      if (issue.code === ISSUE_CODES.TOOL_NOT_FOUND) {
        suggestions.push(`检查工具名称是否正确，可用工具列表请查阅文档`);
        break;
      }
    }

    return [...new Set(suggestions)]; // 去重
  }

  /**
   * 应用自动修复
   */
  private applyAutoFixes(plan: ExecutionPlan, issues: ReflectionIssue[]): ExecutionPlan {
    const corrected = JSON.parse(JSON.stringify(plan)) as ExecutionPlan;

    for (const issue of issues) {
      if (!issue.autoFix) continue;

      Logger.debug(this.MODULE_NAME, "Applying auto-fix", {
        code: issue.code,
        description: issue.autoFix.description,
      });

      // 根据问题类型应用修复
      // 这里可以扩展更多自动修复逻辑
    }

    return corrected;
  }
}

// ============================================================================
// 便捷函数
// ============================================================================

let globalReflection: SelfReflection | null = null;

/**
 * 获取全局反思实例
 */
export function getSelfReflection(): SelfReflection {
  if (!globalReflection) {
    globalReflection = new SelfReflection();
  }
  return globalReflection;
}

/**
 * 创建反思实例
 */
export function createSelfReflection(config?: Partial<ReflectionConfig>): SelfReflection {
  return new SelfReflection(config);
}

/**
 * 快速验证计划
 */
export function validatePlan(plan: ExecutionPlan, toolRegistry?: ToolRegistry): ReflectionResult {
  const reflection = new SelfReflection();
  if (toolRegistry) {
    reflection.setToolRegistry(toolRegistry);
  }
  return reflection.reflect(plan);
}

/**
 * 检查计划是否可执行
 */
export function isPlanExecutable(plan: ExecutionPlan, toolRegistry?: ToolRegistry): boolean {
  const reflection = new SelfReflection({ strictMode: true });
  if (toolRegistry) {
    reflection.setToolRegistry(toolRegistry);
  }
  const result = reflection.quickValidate(plan);
  return result.isValid;
}
