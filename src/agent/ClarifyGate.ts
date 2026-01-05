/**
 * ClarifyGate.ts - 澄清门（P2 澄清机制）
 *
 * v2.9.59: 输出必须是 NextAction 三选一
 *
 * 核心原则：
 * ┌─────────────────────────────────────────────────────┐
 * │  硬规则优先，LLM 兜底                               │
 * │  不靠模型"自觉"，靠代码强制                         │
 * └─────────────────────────────────────────────────────┘
 */

import { NextAction, ClarifyQuestion, Signal, createSignal, SignalCodes } from "./AgentProtocol";
import type { ExecutionPlan } from "./TaskPlanner";

// ========== 类型定义 ==========

/**
 * 工作簿上下文（用于判断是否需要澄清）
 */
export interface WorkbookContext {
  /** 当前选区范围 */
  selectionRange?: string;
  /** 当前活动工作表 */
  activeSheet?: string;
  /** 所有工作表 */
  sheets?: Array<{
    name: string;
    hasData?: boolean;
  }>;
  /** 是否有表格 */
  hasTables?: boolean;
  /** 表格列表 */
  tables?: Array<{
    name: string;
    sheetName: string;
    columns?: string[];
  }>;
}

/**
 * 澄清门配置
 */
export interface ClarifyGateConfig {
  /** 是否启用硬规则 */
  enableHardRules: boolean;
  /** 是否启用 LLM 兜底 */
  enableLLMFallback: boolean;
  /** LLM 置信度阈值 */
  confidenceThreshold: number;
  /** 大范围操作的阈值（超过多少行需要确认） */
  largeRangeThreshold: number;
}

/**
 * 默认配置
 *
 * v2.9.65: 关闭硬规则拦截！
 * Agent 应该先尝试执行，只有真正无法执行时才问用户。
 * 不要在入口处用关键词匹配来阻断，这不智能也不可维护。
 */
export const DEFAULT_CLARIFY_CONFIG: ClarifyGateConfig = {
  enableHardRules: false, // v2.9.65: 关闭！让 Agent 先干活
  enableLLMFallback: false, // v2.9.65: 也关闭，不要在入口拦截
  confidenceThreshold: 0.5, // 降低阈值，更宽容
  largeRangeThreshold: 1000,
};

// ========== 核心类 ==========

/**
 * 澄清门
 *
 * 决定下一步是 clarify / plan / execute
 */
export class ClarifyGate {
  private config: ClarifyGateConfig;

  constructor(config?: Partial<ClarifyGateConfig>) {
    this.config = { ...DEFAULT_CLARIFY_CONFIG, ...config };
  }

  /**
   * 决定下一步动作
   *
   * @param task 用户请求
   * @param workbookCtx 工作簿上下文
   * @param guessedPlan 初步规划（可选）
   * @param llmConfidence LLM 分析的置信度（可选）
   * @returns NextAction 三选一
   */
  decide(
    task: string,
    workbookCtx: WorkbookContext,
    guessedPlan?: ExecutionPlan,
    llmConfidence?: number
  ): NextAction {
    const signals: Signal[] = [];

    // ========== v2.9.63: 快速通道 - 明确输入直接跳过澄清 ==========
    // 如果用户已经明确给了范围（如 A1:D1）或工作表名，就不要再问了
    if (this.canSkipClarify(task, workbookCtx)) {
      console.log("[ClarifyGate] ✅ 快速通道：用户输入足够明确，跳过澄清");
      return {
        kind: "plan",
        plan: this.createEmptyPlan(task),
        signals: [createSignal("info", "SKIP_CLARIFY", "用户输入明确，无需澄清")],
      };
    }

    // ========== 第一层：硬规则检查 ==========
    if (this.config.enableHardRules) {
      const clarifyResult = this.checkHardRules(task, workbookCtx, guessedPlan);
      if (clarifyResult) {
        signals.push(...clarifyResult.signals);
        return {
          kind: "clarify",
          questions: clarifyResult.questions,
          signals,
        };
      }
    }

    // ========== 第二层：LLM 置信度检查 ==========
    if (this.config.enableLLMFallback && llmConfidence !== undefined) {
      if (llmConfidence < this.config.confidenceThreshold) {
        signals.push(
          createSignal(
            "warning",
            "LOW_CONFIDENCE",
            `LLM 置信度过低 (${(llmConfidence * 100).toFixed(0)}%)`,
            { evidence: { confidence: llmConfidence } }
          )
        );

        // 生成通用澄清问题
        const questions = this.generateLowConfidenceQuestions(task, workbookCtx);
        return {
          kind: "clarify",
          questions,
          signals,
        };
      }
    }

    // ========== 第三层：有计划就执行，没计划就规划 ==========
    if (guessedPlan && guessedPlan.steps.length > 0) {
      // 检查是否需要预览确认（大范围写操作）
      const needsConfirmation = this.checkNeedsConfirmation(guessedPlan);
      if (needsConfirmation) {
        signals.push(createSignal("info", "LARGE_OPERATION", "大范围操作，建议确认后执行"));
      }

      return {
        kind: "execute",
        plan: guessedPlan,
        signals,
      };
    }

    // 没有计划，需要规划
    return {
      kind: "plan",
      plan: this.createEmptyPlan(task),
      signals,
    };
  }

  // ========== 硬规则检查 ==========

  /**
   * 检查硬规则（必须写死在代码里，不靠模型自觉）
   */
  private checkHardRules(
    task: string,
    workbookCtx: WorkbookContext,
    guessedPlan?: ExecutionPlan
  ): { questions: ClarifyQuestion[]; signals: Signal[] } | null {
    const questions: ClarifyQuestion[] = [];
    const signals: Signal[] = [];

    // 规则 1：写操作必须知道 sheet + range
    if (guessedPlan?.steps?.some((s) => s.isWriteOperation)) {
      const unknownTargets = guessedPlan.steps.filter(
        (s) =>
          s.isWriteOperation &&
          (!s.parameters?.sheet || !s.parameters?.range) &&
          !s.parameters?.address
      );
      if (unknownTargets.length > 0) {
        signals.push(
          createSignal(
            "warning",
            SignalCodes.MISSING_TARGET_RANGE,
            `${unknownTargets.length} 个写操作缺少目标范围`,
            { evidence: unknownTargets.map((s) => s.action) }
          )
        );
        questions.push(this.createStandardQuestion("range"), this.createStandardQuestion("sheet"));
      }
    }

    // 规则 2：写意图 + 没选区 + 请求文本中也没有指定范围
    const writeIntent = this.hasWriteIntent(task);
    const hasSelection = !!workbookCtx.selectionRange;
    const hasRangeInTask = this.hasRangeInText(task); // v2.9.62: 检查请求文本中是否有范围
    if (writeIntent && !hasSelection && !hasRangeInTask) {
      signals.push(
        createSignal("warning", SignalCodes.MISSING_TARGET_RANGE, "写操作但没有选区或指定范围")
      );
      if (!questions.some((q) => q.id === "range")) {
        questions.push(this.createStandardQuestion("range"));
      }
    }

    // 规则 3：多 sheet + 模糊引用（v2.9.63: 只在真正模糊时触发）
    // 如果用户说"这张表"但工作簿有多个表且没有活动表，才需要问
    if (
      workbookCtx.sheets &&
      workbookCtx.sheets.length > 1 &&
      this.hasAmbiguousSheetReference(task) &&
      !workbookCtx.activeSheet &&
      !this.hasExplicitSheetName(task) // 如果有明确名称就不问
    ) {
      signals.push(
        createSignal("warning", SignalCodes.MISSING_SHEET_NAME, "多个工作表但引用不明确")
      );
      if (!questions.some((q) => q.id === "sheet")) {
        questions.push(this.createSheetQuestion(workbookCtx.sheets.map((s) => s.name)));
      }
    }

    // 规则 4：仅对高危操作需要确认（v2.9.63: 收紧条件）
    // 只有"删除整表"、"清空全部"这种不可逆大范围操作才需要确认
    if (this.isHighRiskDestructive(task)) {
      signals.push(
        createSignal("warning", SignalCodes.OVERWRITE_CONFIRMATION_NEEDED, "高危操作需要确认")
      );
      questions.push(this.createStandardQuestion("overwrite"));
    }

    if (questions.length > 0) {
      return { questions, signals };
    }

    return null;
  }

  // ========== 意图检测 ==========

  /**
   * 检测是否有写意图
   */
  private hasWriteIntent(task: string): boolean {
    const writePatterns =
      /写入|覆盖|清空|删除|填充|替换|设置公式|加边框|格式|生成表|创建|添加|插入|修改|更新/i;
    return writePatterns.test(task);
  }

  /**
   * 检测是否有模糊的工作表引用
   */
  private hasAmbiguousSheetReference(task: string): boolean {
    const ambiguousPatterns = /这张表|该表|当前表|这个表|那张表/i;
    return ambiguousPatterns.test(task);
  }

  /**
   * 检测是否有破坏性意图（旧版，保留用于兼容）
   */
  private hasDestructiveIntent(task: string): boolean {
    const destructivePatterns = /清空|删除|覆盖|替换全部|清除/i;
    return destructivePatterns.test(task);
  }

  /**
   * v2.9.63: 检测是否有明确的工作表名称
   * 匹配：「测试表」、"Sheet1"、叫测试表 等
   */
  private hasExplicitSheetName(task: string): boolean {
    // 用引号/书名号包裹的名称
    const quotedPattern = /[「『"']([^」』"']+)[」』"']/;
    // "叫xxx" 模式
    const namedPattern = /叫[「『"']?([^」』"'\s，,。]+)/;
    // 明确的工作表名（Sheet1、表1 等）
    const explicitPattern = /Sheet\s*\d+|表\s*\d+/i;

    return quotedPattern.test(task) || namedPattern.test(task) || explicitPattern.test(task);
  }

  /**
   * v2.9.63: 检测是否是高危破坏性操作（需要确认）
   *
   * 只有以下情况才需要确认：
   * - 删除整个工作表
   * - 清空整个工作表/表格
   * - 删除全部数据
   *
   * 普通的覆盖、替换、删除行等不需要确认
   */
  private isHighRiskDestructive(task: string): boolean {
    // 高危模式：删除/清空 + 整表/全部/工作表
    const highRiskPatterns = [
      /删除.*(整|全部|所有).*(表|sheet)/i,
      /清空.*(整|全部|所有).*(表|sheet)/i,
      /(整|全部|所有).*(删除|清空)/i,
      /删除.*工作表/i,
      /移除.*工作表/i,
    ];

    return highRiskPatterns.some((pattern) => pattern.test(task));
  }

  /**
   * v2.9.62: 检测请求文本中是否包含范围引用
   * 例如：A1:D1, B2, C列, 第3行 等
   */
  private hasRangeInText(task: string): boolean {
    // 匹配各种范围格式
    const rangePatterns = [
      /[A-Z]+\d+:[A-Z]+\d+/i, // A1:D1, B2:C10
      /[A-Z]+\d+/i, // A1, B2 (单个单元格)
      /[A-Z]+列/i, // A列, B列
      /第\d+行/i, // 第1行, 第2行
      /\d+行\d+列/i, // 1行1列
      /整[表列行]/i, // 整表, 整列, 整行
      /全部|所有/i, // 全部, 所有 (隐含范围)
    ];
    return rangePatterns.some((pattern) => pattern.test(task));
  }

  /**
   * v2.9.63: 判断是否可以跳过澄清（快速通道）
   *
   * 跳过条件（任一满足即可）：
   * 1. 用户明确指定了范围（如 A1:D1）
   * 2. 用户明确指定了工作表名（如「测试表」）
   * 3. 是创建型任务（新建工作表、创建表格等）
   * 4. 有当前选区作为隐含目标
   * 5. v2.9.64: 是后续操作（包含"刚才"、"重新"、"继续"等词）
   * 6. v2.9.64: 包含明确的列名引用（如"单价"、"销售额"等）
   */
  private canSkipClarify(task: string, workbookCtx: WorkbookContext): boolean {
    // 条件 1：有明确范围
    const hasExplicitRange = this.hasRangeInText(task);

    // 条件 2：有明确工作表名（用引号或书名号包裹的名称）
    const hasExplicitSheetName = /[「『"']([^」』"']+)[」』"']|叫[「『"']?([^」』"'\s]+)/.test(
      task
    );

    // 条件 3：是创建型任务（新建/创建 + 工作表/表格）
    const isCreateTask = /新建|创建/.test(task) && /工作表|表格|sheet/i.test(task);

    // 条件 4：有当前选区
    const hasSelection = !!workbookCtx.selectionRange;

    // 条件 5：v2.9.64 是后续操作（follow-up）
    const isFollowUp = this.isFollowUpRequest(task);

    // 条件 6：v2.9.64 包含列名引用（可以从工作簿推断范围）
    const hasColumnReference = this.hasColumnNameReference(task, workbookCtx);

    // 任一条件满足，就可以跳过澄清
    if (
      hasExplicitRange ||
      hasExplicitSheetName ||
      isCreateTask ||
      hasSelection ||
      isFollowUp ||
      hasColumnReference
    ) {
      return true;
    }

    return false;
  }

  /**
   * v2.9.64: 检测是否是后续操作（follow-up request）
   *
   * 包含这些词的请求通常是对前一个操作的延续，不需要重新澄清
   */
  private isFollowUpRequest(task: string): boolean {
    const followUpPatterns = [
      /刚才/,
      /重新/,
      /继续/,
      /再来/,
      /再试/,
      /上次/,
      /之前/,
      /接着/,
      /然后/,
      /现在/,
      /这次/,
      /改成/,
      /换成/,
      /修复/,
      /修正/,
      /改一下/,
      /试试/,
    ];
    return followUpPatterns.some((p) => p.test(task));
  }

  /**
   * v2.9.64: 检测是否包含列名引用
   *
   * 如果用户提到了工作簿中存在的列名，Agent 可以推断操作范围
   */
  private hasColumnNameReference(task: string, workbookCtx: WorkbookContext): boolean {
    // 常见列名关键词（不依赖工作簿状态也能识别）
    const commonColumnPatterns = [
      /日期/,
      /时间/,
      /产品/,
      /商品/,
      /名称/,
      /数量/,
      /单价/,
      /价格/,
      /金额/,
      /销售额/,
      /总计/,
      /合计/,
      /备注/,
      /编号/,
      /序号/,
    ];

    // 检查是否匹配常见列名
    if (commonColumnPatterns.some((p) => p.test(task))) {
      return true;
    }

    // 如果工作簿有表格信息，检查是否匹配实际列名
    if (workbookCtx.tables) {
      for (const table of workbookCtx.tables) {
        if (table.columns) {
          for (const col of table.columns) {
            if (task.includes(col)) {
              return true;
            }
          }
        }
      }
    }

    return false;
  }

  // ========== 标准问题生成 ==========

  /**
   * 创建标准澄清问题
   */
  private createStandardQuestion(type: "sheet" | "range" | "overwrite"): ClarifyQuestion {
    switch (type) {
      case "sheet":
        return {
          id: "sheet",
          question: "你要操作的 **工作表** 是哪个？",
          required: true,
          defaultValue: "当前活动表",
        };
      case "range":
        return {
          id: "range",
          question: "操作的 **范围** 是？（例如 A1:D100 或 整表）",
          required: true,
          defaultValue: "当前选区",
        };
      case "overwrite":
        return {
          id: "overwrite",
          question: "此操作会 **覆盖原数据**，是否确认？",
          required: true,
          options: ["确认覆盖", "不覆盖，另存"],
          defaultValue: "确认覆盖",
        };
      default:
        return {
          id: "unknown",
          question: "请提供更多信息",
          required: false,
        };
    }
  }

  /**
   * 创建工作表选择问题
   */
  private createSheetQuestion(sheetNames: string[]): ClarifyQuestion {
    return {
      id: "sheet",
      question: "你要操作的 **工作表** 是哪个？",
      required: true,
      options: sheetNames,
      defaultValue: sheetNames[0],
    };
  }

  /**
   * 生成低置信度时的通用问题
   */
  private generateLowConfidenceQuestions(
    task: string,
    workbookCtx: WorkbookContext
  ): ClarifyQuestion[] {
    const questions: ClarifyQuestion[] = [];

    // 如果任务很短，问意图
    if (task.length < 10) {
      questions.push({
        id: "intent",
        question: "能具体说一下你想做什么吗？",
        required: true,
      });
    }

    // 如果有写意图但没选区
    if (this.hasWriteIntent(task) && !workbookCtx.selectionRange) {
      questions.push(this.createStandardQuestion("range"));
    }

    // 如果多 sheet
    if (workbookCtx.sheets && workbookCtx.sheets.length > 1) {
      questions.push(this.createSheetQuestion(workbookCtx.sheets.map((s) => s.name)));
    }

    // 至少问一个问题
    if (questions.length === 0) {
      questions.push({
        id: "clarify",
        question: "我不太确定你的意思，能说得更具体一些吗？",
        required: true,
      });
    }

    return questions;
  }

  // ========== 辅助方法 ==========

  /**
   * 检查是否需要确认（大范围操作）
   */
  private checkNeedsConfirmation(plan: ExecutionPlan): boolean {
    for (const step of plan.steps) {
      if (step.isWriteOperation && step.parameters?.range) {
        const range = String(step.parameters.range);
        const rowCount = this.estimateRowCount(range);
        if (rowCount > this.config.largeRangeThreshold) {
          return true;
        }
      }
    }
    return false;
  }

  /**
   * 估算范围的行数
   */
  private estimateRowCount(range: string): number {
    const match = range.match(/[A-Z]+(\d+):[A-Z]+(\d+)/i);
    if (match) {
      const startRow = parseInt(match[1], 10);
      const endRow = parseInt(match[2], 10);
      return endRow - startRow + 1;
    }
    return 0;
  }

  /**
   * 创建空计划（用于返回 kind: "plan"）
   */
  private createEmptyPlan(task: string): ExecutionPlan {
    return {
      id: `plan_${Date.now()}`,
      taskDescription: task,
      taskType: "unknown",
      steps: [],
      taskSuccessConditions: [],
      dependencyCheck: {
        passed: true,
        missingDependencies: [],
        circularDependencies: [],
        warnings: [],
        semanticDependencies: [],
        unresolvedSemanticDeps: [],
      },
      fieldStepIdMap: {},
    } as unknown as ExecutionPlan;
  }
}

// ========== 单例导出 ==========

export const clarifyGate = new ClarifyGate();

/**
 * 便捷函数：检查是否需要澄清
 */
export function needClarify(
  task: string,
  workbookCtx: WorkbookContext,
  guessedPlan?: ExecutionPlan
): boolean {
  const result = clarifyGate.decide(task, workbookCtx, guessedPlan);
  return result.kind === "clarify";
}

/**
 * 便捷函数：获取下一步动作
 */
export function getNextAction(
  task: string,
  workbookCtx: WorkbookContext,
  guessedPlan?: ExecutionPlan,
  llmConfidence?: number
): NextAction {
  return clarifyGate.decide(task, workbookCtx, guessedPlan, llmConfidence);
}
