/**
 * IntentAnalyzer - 意图分析器 v2.9.58
 *
 * P2 核心组件：分析用户请求的置信度，识别需要澄清的场景
 *
 * 核心职责：
 * 1. 评估用户请求的清晰度（confidence score）
 * 2. 识别模糊指代、缺失信息、歧义表达
 * 3. 生成针对性的澄清问题
 * 4. 提供多个可选方案
 *
 * 设计理念：
 * - "不确定就问"，而非"猜测后执行"
 * - 澄清是智能的体现，不是无能的表现
 */

import { DataModel } from "./DataModeler";

// ========== 类型定义 ==========

/**
 * 意图分析结果
 */
export interface IntentAnalysis {
  /** 意图置信度 (0-1) */
  confidence: number;

  /** 识别出的意图类型 */
  intentType: IntentType;

  /** 解析出的实体信息 */
  entities: ParsedEntities;

  /** 需要澄清的问题列表 */
  clarificationNeeded: ClarificationItem[];

  /** 模糊/缺失的信息 */
  ambiguities: AmbiguityInfo[];

  /** 是否可以直接执行（不需要澄清） */
  canProceed: boolean;

  /** 如果需要澄清，推荐的澄清问题 */
  suggestedClarification?: SuggestedClarification;

  /** 如果可以执行，推荐的执行方案 */
  suggestedPlans?: SuggestedPlan[];

  /** 风险评估 */
  riskLevel: "low" | "medium" | "high";

  /** 分析过程的说明 */
  analysisNotes: string[];
}

/**
 * 意图类型
 */
export type IntentType =
  | "query" // 查询/读取
  | "calculate" // 计算/公式
  | "format" // 格式化/美化
  | "modify" // 修改数据
  | "delete" // 删除数据
  | "create" // 创建新内容
  | "analyze" // 分析数据
  | "chart" // 图表操作
  | "sort_filter" // 排序筛选
  | "unclear"; // 无法判断

/**
 * 解析出的实体信息
 */
export interface ParsedEntities {
  /** 工作表名称 */
  sheetNames: string[];

  /** 范围/区域 */
  ranges: string[];

  /** 列名/字段名 */
  columnNames: string[];

  /** 数值 */
  numbers: number[];

  /** 条件/规则 */
  conditions: string[];

  /** 公式 */
  formulas: string[];

  /** 模糊指代词 */
  vaguePronous: VagueReference[];
}

/**
 * 模糊指代
 */
export interface VagueReference {
  /** 原文 */
  text: string;
  /** 类型 */
  type: "this" | "that" | "here" | "there" | "it" | "them" | "above" | "below";
  /** 可能指代的内容 */
  possibleReferences: string[];
}

/**
 * 需要澄清的项
 */
export interface ClarificationItem {
  /** 问题类型 */
  type: ClarificationType;
  /** 缺失/模糊的内容 */
  missing: string;
  /** 为什么需要澄清 */
  reason: string;
  /** 重要程度 */
  importance: "critical" | "important" | "optional";
  /** 可能的选项 */
  options?: string[];
}

export type ClarificationType =
  | "missing_sheet" // 未指定工作表
  | "missing_range" // 未指定范围
  | "missing_column" // 未指定列
  | "vague_reference" // 模糊指代
  | "ambiguous_intent" // 歧义意图
  | "incomplete_condition" // 条件不完整
  | "risky_operation" // 高风险操作需确认
  | "large_scope"; // 影响范围过大

/**
 * 歧义信息
 */
export interface AmbiguityInfo {
  /** 原文 */
  text: string;
  /** 可能的解释 */
  interpretations: string[];
  /** 推荐的解释 */
  recommended?: string;
}

/**
 * 建议的澄清问题
 */
export interface SuggestedClarification {
  /** 主问题 */
  mainQuestion: string;
  /** 补充说明 */
  context?: string;
  /** 选项（如果是选择题） */
  options?: ClarificationOption[];
  /** 是否允许自由回答 */
  allowFreeform: boolean;
}

export interface ClarificationOption {
  /** 选项标识 */
  id: string;
  /** 显示文本 */
  label: string;
  /** 选择后的效果说明 */
  description?: string;
  /** 是否推荐 */
  recommended?: boolean;
}

/**
 * 建议的执行方案
 */
export interface SuggestedPlan {
  /** 方案 ID */
  id: string;
  /** 方案名称 */
  name: string;
  /** 方案描述 */
  description: string;
  /** 预估影响 */
  impact: string;
  /** 置信度 */
  confidence: number;
  /** 是否推荐 */
  recommended: boolean;
}

/**
 * 分析上下文
 */
export interface AnalysisContext {
  /** 用户原始请求 */
  userRequest: string;
  /** 当前工作簿的数据模型 */
  dataModel?: DataModel;
  /** 当前选中的范围 */
  currentSelection?: string;
  /** 当前活动的工作表 */
  activeSheet?: string;
  /** 对话历史 */
  conversationHistory?: Array<{ role: "user" | "assistant"; content: string }>;
  /** 交互配置 */
  clarificationThreshold?: number;
}

// ========== IntentAnalyzer 类 ==========

/**
 * 意图分析器
 */
export class IntentAnalyzer {
  /**
   * 分析用户意图
   */
  analyze(context: AnalysisContext): IntentAnalysis {
    const { userRequest, dataModel, currentSelection, activeSheet } = context;
    const threshold = context.clarificationThreshold ?? 0.7;

    const analysisNotes: string[] = [];

    // 1. 解析实体
    const entities = this.parseEntities(userRequest, dataModel);
    analysisNotes.push(
      `解析到 ${entities.sheetNames.length} 个工作表, ${entities.columnNames.length} 个列名`
    );

    // 2. 识别意图类型
    const intentType = this.classifyIntent(userRequest);
    analysisNotes.push(`识别意图类型: ${intentType}`);

    // 3. 检测模糊/缺失信息
    const clarificationNeeded = this.detectClarificationNeeds(
      userRequest,
      entities,
      intentType,
      dataModel,
      currentSelection,
      activeSheet
    );

    // 4. 检测歧义
    const ambiguities = this.detectAmbiguities(userRequest, entities, dataModel);

    // 5. 计算置信度
    const confidence = this.calculateConfidence(
      userRequest,
      entities,
      intentType,
      clarificationNeeded,
      ambiguities
    );
    analysisNotes.push(`置信度: ${(confidence * 100).toFixed(0)}%`);

    // 6. 评估风险
    const riskLevel = this.assessRisk(intentType, entities, dataModel);

    // 7. 决定是否可以直接执行
    const canProceed =
      confidence >= threshold &&
      clarificationNeeded.filter((c) => c.importance === "critical").length === 0;

    // 8. 生成澄清问题或执行方案
    let suggestedClarification: SuggestedClarification | undefined;
    let suggestedPlans: SuggestedPlan[] | undefined;

    if (!canProceed) {
      suggestedClarification = this.generateClarificationQuestion(
        userRequest,
        clarificationNeeded,
        ambiguities,
        dataModel
      );
      analysisNotes.push("需要澄清后才能执行");
    } else {
      suggestedPlans = this.generateSuggestedPlans(userRequest, entities, intentType, dataModel);
      if (suggestedPlans.length > 1) {
        analysisNotes.push(`生成 ${suggestedPlans.length} 个可选方案`);
      }
    }

    return {
      confidence,
      intentType,
      entities,
      clarificationNeeded,
      ambiguities,
      canProceed,
      suggestedClarification,
      suggestedPlans,
      riskLevel,
      analysisNotes,
    };
  }

  /**
   * 解析实体
   */
  private parseEntities(request: string, dataModel?: DataModel): ParsedEntities {
    const entities: ParsedEntities = {
      sheetNames: [],
      ranges: [],
      columnNames: [],
      numbers: [],
      conditions: [],
      formulas: [],
      vaguePronous: [],
    };

    // 提取工作表名称
    const sheetPatterns = [
      /(?:在|从|到|工作表|表|Sheet)\s*[""']?([^""'\s,，]+)[""']?/gi,
      /'([^']+)'!/g,
      /Sheet(\d+)/gi,
    ];
    for (const pattern of sheetPatterns) {
      let match;
      while ((match = pattern.exec(request)) !== null) {
        const sheetName = match[1];
        if (sheetName && !entities.sheetNames.includes(sheetName)) {
          entities.sheetNames.push(sheetName);
        }
      }
    }

    // 提取范围
    const rangePatterns = [
      /([A-Z]+\d+:[A-Z]+\d+)/gi,
      /([A-Z]+\d+)/gi,
      /([A-Z]+)列/gi,
      /第(\d+)行/gi,
    ];
    for (const pattern of rangePatterns) {
      let match;
      while ((match = pattern.exec(request)) !== null) {
        const range = match[1];
        if (range && !entities.ranges.includes(range)) {
          entities.ranges.push(range);
        }
      }
    }

    // 提取列名（从数据模型匹配）
    if (dataModel) {
      for (const table of dataModel.tables) {
        for (const field of table.fields) {
          if (request.includes(field.name)) {
            if (!entities.columnNames.includes(field.name)) {
              entities.columnNames.push(field.name);
            }
          }
        }
      }
    }

    // 提取数值
    const numbers = request.match(/\d+(\.\d+)?/g);
    if (numbers) {
      entities.numbers = numbers.map(Number);
    }

    // 提取模糊指代
    const vaguePatterns: Array<{ pattern: RegExp; type: VagueReference["type"] }> = [
      { pattern: /这[里个些列行格表]/g, type: "this" },
      { pattern: /那[里个些列行格表]/g, type: "that" },
      { pattern: /这里/g, type: "here" },
      { pattern: /那里/g, type: "there" },
      { pattern: /它们?/g, type: "it" },
      { pattern: /上面/g, type: "above" },
      { pattern: /下面/g, type: "below" },
    ];

    for (const { pattern, type } of vaguePatterns) {
      let match;
      while ((match = pattern.exec(request)) !== null) {
        entities.vaguePronous.push({
          text: match[0],
          type,
          possibleReferences: [], // 后续填充
        });
      }
    }

    // 提取公式
    const formulaMatch = request.match(/=.+?(?=[\s,，。]|$)/g);
    if (formulaMatch) {
      entities.formulas = formulaMatch;
    }

    return entities;
  }

  /**
   * 分类意图
   */
  private classifyIntent(request: string): IntentType {
    const lowerRequest = request.toLowerCase();

    // 删除类
    if (/删除|清空|移除|去掉/.test(request)) {
      return "delete";
    }

    // 查询类
    if (/查看|显示|告诉我|是什么|有多少|统计|查找|搜索/.test(request)) {
      return "query";
    }

    // 计算类
    if (/计算|求和|平均|公式|sum|average|vlookup|加|减|乘|除/.test(lowerRequest)) {
      return "calculate";
    }

    // 格式化类
    if (/格式|美化|颜色|字体|边框|对齐|样式|加粗|居中/.test(request)) {
      return "format";
    }

    // 图表类
    if (/图表|柱状图|折线图|饼图|chart/.test(lowerRequest)) {
      return "chart";
    }

    // 排序筛选类
    if (/排序|筛选|过滤|升序|降序/.test(request)) {
      return "sort_filter";
    }

    // 分析类
    if (/分析|趋势|对比|异常|预测/.test(request)) {
      return "analyze";
    }

    // 创建类
    if (/创建|新建|添加|插入|生成/.test(request)) {
      return "create";
    }

    // 修改类（兜底）
    if (/修改|更改|设置|填|写|改/.test(request)) {
      return "modify";
    }

    return "unclear";
  }

  /**
   * 检测需要澄清的内容
   */
  private detectClarificationNeeds(
    request: string,
    entities: ParsedEntities,
    intentType: IntentType,
    dataModel?: DataModel,
    currentSelection?: string,
    activeSheet?: string
  ): ClarificationItem[] {
    const needs: ClarificationItem[] = [];

    // 1. 检查工作表
    const hasMultipleSheets = dataModel && dataModel.tables.length > 1;
    if (hasMultipleSheets && entities.sheetNames.length === 0 && !activeSheet) {
      needs.push({
        type: "missing_sheet",
        missing: "工作表名称",
        reason: "工作簿有多个工作表，但未指定要操作哪个",
        importance: "critical",
        options: dataModel?.tables.map((t) => t.name),
      });
    }

    // 2. 检查范围（对于修改/删除操作）
    if (["modify", "delete", "calculate"].includes(intentType)) {
      if (entities.ranges.length === 0 && entities.columnNames.length === 0 && !currentSelection) {
        needs.push({
          type: "missing_range",
          missing: "操作范围",
          reason: "未指定要操作的范围或列",
          importance: "critical",
        });
      }
    }

    // 3. 检查模糊指代
    for (const vague of entities.vaguePronous) {
      needs.push({
        type: "vague_reference",
        missing: vague.text,
        reason: `"${vague.text}" 指代不明确`,
        importance: currentSelection ? "optional" : "important",
        options: currentSelection ? [currentSelection] : undefined,
      });
    }

    // 4. 意图不明确
    if (intentType === "unclear") {
      needs.push({
        type: "ambiguous_intent",
        missing: "操作意图",
        reason: "无法确定您想要执行什么操作",
        importance: "critical",
      });
    }

    // 5. 高风险操作
    if (intentType === "delete") {
      needs.push({
        type: "risky_operation",
        missing: "确认",
        reason: "删除操作不可撤销，需要确认",
        importance: "critical",
      });
    }

    // 6. 请求太短
    if (request.length < 10 && intentType !== "query") {
      needs.push({
        type: "incomplete_condition",
        missing: "详细描述",
        reason: "请求太简短，可能遗漏重要信息",
        importance: "important",
      });
    }

    return needs;
  }

  /**
   * 检测歧义
   */
  private detectAmbiguities(
    request: string,
    entities: ParsedEntities,
    dataModel?: DataModel
  ): AmbiguityInfo[] {
    const ambiguities: AmbiguityInfo[] = [];

    // 检测可能匹配多个列的名称
    if (dataModel && entities.columnNames.length === 0) {
      // 模糊匹配
      const words = request.split(/[\s,，。、]+/);
      for (const word of words) {
        if (word.length < 2) continue;

        const matchingFields: string[] = [];
        for (const table of dataModel.tables) {
          for (const field of table.fields) {
            if (field.name.includes(word) || word.includes(field.name)) {
              matchingFields.push(`${table.name}.${field.name}`);
            }
          }
        }

        if (matchingFields.length > 1) {
          ambiguities.push({
            text: word,
            interpretations: matchingFields,
            recommended: matchingFields[0],
          });
        }
      }
    }

    // 检测"表格"歧义（可能指工作表或 Excel Table）
    if (/表格/.test(request)) {
      ambiguities.push({
        text: "表格",
        interpretations: ["工作表 (Sheet)", "Excel 表格 (Table)"],
      });
    }

    return ambiguities;
  }

  /**
   * 计算置信度
   */
  private calculateConfidence(
    request: string,
    entities: ParsedEntities,
    intentType: IntentType,
    clarificationNeeded: ClarificationItem[],
    ambiguities: AmbiguityInfo[]
  ): number {
    let confidence = 1.0;

    // 基础扣分
    if (intentType === "unclear") {
      confidence -= 0.4;
    }

    // 模糊指代扣分
    confidence -= entities.vaguePronous.length * 0.1;

    // 缺失信息扣分
    for (const need of clarificationNeeded) {
      if (need.importance === "critical") {
        confidence -= 0.25;
      } else if (need.importance === "important") {
        confidence -= 0.15;
      } else {
        confidence -= 0.05;
      }
    }

    // 歧义扣分
    confidence -= ambiguities.length * 0.1;

    // 请求太短扣分
    if (request.length < 15) {
      confidence -= 0.15;
    }

    // 有明确实体加分
    if (entities.sheetNames.length > 0) {
      confidence += 0.1;
    }
    if (entities.columnNames.length > 0) {
      confidence += 0.1;
    }
    if (entities.ranges.length > 0) {
      confidence += 0.1;
    }

    return Math.max(0, Math.min(1, confidence));
  }

  /**
   * 评估风险
   */
  private assessRisk(
    intentType: IntentType,
    entities: ParsedEntities,
    _dataModel?: DataModel
  ): "low" | "medium" | "high" {
    // 删除操作总是高风险
    if (intentType === "delete") {
      return "high";
    }

    // 修改操作看范围
    if (intentType === "modify" || intentType === "calculate") {
      // 如果涉及整列或大范围
      if (entities.ranges.some((r) => /[A-Z]+:[A-Z]+/.test(r))) {
        return "high";
      }
      return "medium";
    }

    // 查询和格式化是低风险
    if (intentType === "query" || intentType === "format") {
      return "low";
    }

    return "medium";
  }

  /**
   * 生成澄清问题
   */
  private generateClarificationQuestion(
    request: string,
    needs: ClarificationItem[],
    _ambiguities: AmbiguityInfo[],
    _dataModel?: DataModel
  ): SuggestedClarification {
    // 找到最关键的需要澄清的问题
    const criticalNeeds = needs.filter((n) => n.importance === "critical");
    const primaryNeed = criticalNeeds[0] || needs[0];

    if (!primaryNeed) {
      return {
        mainQuestion: "请提供更多细节，帮助我理解您的需求。",
        allowFreeform: true,
      };
    }

    // 根据类型生成问题
    switch (primaryNeed.type) {
      case "missing_sheet":
        return {
          mainQuestion: "您想在哪个工作表上操作？",
          context: `当前工作簿有 ${primaryNeed.options?.length || "多个"} 个工作表`,
          options: primaryNeed.options?.map((name, i) => ({
            id: `sheet_${i}`,
            label: name,
            recommended: i === 0,
          })),
          allowFreeform: false,
        };

      case "missing_range":
        return {
          mainQuestion: "请指定要操作的范围",
          context: "您可以选择单元格后告诉我，或者直接说明列名/范围",
          options: [
            { id: "selection", label: "使用当前选中区域" },
            { id: "column", label: "指定列（如：A列、金额列）" },
            { id: "range", label: "指定范围（如：A1:D100）" },
          ],
          allowFreeform: true,
        };

      case "vague_reference":
        return {
          mainQuestion: `"${primaryNeed.missing}" 具体指什么？`,
          context: "请明确指出您要操作的位置",
          allowFreeform: true,
        };

      case "ambiguous_intent":
        return {
          mainQuestion: "请告诉我您想要做什么？",
          context: "例如：计算总和、设置格式、查找数据等",
          options: [
            { id: "calculate", label: "计算/公式" },
            { id: "format", label: "格式化/美化" },
            { id: "query", label: "查询/查找" },
            { id: "modify", label: "修改数据" },
          ],
          allowFreeform: true,
        };

      case "risky_operation":
        return {
          mainQuestion: "确认要执行删除操作吗？",
          context: "此操作可能无法撤销",
          options: [
            { id: "confirm", label: "确认删除" },
            { id: "cancel", label: "取消", recommended: true },
          ],
          allowFreeform: false,
        };

      default:
        return {
          mainQuestion: "请提供更多信息",
          context: primaryNeed.reason,
          allowFreeform: true,
        };
    }
  }

  /**
   * 生成建议方案
   */
  private generateSuggestedPlans(
    request: string,
    entities: ParsedEntities,
    intentType: IntentType,
    _dataModel?: DataModel
  ): SuggestedPlan[] {
    const plans: SuggestedPlan[] = [];

    // 根据意图类型生成不同方案
    if (intentType === "calculate" && entities.columnNames.length >= 2) {
      plans.push({
        id: "calc_formula",
        name: "使用公式计算",
        description: `在新列中使用公式引用 ${entities.columnNames.join(" 和 ")}`,
        impact: "在数据末尾添加新列",
        confidence: 0.9,
        recommended: true,
      });

      plans.push({
        id: "calc_values",
        name: "直接计算填值",
        description: "计算结果后直接填入值（不保留公式）",
        impact: "覆盖目标区域",
        confidence: 0.7,
        recommended: false,
      });
    }

    // 如果只有一个方案，就不需要让用户选
    if (plans.length === 0) {
      plans.push({
        id: "default",
        name: "执行请求",
        description: request,
        impact: "按您的要求执行",
        confidence: 0.8,
        recommended: true,
      });
    }

    return plans;
  }
}

// ========== 单例导出 ==========

export const intentAnalyzer = new IntentAnalyzer();

export default IntentAnalyzer;
