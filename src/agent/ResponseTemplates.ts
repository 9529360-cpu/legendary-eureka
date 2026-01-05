/**
 * ResponseTemplates - 自然语言响应模板系统 v2.1
 *
 * v2.1 核心改动：P3 - 模板变为可选，支持 LLM 自由响应
 *
 * v2.0 核心改动：ExecutionState 驱动
 *
 * 致命原则：
 * ┌─────────────────────────────────────────────────────┐
 * │  LLM 可以说 "我建议怎么做"                          │
 * │  但只有 executor 才能说 "我已经做了"                │
 * └─────────────────────────────────────────────────────┘
 *
 * v2.1 改动：
 * - 模板是可选的，不是强制的
 * - allowFreeformResponse = true 时，LLM 可以自由表达
 * - 但即使自由表达，也必须尊重 executionState 的约束
 *
 * 禁止规则：
 * - 没执行 → 绝不说"完成/搞定/已创建"
 * - 执行失败 → 明确说失败
 * - 执行成功 → 只用 executor 的真实结果说话
 */

import ApiService from "../services/ApiService";

// ========== 类型定义 ==========

/**
 * 执行状态（核心！决定能说什么）
 */
export type ExecutionState =
  | "planned" // 已规划，待确认
  | "preview" // 预览中，等用户点确认
  | "executing" // 执行中
  | "executed" // 执行成功
  | "partial" // 部分成功
  | "failed" // 执行失败
  | "rolled_back"; // 已回滚

/**
 * 响应上下文（v2.0: 必须带 executionState）
 */
export interface ResponseContext {
  // 必填
  executionState: ExecutionState;
  taskType: string;

  // v2.1: P3 - 自由响应控制
  allowFreeformResponse?: boolean; // 是否允许 LLM 自由生成
  userRequest?: string; // 原始用户请求（用于 LLM 生成）
  executionSummary?: string; // 执行摘要（用于 LLM 生成）

  // 执行结果（只有 executed 状态才可信）
  result?: ExecutionResult;
  error?: ExecutionError;

  // LLM 提供的受控内容（有长度限制）
  llmSummary?: string; // 1-2句，最多80字
  llmFindings?: string[]; // 最多3条
  llmRiskNote?: string; // 1条风险提示
  llmSuggestion?: string; // 下一步建议

  // 操作上下文
  operationType?: string;
  targetRange?: string;
  sheetName?: string;

  // 兼容旧接口（逐步废弃）
  dataCount?: number;
  columns?: string[];
  chartType?: string;
  formulaType?: string;
  duration?: number;
}

/**
 * 执行结果（来自 executor 的真实数据）
 */
export interface ExecutionResult {
  affectedRange?: string; // 实际修改的范围
  affectedCells?: number; // 实际修改的单元格数
  writtenRows?: number; // 实际写入的行数
  computedValue?: unknown; // 计算结果（如 SUM 的值）
  createdObject?: string; // 创建的对象（如图表名称）
  changes?: ChangeRecord[]; // 变更记录
}

/**
 * 变更记录
 */
export interface ChangeRecord {
  range: string;
  action: "write" | "format" | "formula" | "clear" | "create";
  before?: unknown;
  after?: unknown;
}

/**
 * 执行错误
 */
export interface ExecutionError {
  code: string;
  message: string;
  range?: string;
  recoverable?: boolean;
}

// ========== 禁止词表 ==========

/**
 * 这些词只能在 executionState === "executed" 时使用
 */
const COMPLETION_WORDS = [
  "已完成",
  "完成了",
  "搞定",
  "已经帮你",
  "已创建",
  "已生成",
  "已设置",
  "已修改",
  "已清空",
  "已删除",
  "已排序",
  "已筛选",
  "创建完成",
  "设置完成",
  "修改完成",
  "处理完成",
  "Done",
  "Completed",
  "Created",
  "Set",
  "Modified",
];

// ========== 响应生成器 ==========

/**
 * 响应模板生成器（v2.0: ExecutionState 驱动）
 */
export class ResponseGenerator {
  /**
   * 生成自然语言响应（主入口）
   *
   * 核心逻辑：先看 executionState，再看 taskType
   */
  static generate(context: ResponseContext): string {
    const { executionState } = context;

    // ========== 第一层：按执行状态分流 ==========

    // 1. 规划状态 → 只能说"我打算..."
    if (executionState === "planned") {
      return this.generatePlannedResponse(context);
    }

    // 2. 预览状态 → 只能说"准备..."
    if (executionState === "preview") {
      return this.generatePreviewResponse(context);
    }

    // 3. 执行中 → 只能说"正在..."
    if (executionState === "executing") {
      return this.generateExecutingResponse(context);
    }

    // 4. 失败状态 → 必须说失败
    if (executionState === "failed") {
      return this.generateFailedResponse(context);
    }

    // 5. 回滚状态 → 说明已撤销
    if (executionState === "rolled_back") {
      return this.generateRolledBackResponse(context);
    }

    // 6. 部分成功 → 说明哪些成功哪些失败
    if (executionState === "partial") {
      return this.generatePartialResponse(context);
    }

    // 7. 只有 executed 才能进入"完成"逻辑
    if (executionState === "executed") {
      return this.generateExecutedResponse(context);
    }

    // 兜底：未知状态，保守处理
    return "操作状态未知，请检查 Excel 中的实际结果。";
  }

  // ========== 各状态的响应生成 ==========

  /**
   * 规划状态：只能说"我打算..."
   */
  private static generatePlannedResponse(context: ResponseContext): string {
    const { taskType, llmSummary, llmFindings, llmRiskNote } = context;

    let response = "📋 **我的计划**\n\n";

    // LLM 的解释（受控）
    if (llmSummary) {
      response += this.truncate(llmSummary, 80) + "\n\n";
    }

    // LLM 发现的问题（最多3条）
    if (llmFindings && llmFindings.length > 0) {
      response += "我发现了以下情况：\n";
      llmFindings.slice(0, 3).forEach((finding, i) => {
        response += `${i + 1}. ${this.truncate(finding, 50)}\n`;
      });
      response += "\n";
    }

    // 任务描述
    response += this.getTaskDescription(taskType, context);

    // 风险提示
    if (llmRiskNote) {
      response += `\n\n⚠️ 注意：${this.truncate(llmRiskNote, 60)}`;
    }

    response += "\n\n**需要我执行吗？**";

    return response;
  }

  /**
   * 预览状态：只能说"准备..."
   */
  private static generatePreviewResponse(context: ResponseContext): string {
    const { targetRange, operationType, llmSummary } = context;

    let response = "👀 **操作预览**\n\n";

    if (llmSummary) {
      response += this.truncate(llmSummary, 80) + "\n\n";
    }

    const actionDesc = this.getOperationDescription(operationType);
    response += `准备${actionDesc}`;

    if (targetRange) {
      response += `，目标范围：\`${targetRange}\``;
    }

    response += "\n\n**确认执行？** 点击确认后我才会真正修改 Excel。";

    return response;
  }

  /**
   * 执行中：只能说"正在..."
   */
  private static generateExecutingResponse(context: ResponseContext): string {
    const { operationType, targetRange } = context;

    const actionDesc = this.getOperationDescription(operationType);
    let response = `⏳ 正在${actionDesc}`;

    if (targetRange) {
      response += `（${targetRange}）`;
    }

    response += "...";

    return response;
  }

  /**
   * 执行成功：可以说"完成"，但必须引用真实结果
   */
  private static generateExecutedResponse(context: ResponseContext): string {
    const { taskType, result } = context;

    // 必须有执行结果
    if (!result) {
      return "✅ 操作已执行，但未返回详细结果。请检查 Excel 中的实际变化。";
    }

    // 根据任务类型生成响应
    switch (taskType) {
      case "data_generation":
      case "write":
        return this.generateWriteCompletedResponse(context);
      case "format":
        return this.generateFormatCompletedResponse(context);
      case "formula":
        return this.generateFormulaCompletedResponse(context);
      case "chart":
        return this.generateChartCompletedResponse(context);
      case "sort":
        return this.generateSortCompletedResponse(context);
      case "filter":
        return this.generateFilterCompletedResponse(context);
      case "clear":
        return this.generateClearCompletedResponse(context);
      case "analysis":
      case "query":
        return this.generateQueryCompletedResponse(context);
      default:
        return this.generateGenericCompletedResponse(context);
    }
  }

  /**
   * 执行失败：必须明确说失败
   */
  private static generateFailedResponse(context: ResponseContext): string {
    const { error, llmSuggestion } = context;

    let response = "❌ **操作失败**\n\n";

    if (error) {
      const friendlyError = this.translateError(error);
      response += `原因：${friendlyError}\n`;

      if (error.range) {
        response += `位置：\`${error.range}\`\n`;
      }
    } else {
      response += "执行过程中遇到了问题。\n";
    }

    // LLM 的建议（受控）
    if (llmSuggestion) {
      response += `\n💡 建议：${this.truncate(llmSuggestion, 80)}`;
    } else if (error?.recoverable) {
      response += "\n💡 这个问题可以修复，你可以告诉我更多细节。";
    }

    return response;
  }

  /**
   * 回滚状态
   */
  private static generateRolledBackResponse(context: ResponseContext): string {
    const { targetRange, error } = context;

    let response = "↩️ **操作已撤销**\n\n";
    response += "检测到问题，已自动回滚到操作前的状态。";

    if (targetRange) {
      response += `\n范围 \`${targetRange}\` 已恢复原样。`;
    }

    if (error) {
      response += `\n\n原因：${this.translateError(error)}`;
    }

    return response;
  }

  /**
   * 部分成功
   */
  private static generatePartialResponse(context: ResponseContext): string {
    const { result, error } = context;

    let response = "⚠️ **部分完成**\n\n";

    if (result?.affectedCells) {
      response += `✅ 成功修改了 ${result.affectedCells} 个单元格\n`;
    }

    if (error) {
      response += `❌ 失败：${this.translateError(error)}\n`;
    }

    response += "\n请检查 Excel 中的实际结果。";

    return response;
  }

  // ========== 具体任务的"完成"响应（必须引用 result）==========

  private static generateWriteCompletedResponse(context: ResponseContext): string {
    const { result, columns } = context;

    let response = "✅ **数据写入完成**\n\n";

    if (result?.affectedRange) {
      response += `📍 位置：\`${result.affectedRange}\`\n`;
    }

    if (result?.writtenRows) {
      response += `📊 写入了 ${result.writtenRows} 行数据\n`;
    }

    if (columns && columns.length > 0) {
      const colDesc = columns.slice(0, 4).join("、") + (columns.length > 4 ? " 等" : "");
      response += `📋 包含 ${columns.length} 列：${colDesc}\n`;
    }

    return response;
  }

  private static generateFormatCompletedResponse(context: ResponseContext): string {
    const { result, operationType } = context;

    const formatDesc = this.getOperationDescription(operationType);
    let response = `✅ **格式化完成**\n\n`;

    response += `已${formatDesc}`;

    if (result?.affectedRange) {
      response += `\n📍 范围：\`${result.affectedRange}\``;
    }

    if (result?.affectedCells) {
      response += `\n📊 影响 ${result.affectedCells} 个单元格`;
    }

    return response;
  }

  private static generateFormulaCompletedResponse(context: ResponseContext): string {
    const { result, formulaType } = context;

    let response = "✅ **公式设置完成**\n\n";

    if (result?.computedValue !== undefined) {
      const typeDesc = this.getFormulaTypeDescription(formulaType);
      response += `📊 ${typeDesc}：**${result.computedValue}**\n`;
    }

    if (result?.affectedRange) {
      response += `📍 公式位置：\`${result.affectedRange}\``;
    }

    return response;
  }

  private static generateChartCompletedResponse(context: ResponseContext): string {
    const { result, chartType } = context;

    const chartDesc = this.getChartTypeDescription(chartType);
    let response = `✅ **${chartDesc}创建完成**\n\n`;

    if (result?.createdObject) {
      response += `📈 图表名称：${result.createdObject}\n`;
    }

    if (result?.affectedRange) {
      response += `📊 数据来源：\`${result.affectedRange}\``;
    }

    return response;
  }

  private static generateSortCompletedResponse(context: ResponseContext): string {
    const { result } = context;

    let response = "✅ **排序完成**\n\n";

    if (result?.affectedRange) {
      response += `📍 范围：\`${result.affectedRange}\`\n`;
    }

    if (result?.affectedCells) {
      response += `📊 排序了 ${result.affectedCells} 个单元格`;
    }

    return response;
  }

  private static generateFilterCompletedResponse(context: ResponseContext): string {
    const { result, dataCount } = context;

    let response = "✅ **筛选完成**\n\n";

    const count = dataCount ?? result?.writtenRows;
    if (count !== undefined) {
      response += `🔍 找到 ${count} 条符合条件的数据`;
    }

    if (result?.affectedRange) {
      response += `\n📍 范围：\`${result.affectedRange}\``;
    }

    return response;
  }

  private static generateClearCompletedResponse(context: ResponseContext): string {
    const { result } = context;

    let response = "✅ **清除完成**\n\n";

    if (result?.affectedRange) {
      response += `🧹 已清空：\`${result.affectedRange}\`\n`;
    }

    if (result?.affectedCells) {
      response += `📊 清除了 ${result.affectedCells} 个单元格`;
    }

    return response;
  }

  private static generateQueryCompletedResponse(context: ResponseContext): string {
    const { result, formulaType } = context;

    if (result?.computedValue !== undefined) {
      const typeDesc = this.getFormulaTypeDescription(formulaType);
      return `📊 ${typeDesc}：**${result.computedValue}**`;
    }

    return "📊 查询完成，请查看 Excel 中的结果。";
  }

  private static generateGenericCompletedResponse(context: ResponseContext): string {
    const { result } = context;

    let response = "✅ **操作完成**\n\n";

    if (result?.affectedRange) {
      response += `📍 范围：\`${result.affectedRange}\`\n`;
    }

    if (result?.affectedCells) {
      response += `📊 影响 ${result.affectedCells} 个单元格`;
    }

    if (!result?.affectedRange && !result?.affectedCells) {
      response += "请查看 Excel 中的实际变化。";
    }

    return response;
  }

  // ========== 工具方法 ==========

  /**
   * 获取任务描述
   */
  private static getTaskDescription(taskType: string, context: ResponseContext): string {
    const { targetRange, operationType: _operationType } = context;

    const taskDescriptions: Record<string, string> = {
      data_generation: "生成数据表格",
      write: "写入数据",
      format: "格式化单元格",
      formula: "设置公式",
      chart: "创建图表",
      sort: "排序数据",
      filter: "筛选数据",
      clear: "清除内容",
      analysis: "分析数据",
      query: "查询数据",
    };

    let desc = `我将${taskDescriptions[taskType] || "执行操作"}`;

    if (targetRange) {
      desc += `，目标范围：\`${targetRange}\``;
    }

    return desc;
  }

  /**
   * 获取操作描述
   */
  private static getOperationDescription(operationType?: string): string {
    const operationDescriptions: Record<string, string> = {
      bold: "加粗",
      color: "设置颜色",
      fill: "填充背景",
      border: "添加边框",
      align: "对齐",
      font: "设置字体",
      autofit: "自动调整列宽",
      numberFormat: "设置数字格式",
      write: "写入数据",
      clear: "清除内容",
      formula: "设置公式",
    };

    return operationDescriptions[operationType || ""] || "执行操作";
  }

  /**
   * 获取公式类型描述
   */
  private static getFormulaTypeDescription(formulaType?: string): string {
    const formulaDescriptions: Record<string, string> = {
      sum: "总和",
      average: "平均值",
      count: "计数",
      max: "最大值",
      min: "最小值",
      vlookup: "查找结果",
      xlookup: "查找结果",
    };

    return formulaDescriptions[formulaType || ""] || "计算结果";
  }

  /**
   * 获取图表类型描述
   */
  private static getChartTypeDescription(chartType?: string): string {
    const chartDescriptions: Record<string, string> = {
      column: "柱状图",
      bar: "条形图",
      line: "折线图",
      pie: "饼图",
      area: "面积图",
      scatter: "散点图",
    };

    return chartDescriptions[chartType || ""] || "图表";
  }

  /**
   * 翻译错误信息
   */
  private static translateError(error: ExecutionError): string {
    const errorTranslations: Array<{ pattern: RegExp; friendly: string }> = [
      { pattern: /invalid range/i, friendly: "范围地址格式不正确" },
      { pattern: /permission|protected/i, friendly: "工作表被保护，没有操作权限" },
      { pattern: /network|timeout/i, friendly: "网络连接超时" },
      { pattern: /not found/i, friendly: "找不到指定的内容" },
      { pattern: /empty/i, friendly: "目标区域是空的" },
      { pattern: /busy|conflict/i, friendly: "Excel 正忙，请稍后重试" },
    ];

    for (const { pattern, friendly } of errorTranslations) {
      if (pattern.test(error.message)) {
        return friendly;
      }
    }

    return error.message || "发生未知错误";
  }

  /**
   * 截断文本
   */
  private static truncate(text: string, maxLength: number): string {
    if (!text) return "";
    if (text.length <= maxLength) return text;
    return text.substring(0, maxLength - 3) + "...";
  }

  // ========== 特殊场景模板（保持兼容）==========

  /**
   * 生成问候响应
   */
  static generateGreeting(): string {
    return `你好！👋 我是你的 Excel 助手。

我可以帮你：
• 生成和填充数据
• 设置公式和计算
• 格式化和美化表格
• 创建图表

直接告诉我你想做什么吧！`;
  }

  /**
   * 生成确认响应
   */
  static generateAcknowledgment(request: string): string {
    const text = request.toLowerCase();

    if (/谢谢|thanks|thx/.test(text)) {
      return "不客气！有需要随时说 👍";
    }

    if (/好的|ok|知道了|明白/.test(text)) {
      return "👍";
    }

    if (/拜拜|再见|bye/.test(text)) {
      return "再见！👋";
    }

    return "👍";
  }

  /**
   * 生成帮助响应
   */
  static generateHelp(): string {
    return `# Excel 助手使用指南

## 数据操作
• 「生成一个客户信息表」
• 「把选中的数据求和」

## 格式美化
• 「把标题加粗」
• 「给表格加边框」

## 图表分析
• 「画个柱状图」
• 「按销售额排序」

## 公式设置
• 「计算总和」
• 「求平均值」

---
💡 直接用自然语言描述即可！`;
  }

  /**
   * 生成进度消息
   */
  static generateProgress(current: number, total: number, stepDescription?: string): string {
    const percentage = Math.round((current / total) * 100);
    const progressBar =
      "█".repeat(Math.floor(percentage / 10)) + "░".repeat(10 - Math.floor(percentage / 10));
    return `[${progressBar}] ${percentage}%${stepDescription ? ` - ${stepDescription}` : ""}`;
  }

  // ========== P3: 自由响应生成（v2.1） ==========

  /**
   * 异步生成响应（主入口，支持 LLM 自由表达）
   *
   * @param context 响应上下文
   * @returns 生成的响应文本
   *
   * P3 核心逻辑：
   * 1. 如果 allowFreeformResponse=true，优先使用 LLM 自由生成
   * 2. LLM 必须遵循 executionState 约束（失败不能说成功等）
   * 3. 如果 LLM 调用失败，回退到模板
   */
  static async generateAsync(context: ResponseContext): Promise<string> {
    const { executionState, allowFreeformResponse, userRequest, executionSummary } = context;

    // 如果允许自由响应且有必要信息
    if (allowFreeformResponse && userRequest && executionSummary) {
      try {
        const freeformResponse = await this.generateFreeformResponse(
          userRequest,
          executionSummary,
          executionState,
          context
        );

        // 验证 freeform 响应是否违反约束
        if (this.validateFreeformResponse(freeformResponse, executionState)) {
          return freeformResponse;
        }

        console.warn("[ResponseGenerator] LLM 响应违反状态约束，回退到模板");
      } catch (error) {
        console.warn("[ResponseGenerator] LLM 自由响应失败，回退到模板:", error);
      }
    }

    // 回退到同步模板生成
    return this.generate(context);
  }

  /**
   * P3: 调用 LLM 生成自由响应
   *
   * 关键：给 LLM 明确的约束，但让它自由表达
   */
  private static async generateFreeformResponse(
    userRequest: string,
    executionSummary: string,
    executionState: ExecutionState,
    context: ResponseContext
  ): Promise<string> {
    // 构建约束提示
    const stateConstraint = this.getStateConstraint(executionState);
    const completionWordWarning = this.getCompletionWordWarning(executionState);

    const systemPrompt = `你是 Excel 智能助手。用户刚刚请求了一个操作，你需要用自然、友好的方式告诉用户结果。

## 硬性约束（必须遵守）

${stateConstraint}

${completionWordWarning}

## 风格要求

- 简洁：一般 1-3 句话
- 自然：像人说话，不要生硬
- 具体：如果有具体数据（行数、范围等），可以提及
- 情感适度：成功时可以表达肯定，失败时要诚恳

## 上下文

执行状态: ${executionState}
操作类型: ${context.operationType || "未知"}
目标范围: ${context.targetRange || "未指定"}
`;

    const userPrompt = `用户请求：${userRequest}

执行结果：${executionSummary}

请用 1-3 句话自然地告诉用户结果。`;

    // 调用 API
    const response = await ApiService.sendChatRequest({
      message: userPrompt,
      systemPrompt,
      responseFormat: "text",
    });

    if (response.success && response.message) {
      return response.message.trim();
    }

    throw new Error(response.message || "LLM 响应为空");
  }

  /**
   * P3: 根据执行状态生成硬性约束
   */
  private static getStateConstraint(state: ExecutionState): string {
    const constraints: Record<ExecutionState, string> = {
      planned: '【状态：规划中】你只能说"我打算..."、"计划..."，不能说"已完成"、"已执行"。',
      preview: '【状态：预览中】你只能说"准备..."、"即将..."，不能说"已完成"、"已执行"。',
      executing: '【状态：执行中】你只能说"正在..."、"处理中..."，不能说"已完成"。',
      executed: '【状态：已执行】操作已完成，你可以说"已完成"、"搞定了"等。',
      failed: '【状态：失败】操作失败了，你必须说明失败，不能说"已完成"。要诚恳道歉。',
      partial: "【状态：部分成功】有些操作成功，有些失败。你要说明哪些成功哪些失败。",
      rolled_back: '【状态：已回滚】操作被撤销了，你要说明"已撤销"或"已回滚"。',
    };
    return constraints[state] || "状态未知，请谨慎表达。";
  }

  /**
   * P3: 完成关键词警告
   */
  private static getCompletionWordWarning(state: ExecutionState): string {
    const nonCompleteStates: ExecutionState[] = ["planned", "preview", "executing", "failed"];
    if (nonCompleteStates.includes(state)) {
      return `## 禁用词警告

以下词汇在当前状态下绝对禁止使用：
${COMPLETION_WORDS.map((w) => `- "${w}"`).join("\n")}

使用这些词会导致用户误解操作已完成，这是严重错误！`;
    }
    return "";
  }

  /**
   * P3: 验证 LLM 响应是否违反状态约束
   */
  private static validateFreeformResponse(
    response: string,
    executionState: ExecutionState
  ): boolean {
    // 如果不是完成状态，检查是否误用完成关键词
    const nonCompleteStates: ExecutionState[] = ["planned", "preview", "executing", "failed"];

    if (nonCompleteStates.includes(executionState)) {
      for (const word of COMPLETION_WORDS) {
        if (response.includes(word)) {
          console.warn(
            `[ResponseGenerator] LLM 响应在 ${executionState} 状态下使用了禁用词: ${word}`
          );
          return false;
        }
      }
    }

    // 失败状态必须有失败/道歉语气
    if (executionState === "failed") {
      const failureWords = ["失败", "抱歉", "出错", "没能", "无法", "sorry", "fail", "error"];
      const hasFailureWord = failureWords.some((w) => response.toLowerCase().includes(w));
      if (!hasFailureWord) {
        console.warn("[ResponseGenerator] LLM 响应在 failed 状态下缺少失败语气词");
        return false;
      }
    }

    return true;
  }
}

// ========== 导出 ==========

export default ResponseGenerator;
