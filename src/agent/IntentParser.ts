/**
 * IntentParser - 意图解析器 v4.0
 *
 * Layer 1: 调用 LLM 理解用户意图，输出高层规格
 *
 * ★★★ 核心设计原则 ★★★
 * 1. LLM System Prompt 不包含任何工具名 (如 excel_xxx)
 * 2. LLM 只输出业务意图和规格
 * 3. 工具调用由 SpecCompiler 负责，不是 LLM
 *
 * @module agent/IntentParser
 */

import ApiService from "../services/ApiService";
import { IntentSpec, IntentType, IntentSpecData } from "./types/intent";
import { parseLlmOutput } from "./utils/llmOutputParser";
import { mapToSemanticAtoms, compressToSuperIntent } from "./utils/semanticMapper";

// ========== 解析上下文 ==========

/**
 * 解析上下文
 */
export interface ParseContext {
  /** 用户原始请求 */
  userMessage: string;

  /** 当前选区信息 */
  selection?: {
    address: string;
    values?: unknown[][];
    rowCount?: number;
    columnCount?: number;
  };

  /** 当前工作表名 */
  activeSheet?: string;

  /** 工作簿概要 */
  workbookSummary?: {
    sheetNames: string[];
    tableNames?: string[];
  };

  /** 对话历史 */
  conversationHistory?: Array<{ role: "user" | "assistant"; content: string }>;
}

// ========== IntentParser 类 ==========

/**
 * 意图解析器 - 调用 LLM 理解用户意图
 */
export class IntentParser {
  /**
   * 解析用户意图
   *
   * @param context 解析上下文
   * @returns 意图规格
   */
  async parse(context: ParseContext): Promise<IntentSpec> {
    const systemPrompt = this.buildSystemPrompt();
    const userPrompt = this.buildUserPrompt(context);

    try {
      console.log("[IntentParser] 正在解析用户意图...");

      const response = await ApiService.sendAgentRequest({
        message: userPrompt,
        systemPrompt,
        responseFormat: "json",
      });

      const text = response.message || "";
      console.log("[IntentParser] LLM 返回:", text.substring(0, 300));

      // 先做语义映射（保证在 LLM 未返回结构化数据时仍能做基本路由）
      const atomsFromUser = mapToSemanticAtoms(context.userMessage);
      const compressed = compressToSuperIntent(atomsFromUser);

      // 使用鲁棒解析器先尝试提取 JSON
      const parsedResult = parseLlmOutput(text);
      if (!parsedResult.ok) {
        console.warn("[IntentParser] 无法从 LLM 输出解析 JSON，降级处理：", parsedResult.error);
        const fallback = this.createFallbackSpec(context.userMessage);
        fallback.semanticAtoms = atomsFromUser;
        fallback.compressedIntent = compressed;
        return fallback;
      }

      const spec = this.validateAndConvert(parsedResult.data, context);
      spec.semanticAtoms = atomsFromUser;
      spec.compressedIntent = compressed;
      return spec;
    } catch (error) {
      console.error("[IntentParser] 解析失败:", error);
      return this.createFallbackSpec(context.userMessage);
    }
  }

  /**
   * 构建 System Prompt
   *
   * ★★★ 关键: 不包含任何工具名 ★★★
   *
   * @public 暴露给测试使用
   */
  buildSystemPrompt(): string {
    return `你是 Excel 智能助手的意图理解模块。你的任务是理解用户想做什么，而不是怎么做。

## 你的职责
1. 理解用户的真实意图
2. 提取关键信息（范围、数据、条件等）
3. 判断信息是否完整
4. 如果信息不足，指出需要澄清什么

## 意图类型（选择最匹配的一个）
- create_table: 创建新表格
- write_data: 写入数据
- update_data: 更新数据
- delete_data: 删除数据
- format_range: 格式化（加粗、颜色、边框等）
- create_formula: 创建公式
- analyze_data: 分析数据
- create_chart: 创建图表
- sort_data: 排序
- filter_data: 筛选
- query_data: 查询/读取
- create_sheet: 创建工作表
- clarify: 需要澄清
- respond_only: 只需回复，不需操作

## ★★★ 澄清规则（非常重要）★★★

### 不需要澄清的情况（直接执行！）
- "读取当前表格" → query_data，读取选中区域或活动工作表
- "看看这个表" → query_data，读取并返回概览
- "这是什么数据" → query_data，读取并分析
- "帮我求和" → create_formula，对选中区域求和
- "加粗标题" → format_range，加粗第一行
- "在A1写入xxx" → write_data，明确的写入操作

### 必须澄清的情况
1. **删除类请求但目标不明确**
   - "删除没用的" → 什么是"没用的"？
   - "清理一下" → 清理什么？

2. **范围不明确的破坏性操作**
   - "把数据改一下" → 改什么数据？改成什么？
   
3. **模糊的修改指代**
   - "把这个表格优化一下" → 优化什么？格式？数据？结构？

### ⚠️ 重要：不要过度澄清！
- 查询类操作（读取、查看、分析）不需要澄清，直接返回数据
- 如果用户说"当前表格"、"这个"、"选中的"，就用当前选区
- 宁可多做一点，也不要反复追问用户

## 输出 JSON 格式

{
  "intent": "意图类型",
  "confidence": 0.0-1.0,
  "needsClarification": true/false,
  "clarificationQuestion": "如果需要澄清，问什么",
  "clarificationOptions": ["选项1", "选项2"],
  "spec": {
    // 根据意图类型填写具体规格
    // 如 create_table: { columns: [...], startCell: "A1" }
    // 如 format_range: { range: "A1:D10", format: { bold: true } }
  },
  "reasoning": "你的思考过程"
}

## 规格示例

### create_table
{
  "intent": "create_table",
  "confidence": 0.9,
  "spec": {
    "type": "create_table",
    "columns": [
      { "name": "日期", "type": "date" },
      { "name": "产品", "type": "text" },
      { "name": "数量", "type": "number" },
      { "name": "金额", "type": "currency" }
    ],
    "options": { "hasHeader": true, "hasTotalRow": true }
  }
}

### clarify (删除类)
{
  "intent": "clarify",
  "confidence": 0.3,
  "needsClarification": true,
  "clarificationQuestion": "您想删除哪些内容？",
  "clarificationOptions": [
    "空白的行",
    "空白的列", 
    "重复的数据",
    "特定的列（请指定）"
  ],
  "spec": {
    "type": "clarify",
    "question": "您想删除哪些内容？",
    "reason": "请求不明确，需要确认删除目标"
  }
}

### query_data (读取/查看)
{
  "intent": "query_data",
  "confidence": 0.9,
  "needsClarification": false,
  "spec": {
    "type": "query",
    "target": "selection"
  }
}

### write_data (写入)
{
  "intent": "write_data",
  "confidence": 0.9,
  "needsClarification": false,
  "spec": {
    "type": "write_data",
    "target": "A1",
    "data": "要写入的内容"
  }
}

## 注意
1. 查询类操作不需要澄清，直接执行
2. 只有破坏性操作（删除、覆盖）且目标不明确时才澄清
3. 只输出 JSON，不要其他内容
4. 不要过度澄清，用户会觉得你很蠢`;
  }

  /**
   * 构建用户 Prompt
   *
   * @public 暴露给测试使用
   */
  buildUserPrompt(
    context: ParseContext | { userInput?: string; workbookContext?: Record<string, unknown> }
  ): string {
    // 兼容两种调用方式
    let userMessage: string;
    let selection: ParseContext["selection"] | undefined;
    let activeSheet: string | undefined;
    let workbookSummary: ParseContext["workbookSummary"] | undefined;
    let conversationHistory: ParseContext["conversationHistory"] | undefined;

    if ("userMessage" in context) {
      // ParseContext 格式
      userMessage = context.userMessage;
      selection = context.selection;
      activeSheet = context.activeSheet;
      workbookSummary = context.workbookSummary;
      conversationHistory = context.conversationHistory;
    } else {
      // 简化格式
      userMessage = context.userInput || "";
      activeSheet = context.workbookContext?.activeSheet as string | undefined;
      const workbookName = context.workbookContext?.workbookName;
      if (workbookName) {
        workbookSummary = { sheetNames: [workbookName as string] };
      }
    }

    let prompt = `用户请求: ${userMessage}\n\n`;

    if (selection) {
      prompt += `当前选区: ${selection.address}`;
      if (selection.rowCount && selection.columnCount) {
        prompt += ` (${selection.rowCount}行 × ${selection.columnCount}列)`;
      }
      prompt += "\n";
    }

    if (activeSheet) {
      prompt += `当前工作表: ${activeSheet}\n`;
    }

    if (workbookSummary) {
      prompt += `工作簿包含: ${workbookSummary.sheetNames.join(", ")}\n`;
    }

    if (conversationHistory && conversationHistory.length > 0) {
      prompt += "\n对话历史:\n";
      const recent = conversationHistory.slice(-4);
      for (const msg of recent) {
        prompt += `${msg.role === "user" ? "用户" : "助手"}: ${msg.content.substring(0, 100)}\n`;
      }
    }

    prompt += "\n请分析用户意图并输出 JSON。";
    return prompt;
  }

  /**
   * 验证并转换 LLM 输出
   */
  private validateAndConvert(parsed: Record<string, unknown>, context: ParseContext): IntentSpec {
    // 验证必需字段
    const intent = parsed.intent as IntentType;
    if (!intent) {
      console.warn("[IntentParser] 缺少 intent 字段");
      return this.createFallbackSpec(context.userMessage);
    }

    const confidence = typeof parsed.confidence === "number" ? parsed.confidence : 0.5;
    const needsClarification = Boolean(parsed.needsClarification);

    // 构建规格
    let spec: IntentSpecData;
    if (parsed.spec && typeof parsed.spec === "object") {
      spec = parsed.spec as IntentSpecData;
    } else {
      spec = { type: "respond", message: "无法解析规格" };
    }

    return {
      intent,
      confidence,
      needsClarification,
      clarificationQuestion: parsed.clarificationQuestion as string | undefined,
      clarificationOptions: parsed.clarificationOptions as string[] | undefined,
      spec,
      reasoning: parsed.reasoning as string | undefined,
    };
  }

  /**
   * 创建降级规格
   */
  private createFallbackSpec(_userMessage: string): IntentSpec {
    return {
      intent: "clarify",
      confidence: 0.3,
      needsClarification: true,
      clarificationQuestion: "抱歉，我没有完全理解您的请求。您能具体说明一下想做什么吗？",
      spec: {
        type: "clarify",
        question: "请更具体地描述您的需求",
        reason: "无法解析用户意图",
      },
    };
  }
}

// ========== 单例导出 ==========

export const intentParser = new IntentParser();

export default IntentParser;
