/**
 * AgentOrchestrator - Agent 协调器
 *
 * 单一职责：协调整个 Agent 工作流程
 * 行数上限：400 行
 *
 * 工作流程：
 * 用户输入 → SemanticExtractor → DiagnosticEngine → SolutionBuilder → 执行计划
 */

import { semanticExtractor } from "./semantic/SemanticExtractor";
import { diagnosticEngine } from "./semantic/DiagnosticEngine";
import { solutionBuilder } from "./solutions/SolutionBuilder";
import {
  SemanticExtraction,
  DiagnosticResult,
  LayeredSolution,
  AgentEvent,
  ClarificationRequest,
} from "./types";

// ========== 工作流状态 ==========

/**
 * 工作流阶段
 */
export type WorkflowPhase =
  | "idle"
  | "extracting"
  | "diagnosing"
  | "building_solution"
  | "awaiting_clarification"
  | "executing"
  | "completed"
  | "error";

/**
 * 工作流结果
 */
export interface WorkflowResult {
  phase: WorkflowPhase;
  semanticExtraction?: SemanticExtraction;
  diagnosis?: DiagnosticResult;
  solution?: LayeredSolution;
  clarificationNeeded?: ClarificationRequest;
  error?: Error;
  timestamp: number;
}

/**
 * 编排配置
 */
export interface OrchestratorConfig {
  enableDiagnosis: boolean;
  enableSolutionBuilder: boolean;
  confidenceThreshold: number;
  maxRetries: number;
}

const DEFAULT_CONFIG: OrchestratorConfig = {
  enableDiagnosis: true,
  enableSolutionBuilder: true,
  confidenceThreshold: 0.6,
  maxRetries: 3,
};

// ========== AgentOrchestrator 类 ==========

/**
 * Agent 协调器
 */
export class AgentOrchestrator {
  private config: OrchestratorConfig;
  private eventListeners: Map<string, ((event: AgentEvent) => void)[]>;
  private currentPhase: WorkflowPhase;

  constructor(config: Partial<OrchestratorConfig> = {}) {
    this.config = { ...DEFAULT_CONFIG, ...config };
    this.eventListeners = new Map();
    this.currentPhase = "idle";
  }

  /**
   * 处理用户输入
   */
  async process(userInput: string, context?: Record<string, unknown>): Promise<WorkflowResult> {
    const startTime = Date.now();
    let result: WorkflowResult = {
      phase: "idle",
      timestamp: startTime,
    };

    try {
      // 1. 语义提取
      this.setPhase("extracting");
      this.emit("phase_change", { phase: "extracting", input: userInput });

      const extraction = semanticExtractor.extract(userInput, context);
      result.semanticExtraction = extraction;

      // 检查置信度是否足够
      if (extraction.confidence < this.config.confidenceThreshold) {
        result = this.handleLowConfidence(extraction, result);
        return result;
      }

      // 2. 诊断（如果需要）
      if (this.config.enableDiagnosis && this.needsDiagnosis(extraction)) {
        this.setPhase("diagnosing");
        this.emit("phase_change", { phase: "diagnosing" });

        const diagnosis = diagnosticEngine.diagnose(userInput, context);
        result.diagnosis = diagnosis;
      }

      // 3. 构建解决方案
      if (this.config.enableSolutionBuilder) {
        this.setPhase("building_solution");
        this.emit("phase_change", { phase: "building_solution" });

        if (result.diagnosis) {
          result.solution = solutionBuilder.buildFromDiagnosis(result.diagnosis);
        } else {
          result.solution = solutionBuilder.buildFromSemanticExtraction(extraction);
        }
      }

      // 4. 完成
      this.setPhase("completed");
      result.phase = "completed";
      this.emit("completed", { result });
    } catch (error) {
      this.setPhase("error");
      result.phase = "error";
      result.error = error instanceof Error ? error : new Error(String(error));
      this.emit("error", { error: result.error });
    }

    result.timestamp = Date.now() - startTime;
    return result;
  }

  /**
   * 处理低置信度情况
   */
  private handleLowConfidence(
    extraction: SemanticExtraction,
    result: WorkflowResult
  ): WorkflowResult {
    this.setPhase("awaiting_clarification");

    const clarification: ClarificationRequest = {
      type: "ambiguous_intent",
      message: "我需要更多信息来理解您的需求",
      suggestions: this.generateSuggestions(extraction),
      context: {
        detectedIntent: extraction.intent,
        confidence: extraction.confidence,
      },
    };

    result.phase = "awaiting_clarification";
    result.clarificationNeeded = clarification;

    this.emit("clarification_needed", { clarification });

    return result;
  }

  /**
   * 生成建议
   */
  private generateSuggestions(extraction: SemanticExtraction): string[] {
    const suggestions: string[] = [];

    switch (extraction.intent) {
      case "create_formula":
        suggestions.push("您想创建什么类型的公式？（求和、平均、查找等）");
        suggestions.push("请告诉我需要计算的数据范围");
        break;
      case "format":
        suggestions.push("您想设置什么格式？（颜色、字体、边框等）");
        suggestions.push("请指定要格式化的区域");
        break;
      case "analyze":
        suggestions.push("您想分析什么数据？");
        suggestions.push("您希望得到什么类型的分析结果？");
        break;
      default:
        suggestions.push("请更详细地描述您的需求");
        suggestions.push("您可以提供具体的数据范围或操作类型");
    }

    return suggestions;
  }

  /**
   * 判断是否需要诊断
   */
  private needsDiagnosis(extraction: SemanticExtraction): boolean {
    // 问题排查类意图需要诊断
    const diagnosticIntents = ["diagnose", "debug", "troubleshoot"];
    if (diagnosticIntents.includes(extraction.intent)) {
      return true;
    }

    // 包含问题关键词
    const problemKeywords = ["错误", "不对", "失败", "问题", "为什么", "#REF", "#VALUE", "#NAME"];
    const hasProblems = problemKeywords.some((kw) => extraction.rawInput?.includes(kw));

    return hasProblems;
  }

  /**
   * 设置当前阶段
   */
  private setPhase(phase: WorkflowPhase): void {
    this.currentPhase = phase;
  }

  /**
   * 获取当前阶段
   */
  getPhase(): WorkflowPhase {
    return this.currentPhase;
  }

  /**
   * 注册事件监听器
   */
  on(event: string, handler: (event: AgentEvent) => void): void {
    const handlers = this.eventListeners.get(event) || [];
    handlers.push(handler);
    this.eventListeners.set(event, handlers);
  }

  /**
   * 移除事件监听器
   */
  off(event: string, handler: (event: AgentEvent) => void): void {
    const handlers = this.eventListeners.get(event) || [];
    const index = handlers.indexOf(handler);
    if (index >= 0) {
      handlers.splice(index, 1);
    }
  }

  /**
   * 发射事件
   */
  private emit(event: string, data: Record<string, unknown>): void {
    const handlers = this.eventListeners.get(event) || [];
    const agentEvent: AgentEvent = {
      type: event,
      data,
      timestamp: Date.now(),
    };

    for (const handler of handlers) {
      try {
        handler(agentEvent);
      } catch (error) {
        console.error(`Event handler error for ${event}:`, error);
      }
    }
  }

  /**
   * 格式化完整响应
   */
  formatResponse(result: WorkflowResult): string {
    const lines: string[] = [];

    // 语义理解
    if (result.semanticExtraction) {
      lines.push("【理解您的需求】");
      lines.push(`意图: ${result.semanticExtraction.intent}`);
      lines.push(`置信度: ${(result.semanticExtraction.confidence * 100).toFixed(0)}%`);
      lines.push("");
    }

    // 诊断结果
    if (result.diagnosis) {
      lines.push(diagnosticEngine.formatDiagnosis(result.diagnosis));
      lines.push("");
    }

    // 解决方案
    if (result.solution) {
      lines.push(solutionBuilder.formatSolution(result.solution));
    }

    // 澄清请求
    if (result.clarificationNeeded) {
      lines.push("【需要更多信息】");
      lines.push(result.clarificationNeeded.message);
      lines.push("");
      lines.push("建议：");
      result.clarificationNeeded.suggestions.forEach((s, i) => {
        lines.push(`  ${i + 1}. ${s}`);
      });
    }

    // 错误
    if (result.error) {
      lines.push("【发生错误】");
      lines.push(result.error.message);
    }

    return lines.join("\n");
  }

  /**
   * 更新配置
   */
  updateConfig(config: Partial<OrchestratorConfig>): void {
    this.config = { ...this.config, ...config };
  }

  /**
   * 获取配置
   */
  getConfig(): OrchestratorConfig {
    return { ...this.config };
  }
}

// ========== 单例导出 ==========

export const agentOrchestrator = new AgentOrchestrator();

export default AgentOrchestrator;
