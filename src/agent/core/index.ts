/**
 * Agent Core 模块化架构 v4.1
 *
 * 将原 AgentCore.ts (16000+ 行) 拆分成职责单一的模块
 *
 * 架构：
 * ┌─────────────────────────────────────────────────────────────┐
 * │  AgentOrchestrator (编排器，<300行)                          │
 * │  - 协调各模块工作流                                          │
 * │  - 事件发布/订阅                                             │
 * ├─────────────────────────────────────────────────────────────┤
 * │  IntentParser    SpecCompiler    AgentExecutor              │
 * │  (解析层)        (编译层)         (执行层)                   │
 * ├─────────────────────────────────────────────────────────────┤
 * │  ToolContract (工具契约层)                                   │
 * │  - 模型无关的工具定义                                        │
 * │  - 适配器: OpenAI / Claude / Gemini / DeepSeek              │
 * ├─────────────────────────────────────────────────────────────┤
 * │  DiagnosticEngine  SemanticExtractor  SolutionBuilder       │
 * │  (诊断引擎)        (语义抽取)          (方案生成)            │
 * └─────────────────────────────────────────────────────────────┘
 */

// ========== 核心类型导出 ==========
export * from "./types";

// ========== 工具契约层 ==========
export * from "./contracts/ToolContract";

// ========== 语义处理层 ==========
export * from "./semantic/SemanticExtractor";
export * from "./semantic/DiagnosticEngine";

// ========== 方案生成层 ==========
export * from "./solutions/SolutionBuilder";

// ========== 编排层 ==========
export * from "./AgentOrchestrator";

// ========== 模型适配器层 ==========
export * from "./adapters";
