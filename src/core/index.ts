/**
 * Core 模块统一导出
 *
 * 本文件导出所有核心增强模块，便于统一引用
 *
 * @version 1.0.0
 */

// ============ 内部导入（用于初始化函数）============
import { SecurityManager } from "./SecurityManager";
import { ConfigManager } from "./ConfigManager";
import { ConversationMemory } from "./ConversationMemory";
import { TraceContext } from "./TraceContext";
import { TaskExecutionMonitor } from "./TaskExecutionMonitor";

// ============ 任务执行与监控 ============

export {
  TaskExecutionMonitor,
  monitor,
  TaskPhase,
  AlertLevel,
  type TaskState,
  type TaskAlert,
  type ToolCallInfo,
  type TaskStatistics,
} from "./TaskExecutionMonitor";

export {
  ToolExecutor,
  executor,
  type ExecutionResult,
  type FallbackConfig,
  type RetryConfig,
  type ExecutionStats,
} from "./ToolExecutor";

// ============ 工具注册与协议 ============

export {
  DynamicToolRegistry,
  registry,
  type ToolRegistrationOptions,
  type RegisteredTool,
  type ToolPlugin,
  type ToolQuery,
  type RegistryEvent,
  type RegistryStatistics,
  type HealthCheckResult,
} from "./DynamicToolRegistry";

export {
  ToolProtocolRegistry,
  protocol,
  ToolCapability,
  ToolRiskLevel,
  ToolEnvironment,
  type ParamProtocol,
  type ReturnProtocol,
  type ExceptionProtocol,
  type ToolProtocol,
  type ToolVersion,
} from "./ToolProtocol";

// ============ 链路追踪 ============

export {
  TraceContext,
  trace,
  SpanType,
  type Trace,
  type Span,
  type SpanEvent,
  type TraceStatistics,
  type TraceTree,
  type TraceTimeline,
} from "./TraceContext";

// ============ 配置管理 ============

export {
  ConfigManager,
  config,
  type Environment,
  type ApiConfig,
  type ExcelConfig,
  type AgentConfig,
  type UiConfig,
  type LoggingConfig,
  type SecurityConfig,
  type FeatureFlags,
  type ConfigChangeListener,
} from "./ConfigManager";

// ============ 对话记忆 ============

export {
  ConversationMemory,
  memory,
  IntentType,
  type MessageRole,
  type ConversationMessage,
  type MessageMetadata,
  type ExtractedEntity,
  type ToolCallRecord,
  type ConversationContext,
  type UserPreferences,
  type TaskContext,
  type Reference,
  type ContextWindowConfig,
  type IntentAnalysis,
} from "./ConversationMemory";

// ============ 安全管理 ============

export {
  SecurityManager,
  security,
  PermissionLevel,
  type ApiRequirement,
  type CompatibilityResult,
  type OperationPermission,
  type ValidationRule,
  type ValidationResult,
  type RateLimitConfig,
} from "./SecurityManager";

// ============ 高级 Excel 功能 ============

export {
  AdvancedExcelFunctions,
  advanced,
  TABLE_STYLE_PRESETS,
  type TableStylePreset,
  type CellStyle,
  type BorderStyle,
  type ConditionalFormatRule,
  type ChartConfig,
  type DataValidationRule,
  type PivotTableConfig,
} from "./AdvancedExcelFunctions";

// ============ 现有模块重导出 ============

export { Logger } from "../utils/Logger";
export { ExcelService } from "./ExcelService";
export { ErrorHandler } from "./ErrorHandler";
export { DataAnalyzer } from "./DataAnalyzer";
export { PromptBuilder } from "./PromptBuilder";

// ============ 便捷初始化 ============

/**
 * 初始化所有增强模块
 */
export async function initializeEnhancements(): Promise<{
  success: boolean;
  compatibility: import("./SecurityManager").CompatibilityResult;
  errors: string[];
}> {
  const errors: string[] = [];

  try {
    // 1. 检查兼容性
    const compatibility = SecurityManager.checkCompatibility();
    if (!compatibility.supported) {
      errors.push(...compatibility.missingApis.map((api) => `缺少必需 API: ${api}`));
    }

    // 2. 加载配置
    try {
      ConfigManager.loadFromStorage();
    } catch (e) {
      errors.push(`配置加载失败: ${(e as Error).message}`);
    }

    // 3. 加载对话历史
    try {
      ConversationMemory.loadFromStorage();
    } catch (e) {
      errors.push(`对话历史加载失败: ${(e as Error).message}`);
    }

    // 4. 注册默认工具（如果需要）
    // 这里可以注册一些内置工具

    return {
      success: errors.length === 0,
      compatibility,
      errors,
    };
  } catch (e) {
    errors.push(`初始化失败: ${(e as Error).message}`);
    return {
      success: false,
      compatibility: SecurityManager.checkCompatibility(),
      errors,
    };
  }
}

/**
 * 清理所有模块状态
 */
export function cleanupEnhancements(): void {
  // 保存配置
  ConfigManager.saveToStorage();

  // 保存对话
  ConversationMemory.saveToStorage();

  // 重置追踪
  TraceContext.reset();

  // 重置监控
  TaskExecutionMonitor.reset();
}
