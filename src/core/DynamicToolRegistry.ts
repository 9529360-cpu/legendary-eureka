/**
 * DynamicToolRegistry - 动态工具注册表
 * v1.0.0
 *
 * 功能：
 * 1. 工具动态注册/注销（热插拔）
 * 2. 工具生命周期管理
 * 3. 工具分组与命名空间
 * 4. 工具依赖解析
 * 5. 工具版本管理
 *
 * 解决的问题：
 * - ToolRegistry未实现动态注册/热插拔机制
 * - 新增工具需手动修改多处代码
 * - 工具扩展性不足
 */

import { Tool, ToolResult, ToolParameter } from "../agent/AgentCore";
import { ToolProtocol, ToolCapability, ToolRiskLevel, ToolEnvironment } from "./ToolProtocol";
import { TaskExecutionMonitor } from "./TaskExecutionMonitor";
import { Logger } from "../utils/Logger";

// ========== 类型定义 ==========

/**
 * 工具注册选项
 */
export interface ToolRegistrationOptions {
  /** 是否覆盖已存在的工具 */
  overwrite?: boolean;
  /** 命名空间 */
  namespace?: string;
  /** 分组 */
  group?: string;
  /** 标签 */
  tags?: string[];
  /** 优先级（用于同名工具选择） */
  priority?: number;
  /** 是否启用 */
  enabled?: boolean;
  /** 依赖的其他工具 */
  dependencies?: string[];
}

/**
 * 注册的工具信息
 */
export interface RegisteredTool {
  tool: Tool;
  protocol?: ToolProtocol;
  options: ToolRegistrationOptions;
  registeredAt: Date;
  lastUsed?: Date;
  usageCount: number;
  enabled: boolean;
  status: "active" | "disabled" | "deprecated" | "error";
}

/**
 * 工具查询条件
 */
export interface ToolQuery {
  name?: string;
  namespace?: string;
  group?: string;
  category?: string;
  capabilities?: ToolCapability[];
  tags?: string[];
  enabled?: boolean;
  riskLevel?: ToolRiskLevel;
}

/**
 * 工具注册事件
 */
export interface ToolRegistryEvent {
  type: "registered" | "unregistered" | "enabled" | "disabled" | "updated";
  toolName: string;
  timestamp: Date;
  details?: Record<string, unknown>;
}

/**
 * 工具插件接口
 */
export interface ToolPlugin {
  id: string;
  name: string;
  version: string;
  description?: string;
  tools: Tool[];
  initialize?: () => Promise<void>;
  destroy?: () => Promise<void>;
}

// ========== 动态注册表实现 ==========

/**
 * 动态工具注册表
 */
class DynamicToolRegistryClass {
  private tools: Map<string, RegisteredTool> = new Map();
  private plugins: Map<string, ToolPlugin> = new Map();
  private eventListeners: ((event: ToolRegistryEvent) => void)[] = [];
  private namespaces: Set<string> = new Set();
  private groups: Set<string> = new Set();

  // ========== 工具注册 ==========

  /**
   * 注册单个工具
   */
  register(tool: Tool, options: ToolRegistrationOptions = {}): boolean {
    const fullName = this.getFullToolName(tool.name, options.namespace);

    // 检查是否已存在
    if (this.tools.has(fullName) && !options.overwrite) {
      Logger.warn("DynamicRegistry", `工具已存在且未设置覆盖: ${fullName}`);
      return false;
    }

    // 检查依赖
    if (options.dependencies) {
      const missingDeps = options.dependencies.filter((dep) => !this.tools.has(dep));
      if (missingDeps.length > 0) {
        Logger.warn("DynamicRegistry", `工具 ${fullName} 缺少依赖: ${missingDeps.join(", ")}`);
        // 仍然注册，但标记状态
      }
    }

    const registeredTool: RegisteredTool = {
      tool,
      options: {
        ...options,
        priority: options.priority ?? 0,
        enabled: options.enabled ?? true,
      },
      registeredAt: new Date(),
      usageCount: 0,
      enabled: options.enabled ?? true,
      status: "active",
    };

    this.tools.set(fullName, registeredTool);

    // 更新命名空间和分组
    if (options.namespace) {
      this.namespaces.add(options.namespace);
    }
    if (options.group) {
      this.groups.add(options.group);
    }

    // 同步到任务监控器
    TaskExecutionMonitor.registerTool(fullName);

    // 发送事件
    this.emitEvent({
      type: "registered",
      toolName: fullName,
      timestamp: new Date(),
      details: { options },
    });

    Logger.info("DynamicRegistry", `工具已注册: ${fullName}`);
    return true;
  }

  /**
   * 批量注册工具
   */
  registerAll(
    tools: Tool[],
    options: ToolRegistrationOptions = {}
  ): {
    registered: number;
    failed: number;
    errors: string[];
  } {
    let registered = 0;
    let failed = 0;
    const errors: string[] = [];

    tools.forEach((tool) => {
      try {
        if (this.register(tool, options)) {
          registered++;
        } else {
          failed++;
          errors.push(`${tool.name}: 注册失败（可能已存在）`);
        }
      } catch (error) {
        failed++;
        errors.push(`${tool.name}: ${(error as Error).message}`);
      }
    });

    Logger.info("DynamicRegistry", `批量注册完成: ${registered} 成功, ${failed} 失败`);
    return { registered, failed, errors };
  }

  /**
   * 注销工具
   */
  unregister(toolName: string): boolean {
    if (this.tools.has(toolName)) {
      this.tools.delete(toolName);

      this.emitEvent({
        type: "unregistered",
        toolName,
        timestamp: new Date(),
      });

      Logger.info("DynamicRegistry", `工具已注销: ${toolName}`);
      return true;
    }
    return false;
  }

  /**
   * 批量注销
   */
  unregisterAll(predicate: (tool: RegisteredTool) => boolean): number {
    let count = 0;
    const toRemove: string[] = [];

    this.tools.forEach((rt, name) => {
      if (predicate(rt)) {
        toRemove.push(name);
      }
    });

    toRemove.forEach((name) => {
      if (this.unregister(name)) {
        count++;
      }
    });

    return count;
  }

  // ========== 工具查询 ==========

  /**
   * 获取工具
   */
  get(toolName: string): Tool | undefined {
    const rt = this.tools.get(toolName);
    if (rt && rt.enabled) {
      return rt.tool;
    }
    return undefined;
  }

  /**
   * 获取完整工具信息
   */
  getInfo(toolName: string): RegisteredTool | undefined {
    return this.tools.get(toolName);
  }

  /**
   * 检查工具是否存在
   */
  has(toolName: string): boolean {
    return this.tools.has(toolName);
  }

  /**
   * 检查工具是否启用
   */
  isEnabled(toolName: string): boolean {
    const rt = this.tools.get(toolName);
    return rt?.enabled ?? false;
  }

  /**
   * 获取所有工具
   */
  getAll(includeDisabled: boolean = false): Tool[] {
    return Array.from(this.tools.values())
      .filter((rt) => includeDisabled || rt.enabled)
      .map((rt) => rt.tool);
  }

  /**
   * 获取所有工具信息
   */
  getAllInfo(includeDisabled: boolean = false): RegisteredTool[] {
    return Array.from(this.tools.values()).filter((rt) => includeDisabled || rt.enabled);
  }

  /**
   * 查询工具
   */
  query(conditions: ToolQuery): Tool[] {
    return Array.from(this.tools.values())
      .filter((rt) => {
        if (conditions.enabled !== undefined && rt.enabled !== conditions.enabled) {
          return false;
        }
        if (conditions.name && !rt.tool.name.includes(conditions.name)) {
          return false;
        }
        if (conditions.namespace && rt.options.namespace !== conditions.namespace) {
          return false;
        }
        if (conditions.group && rt.options.group !== conditions.group) {
          return false;
        }
        if (conditions.category && rt.tool.category !== conditions.category) {
          return false;
        }
        if (conditions.tags && conditions.tags.length > 0) {
          const toolTags = rt.options.tags || [];
          if (!conditions.tags.some((t) => toolTags.includes(t))) {
            return false;
          }
        }
        return true;
      })
      .map((rt) => rt.tool);
  }

  /**
   * 按命名空间获取
   */
  getByNamespace(namespace: string): Tool[] {
    return this.query({ namespace });
  }

  /**
   * 按分组获取
   */
  getByGroup(group: string): Tool[] {
    return this.query({ group });
  }

  /**
   * 按分类获取
   */
  getByCategory(category: string): Tool[] {
    return this.query({ category });
  }

  /**
   * 搜索工具
   */
  search(keyword: string): Tool[] {
    const lower = keyword.toLowerCase();
    return Array.from(this.tools.values())
      .filter(
        (rt) =>
          rt.enabled &&
          (rt.tool.name.toLowerCase().includes(lower) ||
            rt.tool.description.toLowerCase().includes(lower))
      )
      .map((rt) => rt.tool);
  }

  // ========== 工具状态管理 ==========

  /**
   * 启用工具
   */
  enable(toolName: string): boolean {
    const rt = this.tools.get(toolName);
    if (rt) {
      rt.enabled = true;
      rt.status = "active";

      this.emitEvent({
        type: "enabled",
        toolName,
        timestamp: new Date(),
      });

      Logger.info("DynamicRegistry", `工具已启用: ${toolName}`);
      return true;
    }
    return false;
  }

  /**
   * 禁用工具
   */
  disable(toolName: string): boolean {
    const rt = this.tools.get(toolName);
    if (rt) {
      rt.enabled = false;
      rt.status = "disabled";

      this.emitEvent({
        type: "disabled",
        toolName,
        timestamp: new Date(),
      });

      Logger.info("DynamicRegistry", `工具已禁用: ${toolName}`);
      return true;
    }
    return false;
  }

  /**
   * 标记工具为已废弃
   */
  deprecate(toolName: string, replacement?: string): boolean {
    const rt = this.tools.get(toolName);
    if (rt) {
      rt.status = "deprecated";

      this.emitEvent({
        type: "updated",
        toolName,
        timestamp: new Date(),
        details: { deprecated: true, replacement },
      });

      Logger.warn(
        "DynamicRegistry",
        `工具已废弃: ${toolName}${replacement ? `, 建议使用 ${replacement}` : ""}`
      );
      return true;
    }
    return false;
  }

  /**
   * 更新工具使用统计
   */
  recordUsage(toolName: string): void {
    const rt = this.tools.get(toolName);
    if (rt) {
      rt.usageCount++;
      rt.lastUsed = new Date();
    }
  }

  // ========== 插件管理 ==========

  /**
   * 加载插件
   */
  async loadPlugin(plugin: ToolPlugin): Promise<boolean> {
    if (this.plugins.has(plugin.id)) {
      Logger.warn("DynamicRegistry", `插件已加载: ${plugin.id}`);
      return false;
    }

    try {
      // 初始化插件
      if (plugin.initialize) {
        await plugin.initialize();
      }

      // 注册插件中的工具
      this.registerAll(plugin.tools, {
        namespace: plugin.id,
        group: `plugin:${plugin.id}`,
      });

      this.plugins.set(plugin.id, plugin);
      Logger.info("DynamicRegistry", `插件已加载: ${plugin.id} (${plugin.tools.length} 个工具)`);
      return true;
    } catch (error) {
      Logger.error("DynamicRegistry", `插件加载失败: ${plugin.id}`, { error });
      return false;
    }
  }

  /**
   * 卸载插件
   */
  async unloadPlugin(pluginId: string): Promise<boolean> {
    const plugin = this.plugins.get(pluginId);
    if (!plugin) {
      return false;
    }

    try {
      // 销毁插件
      if (plugin.destroy) {
        await plugin.destroy();
      }

      // 注销插件中的工具
      this.unregisterAll((rt) => rt.options.namespace === pluginId);

      this.plugins.delete(pluginId);
      Logger.info("DynamicRegistry", `插件已卸载: ${pluginId}`);
      return true;
    } catch (error) {
      Logger.error("DynamicRegistry", `插件卸载失败: ${pluginId}`, { error });
      return false;
    }
  }

  /**
   * 获取已加载的插件
   */
  getPlugins(): ToolPlugin[] {
    return Array.from(this.plugins.values());
  }

  // ========== 工具包装器 ==========

  /**
   * 创建带监控的工具包装器
   */
  wrapWithMonitoring(tool: Tool, taskId?: string): Tool {
    const registry = this;

    return {
      ...tool,
      execute: async (input: Record<string, unknown>): Promise<ToolResult> => {
        const id = taskId || `auto_${Date.now()}`;
        TaskExecutionMonitor.startToolCall(id, tool.name, input);

        try {
          const result = await tool.execute(input);
          TaskExecutionMonitor.completeToolCall(id, tool.name, result.output, result.success);
          registry.recordUsage(tool.name);
          return result;
        } catch (error) {
          TaskExecutionMonitor.failToolCall(id, tool.name, (error as Error).message);
          throw error;
        }
      },
    };
  }

  /**
   * 创建带重试的工具包装器
   */
  wrapWithRetry(tool: Tool, maxRetries: number = 3, delay: number = 1000): Tool {
    return {
      ...tool,
      execute: async (input: Record<string, unknown>): Promise<ToolResult> => {
        let lastError: Error | undefined;

        for (let attempt = 0; attempt <= maxRetries; attempt++) {
          try {
            return await tool.execute(input);
          } catch (error) {
            lastError = error as Error;
            if (attempt < maxRetries) {
              await new Promise((resolve) => setTimeout(resolve, delay * (attempt + 1)));
            }
          }
        }

        throw lastError;
      },
    };
  }

  // ========== 事件管理 ==========

  /**
   * 添加事件监听器
   */
  addEventListener(listener: (event: ToolRegistryEvent) => void): () => void {
    this.eventListeners.push(listener);
    return () => {
      this.eventListeners = this.eventListeners.filter((l) => l !== listener);
    };
  }

  /**
   * 发送事件
   */
  private emitEvent(event: ToolRegistryEvent): void {
    this.eventListeners.forEach((listener) => {
      try {
        listener(event);
      } catch (error) {
        Logger.error("DynamicRegistry", "事件监听器执行失败", { error });
      }
    });
  }

  // ========== 统计与诊断 ==========

  /**
   * 获取统计信息
   */
  getStatistics(): {
    totalTools: number;
    enabledTools: number;
    disabledTools: number;
    deprecatedTools: number;
    namespaces: string[];
    groups: string[];
    categories: string[];
    topUsed: Array<{ name: string; count: number }>;
    plugins: number;
  } {
    const all = Array.from(this.tools.values());
    const categories = new Set<string>();

    all.forEach((rt) => {
      if (rt.tool.category) {
        categories.add(rt.tool.category);
      }
    });

    const topUsed = all
      .filter((rt) => rt.usageCount > 0)
      .sort((a, b) => b.usageCount - a.usageCount)
      .slice(0, 10)
      .map((rt) => ({ name: rt.tool.name, count: rt.usageCount }));

    return {
      totalTools: all.length,
      enabledTools: all.filter((rt) => rt.enabled).length,
      disabledTools: all.filter((rt) => !rt.enabled).length,
      deprecatedTools: all.filter((rt) => rt.status === "deprecated").length,
      namespaces: Array.from(this.namespaces),
      groups: Array.from(this.groups),
      categories: Array.from(categories),
      topUsed,
      plugins: this.plugins.size,
    };
  }

  /**
   * 健康检查
   */
  healthCheck(): {
    healthy: boolean;
    issues: string[];
    warnings: string[];
  } {
    const issues: string[] = [];
    const warnings: string[] = [];

    // 检查是否有工具
    if (this.tools.size === 0) {
      issues.push("没有注册任何工具");
    }

    // 检查禁用的工具
    const disabled = Array.from(this.tools.values()).filter((rt) => !rt.enabled);
    if (disabled.length > 0) {
      warnings.push(`${disabled.length} 个工具被禁用`);
    }

    // 检查废弃的工具
    const deprecated = Array.from(this.tools.values()).filter((rt) => rt.status === "deprecated");
    if (deprecated.length > 0) {
      warnings.push(`${deprecated.length} 个工具已废弃`);
    }

    // 检查错误状态的工具
    const errored = Array.from(this.tools.values()).filter((rt) => rt.status === "error");
    if (errored.length > 0) {
      issues.push(`${errored.length} 个工具处于错误状态`);
    }

    return {
      healthy: issues.length === 0,
      issues,
      warnings,
    };
  }

  /**
   * 导出工具列表
   */
  export(): Array<{ name: string; category: string; description: string; enabled: boolean }> {
    return Array.from(this.tools.values()).map((rt) => ({
      name: rt.tool.name,
      category: rt.tool.category,
      description: rt.tool.description,
      enabled: rt.enabled,
    }));
  }

  // ========== 辅助方法 ==========

  /**
   * 获取完整工具名
   */
  private getFullToolName(name: string, namespace?: string): string {
    return namespace ? `${namespace}.${name}` : name;
  }

  /**
   * 清空注册表
   */
  clear(): void {
    this.tools.clear();
    this.namespaces.clear();
    this.groups.clear();
    Logger.info("DynamicRegistry", "注册表已清空");
  }

  /**
   * 重置（用于测试）
   */
  reset(): void {
    this.clear();
    this.plugins.clear();
    this.eventListeners = [];
  }
}

// 导出单例
export const DynamicToolRegistry = new DynamicToolRegistryClass();

// 便捷方法导出
export const registry = {
  register: (tool: Tool, options?: ToolRegistrationOptions) =>
    DynamicToolRegistry.register(tool, options),
  registerAll: (tools: Tool[], options?: ToolRegistrationOptions) =>
    DynamicToolRegistry.registerAll(tools, options),
  unregister: (name: string) => DynamicToolRegistry.unregister(name),
  get: (name: string) => DynamicToolRegistry.get(name),
  has: (name: string) => DynamicToolRegistry.has(name),
  getAll: (includeDisabled?: boolean) => DynamicToolRegistry.getAll(includeDisabled),
  query: (conditions: ToolQuery) => DynamicToolRegistry.query(conditions),
  enable: (name: string) => DynamicToolRegistry.enable(name),
  disable: (name: string) => DynamicToolRegistry.disable(name),
  loadPlugin: (plugin: ToolPlugin) => DynamicToolRegistry.loadPlugin(plugin),
  unloadPlugin: (id: string) => DynamicToolRegistry.unloadPlugin(id),
  stats: () => DynamicToolRegistry.getStatistics(),
  health: () => DynamicToolRegistry.healthCheck(),
};

export default DynamicToolRegistry;
