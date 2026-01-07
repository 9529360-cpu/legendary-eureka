/**
 * 工具注册中心实现
 *
 * 从 AgentCore.ts 抽取，提供工具的注册、查询和管理功能
 *
 * @packageDocumentation
 */

import type { Tool } from "../types";

/**
 * ToolRegistry - 工具注册中心
 *
 * Agent 通过这里获取可用的工具
 * 支持动态注册/注销工具
 */
export class ToolRegistry {
  private tools: Map<string, Tool> = new Map();
  private categories: Map<string, Set<string>> = new Map();

  /**
   * 注册工具（带验证）
   */
  register(tool: Tool): void {
    // 防御性检查：确保工具对象有效
    if (!tool) {
      console.warn("[ToolRegistry] 跳过无效工具: tool is null/undefined");
      return;
    }

    // 确保必需字段存在
    if (!tool.name) {
      console.warn("[ToolRegistry] 跳过无效工具: 缺少 name 字段");
      return;
    }

    // 自动修复缺失的可选字段
    const safeTool: Tool = {
      ...tool,
      description: tool.description || `工具: ${tool.name}`,
      category: tool.category || "utility",
      parameters: tool.parameters || [],
    };

    // 验证参数数组中的每个参数也有 description
    safeTool.parameters = safeTool.parameters.map((p) => ({
      ...p,
      description: p.description || `参数: ${p.name || "unknown"}`,
    }));

    this.tools.set(safeTool.name, safeTool);

    // 更新分类索引
    if (!this.categories.has(safeTool.category)) {
      this.categories.set(safeTool.category, new Set());
    }
    this.categories.get(safeTool.category)!.add(safeTool.name);

    console.log(`[ToolRegistry] Registered tool: ${safeTool.name} (${safeTool.category})`);
  }

  /**
   * 批量注册工具
   */
  registerMany(tools: Tool[]): void {
    tools.forEach((t) => this.register(t));
  }

  /**
   * 注销工具
   */
  unregister(toolName: string): boolean {
    const tool = this.tools.get(toolName);
    if (tool) {
      this.tools.delete(toolName);
      this.categories.get(tool.category)?.delete(toolName);
      return true;
    }
    return false;
  }

  /**
   * 获取工具
   */
  get(toolName: string): Tool | undefined {
    return this.tools.get(toolName);
  }

  /**
   * 获取工具（别名，兼容 SelfReflection.ToolRegistry 接口）
   */
  getTool(toolName: string): Tool | undefined {
    return this.tools.get(toolName);
  }

  /**
   * 获取所有工具
   */
  getAll(): Tool[] {
    return Array.from(this.tools.values());
  }

  /**
   * 获取所有工具（别名，兼容 SelfReflection.ToolRegistry 接口）
   */
  getAllTools(): Tool[] {
    return Array.from(this.tools.values());
  }

  /**
   * 按分类获取工具
   */
  getByCategory(category: string): Tool[] {
    const toolNames = this.categories.get(category);
    if (!toolNames) return [];
    return Array.from(toolNames)
      .map((name) => this.tools.get(name)!)
      .filter(Boolean);
  }

  /**
   * 获取所有分类
   */
  getCategories(): string[] {
    return Array.from(this.categories.keys());
  }

  /**
   * 生成工具描述（给 LLM 用）
   */
  generateToolsDescription(): string {
    const tools = this.getAll();
    return tools
      .map((tool) => {
        const params = (tool.parameters || [])
          .map(
            (p) =>
              `  - ${p.name || "unknown"}: ${p.description || "无描述"}${p.required ? " (必需)" : ""}`
          )
          .join("\n");
        return `**${tool.name || "unknown"}** [${tool.category || "utility"}]\n${tool.description || "无描述"}\n参数:\n${params}`;
      })
      .join("\n\n");
  }

  /**
   * 列出所有工具名称
   */
  list(): string[] {
    return Array.from(this.tools.keys());
  }

  /**
   * 检查工具是否存在
   */
  has(toolName: string): boolean {
    return this.tools.has(toolName);
  }

  /**
   * 获取工具数量
   */
  get size(): number {
    return this.tools.size;
  }

  /**
   * 清空所有工具
   */
  clear(): void {
    this.tools.clear();
    this.categories.clear();
  }

  /**
   * 根据关键词搜索工具
   */
  search(keyword: string): Tool[] {
    const lowerKeyword = keyword.toLowerCase();
    return this.getAll().filter(
      (tool) =>
        (tool.name || "").toLowerCase().includes(lowerKeyword) ||
        (tool.description || "").toLowerCase().includes(lowerKeyword) ||
        (tool.category || "").toLowerCase().includes(lowerKeyword)
    );
  }

  /**
   * 获取工具的 JSON Schema 格式（用于 LLM function calling）
   */
  getToolSchema(toolName: string): object | null {
    const tool = this.get(toolName);
    if (!tool) return null;

    const params = tool.parameters || [];

    return {
      name: tool.name || toolName,
      description: tool.description || `工具: ${toolName}`,
      parameters: {
        type: "object",
        properties: Object.fromEntries(
          params.map((p) => [
            p.name || "param",
            {
              type: p.type || "string",
              description: p.description || "无描述",
              ...(p.default !== undefined && { default: p.default }),
            },
          ])
        ),
        required: params.filter((p) => p.required).map((p) => p.name || "param"),
      },
    };
  }

  /**
   * 获取所有工具的 JSON Schema 格式
   */
  getAllToolSchemas(): object[] {
    return this.getAll()
      .map((t) => this.getToolSchema(t.name))
      .filter(Boolean) as object[];
  }
}
