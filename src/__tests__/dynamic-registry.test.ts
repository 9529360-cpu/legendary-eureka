/**
 * DynamicToolRegistry 测试
 *
 * 覆盖：
 * - 动态注册/注销
 * - 工具查询
 * - 插件管理
 * - 事件通知
 */

import {
  DynamicToolRegistry,
  ToolRegistrationOptions,
  ToolPlugin,
} from "../core/DynamicToolRegistry";
import { Tool, ToolResult } from "../agent/AgentCore";

// 模拟工具工厂
const createMockTool = (name: string, category: string = "test"): Tool => ({
  name,
  description: `Mock tool: ${name}`,
  category,
  parameters: [],
  execute: async (): Promise<ToolResult> => ({
    success: true,
    output: `Executed: ${name}`,
  }),
});

describe("DynamicToolRegistry", () => {
  beforeEach(() => {
    DynamicToolRegistry.reset();
  });

  describe("工具注册", () => {
    it("应该注册单个工具", () => {
      const tool = createMockTool("my_tool");
      const result = DynamicToolRegistry.register(tool);

      expect(result).toBe(true);
      expect(DynamicToolRegistry.has("my_tool")).toBe(true);
    });

    it("应该批量注册工具", () => {
      const tools = [createMockTool("tool_1"), createMockTool("tool_2"), createMockTool("tool_3")];

      const result = DynamicToolRegistry.registerAll(tools);

      expect(result.registered).toBe(3);
      expect(result.failed).toBe(0);
      expect(DynamicToolRegistry.getAll().length).toBe(3);
    });

    it("应该拒绝重复注册（无覆盖选项）", () => {
      const tool = createMockTool("duplicate");

      DynamicToolRegistry.register(tool);
      const result = DynamicToolRegistry.register(tool);

      expect(result).toBe(false);
    });

    it("应该允许覆盖注册", () => {
      const tool1 = createMockTool("overwrite_me");
      const tool2 = { ...createMockTool("overwrite_me"), description: "Updated" };

      DynamicToolRegistry.register(tool1);
      const result = DynamicToolRegistry.register(tool2, { overwrite: true });

      expect(result).toBe(true);
      expect(DynamicToolRegistry.get("overwrite_me")!.description).toBe("Updated");
    });

    it("应该使用命名空间注册", () => {
      const tool = createMockTool("namespaced_tool");

      DynamicToolRegistry.register(tool, { namespace: "excel" });

      expect(DynamicToolRegistry.has("excel.namespaced_tool")).toBe(true);
    });
  });

  describe("工具注销", () => {
    it("应该注销已注册的工具", () => {
      const tool = createMockTool("to_remove");
      DynamicToolRegistry.register(tool);

      const result = DynamicToolRegistry.unregister("to_remove");

      expect(result).toBe(true);
      expect(DynamicToolRegistry.has("to_remove")).toBe(false);
    });

    it("应该返回 false 当注销不存在的工具", () => {
      const result = DynamicToolRegistry.unregister("nonexistent");

      expect(result).toBe(false);
    });

    it("应该批量注销满足条件的工具", () => {
      DynamicToolRegistry.registerAll([
        createMockTool("cat_a_1", "category_a"),
        createMockTool("cat_a_2", "category_a"),
        createMockTool("cat_b_1", "category_b"),
      ]);

      const count = DynamicToolRegistry.unregisterAll((rt) => rt.tool.category === "category_a");

      expect(count).toBe(2);
      expect(DynamicToolRegistry.getAll().length).toBe(1);
    });
  });

  describe("工具查询", () => {
    beforeEach(() => {
      DynamicToolRegistry.registerAll(
        [
          createMockTool("excel_read", "excel"),
          createMockTool("excel_write", "excel"),
          createMockTool("word_read", "word"),
        ],
        { group: "office" }
      );

      DynamicToolRegistry.register(createMockTool("api_call", "api"), {
        group: "external",
        tags: ["network", "async"],
      });
    });

    it("应该按名称查询", () => {
      const results = DynamicToolRegistry.query({ name: "excel" });

      expect(results.length).toBe(2);
    });

    it("应该按分类查询", () => {
      const results = DynamicToolRegistry.query({ category: "excel" });

      expect(results.length).toBe(2);
    });

    it("应该按分组查询", () => {
      const results = DynamicToolRegistry.query({ group: "office" });

      expect(results.length).toBe(3);
    });

    it("应该按标签查询", () => {
      const results = DynamicToolRegistry.query({ tags: ["network"] });

      expect(results.length).toBe(1);
      expect(results[0].name).toBe("api_call");
    });

    it("应该搜索工具", () => {
      const results = DynamicToolRegistry.search("read");

      expect(results.length).toBe(2);
    });
  });

  describe("工具状态管理", () => {
    it("应该禁用工具", () => {
      const tool = createMockTool("to_disable");
      DynamicToolRegistry.register(tool);

      DynamicToolRegistry.disable("to_disable");

      expect(DynamicToolRegistry.isEnabled("to_disable")).toBe(false);
      expect(DynamicToolRegistry.get("to_disable")).toBeUndefined();
    });

    it("应该启用工具", () => {
      const tool = createMockTool("to_enable");
      DynamicToolRegistry.register(tool, { enabled: false });

      DynamicToolRegistry.enable("to_enable");

      expect(DynamicToolRegistry.isEnabled("to_enable")).toBe(true);
    });

    it("应该标记工具为废弃", () => {
      const tool = createMockTool("old_tool");
      DynamicToolRegistry.register(tool);

      DynamicToolRegistry.deprecate("old_tool", "new_tool");

      const info = DynamicToolRegistry.getInfo("old_tool");
      expect(info!.status).toBe("deprecated");
    });

    it("应该记录使用统计", () => {
      const tool = createMockTool("usage_tool");
      DynamicToolRegistry.register(tool);

      DynamicToolRegistry.recordUsage("usage_tool");
      DynamicToolRegistry.recordUsage("usage_tool");
      DynamicToolRegistry.recordUsage("usage_tool");

      const info = DynamicToolRegistry.getInfo("usage_tool");
      expect(info!.usageCount).toBe(3);
      expect(info!.lastUsed).toBeDefined();
    });
  });

  describe("插件管理", () => {
    const createMockPlugin = (id: string, toolCount: number): ToolPlugin => ({
      id,
      name: `Plugin ${id}`,
      version: "1.0.0",
      tools: Array.from({ length: toolCount }, (_, i) => createMockTool(`${id}_tool_${i}`, id)),
    });

    it("应该加载插件", async () => {
      const plugin = createMockPlugin("test_plugin", 3);

      const result = await DynamicToolRegistry.loadPlugin(plugin);

      expect(result).toBe(true);
      expect(DynamicToolRegistry.getPlugins().length).toBe(1);
      expect(DynamicToolRegistry.has("test_plugin.test_plugin_tool_0")).toBe(true);
    });

    it("应该卸载插件", async () => {
      const plugin = createMockPlugin("removable", 2);
      await DynamicToolRegistry.loadPlugin(plugin);

      const result = await DynamicToolRegistry.unloadPlugin("removable");

      expect(result).toBe(true);
      expect(DynamicToolRegistry.getPlugins().length).toBe(0);
      expect(DynamicToolRegistry.has("removable.removable_tool_0")).toBe(false);
    });

    it("应该拒绝重复加载相同插件", async () => {
      const plugin = createMockPlugin("unique", 1);

      await DynamicToolRegistry.loadPlugin(plugin);
      const result = await DynamicToolRegistry.loadPlugin(plugin);

      expect(result).toBe(false);
    });
  });

  describe("事件通知", () => {
    it("应该在注册时触发事件", () => {
      const events: any[] = [];
      DynamicToolRegistry.addEventListener((e) => events.push(e));

      DynamicToolRegistry.register(createMockTool("event_tool"));

      expect(events.length).toBe(1);
      expect(events[0].type).toBe("registered");
      expect(events[0].toolName).toBe("event_tool");
    });

    it("应该在注销时触发事件", () => {
      const events: any[] = [];
      DynamicToolRegistry.register(createMockTool("to_unreg"));
      DynamicToolRegistry.addEventListener((e) => events.push(e));

      DynamicToolRegistry.unregister("to_unreg");

      expect(events.length).toBe(1);
      expect(events[0].type).toBe("unregistered");
    });

    it("应该在启用/禁用时触发事件", () => {
      const events: any[] = [];
      DynamicToolRegistry.register(createMockTool("toggle_tool"));
      DynamicToolRegistry.addEventListener((e) => events.push(e));

      DynamicToolRegistry.disable("toggle_tool");
      DynamicToolRegistry.enable("toggle_tool");

      expect(events.length).toBe(2);
      expect(events[0].type).toBe("disabled");
      expect(events[1].type).toBe("enabled");
    });
  });

  describe("统计与诊断", () => {
    beforeEach(() => {
      DynamicToolRegistry.registerAll([
        createMockTool("stat_1", "cat_a"),
        createMockTool("stat_2", "cat_a"),
        createMockTool("stat_3", "cat_b"),
      ]);

      DynamicToolRegistry.disable("stat_3");
      DynamicToolRegistry.deprecate("stat_2");

      DynamicToolRegistry.recordUsage("stat_1");
      DynamicToolRegistry.recordUsage("stat_1");
    });

    it("应该返回正确的统计信息", () => {
      const stats = DynamicToolRegistry.getStatistics();

      expect(stats.totalTools).toBe(3);
      expect(stats.enabledTools).toBe(2);
      expect(stats.disabledTools).toBe(1);
      expect(stats.deprecatedTools).toBe(1);
      expect(stats.categories).toContain("cat_a");
      expect(stats.categories).toContain("cat_b");
    });

    it("应该返回使用最多的工具", () => {
      const stats = DynamicToolRegistry.getStatistics();

      expect(stats.topUsed.length).toBe(1);
      expect(stats.topUsed[0].name).toBe("stat_1");
      expect(stats.topUsed[0].count).toBe(2);
    });

    it("应该执行健康检查", () => {
      const health = DynamicToolRegistry.healthCheck();

      expect(health.healthy).toBe(true);
      expect(health.warnings.length).toBeGreaterThan(0);
    });
  });
});
