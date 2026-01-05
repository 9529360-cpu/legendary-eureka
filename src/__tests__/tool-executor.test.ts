/**
 * ToolExecutor 测试
 *
 * 覆盖：
 * - 工具注册与查找
 * - 工具执行与重试
 * - 兜底策略
 * - 参数验证
 */

import { ToolExecutor } from "../core/ToolExecutor";
import { Tool, ToolResult } from "../agent/AgentCore";

// 模拟工具
const createMockTool = (name: string, shouldSucceed: boolean = true, delay: number = 0): Tool => ({
  name,
  description: `Mock tool: ${name}`,
  category: "test",
  parameters: [{ name: "input", type: "string", description: "Test input", required: true }],
  execute: async (input: Record<string, unknown>): Promise<ToolResult> => {
    if (delay > 0) {
      await new Promise((resolve) => setTimeout(resolve, delay));
    }
    if (shouldSucceed) {
      return {
        success: true,
        output: `Success: ${JSON.stringify(input)}`,
        data: input,
      };
    } else {
      throw new Error(`Tool ${name} failed`);
    }
  },
});

describe("ToolExecutor", () => {
  beforeEach(() => {
    ToolExecutor.reset();
  });

  describe("工具注册", () => {
    it("应该成功注册工具", () => {
      const tool = createMockTool("test_tool");
      ToolExecutor.registerTool(tool);

      expect(ToolExecutor.hasTool("test_tool")).toBe(true);
      expect(ToolExecutor.getTool("test_tool")).toBeDefined();
    });

    it("应该批量注册工具", () => {
      const tools = [createMockTool("tool1"), createMockTool("tool2"), createMockTool("tool3")];

      ToolExecutor.registerTools(tools);

      expect(ToolExecutor.getAllTools().length).toBe(3);
      expect(ToolExecutor.hasTool("tool1")).toBe(true);
      expect(ToolExecutor.hasTool("tool2")).toBe(true);
      expect(ToolExecutor.hasTool("tool3")).toBe(true);
    });

    it("应该取消注册工具", () => {
      const tool = createMockTool("to_remove");
      ToolExecutor.registerTool(tool);

      expect(ToolExecutor.hasTool("to_remove")).toBe(true);

      ToolExecutor.unregisterTool("to_remove");

      expect(ToolExecutor.hasTool("to_remove")).toBe(false);
    });
  });

  describe("工具查找", () => {
    beforeEach(() => {
      ToolExecutor.registerTools([
        createMockTool("excel_read_range"),
        createMockTool("excel_write_range"),
        createMockTool("respond_to_user"),
      ]);
    });

    it("应该找到已注册的工具", () => {
      const result = ToolExecutor.lookupTool("excel_read_range");

      expect(result.found).toBe(true);
      expect(result.tool).toBeDefined();
      expect(result.tool!.name).toBe("excel_read_range");
    });

    it("应该返回未找到的工具及建议", () => {
      const result = ToolExecutor.lookupTool("nonexistent_tool");

      expect(result.found).toBe(false);
      expect(result.suggestion).toBeDefined();
    });

    it("应该根据相似名称提供备选工具", () => {
      const result = ToolExecutor.lookupTool("excel_read");

      expect(result.found).toBe(false);
      expect(result.alternatives).toBeDefined();
      expect(result.alternatives!.length).toBeGreaterThan(0);
    });
  });

  describe("工具执行", () => {
    beforeEach(() => {
      ToolExecutor.registerTool(createMockTool("success_tool", true));
      ToolExecutor.registerTool(createMockTool("fail_tool", false));
    });

    it("应该成功执行工具", async () => {
      const result = await ToolExecutor.execute("success_tool", { input: "test" });

      expect(result.success).toBe(true);
      expect(result.toolName).toBe("success_tool");
      expect(result.executionTime).toBeGreaterThanOrEqual(0);
    });

    it("应该处理工具执行失败", async () => {
      const result = await ToolExecutor.execute(
        "fail_tool",
        { input: "test" },
        {
          skipMonitoring: true,
        }
      );

      expect(result.success).toBe(false);
      expect(result.error).toBeDefined();
    });

    it("应该处理未注册的工具", async () => {
      const result = await ToolExecutor.execute(
        "unknown_tool",
        { input: "test" },
        {
          skipMonitoring: true,
        }
      );

      expect(result.success).toBe(false);
      expect(result.error).toContain("TOOL_NOT_FOUND");
    });
  });

  describe("重试机制", () => {
    it("应该在失败后重试", async () => {
      let attempts = 0;
      const retryTool: Tool = {
        name: "retry_tool",
        description: "Tool that fails first then succeeds",
        category: "test",
        parameters: [],
        execute: async (): Promise<ToolResult> => {
          attempts++;
          if (attempts < 3) {
            throw new Error("Not yet");
          }
          return { success: true, output: "Finally!" };
        },
      };

      ToolExecutor.registerTool(retryTool);

      const result = await ToolExecutor.execute(
        "retry_tool",
        {},
        {
          retryOnFailure: true,
          maxRetries: 3,
          retryDelay: 10,
          skipMonitoring: true,
        }
      );

      expect(result.success).toBe(true);
      expect(result.retryCount).toBe(2);
    });

    it("应该在所有重试失败后返回错误", async () => {
      ToolExecutor.registerTool(createMockTool("always_fail", false));

      const result = await ToolExecutor.execute(
        "always_fail",
        { input: "test" },
        {
          retryOnFailure: true,
          maxRetries: 2,
          retryDelay: 10,
          skipMonitoring: true,
        }
      );

      expect(result.success).toBe(false);
      expect(result.retryCount).toBe(2);
    });
  });

  describe("兜底策略", () => {
    beforeEach(() => {
      ToolExecutor.registerTool(createMockTool("respond_to_user", true));
    });

    it("应该使用兜底工具当主工具未注册", async () => {
      const result = await ToolExecutor.execute(
        "nonexistent",
        { input: "test" },
        {
          fallbackTools: ["respond_to_user"],
          skipMonitoring: true,
        }
      );

      expect(result.success).toBe(true);
      expect(result.fallbackUsed).toBe("respond_to_user");
      expect(result.warnings).toBeDefined();
      expect(result.warnings!.length).toBeGreaterThan(0);
    });
  });

  describe("执行统计", () => {
    beforeEach(() => {
      ToolExecutor.registerTool(createMockTool("stats_tool", true));
    });

    it("应该记录执行统计", async () => {
      await ToolExecutor.execute("stats_tool", { input: "1" }, { skipMonitoring: true });
      await ToolExecutor.execute("stats_tool", { input: "2" }, { skipMonitoring: true });
      await ToolExecutor.execute("stats_tool", { input: "3" }, { skipMonitoring: true });

      const stats = ToolExecutor.getExecutionStats();

      expect(stats["stats_tool"]).toBeDefined();
      expect(stats["stats_tool"].calls).toBe(3);
      expect(stats["stats_tool"].successes).toBe(3);
      expect(stats["stats_tool"].successRate).toBe(1);
    });
  });
});
