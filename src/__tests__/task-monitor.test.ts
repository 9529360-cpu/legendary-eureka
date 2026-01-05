/**
 * TaskExecutionMonitor 测试
 *
 * 覆盖：
 * - 任务生命周期追踪
 * - 工具调用监控
 * - 告警系统
 * - 统计分析
 */

import { TaskExecutionMonitor, TaskPhase, AlertLevel } from "../core/TaskExecutionMonitor";

describe("TaskExecutionMonitor", () => {
  beforeEach(() => {
    TaskExecutionMonitor.reset();
  });

  describe("任务生命周期", () => {
    it("应该正确开始任务追踪", () => {
      const record = TaskExecutionMonitor.startTask("task-1", "美化表格");

      expect(record.taskId).toBe("task-1");
      expect(record.request).toBe("美化表格");
      expect(record.status).toBe("running");
      expect(record.phases).toHaveLength(0);
      expect(record.toolCalls).toHaveLength(0);
    });

    it("应该正确追踪阶段", () => {
      TaskExecutionMonitor.startTask("task-2", "创建图表");

      TaskExecutionMonitor.startPhase("task-2", TaskPhase.INTENT_PARSING);
      TaskExecutionMonitor.completePhase("task-2", TaskPhase.INTENT_PARSING);

      TaskExecutionMonitor.startPhase("task-2", TaskPhase.PLANNING);
      TaskExecutionMonitor.completePhase("task-2", TaskPhase.PLANNING);

      const record = TaskExecutionMonitor.getTaskRecord("task-2");

      expect(record).toBeDefined();
      expect(record!.phases).toHaveLength(2);
      expect(record!.phases[0].phase).toBe(TaskPhase.INTENT_PARSING);
      expect(record!.phases[0].status).toBe("completed");
      expect(record!.phases[1].phase).toBe(TaskPhase.PLANNING);
    });

    it("应该正确处理阶段失败", () => {
      TaskExecutionMonitor.startTask("task-3", "失败任务");

      TaskExecutionMonitor.startPhase("task-3", TaskPhase.TOOL_EXECUTION);
      TaskExecutionMonitor.failPhase("task-3", TaskPhase.TOOL_EXECUTION, "工具执行错误");

      const record = TaskExecutionMonitor.getTaskRecord("task-3");

      expect(record!.phases[0].status).toBe("failed");
      expect(record!.phases[0].error).toBe("工具执行错误");
    });

    it("应该正确完成任务", () => {
      TaskExecutionMonitor.startTask("task-4", "完成任务");
      const record = TaskExecutionMonitor.completeTask("task-4", "任务完成");

      expect(record).toBeDefined();
      expect(record!.status).toBe("completed");
      expect(record!.result).toBe("任务完成");
      expect(record!.metrics.totalDuration).toBeGreaterThanOrEqual(0);
    });

    it("应该正确处理任务失败", () => {
      TaskExecutionMonitor.startTask("task-5", "失败任务");
      const record = TaskExecutionMonitor.failTask("task-5", "未知错误");

      expect(record).toBeDefined();
      expect(record!.status).toBe("failed");
      expect(record!.error).toBe("未知错误");
    });
  });

  describe("工具调用追踪", () => {
    beforeEach(() => {
      TaskExecutionMonitor.registerTools([
        "excel_read_range",
        "excel_write_range",
        "excel_format_range",
      ]);
    });

    it("应该追踪工具调用", () => {
      TaskExecutionMonitor.startTask("task-tool", "工具测试");

      TaskExecutionMonitor.startToolCall("task-tool", "excel_read_range", { address: "A1:B10" });
      TaskExecutionMonitor.completeToolCall("task-tool", "excel_read_range", "数据已读取", true);

      const record = TaskExecutionMonitor.getTaskRecord("task-tool");

      expect(record!.toolCalls).toHaveLength(1);
      expect(record!.toolCalls[0].toolName).toBe("excel_read_range");
      expect(record!.toolCalls[0].status).toBe("success");
      expect(record!.metrics.successfulToolCalls).toBe(1);
    });

    it("应该检测未注册的工具调用", () => {
      TaskExecutionMonitor.startTask("task-unreg", "未注册工具");

      const toolCall = TaskExecutionMonitor.startToolCall("task-unreg", "unknown_tool", {
        input: "test",
      });

      expect(toolCall.status).toBe("not_found");
    });

    it("应该记录工具调用失败", () => {
      TaskExecutionMonitor.startTask("task-fail", "失败工具");

      TaskExecutionMonitor.startToolCall("task-fail", "excel_write_range", { address: "A1" });
      TaskExecutionMonitor.failToolCall("task-fail", "excel_write_range", "写入失败");

      const record = TaskExecutionMonitor.getTaskRecord("task-fail");

      expect(record!.toolCalls[0].status).toBe("failed");
      expect(record!.toolCalls[0].error).toBe("写入失败");
      expect(record!.metrics.failedToolCalls).toBe(1);
    });

    it("应该记录兜底操作", () => {
      TaskExecutionMonitor.startTask("task-fallback", "兜底测试");

      TaskExecutionMonitor.recordFallback(
        "task-fallback",
        "excel_format_range",
        "respond_to_user",
        "工具不可用"
      );

      const record = TaskExecutionMonitor.getTaskRecord("task-fallback");

      expect(record!.metrics.fallbackCount).toBe(1);
    });
  });

  describe("告警系统", () => {
    it("应该触发告警", () => {
      const alerts: any[] = [];
      TaskExecutionMonitor.addAlertListener((alert) => alerts.push(alert));

      TaskExecutionMonitor.raiseAlert(AlertLevel.ERROR, "TEST_ERROR", "测试错误消息", {
        detail: "附加信息",
      });

      expect(alerts).toHaveLength(1);
      expect(alerts[0].level).toBe(AlertLevel.ERROR);
      expect(alerts[0].code).toBe("TEST_ERROR");
      expect(alerts[0].acknowledged).toBe(false);
    });

    it("应该获取未确认的告警", () => {
      TaskExecutionMonitor.raiseAlert(AlertLevel.WARNING, "WARN_1", "警告1");
      TaskExecutionMonitor.raiseAlert(AlertLevel.ERROR, "ERR_1", "错误1");

      const unack = TaskExecutionMonitor.getUnacknowledgedAlerts();

      expect(unack).toHaveLength(2);
    });

    it("应该确认告警", () => {
      TaskExecutionMonitor.raiseAlert(AlertLevel.WARNING, "WARN_2", "警告2");

      TaskExecutionMonitor.acknowledgeAlert(0);

      const unack = TaskExecutionMonitor.getUnacknowledgedAlerts();
      expect(unack).toHaveLength(0);
    });
  });

  describe("统计分析", () => {
    beforeEach(() => {
      TaskExecutionMonitor.registerTools(["tool_a", "tool_b"]);

      // 创建一些任务记录
      TaskExecutionMonitor.startTask("stat-1", "任务1");
      TaskExecutionMonitor.startToolCall("stat-1", "tool_a", {});
      TaskExecutionMonitor.completeToolCall("stat-1", "tool_a", "OK", true);
      TaskExecutionMonitor.completeTask("stat-1", "完成");

      TaskExecutionMonitor.startTask("stat-2", "任务2");
      TaskExecutionMonitor.startToolCall("stat-2", "tool_a", {});
      TaskExecutionMonitor.completeToolCall("stat-2", "tool_a", "OK", true);
      TaskExecutionMonitor.startToolCall("stat-2", "tool_b", {});
      TaskExecutionMonitor.failToolCall("stat-2", "tool_b", "失败");
      TaskExecutionMonitor.failTask("stat-2", "任务失败");
    });

    it("应该正确统计任务数据", () => {
      const stats = TaskExecutionMonitor.getStatistics();

      expect(stats.totalTasks).toBe(2);
      expect(stats.completedTasks).toBe(1);
      expect(stats.failedTasks).toBe(1);
    });

    it("应该正确统计工具使用", () => {
      const stats = TaskExecutionMonitor.getStatistics();

      expect(stats.toolUsageStats["tool_a"]).toBeDefined();
      expect(stats.toolUsageStats["tool_a"].calls).toBe(2);
      expect(stats.toolUsageStats["tool_a"].failures).toBe(0);

      expect(stats.toolUsageStats["tool_b"]).toBeDefined();
      expect(stats.toolUsageStats["tool_b"].calls).toBe(1);
      expect(stats.toolUsageStats["tool_b"].failures).toBe(1);
    });
  });

  describe("工具一致性检查", () => {
    it("应该检测未注册但被调用的工具", () => {
      TaskExecutionMonitor.registerTools(["registered_tool"]);

      TaskExecutionMonitor.startTask("cons-1", "一致性测试");
      TaskExecutionMonitor.startToolCall("cons-1", "unregistered_tool", {});

      const consistency = TaskExecutionMonitor.checkToolConsistency();

      expect(consistency.usedButNotRegistered).toContain("unregistered_tool");
    });

    it("应该检测已注册但从未使用的工具", () => {
      TaskExecutionMonitor.registerTools(["never_used_tool", "used_tool"]);

      TaskExecutionMonitor.startTask("cons-2", "使用测试");
      TaskExecutionMonitor.startToolCall("cons-2", "used_tool", {});
      TaskExecutionMonitor.completeToolCall("cons-2", "used_tool", "OK");

      const consistency = TaskExecutionMonitor.checkToolConsistency();

      expect(consistency.registeredButNeverUsed).toContain("never_used_tool");
    });
  });
});
