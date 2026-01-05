/**
 * 反假完成系统单元测试
 */

import {
  AgentState,
  Platform,
  ArtifactType,
  createAgentRun,
  createEmptyChecklist,
  isChecklistComplete,
  Artifact,
  Submission,
} from "../agent/core/gates/types";
import { StateMachine } from "../agent/core/gates/StateMachine";
import { CompletionGate } from "../agent/core/gates/CompletionGate";
import { FormulaValidator } from "../agent/core/gates/FormulaValidator";
import { SubmissionParser } from "../agent/core/gates/SubmissionParser";
import { ValidationEngine } from "../agent/core/gates/ValidationEngine";
import { AntiHallucinationController } from "../agent/core/gates/AntiHallucinationController";

describe("反假完成系统", () => {
  describe("AgentRun 数据结构", () => {
    it("应该正确创建 AgentRun", () => {
      const run = createAgentRun("user1", "task1");

      expect(run.runId).toMatch(/^run_\d+_/);
      expect(run.userId).toBe("user1");
      expect(run.taskId).toBe("task1");
      expect(run.state).toBe(AgentState.INIT);
      expect(run.iteration).toBe(0);
      expect(run.maxIterations).toBe(8);
      expect(run.artifacts).toEqual([]);
      expect(run.history).toEqual([]);
    });

    it("应该正确创建空 Checklist", () => {
      const checklist = createEmptyChecklist();

      expect(checklist.hasExecutableArtifact).toBe(false);
      expect(checklist.hasPlacementInfo).toBe(false);
      expect(checklist.supportsAutoExpand).toBe(false);
      expect(checklist.avoidsSelfReference).toBe(false);
      expect(checklist.has3AcceptanceTests).toBe(false);
      expect(checklist.hasFallbackPlan).toBe(false);
      expect(checklist.hasDeployNotes).toBe(false);
    });

    it("应该正确检查 Checklist 完整性", () => {
      const incomplete = createEmptyChecklist();
      expect(isChecklistComplete(incomplete)).toBe(false);

      const complete = {
        hasExecutableArtifact: true,
        hasPlacementInfo: true,
        supportsAutoExpand: true,
        avoidsSelfReference: true,
        has3AcceptanceTests: true,
        hasFallbackPlan: true,
        hasDeployNotes: true,
      };
      expect(isChecklistComplete(complete)).toBe(true);
    });
  });

  describe("StateMachine 状态机", () => {
    it("应该允许有效的状态转换", () => {
      const run = createAgentRun("user1", "task1");

      const result1 = StateMachine.transition(run, AgentState.ANALYZED);
      expect(result1.success).toBe(true);
      expect(run.state).toBe(AgentState.ANALYZED);

      const result2 = StateMachine.transition(run, AgentState.DESIGNED);
      expect(result2.success).toBe(true);
      expect(run.state).toBe(AgentState.DESIGNED);
    });

    it("应该阻止非法的状态跳转（INIT → DEPLOYED）", () => {
      const run = createAgentRun("user1", "task1");

      const result = StateMachine.transition(run, AgentState.DEPLOYED);
      expect(result.success).toBe(false);
      expect(run.state).toBe(AgentState.INIT); // 状态不变
    });

    it("应该正确判断下一个状态（验证失败后）", () => {
      const run = createAgentRun("user1", "task1");
      const checklist = createEmptyChecklist();

      // 没有产物 → 回到 DESIGNED
      const next1 = StateMachine.nextStateAfterFail(run, checklist);
      expect(next1).toBe(AgentState.DESIGNED);

      // 有产物但验证失败 → 回到 EXECUTED
      checklist.hasExecutableArtifact = true;
      const next2 = StateMachine.nextStateAfterFail(run, checklist);
      expect(next2).toBe(AgentState.EXECUTED);
    });
  });

  describe("FormulaValidator 公式验证器", () => {
    const validator = new FormulaValidator();

    it("应该检测自引用公式", () => {
      const artifact: Artifact = {
        id: "1",
        type: ArtifactType.FORMULA,
        platform: Platform.EXCEL,
        target: { column: "C" },
        content: "=A1+B1+C1", // C 列引用了自己
        version: "1.0",
        createdAt: Date.now(),
      };

      const results = validator.validate(artifact);
      const selfRefCheck = results.find((r) => r.ruleId === "R2_SELF_REFERENCE");

      expect(selfRefCheck).toBeDefined();
      expect(selfRefCheck!.status).toBe("FAIL");
    });

    it("应该通过无自引用的公式", () => {
      const artifact: Artifact = {
        id: "1",
        type: ArtifactType.FORMULA,
        platform: Platform.EXCEL,
        target: { column: "C" },
        content: "=A1+B1", // C 列没有引用自己
        version: "1.0",
        createdAt: Date.now(),
      };

      const results = validator.validate(artifact);
      const selfRefCheck = results.find((r) => r.ruleId === "R2_SELF_REFERENCE");

      expect(selfRefCheck).toBeDefined();
      expect(selfRefCheck!.status).toBe("PASS");
    });

    it("应该识别 Google Sheets ARRAYFORMULA", () => {
      const artifact: Artifact = {
        id: "1",
        type: ArtifactType.FORMULA,
        platform: Platform.GOOGLE_SHEETS,
        target: { column: "C" },
        content: '=ARRAYFORMULA(IF(A2:A<>"",A2:A+B2:B,""))',
        version: "1.0",
        createdAt: Date.now(),
      };

      const results = validator.validate(artifact);
      const autoExpandCheck = results.find((r) => r.ruleId === "R3_AUTO_EXPAND");

      expect(autoExpandCheck).toBeDefined();
      expect(autoExpandCheck!.status).toBe("PASS");
    });

    it("应该警告硬编码行数范围", () => {
      const artifact: Artifact = {
        id: "1",
        type: ArtifactType.FORMULA,
        platform: Platform.GOOGLE_SHEETS,
        target: { column: "C" },
        content: "=SUM(A2:A100)", // 硬编码 100 行
        version: "1.0",
        createdAt: Date.now(),
      };

      const results = validator.validate(artifact);
      const openRangeCheck = results.find((r) => r.ruleId === "GS4_OPEN_RANGE");

      expect(openRangeCheck).toBeDefined();
      expect(openRangeCheck!.status).toBe("WARN");
    });
  });

  describe("SubmissionParser 提交包解析器", () => {
    const parser = new SubmissionParser();

    it("应该解析完整的提交包", () => {
      const modelOutput = `
[STATE]
current_state=EXECUTED
next_state=VERIFIED

[ARTIFACTS]
- type=FORMULA platform=excel target_sheet=Sheet1 target_range=C2 content==A2+B2

[ACCEPTANCE_TESTS]
1) 新增一行数据，结果自动更新
2) 中间插入一行，结果保持正确
3) 某列为空，不报错

[FALLBACK]
- if 数据是文本 then 使用 VALUE 函数转换

[DEPLOY_NOTES]
- protect_ranges: A1:B100
- naming_conventions: 列名使用中文

[NEXT_ACTION]
- system_will_validate: 验证公式正确性
- user_needs_to_provide: 无
- if_fail_agent_will: 重新设计公式
`;

      const result = parser.parse(modelOutput);

      expect(result.success).toBe(true);
      expect(result.missingBlocks).toEqual([]);
      expect(result.submission).toBeDefined();
      expect(result.submission!.acceptanceTests.length).toBeGreaterThanOrEqual(3);
      expect(result.submission!.fallback.length).toBeGreaterThan(0);
    });

    it("应该检测缺少的块", () => {
      const modelOutput = `
[STATE]
current_state=EXECUTED

[ARTIFACTS]
- 公式内容
`;

      const result = parser.parse(modelOutput);

      expect(result.success).toBe(false);
      expect(result.missingBlocks).toContain("[ACCEPTANCE_TESTS]");
      expect(result.missingBlocks).toContain("[FALLBACK]");
      expect(result.missingBlocks).toContain("[DEPLOY_NOTES]");
    });
  });

  describe("CompletionGate 完成门槛", () => {
    it("应该拒绝不完整的提交", () => {
      const run = createAgentRun("user1", "task1");
      const submission: Submission = {
        proposedState: AgentState.EXECUTED,
        artifacts: [],
        acceptanceTests: [],
        fallback: [],
        rawOutput: "",
      };

      const result = CompletionGate.check(run, submission);

      expect(result.passed).toBe(false);
      expect(result.failReasons.length).toBeGreaterThan(0);
    });

    it("应该通过完整的提交", () => {
      const run = createAgentRun("user1", "task1");
      const submission: Submission = {
        proposedState: AgentState.VERIFIED,
        artifacts: [
          {
            id: "1",
            type: ArtifactType.FORMULA,
            platform: Platform.EXCEL,
            target: { sheet: "Sheet1", range: "C2" },
            content: "=A2+B2",
            version: "1.0",
            createdAt: Date.now(),
          },
        ],
        acceptanceTests: [
          { id: "1", description: "测试1", expectedResult: "通过" },
          { id: "2", description: "测试2", expectedResult: "通过" },
          { id: "3", description: "测试3", expectedResult: "通过" },
        ],
        fallback: [{ condition: "数据为空", action: "返回空字符串" }],
        deployNotes: { protectedRanges: ["A1:B100"] },
        rawOutput: "",
      };

      const result = CompletionGate.check(run, submission);

      expect(result.checklist.hasExecutableArtifact).toBe(true);
      expect(result.checklist.hasPlacementInfo).toBe(true);
      expect(result.checklist.has3AcceptanceTests).toBe(true);
      expect(result.checklist.hasFallbackPlan).toBe(true);
      expect(result.checklist.hasDeployNotes).toBe(true);
    });
  });

  describe("AntiHallucinationController 反假完成控制器", () => {
    const controller = new AntiHallucinationController();

    it("应该创建新的运行实例", () => {
      const run = controller.createRun("user1", "task1");

      expect(run.state).toBe(AgentState.INIT);
      expect(run.iteration).toBe(0);
    });

    it("应该拦截不完整的模型输出", () => {
      const run = controller.createRun("user1", "task1");
      controller.handleUserMessage(run, "计算 A+B");

      const result = controller.handleModelOutput(run, "这是一个简单的加法公式 =A1+B1");

      expect(result.allowFinish).toBe(false);
      expect(result.systemMessage).toBeDefined();
    });

    it("应该正确报告运行状态", () => {
      const run = controller.createRun("user1", "task1");
      const summary = controller.getRunSummary(run);

      expect(summary).toContain("运行状态: INIT");
      expect(summary).toContain("迭代次数: 0/8");
    });
  });

  describe("ValidationEngine 验证引擎", () => {
    const engine = new ValidationEngine();

    it("应该生成完整的验证报告", () => {
      const submission: Submission = {
        proposedState: AgentState.VERIFIED,
        artifacts: [
          {
            id: "1",
            type: ArtifactType.FORMULA,
            platform: Platform.EXCEL,
            target: { sheet: "Sheet1", column: "C" },
            content: "=A1+B1",
            version: "1.0",
            createdAt: Date.now(),
          },
        ],
        acceptanceTests: [
          { id: "1", description: "测试1", expectedResult: "通过" },
          { id: "2", description: "测试2", expectedResult: "通过" },
          { id: "3", description: "测试3", expectedResult: "通过" },
        ],
        fallback: [{ condition: "条件", action: "动作" }],
        deployNotes: { protectedRanges: ["A1:B100"] },
        rawOutput: "",
      };

      const report = engine.validate(submission);

      expect(report).toBeDefined();
      expect(report.summary).toBeDefined();
      expect(report.checklist).toBeDefined();
    });
  });
});
