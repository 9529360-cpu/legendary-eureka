/**
 * 提交包解析器 - SubmissionParser
 *
 * 职责：
 * 1. 从模型输出解析提交包
 * 2. 强制模型遵循输出协议
 * 3. 解析失败时触发拦截重试
 *
 * 模型输出必须包含以下块：
 * [STATE] current_state=... next_state=...
 * [ARTIFACTS] ...
 * [ACCEPTANCE_TESTS] ...
 * [FALLBACK] ...
 * [DEPLOY_NOTES] ...
 * [NEXT_ACTION] ...
 */

import {
  Submission,
  AgentState,
  Artifact,
  ArtifactType,
  Platform,
  AcceptanceTest,
  FallbackPlan,
  DeployNotes,
} from "./types";

// ========== 解析结果 ==========

/**
 * 解析结果
 */
export interface ParseResult {
  success: boolean;
  submission?: Submission;
  missingBlocks: string[];
  errors: string[];
}

// ========== SubmissionParser 类 ==========

/**
 * 提交包解析器
 */
export class SubmissionParser {
  /**
   * 解析模型输出（核心方法）
   */
  parse(modelOutput: string): ParseResult {
    const missingBlocks: string[] = [];
    const _errors: string[] = [];

    // 1. 解析状态
    const stateResult = this.parseStateBlock(modelOutput);
    if (!stateResult.found) {
      missingBlocks.push("[STATE]");
    }

    // 2. 解析产物
    const artifactsResult = this.parseArtifactsBlock(modelOutput);
    if (!artifactsResult.found) {
      missingBlocks.push("[ARTIFACTS]");
    }

    // 3. 解析验收测试
    const testsResult = this.parseAcceptanceTestsBlock(modelOutput);
    if (!testsResult.found) {
      missingBlocks.push("[ACCEPTANCE_TESTS]");
    }

    // 4. 解析回退方案
    const fallbackResult = this.parseFallbackBlock(modelOutput);
    if (!fallbackResult.found) {
      missingBlocks.push("[FALLBACK]");
    }

    // 5. 解析部署说明
    const deployResult = this.parseDeployNotesBlock(modelOutput);
    if (!deployResult.found) {
      missingBlocks.push("[DEPLOY_NOTES]");
    }

    // 6. 解析下一步行动
    const nextActionResult = this.parseNextActionBlock(modelOutput);

    // 检查是否有严重缺失
    if (missingBlocks.length > 0) {
      return {
        success: false,
        missingBlocks,
        errors: [`缺少必需块: ${missingBlocks.join(", ")}`],
      };
    }

    const submission: Submission = {
      proposedState: stateResult.state || AgentState.EXECUTED,
      artifacts: artifactsResult.artifacts,
      acceptanceTests: testsResult.tests,
      fallback: fallbackResult.plans,
      deployNotes: deployResult.notes,
      nextAction: nextActionResult.action,
      rawOutput: modelOutput,
    };

    return {
      success: true,
      submission,
      missingBlocks: [],
      errors: [],
    };
  }

  /**
   * 解析 [STATE] 块
   */
  private parseStateBlock(text: string): { found: boolean; state?: AgentState } {
    const stateBlockPattern = /\[STATE\]([\s\S]*?)(?=\[|$)/i;
    const match = text.match(stateBlockPattern);

    if (!match) {
      return { found: false };
    }

    const content = match[1];

    // 提取 current_state 或 next_state
    const statePattern = /(?:current_state|next_state)\s*[=:]\s*(\w+)/i;
    const stateMatch = content.match(statePattern);

    if (stateMatch) {
      const stateValue = stateMatch[1].toUpperCase();
      const validStates = Object.values(AgentState);
      if (validStates.includes(stateValue as AgentState)) {
        return { found: true, state: stateValue as AgentState };
      }
    }

    return { found: true, state: AgentState.EXECUTED };
  }

  /**
   * 解析 [ARTIFACTS] 块
   */
  private parseArtifactsBlock(text: string): { found: boolean; artifacts: Artifact[] } {
    const blockPattern = /\[ARTIFACTS?\]([\s\S]*?)(?=\[|$)/i;
    const match = text.match(blockPattern);

    if (!match) {
      return { found: false, artifacts: [] };
    }

    const content = match[1];
    const artifacts: Artifact[] = [];

    // 解析每个产物条目
    // 格式：- type=FORMULA platform=google_sheets target_sheet=... target_range=... content=...
    const itemPattern =
      /[-•]\s*(?:type\s*[=:]\s*(\w+))?\s*(?:platform\s*[=:]\s*(\w+))?\s*(?:target_sheet\s*[=:]\s*([^\s,]+))?\s*(?:target_range\s*[=:]\s*([^\s,]+))?\s*(?:target_column\s*[=:]\s*([^\s,]+))?\s*(?:content\s*[=:]\s*(.+?))?(?=[-•]|\[|$)/gi;

    let itemMatch;
    let index = 0;
    while ((itemMatch = itemPattern.exec(content)) !== null) {
      artifacts.push({
        id: `artifact_${index++}`,
        type: this.parseArtifactType(itemMatch[1]),
        platform: this.parsePlatform(itemMatch[2]),
        target: {
          sheet: itemMatch[3]?.trim(),
          range: itemMatch[4]?.trim(),
          column: itemMatch[5]?.trim(),
        },
        content: itemMatch[6]?.trim() || "",
        version: "1.0",
        createdAt: Date.now(),
      });
    }

    // 如果正则没匹配到，尝试更宽松的解析
    if (artifacts.length === 0) {
      // 查找公式内容
      const formulaPattern = /=[\s\S]*?(?=\n\n|\[|$)/g;
      const formulaMatches = content.match(formulaPattern);

      if (formulaMatches) {
        formulaMatches.forEach((formula, i) => {
          artifacts.push({
            id: `artifact_${i}`,
            type: ArtifactType.FORMULA,
            platform: Platform.EXCEL,
            target: {},
            content: formula.trim(),
            version: "1.0",
            createdAt: Date.now(),
          });
        });
      }
    }

    return { found: true, artifacts };
  }

  /**
   * 解析 [ACCEPTANCE_TESTS] 块
   */
  private parseAcceptanceTestsBlock(text: string): { found: boolean; tests: AcceptanceTest[] } {
    const blockPattern = /\[ACCEPTANCE_TESTS?\]([\s\S]*?)(?=\[(?!ACCEPTANCE)|$)/i;
    const match = text.match(blockPattern);

    if (!match) {
      return { found: false, tests: [] };
    }

    const content = match[1];
    const tests: AcceptanceTest[] = [];

    // 按行分割，查找编号列表项
    const lines = content.split("\n");
    let index = 0;

    for (const line of lines) {
      const trimmed = line.trim();
      // 匹配 1) ... 或 1. ... 或 - ... 或 • ...
      const itemMatch = trimmed.match(/^(?:\d+[).]\s*|[-•]\s*)(.+)/);
      if (itemMatch) {
        const description = itemMatch[1].trim();
        if (description) {
          tests.push({
            id: `test_${index++}`,
            description,
            expectedResult: "通过",
          });
        }
      }
    }

    return { found: true, tests };
  }

  /**
   * 解析 [FALLBACK] 块
   */
  private parseFallbackBlock(text: string): { found: boolean; plans: FallbackPlan[] } {
    const blockPattern = /\[FALLBACK\]([\s\S]*?)(?=\[(?!FALLBACK)|$)/i;
    const match = text.match(blockPattern);

    if (!match) {
      return { found: false, plans: [] };
    }

    const content = match[1];
    const plans: FallbackPlan[] = [];

    // 按行分割，查找 fallback 条目
    const lines = content.split("\n");

    for (const line of lines) {
      const trimmed = line.trim();
      if (!trimmed || trimmed.startsWith("[")) continue;

      // 匹配 - if ... then ... 或 - 条件 then 动作
      const ifThenMatch = trimmed.match(/^[-•]\s*(?:if\s+)?(.+?)\s+then\s+(.+)/i);
      if (ifThenMatch) {
        plans.push({
          condition: ifThenMatch[1].trim(),
          action: ifThenMatch[2].trim(),
        });
        continue;
      }

      // 匹配 - 条件: 动作 或 - 条件 → 动作
      const colonMatch = trimmed.match(/^[-•]\s*(.+?)(?::\s*|→\s*)(.+)/);
      if (colonMatch) {
        plans.push({
          condition: colonMatch[1].trim(),
          action: colonMatch[2].trim(),
        });
      }
    }

    return { found: true, plans };
  }

  /**
   * 解析 [DEPLOY_NOTES] 块
   */
  private parseDeployNotesBlock(text: string): { found: boolean; notes?: DeployNotes } {
    const blockPattern = /\[DEPLOY_NOTES?\]([\s\S]*?)(?=\[|$)/i;
    const match = text.match(blockPattern);

    if (!match) {
      return { found: false };
    }

    const content = match[1];
    const notes: DeployNotes = {};

    // 解析各项
    const protectedMatch = content.match(/protect(?:ed)?[_\s]*ranges?\s*[:=]\s*(.+)/i);
    if (protectedMatch) {
      notes.protectedRanges = protectedMatch[1].split(/[,;]/).map((s) => s.trim());
    }

    const namingMatch = content.match(/naming[_\s]*conventions?\s*[:=]\s*(.+)/i);
    if (namingMatch) {
      notes.namingConventions = namingMatch[1].split(/[,;]/).map((s) => s.trim());
    }

    const permMatch = content.match(/permissions?\s*[:=]\s*(.+)/i);
    if (permMatch) {
      notes.permissions = permMatch[1].split(/[,;]/).map((s) => s.trim());
    }

    return { found: true, notes };
  }

  /**
   * 解析 [NEXT_ACTION] 块
   */
  private parseNextActionBlock(text: string): {
    found: boolean;
    action?: { systemWillValidate: string; userNeedsToProvide?: string; ifFailAgentWill?: string };
  } {
    const blockPattern = /\[NEXT_ACTION\]([\s\S]*?)(?=\[|$)/i;
    const match = text.match(blockPattern);

    if (!match) {
      return { found: false };
    }

    const content = match[1];

    const systemMatch = content.match(/system[_\s]*will[_\s]*validate\s*[:=]\s*(.+)/i);
    const userMatch = content.match(/user[_\s]*needs[_\s]*to[_\s]*provide\s*[:=]\s*(.+)/i);
    const failMatch = content.match(/if[_\s]*fail[_\s]*(?:agent[_\s]*will)?\s*[:=]\s*(.+)/i);

    return {
      found: true,
      action: {
        systemWillValidate: systemMatch?.[1]?.trim() || "验证产物",
        userNeedsToProvide: userMatch?.[1]?.trim(),
        ifFailAgentWill: failMatch?.[1]?.trim(),
      },
    };
  }

  /**
   * 解析产物类型
   */
  private parseArtifactType(value?: string): ArtifactType {
    if (!value) return ArtifactType.FORMULA;
    const upper = value.toUpperCase();
    if (upper.includes("STEP")) return ArtifactType.STEPS;
    if (upper.includes("TEMPLATE")) return ArtifactType.TEMPLATE;
    if (upper.includes("SCHEMA")) return ArtifactType.SCHEMA_PLAN;
    return ArtifactType.FORMULA;
  }

  /**
   * 解析平台
   */
  private parsePlatform(value?: string): Platform {
    if (!value) return Platform.EXCEL;
    const lower = value.toLowerCase();
    if (lower.includes("google") || lower.includes("sheets")) return Platform.GOOGLE_SHEETS;
    return Platform.EXCEL;
  }

  /**
   * 生成缺失块的提示消息
   */
  generateMissingBlocksMessage(missingBlocks: string[]): string {
    const lines = [
      "输出格式不符合协议。请按以下格式重新提交：",
      "",
      ...missingBlocks.map((block) => this.getBlockTemplate(block)),
    ];
    return lines.join("\n");
  }

  /**
   * 获取块模板
   */
  private getBlockTemplate(block: string): string {
    const templates: Record<string, string> = {
      "[STATE]": `[STATE]
current_state=EXECUTED
next_state=VERIFIED`,
      "[ARTIFACTS]": `[ARTIFACTS]
- type=FORMULA platform=excel target_sheet=Sheet1 target_range=C2 content==公式内容`,
      "[ACCEPTANCE_TESTS]": `[ACCEPTANCE_TESTS]
1) 测试描述1
2) 测试描述2
3) 测试描述3`,
      "[FALLBACK]": `[FALLBACK]
- if 条件1 then 处理方式1`,
      "[DEPLOY_NOTES]": `[DEPLOY_NOTES]
- protect_ranges: 范围
- naming_conventions: 命名规范
- permissions: 权限说明`,
    };
    return templates[block] || block;
  }
}

// ========== 导出单例 ==========

export const submissionParser = new SubmissionParser();
