/**
 * DiagnosticEngine - 诊断引擎
 *
 * 单一职责：诊断表格问题的根因，给出 Top3 原因和验证步骤
 * 行数上限：400 行
 *
 * 遵循协议：
 * B. 诊断与路由（必须给出最可能原因排序）
 *    - Top3 可能原因
 *    - 每个原因的"最短验证方法"
 */

import { DiagnosticResult, DiagnosticCause, ValidationStep } from "../types";
import { mapToSemanticAtoms } from "../../utils/semanticMapper";

// ========== 诊断规则库 ==========

/**
 * 诊断规则
 */
interface DiagnosticRule {
  symptom: string;
  patterns: RegExp[];
  causes: DiagnosticCause[];
  validationSteps: ValidationStep[];
  recommendedFix: string;
  riskNotes: string[];
}

/**
 * 诊断规则库
 */
const DIAGNOSTIC_RULES: DiagnosticRule[] = [
  // 结果为 0
  {
    symptom: "result_is_zero",
    patterns: [/结果.*是0/, /算出来.*0/, /sum.*0/i, /为什么.*0/],
    causes: [
      {
        rank: 1,
        cause: "数据被识别为文本而非数字",
        probability: 0.45,
        shortestValidation: "选中数据列，看左上角是否有绿色三角标记；或用 =ISNUMBER(A1) 检查",
      },
      {
        rank: 2,
        cause: "引用了错误的列或范围",
        probability: 0.3,
        shortestValidation: "双击公式单元格，查看高亮引用区域是否正确",
      },
      {
        rank: 3,
        cause: "有隐藏的空格或不可见字符",
        probability: 0.15,
        shortestValidation: "用 =LEN(A1) 检查字符长度是否异常",
      },
    ],
    validationSteps: [
      { order: 1, description: "检查数据类型", formula: "=ISNUMBER(A1)", expectedResult: "TRUE" },
      { order: 2, description: "检查字符长度", formula: "=LEN(A1)", expectedResult: "与预期相符" },
      { order: 3, description: "强制转数字", formula: "=VALUE(A1)", expectedResult: "数字值" },
    ],
    recommendedFix: '使用 VALUE() 函数强制转换，或选中数据列后用"分列"功能转换格式',
    riskNotes: ["转换后原有格式可能丢失", "批量转换前建议先备份"],
  },

  // 循环引用
  {
    symptom: "circular_reference",
    patterns: [/循环引用/, /circular/i, /#ref!/i, /自引用/],
    causes: [
      {
        rank: 1,
        cause: "公式所在单元格被包含在自身引用范围内",
        probability: 0.5,
        shortestValidation: "检查公式引用范围是否包含公式所在单元格",
      },
      {
        rank: 2,
        cause: "汇总行/列与明细数据混在同一区域",
        probability: 0.3,
        shortestValidation: "确认汇总公式是否放在数据区域之外",
      },
      {
        rank: 3,
        cause: "多个公式互相引用形成环路",
        probability: 0.15,
        shortestValidation: '用"公式" → "追踪引用单元格"功能检查依赖链',
      },
    ],
    validationSteps: [
      { order: 1, description: "追踪引用", formula: '使用"公式"菜单的追踪功能' },
      { order: 2, description: "检查范围边界", formula: "确认汇总行不在数据范围内" },
    ],
    recommendedFix: "将汇总公式移到数据区域外，或使用命名范围明确边界",
    riskNotes: ["修改前建议备份", "如果是复杂公式链，考虑重新设计表结构"],
  },

  // IMPORTRANGE 问题
  {
    symptom: "importrange_issue",
    patterns: [/importrange/i, /跨文件.*0/, /允许访问/, /外部表格/],
    causes: [
      {
        rank: 1,
        cause: '未授权访问（首次使用需点击"允许访问"）',
        probability: 0.4,
        shortestValidation: "在目标表格中手动输入 IMPORTRANGE 并点击授权按钮",
      },
      {
        rank: 2,
        cause: "源文件链接已更改或被删除",
        probability: 0.3,
        shortestValidation: "检查 file_id 是否仍然有效",
      },
      {
        rank: 3,
        cause: "范围地址不正确",
        probability: 0.2,
        shortestValidation: "检查 sheet 名和范围拼写是否正确",
      },
    ],
    validationSteps: [
      { order: 1, description: "检查授权状态", formula: "手动输入 IMPORTRANGE 查看是否提示授权" },
      { order: 2, description: "验证文件链接", formula: "在浏览器中打开 file_id 对应的文件" },
      { order: 3, description: "验证范围", formula: "在源文件中检查 sheet 名和范围是否存在" },
    ],
    recommendedFix: "1) 先在单独单元格测试 IMPORTRANGE 并授权；2) 确认后再嵌套到其他公式中",
    riskNotes: ["每个目标文件只需授权一次", "权限变更可能导致公式失效"],
  },

  // ARRAYFORMULA 问题
  {
    symptom: "arrayformula_issue",
    patterns: [/arrayformula/i, /数组.*溢出/, /spill/i, /#spill!/i, /数组.*扩展/],
    causes: [
      {
        rank: 1,
        cause: "数组输出区域有数据阻挡",
        probability: 0.45,
        shortestValidation: "检查公式下方是否有其他数据",
      },
      {
        rank: 2,
        cause: "ARRAYFORMULA 放在了错误的位置",
        probability: 0.3,
        shortestValidation: "ARRAYFORMULA 应放在输出范围的第一个单元格",
      },
      {
        rank: 3,
        cause: "嵌套公式不兼容数组运算",
        probability: 0.15,
        shortestValidation: "检查内部函数是否支持数组",
      },
    ],
    validationSteps: [
      { order: 1, description: "清空下方区域", formula: "删除公式下方可能阻挡的数据" },
      { order: 2, description: "检查公式位置", formula: "确认公式在第一行（输出区域的起点）" },
    ],
    recommendedFix: "确保 ARRAYFORMULA 输出区域完全为空，公式放在区域起点",
    riskNotes: ["数组区域不要手动输入数据", "修改数组范围需要整体调整公式"],
  },

  // 文本数字混合
  {
    symptom: "text_number_mixed",
    patterns: [/文本.*数字/, /当成文本/, /数字.*文本/, /格式.*不对/],
    causes: [
      {
        rank: 1,
        cause: "数据导入时被识别为文本",
        probability: 0.5,
        shortestValidation: "用 =ISNUMBER(A1) 检查，返回 FALSE 说明是文本",
      },
      {
        rank: 2,
        cause: "有隐藏的空格或特殊字符",
        probability: 0.3,
        shortestValidation: "用 =TRIM(A1) 或 =CLEAN(A1) 清理",
      },
      {
        rank: 3,
        cause: "千分位或小数点格式不一致",
        probability: 0.15,
        shortestValidation: "检查是否混用了中文逗号或欧洲格式",
      },
    ],
    validationSteps: [
      { order: 1, description: "检查类型", formula: "=ISNUMBER(A1)" },
      { order: 2, description: "强制转换", formula: "=VALUE(TRIM(A1))" },
    ],
    recommendedFix: "使用 =VALUE(TRIM(CLEAN(A1))) 组合清理并转换",
    riskNotes: ["批量转换前先在单个单元格测试", "原数据格式可能丢失"],
  },
];

// ========== DiagnosticEngine 类 ==========

/**
 * 诊断引擎
 */
export class DiagnosticEngine {
  /**
   * 诊断问题
   */
  diagnose(userInput: string, context?: Record<string, unknown>): DiagnosticResult {
    const normalizedInput = userInput.toLowerCase();

    // 1. 匹配症状
    const matchedRule = this.matchSymptom(normalizedInput);

    if (matchedRule) {
      return {
        possibleCauses: matchedRule.causes,
        validationSteps: matchedRule.validationSteps,
        recommendedFix: matchedRule.recommendedFix,
        riskNotes: matchedRule.riskNotes,
      };
    }

    // 2. 使用语义原子做通用诊断
    const semanticAtoms = mapToSemanticAtoms(userInput);
    return this.genericDiagnosis(semanticAtoms, context);
  }

  /**
   * 匹配症状规则
   */
  private matchSymptom(input: string): DiagnosticRule | null {
    for (const rule of DIAGNOSTIC_RULES) {
      for (const pattern of rule.patterns) {
        if (pattern.test(input)) {
          return rule;
        }
      }
    }
    return null;
  }

  /**
   * 通用诊断（基于语义原子）
   */
  private genericDiagnosis(atoms: string[], _context?: Record<string, unknown>): DiagnosticResult {
    const causes: DiagnosticCause[] = [];
    const validationSteps: ValidationStep[] = [];
    const riskNotes: string[] = [];

    // 根据语义原子生成诊断
    if (atoms.includes("aggregation_result_unexpected_zero")) {
      causes.push({
        rank: 1,
        cause: "聚合函数返回意外的零值",
        probability: 0.5,
        shortestValidation: "检查数据类型和引用范围",
      });
    }

    if (atoms.includes("text_number_coercion_needed")) {
      causes.push({
        rank: causes.length + 1,
        cause: "需要文本到数字的类型转换",
        probability: 0.4,
        shortestValidation: "使用 VALUE() 函数转换",
      });
    }

    if (atoms.includes("self_reference_detected")) {
      causes.push({
        rank: causes.length + 1,
        cause: "检测到自引用或循环依赖",
        probability: 0.6,
        shortestValidation: "检查公式引用链",
      });
    }

    // 如果没有匹配的原子，给出通用建议
    if (causes.length === 0) {
      causes.push({
        rank: 1,
        cause: "需要更多上下文信息",
        probability: 0.3,
        shortestValidation: "请提供具体的公式和数据示例",
      });
    }

    validationSteps.push({
      order: 1,
      description: "检查公式语法和引用范围",
    });

    riskNotes.push("建议在修改前备份原数据");

    return {
      possibleCauses: causes,
      validationSteps,
      recommendedFix: "根据诊断结果选择对应的修复方案",
      riskNotes,
    };
  }

  /**
   * 格式化诊断结果
   */
  formatDiagnosis(result: DiagnosticResult): string {
    const lines: string[] = ["【快速诊断】"];

    lines.push("");
    lines.push("Top3 可能原因：");
    for (const cause of result.possibleCauses.slice(0, 3)) {
      lines.push(`  ${cause.rank}. ${cause.cause} (${(cause.probability * 100).toFixed(0)}%)`);
      lines.push(`     验证方法: ${cause.shortestValidation}`);
    }

    lines.push("");
    lines.push("验证步骤：");
    for (const step of result.validationSteps) {
      lines.push(`  ${step.order}. ${step.description}`);
      if (step.formula) {
        lines.push(`     公式: ${step.formula}`);
      }
    }

    lines.push("");
    lines.push(`推荐修复: ${result.recommendedFix}`);

    if (result.riskNotes.length > 0) {
      lines.push("");
      lines.push("⚠️ 注意事项：");
      for (const note of result.riskNotes) {
        lines.push(`  - ${note}`);
      }
    }

    return lines.join("\n");
  }
}

// ========== 单例导出 ==========

export const diagnosticEngine = new DiagnosticEngine();

export default DiagnosticEngine;
