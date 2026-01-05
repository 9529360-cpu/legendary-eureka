/**
 * SolutionBuilder - 解决方案构建器
 * 
 * 单一职责：根据诊断结果生成分层解决方案
 * 行数上限：400 行
 * 
 * 遵循协议：
 * C. 分层解决方案（必须输出三个层次）
 *    - 🚀 最小可行: 立刻能跑的最小改动
 *    - ✅ 推荐方案: 长期稳定、易理解
 *    - 🏗️ 结构优化: 带来整体效率提升的重构
 */

import {
  LayeredSolution,
  SolutionOption,
  DiagnosticResult,
  SemanticExtraction,
  IntentType,
} from '../types';

// ========== 解决方案模板库 ==========

/**
 * 解决方案模板
 */
interface SolutionTemplate {
  intent: IntentType;
  minimal: SolutionOption;
  recommended: SolutionOption;
  structural: SolutionOption;
}

/**
 * 解决方案模板库
 */
const SOLUTION_TEMPLATES: SolutionTemplate[] = [
  // 创建公式
  {
    intent: 'create_formula',
    minimal: {
      tier: 'minimal',
      emoji: '🚀',
      title: '最小可行方案',
      description: '直接在目标单元格输入公式',
      steps: ['定位到目标单元格', '输入公式', '按 Enter 确认'],
      code: '=SUM(A1:A10)',
      pros: ['立即生效', '无需额外设置'],
      cons: ['手动操作', '不适合批量'],
    },
    recommended: {
      tier: 'recommended',
      emoji: '✅',
      title: '推荐方案',
      description: '使用命名范围，提高可读性和维护性',
      steps: [
        '选中数据区域',
        '定义命名范围（如 "销售数据"）',
        '使用 =SUM(销售数据) 代替绝对引用',
      ],
      code: '=SUM(销售数据)',
      pros: ['可读性好', '易于维护', '修改范围时自动更新'],
      cons: ['需要额外定义命名范围'],
    },
    structural: {
      tier: 'structural',
      emoji: '🏗️',
      title: '结构优化方案',
      description: '重构为标准化表格结构',
      steps: [
        '将数据转换为 Excel 表格（Ctrl+T）',
        '使用结构化引用（如 =SUM(Table1[销售额])）',
        '考虑添加数据验证规则',
      ],
      code: '=SUM(Table1[销售额])',
      pros: ['结构清晰', '自动扩展', '支持高级功能'],
      cons: ['需要重构现有数据', '学习成本较高'],
    },
  },

  // 格式化
  {
    intent: 'format',
    minimal: {
      tier: 'minimal',
      emoji: '🚀',
      title: '最小可行方案',
      description: '手动设置单元格格式',
      steps: ['选中目标区域', '右键 → 设置单元格格式', '选择所需格式'],
      pros: ['直观简单', '立即生效'],
      cons: ['手动操作', '不可复用'],
    },
    recommended: {
      tier: 'recommended',
      emoji: '✅',
      title: '推荐方案',
      description: '使用条件格式化规则',
      steps: [
        '选中数据区域',
        '开始 → 条件格式化 → 新建规则',
        '设置条件和格式',
      ],
      pros: ['自动应用', '可视化效果好', '易于管理'],
      cons: ['规则过多可能影响性能'],
    },
    structural: {
      tier: 'structural',
      emoji: '🏗️',
      title: '结构优化方案',
      description: '建立格式模板和样式库',
      steps: [
        '定义统一的单元格样式',
        '创建自定义格式模板',
        '使用格式刷批量应用',
      ],
      pros: ['一致性好', '易于维护', '可跨文件复用'],
      cons: ['需要前期规划', '团队需要统一标准'],
    },
  },

  // 数据清洗
  {
    intent: 'clean_data',
    minimal: {
      tier: 'minimal',
      emoji: '🚀',
      title: '最小可行方案',
      description: '使用 TRIM/CLEAN 函数处理',
      steps: ['在空列输入清洗公式', '下拉填充', '复制粘贴为值'],
      code: '=TRIM(CLEAN(A1))',
      pros: ['快速解决当前问题', '不破坏原数据'],
      cons: ['临时解决方案', '每次都需重复操作'],
    },
    recommended: {
      tier: 'recommended',
      emoji: '✅',
      title: '推荐方案',
      description: '使用数据验证规则防止脏数据',
      steps: [
        '对输入列设置数据验证',
        '使用 ARRAYFORMULA 自动清洗新数据',
        '建立错误数据提示',
      ],
      pros: ['防患于未然', '自动化处理'],
      cons: ['需要设置验证规则'],
    },
    structural: {
      tier: 'structural',
      emoji: '🏗️',
      title: '结构优化方案',
      description: '建立 ETL 流程',
      steps: [
        '设计数据输入表单',
        '创建数据清洗中间层',
        '使用 Power Query 自动化处理',
      ],
      pros: ['完整的数据治理', '可追溯', '企业级方案'],
      cons: ['实施成本较高', '需要技术支持'],
    },
  },

  // 数据分析
  {
    intent: 'analyze',
    minimal: {
      tier: 'minimal',
      emoji: '🚀',
      title: '最小可行方案',
      description: '使用基础统计函数',
      steps: ['用 SUM/AVERAGE/COUNT 等函数', '手动创建汇总表'],
      code: '=AVERAGE(A:A)',
      pros: ['简单直接', '立即获得结果'],
      cons: ['手动维护', '不自动更新'],
    },
    recommended: {
      tier: 'recommended',
      emoji: '✅',
      title: '推荐方案',
      description: '使用数据透视表',
      steps: [
        '选中数据区域',
        '插入 → 数据透视表',
        '配置行/列/值字段',
      ],
      pros: ['交互式分析', '自动刷新', '灵活性高'],
      cons: ['需要学习透视表操作'],
    },
    structural: {
      tier: 'structural',
      emoji: '🏗️',
      title: '结构优化方案',
      description: '建立数据分析仪表板',
      steps: [
        '创建独立的分析工作表',
        '使用切片器控制透视表',
        '添加图表可视化',
      ],
      pros: ['专业仪表板', '一目了然', '可分享'],
      cons: ['需要设计和规划', '维护成本'],
    },
  },
];

// ========== SolutionBuilder 类 ==========

/**
 * 解决方案构建器
 */
export class SolutionBuilder {
  /**
   * 从语义提取结果构建解决方案
   */
  buildFromSemanticExtraction(extraction: SemanticExtraction): LayeredSolution {
    const template = this.findTemplate(extraction.intent);
    
    if (template) {
      return {
        minimal: this.customizeOption(template.minimal, extraction),
        recommended: this.customizeOption(template.recommended, extraction),
        structural: this.customizeOption(template.structural, extraction),
      };
    }
    
    return this.buildGenericSolution(extraction);
  }

  /**
   * 从诊断结果构建解决方案
   */
  buildFromDiagnosis(diagnosis: DiagnosticResult): LayeredSolution {
    const mainCause = diagnosis.possibleCauses[0];
    
    return {
      minimal: {
        tier: 'minimal',
        emoji: '🚀',
        title: '快速修复',
        description: mainCause?.shortestValidation || '验证问题后手动修复',
        steps: diagnosis.validationSteps.map(s => s.description),
        pros: ['立即解决问题'],
        cons: ['可能只是临时方案'],
      },
      recommended: {
        tier: 'recommended',
        emoji: '✅',
        title: '推荐方案',
        description: diagnosis.recommendedFix,
        steps: [
          '按验证步骤确认问题',
          ...diagnosis.validationSteps.map(s => s.description),
          '应用修复方案',
        ],
        pros: ['解决根本问题', '防止复发'],
        cons: ['可能需要更多时间'],
      },
      structural: {
        tier: 'structural',
        emoji: '🏗️',
        title: '结构优化',
        description: '从数据结构层面解决问题',
        steps: [
          '审视当前数据架构',
          '考虑是否需要重构表结构',
          '建立数据验证机制',
        ],
        pros: ['长期收益', '系统性解决'],
        cons: ['需要更多投入', '可能影响现有流程'],
      },
    };
  }

  /**
   * 查找匹配的模板
   */
  private findTemplate(intent: IntentType): SolutionTemplate | null {
    return SOLUTION_TEMPLATES.find(t => t.intent === intent) || null;
  }

  /**
   * 定制化选项
   */
  private customizeOption(
    option: SolutionOption,
    extraction: SemanticExtraction
  ): SolutionOption {
    const customized = { ...option };
    
    // 根据约束条件调整
    if (extraction.constraints.urgent) {
      customized.description = `【紧急】${customized.description}`;
    }
    
    if (extraction.constraints.noCode) {
      customized.code = undefined;
    }
    
    if (extraction.constraints.preserveFormat && customized.steps) {
      customized.steps = [
        '备份原有格式',
        ...customized.steps,
        '验证格式保持一致',
      ];
    }
    
    return customized;
  }

  /**
   * 构建通用解决方案
   */
  private buildGenericSolution(extraction: SemanticExtraction): LayeredSolution {
    return {
      minimal: {
        tier: 'minimal',
        emoji: '🚀',
        title: '最小可行方案',
        description: `快速处理${extraction.intent}需求`,
        steps: ['分析当前数据', '执行基础操作', '验证结果'],
        pros: ['快速完成'],
        cons: ['可能不够完善'],
      },
      recommended: {
        tier: 'recommended',
        emoji: '✅',
        title: '推荐方案',
        description: `标准化处理${extraction.intent}需求`,
        steps: ['规划操作步骤', '执行标准流程', '验证并记录'],
        pros: ['稳定可靠'],
        cons: ['需要更多时间'],
      },
      structural: {
        tier: 'structural',
        emoji: '🏗️',
        title: '结构优化方案',
        description: '从根本上改进数据结构',
        steps: ['评估当前架构', '设计优化方案', '逐步实施重构'],
        pros: ['长期收益大'],
        cons: ['需要投入资源'],
      },
    };
  }

  /**
   * 格式化分层解决方案
   */
  formatSolution(solution: LayeredSolution): string {
    const lines: string[] = ['【分层解决方案】'];
    
    for (const tier of ['minimal', 'recommended', 'structural'] as const) {
      const opt = solution[tier];
      if (!opt) continue;
      
      lines.push('');
      lines.push(`${opt.emoji} ${opt.title}`);
      lines.push(`   ${opt.description}`);
      
      if (opt.steps && opt.steps.length > 0) {
        lines.push('   步骤：');
        opt.steps.forEach((step, i) => {
          lines.push(`     ${i + 1}. ${step}`);
        });
      }
      
      if (opt.code) {
        lines.push(`   代码: ${opt.code}`);
      }
      
      if (opt.pros && opt.pros.length > 0) {
        lines.push(`   优点: ${opt.pros.join('、')}`);
      }
      
      if (opt.cons && opt.cons.length > 0) {
        lines.push(`   注意: ${opt.cons.join('、')}`);
      }
    }
    
    return lines.join('\n');
  }
}

// ========== 单例导出 ==========

export const solutionBuilder = new SolutionBuilder();

export default SolutionBuilder;
