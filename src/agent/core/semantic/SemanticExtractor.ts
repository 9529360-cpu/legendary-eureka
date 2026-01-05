/**
 * SemanticExtractor - 语义抽取器
 * 
 * 单一职责：从用户输入中抽取 Intent / Entities / Constraints
 * 行数上限：400 行
 * 
 * 遵循协议：
 * A. 语义抽取（必须显式写出来）
 *    - Intent: 这句话想解决什么
 *    - Entities: 关键实体
 *    - Constraints: 用户明确或隐含约束
 */

import {
  SemanticExtraction,
  IntentType,
  ExtractedEntities,
  ExtractedConstraints,
} from '../types';
import { mapToSemanticAtoms, compressToSuperIntent } from '../../utils/semanticMapper';

// ========== 意图识别规则 ==========

/**
 * 意图识别规则
 */
interface IntentRule {
  intent: IntentType;
  patterns: RegExp[];
  priority: number;
}

/**
 * 意图识别规则表
 */
const INTENT_RULES: IntentRule[] = [
  // 跨表汇总
  {
    intent: 'cross_sheet_summary',
    patterns: [/跨表/, /跨文件/, /多表汇总/, /importrange/i, /汇总.*表/, /合并.*表/],
    priority: 10,
  },
  // 自动填充
  {
    intent: 'auto_fill',
    patterns: [/自动填充/, /不.*拖/, /自动扩展/, /整列/, /arrayformula/i, /新行.*自动/],
    priority: 10,
  },
  // 结果为0排查
  {
    intent: 'diagnose_zero',
    patterns: [/为什么.*0/, /结果.*是0/, /算出来.*0/, /sum.*0/i, /count.*0/i],
    priority: 15,
  },
  // 循环引用修复
  {
    intent: 'fix_circular_ref',
    patterns: [/循环引用/, /circular/i, /#ref!/i, /自引用/, /环路/],
    priority: 15,
  },
  // 下拉联动
  {
    intent: 'dropdown_cascade',
    patterns: [/下拉.*联动/, /下拉.*选择/, /选择.*联动/, /级联/, /数据验证.*联动/],
    priority: 10,
  },
  // 数据清洗
  {
    intent: 'data_cleanup',
    patterns: [/清洗/, /去重/, /清理/, /duplicate/i, /空白.*删除/, /格式.*统一/],
    priority: 8,
  },
  // 公式创建
  {
    intent: 'formula_creation',
    patterns: [/写.*公式/, /公式.*怎么写/, /vlookup/i, /sumif/i, /countif/i, /index.*match/i],
    priority: 5,
  },
  // 格式化范围
  {
    intent: 'format_range',
    patterns: [/格式化/, /加粗/, /颜色/, /条件格式/, /高亮/],
    priority: 5,
  },
  // 图表创建
  {
    intent: 'chart_creation',
    patterns: [/图表/, /柱状图/, /折线图/, /饼图/, /chart/i],
    priority: 5,
  },
  // 数据分析
  {
    intent: 'data_analysis',
    patterns: [/分析/, /统计/, /趋势/, /透视表/, /pivot/i],
    priority: 8,
  },
  // 结构重构
  {
    intent: 'structure_refactor',
    patterns: [/重构/, /拆分/, /设计.*问题/, /表.*结构/, /长期.*稳定/],
    priority: 12,
  },
];

// ========== 实体抽取规则 ==========

/**
 * 列名模式
 */
const COLUMN_PATTERNS = [
  /([A-Z]+)列/g,
  /列\s*([A-Z]+)/g,
  /["「]([^"」]+)["」]\s*列/g,
  /(\w+)\s*(?:字段|栏)/g,
];

/**
 * 范围模式
 */
const RANGE_PATTERNS = [
  /([A-Z]+\d+:[A-Z]+\d+)/gi,
  /([A-Z]+:\s*[A-Z]+)/gi,
  /(\d+:\d+)/g,
];

/**
 * 工作表名模式
 */
const SHEET_PATTERNS = [
  /['「]([^'」]+)['」]\s*(?:表|sheet)/gi,
  /(?:表|sheet)\s*['「]([^'」]+)['」]/gi,
  /(\w+)\s*(?:工作表|worksheet)/gi,
];

// ========== 约束抽取规则 ==========

/**
 * 约束关键词
 */
const CONSTRAINT_KEYWORDS: Record<string, string[]> = {
  noDragging: ['不拖', '不用拖', '不想拖', '不复制', '不要拖'],
  noManualCopy: ['不复制', '不手动', '自动'],
  autoExpand: ['自动扩展', '新行自动', '整列', 'arrayformula'],
  crossFile: ['跨文件', '跨表格', '另一个文件', '其他文件', 'importrange'],
  multiUser: ['多人', '协作', '同事', '共享', '别人'],
  reusable: ['复用', '通用', '模板', '重复使用'],
  longTermStable: ['长期', '稳定', '不炸', '以后', '维护'],
  urgent: ['紧急', '马上', '立刻', '急', '赶紧', '快'],
  preserveFormat: ['保持格式', '原有格式', '保留格式', '格式不变'],
  readOnly: ['只读', '只查看', '只看', '不修改', '不要改'],
  noCode: ['不用代码', '不写代码', '简单点', '直接'],
};

// ========== SemanticExtractor 类 ==========

/**
 * 语义抽取器
 */
export class SemanticExtractor {
  /**
   * 从用户输入中抽取语义信息
   */
  extract(userInput: string, _context?: Record<string, unknown>): SemanticExtraction {
    const normalizedInput = userInput.toLowerCase();
    
    // 1. 抽取意图
    const intent = this.extractIntent(normalizedInput);
    
    // 2. 抽取实体
    const entities = this.extractEntities(userInput);
    
    // 3. 抽取约束
    const constraints = this.extractConstraints(normalizedInput);
    
    // 4. 使用语义原子增强置信度
    const semanticAtoms = mapToSemanticAtoms(userInput);
    const compressedIntent = compressToSuperIntent(semanticAtoms);
    
    // 5. 计算置信度
    const confidence = this.calculateConfidence(intent, entities, constraints, semanticAtoms);
    
    return {
      intent,
      entities,
      constraints,
      confidence,
      rawInput: userInput,
    };
  }

  /**
   * 抽取意图
   */
  private extractIntent(input: string): IntentType {
    let bestMatch: IntentType = 'unknown';
    let highestPriority = -1;
    
    for (const rule of INTENT_RULES) {
      for (const pattern of rule.patterns) {
        if (pattern.test(input)) {
          if (rule.priority > highestPriority) {
            highestPriority = rule.priority;
            bestMatch = rule.intent;
          }
          break;
        }
      }
    }
    
    return bestMatch;
  }

  /**
   * 抽取实体
   */
  private extractEntities(input: string): ExtractedEntities {
    const entities: ExtractedEntities = {};
    
    // 抽取列名
    const columns: string[] = [];
    for (const pattern of COLUMN_PATTERNS) {
      let match;
      while ((match = pattern.exec(input)) !== null) {
        columns.push(match[1]);
      }
    }
    if (columns.length > 0) {
      entities.columns = [...new Set(columns)];
    }
    
    // 抽取范围
    const ranges: string[] = [];
    for (const pattern of RANGE_PATTERNS) {
      let match;
      while ((match = pattern.exec(input)) !== null) {
        ranges.push(match[1]);
      }
    }
    if (ranges.length > 0) {
      entities.ranges = [...new Set(ranges)];
    }
    
    // 抽取工作表名
    const sheets: string[] = [];
    for (const pattern of SHEET_PATTERNS) {
      let match;
      while ((match = pattern.exec(input)) !== null) {
        sheets.push(match[1]);
      }
    }
    if (sheets.length > 0) {
      entities.sheets = [...new Set(sheets)];
    }
    
    // 抽取指标（简化版）
    const metricPatterns = [/计算\s*(\w+)/, /(\w+)\s*率/, /(\w+)\s*值/];
    const metrics: string[] = [];
    for (const pattern of metricPatterns) {
      const match = input.match(pattern);
      if (match) {
        metrics.push(match[1]);
      }
    }
    if (metrics.length > 0) {
      entities.metrics = [...new Set(metrics)];
    }
    
    return entities;
  }

  /**
   * 抽取约束
   */
  private extractConstraints(input: string): ExtractedConstraints {
    const constraints: ExtractedConstraints = {
      noDragging: false,
      noManualCopy: false,
      autoExpand: false,
      crossFile: false,
      multiUser: false,
      reusable: false,
      longTermStable: false,
      urgent: false,
      preserveFormat: false,
      readOnly: false,
      noCode: false,
    };
    
    for (const [key, keywords] of Object.entries(CONSTRAINT_KEYWORDS)) {
      for (const keyword of keywords) {
        if (input.includes(keyword.toLowerCase())) {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          (constraints as any)[key] = true;
          break;
        }
      }
    }
    
    // 隐含约束推断
    // 如果提到"每天加数据"，隐含 autoExpand
    if (/每天.*加|每天.*新增|日常.*录入/.test(input)) {
      constraints.autoExpand = true;
    }
    
    // 如果提到"部门"、"团队"，隐含 multiUser
    if (/部门|团队|组/.test(input)) {
      constraints.multiUser = true;
    }
    
    return constraints;
  }

  /**
   * 计算置信度
   */
  private calculateConfidence(
    intent: IntentType,
    entities: ExtractedEntities,
    constraints: ExtractedConstraints,
    semanticAtoms: string[]
  ): number {
    let confidence = 0.5; // 基础置信度
    
    // 意图明确加分
    if (intent !== 'unknown') {
      confidence += 0.2;
    }
    
    // 有实体加分
    const entityCount = Object.values(entities).filter(Boolean).length;
    confidence += Math.min(entityCount * 0.05, 0.15);
    
    // 有约束加分
    const constraintCount = Object.values(constraints).filter(Boolean).length;
    confidence += Math.min(constraintCount * 0.03, 0.1);
    
    // 语义原子匹配加分
    if (semanticAtoms.length > 0) {
      confidence += Math.min(semanticAtoms.length * 0.02, 0.1);
    }
    
    return Math.min(confidence, 1.0);
  }

  /**
   * 格式化输出（用于日志和调试）
   */
  formatExtraction(extraction: SemanticExtraction): string {
    return [
      '【语义抽取结果】',
      `Intent: ${extraction.intent}`,
      `Entities: ${JSON.stringify(extraction.entities)}`,
      `Constraints: ${JSON.stringify(extraction.constraints)}`,
      `Confidence: ${(extraction.confidence * 100).toFixed(1)}%`,
    ].join('\n');
  }
}

// ========== 单例导出 ==========

export const semanticExtractor = new SemanticExtractor();

export default SemanticExtractor;
