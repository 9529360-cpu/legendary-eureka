/**
 * semanticMapper.ts
 * 把用户自然语言映射到预定义的语义原子（semantic atoms）和四个压缩意图
 */

export const SUPER_INTENTS = {
  AUTOMATION: 'automation',
  FAILURE: 'failure',
  STRUCTURE: 'structure',
  MAINTAINABILITY: 'maintainability',
} as const;

export type SuperIntent = typeof SUPER_INTENTS[keyof typeof SUPER_INTENTS];

export const SEMANTIC_ATOMS = [
  'auto_expand_required',
  'no_manual_dragging',
  'future_rows_unknown',
  'aggregation_result_unexpected_zero',
  'aggregation_result_empty',
  'formula_logic_not_understood',
  'error_message_not_actionable',
  'cross_sheet_data_missing',
  'cross_file_permission_issue',
  'self_reference_detected',
  'arrayformula_misplacement',
  'text_number_coercion_needed',
  'schema_role_confusion',
  'summary_and_detail_mixed',
  'dropdown_driven_calculation',
  'selection_change_no_effect',
  'long_term_maintainability_concern',
  'multi_user_safety_required',
  'structure_should_be_refactored',
  'fear_of_future_breakage',
] as const;

export type SemanticAtom = typeof SEMANTIC_ATOMS[number];

// 简单的关键词到原子映射表（可扩展）
const KEYWORD_MAP: Array<{ keywords: RegExp[]; atoms: SemanticAtom[] }> = [
  { keywords: [/拖公式/, /不想.*拖/, /不想.*复制/, /不用拖/], atoms: ['no_manual_dragging', 'auto_expand_required'] },
  { keywords: [/新行/, /新增.*行/, /每天.*加数据/, /每天.*加/], atoms: ['future_rows_unknown', 'auto_expand_required'] },
  { keywords: [/结果.*是0/, /算出来是0/, /SUM.*0/, /COUNT.*0/], atoms: ['aggregation_result_unexpected_zero'] },
  { keywords: [/空$/, /空的$/, /为空/, /结果为空/], atoms: ['aggregation_result_empty'] },
  { keywords: [/循环依赖/, /自引用/, /self[- ]reference/, /循环/], atoms: ['self_reference_detected'] },
  { keywords: [/ARRAYFORMULA/, /ARRAYFORMULA.*报错/, /数组公式/, /数组无法扩展/], atoms: ['arrayformula_misplacement', 'auto_expand_required'] },
  { keywords: [/文本.*数字/, /当成文本/, /被当成文本/, /转成数字/], atoms: ['text_number_coercion_needed'] },
  { keywords: [/跨表/, /跨文件/, /IMPORTRANGE/, /允许访问/, /链接失效/], atoms: ['cross_sheet_data_missing', 'cross_file_permission_issue'] },
  { keywords: [/下拉/, /联动/, /选择器/, /下拉选项/], atoms: ['dropdown_driven_calculation', 'selection_change_no_effect'] },
  { keywords: [/设计.*问题/, /拆成/, /结构.*问题/, /拆分/], atoms: ['structure_should_be_refactored', 'schema_role_confusion'] },
  { keywords: [/长期/, /复用/, /维护/, /以后.*要/, /每天.*不想管/], atoms: ['long_term_maintainability_concern', 'fear_of_future_breakage'] },
  { keywords: [/别人用/, /同事/, /权限/, /账号/, /多用户/], atoms: ['multi_user_safety_required'] },
  { keywords: [/汇总/, /明细/, /汇总表/, /不是明细/, /总计/, /明细/], atoms: ['schema_role_confusion', 'summary_and_detail_mixed'] },
  { keywords: [/公式.*不懂/, /不会写公式/, /QUERY.*不会/, /公式.*写错/], atoms: ['formula_logic_not_understood'] },
];

export function mapToSemanticAtoms(text: string): SemanticAtom[] {
  const lower = (text || '').toLowerCase();
  const found = new Set<SemanticAtom>();
  for (const entry of KEYWORD_MAP) {
    for (const re of entry.keywords) {
      if (re.test(lower)) {
        for (const a of entry.atoms) found.add(a);
        break;
      }
    }
  }

  // 如果包含请求自动扩展的语句，但没有明确提到拖拽，则补充 auto_expand_required
  if ((/整列|整列算|一行公式管一整列|不用拖/).test(lower)) {
    found.add('auto_expand_required');
  }

  return Array.from(found);
}

export function compressToSuperIntent(atoms: SemanticAtom[]): SuperIntent {
  if (atoms.length === 0) return SUPER_INTENTS.AUTOMATION;
  // Heuristic: failure-related atoms map to FAILURE, automation-related to AUTOMATION, structure to STRUCTURE, maintainability to MAINTAINABILITY
  const failureKeys = new Set(['aggregation_result_unexpected_zero','aggregation_result_empty','self_reference_detected','arrayformula_misplacement','error_message_not_actionable','text_number_coercion_needed','cross_sheet_data_missing']);
  const structureKeys = new Set(['structure_should_be_refactored','schema_role_confusion','summary_and_detail_mixed']);
  const maintainKeys = new Set(['long_term_maintainability_concern','multi_user_safety_required','fear_of_future_breakage']);

  let score = { failure: 0, structure: 0, maintain: 0, automation: 0 };
  for (const a of atoms) {
    if (failureKeys.has(a)) score.failure++;
    if (structureKeys.has(a)) score.structure++;
    if (maintainKeys.has(a)) score.maintain++;
    if (a === 'auto_expand_required' || a === 'no_manual_dragging' || a === 'dropdown_driven_calculation') score.automation++;
  }

  // pick highest
  const max = Math.max(score.failure, score.structure, score.maintain, score.automation);
  if (max === score.failure) return SUPER_INTENTS.FAILURE;
  if (max === score.structure) return SUPER_INTENTS.STRUCTURE;
  if (max === score.maintain) return SUPER_INTENTS.MAINTAINABILITY;
  return SUPER_INTENTS.AUTOMATION;
}

export default { mapToSemanticAtoms, compressToSuperIntent };
