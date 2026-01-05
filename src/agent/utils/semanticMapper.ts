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
  'circular_reference',
  'date_time_parsing_issue',
  'locale_decimal_comma',
  'precision_rounding_issue',
  'hidden_rows_columns',
  'merged_cells_issue',
  'named_range_missing',
  'external_link_broken',
  'formula_performance_issue',
  'structure_should_be_refactored',
  'fear_of_future_breakage',
  // Batch 2
  'vlookup_match_failure',
  'index_match_issue',
  'sorting_issue',
  'filter_issue',
  'data_cleanup_needed',
  'duplicate_data',
  'blank_cells_issue',
  'column_type_mismatch',
  'conditional_format_issue',
  'chart_data_issue',
  'print_layout_issue',
  'copy_paste_issue',
  'undo_redo_issue',
  'file_corruption_concern',
  'version_compatibility_issue',
] as const;

export type SemanticAtom = typeof SEMANTIC_ATOMS[number];

// 简单的关键词到原子映射表（可扩展）
const KEYWORD_MAP: Array<{ keywords: RegExp[]; atoms: SemanticAtom[] }> = [
  { keywords: [/拖公式/, /不想.*拖/, /不想.*复制/, /不用拖/], atoms: ['no_manual_dragging', 'auto_expand_required'] },
  { keywords: [/新行/, /新增.*行/, /每天.*加数据/, /每天.*加/], atoms: ['future_rows_unknown', 'auto_expand_required'] },
  { keywords: [/结果.*是0/, /算出来是0/, /sum.*0/, /count.*0/, /#div\/0!/, /除数为0/], atoms: ['aggregation_result_unexpected_zero'] },
  { keywords: [/空$/, /空的$/, /为空/, /结果为空/, /#n\/a/, /#na/], atoms: ['aggregation_result_empty'] },
  { keywords: [/循环依赖/, /自引用/, /self[- ]reference/, /循环/, /#ref!/, /引用错误/], atoms: ['self_reference_detected'] },
  { keywords: [/arrayformula/, /arrayformula.*报错/, /数组公式/, /数组无法扩展/, /溢出/, /spill/, /#spill!/], atoms: ['arrayformula_misplacement', 'auto_expand_required'] },
  { keywords: [/文本.*数字/, /当成文本/, /被当成文本/, /转成数字/, /格式.*文本/, /数字转换/], atoms: ['text_number_coercion_needed'] },
  { keywords: [/跨表/, /跨文件/, /importrange/, /允许访问/, /链接失效/, /无法导入/, /外部表格/], atoms: ['cross_sheet_data_missing', 'cross_file_permission_issue'] },
  { keywords: [/下拉/, /联动/, /选择器/, /下拉选项/, /数据验证/], atoms: ['dropdown_driven_calculation', 'selection_change_no_effect'] },
  { keywords: [/设计.*问题/, /拆成/, /结构.*问题/, /拆分/, /合并单元格/, /命名范围/, /命名范围缺失/, /schema/], atoms: ['structure_should_be_refactored', 'schema_role_confusion'] },
  { keywords: [/长期/, /复用/, /维护/, /以后.*要/, /每天.*不想管/, /难以维护/, /以后会出问题/], atoms: ['long_term_maintainability_concern', 'fear_of_future_breakage'] },
  { keywords: [/别人用/, /别人/, /同事/, /权限/, /账号/, /多用户/, /并发编辑/, /冲突/, /怕被改/, /怕.*改/], atoms: ['multi_user_safety_required'] },
  { keywords: [/汇总/, /明细/, /汇总表/, /不是明细/, /总计/, /明细/], atoms: ['schema_role_confusion', 'summary_and_detail_mixed'] },
  { keywords: [/公式.*不懂/, /不会写公式/, /query.*不会/, /公式.*写错/, /vlookup.*?不匹配/, /match.*不会/], atoms: ['formula_logic_not_understood'] },
  { keywords: [/提示.*无用/, /错误信息.*无法理解/, /错误信息.*不可操作/, /#value!/, /#value/, /错误提示/], atoms: ['error_message_not_actionable'] },
  { keywords: [/透视表/, /pivot/, /过滤视图/, /filter view/, /筛选视图/, /筛选条件/], atoms: ['structure_should_be_refactored'] },
  { keywords: [/保护工作表/, /保护/, /只读/, /锁定/, /权限不足/], atoms: ['multi_user_safety_required'] },
  { keywords: [/循环引用/, /circular reference/, /循环引用/, /环路/, /circular/], atoms: ['circular_reference', 'self_reference_detected'] },
  { keywords: [/日期.*解析/, /日期.*格式/, /时间.*解析/, /时间.*格式/, /日期.*不对/], atoms: ['date_time_parsing_issue'] },
  { keywords: [/小数点.*逗号/, /小数.*逗号/, /千分位/, /逗号作为小数/, /locale/, /区域设置/], atoms: ['locale_decimal_comma', 'text_number_coercion_needed'] },
  { keywords: [/精度/, /四舍五入/, /保留.*位/, /浮点误差/, /舍入误差/], atoms: ['precision_rounding_issue'] },
  { keywords: [/隐藏行/, /隐藏列/, /看不到/, /被隐藏/], atoms: ['hidden_rows_columns'] },
  { keywords: [/合并单元格/, /合并了/, /merged cells/, /合并导致/], atoms: ['merged_cells_issue'] },
  { keywords: [/命名范围/, /命名区域/, /named range/, /命名范围缺失/], atoms: ['named_range_missing', 'schema_role_confusion'] },
  { keywords: [/外部链接/, /链接断开/, /链接失效/, /external link/, /连接不到外部/], atoms: ['external_link_broken', 'cross_sheet_data_missing'] },
  { keywords: [/太慢/, /公式太慢/, /性能问题/, /计算过慢/, /性能瓶颈/], atoms: ['formula_performance_issue'] },
  // Batch 2 keyword mappings
  { keywords: [/vlookup.*找不到/, /vlookup.*错误/, /vlookup.*不匹配/, /查找.*失败/, /查不到/], atoms: ['vlookup_match_failure', 'formula_logic_not_understood'] },
  { keywords: [/index.*match/, /索引.*匹配/, /index.*不对/, /match.*返回错误/], atoms: ['index_match_issue', 'formula_logic_not_understood'] },
  { keywords: [/排序.*乱了/, /排序.*不对/, /排序.*问题/, /sorting/, /排序后.*丢失/], atoms: ['sorting_issue'] },
  { keywords: [/筛选.*不对/, /筛选.*丢失/, /筛选.*问题/, /filter.*不工作/, /自动筛选/], atoms: ['filter_issue'] },
  { keywords: [/数据清洗/, /清理数据/, /数据.*脏/, /格式.*乱/, /需要清理/], atoms: ['data_cleanup_needed'] },
  { keywords: [/重复.*数据/, /去重/, /duplicate/, /重复值/, /重复项/], atoms: ['duplicate_data', 'data_cleanup_needed'] },
  { keywords: [/空白.*单元格/, /空格.*问题/, /有空格/, /空白行/, /空白列/], atoms: ['blank_cells_issue', 'data_cleanup_needed'] },
  { keywords: [/列.*类型/, /类型.*不一致/, /数据类型.*不匹配/, /混合类型/], atoms: ['column_type_mismatch', 'text_number_coercion_needed'] },
  { keywords: [/条件格式.*不显示/, /条件格式.*问题/, /conditional format/, /高亮.*不工作/], atoms: ['conditional_format_issue'] },
  { keywords: [/图表.*不更新/, /图表.*数据.*错/, /chart.*问题/, /图表.*显示不对/], atoms: ['chart_data_issue'] },
  { keywords: [/打印.*问题/, /打印.*乱/, /print.*问题/, /页面设置/, /分页.*问题/], atoms: ['print_layout_issue'] },
  { keywords: [/复制.*粘贴.*问题/, /粘贴.*格式/, /copy.*paste/, /粘贴.*不对/], atoms: ['copy_paste_issue'] },
  { keywords: [/撤销.*不行/, /undo.*问题/, /恢复不了/, /ctrl.*z.*不工作/], atoms: ['undo_redo_issue'] },
  { keywords: [/文件.*损坏/, /打不开/, /文件.*坏了/, /corrupt/, /无法打开/], atoms: ['file_corruption_concern'] },
  { keywords: [/版本.*不兼容/, /旧版本/, /新版本.*打不开/, /兼容性/, /xlsx.*xls/], atoms: ['version_compatibility_issue'] },
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
  // external links (IMPORTRANGE / 外部链接) indicate data access failure
  failureKeys.add('external_link_broken');
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
