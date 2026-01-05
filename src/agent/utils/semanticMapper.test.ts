import { mapToSemanticAtoms, compressToSuperIntent, SUPER_INTENTS } from './semanticMapper';

describe('semanticMapper', () => {
  it('maps auto expand and no drag phrases', () => {
    const atoms = mapToSemanticAtoms('我不想每次新增一行就拖公式');
    expect(atoms).toContain('no_manual_dragging');
    expect(atoms).toContain('auto_expand_required');
    expect(compressToSuperIntent(atoms)).toBe(SUPER_INTENTS.AUTOMATION);
  });

  it('maps aggregation zero cases to failure', () => {
    const atoms = mapToSemanticAtoms('为什么SUM算出来是0');
    expect(atoms).toContain('aggregation_result_unexpected_zero');
    expect(compressToSuperIntent(atoms)).toBe(SUPER_INTENTS.FAILURE);
  });

  it('maps cross file issues to failure', () => {
    const atoms = mapToSemanticAtoms('IMPORTRANGE怎么一直是0 点了允许访问还是不显示');
    expect(atoms).toContain('cross_file_permission_issue');
    expect(atoms).toContain('cross_sheet_data_missing');
    expect(compressToSuperIntent(atoms)).toBe(SUPER_INTENTS.FAILURE);
  });

  it('detects arrayformula spill and maps to automation/failure', () => {
    const atoms = mapToSemanticAtoms('ARRAYFORMULA 溢出 报错，数组无法扩展');
    expect(atoms).toContain('arrayformula_misplacement');
    expect(atoms).toContain('auto_expand_required');
    const superIntent = compressToSuperIntent(atoms);
    expect([SUPER_INTENTS.AUTOMATION, SUPER_INTENTS.FAILURE]).toContain(superIntent);
  });

  it('detects text-number coercion', () => {
    const atoms = mapToSemanticAtoms('列里都是文本格式，能不能帮我转成数字');
    expect(atoms).toContain('text_number_coercion_needed');
    expect(compressToSuperIntent(atoms)).toBe(SUPER_INTENTS.FAILURE);
  });

  it('detects maintainability concerns and multi-user safety', () => {
    const atoms = mapToSemanticAtoms('这个表太难维护了，别人也会编辑，怕被改坏');
    expect(atoms).toContain('long_term_maintainability_concern');
    expect(atoms).toContain('multi_user_safety_required');
    expect(compressToSuperIntent(atoms)).toBe(SUPER_INTENTS.MAINTAINABILITY);
  });

  it('detects dropdown driven calculations', () => {
    const atoms = mapToSemanticAtoms('下拉选项改变后，计算没有联动');
    expect(atoms).toContain('dropdown_driven_calculation');
    expect(atoms).toContain('selection_change_no_effect');
  });

  it('detects circular reference mentions', () => {
    const atoms = mapToSemanticAtoms('提示有循环引用，REF错误');
    expect(atoms).toContain('circular_reference');
    expect(atoms).toContain('self_reference_detected');
  });

  it('detects date/time parsing problems', () => {
    const atoms = mapToSemanticAtoms('这个日期列格式不对，导入后日期全错');
    expect(atoms).toContain('date_time_parsing_issue');
  });

  it('detects locale decimal comma issues', () => {
    const atoms = mapToSemanticAtoms('有些地区用逗号作为小数点，计算结果不对');
    expect(atoms).toContain('locale_decimal_comma');
    expect(atoms).toContain('text_number_coercion_needed');
  });

  it('detects merged cells and hidden rows', () => {
    const atoms = mapToSemanticAtoms('合并单元格导致筛选和公式不对，有些行被隐藏看不到');
    expect(atoms).toContain('merged_cells_issue');
    expect(atoms).toContain('hidden_rows_columns');
  });

  it('detects named range and external link issues', () => {
    const atoms = mapToSemanticAtoms('命名范围找不到，外部链接失效无法导入');
    expect(atoms).toContain('named_range_missing');
    expect(atoms).toContain('external_link_broken');
    expect(compressToSuperIntent(atoms)).toBe(SUPER_INTENTS.FAILURE);
  });

  it('detects formula performance issues', () => {
    const atoms = mapToSemanticAtoms('表格一打开就很慢，公式计算太慢，可能是性能问题');
    expect(atoms).toContain('formula_performance_issue');
  });

  // Batch 2 tests
  it('detects vlookup match failures', () => {
    const atoms = mapToSemanticAtoms('VLOOKUP查不到数据，返回错误');
    expect(atoms).toContain('vlookup_match_failure');
    expect(atoms).toContain('formula_logic_not_understood');
  });

  it('detects index/match issues', () => {
    const atoms = mapToSemanticAtoms('INDEX MATCH 公式 返回错误');
    expect(atoms).toContain('index_match_issue');
    expect(atoms).toContain('formula_logic_not_understood');
  });

  it('detects sorting issues', () => {
    const atoms = mapToSemanticAtoms('排序后数据乱了，排序不对');
    expect(atoms).toContain('sorting_issue');
  });

  it('detects filter issues', () => {
    const atoms = mapToSemanticAtoms('筛选后数据丢失，筛选不对');
    expect(atoms).toContain('filter_issue');
  });

  it('detects data cleanup and duplicate data needs', () => {
    const atoms = mapToSemanticAtoms('数据很脏，有大量重复值需要去重');
    expect(atoms).toContain('data_cleanup_needed');
    expect(atoms).toContain('duplicate_data');
  });

  it('detects blank cells issues', () => {
    const atoms = mapToSemanticAtoms('有很多空白单元格和空格问题');
    expect(atoms).toContain('blank_cells_issue');
    expect(atoms).toContain('data_cleanup_needed');
  });

  it('detects column type mismatch', () => {
    const atoms = mapToSemanticAtoms('列里混合类型，数据类型不匹配');
    expect(atoms).toContain('column_type_mismatch');
    expect(atoms).toContain('text_number_coercion_needed');
  });

  it('detects conditional format issues', () => {
    const atoms = mapToSemanticAtoms('条件格式不显示，高亮不工作');
    expect(atoms).toContain('conditional_format_issue');
  });

  it('detects chart data issues', () => {
    const atoms = mapToSemanticAtoms('图表数据显示不对，chart问题');
    expect(atoms).toContain('chart_data_issue');
  });

  it('detects print layout issues', () => {
    const atoms = mapToSemanticAtoms('打印问题，页面设置不对，分页问题');
    expect(atoms).toContain('print_layout_issue');
  });

  it('detects copy paste issues', () => {
    const atoms = mapToSemanticAtoms('复制粘贴后格式不对');
    expect(atoms).toContain('copy_paste_issue');
  });

  it('detects file corruption concerns', () => {
    const atoms = mapToSemanticAtoms('文件损坏打不开');
    expect(atoms).toContain('file_corruption_concern');
  });

  it('detects version compatibility issues', () => {
    const atoms = mapToSemanticAtoms('旧版本打开新文件，兼容性问题');
    expect(atoms).toContain('version_compatibility_issue');
  });
});
