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
});
