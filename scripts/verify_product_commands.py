"""Product CRUD, movement integrity and rollback without a database or runtime."""
from copy import deepcopy
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def main():
    from lugest_modules.inventory.application.product_commands import ProductCommands
    from lugest_modules.inventory.application.product_definition import ProductDefinitionRules, normalize_product, price_preview
    from lugest_modules.inventory.infrastructure.legacy_product_write_repository import LegacyProductWriteRepository
    state = [{'produtos': [], 'produtos_mov': [], 'seq': {'produto': 1}, 'other': {'keep': True}}]
    persisted = {}
    fail = [False]
    parse = lambda value, default=0: float(value or default)
    rules = ProductDefinitionRules(
        unit_price=lambda row: row.get('p_compra', 0), _parse_float=parse,
        _product_dimensoes=lambda row: '10x20x2', _product_resolve_catalog_fields=lambda row: {},
        product_copilot_analysis=lambda description, code: {}, product_next_code=lambda: 'PRD-0001',
    )
    def normalize(payload):
        # Reproduce a dependency publishing a fresh snapshot while normalizing.
        state[0] = deepcopy(state[0])
        return normalize_product(rules, payload)
    def save(**kwargs):
        if fail[0]:
            raise OSError('write failed')
        persisted.clear()
        persisted.update(deepcopy(state[0]))
        state[0] = deepcopy(state[0])
    def movement(data, **row):
        data.setdefault('produtos_mov', []).append(row)
    def sequence(data, code):
        data.setdefault('seq', {})['produto'] = 2
    repository = LegacyProductWriteRepository(get_data=lambda: state[0], save_dataset=save,
                                              add_movement=movement, ensure_sequence=sequence)
    commands = ProductCommands(repository, normalize=normalize, parse_float=parse,
                               format_number=str, now=lambda: '2026-09-07T12:00:00')
    payload = {'descricao': 'Parafuso', 'qty': 10, 'p_compra': 2}
    assert commands.save(payload, actor='Luis') == 'PRD-0001'
    assert persisted['produtos'][0]['qty'] == 10
    assert persisted['produtos_mov'][0]['tipo'] == 'ENTRADA_INICIAL'
    assert persisted['produtos_mov'][0]['qtd'] == 10
    assert 'codigo' not in payload
    state[0]['produtos'][0]['custom_metadata'] = {'keep': True}
    edited = {'codigo': 'PRD-0001', 'descricao': 'Alterado', 'qty': 8, 'p_compra': 2}
    commands.save(edited, actor='Luis')
    assert persisted['produtos'][0]['custom_metadata']['keep']
    adjustment = persisted['produtos_mov'][-1]
    assert adjustment['tipo'] == 'AJUSTE_STOCK' and adjustment['qtd'] == 2
    assert adjustment['antes'] == 10 and adjustment['depois'] == 8
    commands.save(edited, actor='Luis')
    assert len(persisted['produtos_mov']) == 2, 'Unchanged stock created a movement'
    before = deepcopy(state[0])
    fail[0] = True
    for operation in [lambda: commands.save({**edited, 'qty': 4}, actor='Luis'),
                      lambda: commands.remove(['PRD-0001']),
                      lambda: commands.save({'codigo': 'P2', 'descricao': 'Novo', 'qty': 3}, actor='Luis')]:
        try:
            operation()
        except OSError:
            pass
        else:
            raise AssertionError('Write failure swallowed')
        assert state[0] == before, 'Failed operation left product, movement or sequence changes'
    fail[0] = False
    assert commands.remove([' PRD-0001 ', 'PRD-0001']) == 1
    assert persisted['produtos'] == [] and len(persisted['produtos_mov']) == 2
    for codes in [[], ['missing']]:
        try:
            commands.remove(codes)
        except ValueError:
            pass
        else:
            raise AssertionError('Empty or unknown selection accepted')
    assert price_preview(rules, {'qty': 2, 'p_compra': 3})['valor_stock'] == 6
    assert state[0]['other'] == {'keep': True}
    assert 'main' not in sys.modules
    print('product-commands-ok create=yes edit=yes remove=yes movements=yes snapshot-replacement=yes failure-recovery=yes')


if __name__ == '__main__':
    main()
