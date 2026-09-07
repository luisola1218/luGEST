"""Inventory queries and stock issues with no global runtime or database."""
from copy import deepcopy
from datetime import date
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def main():
    from lugest_modules.inventory.application.product_queries import ProductQueries, ProductQueryRules
    from lugest_modules.inventory.application.consumption import StockIssueService
    from lugest_modules.inventory.infrastructure.legacy_product_repository import LegacyProductReadRepository
    from lugest_modules.inventory.infrastructure.legacy_stock_issue_repository import LegacyStockIssueRepository
    state = [{'produtos': [{'codigo': 'P1', 'descricao': 'Parafuso', 'qty': 8,
                           'quality_pending_qty': 2, 'p_compra': 4, 'meta': {'keep': True}}],
              'produtos_mov': []}]
    parse = lambda value, default=0: float(value or default)
    reads = LegacyProductReadRepository(products=lambda: state[0]['produtos'],
                                         movements=lambda: state[0]['produtos_mov'])
    rules = ProductQueryRules(
        _fmt=str, _parse_float=parse, _product_resolve_catalog_fields=lambda row: {},
        _quality_status_is_available=lambda status: status == 'LIBERTADO',
        inventory_scan_code=lambda kind, code: f'{kind}|{code}',
        produto_preco_unitario=lambda row: row.get('p_compra', 0),
        produto_preco_venda=lambda row: row.get('pvp1', 0),
    )
    queries = ProductQueries(reads, rules, today=lambda: date(2026, 9, 7))
    rows = queries.product_rows('paraf')
    assert len(rows) == 1 and rows[0]['qty'] == 10 and rows[0]['available_qty'] == 8
    assert rows[0]['valor_stock'] == 40 and rows[0]['severity'] == 'warning'
    assert queries.product_rows('missing') == []
    detail = queries.product_detail('P1')
    detail['meta']['keep'] = False
    assert state[0]['produtos'][0]['meta']['keep'] is True
    assert queries.product_movement_years() == ['2026']
    fail = [False]
    def save(**kwargs):
        if fail[0]:
            raise OSError('Persistence failed')
        state[0] = deepcopy(state[0])
    def add(data, **movement):
        data['produtos_mov'].append({'data': '2026-09-07T12:00:00', **movement})
    repository = LegacyStockIssueRepository(get_data=lambda: state[0], save_dataset=save,
                                            add_movement=add, parse_float=parse)
    service = StockIssueService(repository, parse_float=parse, unit_price=rules.produto_preco_unitario,
                                now=lambda: '2026-09-07T12:00:00')
    before = deepcopy(state[0])
    for quantity, kwargs in [(0, {}), (9, {}), (float('nan'), {}), (float('inf'), {}),
                             (2, {'issue_mode': 'operator'})]:
        try:
            service.consume('P1', quantity, actor='admin', **kwargs)
        except ValueError:
            pass
        else:
            raise AssertionError('Invalid consumption accepted')
        assert state[0] == before, 'Validation mutated stock'
    fail[0] = True
    try:
        service.consume('P1', 2, actor='admin')
    except OSError:
        pass
    else:
        raise AssertionError('Failed stock write accepted')
    assert state[0] == before, 'Failed persistence left a stock change or movement'
    fail[0] = False
    result = service.consume('P1', 2, actor='admin', issue_mode='operator', target_operator='Luis')
    assert result['antes'] == 8 and result['depois'] == 6
    assert state[0]['produtos'][0]['qty'] == 6
    summary = queries.product_issue_summary(operator_name='luis', year='2026', codigo='P1')
    assert summary == {'linhas': 1, 'qtd_total': 2.0, 'valor_total': 8.0}
    assert queries.product_movements(year='2025') == []
    try:
        repository.issue('P1', expected_quantity=8, quantity_after=5, updated_at='', movement={})
    except ValueError:
        pass
    else:
        raise AssertionError('Stale quantity accepted')
    assert state[0]['produtos'][0]['qty'] == 6
    assert 'main' not in sys.modules
    print('inventory-services-ok physical-stock=yes detached=yes operator-validation=yes rollback=yes stale-quantity=yes')


if __name__ == '__main__':
    main()
