"""Quote prices, reference allocation and write failures without the legacy app."""
from copy import deepcopy
from dataclasses import fields
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def main():
    from lugest_modules.quotes.application.commands import QuoteCommands, QuoteCommandRules
    from lugest_modules.quotes.infrastructure.legacy_quote_repository import LegacyQuoteWriteRepository
    state = [{'orcamentos': []}]
    baseline = {'orcamentos': []}
    stored = {}
    fail = [False]
    def upsert(data, quote):
        assert data is state[0]
        if fail[0]:
            raise OSError('quote write failed')
        stored[quote['numero']] = deepcopy(quote)
    def dataset(**kwargs):
        if fail[0]:
            raise OSError('dataset write failed')
        stored.clear()
        stored.update({row['numero']: deepcopy(row) for row in state[0]['orcamentos']})
        state[0] = deepcopy(state[0])
    def sync(quote):
        state[0].setdefault('orc_refs', {})['touched'] = quote['numero']
        state[0] = deepcopy(state[0])
    repository = LegacyQuoteWriteRepository(
        get_data=lambda: state[0], get_baseline=lambda: baseline, save_dataset=dataset,
        next_number=lambda: 'ORC-2026-0001', next_reference=lambda client, reserved: f'REF-{len(reserved)+1}',
        peek_number=lambda: 'unused', sync_registry=sync, upsert=upsert, delete_studies=lambda number: None,
    )
    rules = QuoteCommandRules(
        _active_client_ref_usage=lambda *args, **kwargs: (set(), set()),
        _known_client_ref_for_external=lambda *args: '', _known_client_ref_pairs=lambda *args: set(),
        _normalize_orc_client=deepcopy, _normalize_orc_line=deepcopy,
        _normalize_quote_discount_groups=lambda value: list(value or []),
        _normalize_quote_discount_mode=lambda value: value or 'total',
        _normalize_supplier_reference=lambda *args: ('', '', ''), _normalize_workcenter_value=str,
        _parse_float=lambda value, default=0: float(value or default),
        _quote_default_delivery_text=lambda: 'A combinar', _quote_line_is_raw_material=lambda line: False,
        _quote_standard_iva_perc=lambda: 23, _ref_client_code=lambda code: code,
        _repair_orc_ref_history=lambda code: None, current_year=lambda: 2026,
        now_iso=lambda: '2026-09-07T12:00:00', orc_line_is_piece=lambda line: line.get('tipo_item') == 'peca',
    )
    service = QuoteCommands(repository, rules)
    payload = {'cliente': {'codigo': 'CL1'}, 'linhas': [
        {'tipo_item': 'peca', 'qtd': 3, 'preco_unit': 10.1234, 'discount_group_key': 'A'},
        {'tipo_item': 'servico', 'qtd': 2, 'preco_unit': 4, 'discount_group_key': 'B'},
    ], 'desconto_perc': 25, 'desconto_modo': 'lotes_espessura', 'desconto_grupos': ['A'],
        'incremento_preco_perc': 10, 'preco_transporte': 5}
    original = deepcopy(payload)
    number = service.save(payload)
    quote = stored[number]
    assert quote['linhas'][0]['ref_interna'] == 'REF-1'
    assert quote['linhas'][1]['ref_interna'] == ''
    assert quote['linhas'][0]['preco_unit_desconto'] == 8.3518
    assert quote['linhas'][1]['preco_unit_desconto'] == 4.4
    assert quote['subtotal'] == 38.86 and quote['total'] == 47.80
    assert quote['desconto_valor'] == 8.35 and payload == original
    assert baseline['orcamentos'][0] == quote
    from lugest_modules.quotes.application.queries import QuoteQueries, QuoteQueryRules
    from lugest_modules.quotes.infrastructure.legacy_quote_read_repository import LegacyQuoteReadRepository
    from lugest_core.search import search_normalize
    query_values = {field.name: (lambda *args, **kwargs: {}) for field in fields(QuoteQueryRules)}
    for name in ('_normalize_orc_client', '_normalize_quote_discount_groups', '_normalize_quote_discount_mode',
                 '_normalize_workcenter_value', '_parse_float', '_quote_default_delivery_text', '_quote_standard_iva_perc'):
        query_values[name] = getattr(rules, name)
    query_values.update(_fmt=str, _orc_number_sort_key=lambda number: number,
                        current_year=lambda: 2026, extract_year=lambda stamp, number, year: year or str(stamp)[:4],
                        norm_text=search_normalize, normalize_orc_line_type=lambda value: value,
                        orc_line_is_piece=rules.orc_line_is_piece)
    queries = QuoteQueries(LegacyQuoteReadRepository(lambda: state[0]['orcamentos']), QuoteQueryRules(**query_values))
    assert queries.rows(year='2026')[0]['numero'] == number
    assert queries.rows()[0]['linhas'] == 2
    class HeavyStudy:
        def __deepcopy__(self, memo):
            raise AssertionError('The list copied a heavy field it does not display')
    state[0]['orcamentos'][0]['nesting_studies'] = HeavyStudy()
    assert queries.rows()[0]['linhas'] == 2
    state[0]['orcamentos'][0].pop('nesting_studies')
    assert queries.rows(year='2025') == [] and queries.rows(filter_text='missing') == []
    assert queries.years() == ['2026']
    detail = queries.detail(number)
    assert detail['total'] == quote['total']
    detail['cliente']['codigo'] = 'CHANGED'
    assert state[0]['orcamentos'][0]['cliente']['codigo'] == 'CL1'
    service.set_state(number, 'Aprovado')
    assert stored[number]['estado'] == 'Aprovado'
    before = deepcopy(state[0])
    before_baseline = deepcopy(baseline)
    fail[0] = True
    for operation in [lambda: service.save({**payload, 'numero': number, 'nota_cliente': 'failed'}),
                      lambda: service.set_state(number, 'Rejeitado'), lambda: service.remove(number)]:
        try:
            operation()
        except OSError:
            pass
        else:
            raise AssertionError('Quote write failure swallowed')
        assert state[0] == before and baseline == before_baseline
    fail[0] = False
    service.remove(number)
    assert number not in stored
    try:
        service.set_state('missing', 'Aprovado')
    except ValueError:
        pass
    else:
        raise AssertionError('Missing quote accepted')
    assert 'main' not in sys.modules
    print('quote-commands-ok group-discount=yes rounding=yes references=yes queries=yes detached=yes snapshot-refresh=yes failure-recovery=yes')


if __name__ == '__main__':
    main()
