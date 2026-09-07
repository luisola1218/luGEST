"""Exercise independent quote editors with explicit in-memory ports, no runtime."""
from pathlib import Path
import os
import sys
from dataclasses import fields
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')


def main():
    from PySide6.QtWidgets import QApplication, QWidget, QDialog, QMessageBox, QPushButton, QDialogButtonBox, QFileDialog
    from lugest_core.search import search_normalize
    from lugest_modules.quotes.application.editor_ports import QuoteEditorPorts
    from lugest_modules.quotes.presentation.material_prices import edit_material_prices, sync_stock_price
    from lugest_modules.quotes.presentation.material_editor import edit_material
    from lugest_modules.quotes.presentation.product_editor import edit_product
    from lugest_modules.quotes.presentation.labor_editor import edit_labor
    from lugest_modules.quotes.presentation.consumable_editor import edit_consumable
    from lugest_modules.quotes.presentation.structure_editor import edit_structure
    from lugest_modules.quotes.presentation.line_editor import edit_line
    from lugest_modules.quotes.presentation.profile_editor import edit_profile, ProfileEditorPorts
    from lugest_modules.quotes.domain.lines import service_line, product_line

    def forbidden(*args, **kwargs):
        raise AssertionError('Unexpected persistence or network access')

    values = {field.name: forbidden for field in fields(QuoteEditorPorts)}
    for name in ['material_rows', 'material_section_options', 'material_family_options',
                 'material_profile_size_options', 'material_price_rows', 'order_reference_rows',
                 'quote_parse_operacoes_lista', 'build_operacoes_fluxo']:
        values[name] = lambda *args, **kwargs: []
    for name in ['material_price_preview', 'material_geometry_preview', 'operation_cost_estimate',
                 'order_presets', 'material_by_id', 'product_catalog_options']:
        values[name] = lambda *args, **kwargs: {}
    product = {'codigo': 'P-1', 'descricao': 'Parafuso', 'pvp1': 2.5, 'unid': 'UN'}
    values.update(
        ORC_LINE_TYPE_PIECE='peca', ORC_LINE_TYPE_PRODUCT='produto', ORC_LINE_TYPE_SERVICE='servico',
        norm_text=search_normalize, normalize_operacao_nome=lambda value: str(value),
        normalize_orc_line_type=lambda value: str(value), detect_materia_formato=lambda *args: 'Chapa',
        material_default_price_kg=lambda *args: 3.15,
        material_family_profile=lambda *args: {'key': 'steel', 'density': 7.85},
        _fmt=lambda value: str(value), _parse_float=lambda value, default=0: float(value or default),
        ne_product_options=lambda *args: [dict(product)], orc_suggest_ref_interna=lambda *args, **kwargs: 'REF-1',
    )
    ports = QuoteEditorPorts(**values)
    app = QApplication.instance() or QApplication([])
    owner = QWidget()  # Deliberately has no backend, presets, line_rows or page methods.
    cases = [
        (edit_material_prices, {}),
        (edit_material, {'presets': {}, 'initial': {'calc_mode': 'Manual'}}),
        (edit_product, {'initial': {'produto_codigo': 'P-1', 'qtd': 3}}),
        (edit_labor, {'presets': {}, 'initial': {'hours': 2, 'hour_rate': 25}}),
        (edit_consumable, {'initial': {'quantity_units': 4, 'unit_price': 3}}),
        (edit_structure, {'workcenter': 'Montagem'}),
        (edit_line, {'presets': {}, 'client_code': 'C-1', 'current_number': '', 'line_rows': []}),
    ]
    errors = []
    with patch('socket.socket.connect', forbidden), patch('socket.create_connection', forbidden), \
         patch.object(QMessageBox, 'warning', forbidden), \
         patch.object(sys, 'excepthook', lambda *args: errors.append(args)):
        for editor, kwargs in cases:
            with patch.object(QDialog, 'exec', lambda dialog: QDialog.Rejected):
                assert editor(owner, ports, **kwargs) is None, editor.__name__
        for editor, kwargs in cases[2:6]:
            with patch.object(QDialog, 'exec', lambda dialog: QDialog.Accepted):
                result = editor(owner, ports, **kwargs)
                assert isinstance(result, dict), editor.__name__
                if editor is edit_labor:
                    assert result['total_cost'] == 50 and result['line']['qtd'] == 2
                if editor is edit_consumable:
                    assert result['total_cost'] == 12
                if editor is edit_product:
                    assert result['line']['produto_codigo'] == 'P-1'
        assert sync_stock_price(owner, ports, '', 10, 5) is None
        assert sync_stock_price(owner, ports, 'M-1', 5, 5) is None
        profile_ports = ProfileEditorPorts(
            ORC_LINE_TYPE_SERVICE='servico', laser_quote_settings=lambda: {},
            laser_quote_save_settings=forbidden, material_presets=lambda: {},
            profile_laser_quote_analyze=lambda payload: {},
            profile_laser_quote_build_line=lambda payload: {'line': {'descricao': 'Perfil', 'qtd': 1}},
        )
        with patch.object(QDialog, 'exec', lambda dialog: QDialog.Rejected):
            assert edit_profile(owner, profile_ports) is None
        def accept_profile(dialog):
            add = next(button for button in dialog.findChildren(QPushButton)
                       if button.text() == 'Adicionar STEP/IGS')
            add.click()
            dialog.findChild(QDialogButtonBox).button(QDialogButtonBox.Ok).click()
            return dialog.result()
        with patch.object(QDialog, 'exec', accept_profile), \
             patch.object(QFileDialog, 'getOpenFileNames', return_value=(['test-IPE100.step'], '')), \
             patch('lugest_modules.quotes.presentation.profile_editor.analyze_profile_cut_features', return_value={}), \
             patch('lugest_modules.quotes.presentation.profile_editor.render_step_preview_image', return_value={}):
            lines = edit_profile(owner, profile_ports)
            assert lines and lines[0]['line_origin'] == 'step_igs_profile_laser'
            assert lines[0]['tipo_item'] == 'servico'
    assert not errors, errors
    assert service_line('Servico', 0, 'h', 2, 'Montagem', line_type='servico') is None
    assert product_line(product, -1, line_type='produto') is None
    assert product_line(product, 3, line_type='produto')['preco_unit'] == 2.5
    assert product == {'codigo': 'P-1', 'descricao': 'Parafuso', 'pvp1': 2.5, 'unid': 'UN'}
    assert 'main' not in sys.modules
    assert 'lugest_qt.ui.pages.quotes_page' not in sys.modules
    owner.close()
    print('quote-editors-ok cancel=8 accept=5 page-independent=yes no-main=yes no-network=yes')


if __name__ == '__main__':
    main()
