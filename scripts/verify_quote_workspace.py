"""Construct quote panels without a controller; exercise signals and debounce."""
from dataclasses import fields
import os
from pathlib import Path
import sys
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
os.environ.setdefault('QT_QPA_PLATFORM', 'offscreen')


def main():
    from PySide6.QtWidgets import QApplication, QComboBox, QDialog, QMessageBox
    from PySide6.QtTest import QTest
    from lugest_modules.quotes.presentation.workspace import QuoteWorkspace, QuoteWorkspaceActions
    app = QApplication.instance() or QApplication([])
    calls = []
    actions = QuoteWorkspaceActions(**{
        field.name: (lambda *args, name=field.name, **kwargs: calls.append((name, args)))
        for field in fields(QuoteWorkspaceActions)
    })
    errors = []
    with patch.object(sys, 'excepthook', lambda *args: errors.append(args)):
        view = QuoteWorkspace()
        view.build(actions)
        assert view.view_stack.count() == 2
        assert view.lines_table.columnCount() == 14
        calls.clear()
        view.new_quote_btn.click()
        assert calls[-1][0] == '_new_quote'
        calls.clear()
        for text in ['O', 'OR', 'ORC-123']:
            view.filter_edit.lineEdit().setText(text)
        assert not any(name == 'refresh' for name, _ in calls)
        QTest.qWait(240)
        assert [name for name, _ in calls].count('refresh') == 1
        assert not errors, errors
        assert not hasattr(view, 'backend') and not hasattr(view, 'line_rows')
        view.close()
        from lugest_modules.quotes.presentation.page import QuotePage
        from lugest_modules.quotes.presentation.page_services import QuotePageServices
        values = {field.name: (lambda *args, **kwargs: None) for field in fields(QuotePageServices)}
        paired_saves = []
        values["save_assembly_pair"] = lambda template, live: paired_saves.append((template, live))
        values.update(ORC_LINE_TYPE_PIECE='peca', ORC_LINE_TYPE_PRODUCT='produto', ORC_LINE_TYPE_SERVICE='servico',
                      current_user=lambda: {}, branding=lambda: {}, norm_text=lambda value: str(value).lower(),
                      normalize_orc_line_type=lambda value: str(value or 'peca'),
                      orc_line_type_label=lambda value: str(value), orc_line_is_piece=lambda line: False,
                      orc_line_is_product=lambda line: line.get('tipo_item') == 'produto',
                      quote_parse_operacoes_lista=lambda value: [])
        page = QuotePage(QuotePageServices(**values))
        page.line_rows = [{'tipo_item': 'servico', 'descricao': 'Teste', 'qtd': 2, 'preco_unit': 10}]
        page._render_quote_lines()
        assert 'Teste' in page.view.lines_table.item(0, page.LINE_COL_DESCRIPTION).text()
        page.view.lines_table.selectRow(0)
        def choose_both(dialog):
            combo = next(widget for widget in dialog.findChildren(QComboBox) if widget.findData("both") >= 0)
            combo.setCurrentIndex(combo.findData("both"))
            return QDialog.Accepted
        with patch.object(QDialog, "exec", choose_both), patch.object(QMessageBox, "information", lambda *args: None):
            page._save_selected_lines_as_group()
        assert len(paired_saves) == 1 and paired_saves[0][0]["itens"]
        assert not hasattr(page, 'backend')
        assert not errors, errors
        page.close()
    assert 'main' not in sys.modules
    assert 'lugest_qt.ui.pages.quotes_page' not in sys.modules
    print('quote-workspace-ok standalone-view=yes standalone-controller=yes panels=yes signal-routing=yes filter-debounce=yes no-backend=yes')


if __name__ == '__main__':
    main()
