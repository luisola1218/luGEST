"""Enforce module boundaries and imports, including editor callback globals."""
import ast
import builtins
import importlib
from pathlib import Path
import symtable
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def references(table):
    names = {s.get_name() for s in table.get_symbols() if s.is_referenced() and s.is_global()}
    for child in table.get_children():
        names |= references(child)
    return names


def main():
    for module in ('clients', 'inventory', 'quotes', 'transport', 'billing', 'quality'):
        importlib.import_module(f'lugest_modules.{module}.api')
    assert not any(name.startswith('PySide6') for name in sys.modules), 'Business APIs loaded Qt'
    assert 'main' not in sys.modules
    files = list((ROOT / 'lugest_modules').rglob('*.py'))
    for path in files:
        relative = path.relative_to(ROOT)
        parts = relative.parts
        source = path.read_text(encoding='utf-8')
        tree = ast.parse(source)
        imports = []
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                imports.extend(alias.name for alias in node.names)
            elif isinstance(node, ast.ImportFrom) and node.module:
                imports.append(node.module)
            if any(layer in parts for layer in ('presentation', 'application', 'domain')):
                assert not (isinstance(node, ast.Attribute) and node.attr in {'desktop_main', 'ensure_data', '_base_data_snapshot'}), relative
        for imported in imports:
            assert imported != 'main' and not imported.startswith('lugest_desktop'), (relative, imported)
            if 'domain' in parts or 'application' in parts:
                assert not imported.startswith(('lugest_qt', 'lugest_infra', 'PySide6', 'pymysql')), (relative, imported)
                assert '.presentation' not in imported and '.infrastructure' not in imported, (relative, imported)
            if imported.startswith('lugest_modules.'):
                target = imported.split('.')
                assert target[1] == parts[1] or target[2:] == ['api'], (relative, 'Cross-module import must use api', imported)
        module_name = '.'.join(relative.with_suffix('').parts)
        if module_name.endswith('.__init__'):
            module_name = module_name[:-9]
        module = importlib.import_module(module_name)
        missing = references(symtable.symtable(source, str(path), 'exec'))
        missing -= set(vars(module)) | set(vars(builtins)) | {'__class__'}
        assert not missing, (relative, missing)
    assert 'main' not in sys.modules
    assert 'lugest_qt.ui.pages.quotes_page' not in sys.modules
    from backend_map import backend_methods, business_implementations
    methods = backend_methods()
    for method, target in {'client_rows': 'ClientService.rows',
                           'orc_save_nesting_study': 'NestingStudyService.save',
                           'product_consume': 'StockIssueService.consume',
                           'product_save': 'ProductCommands.save',
                           'orc_save': 'QuoteCommands.save',
                           'orc_detail': 'QuoteQueries.detail',
                           'conjunto_refresh_prices': 'AssemblyRefresh.refresh',
                           'conjunto_save': 'AssemblyCatalog.save',
                           'orc_purchase_needs': 'PurchaseNeeds.rows',
                           'orc_convert_to_order': 'QuoteConversion.convert',
                           'transport_tariff_save': 'TariffService.save',
                           'transport_move_stop': 'TransportStops.move',
                           'billing_add_payment': 'Payments.add',
                           'quality_nc_save': 'Nonconformities.save',
                           'assembly_model_remove': 'AssemblyCatalog.remove',
                           '_normalize_assembly_model_item': 'normalize_item',
                           'conjunto_expand': 'expand_model',
                           '_normalize_orc_line': 'normalize_line'}.items():
        assert target in business_implementations(methods[method][0], methods), method
    assert 'main' not in sys.modules
    print(f'business-modules-ok files={len(files)} boundaries=yes callback-globals=yes no-main=yes')


if __name__ == '__main__':
    main()
