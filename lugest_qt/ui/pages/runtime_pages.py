"""Compatibility imports for historical callers; implementations live by area.

New code imports the dedicated page module. Lazy resolution keeps importing
one historical page from loading every other workspace.
"""
from importlib import import_module

_EXPORTS = {'AvariasPage': 'avarias_page',
 'ClientsPage': 'partners_pages',
 'ExpeditionPage': 'shipping_workspace',
 'LIST_TABLE_FONT_PX': 'runtime_support',
 'LIST_TABLE_ROW_PX': 'runtime_support',
 'LegacyExpeditionPage': 'shipping_workspace',
 'LegacyOperatorPage': 'operator_workspace',
 'LegacyOrdersPage': 'orders_workspace',
 'LegacyPlanningPage': 'planning_workspace',
 'LegacyPurchaseNotesPage': 'purchase_notes_workspace',
 'OperatorPage': 'operator_workspace',
 'OppPage': 'opp_page',
 'OrdersPage': 'orders_workspace',
 'PlanningBacklogTable': 'planning_workspace',
 'PlanningGridTable': 'planning_workspace',
 'PlanningPage': 'planning_workspace',
 'PulsePage': 'pulse_page',
 'PurchaseNotesPage': 'purchase_notes_page',
 'QuotesPage': 'quotes_page',
 'SuppliersPage': 'partners_pages',
 'TransportsPage': 'transports_page',
 '_OPERATION_PRICING_MODE_ITEMS': 'runtime_support',
 '_OperatorLabelsDialog': 'operator_workspace',
 '_adopt_layout_item': 'runtime_support',
 '_apply_progress_style': 'runtime_support',
 '_build_operation_selector': 'runtime_support',
 '_clear_layout_widgets': 'runtime_support',
 '_fmt_eur': 'runtime_support',
 '_format_client_label': 'runtime_support',
 '_is_dark': 'runtime_support',
 '_make_inline_progress': 'runtime_support',
 '_normalize_operation_text': 'runtime_support',
 '_open_operation_cost_profiles_dialog': 'runtime_support',
 '_open_quote_operation_detail_dialog': 'runtime_support',
 '_operation_pricing_mode_label': 'runtime_support',
 '_operation_tokens': 'runtime_support',
 '_operations_for_posto': 'runtime_support',
 '_piece_ops_progress': 'runtime_support',
 '_reference_catalog_dialog': 'runtime_support',
 '_split_client_label': 'runtime_support',
 '_take_layout_items': 'runtime_support'}

__all__ = [name for name in _EXPORTS if not name.startswith("_")]

def __getattr__(name: str):
    module = _EXPORTS.get(name)
    if module is None:
        raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
    value = getattr(import_module(f".{module}", __package__), name)
    globals()[name] = value
    return value


def __dir__():
    return sorted(set(globals()) | set(_EXPORTS))
