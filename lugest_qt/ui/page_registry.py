"""Page composition: import and construct only the workspace being opened."""
from __future__ import annotations

from functools import partial
from importlib import import_module


PAGE_DEFINITIONS = {
    'stock_dashboard': ('stock_dashboard_page', 'StockDashboardPage', ('backend',)),
    'pulse': ('pulse_page', 'PulsePage', ('runtime_service', 'backend')),
    'operator': ('operator_page', 'OperatorPage', ('runtime_service', 'backend')),
    'planning': ('planning_page', 'PlanningPage', ('runtime_service', 'backend')),
    'avarias': ('avarias_page', 'AvariasPage', ('runtime_service',)),
    'purchase_notes': ('purchase_notes_legacy_page', 'PurchaseNotesPage', ('backend',)),
    'shipping': ('shipping_page', 'ExpeditionPage', ('backend',)),
    'billing': ('billing_page', 'BillingPage', ('backend',)),
    'materials': ('materials_page', 'MaterialsPage', ('backend',)),
    'products': ('products_page', 'ProductsPage', ('backend',)),
    'direct_services': ('direct_services_page', 'DirectServicesPage', ('backend',)),
    'clients': ('partners_pages', 'ClientsPage', ('backend',)),
    'suppliers': ('partners_pages', 'SuppliersPage', ('backend',)),
    'orders': ('orders_page', 'OrdersPage', ('backend',)),
    'quotes': ('quotes_page', 'QuotesPage', ('backend',)),
    'opp': ('opp_page', 'OppPage', ('backend',)),
    'material_assistant': ('material_assistant_page', 'MaterialAssistantPage', ('backend',)),
    'transportes': ('transports_page', 'TransportsPage', ('backend',)),
    'quality': ('quality_page', 'QualityPage', ('backend',)),
    'diagnostics': ('diagnostics_page', 'DiagnosticsPage', ('backend',)),
}


def _create_page(module: str, class_name: str, arguments: tuple):
    page_class = getattr(import_module(f".{module}", __package__ + ".pages"), class_name)
    return page_class(*arguments)


def build_page_factories(backend, runtime_service) -> dict:
    dependencies = {"backend": backend, "runtime_service": runtime_service}
    return {
        key: partial(_create_page, module, class_name, tuple(dependencies[name] for name in names))
        for key, (module, class_name, names) in PAGE_DEFINITIONS.items()
    }
