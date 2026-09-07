from __future__ import annotations

import os
from lugest_infra.app_paths import AppPaths
from lugest_qt.services.bridge_mixins import (
    BillingBridgeMixin,
    DashboardBridgeMixin,
    DirectServicesBridgeMixin,
    PlanningBridgeMixin,
    PurchasingBridgeMixin,
    QuotesBridgeMixin,
    ShippingBridgeMixin,
    TransportBridgeMixin,
)
from lugest_qt.services.bridge_mixins.updates import UpdatesBridgeMixin
from lugest_qt.services.legacy_runtime import LegacyRuntime, load_legacy_runtime
from pathlib import Path
from typing import Any
from .bridge_helpers import (
    _ANGLE_SECTION_OPTIONS,
    _BAR_SECTION_OPTIONS,
    _CHECKER_PLATE_WEIGHT_KG,
    _PROFILE_SECTION_OPTIONS,
    _PROFILE_STANDARD_KG_M,
    _STEEL_DENSITY_G_CM3,
    _TUBE_SECTION_OPTIONS,
    _ValueHolder,
    _checker_plate_weight_kg,
    _commercial_thickness_base_mm,
    _commercial_thickness_label,
    _detect_profile_catalog_from_text,
    _profile_catalog_lookup_key,
    _profile_size_lookup_key,
    _search_normalize,
    _search_terms,
)
from .bridge_mixins.assembly_stock import AssemblyStockBackendMixin
from .bridge_mixins.branding import BrandingBackendMixin
from .bridge_mixins.clients import ClientsBackendMixin
from .bridge_mixins.common import CommonBackendMixin
from .bridge_mixins.configuration import ConfigurationBackendMixin
from .bridge_mixins.data_runtime import DataRuntimeBackendMixin
from .bridge_mixins.diagnostics import DiagnosticsBackendMixin
from .bridge_mixins.documents import DocumentsBackendMixin
from .bridge_mixins.intelligence import IntelligenceBackendMixin
from .bridge_mixins.inventory_scanning import InventoryScanningBackendMixin
from .bridge_mixins.material_assistant import MaterialAssistantBackendMixin
from .bridge_mixins.material_assistant_reports import MaterialAssistantReportsBackendMixin
from .bridge_mixins.material_reports import MaterialReportsBackendMixin
from .bridge_mixins.materials import MaterialsBackendMixin
from .bridge_mixins.operation_catalog import OperationCatalogBackendMixin
from .bridge_mixins.operation_costing import OperationCostingBackendMixin
from .bridge_mixins.operator import OperatorBackendMixin
from .bridge_mixins.operator_labels import OperatorLabelsBackendMixin
from .bridge_mixins.order_reports import OrderReportsBackendMixin
from .bridge_mixins.orders import OrdersBackendMixin
from .bridge_mixins.planning_delays import PlanningDelaysBackendMixin
from .bridge_mixins.planning_operations import PlanningOperationsBackendMixin
from .bridge_mixins.product_catalog import ProductCatalogBackendMixin
from .bridge_mixins.product_reports import ProductReportsBackendMixin
from .bridge_mixins.production_orders import ProductionOrdersBackendMixin
from .bridge_mixins.products import ProductsBackendMixin
from .bridge_mixins.purchasing_documents import PurchasingDocumentsBackendMixin
from .bridge_mixins.quality import QualityBackendMixin
from .bridge_mixins.quality_reports import QualityReportsBackendMixin
from .bridge_mixins.quote_references import QuoteReferencesBackendMixin
from .bridge_mixins.users import UsersBackendMixin
from .bridge_mixins.workcenters import WorkcentersBackendMixin


class LegacyBackend(
    AssemblyStockBackendMixin,
    BrandingBackendMixin,
    ClientsBackendMixin,
    CommonBackendMixin,
    ConfigurationBackendMixin,
    DataRuntimeBackendMixin,
    DiagnosticsBackendMixin,
    DocumentsBackendMixin,
    IntelligenceBackendMixin,
    InventoryScanningBackendMixin,
    MaterialAssistantBackendMixin,
    MaterialAssistantReportsBackendMixin,
    MaterialReportsBackendMixin,
    MaterialsBackendMixin,
    OperationCatalogBackendMixin,
    OperationCostingBackendMixin,
    OperatorBackendMixin,
    OperatorLabelsBackendMixin,
    OrderReportsBackendMixin,
    OrdersBackendMixin,
    PlanningDelaysBackendMixin,
    PlanningOperationsBackendMixin,
    ProductCatalogBackendMixin,
    ProductReportsBackendMixin,
    ProductionOrdersBackendMixin,
    ProductsBackendMixin,
    PurchasingDocumentsBackendMixin,
    QualityBackendMixin,
    QualityReportsBackendMixin,
    QuoteReferencesBackendMixin,
    UsersBackendMixin,
    WorkcentersBackendMixin,
    UpdatesBridgeMixin,
    DirectServicesBridgeMixin,
    BillingBridgeMixin,
    PurchasingBridgeMixin,
    QuotesBridgeMixin,
    PlanningBridgeMixin,
    ShippingBridgeMixin,
    TransportBridgeMixin,
    DashboardBridgeMixin,
):
    """Composition root for legacy adapters. Public callers use legacy_backend.py."""

    def __init__(self, *, runtime: LegacyRuntime | None = None) -> None:
        legacy = runtime if runtime is not None else load_legacy_runtime()
        self.app_misc_actions = legacy.app_misc_actions
        self.billing_pdf_actions = legacy.billing_pdf_actions
        self.encomendas_actions = legacy.encomendas_actions
        self.desktop_main = legacy.desktop_main
        self.materia_actions = legacy.materia_actions
        self.ne_expedicao_actions = legacy.ne_expedicao_actions
        self.orc_actions = legacy.orc_actions
        self.operador_actions = legacy.operador_ordens_actions
        self.plan_actions = legacy.plan_actions
        self.produtos_actions = legacy.produtos_actions
        self.tax_compliance = legacy.tax_compliance
        desktop_main = legacy.desktop_main
        self.base_dir = Path(getattr(desktop_main, "BASE_DIR", Path.cwd()))
        self.app_paths = AppPaths(self.base_dir)
        self.data: dict[str, Any] | None = None
        self._base_data_snapshot: dict[str, Any] | None = None
        self._data_loaded_at = 0.0
        self._data_cache_generation = 0
        self._reload_cache_ttl_sec = self._env_float("LUGEST_RELOAD_CACHE_TTL_SEC", 30.0, minimum=0.0, maximum=300.0)
        self._op_mysql_ops_status_cache: dict[tuple[str, str], tuple[float, list[dict[str, Any]]]] = {}
        self._op_mysql_ops_status_ttl_sec = self._env_float("LUGEST_OPERATOR_OPS_STATUS_TTL_SEC", 2.0, minimum=0.0, maximum=30.0)
        self._trial_status_cache: dict[str, Any] | None = None
        self._trial_status_loaded_at = 0.0
        self._trial_status_cache_ttl_sec = self._env_float(
            "LUGEST_TRIAL_STATUS_CACHE_TTL_SEC",
            60.0,
            minimum=5.0,
            maximum=600.0,
        )
        self.user: dict[str, Any] | None = None
        self._qt_config_cache: dict[str, Any] | None = None
        self._qt_config_last_error = ""
        self._product_taxonomy_nodes_cache: tuple[
            dict[str, Any],
            dict[tuple[str, str], dict[str, Any]],
            dict[tuple[str, str, str], dict[str, Any]],
        ] | None = None
        self._operation_catalog_cache: tuple[int, list[dict[str, Any]]] | None = None

    def _env_float(self, name: str, default: float, *, minimum: float | None = None, maximum: float | None = None) -> float:
        try:
            value = float(os.environ.get(name, "") or default)
        except Exception:
            value = float(default)
        if minimum is not None:
            value = max(float(minimum), value)
        if maximum is not None:
            value = min(float(maximum), value)
        return value
