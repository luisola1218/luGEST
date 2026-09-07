"""Explicit capabilities consumed by quote editors; no page or runtime state."""
from dataclasses import dataclass
from typing import Any, Callable

@dataclass(frozen=True)
class QuoteEditorPorts:
    ORC_LINE_TYPE_PIECE: str
    ORC_LINE_TYPE_PRODUCT: str
    ORC_LINE_TYPE_SERVICE: str
    _fmt: Callable[..., Any]
    _parse_float: Callable[..., Any]
    add_material: Callable[..., Any]
    build_operacoes_fluxo: Callable[..., Any]
    detect_materia_formato: Callable[..., Any]
    material_by_id: Callable[..., Any]
    material_default_price_kg: Callable[..., Any]
    material_family_options: Callable[..., Any]
    material_family_profile: Callable[..., Any]
    material_geometry_preview: Callable[..., Any]
    material_price_preview: Callable[..., Any]
    material_price_rows: Callable[..., Any]
    material_profile_size_options: Callable[..., Any]
    material_rows: Callable[..., Any]
    material_section_options: Callable[..., Any]
    material_update_price_kg: Callable[..., Any]
    ne_product_options: Callable[..., Any]
    norm_text: Callable[..., Any]
    normalize_operacao_nome: Callable[..., Any]
    normalize_orc_line_type: Callable[..., Any]
    operation_cost_estimate: Callable[..., Any]
    orc_suggest_ref_interna: Callable[..., Any]
    order_presets: Callable[..., Any]
    order_reference_rows: Callable[..., Any]
    product_catalog_options: Callable[..., Any]
    product_next_code: Callable[..., Any]
    product_save: Callable[..., Any]
    quote_parse_operacoes_lista: Callable[..., Any]
