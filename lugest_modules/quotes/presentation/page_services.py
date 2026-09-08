"""Capabilities of the quote controller; composition supplies implementations."""
from dataclasses import dataclass
from typing import Any, Callable
from lugest_modules.quotes.application.editor_ports import QuoteEditorPorts
from lugest_modules.quotes.presentation.profile_editor import ProfileEditorPorts


@dataclass(frozen=True)
class QuotePageServices:
    editor_ports: QuoteEditorPorts
    profile_editor_ports: ProfileEditorPorts
    ORC_LINE_TYPE_PIECE: str
    ORC_LINE_TYPE_PRODUCT: str
    ORC_LINE_TYPE_SERVICE: str
    _conjunto_find_quote_source: Callable[..., Any] | None
    assembly_model_detail: Callable[..., Any]
    assembly_model_expand: Callable[..., Any] | None
    assembly_model_remove: Callable[..., Any]
    assembly_model_rows: Callable[..., Any]
    assembly_model_save: Callable[..., Any]
    branding: Callable[..., Any]
    client_rows: Callable[..., Any]
    conjunto_detail: Callable[..., Any]
    conjunto_expand: Callable[..., Any]
    conjunto_next_param_codigo: Callable[..., Any]
    conjunto_open_sheet_pdf: Callable[..., Any]
    conjunto_remove: Callable[..., Any]
    conjunto_rows: Callable[..., Any]
    conjunto_save: Callable[..., Any]
    current_user: Callable[..., Any]
    laser_batch_dialog: Callable[..., Any]
    laser_dialog: Callable[..., Any]
    laser_quote_analyze: Callable[..., Any]
    laser_settings_dialog: Callable[..., Any]
    logo_path: Callable[..., Any] | None
    ne_suppliers: Callable[..., Any]
    nesting_dialog: Callable[..., Any]
    norm_text: Callable[..., Any]
    normalize_orc_line_type: Callable[..., Any]
    open_operation_profiles: Callable[..., Any]
    orc_available_years: Callable[..., Any]
    orc_clients: Callable[..., Any]
    orc_convert_to_order: Callable[..., Any]
    orc_create_purchase_quote: Callable[..., Any]
    orc_detail: Callable[..., Any]
    orc_line_is_piece: Callable[..., Any]
    orc_line_is_product: Callable[..., Any]
    orc_line_type_label: Callable[..., Any]
    orc_next_number: Callable[..., Any]
    orc_open_pdf: Callable[..., Any]
    orc_print_pdf: Callable[..., Any]
    orc_remove: Callable[..., Any]
    orc_render_pdf: Callable[..., Any]
    orc_rows: Callable[..., Any]
    orc_save: Callable[..., Any]
    orc_set_state: Callable[..., Any]
    orc_suggest_notes: Callable[..., Any]
    order_presets: Callable[..., Any]
    quote_authors: Callable[..., Any]
    quote_parse_operacoes_lista: Callable[..., Any]
    transport_zone_options: Callable[..., Any]
