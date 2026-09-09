"""Compose a quote page without exposing the runtime to its controller."""
from copy import deepcopy
from lugest_qt.services.assembly_composition import assembly_pair
from lugest_modules.quotes.presentation.page_services import QuotePageServices
from lugest_qt.services.quote_editor_composition import quote_editor_ports, profile_editor_ports
from lugest_qt.ui.pages.laser_quote_dialogs import LaserQuoteDialog, LaserSettingsDialog
from lugest_qt.ui.pages.laser_batch_quote_dialog import LaserBatchQuoteDialog
from lugest_qt.ui.pages.laser_nesting_dialog import LaserNestingDialog
from lugest_qt.ui.pages.runtime_support import _open_operation_cost_profiles_dialog


def quote_page_services(backend) -> QuotePageServices:
    return QuotePageServices(
        save_assembly_pair=lambda template, live: assembly_pair(backend).save(template, live),
        editor_ports=quote_editor_ports(backend),
        profile_editor_ports=profile_editor_ports(backend),
        quote_authors=lambda: deepcopy(backend.ensure_data().get('orcamentistas', [])),
        current_user=lambda: deepcopy(getattr(backend, 'user', {}) or {}),
        branding=lambda: deepcopy(getattr(backend, 'branding', {}) or {}),
        open_operation_profiles=lambda owner: _open_operation_cost_profiles_dialog(owner, backend),
        laser_settings_dialog=lambda *args, **kwargs: LaserSettingsDialog(backend, *args, **kwargs),
        laser_batch_dialog=lambda *args, **kwargs: LaserBatchQuoteDialog(backend, *args, **kwargs),
        laser_dialog=lambda *args, **kwargs: LaserQuoteDialog(backend, *args, **kwargs),
        nesting_dialog=lambda *args, **kwargs: LaserNestingDialog(backend, *args, **kwargs),
        ORC_LINE_TYPE_PIECE=backend.desktop_main.ORC_LINE_TYPE_PIECE,
        ORC_LINE_TYPE_PRODUCT=backend.desktop_main.ORC_LINE_TYPE_PRODUCT,
        ORC_LINE_TYPE_SERVICE=backend.desktop_main.ORC_LINE_TYPE_SERVICE,
        _conjunto_find_quote_source=getattr(backend, '_conjunto_find_quote_source', None),
        assembly_model_detail=lambda *args, **kwargs: backend.assembly_model_detail(*args, **kwargs),
        assembly_model_expand=getattr(backend, 'assembly_model_expand', None),
        assembly_model_remove=lambda *args, **kwargs: backend.assembly_model_remove(*args, **kwargs),
        assembly_model_rows=lambda *args, **kwargs: backend.assembly_model_rows(*args, **kwargs),
        assembly_model_save=lambda *args, **kwargs: backend.assembly_model_save(*args, **kwargs),
        client_rows=lambda *args, **kwargs: backend.client_rows(*args, **kwargs),
        conjunto_detail=lambda *args, **kwargs: backend.conjunto_detail(*args, **kwargs),
        conjunto_expand=lambda *args, **kwargs: backend.conjunto_expand(*args, **kwargs),
        conjunto_next_param_codigo=lambda *args, **kwargs: backend.conjunto_next_param_codigo(*args, **kwargs),
        conjunto_open_sheet_pdf=lambda *args, **kwargs: backend.conjunto_open_sheet_pdf(*args, **kwargs),
        conjunto_remove=lambda *args, **kwargs: backend.conjunto_remove(*args, **kwargs),
        conjunto_rows=lambda *args, **kwargs: backend.conjunto_rows(*args, **kwargs),
        conjunto_save=lambda *args, **kwargs: backend.conjunto_save(*args, **kwargs),
        laser_quote_analyze=lambda *args, **kwargs: backend.laser_quote_analyze(*args, **kwargs),
        logo_path=getattr(backend, 'logo_path', None),
        ne_suppliers=lambda *args, **kwargs: backend.ne_suppliers(*args, **kwargs),
        norm_text=lambda *args, **kwargs: backend.desktop_main.norm_text(*args, **kwargs),
        normalize_orc_line_type=lambda *args, **kwargs: backend.desktop_main.normalize_orc_line_type(*args, **kwargs),
        orc_available_years=lambda *args, **kwargs: backend.orc_available_years(*args, **kwargs),
        orc_clients=lambda *args, **kwargs: backend.orc_clients(*args, **kwargs),
        orc_convert_to_order=lambda *args, **kwargs: backend.orc_convert_to_order(*args, **kwargs),
        orc_create_purchase_quote=lambda *args, **kwargs: backend.orc_create_purchase_quote(*args, **kwargs),
        orc_detail=lambda *args, **kwargs: backend.orc_detail(*args, **kwargs),
        orc_line_is_piece=lambda *args, **kwargs: backend.desktop_main.orc_line_is_piece(*args, **kwargs),
        orc_line_is_product=lambda *args, **kwargs: backend.desktop_main.orc_line_is_product(*args, **kwargs),
        orc_line_type_label=lambda *args, **kwargs: backend.desktop_main.orc_line_type_label(*args, **kwargs),
        orc_next_number=lambda *args, **kwargs: backend.orc_next_number(*args, **kwargs),
        orc_open_pdf=lambda *args, **kwargs: backend.orc_open_pdf(*args, **kwargs),
        orc_print_pdf=lambda *args, **kwargs: backend.orc_print_pdf(*args, **kwargs),
        orc_remove=lambda *args, **kwargs: backend.orc_remove(*args, **kwargs),
        orc_render_pdf=lambda *args, **kwargs: backend.orc_render_pdf(*args, **kwargs),
        orc_rows=lambda *args, **kwargs: backend.orc_rows(*args, **kwargs),
        orc_save=lambda *args, **kwargs: backend.orc_save(*args, **kwargs),
        orc_set_state=lambda *args, **kwargs: backend.orc_set_state(*args, **kwargs),
        orc_suggest_notes=lambda *args, **kwargs: backend.orc_suggest_notes(*args, **kwargs),
        order_presets=lambda *args, **kwargs: backend.order_presets(*args, **kwargs),
        quote_parse_operacoes_lista=lambda *args, **kwargs: backend.quote_parse_operacoes_lista(*args, **kwargs),
        transport_zone_options=lambda *args, **kwargs: backend.transport_zone_options(*args, **kwargs),
    )
