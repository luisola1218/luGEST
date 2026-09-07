"""Legacy bindings for quote editors. Only this adapter knows desktop_main."""
from lugest_modules.quotes.application.editor_ports import QuoteEditorPorts
from lugest_modules.quotes.presentation.profile_editor import ProfileEditorPorts


def profile_editor_ports(backend) -> ProfileEditorPorts:
    return ProfileEditorPorts(
        ORC_LINE_TYPE_SERVICE=backend.desktop_main.ORC_LINE_TYPE_SERVICE,
        laser_quote_settings=backend.laser_quote_settings,
        laser_quote_save_settings=backend.laser_quote_save_settings,
        material_presets=backend.material_presets,
        profile_laser_quote_analyze=backend.profile_laser_quote_analyze,
        profile_laser_quote_build_line=backend.profile_laser_quote_build_line,
    )

def quote_editor_ports(backend) -> QuoteEditorPorts:
    return QuoteEditorPorts(
        ORC_LINE_TYPE_PIECE=backend.desktop_main.ORC_LINE_TYPE_PIECE,
        ORC_LINE_TYPE_PRODUCT=backend.desktop_main.ORC_LINE_TYPE_PRODUCT,
        ORC_LINE_TYPE_SERVICE=backend.desktop_main.ORC_LINE_TYPE_SERVICE,
        _fmt=backend._fmt,
        _parse_float=backend._parse_float,
        add_material=backend.add_material,
        build_operacoes_fluxo=backend.desktop_main.build_operacoes_fluxo,
        detect_materia_formato=backend.desktop_main.detect_materia_formato,
        material_by_id=backend.material_by_id,
        material_default_price_kg=backend.material_default_price_kg,
        material_family_options=backend.material_family_options,
        material_family_profile=backend.material_family_profile,
        material_geometry_preview=backend.material_geometry_preview,
        material_price_preview=backend.material_price_preview,
        material_price_rows=backend.material_price_rows,
        material_profile_size_options=backend.material_profile_size_options,
        material_rows=backend.material_rows,
        material_section_options=backend.material_section_options,
        material_update_price_kg=backend.material_update_price_kg,
        ne_product_options=backend.ne_product_options,
        norm_text=backend.desktop_main.norm_text,
        normalize_operacao_nome=backend.desktop_main.normalize_operacao_nome,
        normalize_orc_line_type=backend.desktop_main.normalize_orc_line_type,
        operation_cost_estimate=backend.operation_cost_estimate,
        orc_suggest_ref_interna=backend.orc_suggest_ref_interna,
        order_presets=backend.order_presets,
        order_reference_rows=backend.order_reference_rows,
        product_catalog_options=backend.product_catalog_options,
        product_next_code=backend.product_next_code,
        product_save=backend.product_save,
        quote_parse_operacoes_lista=backend.quote_parse_operacoes_lista,
    )
