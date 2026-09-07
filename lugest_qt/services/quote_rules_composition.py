"""Explicit legacy bindings for quote rules and document generation."""
from lugest_modules.quotes.application.line_normalization import LineNormalizationPorts

def normalize_line_ports(backend) -> LineNormalizationPorts:
    return LineNormalizationPorts(
        ORC_LINE_TYPE_PIECE=backend.desktop_main.ORC_LINE_TYPE_PIECE,
        ORC_LINE_TYPE_PRODUCT=backend.desktop_main.ORC_LINE_TYPE_PRODUCT,
        _normalize_conjunto_technical_sheet=backend._normalize_conjunto_technical_sheet,
        _parse_float=backend._parse_float,
        _product_lookup=backend._product_lookup,
        _quote_collect_non_laser_map=backend._quote_collect_non_laser_map,
        _quote_line_looks_stock_material_ref=backend._quote_line_looks_stock_material_ref,
        _quote_line_operation_snapshot=backend._quote_line_operation_snapshot,
        _quote_line_operations_text=backend._quote_line_operations_text,
        _quote_line_operations_value=backend._quote_line_operations_value,
        norm_text=backend.desktop_main.norm_text,
        normalize_operacao_nome=backend.desktop_main.normalize_operacao_nome,
        normalize_orc_line_type=backend.desktop_main.normalize_orc_line_type,
        produto_preco_venda=backend.desktop_main.produto_preco_venda,
    )

from lugest_modules.quotes.infrastructure.assembly_report import AssemblyReportPorts

def render_assembly_sheet_ports(backend) -> AssemblyReportPorts:
    return AssemblyReportPorts(
        _draw_code128_fit=backend._draw_code128_fit,
        _draw_operator_logo_plate=backend._draw_operator_logo_plate,
        _fmt=backend._fmt,
        _operator_label_palette=backend._operator_label_palette,
        _operator_pdf_text=backend._operator_pdf_text,
        _storage_output_path=backend._storage_output_path,
        branding_settings=backend.branding_settings,
        conjunto_detail=backend.conjunto_detail,
        now_iso=backend.desktop_main.now_iso,
        orc_line_is_piece=backend.desktop_main.orc_line_is_piece,
        orc_line_is_product=backend.desktop_main.orc_line_is_product,
        orc_line_is_service=backend.desktop_main.orc_line_is_service,
    )

from lugest_modules.quotes.infrastructure.nesting_report import NestingReportPorts

def render_nesting_study_ports(backend) -> NestingReportPorts:
    return NestingReportPorts(
        _draw_operator_logo_plate=backend._draw_operator_logo_plate,
        _fmt=backend._fmt,
        _fmt_eur=backend._fmt_eur,
        _operator_label_palette=backend._operator_label_palette,
        _parse_float=backend._parse_float,
        branding_settings=backend.branding_settings,
        orc_detail=backend.orc_detail,
        orc_nesting_studies=backend.orc_nesting_studies,
    )

