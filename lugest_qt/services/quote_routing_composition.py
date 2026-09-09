"""Bind normalized business values to the quote routing rules."""
from lugest_modules.quotes.domain.routing import LineRouting, RoutingRules


def line_routing(backend) -> LineRouting:
    return LineRouting(RoutingRules(
        backend.desktop_main.ORC_LINE_TYPE_PIECE,
        backend.desktop_main.ORC_LINE_TYPE_PRODUCT,
        backend.desktop_main.ORC_LINE_TYPE_SERVICE,
        backend.desktop_main.normalize_orc_line_type,
        backend.desktop_main.normalize_operacao_nome,
        backend.desktop_main.norm_text,
        backend._parse_float,
        backend.quote_parse_operacoes_lista,
        backend.quote_format_operacoes,
    ))
