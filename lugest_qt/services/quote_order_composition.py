"""Legacy identifier allocation and routing for quote-to-order preparation."""
from lugest_modules.quotes.application.order_lines import OrderLinePorts


def order_line_ports(backend, data, client_code: str) -> OrderLinePorts:
    return OrderLinePorts(
        parse_float=backend._parse_float,
        normalize_orc_line_type=backend.desktop_main.normalize_orc_line_type,
        normalize_operacao_nome=backend.desktop_main.normalize_operacao_nome,
        production_route=backend._quote_line_production_route,
        operations=backend._quote_line_operations_value,
        operations_text=backend._quote_line_operations_text,
        default_resource=backend.workcenter_default_resource,
        next_reference=lambda used: backend.desktop_main.next_ref_interna_unique(data, client_code, used),
        next_piece_order=lambda: backend.desktop_main.next_opp_numero(data),
        build_operacoes_fluxo=backend.desktop_main.build_operacoes_fluxo,
        now_iso=backend.desktop_main.now_iso,
    )
