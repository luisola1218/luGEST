"""Legacy identifier allocation and routing for quote-to-order preparation."""
from lugest_modules.quotes.application.order_lines import OrderLinePorts
from lugest_modules.quotes.application.conversion import QuoteConversion, ConversionRules
from lugest_modules.quotes.infrastructure.legacy_conversion_repository import LegacyConversionRepository


def quote_conversion(backend) -> QuoteConversion:
    def order_code(candidate, order):
        code = backend._order_expected_of_code(order) or str(backend.desktop_main.next_of_numero(candidate) or "").strip()
        if code:
            order["of_codigo"] = code
            order["ordem_fabrico"] = {
                "id": code, "encomenda_id": str(order.get("numero", "") or "").strip(),
                "estado": str(order.get("estado", "") or "Preparacao").strip() or "Preparacao",
                "data": str(order.get("data_criacao", "") or backend.desktop_main.now_iso()).strip()[:10],
            }
        return code

    repository = LegacyConversionRepository(
        backend.ensure_data, backend._save,
        next_client=backend.desktop_main.next_cliente_codigo,
        next_order=backend.desktop_main.next_encomenda_numero,
        order_code=order_code,
        make_line_ports=lambda data, client: order_line_ports(backend, data, client),
        register_reference=backend.desktop_main.update_refs,
    )
    return QuoteConversion(repository, ConversionRules(
        normalize_client=backend._normalize_orc_client,
        now_iso=backend.desktop_main.now_iso,
        parse_float=backend._parse_float,
        normalize_workcenter=backend._normalize_workcenter_value,
        assembly_detail=backend.conjunto_detail,
        update_order_state=backend.desktop_main.update_estado_encomenda_por_espessuras,
    ))


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
