from lugest_modules.transport.infrastructure.route_report import RouteReport, RouteReportRules
"""Explicit adapters for the transport module."""
from lugest_modules.transport.application.tariffs import TariffService, TariffRules
from lugest_modules.transport.infrastructure.legacy_tariff_repository import LegacyTariffRepository
from lugest_modules.transport.application.stops import TransportStops, StopRules
from lugest_modules.transport.infrastructure.legacy_trip_repository import LegacyTripRepository
from lugest_modules.transport.application.trips import TripCommands, TripRules
from lugest_modules.transport.application.assignments import TripAssignments, AssignmentRules
from lugest_modules.transport.application.queries import TransportQueries, TransportQueryRules
from lugest_modules.transport.infrastructure.legacy_transport_read_repository import LegacyTransportReadRepository


def transport_queries(backend) -> TransportQueries:
    return TransportQueries(
        LegacyTransportReadRepository(backend.ensure_data),
        TransportQueryRules(backend._parse_float, lambda value: backend.desktop_main.norm_text(value),
                            lambda *args: backend._normalize_supplier_reference(*args),
                            backend._transport_tariff_suggestion,
                            lambda order: backend.desktop_main.update_estado_expedicao_encomenda(order),
                            lambda order: backend.desktop_main.encomenda_pecas(order),
                            lambda piece: backend.desktop_main.peca_qtd_disponivel_expedicao(piece),
                            backend._transport_stop_checklist_state, backend._transport_defaults),
    )


def trip_assignments(backend) -> TripAssignments:
    normalize = lambda value: backend.desktop_main.norm_text(value)
    return TripAssignments(
        LegacyTripRepository(backend.ensure_data, backend._save, normalize),
        AssignmentRules(backend._parse_float, normalize, lambda: backend.desktop_main.now_iso(),
                        backend._transport_is_own_cargo, backend._transport_latest_guide_for_order,
                        backend._transport_metrics_for_order, backend._transport_tariff_suggestion,
                        backend._transport_reindex_stops, backend.transport_detail),
    )


def trip_commands(backend) -> TripCommands:
    normalize = lambda value: backend.desktop_main.norm_text(value)

    def allocate(candidate, requested):
        number = str(requested or backend.desktop_main.next_transporte_numero(candidate)).strip()
        backend.desktop_main.reserve_transporte_numero(candidate, number)
        return number

    return TripCommands(
        LegacyTripRepository(backend.ensure_data, backend._save, normalize, allocate),
        TripRules(backend._parse_float, normalize, lambda: backend.desktop_main.now_iso(),
                  lambda: str((backend.user or {}).get("username", "") or "").strip(),
                  lambda *args: backend._normalize_supplier_reference(*args),
                  backend._transport_defaults, backend._transport_reindex_stops),
    )


def transport_stops(backend) -> TransportStops:
    normalize = lambda value: backend.desktop_main.norm_text(value)
    return TransportStops(
        LegacyTripRepository(backend.ensure_data, backend._save, normalize),
        StopRules(backend._parse_float, normalize, lambda: backend.desktop_main.now_iso(),
                  lambda: str((backend.user or {}).get("username", "") or "").strip(),
                  lambda *args: backend._normalize_supplier_reference(*args),
                  lambda number: backend.transport_guide_options(number)),
    )


def tariff_service(backend) -> TariffService:
    return TariffService(
        LegacyTariffRepository(backend.ensure_data, backend._save),
        TariffRules(backend._parse_float, lambda value: backend.desktop_main.norm_text(value),
                    lambda *args: backend._normalize_supplier_reference(*args)),
    )


def route_report(backend) -> RouteReport:
    return RouteReport(RouteReportRules(
        backend._operator_label_palette, backend.branding_settings,
        lambda: backend.desktop_main.now_iso(), backend._operator_pdf_text,
        backend._draw_operator_logo_plate, backend._fmt, backend._fmt_eur,
    ))
