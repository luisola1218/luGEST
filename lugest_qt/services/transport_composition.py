"""Explicit adapters for the transport module."""
from lugest_modules.transport.application.tariffs import TariffService, TariffRules
from lugest_modules.transport.infrastructure.legacy_tariff_repository import LegacyTariffRepository
from lugest_modules.transport.application.stops import TransportStops, StopRules
from lugest_modules.transport.infrastructure.legacy_trip_repository import LegacyTripRepository


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
