"""Bindings for quote read models; the business module has no runtime imports."""
from datetime import datetime
from lugest_modules.quotes.application.queries import QuoteQueries, QuoteQueryRules
from lugest_modules.quotes.infrastructure.legacy_quote_read_repository import LegacyQuoteReadRepository


def quote_queries(backend) -> QuoteQueries:
    repository = LegacyQuoteReadRepository(lambda: backend.ensure_data().get('orcamentos', []))
    rules = QuoteQueryRules(
        _fmt=lambda *args, **kwargs: backend._fmt(*args, **kwargs),
        _normalize_conjunto_technical_sheet=lambda *args, **kwargs: backend._normalize_conjunto_technical_sheet(*args, **kwargs),
        _normalize_orc_client=lambda *args, **kwargs: backend._normalize_orc_client(*args, **kwargs),
        _normalize_quote_discount_groups=lambda *args, **kwargs: backend._normalize_quote_discount_groups(*args, **kwargs),
        _normalize_quote_discount_mode=lambda *args, **kwargs: backend._normalize_quote_discount_mode(*args, **kwargs),
        _normalize_workcenter_value=lambda *args, **kwargs: backend._normalize_workcenter_value(*args, **kwargs),
        _orc_number_sort_key=lambda *args, **kwargs: backend._orc_number_sort_key(*args, **kwargs),
        _parse_float=lambda *args, **kwargs: backend._parse_float(*args, **kwargs),
        _quote_collect_non_laser_map=lambda *args, **kwargs: backend._quote_collect_non_laser_map(*args, **kwargs),
        _quote_default_delivery_text=lambda *args, **kwargs: backend._quote_default_delivery_text(*args, **kwargs),
        _quote_line_operation_snapshot=lambda *args, **kwargs: backend._quote_line_operation_snapshot(*args, **kwargs),
        _quote_standard_iva_perc=lambda *args, **kwargs: backend._quote_standard_iva_perc(*args, **kwargs),
        current_year=lambda: datetime.now().year,
        extract_year=lambda *args, **kwargs: backend.orc_actions._orc_extract_year(*args, **kwargs),
        norm_text=lambda *args, **kwargs: backend.desktop_main.norm_text(*args, **kwargs),
        normalize_orc_line_type=lambda *args, **kwargs: backend.desktop_main.normalize_orc_line_type(*args, **kwargs),
        orc_line_is_piece=lambda *args, **kwargs: backend.desktop_main.orc_line_is_piece(*args, **kwargs),
    )
    return QuoteQueries(repository, rules)
