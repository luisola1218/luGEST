"""Composition for quote persistence and reference operations."""
from datetime import datetime
from lugest_modules.quotes.application.commands import QuoteCommands, QuoteCommandRules
from lugest_modules.quotes.infrastructure.legacy_quote_repository import LegacyQuoteWriteRepository


def quote_commands(backend) -> QuoteCommands:
    repository = LegacyQuoteWriteRepository(
        get_data=backend.ensure_data, get_baseline=lambda: backend._base_data_snapshot,
        save_dataset=backend._save,
        next_number=lambda: backend.desktop_main.next_orc_numero(backend.ensure_data()),
        next_reference=lambda client, reserved: backend.desktop_main.next_ref_interna_unique(backend.ensure_data(), client, reserved),
        peek_number=backend._peek_next_orc_number,
        sync_registry=backend._sync_quote_piece_registry,
        upsert=getattr(backend.desktop_main, 'mysql_upsert_orcamento_com_linhas', None),
        delete_studies=backend._mysql_delete_orc_nesting_studies,
    )
    rules = QuoteCommandRules(
        _active_client_ref_usage=lambda *args, **kwargs: backend._active_client_ref_usage(*args, **kwargs),
        _known_client_ref_for_external=lambda *args, **kwargs: backend._known_client_ref_for_external(*args, **kwargs),
        _known_client_ref_pairs=lambda *args, **kwargs: backend._known_client_ref_pairs(*args, **kwargs),
        _normalize_orc_client=lambda *args, **kwargs: backend._normalize_orc_client(*args, **kwargs),
        _normalize_orc_line=lambda *args, **kwargs: backend._normalize_orc_line(*args, **kwargs),
        _normalize_quote_discount_groups=lambda *args, **kwargs: backend._normalize_quote_discount_groups(*args, **kwargs),
        _normalize_quote_discount_mode=lambda *args, **kwargs: backend._normalize_quote_discount_mode(*args, **kwargs),
        _normalize_supplier_reference=lambda *args, **kwargs: backend._normalize_supplier_reference(*args, **kwargs),
        _normalize_workcenter_value=lambda *args, **kwargs: backend._normalize_workcenter_value(*args, **kwargs),
        _parse_float=lambda *args, **kwargs: backend._parse_float(*args, **kwargs),
        _quote_default_delivery_text=lambda *args, **kwargs: backend._quote_default_delivery_text(*args, **kwargs),
        _quote_line_is_raw_material=lambda *args, **kwargs: backend._quote_line_is_raw_material(*args, **kwargs),
        _quote_standard_iva_perc=lambda *args, **kwargs: backend._quote_standard_iva_perc(*args, **kwargs),
        _ref_client_code=lambda *args, **kwargs: backend._ref_client_code(*args, **kwargs),
        _repair_orc_ref_history=lambda *args, **kwargs: backend._repair_orc_ref_history(*args, **kwargs),
        current_year=lambda: datetime.now().year,
        now_iso=lambda *args, **kwargs: backend.desktop_main.now_iso(*args, **kwargs),
        orc_line_is_piece=lambda *args, **kwargs: backend.desktop_main.orc_line_is_piece(*args, **kwargs),
    )
    return QuoteCommands(repository, rules)
