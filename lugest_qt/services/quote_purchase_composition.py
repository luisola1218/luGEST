"""Bind purchase planning to explicit read projections and classification rules."""
from lugest_modules.quotes.application.purchase_needs import PurchaseNeeds, PurchaseNeedsRules
from lugest_modules.quotes.infrastructure.legacy_purchase_needs_repository import LegacyPurchaseNeedsRepository


def purchase_needs(backend) -> PurchaseNeeds:
    return PurchaseNeeds(
        LegacyPurchaseNeedsRepository(backend.ensure_data, lambda code: backend.material_by_id(code)),
        PurchaseNeedsRules(
            ORC_LINE_TYPE_PRODUCT=backend.desktop_main.ORC_LINE_TYPE_PRODUCT,
            ORC_LINE_TYPE_PIECE=backend.desktop_main.ORC_LINE_TYPE_PIECE,
            norm_text=lambda value: backend.desktop_main.norm_text(value),
            normalize_orc_line_type=lambda value: backend.desktop_main.normalize_orc_line_type(value),
            parse_float=backend._parse_float,
            is_raw_material=lambda line: backend._quote_line_is_raw_material(line),
            detect_materia_formato=lambda material: backend.desktop_main.detect_materia_formato(material),
        ),
    )
