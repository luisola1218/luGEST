"""Connect purchase use cases to historical persistence and value rules."""
from lugest_modules.purchasing.application.note_lifecycle import NoteLifecycle, NoteRules
from lugest_modules.purchasing.infrastructure.legacy_note_repository import LegacyNoteRepository


def note_lifecycle(backend) -> NoteLifecycle:
    return NoteLifecycle(
        LegacyNoteRepository(backend.ensure_data, backend._save, backend.desktop_main.next_ne_numero),
        NoteRules(backend.desktop_main.now_iso, backend._note_kind,
                  backend._resolve_supplier, backend._recalculate_note_totals),
    )


from dataclasses import replace
from copy import deepcopy
from lugest_modules.purchasing.application.note_commands import NoteCommands, PurchaseWriteRules
from lugest_modules.purchasing.application.pricing import PurchasePricing, PurchasePriceRules
from lugest_modules.purchasing.application.material_lines import MaterialLineRules, sync_material_lines
from lugest_modules.purchasing.infrastructure.legacy_purchase_repository import LegacyPurchaseRepository
from lugest_modules.quotes.application.assemblies import refresh_model
from lugest_qt.services.assembly_composition import assembly_rules


def material_line_rules(backend) -> MaterialLineRules:
    return MaterialLineRules(
        backend.desktop_main.origem_is_materia, backend.desktop_main.materia_preco_unitario,
        backend._parse_float, backend.desktop_main.detect_materia_formato,
        backend.ne_expedicao_actions._ne_build_material_desc,
    )

def purchase_pricing(backend) -> PurchasePricing:
    material_rules = material_line_rules(backend)

    def refresh(assemblies, products, materials):
        product_map = {str(row.get("codigo", "") or "").strip(): row for row in products}
        material_map = {str(row.get("id", "") or "").strip(): row for row in materials}
        rules = replace(assembly_rules(backend),
                        product_lookup=lambda code: deepcopy(product_map.get(str(code).strip())),
                        material_by_id=lambda code: deepcopy(material_map.get(str(code).strip())),
                        material_candidates=lambda: deepcopy(materials))
        return [refresh_model(rules, model)[0] if isinstance(model, dict) else model for model in assemblies]

    return PurchasePricing(PurchasePriceRules(
        backend._parse_float, backend.desktop_main.origem_is_materia,
        backend.desktop_main.produto_preco_unitario, backend.desktop_main.detect_materia_formato,
        backend.desktop_main.now_iso, backend.desktop_main.produto_modo_preco,
        lambda note, materials: sync_material_lines(note, materials, material_rules), refresh,
    ))


def note_commands(backend) -> NoteCommands:
    return NoteCommands(
        LegacyPurchaseRepository(backend.ensure_data, backend._save, backend.desktop_main.next_ne_numero),
        PurchaseWriteRules(backend._ne_normalize_line, backend._normalize_supplier_reference,
                           backend._resolve_supplier, backend.desktop_main.now_iso, backend._parse_float,
                           backend._note_kind, backend._recalculate_note_totals), purchase_pricing(backend),
    )


from lugest_modules.purchasing.application.line_normalization import PurchaseLines, PurchaseLineRules
from lugest_modules.purchasing.infrastructure.legacy_line_repository import LegacyLineRepository


def purchase_lines(backend) -> PurchaseLines:
    return PurchaseLines(LegacyLineRepository(backend.ensure_data), PurchaseLineRules(
        backend.desktop_main.norm_text, backend._parse_float, backend._parse_dimension_mm,
        backend._fmt, backend.material_geometry_preview, backend.desktop_main.origem_is_materia,
        backend.desktop_main.detect_materia_formato, backend._localizacao, backend.purchase_advice_for_line,
    ))
