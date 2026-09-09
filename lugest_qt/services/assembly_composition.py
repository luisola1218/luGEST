"""Bind assembly use cases to the current catalog and quote adapters."""
from copy import deepcopy

from lugest_modules.quotes.application.assemblies import AssemblyRules
from lugest_modules.quotes.application.assembly_refresh import AssemblyRefresh
from lugest_modules.quotes.application.assembly_catalog import AssemblyCatalog
from lugest_modules.quotes.application.assembly_queries import AssemblyQueries
from lugest_modules.quotes.infrastructure.legacy_assembly_repository import LegacyAssemblyRepository


def assembly_refresh(backend) -> AssemblyRefresh:
    return AssemblyRefresh(LegacyAssemblyRepository(backend.ensure_data, backend._save), assembly_rules(backend))


def assembly_catalog(backend, *, templates: bool = False) -> AssemblyCatalog:
    repository = LegacyAssemblyRepository(backend.ensure_data, backend._save,
                                          "conjuntos_modelo" if templates else "conjuntos")
    return AssemblyCatalog(repository, assembly_rules(backend), backend._next_assembly_model_code,
                           live_prices=not templates)


def assembly_queries(backend, *, templates: bool = False) -> AssemblyQueries:
    repository = LegacyAssemblyRepository(backend.ensure_data, backend._save,
                                          "conjuntos_modelo" if templates else "conjuntos")
    return AssemblyQueries(repository, assembly_rules(backend),
                           lambda item: backend.desktop_main.orc_line_is_service(item))


def assembly_rules(backend) -> AssemblyRules:
    return AssemblyRules(
        ORC_LINE_TYPE_PIECE=backend.desktop_main.ORC_LINE_TYPE_PIECE,
        ORC_LINE_TYPE_PRODUCT=backend.desktop_main.ORC_LINE_TYPE_PRODUCT,
        parse_float=backend._parse_float,
        normalize_orc_line_type=lambda value: backend.desktop_main.normalize_orc_line_type(value),
        product_lookup=lambda code: deepcopy(backend._product_lookup(code)),
        produto_preco_venda=lambda product: backend.desktop_main.produto_preco_venda(product),
        produto_preco_unitario=lambda product: backend.desktop_main.produto_preco_unitario(product),
        orc_line_is_product=lambda item: backend.desktop_main.orc_line_is_product(item),
        orc_line_is_piece=lambda item: backend.desktop_main.orc_line_is_piece(item),
        norm_text=lambda value: backend.desktop_main.norm_text(value),
        detect_materia_formato=lambda item: backend.desktop_main.detect_materia_formato(item),
        material_by_id=lambda code: deepcopy(backend.material_by_id(code)),
        material_candidates=lambda: deepcopy(list(backend.ensure_data().get("materiais", []) or [])),
        material_price_preview=lambda item: backend.material_price_preview(item),
        quote_source=lambda item, code: deepcopy(backend._conjunto_find_quote_source(item, code)),
        now_iso=lambda: backend.desktop_main.now_iso(),
    )


from lugest_modules.quotes.application.assembly_pair import AssemblyPair
from lugest_modules.quotes.infrastructure.legacy_assembly_pair_repository import LegacyAssemblyPairRepository


def assembly_pair(backend) -> AssemblyPair:
    return AssemblyPair(assembly_catalog(backend, templates=True), assembly_catalog(backend),
                        LegacyAssemblyPairRepository(backend.ensure_data, backend._save))
