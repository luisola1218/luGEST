"""Bind inventory read use cases to the existing runtime at the application edge."""
from lugest_modules.inventory.application.product_queries import ProductQueries, ProductQueryRules
from lugest_modules.inventory.infrastructure.legacy_product_repository import LegacyProductReadRepository
from lugest_modules.inventory.application.consumption import StockIssueService
from lugest_modules.inventory.infrastructure.legacy_stock_issue_repository import LegacyStockIssueRepository
from lugest_modules.inventory.application.product_definition import ProductDefinitionRules, normalize_product
from lugest_modules.inventory.application.product_commands import ProductCommands
from lugest_modules.inventory.infrastructure.legacy_product_write_repository import LegacyProductWriteRepository


def product_definition_rules(backend) -> ProductDefinitionRules:
    return ProductDefinitionRules(
        unit_price=backend.desktop_main.produto_preco_unitario,
        _parse_float=backend._parse_float, _product_dimensoes=backend._product_dimensoes,
        _product_resolve_catalog_fields=backend._product_resolve_catalog_fields,
        product_copilot_analysis=backend.product_copilot_analysis,
        product_next_code=backend.product_next_code,
    )


def product_commands(backend) -> ProductCommands:
    repository = LegacyProductWriteRepository(
        get_data=backend.ensure_data, save_dataset=backend._save,
        add_movement=backend.desktop_main.add_produto_mov,
        ensure_sequence=backend.desktop_main.ensure_produto_seq,
    )
    return ProductCommands(repository, normalize=lambda payload: normalize_product(product_definition_rules(backend), payload),
                           parse_float=backend._parse_float, format_number=backend._fmt,
                           now=backend.desktop_main.now_iso)


def stock_issue_service(backend) -> StockIssueService:
    repository = LegacyStockIssueRepository(
        get_data=backend.ensure_data, save_dataset=backend._save,
        add_movement=backend.desktop_main.add_produto_mov, parse_float=backend._parse_float,
    )
    return StockIssueService(repository, parse_float=backend._parse_float,
                             unit_price=backend.desktop_main.produto_preco_unitario,
                             now=backend.desktop_main.now_iso)

def product_queries(backend) -> ProductQueries:
    repository = LegacyProductReadRepository(
        products=lambda: backend.ensure_data().get('produtos', []),
        movements=lambda: backend.ensure_data().get('produtos_mov', []),
    )
    rules = ProductQueryRules(
        _fmt=backend._fmt,
        _parse_float=backend._parse_float,
        _product_resolve_catalog_fields=backend._product_resolve_catalog_fields,
        _quality_status_is_available=backend._quality_status_is_available,
        inventory_scan_code=backend.inventory_scan_code,
        produto_preco_unitario=backend.desktop_main.produto_preco_unitario,
        produto_preco_venda=backend.desktop_main.produto_preco_venda,
    )
    return ProductQueries(repository, rules)
