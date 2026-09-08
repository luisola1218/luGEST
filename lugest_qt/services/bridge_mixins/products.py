from __future__ import annotations
from lugest_modules.inventory.application.product_definition import normalize_product, price_preview
from lugest_qt.services.inventory_composition import product_definition_rules, product_commands
from lugest_qt.services.inventory_composition import product_queries, stock_issue_service

from typing import Any


class ProductsBackendMixin:
    """Legacy adapter for products; see BACKEND_GUIDE.md."""

    def product_next_code(self) -> str:
        return str(self.desktop_main.peek_next_produto_numero(self.ensure_data()))

    def product_price_preview(self, payload: dict[str, Any]) -> dict[str, Any]:
        return price_preview(product_definition_rules(self), payload)

    def _product_dimensoes(self, prod: dict[str, Any]) -> str:
        return product_queries(self)._product_dimensoes(prod)

    def _product_normalize_payload(self, payload: dict[str, Any]) -> dict[str, Any]:
        return normalize_product(product_definition_rules(self), payload)

    def product_rows(self, filter_text: str = "", in_stock_only: bool = False) -> list[dict[str, Any]]:
        return product_queries(self).product_rows(filter_text, in_stock_only)

    def _product_issue_meta_from_obs(self, obs: str) -> dict[str, Any]:
        return product_queries(self)._product_issue_meta_from_obs(obs)

    def product_movement_years(self, codigo: str = "") -> list[str]:
        return product_queries(self).product_movement_years(codigo)

    def product_movements(
        self,
        codigo: str = "",
        limit: int = 120,
        operator_name: str = "",
        year: str = "",
        issue_only: bool = False,
    ) -> list[dict[str, Any]]:
        return product_queries(self).product_movements(codigo, limit, operator_name, year, issue_only)

    def product_issue_summary(self, operator_name: str = "", year: str = "", codigo: str = "") -> dict[str, Any]:
        return product_queries(self).product_issue_summary(operator_name, year, codigo)

    def product_detail(self, codigo: str) -> dict[str, Any]:
        return product_queries(self).product_detail(codigo)

    def product_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        code = product_commands(self).save(payload, actor=str((self.user or {}).get("username", "") or "Sistema"))
        try:
            self.conjunto_refresh_prices()
        except Exception:
            pass
        return self.product_detail(code)

    def product_remove(self, codigo: str) -> None:
        self.product_remove_many([codigo])

    def product_remove_many(self, codigos: list[str]) -> int:
        removed = product_commands(self).remove(codigos)
        try:
            self.ensure_inventory_scan_codes(persist=True)
        except Exception:
            pass
        try:
            self.conjunto_refresh_prices()
        except Exception:
            pass
        return removed

    def product_consume(
        self,
        codigo: str,
        quantidade: Any,
        obs: str = "",
        target_operator: str = "",
        issue_mode: str = "stock",
    ) -> dict[str, Any]:
        stock_issue_service(self).consume(
            codigo, quantidade, actor=str((self.user or {}).get("username", "") or "Sistema"),
            observation=obs, target_operator=target_operator, issue_mode=issue_mode,
        )
        return self.product_detail(codigo)
