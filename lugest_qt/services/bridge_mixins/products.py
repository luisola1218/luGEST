from __future__ import annotations
from lugest_qt.services.inventory_composition import product_queries, stock_issue_service

from datetime import date
from typing import Any


class ProductsBackendMixin:
    """Legacy adapter for products; see BACKEND_GUIDE.md."""

    def product_next_code(self) -> str:
        return str(self.desktop_main.peek_next_produto_numero(self.ensure_data()))

    def product_price_preview(self, payload: dict[str, Any]) -> dict[str, Any]:
        """Calcula os indicadores editáveis sem executar classificação inteligente."""
        product = dict(payload or {})
        quantity = self._parse_float(product.get("qty", product.get("quantidade", 0)), 0)
        unit = str(product.get("unid", "UN") or "UN").strip() or "UN"
        unit_price = self._parse_float(self.desktop_main.produto_preco_unitario(product), 0)
        return {
            "preco_unid": round(unit_price, 4),
            "qty": quantity,
            "unit": unit,
            "valor_stock": round(unit_price * quantity, 2),
        }

    def _product_dimensoes(self, prod: dict[str, Any]) -> str:
        return product_queries(self)._product_dimensoes(prod)

    def _product_normalize_payload(self, payload: dict[str, Any]) -> dict[str, Any]:
        code = str(payload.get("codigo", "") or "").strip() or self.product_next_code()
        descricao = str(payload.get("descricao", "") or "").strip()
        if not code:
            raise ValueError("Codigo do produto em falta.")
        if not descricao:
            raise ValueError("Descricao do produto em falta.")
        normalized_payload = dict(payload)
        intelligence = self.product_copilot_analysis(descricao, code)
        suggestion = intelligence
        for field in ("categoria", "subcat", "tipo", "dimensoes"):
            if not str(normalized_payload.get(field, "") or "").strip():
                suggested_value = str(suggestion.get(field, "") or "").strip()
                if suggested_value:
                    normalized_payload[field] = suggested_value
        catalog_fields = self._product_resolve_catalog_fields(normalized_payload)
        categoria = str(catalog_fields.get("categoria", "") or "").strip()
        tipo = str(catalog_fields.get("tipo", "") or "").strip()
        metros_unidade = self._parse_float(payload.get("metros_unidade", payload.get("metros", 0)), 0)
        prod = {
            "codigo": code,
            "descricao": descricao,
            "categoria": categoria,
            "category_id": str(catalog_fields.get("category_id", "") or "").strip(),
            "subcat": str(catalog_fields.get("subcat", "") or "").strip(),
            "subcategory_id": str(catalog_fields.get("subcategory_id", "") or "").strip(),
            "tipo": tipo,
            "type_id": str(catalog_fields.get("type_id", "") or "").strip(),
            "category_icon": str(catalog_fields.get("category_icon", "") or "").strip(),
            "category_badge": str(catalog_fields.get("category_badge", "") or "").strip(),
            "category_tone": str(catalog_fields.get("category_tone", "") or "").strip(),
            "dimensoes": str(normalized_payload.get("dimensoes", "") or "").strip(),
            "comprimento": self._parse_float(payload.get("comprimento", 0), 0),
            "largura": self._parse_float(payload.get("largura", 0), 0),
            "espessura": self._parse_float(payload.get("espessura", 0), 0),
            "metros_unidade": metros_unidade,
            "metros": metros_unidade,
            "peso_unid": self._parse_float(payload.get("peso_unid", 0), 0),
            "fabricante": str(payload.get("fabricante", "") or "").strip(),
            "modelo": str(payload.get("modelo", "") or "").strip(),
            "unid": str(payload.get("unid", "UN") or "UN").strip() or "UN",
            "qty": self._parse_float(payload.get("qty", payload.get("quantidade", 0)), 0),
            "alerta": self._parse_float(payload.get("alerta", 0), 0),
            "p_compra": self._parse_float(payload.get("p_compra", 0), 0),
            "pvp1": self._parse_float(payload.get("pvp1", 0), 0),
            "pvp2": self._parse_float(payload.get("pvp2", 0), 0),
            "obs": str(payload.get("obs", "") or "").strip(),
            "catalog_intelligence": {
                "engine": str(intelligence.get("engine", "") or ""),
                "confidence": round(float(intelligence.get("confidence", 0) or 0), 4),
                "reason": str(intelligence.get("reason", "") or ""),
                "attributes": dict(intelligence.get("atributos", {}) or {}),
                "normalized_description": str(intelligence.get("descricao_normalizada", "") or ""),
            },
        }
        if not prod["dimensoes"] and (prod["comprimento"] > 0 or prod["largura"] > 0 or prod["espessura"] > 0):
            prod["dimensoes"] = self._product_dimensoes(prod)
        prod["preco_unid"] = round(self._parse_float(self.desktop_main.produto_preco_unitario(prod), 0), 4)
        return prod

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
        data = self.ensure_data()
        prod = self._product_normalize_payload(payload)
        code = str(prod.get("codigo", "") or "").strip()
        rows = data.setdefault("produtos", [])
        existing = next((row for row in rows if str(row.get("codigo", "") or "").strip() == code), None)
        old_qty = self._parse_float((existing or {}).get("qty", 0), 0)
        if existing is None:
            rows.append(prod)
            target = prod
        else:
            existing.update(prod)
            target = existing
        target["atualizado_em"] = self.desktop_main.now_iso()
        new_qty = self._parse_float(target.get("qty", 0), 0)
        operador = str((self.user or {}).get("username", "") or "Sistema")
        if existing is None and new_qty > 1e-9:
            self.desktop_main.add_produto_mov(
                data,
                tipo="ENTRADA_INICIAL",
                operador=operador,
                codigo=code,
                descricao=str(target.get("descricao", "") or "").strip(),
                qtd=new_qty,
                antes=0.0,
                depois=new_qty,
                obs="Stock inicial no registo do produto",
                origem="PRODUTOS",
                ref_doc=code,
            )
        elif existing is not None and abs(new_qty - old_qty) > 1e-9:
            delta = new_qty - old_qty
            self.desktop_main.add_produto_mov(
                data,
                tipo="AJUSTE_STOCK",
                operador=operador,
                codigo=code,
                descricao=str(target.get("descricao", "") or "").strip(),
                qtd=abs(delta),
                antes=old_qty,
                depois=new_qty,
                obs=f"Ajuste manual no cadastro ({self._fmt(delta)})",
                origem="PRODUTOS",
                ref_doc=code,
            )
        self.desktop_main.ensure_produto_seq(data, code)
        self._save(force=True)
        try:
            self.conjunto_refresh_prices()
        except Exception:
            pass
        return self.product_detail(code)

    def product_remove(self, codigo: str) -> None:
        self.product_remove_many([codigo])

    def product_remove_many(self, codigos: list[str]) -> int:
        codes = {str(value or "").strip() for value in codigos if str(value or "").strip()}
        if not codes:
            raise ValueError("Seleciona pelo menos um produto.")
        data = self.ensure_data()
        rows = list(data.get("produtos", []) or [])
        removed = [row for row in rows if str(row.get("codigo", "") or "").strip() in codes]
        if not removed:
            raise ValueError("Os produtos selecionados já não existem.")
        data["produtos"] = [row for row in rows if str(row.get("codigo", "") or "").strip() not in codes]
        self._save(force=True)
        try:
            self.ensure_inventory_scan_codes(persist=True)
        except Exception:
            pass
        try:
            self.conjunto_refresh_prices()
        except Exception:
            pass
        return len(removed)

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
