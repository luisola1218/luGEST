from __future__ import annotations

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
        dim = str(prod.get("dimensoes", "") or "").strip()
        if dim:
            return dim
        comp = self._parse_float(prod.get("comprimento", 0), 0)
        larg = self._parse_float(prod.get("largura", 0), 0)
        esp = self._parse_float(prod.get("espessura", 0), 0)
        if comp > 0 or larg > 0 or esp > 0:
            return f"{self._fmt(comp)}x{self._fmt(larg)}x{self._fmt(esp)}"
        return "-"

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
        query = str(filter_text or "").strip().lower()
        rows = []
        for index, prod in enumerate(list(self.ensure_data().get("produtos", []) or [])):
            price_unit = round(self._parse_float(self.desktop_main.produto_preco_unitario(prod), 0), 4)
            sale_unit = round(self._parse_float(self.desktop_main.produto_preco_venda(prod), 0), 4)
            catalog_fields = self._product_resolve_catalog_fields(prod)
            qty = self._parse_float(prod.get("qty", 0), 0)
            physical_qty = qty + max(0.0, self._parse_float(prod.get("quality_pending_qty", 0), 0))
            if in_stock_only and qty <= 0:
                continue
            quality_status = str(prod.get("quality_status", "") or "").strip()
            quality_display_status = (
                "EM_INSPECAO"
                if self._parse_float(prod.get("quality_pending_qty", 0), 0) > 0
                else quality_status
            )
            quality_blocked = bool(prod.get("quality_blocked")) or (
                bool(quality_display_status) and not self._quality_status_is_available(quality_display_status)
            )
            alerta = self._parse_float(prod.get("alerta", 0), 0)
            row = {
                "codigo": str(prod.get("codigo", "") or "").strip(),
                "scan_code": str(prod.get("scan_code", "") or self.inventory_scan_code("PRD", prod.get("codigo"))).strip(),
                "descricao": str(prod.get("descricao", "") or "").strip(),
                "categoria": str(prod.get("categoria", catalog_fields.get("categoria", "")) or "").strip(),
                "category_id": str(prod.get("category_id", catalog_fields.get("category_id", "")) or "").strip(),
                "subcat": str(prod.get("subcat", catalog_fields.get("subcat", "")) or "").strip(),
                "subcategory_id": str(prod.get("subcategory_id", catalog_fields.get("subcategory_id", "")) or "").strip(),
                "tipo": str(prod.get("tipo", catalog_fields.get("tipo", "")) or "").strip(),
                "type_id": str(prod.get("type_id", catalog_fields.get("type_id", "")) or "").strip(),
                "category_icon": str(prod.get("category_icon", catalog_fields.get("category_icon", "")) or "").strip(),
                "category_badge": str(prod.get("category_badge", catalog_fields.get("category_badge", "")) or "").strip(),
                "category_tone": str(prod.get("category_tone", catalog_fields.get("category_tone", "")) or "").strip(),
                "dimensoes": self._product_dimensoes(prod),
                "unid": str(prod.get("unid", "UN") or "UN").strip() or "UN",
                "qty": physical_qty,
                "available_qty": qty,
                "alerta": alerta,
                "p_compra": round(self._parse_float(prod.get("p_compra", 0), 0), 4),
                "pvp1": round(self._parse_float(prod.get("pvp1", 0), 0), 4),
                "pvp2": round(self._parse_float(prod.get("pvp2", 0), 0), 4),
                "preco_unid": price_unit,
                "preco_venda": sale_unit,
                "valor_stock": round(physical_qty * price_unit, 2),
                "metros_unidade": round(self._parse_float(prod.get("metros_unidade", prod.get("metros", 0)), 0), 4),
                "peso_unid": round(self._parse_float(prod.get("peso_unid", 0), 0), 4),
                "fabricante": str(prod.get("fabricante", "") or "").strip(),
                "modelo": str(prod.get("modelo", "") or "").strip(),
                "obs": str(prod.get("obs", "") or "").strip(),
                "quality_status": quality_display_status,
                "quality_pending_qty": round(self._parse_float(prod.get("quality_pending_qty", 0), 0), 4),
                "updated_at": str(prod.get("atualizado_em", "") or "").strip(),
            }
            if row["subcat"] and row["tipo"]:
                row["type_display"] = f"{row['subcat']} / {row['tipo']}"
            else:
                row["type_display"] = row["subcat"] or row["tipo"] or "-"
            row["category_display"] = (
                f"{row['category_icon']} {row['categoria']}".strip()
                if row["category_icon"]
                else (row["categoria"] or "-")
            )
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            if quality_blocked:
                severity = "warning"
            elif qty <= 0 or (alerta > 0 and qty <= alerta):
                severity = "warning"
            else:
                severity = "ok"
            row["severity"] = severity
            row["band"] = "even" if index % 2 == 0 else "odd"
            rows.append(row)
        rows.sort(key=lambda item: (item.get("codigo") or "", item.get("descricao") or ""))
        return rows

    def _product_issue_meta_from_obs(self, obs: str) -> dict[str, Any]:
        text = str(obs or "").strip()
        meta: dict[str, Any] = {}
        if "|meta|" not in text:
            return meta
        for part in text.split("|meta|")[1:]:
            chunk = part.strip()
            if "=" not in chunk:
                continue
            key, value = chunk.split("=", 1)
            meta[key.strip()] = value.strip()
        return meta

    def product_movement_years(self, codigo: str = "") -> list[str]:
        years: set[str] = set()
        for row in self.product_movements(codigo=codigo, limit=5000):
            stamp = str(row.get("data", "") or "")
            if len(stamp) >= 4 and stamp[:4].isdigit():
                years.add(stamp[:4])
        if not years:
            years.add(str(date.today().year))
        return sorted(years, reverse=True)

    def product_movements(
        self,
        codigo: str = "",
        limit: int = 120,
        operator_name: str = "",
        year: str = "",
        issue_only: bool = False,
    ) -> list[dict[str, Any]]:
        code = str(codigo or "").strip()
        operator_name = str(operator_name or "").strip().lower()
        year = str(year or "").strip()
        rows = []
        for row in reversed(list(self.ensure_data().get("produtos_mov", []) or [])):
            mov = dict(row or {})
            mov_code = str(mov.get("codigo", "") or mov.get("produto", "") or "").strip()
            if code and mov_code != code:
                continue
            mov_tipo = str(mov.get("tipo", "") or "").strip()
            if issue_only and mov_tipo != "ENTREGA_OPERADOR":
                continue
            mov_operator = str(mov.get("operador", "") or "").strip()
            if operator_name and mov_operator.lower() != operator_name:
                continue
            mov_data = str(mov.get("data", "") or "").replace("T", " ")[:19]
            if year and year != "Todos" and not mov_data.startswith(year):
                continue
            meta = self._product_issue_meta_from_obs(mov.get("obs", ""))
            rows.append(
                {
                    "data": mov_data,
                    "tipo": mov_tipo,
                    "operador": mov_operator,
                    "codigo": mov_code,
                    "descricao": str(mov.get("descricao", "") or "").strip(),
                    "qtd": round(self._parse_float(mov.get("qtd", 0), 0), 2),
                    "antes": round(self._parse_float(mov.get("antes", 0), 0), 2),
                    "depois": round(self._parse_float(mov.get("depois", 0), 0), 2),
                    "obs": str(mov.get("obs", "") or "").strip(),
                    "origem": str(mov.get("origem", "") or "").strip(),
                    "valor_unit": round(self._parse_float(meta.get("valor_unit", 0), 0), 4),
                    "valor_total": round(self._parse_float(meta.get("valor_total", 0), 0), 2),
                }
            )
            if len(rows) >= limit:
                break
        return rows

    def product_issue_summary(self, operator_name: str = "", year: str = "", codigo: str = "") -> dict[str, Any]:
        rows = self.product_movements(codigo=codigo, limit=5000, operator_name=operator_name, year=year, issue_only=True)
        total_qtd = sum(self._parse_float(row.get("qtd", 0), 0) for row in rows)
        total_valor = sum(self._parse_float(row.get("valor_total", 0), 0) for row in rows)
        return {
            "linhas": len(rows),
            "qtd_total": round(total_qtd, 2),
            "valor_total": round(total_valor, 2),
        }

    def product_detail(self, codigo: str) -> dict[str, Any]:
        code = str(codigo or "").strip()
        prod = next((row for row in list(self.ensure_data().get("produtos", []) or []) if str(row.get("codigo", "") or "").strip() == code), None)
        if prod is None:
            raise ValueError("Produto não encontrado.")
        detail = dict(prod)
        detail["preco_unid"] = round(self._parse_float(self.desktop_main.produto_preco_unitario(detail), 0), 4)
        detail["valor_stock"] = round(self._parse_float(detail.get("qty", 0), 0) * detail["preco_unid"], 2)
        detail["dimensoes"] = self._product_dimensoes(detail)
        detail["movimentos"] = self.product_movements(code, limit=80)
        return detail

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
        code = str(codigo or "").strip()
        qty = self._parse_float(quantidade, 0)
        if qty <= 0:
            raise ValueError("Quantidade invalida.")
        prod = next((row for row in list(self.ensure_data().get("produtos", []) or []) if str(row.get("codigo", "") or "").strip() == code), None)
        if prod is None:
            raise ValueError("Produto não encontrado.")
        before = self._parse_float(prod.get("qty", 0), 0)
        if qty > before + 1e-9:
            raise ValueError("Quantidade superior ao stock disponivel.")
        prod["qty"] = max(0.0, before - qty)
        prod["atualizado_em"] = self.desktop_main.now_iso()
        actor = str((self.user or {}).get("username", "") or "Sistema")
        issue_mode = str(issue_mode or "stock").strip().lower()
        operator_txt = str(target_operator or "").strip()
        movement_type = "BAIXA"
        movement_operator = operator_txt or actor
        detail_obs = str(obs or "").strip() or "Baixa manual no desktop Qt"
        if issue_mode == "operator":
            if not operator_txt:
                raise ValueError("Seleciona o operador que recebe o material.")
            movement_type = "ENTREGA_OPERADOR"
            movement_operator = operator_txt
            valor_unit = round(self._parse_float(self.desktop_main.produto_preco_unitario(prod), 0), 4)
            valor_total = round(valor_unit * qty, 2)
            note = str(obs or "").strip() or "Entrega a operador"
            detail_obs = (
                f"{note} |meta|actor={actor} |meta|valor_unit={valor_unit:.4f} "
                f"|meta|valor_total={valor_total:.2f}"
            )
        self.desktop_main.add_produto_mov(
            self.ensure_data(),
            tipo=movement_type,
            operador=movement_operator,
            codigo=code,
            descricao=str(prod.get("descricao", "") or "").strip(),
            qtd=qty,
            antes=before,
            depois=self._parse_float(prod.get("qty", 0), 0),
            obs=detail_obs,
            origem="OPERADOR" if movement_type == "ENTREGA_OPERADOR" else "PRODUTOS",
            ref_doc=code,
        )
        self._save(force=True)
        return self.product_detail(code)
