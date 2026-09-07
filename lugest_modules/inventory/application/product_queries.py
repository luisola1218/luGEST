"""Product queries depend on a read repository and explicit calculation rules."""
from __future__ import annotations
from dataclasses import dataclass
from datetime import date
from typing import Any, Callable, Protocol

class ProductReadRepository(Protocol):
    def products(self) -> list[dict]: ...
    def movements(self) -> list[dict]: ...

@dataclass(frozen=True)
class ProductQueryRules:
    _fmt: Callable[..., Any]
    _parse_float: Callable[..., Any]
    _product_resolve_catalog_fields: Callable[..., Any]
    _quality_status_is_available: Callable[..., Any]
    inventory_scan_code: Callable[..., Any]
    produto_preco_unitario: Callable[..., Any]
    produto_preco_venda: Callable[..., Any]

class ProductQueries:
    def __init__(self, repository: ProductReadRepository, rules: ProductQueryRules,
                 *, today: Callable[[], date] = date.today):
        self.repository = repository
        self.rules = rules
        self.today = today

    def _product_dimensoes(self, prod: dict[str, Any]) -> str:
        dim = str(prod.get("dimensoes", "") or "").strip()
        if dim:
            return dim
        comp = self.rules._parse_float(prod.get("comprimento", 0), 0)
        larg = self.rules._parse_float(prod.get("largura", 0), 0)
        esp = self.rules._parse_float(prod.get("espessura", 0), 0)
        if comp > 0 or larg > 0 or esp > 0:
            return f"{self.rules._fmt(comp)}x{self.rules._fmt(larg)}x{self.rules._fmt(esp)}"
        return "-"

    def product_rows(self, filter_text: str = "", in_stock_only: bool = False) -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        rows = []
        for index, prod in enumerate(self.repository.products()):
            price_unit = round(self.rules._parse_float(self.rules.produto_preco_unitario(prod), 0), 4)
            sale_unit = round(self.rules._parse_float(self.rules.produto_preco_venda(prod), 0), 4)
            catalog_fields = self.rules._product_resolve_catalog_fields(prod)
            qty = self.rules._parse_float(prod.get("qty", 0), 0)
            physical_qty = qty + max(0.0, self.rules._parse_float(prod.get("quality_pending_qty", 0), 0))
            if in_stock_only and qty <= 0:
                continue
            quality_status = str(prod.get("quality_status", "") or "").strip()
            quality_display_status = (
                "EM_INSPECAO"
                if self.rules._parse_float(prod.get("quality_pending_qty", 0), 0) > 0
                else quality_status
            )
            quality_blocked = bool(prod.get("quality_blocked")) or (
                bool(quality_display_status) and not self.rules._quality_status_is_available(quality_display_status)
            )
            alerta = self.rules._parse_float(prod.get("alerta", 0), 0)
            row = {
                "codigo": str(prod.get("codigo", "") or "").strip(),
                "scan_code": str(prod.get("scan_code", "") or self.rules.inventory_scan_code("PRD", prod.get("codigo"))).strip(),
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
                "p_compra": round(self.rules._parse_float(prod.get("p_compra", 0), 0), 4),
                "pvp1": round(self.rules._parse_float(prod.get("pvp1", 0), 0), 4),
                "pvp2": round(self.rules._parse_float(prod.get("pvp2", 0), 0), 4),
                "preco_unid": price_unit,
                "preco_venda": sale_unit,
                "valor_stock": round(physical_qty * price_unit, 2),
                "metros_unidade": round(self.rules._parse_float(prod.get("metros_unidade", prod.get("metros", 0)), 0), 4),
                "peso_unid": round(self.rules._parse_float(prod.get("peso_unid", 0), 0), 4),
                "fabricante": str(prod.get("fabricante", "") or "").strip(),
                "modelo": str(prod.get("modelo", "") or "").strip(),
                "obs": str(prod.get("obs", "") or "").strip(),
                "quality_status": quality_display_status,
                "quality_pending_qty": round(self.rules._parse_float(prod.get("quality_pending_qty", 0), 0), 4),
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
            years.add(str(self.today().year))
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
        for row in reversed(self.repository.movements()):
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
                    "qtd": round(self.rules._parse_float(mov.get("qtd", 0), 0), 2),
                    "antes": round(self.rules._parse_float(mov.get("antes", 0), 0), 2),
                    "depois": round(self.rules._parse_float(mov.get("depois", 0), 0), 2),
                    "obs": str(mov.get("obs", "") or "").strip(),
                    "origem": str(mov.get("origem", "") or "").strip(),
                    "valor_unit": round(self.rules._parse_float(meta.get("valor_unit", 0), 0), 4),
                    "valor_total": round(self.rules._parse_float(meta.get("valor_total", 0), 0), 2),
                }
            )
            if len(rows) >= limit:
                break
        return rows

    def product_issue_summary(self, operator_name: str = "", year: str = "", codigo: str = "") -> dict[str, Any]:
        rows = self.product_movements(codigo=codigo, limit=5000, operator_name=operator_name, year=year, issue_only=True)
        total_qtd = sum(self.rules._parse_float(row.get("qtd", 0), 0) for row in rows)
        total_valor = sum(self.rules._parse_float(row.get("valor_total", 0), 0) for row in rows)
        return {
            "linhas": len(rows),
            "qtd_total": round(total_qtd, 2),
            "valor_total": round(total_valor, 2),
        }

    def product_detail(self, codigo: str) -> dict[str, Any]:
        code = str(codigo or "").strip()
        prod = next((row for row in self.repository.products() if str(row.get("codigo", "") or "").strip() == code), None)
        if prod is None:
            raise ValueError("Produto não encontrado.")
        detail = dict(prod)
        detail["preco_unid"] = round(self.rules._parse_float(self.rules.produto_preco_unitario(detail), 0), 4)
        detail["valor_stock"] = round(self.rules._parse_float(detail.get("qty", 0), 0) * detail["preco_unid"], 2)
        detail["dimensoes"] = self._product_dimensoes(detail)
        detail["movimentos"] = self.product_movements(code, limit=80)
        return detail
