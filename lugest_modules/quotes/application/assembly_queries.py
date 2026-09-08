"""Detached assembly read models for the saved and template catalogs."""
from typing import Any, Callable
from lugest_modules.quotes.application.assemblies import AssemblyRules, normalize_item, technical_sheet
from lugest_modules.quotes.application.assembly_refresh import AssemblyCatalogRepository

class AssemblyQueries:
    def __init__(self, repository: AssemblyCatalogRepository, rules: AssemblyRules, is_service: Callable):
        self.repository = repository
        self.rules = rules
        self.is_service = is_service

    def template_rows(self, filter_text: str = "") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        rows: list[dict[str, Any]] = []
        for model in self.repository.models():
            if not isinstance(model, dict):
                continue
            items = list(model.get("itens", []) or [])
            row = {
                "codigo": str(model.get("codigo", "") or "").strip(),
                "descricao": str(model.get("descricao", "") or "").strip(),
                "ativo": bool(model.get("ativo", True)),
                "template": bool(model.get("template", False)),
                "origem": str(model.get("origem", "") or "").strip(),
                "itens": len(items),
                "pecas": sum(1 for item in items if self.rules.orc_line_is_piece(item)),
                "produtos": sum(1 for item in items if self.rules.orc_line_is_product(item)),
                "servicos": sum(1 for item in items if self.is_service(item)),
                "total_base": round(sum(self.rules.parse_float(item.get("qtd", 0), 0) * self.rules.parse_float(item.get("preco_unit", 0), 0) for item in items), 2),
                "notas": str(model.get("notas", "") or "").strip(),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: (item.get("codigo", ""), item.get("descricao", "")))
        return rows


    def template_detail(self, codigo: str) -> dict[str, Any]:
        code = str(codigo or "").strip()
        model = next(
            (
                row
                for row in self.repository.models()
                if str(row.get("codigo", "") or "").strip() == code
            ),
            None,
        )
        if model is None:
            raise ValueError("Conjunto nao encontrado.")
        items = [normalize_item(self.rules, dict(item or {})) for item in list(model.get("itens", []) or [])]
        return {
            "codigo": str(model.get("codigo", "") or "").strip(),
            "param_codigo": str(model.get("param_codigo", "") or "").strip(),
            "descricao": str(model.get("descricao", "") or "").strip(),
            "notas": str(model.get("notas", "") or "").strip(),
            "ativo": bool(model.get("ativo", True)),
            "template": bool(model.get("template", False)),
            "origem": str(model.get("origem", "") or "").strip(),
            "created_at": str(model.get("created_at", "") or "").strip(),
            "updated_at": str(model.get("updated_at", "") or "").strip(),
            "ficha_tecnica": technical_sheet(model.get("ficha_tecnica", {})),
            "itens": items,
        }


    def rows(self, filter_text: str = "") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        rows: list[dict[str, Any]] = []
        for model in self.repository.models():
            if not isinstance(model, dict):
                continue
            items = list(model.get("itens", []) or [])
            row = {
                "codigo": str(model.get("codigo", "") or "").strip(),
                "param_codigo": str(model.get("param_codigo", "") or "").strip(),
                "descricao": str(model.get("descricao", "") or "").strip(),
                "ativo": bool(model.get("ativo", True)),
                "template": bool(model.get("template", False)),
                "origem": str(model.get("origem", "") or "").strip(),
                "itens": len(items),
                "pecas": sum(1 for item in items if self.rules.orc_line_is_piece(item)),
                "produtos": sum(1 for item in items if self.rules.orc_line_is_product(item)),
                "servicos": sum(1 for item in items if self.is_service(item)),
                "total_custo": round(self.rules.parse_float(model.get("total_custo", 0), 0), 2),
                "total_final": round(self.rules.parse_float(model.get("total_final", 0), 0), 2),
                "margem_perc": round(self.rules.parse_float(model.get("margem_perc", 0), 0), 2),
                "notas": str(model.get("notas", "") or "").strip(),
                "created_at": str(model.get("created_at", "") or "").strip(),
                "updated_at": str(model.get("updated_at", "") or "").strip(),
                "precos_atualizados_em": str(model.get("precos_atualizados_em", "") or "").strip(),
                "itens_ligados": sum(1 for item in items if bool(item.get("pricing_linked"))),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: (item.get("codigo", ""), item.get("descricao", "")))
        return rows


    def detail(self, codigo: str) -> dict[str, Any]:
        code = str(codigo or "").strip()
        model = next(
            (
                row
                for row in self.repository.models()
                if str(row.get("codigo", "") or "").strip() == code
            ),
            None,
        )
        if model is None:
            raise ValueError("Conjunto nao encontrado.")
        items = [dict(item or {}) for item in list(model.get("itens", []) or [])]
        return {
            "codigo": str(model.get("codigo", "") or "").strip(),
            "param_codigo": str(model.get("param_codigo", "") or "").strip(),
            "descricao": str(model.get("descricao", "") or "").strip(),
            "notas": str(model.get("notas", "") or "").strip(),
            "ativo": bool(model.get("ativo", True)),
            "template": bool(model.get("template", False)),
            "origem": str(model.get("origem", "") or "").strip(),
            "margem_perc": round(self.rules.parse_float(model.get("margem_perc", 0), 0), 2),
            "total_custo": round(self.rules.parse_float(model.get("total_custo", 0), 0), 2),
            "total_final": round(self.rules.parse_float(model.get("total_final", 0), 0), 2),
            "created_at": str(model.get("created_at", "") or "").strip(),
            "updated_at": str(model.get("updated_at", "") or "").strip(),
            "precos_atualizados_em": str(model.get("precos_atualizados_em", "") or "").strip(),
            "ficha_tecnica": technical_sheet(model.get("ficha_tecnica", {})),
            "itens": items,
        }


