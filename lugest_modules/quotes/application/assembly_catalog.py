"""Create, update and remove assembly aggregates before touching persistence."""
from copy import deepcopy
from typing import Any, Callable

from lugest_modules.quotes.application.assemblies import AssemblyRules, normalize_item, refresh_model, technical_sheet
from lugest_modules.quotes.application.assembly_refresh import AssemblyCatalogRepository, assign_parameter_codes


class AssemblyCatalog:
    def __init__(self, repository: AssemblyCatalogRepository, rules: AssemblyRules,
                 next_code: Callable[[], str], *, live_prices: bool):
        self.repository = repository
        self.rules = rules
        self.next_code = next_code
        self.live_prices = live_prices

    def save(self, payload: dict[str, Any]) -> str:
        payload = deepcopy(payload)
        description = str(payload.get("descricao", "") or "").strip()
        if not description:
            raise ValueError("Descricao obrigatoria no conjunto.")
        items = [normalize_item(self.rules, dict(row or {})) for row in list(payload.get("itens", []) or [])]
        if not items:
            raise ValueError("O conjunto precisa de pelo menos um item.")
        code = str(payload.get("codigo", "") or "").strip() or self.next_code()
        original = self.repository.models()
        models = deepcopy(original)
        if self.live_prices:
            assign_parameter_codes(models)
        existing = next((row for row in models if str(row.get("codigo", "") or "").strip() == code), None)
        model = deepcopy(existing) if existing is not None else {}
        created = str((existing or {}).get("created_at", "") or payload.get("created_at", "") or "").strip()
        model.update({
            "codigo": code,
            "param_codigo": str(payload.get("param_codigo", "") or "").strip(),
            "descricao": description,
            "notas": str(payload.get("notas", "") or "").strip(),
            "ativo": bool(payload.get("ativo", True)),
            "template": bool(payload.get("template", False)),
            "origem": str(payload.get("origem", "") or "").strip(),
            "created_at": created or self.rules.now_iso(),
            "updated_at": self.rules.now_iso(),
            "ficha_tecnica": technical_sheet(payload.get("ficha_tecnica", (existing or {}).get("ficha_tecnica", {}))),
            "itens": [{**item, "linha_ordem": index} for index, item in enumerate(items, start=1)],
        })
        if self.live_prices:
            used = {str(row.get("param_codigo", "") or "").strip() for row in models if isinstance(row, dict)}
            highest = max((int(value) for value in used if value.isdigit()), default=0)
            if existing is not None:
                model["param_codigo"] = str(existing.get("param_codigo", "") or model["param_codigo"] or f"{highest + 1:04d}").strip()
            elif not model["param_codigo"] or model["param_codigo"] in used:
                model["param_codigo"] = f"{highest + 1:04d}"
            model["margem_perc"] = round(self.rules.parse_float(payload.get("margem_perc", 0), 0), 2)
            model, _ = refresh_model(self.rules, model)
        if existing is None:
            models.append(model)
        else:
            models[models.index(existing)] = model
        self.repository.replace(models, expected=original)
        return code

    def remove(self, codigo: str) -> None:
        code = str(codigo or "").strip()
        original = self.repository.models()
        models = [row for row in original if str(row.get("codigo", "") or "").strip() != code]
        if len(models) == len(original):
            raise ValueError("Conjunto nao encontrado.")
        self.repository.replace(models, expected=original)
