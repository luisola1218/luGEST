"""Reprice an assembly catalog as a single prepared change."""
from copy import deepcopy
from typing import Any, Protocol

from lugest_modules.quotes.application.assemblies import AssemblyRules, refresh_model


class AssemblyCatalogRepository(Protocol):
    def models(self) -> list[dict[str, Any]]: ...
    def replace(self, models: list[dict[str, Any]], *, expected: list[dict[str, Any]]) -> None: ...


def assign_parameter_codes(models: list[dict[str, Any]]) -> bool:
    """Normalize unique parameter codes in a caller-owned catalog draft."""
    changed = False
    used: set[str] = set()
    highest = 0
    rows = [row for row in models if isinstance(row, dict)]
    for row in rows:
        raw = str(row.get("param_codigo", "") or "").strip()
        digits = "".join(ch for ch in raw if ch.isdigit())
        if digits:
            normalized = f"{int(digits):04d}"
            if normalized not in used:
                used.add(normalized)
                highest = max(highest, int(normalized))
                if raw != normalized:
                    row["param_codigo"] = normalized
                    changed = True
                continue
        row["param_codigo"] = ""
    for row in rows:
        if str(row.get("param_codigo", "") or "").strip():
            continue
        highest += 1
        while f"{highest:04d}" in used:
            highest += 1
        row["param_codigo"] = f"{highest:04d}"
        used.add(row["param_codigo"])
        changed = True
    return changed


class AssemblyRefresh:
    def __init__(self, repository: AssemblyCatalogRepository, rules: AssemblyRules):
        self.repository = repository
        self.rules = rules

    def refresh(self, codigo: str = "") -> dict[str, Any]:
        code = str(codigo or "").strip()
        original = self.repository.models()
        models = deepcopy(original)
        matched = next((row for row in models if isinstance(row, dict)
                        and str(row.get("codigo", "") or "").strip() == code), None)
        if code and matched is None:
            raise ValueError("Conjunto nao encontrado.")
        changed = assign_parameter_codes(models)
        result = None
        for index, model in enumerate(models):
            if not isinstance(model, dict) or (code and str(model.get("codigo", "") or "").strip() != code):
                continue
            refreshed, item_changed = refresh_model(self.rules, model)
            models[index] = refreshed
            changed = item_changed or changed
            if code:
                result = refreshed
        if changed:
            self.repository.replace(models, expected=original)
        return deepcopy(result) if code else {"updated": changed}
