"""Prepare a template and its live assembly before publishing either catalog."""
from copy import deepcopy
from typing import Any, Protocol
from .assembly_catalog import AssemblyCatalog, PreparedAssembly


class AssemblyPairRepository(Protocol):
    def replace(self, template: PreparedAssembly, live: PreparedAssembly) -> None: ...


class AssemblyPair:
    def __init__(self, templates: AssemblyCatalog, live: AssemblyCatalog, repository: AssemblyPairRepository):
        self.templates = templates
        self.live = live
        self.repository = repository

    def save(self, template_payload: dict[str, Any], live_payload: dict[str, Any]) -> str:
        template = self.templates.prepare(template_payload)
        payload = deepcopy(live_payload)
        requested = str(payload.get("codigo", "") or "").strip()
        if requested and requested != template.code:
            raise ValueError("O modelo e o conjunto devem ter o mesmo codigo.")
        payload["codigo"] = template.code
        live = self.live.prepare(payload)
        parameter = next(row.get("param_codigo", "") for row in live.models if row.get("codigo") == live.code)
        next(row for row in template.models if row.get("codigo") == template.code)["param_codigo"] = parameter
        self.repository.replace(template, live)
        return live.code
