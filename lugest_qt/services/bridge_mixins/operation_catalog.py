from __future__ import annotations

from typing import Any


class OperationCatalogBackendMixin:
    """Legacy adapter for operation catalog; see BACKEND_GUIDE.md."""

    def planning_operation_options(self) -> list[str]:
        return [
            str(row.get("name", "") or "").strip()
            for row in self.operation_catalog_rows()
            if bool(row.get("active", True)) and bool(row.get("planeavel", False))
        ]

    def _default_operation_catalog(self) -> list[dict[str, Any]]:
        planning_defaults = [
            "Corte Laser",
            "Quinagem",
            "Serralharia",
            "Maquinacao",
            "Roscagem",
            "Lacagem",
            "Montagem",
            "Embalamento",
            "Expedicao",
            "Furo Manual",
        ]
        non_planning_defaults = [
            "Departamento de Desenho",
            "Departamento de Orcamentacao",
            "Outros",
        ]
        names = list(dict.fromkeys(
            list(getattr(self.desktop_main, "OFF_OPERACOES_DISPONIVEIS", []) or [])
            + list(getattr(self.desktop_main, "PLANEAMENTO_OPERACOES_DISPONIVEIS", []) or [])
            + planning_defaults
            + non_planning_defaults
        ))
        planning_keys = {self.desktop_main.norm_text(name) for name in planning_defaults}
        return [
            {
                "name": self._planning_normalize_operation(name, default=str(name).strip()) if self.desktop_main.norm_text(name) in planning_keys else str(name).strip(),
                "active": True,
                "planeavel": self.desktop_main.norm_text(name) in planning_keys,
            }
            for name in names
            if str(name or "").strip()
        ]

    def operation_catalog_rows(self) -> list[dict[str, Any]]:
        cached = getattr(self, "_operation_catalog_cache", None)
        generation_now = int(getattr(self, "_data_cache_generation", 0))
        if cached is not None:
            generation, cached_rows = cached
            if generation == generation_now:
                return [dict(row) for row in cached_rows]
        data = self.ensure_data()
        raw_rows = list(data.get("operations_catalog", []) or [])
        seed_rows = self._default_operation_catalog()
        seed_rows.extend(dict(row or {}) for row in raw_rows if isinstance(row, dict))
        for row in list(self._workcenter_catalog(sync_legacy=False) or []):
            op_name = self._planning_normalize_operation(row.get("operation", row.get("name", "")), default=str(row.get("operation", row.get("name", "")) or "").strip())
            if op_name:
                seed_rows.append({"name": op_name})
        merged: dict[str, dict[str, Any]] = {}
        for raw in seed_rows:
            name = str(raw.get("name", raw.get("nome", "")) or "").strip()
            if not name:
                continue
            normalized = self._planning_normalize_operation(name, default=name)
            if self.desktop_main.norm_text(name).startswith("departamento") or self.desktop_main.norm_text(name) == "outros":
                normalized = name
            key = normalized.casefold()
            current = merged.get(key)
            if current is None:
                current = {"name": normalized, "active": True, "planeavel": False}
                merged[key] = current
            if "active" in raw or "ativo" in raw:
                current["active"] = bool(raw.get("active", raw.get("ativo")))
            if "planeavel" in raw:
                current["planeavel"] = bool(raw.get("planeavel"))
        ordered_names = [row["name"] for row in self._default_operation_catalog()]
        order_index = {name.casefold(): index for index, name in enumerate(ordered_names)}
        cleaned = sorted(
            merged.values(),
            key=lambda row: (order_index.get(str(row.get("name", "")).casefold(), 999), str(row.get("name", "")).casefold()),
        )
        data["operations_catalog"] = cleaned
        self._operation_catalog_cache = (
            generation_now,
            [dict(row) for row in cleaned],
        )
        return [dict(row) for row in cleaned]

    def operation_catalog_options(self, *, include_inactive: bool = False, planeavel_only: bool = False) -> list[str]:
        rows = []
        for row in self.operation_catalog_rows():
            if not include_inactive and not bool(row.get("active", True)):
                continue
            if planeavel_only and not bool(row.get("planeavel", False)):
                continue
            name = str(row.get("name", "") or "").strip()
            if name:
                rows.append(name)
        return rows

    def _operation_usage_counts(self, operation: str, *, data: dict[str, Any] | None = None) -> dict[str, int]:
        op_txt = self._planning_normalize_operation(operation, default=str(operation or "").strip())
        if not op_txt:
            return {"workcenters": 0, "quotes": 0, "orders": 0, "planning": 0, "total": 0}
        data = data if isinstance(data, dict) else self.ensure_data()
        op_key = self.desktop_main.norm_text(op_txt)

        def op_matches(value: Any) -> bool:
            normalized = self._planning_normalize_operation(value, default=str(value or "").strip())
            return bool(normalized) and self.desktop_main.norm_text(normalized) == op_key

        workcenters = sum(1 for row in list(data.get("workcenter_catalog", []) or []) if op_matches((row or {}).get("operation", (row or {}).get("name", ""))))
        quotes = 0
        for quote in list(data.get("orcamentos", []) or []):
            for line in list((quote or {}).get("linhas", []) or []):
                if any(op_matches(op_name) for op_name in self._planning_ops_from_ops_value((line or {}).get("operacao", ""))):
                    quotes += 1
                    break
        orders = 0
        for order in list(data.get("encomendas", []) or []):
            order_used = False
            for mat in list((order or {}).get("materiais", []) or []):
                for esp in list((mat or {}).get("espessuras", []) or []):
                    maps = [
                        dict((esp or {}).get("tempos_operacao", {}) or {}),
                        dict((esp or {}).get("maquinas_operacao", (esp or {}).get("recursos_operacao", {})) or {}),
                    ]
                    if any(op_matches(op_name) for values in maps for op_name in values.keys()):
                        order_used = True
                        break
                if order_used:
                    break
            if order_used:
                orders += 1
        planning = sum(
            1
            for bucket_name in ("plano", "plano_hist")
            for row in list(data.get(bucket_name, []) or [])
            if op_matches((row or {}).get("operacao", ""))
        )
        total = workcenters + quotes + orders + planning
        return {"workcenters": workcenters, "quotes": quotes, "orders": orders, "planning": planning, "total": total}

    def save_operation_catalog_row(self, name: str, *, current_name: str = "", active: bool = True, planeavel: bool = False) -> dict[str, Any]:
        data = self.ensure_data()
        new_name = str(name or "").strip()
        current_txt = str(current_name or "").strip()
        if not new_name:
            raise ValueError("Nome da operação obrigatório.")
        rows = self.operation_catalog_rows()
        new_key = new_name.casefold()
        current_key = current_txt.casefold()
        if any(str(row.get("name", "") or "").strip().casefold() == new_key and str(row.get("name", "") or "").strip().casefold() != current_key for row in rows):
            raise ValueError("Já existe uma operação com esse nome.")
        target = next((row for row in rows if str(row.get("name", "") or "").strip().casefold() == current_key), None) if current_txt else None
        if target is None:
            target = {"name": new_name}
            rows.append(target)
        old_name = str(target.get("name", "") or "").strip()
        target["name"] = self._planning_normalize_operation(new_name, default=new_name)
        if self.desktop_main.norm_text(new_name).startswith("departamento") or self.desktop_main.norm_text(new_name) == "outros":
            target["name"] = new_name
        target["active"] = bool(active)
        target["planeavel"] = bool(planeavel)
        if old_name and old_name.casefold() != str(target["name"]).casefold():
            for wc in list(data.get("workcenter_catalog", []) or []):
                if str((wc or {}).get("operation", "") or "").strip().casefold() == old_name.casefold():
                    wc["operation"] = target["name"]
            for bucket_name in ("plano", "plano_hist"):
                for block in list(data.get(bucket_name, []) or []):
                    if str((block or {}).get("operacao", "") or "").strip().casefold() == old_name.casefold():
                        block["operacao"] = target["name"]
        data["operations_catalog"] = rows
        self.operation_catalog_rows()
        self._workcenter_catalog()
        self._save(force=True)
        return next(row for row in self.operation_catalog_rows() if str(row.get("name", "") or "").strip().casefold() == str(target["name"]).casefold())

    def remove_operation_catalog_row(self, name: str) -> None:
        data = self.ensure_data()
        target = str(name or "").strip()
        if not target:
            raise ValueError("Operação inválida.")
        rows = self.operation_catalog_rows()
        current = next((row for row in rows if str(row.get("name", "") or "").strip().casefold() == target.casefold()), None)
        if current is None:
            raise ValueError("Operação não encontrada.")
        usage = self._operation_usage_counts(target, data=data)
        if usage["total"] > 0:
            raise ValueError(
                "Não é possível remover esta operação porque ainda está em uso "
                f"(postos: {usage['workcenters']}, orçamentos: {usage['quotes']}, encomendas: {usage['orders']}, planeamento: {usage['planning']})."
            )
        data["operations_catalog"] = [
            row
            for row in rows
            if str(row.get("name", "") or "").strip().casefold() != target.casefold()
        ]
        self.operation_catalog_rows()
        self._save(force=True)
