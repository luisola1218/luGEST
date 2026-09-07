from __future__ import annotations

from typing import Any


class WorkcentersBackendMixin:
    """Legacy adapter for workcenters; see BACKEND_GUIDE.md."""

    def quote_workcenter_options(self) -> list[str]:
        ordered: list[str] = []
        seen: set[str] = set()
        preferred = list(self.workcenter_machine_options("Corte Laser") or [])
        for value in preferred:
            key = str(value or "").strip().lower()
            if not key or key in seen or key == "geral":
                continue
            seen.add(key)
            ordered.append(str(value).strip())
        for group in self.workcenter_group_options():
            group_txt = str(group or "").strip()
            group_key = group_txt.lower()
            if group_txt and group_key not in seen and group_key != "geral":
                seen.add(group_key)
                ordered.append(group_txt)
            for machine in self.workcenter_machine_options(group_txt):
                machine_txt = str(machine or "").strip()
                machine_key = machine_txt.lower()
                if not machine_txt or machine_key in seen or machine_key == "geral":
                    continue
                seen.add(machine_key)
                ordered.append(machine_txt)
        return ordered or ["Laser", "Maquina 3030", "Maquina 5030", "Maquina 5040"]

    def _default_workcenter_catalog(self) -> list[dict[str, Any]]:
        raw_groups = [
            ("Corte Laser", ["Maquina 3030", "Maquina 5030", "Maquina 5040"]),
            ("Quinagem", []),
            ("Serralharia", []),
            ("Maquinacao", []),
            ("Roscagem", []),
            ("Lacagem", []),
            ("Montagem", []),
            ("Soldadura", []),
            ("Embalamento", []),
            ("Furo Manual", []),
            ("Expedicao", []),
        ]
        return [
            {
                "name": str(name),
                "operation": self._planning_normalize_operation(name, default=str(name)),
                "active": True,
                "machines": [{"name": str(machine), "active": True} for machine in list(machines or []) if str(machine).strip()],
            }
            for name, machines in raw_groups
        ]

    def _workcenter_group_aliases(self, name: str) -> set[str]:
        raw = str(name or "").strip()
        norm = self.desktop_main.norm_text(raw)
        aliases = {raw.lower(), norm}
        if "laser" in norm:
            aliases.update({"laser", "corte laser"})
        if "quin" in norm:
            aliases.update({"quinagem", "quin"})
        if "serralh" in norm:
            aliases.update({"serralharia", "serralh"})
        if "maquin" in norm:
            aliases.update({"maquinacao", "maquinação", "maquin"})
        if "rosc" in norm:
            aliases.update({"roscagem", "rosc"})
        if "laca" in norm or "pint" in norm:
            aliases.update({"lacagem", "pintura"})
        if "mont" in norm:
            aliases.update({"montagem", "mont"})
        if "sold" in norm:
            aliases.update({"soldadura", "sold"})
        if "embal" in norm:
            aliases.update({"embalamento", "embal"})
        if "furo" in norm:
            aliases.update({"furo manual", "furo"})
        if "exped" in norm:
            aliases.update({"expedicao", "expedição", "shipping"})
        return {alias for alias in aliases if alias}

    def _legacy_workcenter_group_name(self, value: Any) -> str:
        raw = str(value or "").strip()
        if not raw:
            return ""
        catalog = list(self._workcenter_catalog(sync_legacy=False) or [])
        raw_lower = raw.lower()
        raw_norm = self.desktop_main.norm_text(raw)
        for row in catalog:
            group_name = str(row.get("name", "") or "").strip()
            if not group_name:
                continue
            if raw_lower == group_name.lower() or raw_norm in self._workcenter_group_aliases(group_name):
                return group_name
            for machine in list(row.get("machines", []) or []):
                machine_txt = str(machine or "").strip()
                if machine_txt and machine_txt.lower() == raw_lower:
                    return group_name
        if raw.replace(" ", "").isdigit() or raw_norm.startswith("maquina "):
            return "Corte Laser"
        alias_map = {
            "laser": "Corte Laser",
            "quin": "Quinagem",
            "serralh": "Serralharia",
            "maquin": "Maquinacao",
            "rosc": "Roscagem",
            "laca": "Lacagem",
            "pint": "Lacagem",
            "mont": "Montagem",
            "sold": "Soldadura",
            "embal": "Embalamento",
            "furo": "Furo Manual",
            "exped": "Expedicao",
        }
        for token, canonical in alias_map.items():
            if token in raw_norm:
                return canonical
        return raw

    def _workcenter_catalog(self, *, sync_legacy: bool = True) -> list[dict[str, Any]]:
        data = self.ensure_data()
        raw_catalog = list(data.get("workcenter_catalog", []) or [])
        default_catalog = list(self._default_workcenter_catalog() or [])
        groups_by_name: dict[str, dict[str, Any]] = {}

        def has_meaningful_usage() -> bool:
            for row in list(data.get("users", []) or []):
                if any(str(row.get(key, "") or "").strip() for key in ("posto", "posto_trabalho", "work_center", "workcenter")):
                    return True
            for row in list(data.get("orcamentos", []) or []):
                if str(row.get("posto_trabalho", "") or "").strip():
                    return True
            for row in list(data.get("encomendas", []) or []):
                if any(str(row.get(key, "") or "").strip() for key in ("posto_trabalho", "posto", "maquina")):
                    return True
                for mat in list(row.get("materiais", []) or []):
                    for esp in list(mat.get("espessuras", []) or []):
                        if dict(esp.get("maquinas_operacao", esp.get("recursos_operacao", {})) or {}):
                            return True
            for bucket_name in ("plano", "plano_hist"):
                for row in list(data.get(bucket_name, []) or []):
                    if any(str(row.get(key, "") or "").strip() for key in ("posto", "posto_trabalho", "maquina")):
                        return True
            return False

        suspicious_catalog = False
        raw_group_names = {str((row or {}).get("name", "") or "").strip().lower() for row in raw_catalog if isinstance(row, dict) and str((row or {}).get("name", "") or "").strip()}
        for row in raw_catalog:
            if not isinstance(row, dict):
                continue
            group_name = str(row.get("name", "") or "").strip()
            if group_name.replace(" ", "").isdigit():
                suspicious_catalog = True
                break
            for machine in list(row.get("machines", []) or []):
                machine_txt = str(machine or "").strip()
                if machine_txt and machine_txt.lower() in raw_group_names and machine_txt.lower() != group_name.lower():
                    suspicious_catalog = True
                    break
            if suspicious_catalog:
                break
        if raw_catalog and suspicious_catalog and not has_meaningful_usage():
            raw_catalog = []
            data["workcenter_catalog"] = []
            data["postos_trabalho"] = []

        def ensure_group(name: str, operation: str = "") -> dict[str, Any]:
            group_name = str(name or "").strip()
            if not group_name:
                group_name = "Sem grupo"
            key = group_name.lower()
            row = groups_by_name.get(key)
            if row is None:
                row = {
                    "name": group_name,
                    "operation": str(operation or "").strip() or self._planning_normalize_operation(group_name, default=group_name),
                    "active": True,
                    "machines": [],
                }
                groups_by_name[key] = row
            elif operation and not str(row.get("operation", "") or "").strip():
                row["operation"] = str(operation).strip()
            return row

        def machine_name_from(raw_machine: Any) -> str:
            if isinstance(raw_machine, dict):
                return str(raw_machine.get("name", raw_machine.get("nome", "")) or "").strip()
            return str(raw_machine or "").strip()

        def machine_active_from(raw_machine: Any) -> bool:
            if isinstance(raw_machine, dict):
                return bool(raw_machine.get("active", raw_machine.get("ativo", True)))
            return True

        def add_machine(group_name: str, machine_name: Any, active: bool = True) -> None:
            machine_txt = str(machine_name or "").strip()
            if not machine_txt:
                return
            group_row = ensure_group(group_name)
            for owner in groups_by_name.values():
                if owner is group_row:
                    continue
                if any(machine_name_from(value).lower() == machine_txt.lower() for value in list(owner.get("machines", []) or [])):
                    return
            existing = {
                machine_name_from(value).lower()
                for value in list(group_row.get("machines", []) or [])
                if machine_name_from(value)
            }
            if machine_txt.lower() not in existing and machine_txt.lower() != str(group_row.get("name", "") or "").strip().lower():
                group_row.setdefault("machines", []).append({"name": machine_txt, "active": bool(active)})

        seed_catalog = default_catalog
        for default_row in seed_catalog:
            group_row = ensure_group(str(default_row.get("name", "") or "").strip(), str(default_row.get("operation", "") or "").strip())
            group_row["active"] = bool(default_row.get("active", True))
            for machine in list(default_row.get("machines", []) or []):
                add_machine(str(group_row.get("name", "") or "").strip(), machine_name_from(machine), machine_active_from(machine))

        for raw_row in raw_catalog:
            if not isinstance(raw_row, dict):
                continue
            group_name = str(raw_row.get("name", "") or "").strip()
            if not group_name:
                continue
            group_row = ensure_group(group_name, str(raw_row.get("operation", "") or "").strip())
            group_row["active"] = bool(raw_row.get("active", raw_row.get("ativo", True)))
            for machine in list(raw_row.get("machines", []) or []):
                add_machine(str(group_row.get("name", "") or "").strip(), machine_name_from(machine), machine_active_from(machine))

        if sync_legacy and not raw_catalog:
            legacy_postos = [str(value or "").strip() for value in list(data.get("postos_trabalho", []) or []) if str(value or "").strip()]
            for legacy_name in legacy_postos:
                if legacy_name.lower() == "geral":
                    continue
                group_name = self._legacy_workcenter_group_name(legacy_name)
                if not group_name:
                    continue
                normalized_group = str(group_name or "").strip()
                ensure_group(normalized_group)
                if legacy_name.lower() != normalized_group.lower():
                    add_machine(normalized_group, legacy_name)

        cleaned: list[dict[str, Any]] = []
        for row in sorted(groups_by_name.values(), key=lambda item: str(item.get("name", "") or "").lower()):
            machines_by_key: dict[str, dict[str, Any]] = {}
            for machine in list(row.get("machines", []) or []):
                machine_txt = machine_name_from(machine)
                if machine_txt:
                    machines_by_key[machine_txt.lower()] = {"name": machine_txt, "active": machine_active_from(machine)}
            machines = sorted(machines_by_key.values(), key=lambda value: str(value.get("name", "")).lower())
            cleaned.append(
                {
                    "name": str(row.get("name", "") or "").strip(),
                    "operation": str(row.get("operation", "") or "").strip()
                    or self._planning_normalize_operation(row.get("name", ""), default=str(row.get("name", "") or "").strip()),
                    "active": bool(row.get("active", True)),
                    "machines": machines,
                }
            )

        flattened: list[str] = []
        seen_flat: set[str] = set()
        for row in cleaned:
            for value in [str(row.get("name", "") or "").strip(), *[machine_name_from(machine) for machine in list(row.get("machines", []) or [])]]:
                key = str(value or "").strip().lower()
                if not key or key == "geral" or key in seen_flat:
                    continue
                seen_flat.add(key)
                flattened.append(str(value).strip())
        data["workcenter_catalog"] = cleaned
        data["postos_trabalho"] = flattened
        return cleaned

    def workcenter_group_options(self, operation: Any = "", include_general: bool = False) -> list[str]:
        target_operation = self._planning_normalize_operation(operation, default="") if str(operation or "").strip() else ""
        rows = []
        for row in list(self._workcenter_catalog() or []):
            if not bool(row.get("active", True)):
                continue
            group_name = str(row.get("name", "") or "").strip()
            if not group_name:
                continue
            row_operation = self._planning_normalize_operation(row.get("operation", group_name), default=group_name)
            if target_operation and row_operation != target_operation:
                continue
            rows.append(group_name)
        rows = sorted(dict.fromkeys(rows), key=lambda value: value.lower())
        if include_general:
            return ["Geral"] + rows
        return rows

    def workcenter_machine_options(self, group: str = "", operation: Any = "") -> list[str]:
        target_group = str(group or "").strip()
        target_operation = self._planning_normalize_operation(operation, default="") if str(operation or "").strip() else ""
        rows: list[str] = []
        for row in list(self._workcenter_catalog() or []):
            if not bool(row.get("active", True)):
                continue
            group_name = str(row.get("name", "") or "").strip()
            if not group_name:
                continue
            row_operation = self._planning_normalize_operation(row.get("operation", group_name), default=group_name)
            if target_group and group_name.lower() != target_group.lower():
                continue
            if target_operation and row_operation != target_operation:
                continue
            for machine in list(row.get("machines", []) or []):
                machine_name = str((machine or {}).get("name", "") or "").strip() if isinstance(machine, dict) else str(machine or "").strip()
                machine_active = bool((machine or {}).get("active", True)) if isinstance(machine, dict) else True
                if machine_name and machine_active:
                    rows.append(machine_name)
        return sorted(dict.fromkeys(rows), key=lambda value: value.lower())

    def workcenter_resource_options(self, operation: Any = "", include_all: bool = False) -> list[str]:
        target_operation = self._planning_normalize_operation(operation, default="") if str(operation or "").strip() else ""
        resources: list[str] = []
        for row in list(self._workcenter_catalog() or []):
            if not bool(row.get("active", True)):
                continue
            group_name = str(row.get("name", "") or "").strip()
            if not group_name:
                continue
            row_operation = self._planning_normalize_operation(row.get("operation", group_name), default=group_name)
            if target_operation and row_operation != target_operation:
                continue
            machines = [
                str((machine or {}).get("name", "") or "").strip() if isinstance(machine, dict) else str(machine or "").strip()
                for machine in list(row.get("machines", []) or [])
                if (bool((machine or {}).get("active", True)) if isinstance(machine, dict) else True)
                and (str((machine or {}).get("name", "") or "").strip() if isinstance(machine, dict) else str(machine or "").strip())
            ]
            if machines:
                resources.extend(machines)
            else:
                resources.append(group_name)
        ordered = sorted(dict.fromkeys(resources), key=lambda value: value.lower())
        if include_all:
            return ["Todos"] + ordered
        return ordered

    def workcenter_group_for_resource(self, resource: Any, operation: Any = "") -> str:
        resource_txt = self._normalize_workcenter_value(resource)
        if not resource_txt:
            return ""
        for row in list(self._workcenter_catalog() or []):
            group_name = str(row.get("name", "") or "").strip()
            if not group_name:
                continue
            row_operation = self._planning_normalize_operation(row.get("operation", group_name), default=group_name)
            if str(operation or "").strip() and row_operation != self._planning_normalize_operation(operation):
                continue
            if resource_txt.lower() == group_name.lower():
                return group_name
            for machine in list(row.get("machines", []) or []):
                machine_txt = str((machine or {}).get("name", "") or "").strip() if isinstance(machine, dict) else str(machine or "").strip()
                if machine_txt and resource_txt.lower() == machine_txt.lower():
                    return group_name
        return self._legacy_workcenter_group_name(resource_txt)

    def workcenter_default_resource(self, operation: Any = "", preferred: Any = "") -> str:
        preferred_txt = self._normalize_workcenter_value(preferred)
        options = list(self.workcenter_resource_options(operation) or [])
        if preferred_txt and any(preferred_txt.lower() == str(value or "").strip().lower() for value in options):
            return preferred_txt
        return str(options[0] if options else "").strip()

    def _normalize_workcenter_value(self, value: Any) -> str:
        raw = str(value or "").strip()
        if not raw:
            return ""
        catalog = list(self._workcenter_catalog() or [])
        for row in catalog:
            group_name = str(row.get("name", "") or "").strip()
            if group_name and raw.lower() == group_name.lower():
                return group_name
            for machine in list(row.get("machines", []) or []):
                machine_txt = str((machine or {}).get("name", "") or "").strip() if isinstance(machine, dict) else str(machine or "").strip()
                if machine_txt and raw.lower() == machine_txt.lower():
                    return machine_txt
            if group_name and self.desktop_main.norm_text(raw) in self._workcenter_group_aliases(group_name):
                return group_name
        return raw

    def _workcenter_machine_name(self, machine: Any) -> str:
        if isinstance(machine, dict):
            return str(machine.get("name", machine.get("nome", "")) or "").strip()
        return str(machine or "").strip()

    def _workcenter_machine_active(self, machine: Any) -> bool:
        if isinstance(machine, dict):
            return bool(machine.get("active", machine.get("ativo", True)))
        return True

    def _workcenter_machine_entry(self, name: Any, active: bool = True) -> dict[str, Any]:
        return {"name": str(name or "").strip(), "active": bool(active)}

    def available_postos(self) -> list[str]:
        postos = ["Geral"] + self.quote_workcenter_options()
        seen: set[str] = set()
        ordered: list[str] = []
        for posto in postos:
            key = posto.lower()
            if key in seen:
                continue
            seen.add(key)
            ordered.append(posto)
        return ordered

    def operator_posto_options(self) -> list[str]:
        return self.workcenter_group_options(include_general=True)

    def _workcenter_usage_counts(
        self,
        workcenter: str,
        *,
        data: dict[str, Any] | None = None,
        profiles: dict[str, Any] | None = None,
    ) -> dict[str, int]:
        name = str(workcenter or "").strip()
        if not name:
            return {"users": 0, "quotes": 0, "orders": 0, "planning": 0, "operator_map": 0, "total": 0}
        data = data if isinstance(data, dict) else self.ensure_data()
        profiles = profiles if isinstance(profiles, dict) else self._user_profiles()
        norm = self._normalize_workcenter_value(name).lower()
        tracked_names = {norm}
        group_name = self._legacy_workcenter_group_name(name)
        if group_name and group_name.lower() == norm:
            tracked_names.update(str(machine or "").strip().lower() for machine in self.workcenter_machine_options(group_name))
        users = 0
        for user in list(data.get("users", []) or []):
            username = str(user.get("username", "") or "").strip().lower()
            profile = dict(profiles.get(username, {}) or {})
            posto = str(profile.get("posto", "") or user.get("posto", "") or "").strip()
            if self._normalize_workcenter_value(posto).lower() in tracked_names:
                users += 1
        quotes = sum(
            1
            for row in list(data.get("orcamentos", []) or [])
            if self._normalize_workcenter_value(str(row.get("posto_trabalho", "") or "")).lower() in tracked_names
        )
        orders = sum(
            1
            for row in list(data.get("encomendas", []) or [])
            if self._normalize_workcenter_value(
                str(row.get("posto_trabalho", row.get("posto", row.get("maquina", ""))) or "")
            ).lower()
            in tracked_names
        )
        for row in list(data.get("encomendas", []) or []):
            if not isinstance(row, dict):
                continue
            for mat in list(row.get("materiais", []) or []):
                for esp in list(mat.get("espessuras", []) or []):
                    for raw_resource in dict(esp.get("maquinas_operacao", esp.get("recursos_operacao", {})) or {}).values():
                        if self._normalize_workcenter_value(str(raw_resource or "")).lower() in tracked_names:
                            orders += 1
                            break
        planning = 0
        for bucket_name in ("plano", "plano_hist"):
            for row in list(data.get(bucket_name, []) or []):
                posto = str(row.get("maquina", row.get("posto", row.get("posto_trabalho", ""))) or "").strip()
                if self._normalize_workcenter_value(posto).lower() in tracked_names:
                    planning += 1
        operator_map = 0
        for value in dict(data.get("operador_posto_map", {}) or {}).values():
            if self._normalize_workcenter_value(str(value or "")).lower() in tracked_names:
                operator_map += 1
        total = users + quotes + orders + planning + operator_map
        return {
            "users": users,
            "quotes": quotes,
            "orders": orders,
            "planning": planning,
            "operator_map": operator_map,
            "total": total,
        }

    def workcenter_rows(self) -> list[dict[str, Any]]:
        rows: list[dict[str, Any]] = []
        data = self.ensure_data()
        for row in list(self._workcenter_catalog() or []):
            group_name = str(row.get("name", "") or "").strip()
            if not group_name:
                continue
            usage = self._workcenter_usage_counts(group_name, data=data)
            rows.append(
                {
                    "entry_type": "group",
                    "name": group_name,
                    "group": group_name,
                    "operation": self._planning_normalize_operation(row.get("operation", group_name), default=group_name),
                    "kind": "Posto",
                    "protected": group_name == "Geral",
                    "active": bool(row.get("active", True)),
                    "machine_count": len(list(row.get("machines", []) or [])),
                    "users": usage["users"],
                    "quotes": usage["quotes"],
                    "orders": usage["orders"],
                    "planning": usage["planning"],
                    "operator_map": usage["operator_map"],
                    "usage_total": usage["total"],
                }
            )
            for machine_name in list(row.get("machines", []) or []):
                machine_txt = str((machine_name or {}).get("name", "") or "").strip() if isinstance(machine_name, dict) else str(machine_name or "").strip()
                machine_active = bool((machine_name or {}).get("active", True)) if isinstance(machine_name, dict) else True
                if not machine_txt:
                    continue
                machine_usage = self._workcenter_usage_counts(machine_txt, data=data)
                rows.append(
                    {
                        "entry_type": "machine",
                        "name": machine_txt,
                        "group": group_name,
                        "operation": self._planning_normalize_operation(row.get("operation", group_name), default=group_name),
                        "kind": "Maquina",
                        "protected": False,
                        "active": machine_active,
                        "machine_count": 0,
                        "users": machine_usage["users"],
                        "quotes": machine_usage["quotes"],
                        "orders": machine_usage["orders"],
                        "planning": machine_usage["planning"],
                        "operator_map": machine_usage["operator_map"],
                        "usage_total": machine_usage["total"],
                    }
                )
        return rows

    def _replace_removed_resource_references(
        self,
        target_name: str,
        *,
        fallback_resource: str = "",
        operation: Any = "",
        remove_active_planning: bool = False,
    ) -> dict[str, int]:
        data = self.ensure_data()
        profiles = self._user_profiles()
        target_txt = self._normalize_workcenter_value(target_name)
        fallback_txt = self._normalize_workcenter_value(fallback_resource)
        fallback_group = self.workcenter_group_for_resource(fallback_txt, operation) if fallback_txt else ""
        target_key = target_txt.lower()
        stats = {"profiles": 0, "users": 0, "quotes": 0, "orders": 0, "maps": 0, "planning_removed": 0, "planning_updated": 0, "history_updated": 0}
        if not target_key:
            return stats

        def match(value: Any) -> bool:
            return self._normalize_workcenter_value(value).lower() == target_key

        replacement_for_order = fallback_txt or fallback_group
        replacement_for_history = fallback_txt or fallback_group

        for username, profile in list(profiles.items()):
            if match(profile.get("posto", "")):
                profile["posto"] = replacement_for_order
                profiles[username] = profile
                stats["profiles"] += 1
        for user in list(data.get("users", []) or []):
            if match(user.get("posto", "")):
                user["posto"] = replacement_for_order
                stats["users"] += 1
        for row in list(data.get("orcamentos", []) or []):
            if match(row.get("posto_trabalho", "")):
                row["posto_trabalho"] = replacement_for_order
                stats["quotes"] += 1
        for row in list(data.get("encomendas", []) or []):
            changed_order = False
            for key in ("posto_trabalho", "posto", "maquina"):
                if match(row.get(key, "")):
                    row[key] = replacement_for_order
                    changed_order = True
            for mat in list(row.get("materiais", []) or []):
                for esp in list(mat.get("espessuras", []) or []):
                    machine_map = dict(esp.get("maquinas_operacao", esp.get("recursos_operacao", {})) or {})
                    changed_map = False
                    for op_name, raw_value in list(machine_map.items()):
                        if not match(raw_value):
                            continue
                        repl = fallback_txt or self.workcenter_default_resource(op_name, preferred=fallback_group)
                        if repl:
                            machine_map[op_name] = repl
                        else:
                            machine_map.pop(op_name, None)
                        changed_map = True
                    if changed_map:
                        esp["maquinas_operacao"] = machine_map
                        changed_order = True
                        stats["maps"] += 1
            if changed_order:
                stats["orders"] += 1

        active_rows = []
        for row in list(data.get("plano", []) or []):
            if match(self._planning_row_resource(row)):
                if remove_active_planning:
                    stats["planning_removed"] += 1
                    continue
                self._planning_apply_resource_to_row(row, replacement_for_order, row.get("operacao", operation))
                stats["planning_updated"] += 1
            active_rows.append(row)
        data["plano"] = active_rows

        for row in list(data.get("plano_hist", []) or []):
            if not match(self._planning_row_resource(row)):
                continue
            self._planning_apply_resource_to_row(row, replacement_for_history, row.get("operacao", operation))
            stats["history_updated"] += 1

        posto_map = dict(data.get("operador_posto_map", {}) or {})
        for username, posto in list(posto_map.items()):
            if match(posto):
                posto_map[username] = replacement_for_order
        if posto_map:
            data["operador_posto_map"] = posto_map
        self._save_user_profiles(profiles)
        return stats

    def save_workcenter_group(self, name: str, operation: Any = "", current_name: str = "", active: bool = True) -> dict[str, Any]:
        data = self.ensure_data()
        profiles = self._user_profiles()
        new_name = str(name or "").strip()
        current_txt = str(current_name or "").strip()
        if not new_name:
            raise ValueError("Nome do posto obrigatório.")
        if new_name.lower() == "geral":
            raise ValueError("O posto 'Geral' já existe no sistema e não pode ser redefinido.")
        new_operation = self._planning_normalize_operation(operation or new_name, default=new_name)
        catalog = list(self._workcenter_catalog() or [])
        group_names = {str(row.get("name", "") or "").strip().lower() for row in catalog if str(row.get("name", "") or "").strip()}
        machine_names = {
            self._workcenter_machine_name(machine).lower()
            for row in catalog
            for machine in list(row.get("machines", []) or [])
            if self._workcenter_machine_name(machine)
        }
        current_key = current_txt.lower()
        if not current_txt:
            if new_name.lower() in group_names or new_name.lower() in machine_names:
                raise ValueError("Já existe um posto de trabalho com esse nome.")
            catalog.append({"name": new_name, "operation": new_operation, "active": bool(active), "machines": []})
            data["workcenter_catalog"] = catalog
            self._workcenter_catalog()
            self._save(force=True)
            return next(
                (
                    row
                    for row in self.workcenter_rows()
                    if str(row.get("entry_type", "") or "") == "group"
                    and str(row.get("name", "") or "").strip().lower() == new_name.lower()
                ),
                {"name": new_name, "entry_type": "group"},
            )
        current_group = next((row for row in catalog if str(row.get("name", "") or "").strip().lower() == current_key), None)
        if current_group is None:
            raise ValueError("Só é possível editar postos existentes.")
        if new_name.lower() != current_key and (new_name.lower() in group_names or new_name.lower() in machine_names):
            raise ValueError("Já existe um posto de trabalho com esse nome.")
        current_group["name"] = new_name
        current_group["operation"] = new_operation
        current_group["active"] = bool(active)
        for username, profile in list(profiles.items()):
            if self._normalize_workcenter_value(str(profile.get("posto", "") or "")).lower() == current_key:
                profile["posto"] = new_name
                profiles[username] = profile
        for user in list(data.get("users", []) or []):
            if self._normalize_workcenter_value(str(user.get("posto", "") or "")).lower() == current_key:
                user["posto"] = new_name
        for row in list(data.get("orcamentos", []) or []):
            if self._normalize_workcenter_value(str(row.get("posto_trabalho", "") or "")).lower() == current_key:
                row["posto_trabalho"] = new_name
        for row in list(data.get("encomendas", []) or []):
            for key in ("posto_trabalho", "posto", "maquina"):
                if self._normalize_workcenter_value(str(row.get(key, "") or "")).lower() == current_key:
                    row[key] = new_name
            for mat in list(row.get("materiais", []) or []):
                for esp in list(mat.get("espessuras", []) or []):
                    machine_map = dict(esp.get("maquinas_operacao", esp.get("recursos_operacao", {})) or {})
                    changed_map = False
                    for op_name, raw_value in list(machine_map.items()):
                        if self._normalize_workcenter_value(str(raw_value or "")).lower() == current_key:
                            machine_map[op_name] = new_name
                            changed_map = True
                    if changed_map:
                        esp["maquinas_operacao"] = machine_map
        for bucket_name in ("plano", "plano_hist"):
            for row in list(data.get(bucket_name, []) or []):
                for key in ("posto", "posto_trabalho", "maquina"):
                    if self._normalize_workcenter_value(str(row.get(key, "") or "")).lower() == current_key:
                        row[key] = new_name
        posto_map = dict(data.get("operador_posto_map", {}) or {})
        for username, posto in list(posto_map.items()):
            if self._normalize_workcenter_value(str(posto or "")).lower() == current_key:
                posto_map[username] = new_name
        if posto_map:
            data["operador_posto_map"] = posto_map
        data["workcenter_catalog"] = catalog
        self._workcenter_catalog()
        self._save_user_profiles(profiles)
        self._save(force=True)
        return next(
            (
                row
                for row in self.workcenter_rows()
                if str(row.get("entry_type", "") or "") == "group"
                and str(row.get("name", "") or "").strip().lower() == new_name.lower()
            ),
            {"name": new_name, "entry_type": "group"},
        )

    def remove_workcenter_group(self, name: str) -> None:
        data = self.ensure_data()
        target = str(name or "").strip()
        if not target:
            raise ValueError("Posto de trabalho inválido.")
        catalog = list(self._workcenter_catalog() or [])
        current_group = next((row for row in catalog if str(row.get("name", "") or "").strip().lower() == target.lower()), None)
        if current_group is None:
            raise ValueError("Posto de trabalho não encontrado.")
        if list(current_group.get("machines", []) or []):
            raise ValueError("Remove primeiro as máquinas associadas a este posto.")
        usage = self._workcenter_usage_counts(target, data=data)
        if usage["total"] > 0:
            raise ValueError(
                "Não é possível remover este posto porque ainda está em uso "
                f"(utilizadores: {usage['users']}, orçamentos: {usage['quotes']}, encomendas: {usage['orders']}, planeamento: {usage['planning']})."
            )
        data["workcenter_catalog"] = [row for row in catalog if str(row.get("name", "") or "").strip().lower() != target.lower()]
        self._workcenter_catalog()
        self._save(force=True)

    def save_workcenter_machine(self, group_name: str, machine_name: str, current_name: str = "", active: bool = True) -> dict[str, Any]:
        data = self.ensure_data()
        profiles = self._user_profiles()
        parent_group = str(group_name or "").strip()
        new_name = str(machine_name or "").strip()
        current_txt = str(current_name or "").strip()
        if not parent_group:
            raise ValueError("Seleciona o posto de trabalho da máquina.")
        if not new_name:
            raise ValueError("Nome da máquina obrigatório.")
        catalog = list(self._workcenter_catalog() or [])
        group_row = next((row for row in catalog if str(row.get("name", "") or "").strip().lower() == parent_group.lower()), None)
        if group_row is None:
            raise ValueError("Posto de trabalho não encontrado.")
        all_group_names = {str(row.get("name", "") or "").strip().lower() for row in catalog if str(row.get("name", "") or "").strip()}
        all_machine_names = {
            self._workcenter_machine_name(machine).lower()
            for row in catalog
            for machine in list(row.get("machines", []) or [])
            if self._workcenter_machine_name(machine)
        }
        current_key = current_txt.lower()
        if not current_txt:
            if new_name.lower() in all_group_names or new_name.lower() in all_machine_names:
                raise ValueError("Já existe um posto ou máquina com esse nome.")
            group_row.setdefault("machines", []).append(self._workcenter_machine_entry(new_name, active))
            group_row["machines"] = sorted(
                {self._workcenter_machine_name(value).lower(): self._workcenter_machine_entry(self._workcenter_machine_name(value), self._workcenter_machine_active(value)) for value in list(group_row.get("machines", []) or []) if self._workcenter_machine_name(value)}.values(),
                key=lambda value: str(value.get("name", "")).lower(),
            )
            data["workcenter_catalog"] = catalog
            self._workcenter_catalog()
            self._save(force=True)
            return next(
                (
                    row
                    for row in self.workcenter_rows()
                    if str(row.get("entry_type", "") or "") == "machine"
                    and str(row.get("name", "") or "").strip().lower() == new_name.lower()
                ),
                {"name": new_name, "entry_type": "machine", "group": parent_group},
            )
        machine_owner = next(
            (
                row
                for row in catalog
                if any(self._workcenter_machine_name(machine).lower() == current_key for machine in list(row.get("machines", []) or []))
            ),
            None,
        )
        if machine_owner is None:
            raise ValueError("Máquina não encontrada.")
        if new_name.lower() != current_key and (new_name.lower() in all_group_names or new_name.lower() in all_machine_names):
            raise ValueError("Já existe um posto ou máquina com esse nome.")
        machine_owner["machines"] = [
            self._workcenter_machine_entry(new_name, active)
            if self._workcenter_machine_name(machine).lower() == current_key
            else self._workcenter_machine_entry(self._workcenter_machine_name(machine), self._workcenter_machine_active(machine))
            for machine in list(machine_owner.get("machines", []) or [])
            if self._workcenter_machine_name(machine)
        ]
        machine_owner["machines"] = sorted(
            {self._workcenter_machine_name(value).lower(): value for value in machine_owner["machines"]}.values(),
            key=lambda value: str(value.get("name", "")).lower(),
        )
        if machine_owner is not group_row:
            machine_owner["machines"] = [value for value in list(machine_owner.get("machines", []) or []) if self._workcenter_machine_name(value).lower() != new_name.lower()]
            group_row.setdefault("machines", []).append(self._workcenter_machine_entry(new_name, True))
            group_row["machines"] = sorted(
                {self._workcenter_machine_name(value).lower(): self._workcenter_machine_entry(self._workcenter_machine_name(value), self._workcenter_machine_active(value)) for value in list(group_row.get("machines", []) or []) if self._workcenter_machine_name(value)}.values(),
                key=lambda value: str(value.get("name", "")).lower(),
            )
        for username, profile in list(profiles.items()):
            if self._normalize_workcenter_value(str(profile.get("posto", "") or "")).lower() == current_key:
                profile["posto"] = new_name
                profiles[username] = profile
        for user in list(data.get("users", []) or []):
            if self._normalize_workcenter_value(str(user.get("posto", "") or "")).lower() == current_key:
                user["posto"] = new_name
        for row in list(data.get("orcamentos", []) or []):
            if self._normalize_workcenter_value(str(row.get("posto_trabalho", "") or "")).lower() == current_key:
                row["posto_trabalho"] = new_name
        for row in list(data.get("encomendas", []) or []):
            for key in ("posto_trabalho", "posto", "maquina"):
                if self._normalize_workcenter_value(str(row.get(key, "") or "")).lower() == current_key:
                    row[key] = new_name
            for mat in list(row.get("materiais", []) or []):
                for esp in list(mat.get("espessuras", []) or []):
                    machine_map = dict(esp.get("maquinas_operacao", esp.get("recursos_operacao", {})) or {})
                    changed_map = False
                    for op_name, raw_value in list(machine_map.items()):
                        if self._normalize_workcenter_value(str(raw_value or "")).lower() == current_key:
                            machine_map[op_name] = new_name
                            changed_map = True
                    if changed_map:
                        esp["maquinas_operacao"] = machine_map
        for bucket_name in ("plano", "plano_hist"):
            for row in list(data.get(bucket_name, []) or []):
                for key in ("posto", "posto_trabalho", "maquina"):
                    if self._normalize_workcenter_value(str(row.get(key, "") or "")).lower() == current_key:
                        row[key] = new_name
        posto_map = dict(data.get("operador_posto_map", {}) or {})
        for username, posto in list(posto_map.items()):
            if self._normalize_workcenter_value(str(posto or "")).lower() == current_key:
                posto_map[username] = new_name
        if posto_map:
            data["operador_posto_map"] = posto_map
        data["workcenter_catalog"] = catalog
        self._workcenter_catalog()
        self._save_user_profiles(profiles)
        self._save(force=True)
        return next(
            (
                row
                for row in self.workcenter_rows()
                if str(row.get("entry_type", "") or "") == "machine"
                and str(row.get("name", "") or "").strip().lower() == new_name.lower()
            ),
            {"name": new_name, "entry_type": "machine", "group": parent_group},
        )

    def remove_workcenter_machine(self, machine_name: str) -> None:
        data = self.ensure_data()
        target = str(machine_name or "").strip()
        if not target:
            raise ValueError("Máquina inválida.")
        catalog = list(self._workcenter_catalog() or [])
        owner_row = next(
            (
                row
                for row in catalog
                if any(self._workcenter_machine_name(machine).lower() == target.lower() for machine in list(row.get("machines", []) or []))
            ),
            None,
        )
        if owner_row is None:
            raise ValueError("Máquina não encontrada.")
        usage = self._workcenter_usage_counts(target, data=data)
        if usage["total"] > 0:
            raise ValueError(
                "Não é possível remover esta máquina porque ainda está em uso "
                f"(utilizadores: {usage['users']}, orçamentos: {usage['quotes']}, encomendas: {usage['orders']}, planeamento: {usage['planning']})."
            )
        owner_row["machines"] = [value for value in list(owner_row.get("machines", []) or []) if self._workcenter_machine_name(value).lower() != target.lower()]
        data["workcenter_catalog"] = catalog
        self._workcenter_catalog()
        self._save(force=True)

    def _order_workcenter(self, enc_or_numero: dict[str, Any] | str | None = None) -> str:
        enc: dict[str, Any] | None
        if isinstance(enc_or_numero, dict):
            enc = enc_or_numero
        else:
            enc = self.get_encomenda_by_numero(str(enc_or_numero or "").strip()) if str(enc_or_numero or "").strip() else None
        if not isinstance(enc, dict):
            return ""
        for key in ("posto_trabalho", "posto", "maquina"):
            value = str(enc.get(key, "") or "").strip()
            if value:
                return self._normalize_workcenter_value(value)
        return ""

    def _order_esp_machine_map(self, esp_obj: dict[str, Any] | None) -> dict[str, str]:
        if not isinstance(esp_obj, dict):
            return {}
        raw_map = dict(esp_obj.get("maquinas_operacao", esp_obj.get("recursos_operacao", {})) or {})
        cleaned: dict[str, str] = {}
        for raw_op, raw_resource in raw_map.items():
            op_txt = self._planning_normalize_operation(raw_op, default="")
            resource_txt = self._sanitize_operation_resource(op_txt, raw_resource)
            if not op_txt or not resource_txt:
                continue
            cleaned[op_txt] = resource_txt
        return cleaned

    def _sanitize_operation_resource(self, operation: Any = "", resource: Any = "") -> str:
        op_txt = self._planning_normalize_operation(operation, default="") if str(operation or "").strip() else ""
        resource_txt = self._normalize_workcenter_value(resource)
        if not op_txt:
            return resource_txt
        available_resources = [
            str(value or "").strip()
            for value in list(self.workcenter_resource_options(op_txt) or [])
            if str(value or "").strip()
        ]
        if not available_resources:
            return resource_txt
        if resource_txt and any(resource_txt.lower() == value.lower() for value in available_resources):
            return next((value for value in available_resources if value.lower() == resource_txt.lower()), resource_txt)
        if len(available_resources) == 1:
            return str(available_resources[0] or "").strip()
        return ""

    def _order_operation_resource(
        self,
        enc_or_numero: dict[str, Any] | str | None,
        material: str = "",
        espessura: str = "",
        operation: Any = "",
    ) -> str:
        op_txt = self._planning_normalize_operation(operation, default="") if str(operation or "").strip() else ""
        enc = enc_or_numero if isinstance(enc_or_numero, dict) else self.get_encomenda_by_numero(str(enc_or_numero or "").strip())
        if not isinstance(enc, dict):
            return ""
        esp_obj = self._order_find_espessura(enc, material, espessura) if str(material or "").strip() and str(espessura or "").strip() else None
        machine_map = self._order_esp_machine_map(esp_obj)
        if op_txt and machine_map.get(op_txt):
            return str(machine_map.get(op_txt) or "").strip()
        if op_txt:
            return self.workcenter_default_resource(op_txt, preferred=self._order_workcenter(enc))
        return self._order_workcenter(enc)
