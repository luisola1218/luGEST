"""Operation costing rules, independent of widgets, storage and the legacy app.

Name/number parsing is supplied explicitly to preserve the existing company's
normalization rules. These callbacks must be pure: no database or UI access.
"""
from __future__ import annotations

from typing import Any, Callable, Iterable


class OperationCostingEngine:
    def __init__(self, *, available_operations: Iterable[str],
                 normalize_operation: Callable[[Any], str],
                 parse_number: Callable[[Any, Any], float],
                 parse_operations: Callable[[Any], list[str]]) -> None:
        self.available_operations = tuple(available_operations)
        self.normalize_operation = normalize_operation
        self.parse_number = parse_number
        self.parse_operations = parse_operations

    @staticmethod
    def default_settings() -> dict[str, Any]:
        def _profile(
            pricing_mode: str,
            driver_label: str,
            *,
            default_units: float = 1.0,
            setup_min: float = 0.0,
            unit_time_min: float = 0.0,
            hour_rate_eur: float = 0.0,
            fixed_unit_eur: float = 0.0,
            min_unit_eur: float = 0.0,
            extra_unit_eur: float = 0.0,
            requires_driver_input: bool = False,
            note: str = "",
        ) -> dict[str, Any]:
            return {
                "pricing_mode": pricing_mode,
                "driver_label": driver_label,
                "default_units": float(default_units),
                "setup_min": float(setup_min),
                "unit_time_min": float(unit_time_min),
                "hour_rate_eur": float(hour_rate_eur),
                "fixed_unit_eur": float(fixed_unit_eur),
                "min_unit_eur": float(min_unit_eur),
                "extra_unit_eur": float(extra_unit_eur),
                "requires_driver_input": bool(requires_driver_input),
                "note": str(note or "").strip(),
            }

        return {
            "active_profile": "Base",
            "profiles": {
                "Base": {
                    "Corte Laser": _profile(
                        "manual",
                        "Programa",
                        note="Usar o motor de corte laser ou detalhe manual quando esta operacao aparece combinada com outras.",
                    ),
                    "Quinagem": _profile("per_feature", "Dobras/peca", setup_min=6.0, unit_time_min=0.35, hour_rate_eur=42.0, requires_driver_input=True),
                    "Roscagem": _profile("per_feature", "Roscas/peca", setup_min=4.0, unit_time_min=0.2, hour_rate_eur=38.0, requires_driver_input=True),
                    "Serralharia": _profile("per_piece", "Operacoes/peca", setup_min=10.0, unit_time_min=4.0, hour_rate_eur=40.0, requires_driver_input=True),
                    "Lacagem": _profile("per_area_m2", "m2/peca", default_units=1.0, setup_min=6.0, fixed_unit_eur=14.0),
                    "Maquinacao": _profile("per_feature", "Operacoes/peca", setup_min=12.0, unit_time_min=3.0, hour_rate_eur=55.0, requires_driver_input=True),
                    "Soldadura": _profile("per_feature", "Pontos cordoes/peca", setup_min=8.0, unit_time_min=1.5, hour_rate_eur=42.0, requires_driver_input=True),
                    "Montagem": _profile("per_piece", "Operacoes/peca", setup_min=5.0, unit_time_min=2.5, hour_rate_eur=30.0),
                    "Embalamento": _profile("per_piece", "Volumes/peca", setup_min=2.0, unit_time_min=0.8, hour_rate_eur=24.0, min_unit_eur=0.25),
                }
            },
        }

    def merge_settings(self, stored: dict[str, Any] | None = None) -> dict[str, Any]:
        base = self.default_settings()
        raw = dict(stored or {})
        merged_profiles: dict[str, dict[str, Any]] = {}
        stored_profiles = dict(raw.get("profiles", {}) or {})
        active_profile = str(raw.get("active_profile", base.get("active_profile", "Base")) or "Base").strip() or "Base"

        profile_names = set(stored_profiles.keys()) | set(base.get("profiles", {}).keys()) | {active_profile}
        for profile_name in sorted(profile_names):
            base_profile = dict(dict(base.get("profiles", {}) or {}).get(profile_name, {}) or {})
            current_profile = dict(stored_profiles.get(profile_name, {}) or {})
            merged_profile: dict[str, Any] = {}
            operation_names = set(base_profile.keys()) | set(current_profile.keys()) | set(self.available_operations)
            for operation_name in sorted(operation_names):
                raw_name = self.normalize_operation(operation_name) or str(operation_name or "").strip()
                if not raw_name:
                    continue
                template = dict(base_profile.get(raw_name, {}) or {})
                current = dict(current_profile.get(raw_name, current_profile.get(operation_name, {})) or {})
                merged_profile[raw_name] = {
                    "pricing_mode": str(current.get("pricing_mode", template.get("pricing_mode", "manual")) or "manual").strip() or "manual",
                    "driver_label": str(current.get("driver_label", template.get("driver_label", "Qtd./peca")) or "Qtd./peca").strip(),
                    "default_units": round(self.parse_number(current.get("default_units", template.get("default_units", 1)), 1), 4),
                    "setup_min": round(self.parse_number(current.get("setup_min", template.get("setup_min", 0)), 0), 4),
                    "unit_time_min": round(self.parse_number(current.get("unit_time_min", template.get("unit_time_min", 0)), 0), 4),
                    "hour_rate_eur": round(self.parse_number(current.get("hour_rate_eur", template.get("hour_rate_eur", 0)), 0), 4),
                    "fixed_unit_eur": round(self.parse_number(current.get("fixed_unit_eur", template.get("fixed_unit_eur", 0)), 0), 4),
                    "min_unit_eur": round(self.parse_number(current.get("min_unit_eur", template.get("min_unit_eur", 0)), 0), 4),
                    "extra_unit_eur": round(self.parse_number(current.get("extra_unit_eur", template.get("extra_unit_eur", 0)), 0), 4),
                    "requires_driver_input": bool(current.get("requires_driver_input", template.get("requires_driver_input", False))),
                    "note": str(current.get("note", template.get("note", "")) or "").strip(),
                }
            merged_profiles[profile_name] = merged_profile

        if active_profile not in merged_profiles:
            active_profile = next(iter(merged_profiles.keys()), "Base")
        return {"active_profile": active_profile, "profiles": merged_profiles}

    def estimate(self, payload: dict[str, Any], settings: dict[str, Any]) -> dict[str, Any]:
        row = dict(payload or {})
        active_profile = str(settings.get("active_profile", "Base") or "Base").strip() or "Base"
        profile_map = dict(dict(settings.get("profiles", {}) or {}).get(active_profile, {}) or {})
        raw_costing_operations = row.get("costing_operations")
        if isinstance(raw_costing_operations, (list, tuple, set)):
            operations = []
            for raw_name in list(raw_costing_operations):
                normalized = str(self.normalize_operation(raw_name) or raw_name or "").strip()
                if normalized and normalized not in operations:
                    operations.append(normalized)
        else:
            operations = [
                str(op or "").strip()
                for op in list(self.parse_operations(row.get("operacao", "")) or [])
                if str(op or "").strip()
            ]
        qty = max(1.0, self.parse_number(row.get("qtd", 0), 1))
        raw_detail = {
            str(self.normalize_operation(item.get("nome", "")) or item.get("nome", "") or "").strip(): dict(item or {})
            for item in list(row.get("operacoes_detalhe", []) or [])
            if isinstance(item, dict) and str(item.get("nome", "") or "").strip()
        }
        area_m2 = max(0.0, self.parse_number(row.get("area_m2", row.get("net_area_m2", 0)), 0))
        detail_rows: list[dict[str, Any]] = []
        complete = True
        partial = False
        total_time_unit = 0.0
        total_cost_unit = 0.0
        for index, op_name in enumerate(operations, start=1):
            profile = dict(profile_map.get(op_name, {}) or {})
            existing = dict(raw_detail.get(op_name, {}) or {})
            pricing_mode = str(existing.get("pricing_mode", profile.get("pricing_mode", "manual")) or "manual").strip() or "manual"
            if pricing_mode == "per_area_m2" and area_m2 > 0:
                default_units = area_m2
            else:
                default_units = profile.get("default_units", 1)
            driver_units = round(self.parse_number(existing.get("driver_units", default_units), default_units), 4)
            driver_label = str(existing.get("driver_label", profile.get("driver_label", "Qtd./peca")) or "Qtd./peca").strip()
            setup_min = round(self.parse_number(existing.get("setup_min", profile.get("setup_min", 0)), 0), 4)
            unit_time_base_min = round(self.parse_number(existing.get("unit_time_base_min", existing.get("unit_time_min", profile.get("unit_time_min", 0))), 0), 4)
            hour_rate_eur = round(self.parse_number(existing.get("hour_rate_eur", profile.get("hour_rate_eur", 0)), 0), 4)
            fixed_unit_eur = round(self.parse_number(existing.get("fixed_unit_eur", profile.get("fixed_unit_eur", 0)), 0), 4)
            min_unit_eur = round(self.parse_number(existing.get("min_unit_eur", profile.get("min_unit_eur", 0)), 0), 4)
            extra_unit_eur = round(self.parse_number(existing.get("extra_unit_eur", profile.get("extra_unit_eur", 0)), 0), 4)
            requires_driver_input = bool(existing.get("requires_driver_input", profile.get("requires_driver_input", False)))
            driver_units_confirmed = bool(existing.get("driver_units_confirmed", False))
            if not driver_units_confirmed and pricing_mode != "manual":
                driver_units_confirmed = any(existing.get(field_name) not in (None, "") for field_name in ("tempo_unit_min", "custo_unit_eur"))
            note = str(existing.get("note", profile.get("note", "")) or "").strip()
            manual_time = existing.get("tempo_unit_min")
            manual_cost = existing.get("custo_unit_eur")
            manual_values_confirmed = bool(existing.get("manual_values_confirmed", False))
            if not manual_values_confirmed and (manual_time not in (None, "") or manual_cost not in (None, "")):
                manual_values_confirmed = abs(self.parse_number(manual_time, 0)) > 0.000001 or abs(self.parse_number(manual_cost, 0)) > 0.000001
            computed_time_unit: float | None = None
            computed_cost_unit: float | None = None
            resolved = False
            profile_source = "manual"
            if pricing_mode == "manual":
                if manual_values_confirmed and manual_time not in (None, ""):
                    computed_time_unit = round(self.parse_number(manual_time, 0), 4)
                if manual_values_confirmed and manual_cost not in (None, ""):
                    computed_cost_unit = round(self.parse_number(manual_cost, 0), 4)
                resolved = computed_time_unit is not None and computed_cost_unit is not None
            else:
                profile_source = active_profile
                if requires_driver_input and not driver_units_confirmed:
                    resolved = False
                else:
                    computed_time_unit = round((setup_min / qty) + (driver_units * unit_time_base_min), 4)
                    computed_cost_unit = round(max(min_unit_eur, extra_unit_eur + (driver_units * fixed_unit_eur) + ((hour_rate_eur * computed_time_unit) / 60.0)), 4)
                    resolved = True
            if resolved:
                partial = True
                total_time_unit += float(computed_time_unit or 0)
                total_cost_unit += float(computed_cost_unit or 0)
            else:
                complete = False
                if computed_time_unit is not None or computed_cost_unit is not None:
                    partial = True
            detail_rows.append(
                {
                    "seq": index,
                    "nome": op_name,
                    "pricing_mode": pricing_mode,
                    "profile_name": active_profile,
                    "profile_source": profile_source,
                    "driver_label": driver_label,
                    "driver_units": driver_units,
                    "setup_min": setup_min,
                    "unit_time_base_min": unit_time_base_min,
                    "hour_rate_eur": hour_rate_eur,
                    "fixed_unit_eur": fixed_unit_eur,
                    "min_unit_eur": min_unit_eur,
                    "extra_unit_eur": extra_unit_eur,
                    "requires_driver_input": requires_driver_input,
                    "driver_units_confirmed": driver_units_confirmed,
                    "manual_values_confirmed": manual_values_confirmed,
                    "missing_driver_input": bool(requires_driver_input and pricing_mode != "manual" and not driver_units_confirmed),
                    "tempo_unit_min": computed_time_unit,
                    "custo_unit_eur": computed_cost_unit,
                    "resolved": resolved,
                    "tem_detalhe": resolved,
                    "note": note,
                }
            )
        if detail_rows and complete:
            costing_mode = "detailed"
        elif partial:
            costing_mode = "partial_detail"
        elif len(operations) <= 1:
            costing_mode = "single_operation_total"
        else:
            costing_mode = "aggregate_pending"
        return {
            "active_profile": active_profile,
            "operations": detail_rows,
            "summary": {
                "complete": complete and bool(detail_rows),
                "partial": partial and not (complete and bool(detail_rows)),
                "tempo_unit_total_min": round(total_time_unit, 4),
                "custo_unit_total_eur": round(total_cost_unit, 4),
                "tempo_total_min": round(total_time_unit * qty, 2),
                "custo_total_eur": round(total_cost_unit * qty, 2),
                "costing_mode": costing_mode,
                "qtd": round(qty, 2),
            },
        }
