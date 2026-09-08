"""Transport tariff use cases without the shared ERP snapshot."""
from copy import deepcopy
from dataclasses import dataclass
from typing import Any, Callable, Protocol

class TariffRepository(Protocol):
    def rows(self) -> list[dict[str, Any]]: ...
    def replace(self, rows: list[dict[str, Any]], *, expected: list[dict[str, Any]]) -> None: ...

@dataclass(frozen=True)
class TariffRules:
    parse_float: Callable
    norm_text: Callable
    normalize_supplier: Callable

class TariffService:
    def __init__(self, repository: TariffRepository, rules: TariffRules):
        self.repository = repository
        self.rules = rules

    def defaults(self) -> dict[str, Any]:
        return {
            "id": "",
            "transportadora_id": "",
            "transportadora_nome": "",
            "zona": "",
            "valor_base": 0.0,
            "valor_por_palete": 0.0,
            "valor_por_kg": 0.0,
            "valor_por_m3": 0.0,
            "custo_minimo": 0.0,
            "ativo": True,
            "observacoes": "",
        }


    def signature(self, row: dict[str, Any] | None) -> str:
        row = dict(row or {})
        carrier_parts = [
            part
            for part in [
                str(row.get("transportadora_id", "") or "").strip(),
                str(row.get("transportadora_nome", "") or "").strip(),
            ]
            if part
        ]
        carrier = " - ".join(carrier_parts).strip(" -") or "Sem transportadora"
        zone = str(row.get("zona", "") or "").strip() or "Sem zona"
        return f"{carrier} | {zone}"


    def next_id(self) -> int:
        highest = 0
        for row in self.repository.rows():
            highest = max(highest, int(self.rules.parse_float((row or {}).get("id", 0), 0) or 0))
        return highest + 1


    def match(self, transportadora_id: Any = "", transportadora_nome: Any = "", zona: Any = "") -> dict[str, Any] | None:
        zone_norm = self.rules.norm_text(zona)
        if not zone_norm:
            return None
        supplier_id = str(transportadora_id or "").strip()
        supplier_name_norm = self.rules.norm_text(transportadora_nome)
        best: tuple[int, int, dict[str, Any]] | None = None
        for raw in self.repository.rows():
            if not isinstance(raw, dict) or not bool(raw.get("ativo", True)):
                continue
            if self.rules.norm_text(raw.get("zona", "")) != zone_norm:
                continue
            row_supplier_id = str(raw.get("transportadora_id", "") or "").strip()
            row_supplier_name_norm = self.rules.norm_text(raw.get("transportadora_nome", ""))
            score = 0
            if supplier_id and row_supplier_id and row_supplier_id == supplier_id:
                score = 3
            elif supplier_name_norm and row_supplier_name_norm and row_supplier_name_norm == supplier_name_norm:
                score = 2
            elif not row_supplier_id and not row_supplier_name_norm:
                score = 1
            if score <= 0:
                continue
            row_id = int(self.rules.parse_float(raw.get("id", 0), 0) or 0)
            candidate = (score, -row_id, raw)
            if best is None or candidate[:2] > best[:2]:
                best = candidate
        return deepcopy(best[2]) if best else None


    def cost(self, row: dict[str, Any] | None, paletes: Any = 0, peso_bruto_kg: Any = 0, volume_m3: Any = 0) -> float:
        row = dict(row or {})
        base = round(self.rules.parse_float(row.get("valor_base", 0), 0), 2)
        per_pal = round(self.rules.parse_float(row.get("valor_por_palete", 0), 0), 2)
        per_kg = round(self.rules.parse_float(row.get("valor_por_kg", 0), 0), 4)
        per_m3 = round(self.rules.parse_float(row.get("valor_por_m3", 0), 0), 2)
        minimum = round(self.rules.parse_float(row.get("custo_minimo", 0), 0), 2)
        pal = max(0.0, round(self.rules.parse_float(paletes, 0), 2))
        peso = max(0.0, round(self.rules.parse_float(peso_bruto_kg, 0), 2))
        volume = max(0.0, round(self.rules.parse_float(volume_m3, 0), 3))
        total = round(base + (pal * per_pal) + (peso * per_kg) + (volume * per_m3), 2)
        if minimum > 0:
            total = max(total, minimum)
        return round(total, 2)


    def suggestion(
        self,
        transportadora_id: Any = "",
        transportadora_nome: Any = "",
        zona: Any = "",
        paletes: Any = 0,
        peso_bruto_kg: Any = 0,
        volume_m3: Any = 0,
    ) -> dict[str, Any]:
        tariff = self.match(transportadora_id, transportadora_nome, zona)
        if tariff is None:
            return {
                "tarifario_id": "",
                "tarifario_label": "",
                "custo_sugerido": 0.0,
            }
        return {
            "tarifario_id": tariff.get("id", ""),
            "tarifario_label": self.signature(tariff),
            "custo_sugerido": self.cost(tariff, paletes, peso_bruto_kg, volume_m3),
        }


    def rows(self, filter_text: str = "") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        rows: list[dict[str, Any]] = []
        for raw in self.repository.rows():
            if not isinstance(raw, dict):
                continue
            row = {
                "id": int(self.rules.parse_float(raw.get("id", 0), 0) or 0),
                "transportadora_id": str(raw.get("transportadora_id", "") or "").strip(),
                "transportadora_nome": str(raw.get("transportadora_nome", "") or "").strip(),
                "zona": str(raw.get("zona", "") or "").strip(),
                "valor_base": round(self.rules.parse_float(raw.get("valor_base", 0), 0), 2),
                "valor_por_palete": round(self.rules.parse_float(raw.get("valor_por_palete", 0), 0), 2),
                "valor_por_kg": round(self.rules.parse_float(raw.get("valor_por_kg", 0), 0), 4),
                "valor_por_m3": round(self.rules.parse_float(raw.get("valor_por_m3", 0), 0), 2),
                "custo_minimo": round(self.rules.parse_float(raw.get("custo_minimo", 0), 0), 2),
                "ativo": bool(raw.get("ativo", True)),
                "observacoes": str(raw.get("observacoes", "") or "").strip(),
                "label": self.signature(raw),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: (self.rules.norm_text(item.get("transportadora_nome", "")), self.rules.norm_text(item.get("zona", "")), item.get("id", 0)))
        return rows


    def save(self, payload: dict[str, Any]) -> dict[str, Any]:
        payload = deepcopy(payload)
        tariff_id = int(self.rules.parse_float(payload.get("id", 0), 0) or 0)
        zone = str(payload.get("zona", "") or "").strip()
        if not zone:
            raise ValueError("Zona obrigatoria no tarifario.")
        transportadora_id, transportadora_nome, _contact = self.rules.normalize_supplier(
            payload.get("transportadora_id", ""),
            payload.get("transportadora_nome", ""),
        )
        original = self.repository.rows()
        rows = deepcopy(original)
        zone_norm = self.rules.norm_text(zone)
        for row in rows:
            if not isinstance(row, dict):
                continue
            if int(self.rules.parse_float(row.get("id", 0), 0) or 0) == tariff_id:
                continue
            same_zone = self.rules.norm_text(row.get("zona", "")) == zone_norm
            same_supplier = (
                str(row.get("transportadora_id", "") or "").strip() == transportadora_id
                and self.rules.norm_text(row.get("transportadora_nome", "")) == self.rules.norm_text(transportadora_nome)
            )
            if same_zone and same_supplier:
                raise ValueError("Ja existe um tarifario para essa transportadora e zona.")
        target = next((row for row in rows if int(self.rules.parse_float((row or {}).get("id", 0), 0) or 0) == tariff_id), None) if tariff_id > 0 else None
        if target is None:
            target = self.defaults()
            target["id"] = self.next_id()
            rows.append(target)
        target["transportadora_id"] = transportadora_id
        target["transportadora_nome"] = transportadora_nome
        target["zona"] = zone
        target["valor_base"] = round(self.rules.parse_float(payload.get("valor_base", target.get("valor_base", 0)), 0), 2)
        target["valor_por_palete"] = round(self.rules.parse_float(payload.get("valor_por_palete", target.get("valor_por_palete", 0)), 0), 2)
        target["valor_por_kg"] = round(self.rules.parse_float(payload.get("valor_por_kg", target.get("valor_por_kg", 0)), 0), 4)
        target["valor_por_m3"] = round(self.rules.parse_float(payload.get("valor_por_m3", target.get("valor_por_m3", 0)), 0), 2)
        target["custo_minimo"] = round(self.rules.parse_float(payload.get("custo_minimo", target.get("custo_minimo", 0)), 0), 2)
        target["ativo"] = bool(payload.get("ativo", target.get("ativo", True)))
        target["observacoes"] = str(payload.get("observacoes", target.get("observacoes", "")) or "").strip()
        self.repository.replace(rows, expected=original)
        return deepcopy(target)


    def remove(self, tariff_id: Any) -> None:
        target_id = int(self.rules.parse_float(tariff_id, 0))
        rows = self.repository.rows()
        filtered = [row for row in rows if int(self.rules.parse_float((row or {}).get("id", 0), 0) or 0) != target_id]
        if len(filtered) == len(rows):
            raise ValueError("Tarifario nao encontrado.")
        self.repository.replace(filtered, expected=rows)


