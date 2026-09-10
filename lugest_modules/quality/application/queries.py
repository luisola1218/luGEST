"""Quality dashboard, integrity checks and entity selection read models."""
from dataclasses import dataclass
from typing import Any, Callable
from .stock_policy import status_code

@dataclass
class QualityCatalogs:
    nonconformities: list
    documents: list
    audit: list
    orders: list
    materials: list
    products: list
    suppliers: list
    clients: list
    plan: list

@dataclass(frozen=True)
class QualityQueryRules:
    parse_float: Callable
    now_iso: Callable
    today: Callable
    path_exists: Callable
    material_blocked: Callable

class QualityQueries:
    def __init__(self, repository, rules: QualityQueryRules, movement_rows):
        self.repository, self.rules, self.movement_rows = repository, rules, movement_rows

    def summary(self) -> dict[str, Any]:
        data = self.repository.load()
        ncs = [row for row in data.nonconformities if isinstance(row, dict)]
        docs = [row for row in data.documents if isinstance(row, dict)]
        health = self.health()
        open_ncs = [row for row in ncs if str(row.get("estado", "") or "Aberta").strip().lower() not in {"fechada", "cancelada"}]
        overdue = 0
        today = self.rules.today()
        for row in open_ncs:
            due = str(row.get("prazo", "") or "").strip()[:10]
            if due and due < today:
                overdue += 1
        supplier_ncs = [
            row
            for row in open_ncs
            if str(row.get("tipo", "") or "").strip().casefold() == "fornecedor"
            or str(row.get("fornecedor_id", "") or row.get("fornecedor_nome", "") or "").strip()
        ]
        blocked_materials = [
            row
            for row in self.movement_rows()
            if str(row.get("entity_type", "") or "") == "Material"
            and (
                self.rules.parse_float(row.get("pending_qty", 0), 0) > 0
                or not status_code(row.get("status", "")) == "APROVADO"
            )
        ]
        return {
            "open_nc": len(open_ncs),
            "overdue_nc": overdue,
            "supplier_nc": len(supplier_ncs),
            "blocked_materials": len(blocked_materials),
            "documents": len(docs),
            "audit_events": len(data.audit),
            "quality_issues": len(list(health.get("issues", []) or [])),
            "updated_at": str(self.rules.now_iso() or "").strip(),
        }

    def health(self) -> dict[str, Any]:
        data = self.repository.load()
        known: dict[str, set[str]] = {
            "Encomenda": {str(row.get("numero", "") or "").strip() for row in data.orders if isinstance(row, dict)},
            "Material": {str(row.get("id", "") or "").strip() for row in data.materials if isinstance(row, dict)},
            "Produto": {str(row.get("codigo", "") or "").strip() for row in data.products if isinstance(row, dict)},
            "Fornecedor": {
                str(row.get("id", "") or row.get("nome", "") or "").strip()
                for row in data.suppliers
                if isinstance(row, dict)
            },
            "Cliente": {str(row.get("codigo", "") or row.get("nome", "") or "").strip() for row in data.clients if isinstance(row, dict)},
            "Documento": {str(row.get("id", "") or row.get("titulo", "") or "").strip() for row in data.documents if isinstance(row, dict)},
        }
        reception_keys = {(row["entity_type"], row["entity_id"]) for row in self.movement_rows()}
        issues: list[dict[str, str]] = []
        for row in data.nonconformities:
            if not isinstance(row, dict):
                continue
            entity_type = str(row.get("entidade_tipo", "") or "").strip()
            entity_id = str(row.get("entidade_id", "") or "").strip()
            if entity_type and entity_type != "Livre" and entity_id:
                if entity_type in {"Material", "Produto"} and (entity_type, entity_id) not in reception_keys and str(row.get("origem", "") or "").strip().casefold() == "receção fornecedor":
                    continue
                known_ids = known.get(entity_type)
                if known_ids is not None and entity_id not in known_ids:
                    issues.append({"tipo": "NC", "id": str(row.get("id", "") or ""), "problema": f"Ligacao inexistente: {entity_type} {entity_id}"})
        for row in data.documents:
            if not isinstance(row, dict):
                continue
            path_txt = str(row.get("caminho", "") or "").strip()
            if path_txt and not self.rules.path_exists(path_txt):
                issues.append({"tipo": "Documento", "id": str(row.get("id", "") or ""), "problema": f"Ficheiro nao encontrado: {path_txt}"})
        open_nc_ids = {
            str(row.get("id", "") or "").strip()
            for row in data.nonconformities
            if isinstance(row, dict) and str(row.get("estado", "") or "Aberta").strip().lower() not in {"fechada", "cancelada"}
        }
        for material in data.materials:
            if not isinstance(material, dict) or not self.rules.material_blocked(material):
                continue
            material_id = str(material.get("id", "") or "").strip()
            if ("Material", material_id) not in reception_keys:
                continue
            status_norm = str(material.get("quality_status", "") or material.get("inspection_status", "") or "").strip().casefold()
            if "inspe" in status_norm and not any(token in status_norm for token in ("bloque", "reclam", "rejeit")):
                continue
            nc_id = str(material.get("quality_nc_id", "") or material.get("supplier_claim_id", "") or "").strip()
            if not nc_id:
                issues.append({"tipo": "Material", "id": str(material.get("id", "") or ""), "problema": "Material bloqueado sem NC/reclamacao ligada."})
            elif nc_id not in open_nc_ids:
                issues.append({"tipo": "Material", "id": str(material.get("id", "") or ""), "problema": f"Material bloqueado com NC inexistente/fechada: {nc_id}"})
        return {"issues": issues, "ok": not issues}

    def link_options(self) -> dict[str, list[dict[str, str]]]:
        data = self.repository.load()
        options: dict[str, list[dict[str, str]]] = {
            "Livre": [{"id": "", "label": ""}],
            "OPP": [],
            "Encomenda": [],
            "Material": [],
            "Produto": [],
            "Fornecedor": [],
            "Cliente": [],
            "Documento": [],
        }
        for row in data.plan:
            if not isinstance(row, dict):
                continue
            opp = str(row.get("opp", "") or row.get("OPP", "") or "").strip()
            if opp:
                options["OPP"].append({"id": opp, "label": f"{opp} | {str(row.get('encomenda', '') or '').strip()}".strip(" |")})
        for enc in data.orders:
            numero = str((enc or {}).get("numero", "") or "").strip()
            if numero:
                options["Encomenda"].append({"id": numero, "label": f"{numero} | {str((enc or {}).get('cliente', '') or '').strip()}"})
        for row in data.materials:
            if not isinstance(row, dict):
                continue
            material_id = str(row.get("id", "") or "").strip()
            if material_id:
                options["Material"].append(
                    {
                        "id": material_id,
                        "label": " | ".join(
                            part
                            for part in (
                                material_id,
                                str(row.get("material", "") or "").strip(),
                                str(row.get("espessura", "") or "").strip(),
                                str(row.get("formato", "") or "").strip(),
                            )
                            if part
                        ),
                    }
                )
        for row in data.products:
            if not isinstance(row, dict):
                continue
            code = str(row.get("codigo", "") or "").strip()
            if code:
                options["Produto"].append({"id": code, "label": f"{code} | {str(row.get('descricao', '') or '').strip()}".strip(" |")})
        for row in data.suppliers:
            supplier_id = str((row or {}).get("id", "") or "").strip()
            name = str((row or {}).get("nome", "") or "").strip()
            if supplier_id or name:
                options["Fornecedor"].append({"id": supplier_id or name, "label": f"{supplier_id} | {name}".strip(" |")})
        for row in data.clients:
            code = str((row or {}).get("codigo", "") or "").strip()
            name = str((row or {}).get("nome", "") or "").strip()
            if code or name:
                options["Cliente"].append({"id": code or name, "label": f"{code} | {name}".strip(" |")})
        for row in data.documents:
            doc_id = str((row or {}).get("id", "") or "").strip()
            title = str((row or {}).get("titulo", "") or "").strip()
            if doc_id or title:
                options["Documento"].append({"id": doc_id or title, "label": f"{doc_id} | {title}".strip(" |")})
        for rows in options.values():
            rows.sort(key=lambda item: str(item.get("label", "") or item.get("id", "") or "").lower())
        return options

    def link_label(self, entity_type: str, entity_id: str) -> str:
        entity_type_txt = str(entity_type or "Livre").strip() or "Livre"
        entity_id_txt = str(entity_id or "").strip()
        if not entity_id_txt:
            return ""
        for row in list(self.link_options().get(entity_type_txt, []) or []):
            if str(row.get("id", "") or "").strip() == entity_id_txt:
                return str(row.get("label", "") or "").strip()
        return entity_id_txt

