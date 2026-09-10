"""Purchase line validation and material inference over detached catalog reads."""
from copy import deepcopy
from dataclasses import dataclass
from typing import Any, Callable, Protocol
import math
import re

class PurchaseLineRepository(Protocol):
    def material(self, code: str) -> dict[str, Any] | None: ...
    def product(self, code: str) -> dict[str, Any] | None: ...

@dataclass(frozen=True)
class PurchaseLineRules:
    norm_text: Callable
    parse_float: Callable
    dimension: Callable
    format_number: Callable
    geometry: Callable
    is_material: Callable
    material_format: Callable
    location: Callable
    advice: Callable

class PurchaseLines:
    def __init__(self, repository: PurchaseLineRepository, rules: PurchaseLineRules):
        self.repository = repository
        self.rules = rules

    def infer_material(self, line: dict[str, Any]) -> dict[str, Any]:
        row = dict(line or {})
        text = " ".join(
            str(row.get(key, "") or "")
            for key in ("descricao", "material", "dimensao", "dimensoes", "ref")
        )
        norm = self.rules.norm_text(text)

        def _number(value: Any) -> float:
            return self.rules.parse_float(str(value or "").replace(",", "."), 0)

        qty_len = re.search(r"(\d+(?:[.,]\d+)?)\s*un\s*x\s*(\d+(?:[.,]\d+)?)\s*m", text, re.IGNORECASE)
        if qty_len and self.rules.parse_float(row.get("metros", 0), 0) <= 0:
            row["metros"] = _number(qty_len.group(2))
        kg_m_match = re.search(r"(\d+(?:[.,]\d+)?)\s*kg\s*/\s*m", text, re.IGNORECASE)
        if kg_m_match and self.rules.parse_float(row.get("kg_m", 0), 0) <= 0:
            row["kg_m"] = _number(kg_m_match.group(1))
        price_match = re.search(r"(\d+(?:[.,]\d+)?)\s*EUR\s*/\s*(kg|m)", text, re.IGNORECASE)
        if price_match and self.rules.parse_float(row.get("p_compra", row.get("preco", 0)), 0) <= 0:
            row["p_compra"] = _number(price_match.group(1))
            row["price_base_label"] = f"EUR/{price_match.group(2).lower()}"

        if "nervurado" in norm:
            diameter_match = re.search(r"[Øø]\s*(\d+(?:[.,]\d+)?)", text)
            diameter = _number(diameter_match.group(1)) if diameter_match else self.rules.dimension(row.get("espessura", 0), 0)
            row["formato"] = "Varão nervurado"
            row["material"] = str(row.get("material", "") or "Ferro nervurado").strip() or "Ferro nervurado"
            row["espessura"] = self.rules.format_number(diameter) if diameter > 0 else str(row.get("espessura", "") or "").strip()
            row["diametro"] = diameter
            row["secao_tipo"] = "nervurado"
        elif "cantoneira" in norm:
            match = re.search(r"(\d+(?:[.,]\d+)?)\s*x\s*(\d+(?:[.,]\d+)?)\s*x\s*(\d+(?:[.,]\d+)?)\s*mm", text, re.IGNORECASE)
            row["formato"] = "Cantoneira"
            if match:
                side_a = _number(match.group(1))
                side_b = _number(match.group(2))
                thickness = _number(match.group(3))
                row["comprimento"] = side_a
                row["largura"] = side_b
                row["espessura"] = self.rules.format_number(thickness)
                row["secao_tipo"] = "abas_iguais" if abs(side_a - side_b) <= 1e-6 else "abas_desiguais"
        elif "tubo" in norm:
            match = re.search(r"(\d+(?:[.,]\d+)?)\s*x\s*(\d+(?:[.,]\d+)?)\s*x\s*(\d+(?:[.,]\d+)?)\s*mm", text, re.IGNORECASE)
            round_match = re.search(r"[Øø]\s*(\d+(?:[.,]\d+)?)\s*x\s*(\d+(?:[.,]\d+)?)\s*mm", text, re.IGNORECASE)
            row["formato"] = "Tubo"
            if match:
                side_a = _number(match.group(1))
                side_b = _number(match.group(2))
                thickness = _number(match.group(3))
                row["comprimento"] = side_a
                row["largura"] = side_b
                row["altura"] = side_b
                row["espessura"] = self.rules.format_number(thickness)
                row["secao_tipo"] = "quadrado" if abs(side_a - side_b) <= 1e-6 else "retangular"
            elif round_match:
                row["diametro"] = _number(round_match.group(1))
                row["espessura"] = self.rules.format_number(_number(round_match.group(2)))
                row["secao_tipo"] = "redondo"
        elif "barra" in norm:
            match = re.search(r"(\d+(?:[.,]\d+)?)\s*x\s*(\d+(?:[.,]\d+)?)\s*mm", text, re.IGNORECASE)
            row["formato"] = "Barra"
            row["secao_tipo"] = "chata"
            if match:
                side_a = _number(match.group(1))
                side_b = _number(match.group(2))
                row["comprimento"] = side_a
                row["largura"] = side_b
                row["espessura"] = self.rules.format_number(side_b)
        else:
            profile_match = re.search(r"\b(IPE|IPN|UPN|HEA|HEB|HEM)\s*[- ]?(\d{2,4})\b", text, re.IGNORECASE)
            if profile_match or "perfil" in norm:
                row["formato"] = "Perfil"
                if profile_match:
                    row["secao_tipo"] = str(profile_match.group(1) or "").strip().upper()
                    row["altura"] = _number(profile_match.group(2))
                    row["espessura"] = self.rules.format_number(row["altura"])

        formato = str(row.get("formato", "") or "").strip()
        if formato:
            preview = self.rules.geometry(row)
            for key in ("comprimento", "largura", "altura", "diametro", "metros", "kg_m", "peso_unid", "secao_tipo"):
                if preview.get(key) not in (None, ""):
                    row[key] = preview.get(key)
            row["dimensao"] = str(preview.get("dimension_label", "") or row.get("dimensao", row.get("dimensoes", "")) or "").strip()
            row["dimensoes"] = row["dimensao"]
            if str(preview.get("espessura", "") or "").strip():
                row["espessura"] = str(preview.get("espessura", "") or "").strip()
        return row


    def normalize(self, payload: dict[str, Any]) -> dict[str, Any]:
        payload = deepcopy(payload)
        origem = str(payload.get("origem", "Produto") or "Produto").strip() or "Produto"
        ref = str(payload.get("ref", "") or "").strip()
        descricao = str(payload.get("descricao", "") or "").strip()
        fornecedor_linha = str(payload.get("fornecedor_linha", "") or "").strip()
        unid = str(payload.get("unid", "UN") or "UN").strip() or "UN"
        qtd = self.rules.parse_float(payload.get("qtd", 0), 0)
        preco = self.rules.parse_float(payload.get("preco", 0), 0)
        desconto = max(0.0, min(100.0, self.rules.parse_float(payload.get("desconto", 0), 0)))
        iva = max(0.0, min(100.0, self.rules.parse_float(payload.get("iva", 23), 23)))
        if not descricao:
            raise ValueError("Descrição da linha obrigatória.")
        if not math.isfinite(qtd) or qtd <= 0:
            raise ValueError("Quantidade da linha inválida.")
        if not math.isfinite(preco):
            raise ValueError("Preco da linha invalido.")
        base = (qtd * preco) * (1.0 - (desconto / 100.0))
        iva_amt = base * (iva / 100.0)
        total = round(base + iva_amt, 4)
        line = {
            "ref": ref,
            "descricao": descricao,
            "fornecedor_linha": fornecedor_linha,
            "origem": origem,
            "qtd": qtd,
            "unid": unid,
            "preco": preco,
            "total": total,
            "desconto": desconto,
            "iva": iva,
            "entregue": bool(payload.get("entregue")),
            "qtd_entregue": self.rules.parse_float(payload.get("qtd_entregue", qtd if payload.get("entregue") else 0), 0),
        }
        if self.rules.is_material(origem):
            material = self.repository.material(ref)
            if material:
                metrics = self.rules.geometry(material)
                line.update(
                    {
                        "material": material.get("material", ""),
                        "espessura": material.get("espessura", ""),
                        "comprimento": self.rules.dimension(metrics.get("comprimento", material.get("comprimento", 0)), 0),
                        "largura": self.rules.dimension(metrics.get("largura", material.get("largura", 0)), 0),
                        "altura": self.rules.dimension(metrics.get("altura", material.get("altura", 0)), 0),
                        "diametro": self.rules.dimension(metrics.get("diametro", material.get("diametro", 0)), 0),
                        "dimensao": str(material.get("dimensao", material.get("dimensoes", "")) or "").strip(),
                        "dimensoes": str(material.get("dimensao", material.get("dimensoes", "")) or "").strip(),
                        "metros": self.rules.parse_float(metrics.get("metros", material.get("metros", 0)), 0),
                        "kg_m": self.rules.parse_float(metrics.get("kg_m", material.get("kg_m", 0)), 0),
                        "localizacao": self.rules.location(material),
                        "lote_fornecedor": material.get("lote_fornecedor", ""),
                        "peso_unid": self.rules.parse_float(metrics.get("peso_unid", material.get("peso_unid", 0)), 0),
                        "p_compra": self.rules.parse_float(material.get("p_compra", 0), 0),
                        "formato": material.get("formato", self.rules.material_format(material)),
                        "secao_tipo": str(metrics.get("secao_tipo", material.get("secao_tipo", "")) or "").strip(),
                        "material_familia": str(material.get("material_familia", "") or "").strip(),
                        "_material_pending_create": False,
                        "_material_manual": False,
                    }
                )
            else:
                payload = self.infer_material(payload)
                formato_txt = str(payload.get("formato", "") or self.rules.material_format(payload) or "Chapa").strip() or "Chapa"
                material_txt = str(payload.get("material", "") or "").strip()
                esp_txt = str(payload.get("espessura", "") or "").strip()
                if not material_txt:
                    raise ValueError("Qualidade da matéria-prima obrigatória.")
                if formato_txt in {"Chapa", "Tubo", "Cantoneira", "Varão nervurado"} and not esp_txt:
                    raise ValueError("Espessura obrigatória para chapa, tubo, cantoneira e varão nervurado.")
                metrics = self.rules.geometry(payload)
                line.update(
                    {
                        "material": material_txt,
                        "espessura": esp_txt,
                        "comprimento": self.rules.dimension(metrics.get("comprimento", payload.get("comprimento", 0)), 0),
                        "largura": self.rules.dimension(metrics.get("largura", payload.get("largura", 0)), 0),
                        "altura": self.rules.dimension(metrics.get("altura", payload.get("altura", 0)), 0),
                        "diametro": self.rules.dimension(metrics.get("diametro", payload.get("diametro", 0)), 0),
                        "dimensao": str(payload.get("dimensao", payload.get("dimensoes", "")) or "").strip(),
                        "dimensoes": str(payload.get("dimensao", payload.get("dimensoes", "")) or "").strip(),
                        "metros": self.rules.parse_float(metrics.get("metros", payload.get("metros", 0)), 0),
                        "kg_m": self.rules.parse_float(metrics.get("kg_m", payload.get("kg_m", 0)), 0),
                        "localizacao": str(payload.get("localizacao", "") or "").strip(),
                        "lote_fornecedor": str(payload.get("lote_fornecedor", "") or "").strip(),
                        "peso_unid": self.rules.parse_float(metrics.get("peso_unid", payload.get("peso_unid", 0)), 0),
                        "p_compra": self.rules.parse_float(payload.get("p_compra", payload.get("preco", 0)), 0),
                        "formato": formato_txt,
                        "secao_tipo": str(metrics.get("secao_tipo", payload.get("secao_tipo", "")) or "").strip(),
                        "material_familia": str(payload.get("material_familia", "") or "").strip(),
                        "_material_pending_create": bool(payload.get("_material_pending_create", True)),
                        "_material_manual": bool(payload.get("_material_manual", True)),
                    }
                )
        else:
            product = self.repository.product(ref)
            line.update(
                {
                    "categoria": str(payload.get("categoria", (product or {}).get("categoria", "")) or "").strip(),
                    "tipo": str(payload.get("tipo", (product or {}).get("tipo", "")) or "").strip(),
                    "dimensoes": str(payload.get("dimensoes", (product or {}).get("dimensoes", "")) or "").strip(),
                    "peso_unid": self.rules.parse_float(payload.get("peso_unid", (product or {}).get("peso_unid", 0)), 0),
                    "metros_unidade": self.rules.parse_float(
                        payload.get("metros_unidade", (product or {}).get("metros_unidade", 0)),
                        0,
                    ),
                    "price_basis": str(payload.get("price_basis", (product or {}).get("price_basis", "")) or "").strip(),
                    "_product_pending_create": bool(payload.get("_product_pending_create", product is None)),
                }
            )
        line["purchase_advice"] = self.rules.advice(line)
        return line
