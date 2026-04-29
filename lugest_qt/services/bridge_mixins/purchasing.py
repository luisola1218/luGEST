from __future__ import annotations

import os
import re
import tempfile
from pathlib import Path
from typing import Any


class PurchasingBridgeMixin:
    """Supplier, purchase note, and delivery operations for the Qt bridge."""

    def ne_suppliers(self) -> list[dict[str, str]]:
        self._maybe_normalize_single_supplier_catalog()
        rows = []
        for raw in list(self.ensure_data().get("fornecedores", []) or []):
            row = {
                "id": str(raw.get("id", "") or "").strip(),
                "nome": str(raw.get("nome", "") or "").strip(),
                "contacto": str(raw.get("contacto", "") or "").strip(),
                "nif": str(raw.get("nif", "") or "").strip(),
                "morada": str(raw.get("morada", "") or "").strip(),
                "email": str(raw.get("email", "") or "").strip(),
                "codigo_postal": str(raw.get("codigo_postal", "") or "").strip(),
                "localidade": str(raw.get("localidade", "") or "").strip(),
                "pais": str(raw.get("pais", "") or "").strip(),
                "cond_pagamento": str(raw.get("cond_pagamento", "") or "").strip(),
                "prazo_entrega_dias": int(self._parse_float(raw.get("prazo_entrega_dias", 0), 0)),
                "website": str(raw.get("website", "") or "").strip(),
                "obs": str(raw.get("obs", "") or "").strip(),
            }
            rows.append(row)
        def _supplier_sort_key(item: dict[str, Any]) -> tuple[int, str]:
            supplier_id = str(item.get("id", "") or "").strip().upper()
            if supplier_id.startswith("FOR-") and supplier_id[4:].isdigit():
                return (int(supplier_id[4:]), supplier_id)
            return (10**9, supplier_id)
        rows.sort(key=_supplier_sort_key)
        return rows

    def supplier_next_id(self) -> str:
        self._maybe_normalize_single_supplier_catalog()
        return str(self.desktop_main.peek_next_fornecedor_numero(self.ensure_data()))

    def _maybe_normalize_single_supplier_catalog(self) -> None:
        if getattr(self, "_supplier_catalog_fixing", False):
            return
        data = self.ensure_data()
        suppliers = list(data.get("fornecedores", []) or [])
        if len(suppliers) != 1:
            return
        supplier = suppliers[0]
        old_id = str(supplier.get("id", "") or "").strip()
        supplier_name = str(supplier.get("nome", "") or "").strip()
        target_id = "FOR-0001"
        current_next = str(self.desktop_main.peek_next_fornecedor_numero(data) or "").strip().upper()
        notes = list(data.get("notas_encomenda", []) or [])
        needs_fix = old_id != target_id or current_next != "FOR-0002"
        if not needs_fix and not any(str(note.get("fornecedor", "") or "").strip() == old_id for note in notes):
            return
        self._supplier_catalog_fixing = True
        changed = False
        try:
            if old_id != target_id:
                supplier["id"] = target_id
                changed = True
            for note in notes:
                note_supplier_id = str(note.get("fornecedor_id", "") or "").strip()
                note_supplier_txt = str(note.get("fornecedor", "") or "").strip()
                if note_supplier_id in {old_id, target_id}:
                    if note_supplier_id != target_id:
                        note["fornecedor_id"] = target_id
                        changed = True
                    if note_supplier_txt != supplier_name:
                        note["fornecedor"] = supplier_name
                        changed = True
                elif note_supplier_txt == old_id:
                    note["fornecedor"] = supplier_name
                    changed = True
                for line in list(note.get("linhas", []) or []):
                    line_supplier = str(line.get("fornecedor_linha", "") or "").strip()
                    if line_supplier == old_id or line_supplier == target_id or line_supplier.startswith(f"{old_id} - "):
                        if line_supplier != supplier_name:
                            line["fornecedor_linha"] = supplier_name
                            changed = True
            self.desktop_main.reserve_fornecedor_numero(data, target_id)
            if str(self.desktop_main.peek_next_fornecedor_numero(data) or "").strip().upper() != "FOR-0002":
                self.desktop_main._store_fornecedor_sequence_next(data, 2)
                changed = True
            if changed:
                self._save(force=True)
        finally:
            self._supplier_catalog_fixing = False

    def supplier_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        self._maybe_normalize_single_supplier_catalog()
        data = self.ensure_data()
        supplier_id = str(payload.get("id", "") or "").strip()
        nome = str(payload.get("nome", "") or "").strip()
        if not nome:
            raise ValueError("Nome do fornecedor obrigatorio.")
        rows = data.setdefault("fornecedores", [])
        existing = next((item for item in rows if str(item.get("id", "") or "").strip() == supplier_id), None) if supplier_id else None
        if not supplier_id:
            supplier_id = str(self.desktop_main.next_fornecedor_numero(data))
        elif existing is None:
            self.desktop_main.reserve_fornecedor_numero(data, supplier_id)
        row = {
            "id": supplier_id,
            "nome": nome,
            "nif": str(payload.get("nif", "") or "").strip(),
            "morada": str(payload.get("morada", "") or "").strip(),
            "contacto": str(payload.get("contacto", "") or "").strip(),
            "email": str(payload.get("email", "") or "").strip(),
            "codigo_postal": str(payload.get("codigo_postal", "") or "").strip(),
            "localidade": str(payload.get("localidade", "") or "").strip(),
            "pais": str(payload.get("pais", "") or "").strip(),
            "cond_pagamento": str(payload.get("cond_pagamento", "") or "").strip(),
            "prazo_entrega_dias": int(self._parse_float(payload.get("prazo_entrega_dias", 0), 0)),
            "website": str(payload.get("website", "") or "").strip(),
            "obs": str(payload.get("obs", "") or "").strip(),
        }
        if existing is None:
            rows.append(row)
            target = row
        else:
            existing.update(row)
            target = existing
        self._save(force=True)
        return dict(target)

    def supplier_remove(self, supplier_id: str) -> None:
        data = self.ensure_data()
        value = str(supplier_id or "").strip()
        if not value:
            raise ValueError("Fornecedor inv?lido.")
        if any(str(note.get("fornecedor_id", "") or "").strip() == value for note in list(data.get("notas_encomenda", []) or [])):
            raise ValueError("Nao e possivel remover um fornecedor usado em notas de encomenda.")
        before = len(list(data.get("fornecedores", []) or []))
        data["fornecedores"] = [row for row in list(data.get("fornecedores", []) or []) if str(row.get("id", "") or "").strip() != value]
        if len(data["fornecedores"]) == before:
            raise ValueError("Fornecedor n?o encontrado.")
        self._save(force=True)

    def ne_next_number(self) -> str:
        return str(self.desktop_main.peek_next_ne_numero(self.ensure_data()))

    def _infer_purchase_material_line(self, line: dict[str, Any]) -> dict[str, Any]:
        row = dict(line or {})
        text = " ".join(
            str(row.get(key, "") or "")
            for key in ("descricao", "material", "dimensao", "dimensoes", "ref")
        )
        norm = self.desktop_main.norm_text(text)

        def _number(value: Any) -> float:
            return self._parse_float(str(value or "").replace(",", "."), 0)

        qty_len = re.search(r"(\d+(?:[.,]\d+)?)\s*un\s*x\s*(\d+(?:[.,]\d+)?)\s*m", text, re.IGNORECASE)
        if qty_len and self._parse_float(row.get("metros", 0), 0) <= 0:
            row["metros"] = _number(qty_len.group(2))
        kg_m_match = re.search(r"(\d+(?:[.,]\d+)?)\s*kg\s*/\s*m", text, re.IGNORECASE)
        if kg_m_match and self._parse_float(row.get("kg_m", 0), 0) <= 0:
            row["kg_m"] = _number(kg_m_match.group(1))
        price_match = re.search(r"(\d+(?:[.,]\d+)?)\s*EUR\s*/\s*(kg|m)", text, re.IGNORECASE)
        if price_match and self._parse_float(row.get("p_compra", row.get("preco", 0)), 0) <= 0:
            row["p_compra"] = _number(price_match.group(1))
            row["price_base_label"] = f"EUR/{price_match.group(2).lower()}"

        if "nervurado" in norm:
            diameter_match = re.search(r"[Øø]\s*(\d+(?:[.,]\d+)?)", text)
            diameter = _number(diameter_match.group(1)) if diameter_match else self._parse_dimension_mm(row.get("espessura", 0), 0)
            row["formato"] = "Varão nervurado"
            row["material"] = str(row.get("material", "") or "Ferro nervurado").strip() or "Ferro nervurado"
            row["espessura"] = self._fmt(diameter) if diameter > 0 else str(row.get("espessura", "") or "").strip()
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
                row["espessura"] = self._fmt(thickness)
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
                row["espessura"] = self._fmt(thickness)
                row["secao_tipo"] = "quadrado" if abs(side_a - side_b) <= 1e-6 else "retangular"
            elif round_match:
                row["diametro"] = _number(round_match.group(1))
                row["espessura"] = self._fmt(_number(round_match.group(2)))
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
                row["espessura"] = self._fmt(side_b)
        else:
            profile_match = re.search(r"\b(IPE|IPN|UPN|HEA|HEB|HEM)\s*[- ]?(\d{2,4})\b", text, re.IGNORECASE)
            if profile_match or "perfil" in norm:
                row["formato"] = "Perfil"
                if profile_match:
                    row["secao_tipo"] = str(profile_match.group(1) or "").strip().upper()
                    row["altura"] = _number(profile_match.group(2))
                    row["espessura"] = self._fmt(row["altura"])

        formato = str(row.get("formato", "") or "").strip()
        if formato:
            preview = self.material_geometry_preview(row)
            for key in ("comprimento", "largura", "altura", "diametro", "metros", "kg_m", "peso_unid", "secao_tipo"):
                if preview.get(key) not in (None, ""):
                    row[key] = preview.get(key)
            row["dimensao"] = str(preview.get("dimension_label", "") or row.get("dimensao", row.get("dimensoes", "")) or "").strip()
            row["dimensoes"] = row["dimensao"]
            if str(preview.get("espessura", "") or "").strip():
                row["espessura"] = str(preview.get("espessura", "") or "").strip()
        return row

    def ne_rows(self, filter_text: str = "", state_filter: str = "Ativas") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        state_raw = str(state_filter or "Ativas").strip().lower()
        rows = []
        for note in sorted(list(self.ensure_data().get("notas_encomenda", []) or []), key=lambda item: str(item.get("numero", "") or "")):
            if note.get("oculta") and "convert" not in state_raw and "todas" not in state_raw and "todos" not in state_raw:
                continue
            estado = str(note.get("estado", "Em edicao") or "Em edicao").strip()
            estado_norm = self.desktop_main.norm_text(estado)
            is_partial = "parcial" in estado_norm
            is_delivered = "entreg" in estado_norm and not is_partial
            is_converted = "convert" in estado_norm
            if state_raw and state_raw not in ("todas", "todos", "all"):
                if "ativ" in state_raw and (is_delivered or is_converted):
                    continue
                if "edi" in state_raw and "edi" not in estado_norm:
                    continue
                if "apro" in state_raw and "apro" not in estado_norm:
                    continue
                if "enviad" in state_raw and "enviad" not in estado_norm:
                    continue
                if "parcial" in state_raw and not is_partial:
                    continue
                if "entreg" in state_raw and not is_delivered:
                    continue
                if "convert" in state_raw and not is_converted:
                    continue
            row = {
                "numero": str(note.get("numero", "") or "").strip(),
                "fornecedor": str(note.get("fornecedor", "") or "").strip() or ("Multi-fornecedor" if self._note_kind(note) == "rfq" else "Por adjudicar"),
                "data_entrega": str(note.get("data_entrega", "") or "").strip(),
                "estado": estado,
                "total": round(self._parse_float(note.get("total", 0), 0), 2),
                "linhas": len(list(note.get("linhas", []) or [])),
                "draft": bool(note.get("_draft")),
                "oculta": bool(note.get("oculta")),
                "kind": self._note_kind(note),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        return rows

    def ne_material_options(self, filter_text: str = "") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        rows = []
        for record in list(self.ensure_data().get("materiais", []) or []):
            material_id = str(record.get("id", "") or "").strip()
            material = str(record.get("material", "") or "").strip()
            esp_raw = str(record.get("espessura", "") or "").strip()
            esp = self._fmt(esp_raw) if esp_raw else ""
            formato = str(record.get("formato") or self.desktop_main.detect_materia_formato(record) or "Chapa").strip()
            metrics = self.material_price_preview(record)
            preco_unid = float(metrics.get("preco_unid", 0.0) or 0.0)
            comp = round(self._parse_dimension_mm(metrics.get("comprimento", record.get("comprimento", 0)), 0), 3)
            larg = round(self._parse_dimension_mm(metrics.get("largura", record.get("largura", 0)), 0), 3)
            altura = round(self._parse_dimension_mm(metrics.get("altura", record.get("altura", 0)), 0), 3)
            diametro = round(self._parse_dimension_mm(metrics.get("diametro", record.get("diametro", 0)), 0), 3)
            metros = round(self._parse_float(metrics.get("metros", record.get("metros", 0)), 0), 4)
            dim_txt = str(metrics.get("dimension_label", "") or "").strip() or "-"
            esp_txt = f" | {esp} mm" if esp else ""
            lote_txt = str(record.get("lote_fornecedor", "") or "").strip()
            desc = f"{formato} | {material}{esp_txt} | {dim_txt}"
            if metros > 0:
                desc = f"{desc} | {self._fmt(metros)} m"
            if lote_txt:
                desc = f"{desc} | Lote {lote_txt}"
            row = {
                "id": material_id,
                "descricao": desc,
                "material": material,
                "espessura": esp,
                "formato": formato,
                "preco": round(preco_unid, 4),
                "preco_base": round(self._parse_float(record.get("p_compra", 0), 0), 4),
                "preco_base_label": str(metrics.get("base_label", "EUR/kg")),
                "unid": "UN",
                "lote": str(record.get("lote_fornecedor", "") or "").strip(),
                "localizacao": self._localizacao(record),
                "comprimento": round(comp, 3),
                "largura": round(larg, 3),
                "altura": round(altura, 3),
                "diametro": round(diametro, 3),
                "dimensao": dim_txt,
                "secao_tipo": str(metrics.get("secao_tipo", record.get("secao_tipo", "")) or "").strip(),
                "secao_label": str(metrics.get("secao_label", "") or "").strip(),
                "kg_m": round(self._parse_float(metrics.get("kg_m", record.get("kg_m", 0)), 0), 4),
                "metros": round(metros, 4),
                "peso_unid": round(self._parse_float(metrics.get("peso_unid", record.get("peso_unid", 0)), 0), 4),
                "material_familia": str(record.get("material_familia", "") or "").strip(),
                "material_familia_resolved": str(metrics.get("material_familia_resolved", "") or "").strip(),
                "material_familia_label": str(metrics.get("material_familia_label", "") or "").strip(),
                "densidade": round(self._parse_float(metrics.get("densidade", 0), 0), 3),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: (item.get("material") or "", self._parse_float(item.get("espessura", 0), 0), item.get("id") or ""))
        return rows

    def ne_product_options(self, filter_text: str = "") -> list[dict[str, Any]]:
        rows = []
        for raw in self.product_rows(filter_text):
            categoria = str(raw.get("categoria", "") or "").strip()
            tipo = str(raw.get("tipo", "") or "").strip()
            rows.append(
                {
                    "codigo": str(raw.get("codigo", "") or "").strip(),
                    "descricao": str(raw.get("descricao", "") or "").strip(),
                    "origem": "Produto",
                    "stock": round(self._parse_float(raw.get("qty", 0), 0), 2),
                    "unid": str(raw.get("unid", "UN") or "UN").strip(),
                    "preco": round(self._parse_float(raw.get("preco_unid", 0), 0), 4),
                    "preco_unid": round(self._parse_float(raw.get("preco_unid", 0), 0), 4),
                    "p_compra": round(self._parse_float(raw.get("p_compra", 0), 0), 4),
                    "categoria": categoria,
                    "tipo": tipo,
                    "dimensoes": str(raw.get("dimensoes", "") or "").strip(),
                    "peso_unid": round(self._parse_float(raw.get("peso_unid", 0), 0), 4),
                    "metros_unidade": round(self._parse_float(raw.get("metros_unidade", 0), 0), 4),
                    "price_mode": str(self.desktop_main.produto_modo_preco(categoria, tipo) or "compra").strip(),
                }
            )
        return rows

    def ne_detail(self, numero: str) -> dict[str, Any]:
        numero = str(numero or "").strip()
        note = next((row for row in self.ensure_data().get("notas_encomenda", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if note is None:
            raise ValueError("Nota de Encomenda n?o encontrada.")
        documents = self._ne_document_rows(note)
        lines = []
        for line in list(note.get("linhas", []) or []):
            if self.desktop_main.origem_is_materia(line.get("origem", "")):
                line = self._infer_purchase_material_line(line)
            qtd = self._parse_float(line.get("qtd", 0), 0)
            qtd_ent = self._parse_float(line.get("qtd_entregue", qtd if line.get("entregue") else 0), 0)
            origem = str(line.get("origem", "Produto") or "Produto").strip()
            product = None
            if not self.desktop_main.origem_is_materia(origem):
                ref = str(line.get("ref", "") or "").strip()
                if ref:
                    product = next(
                        (
                            row
                            for row in list(self.ensure_data().get("produtos", []) or [])
                            if str(row.get("codigo", "") or "").strip() == ref
                        ),
                        None,
                    )
            peso_unid = round(
                self._parse_float(
                    line.get("peso_unid", (product or {}).get("peso_unid", 0)),
                    0,
                ),
                4,
            )
            if qtd_ent <= 0:
                entrega = "PENDENTE"
            elif qtd_ent < max(0.0, qtd - 1e-9):
                entrega = f"PARCIAL ({self._fmt(qtd_ent)}/{self._fmt(qtd)})"
            else:
                entrega = "ENTREGUE"
            lines.append(
                {
                    "ref": str(line.get("ref", "") or "").strip(),
                    "descricao": str(line.get("descricao", "") or "").strip(),
                    "fornecedor_linha": str(line.get("fornecedor_linha", "") or "").strip(),
                    "origem": origem,
                    "qtd": round(qtd, 4),
                    "unid": str(line.get("unid", "") or "").strip(),
                    "preco": round(self._parse_float(line.get("preco", 0), 0), 4),
                    "desconto": round(self._parse_float(line.get("desconto", 0), 0), 2),
                    "iva": round(self._parse_float(line.get("iva", 23), 23), 2),
                    "total": round(self._parse_float(line.get("total", 0), 0), 4),
                    "entrega": entrega,
                    "material": str(line.get("material", "") or "").strip(),
                    "espessura": str(line.get("espessura", "") or "").strip(),
                    "comprimento": round(self._parse_dimension_mm(line.get("comprimento", 0), 0), 3),
                    "largura": round(self._parse_dimension_mm(line.get("largura", 0), 0), 3),
                    "altura": round(self._parse_dimension_mm(line.get("altura", 0), 0), 3),
                    "diametro": round(self._parse_dimension_mm(line.get("diametro", 0), 0), 3),
                    "metros": round(self._parse_float(line.get("metros", 0), 0), 4),
                    "kg_m": round(self._parse_float(line.get("kg_m", 0), 0), 4),
                    "localizacao": str(line.get("localizacao", "") or "").strip(),
                    "lote_fornecedor": str(line.get("lote_fornecedor", "") or "").strip(),
                    "peso_unid": peso_unid,
                    "peso_total": round(peso_unid * qtd, 4) if peso_unid > 0 and qtd > 0 else 0.0,
                    "p_compra": round(self._parse_float(line.get("p_compra", 0), 0), 6),
                    "formato": str(line.get("formato", "") or "").strip(),
                    "secao_tipo": str(line.get("secao_tipo", "") or "").strip(),
                    "material_familia": str(line.get("material_familia", "") or "").strip(),
                    "categoria": str(line.get("categoria", (product or {}).get("categoria", "")) or "").strip(),
                    "tipo": str(line.get("tipo", (product or {}).get("tipo", "")) or "").strip(),
                    "dimensoes": str(line.get("dimensoes", (product or {}).get("dimensoes", "")) or "").strip(),
                    "metros_unidade": round(
                        self._parse_float(
                            line.get("metros_unidade", (product or {}).get("metros_unidade", 0)),
                            0,
                        ),
                        4,
                    ),
                    "price_basis": str(line.get("price_basis", (product or {}).get("price_basis", "")) or "").strip(),
                    "_material_manual": bool(line.get("_material_manual")),
                    "_material_pending_create": bool(line.get("_material_pending_create")),
                    "_product_pending_create": bool(line.get("_product_pending_create")),
                }
            )
        return {
            "numero": str(note.get("numero", "") or "").strip(),
            "fornecedor": str(note.get("fornecedor", "") or "").strip(),
            "fornecedor_id": str(note.get("fornecedor_id", "") or "").strip(),
            "contacto": str(note.get("contacto", "") or "").strip(),
            "data_entrega": str(note.get("data_entrega", "") or "").strip(),
            "obs": str(note.get("obs", "") or "").strip(),
            "local_descarga": str(note.get("local_descarga", "") or "").strip(),
            "meio_transporte": str(note.get("meio_transporte", "") or "").strip(),
            "estado": str(note.get("estado", "Em edicao") or "Em edicao").strip(),
            "total": round(self._parse_float(note.get("total", 0), 0), 2),
            "draft": bool(note.get("_draft")),
            "kind": self._note_kind(note),
            "ne_geradas": list(note.get("ne_geradas", []) or []),
            "origem_cotacao": str(note.get("origem_cotacao", "") or "").strip(),
            "guia_ultima": str(note.get("guia_ultima", "") or "").strip(),
            "fatura_ultima": str(note.get("fatura_ultima", "") or "").strip(),
            "fatura_caminho_ultima": str(note.get("fatura_caminho_ultima", "") or "").strip(),
            "data_doc_ultima": str(note.get("data_doc_ultima", "") or "").strip(),
            "data_ultima_entrega": str(note.get("data_ultima_entrega", "") or "").strip(),
            "documents": documents,
            "document_count": len(documents),
            "lines": lines,
        }

    def _find_existing_material_from_note_line(self, line: dict[str, Any]) -> dict[str, Any] | None:
        line = self._infer_purchase_material_line(line)
        material_id = str(line.get("ref", "") or "").strip()
        material_txt = self._norm_material_token(line.get("material", ""))
        esp_txt = self._norm_esp_token(line.get("espessura", ""))
        formato_txt = str(line.get("formato", "") or "Chapa").strip() or "Chapa"
        lote_txt = str(line.get("lote_fornecedor", "") or "").strip().lower()
        local_txt = str(line.get("localizacao", "") or "").strip().lower()
        family_txt = ""
        if str(line.get("material_familia", "") or "").strip():
            family_txt = str(self.material_family_profile(line.get("material", ""), line.get("material_familia", "")).get("key", "") or "").strip()
        probe = self.material_geometry_preview(line)
        if not material_txt:
            return None
        records = [record for record in list(self.ensure_data().get("materiais", []) or []) if isinstance(record, dict)]
        if material_id:
            records.sort(key=lambda record: 0 if str(record.get("id", "") or "").strip() == material_id else 1)
        for record in records:
            if self._norm_material_token(record.get("material", "")) != material_txt:
                continue
            if self._norm_esp_token(record.get("espessura", "")) != esp_txt:
                continue
            if str(record.get("formato") or self.desktop_main.detect_materia_formato(record) or "").strip() != formato_txt:
                continue
            if lote_txt and str(record.get("lote_fornecedor", "") or "").strip().lower() != lote_txt:
                continue
            if local_txt and self._localizacao(record).strip().lower() != local_txt:
                continue
            rec_family = ""
            if str(record.get("material_familia", "") or "").strip():
                rec_family = str(self.material_family_profile(record.get("material", ""), record.get("material_familia", "")).get("key", "") or "").strip()
            if family_txt and rec_family and rec_family != family_txt:
                continue
            candidate = self.material_geometry_preview(record)
            if formato_txt == "Tubo":
                if str(candidate.get("secao_tipo", "") or "").strip() != str(probe.get("secao_tipo", "") or "").strip():
                    continue
                if abs(float(candidate.get("metros", 0) or 0) - float(probe.get("metros", 0) or 0)) > 1e-6:
                    continue
                if str(probe.get("secao_tipo", "") or "").strip() == "redondo":
                    if abs(float(candidate.get("diametro", 0) or 0) - float(probe.get("diametro", 0) or 0)) > 1e-6:
                        continue
                else:
                    if abs(float(candidate.get("comprimento", 0) or 0) - float(probe.get("comprimento", 0) or 0)) > 1e-6:
                        continue
                    if abs(float(candidate.get("largura", 0) or 0) - float(probe.get("largura", 0) or 0)) > 1e-6:
                        continue
            elif formato_txt == "Perfil":
                if str(candidate.get("secao_tipo", "") or "").strip() != str(probe.get("secao_tipo", "") or "").strip():
                    continue
                if abs(float(candidate.get("altura", 0) or 0) - float(probe.get("altura", 0) or 0)) > 1e-6:
                    continue
                if abs(float(candidate.get("metros", 0) or 0) - float(probe.get("metros", 0) or 0)) > 1e-6:
                    continue
                if abs(float(candidate.get("kg_m", 0) or 0) - float(probe.get("kg_m", 0) or 0)) > 1e-4:
                    continue
            elif formato_txt == "Varão nervurado":
                if abs(float(candidate.get("diametro", 0) or 0) - float(probe.get("diametro", 0) or 0)) > 1e-6:
                    continue
                if abs(float(candidate.get("metros", 0) or 0) - float(probe.get("metros", 0) or 0)) > 1e-6:
                    continue
            else:
                if abs(float(candidate.get("comprimento", 0) or 0) - float(probe.get("comprimento", 0) or 0)) > 1e-6:
                    continue
                if abs(float(candidate.get("largura", 0) or 0) - float(probe.get("largura", 0) or 0)) > 1e-6:
                    continue
                if formato_txt in {"Barra", "Cantoneira"} and abs(float(candidate.get("metros", 0) or 0) - float(probe.get("metros", 0) or 0)) > 1e-6:
                    continue
            return record
        return None

    def _create_material_placeholder_from_note_line(
        self,
        line: dict[str, Any],
        note_number: str,
        *,
        quantity: float = 0.0,
        lote_override: str = "",
        localizacao_override: str = "",
    ) -> dict[str, Any] | None:
        line = self._infer_purchase_material_line(line)
        data = self.ensure_data()
        material_txt = str(line.get("material", "") or "").strip()
        if not material_txt:
            return None
        formato_txt = str(line.get("formato", "") or self.desktop_main.detect_materia_formato(line) or "Chapa").strip() or "Chapa"
        lote_txt = str(lote_override or line.get("lote_fornecedor", "") or "").strip()
        local_txt = str(localizacao_override or line.get("localizacao", "") or "").strip()
        geometry = self.material_geometry_preview(line)
        record = {
            "id": self._next_material_id(),
            "formato": formato_txt,
            "material": material_txt,
            "material_familia": str(line.get("material_familia", "") or "").strip(),
            "espessura": str(line.get("espessura", "") or "").strip(),
            "comprimento": self._parse_dimension_mm(geometry.get("comprimento", line.get("comprimento", 0)), 0),
            "largura": self._parse_dimension_mm(geometry.get("largura", line.get("largura", 0)), 0),
            "altura": self._parse_dimension_mm(geometry.get("altura", line.get("altura", 0)), 0),
            "diametro": self._parse_dimension_mm(geometry.get("diametro", line.get("diametro", 0)), 0),
            "dimensao": str(line.get("dimensao", line.get("dimensoes", "")) or "").strip(),
            "dimensoes": str(line.get("dimensao", line.get("dimensoes", "")) or "").strip(),
            "metros": self._parse_float(geometry.get("metros", line.get("metros", 0)), 0),
            "kg_m": self._parse_float(geometry.get("kg_m", line.get("kg_m", 0)), 0),
            "quantidade": max(0.0, self._parse_float(quantity, 0)),
            "reservado": 0.0,
            "Localização": local_txt,
            "Localizacao": local_txt,
            "lote_fornecedor": lote_txt,
            "secao_tipo": str(geometry.get("secao_tipo", line.get("secao_tipo", "")) or "").strip(),
            "peso_unid": self._parse_float(geometry.get("peso_unid", line.get("peso_unid", 0)), 0),
            "p_compra": self._parse_float(line.get("p_compra", line.get("preco", 0)), 0),
            "fornecedor": str(line.get("fornecedor", "") or "").strip(),
            "fornecedor_id": str(line.get("fornecedor_id", "") or "").strip(),
            "origem_ne": str(note_number or "").strip(),
            "contorno_points": [],
            "is_sobra": False,
            "atualizado_em": self.desktop_main.now_iso(),
        }
        record["preco_unid"] = float(self.materia_actions._materia_preco_unid_record(record))
        record = self.materia_actions._hydrate_retalho_record(data, record)
        data.setdefault("materiais", []).append(record)
        self.desktop_main.push_unique(data.setdefault("materiais_hist", []), material_txt)
        if str(record.get("espessura", "") or "").strip():
            self.desktop_main.push_unique(data.setdefault("espessuras_hist", []), str(record.get("espessura", "") or "").strip())
        self.desktop_main.log_stock(data, "CRIAR_NE", f"{record.get('id', '')} via {note_number}")
        return record

    def _delivery_inspection_payload(self, update: dict[str, Any]) -> dict[str, str]:
        raw_status = str(update.get("inspection_status", "") or update.get("estado_inspecao", "") or "Aprovado").strip()
        status_key = raw_status.casefold()
        if "rejeit" in status_key:
            status = "Rejeitado"
        elif "reclam" in status_key:
            status = "Reclamado"
        elif "bloque" in status_key:
            status = "Bloqueado"
        elif "inspe" in status_key:
            status = "Em inspeção"
        else:
            status = "Aprovado"
        defect = str(update.get("inspection_defect", "") or update.get("defeito", "") or "").strip()
        decision = str(update.get("inspection_decision", "") or update.get("decisao", "") or "").strip()
        decision_key = decision.casefold()
        if status == "Aprovado":
            if "rejeit" in decision_key or "devolver" in decision_key:
                status = "Rejeitado"
            elif "reclam" in decision_key:
                status = "Reclamado"
            elif "bloque" in decision_key:
                status = "Bloqueado"
            elif "inspe" in decision_key:
                status = "Em inspeção"
        if not decision:
            if status == "Aprovado":
                decision = "Entrada normal"
            elif status == "Em inspeção":
                decision = "Aguardar inspeção"
            elif status == "Reclamado":
                decision = "Reclamar fornecedor"
            elif status == "Rejeitado":
                decision = "Devolver ao fornecedor"
            else:
                decision = "Bloquear stock"
        return {"status": status, "defect": defect, "decision": decision}

    def _material_quality_is_blocked(self, material: dict[str, Any] | None) -> bool:
        if not isinstance(material, dict):
            return False
        status = str(material.get("quality_status", "") or material.get("inspection_status", "") or "").strip().casefold()
        return bool(material.get("quality_blocked")) or any(token in status for token in ("inspe", "bloque", "reclam", "rejeit"))

    def _apply_material_quality_from_delivery(
        self,
        material: dict[str, Any],
        line: dict[str, Any],
        note: dict[str, Any],
        inspection: dict[str, str],
        *,
        quantity: float,
        registo_ts: str,
        guia: str = "",
        fatura: str = "",
    ) -> dict[str, Any] | None:
        if not isinstance(material, dict):
            return None
        status = str(inspection.get("status", "") or "Aprovado").strip() or "Aprovado"
        defect = str(inspection.get("defect", "") or "").strip()
        decision = str(inspection.get("decision", "") or "").strip()
        blocked = status != "Aprovado"
        current_blocked = self._material_quality_is_blocked(material)
        if blocked or not current_blocked:
            material["quality_status"] = status
            material["inspection_status"] = status
            material["quality_blocked"] = bool(blocked)
        material["inspection_defect"] = defect
        material["inspection_decision"] = decision
        material["inspection_at"] = registo_ts
        material["inspection_by"] = str((self.user or {}).get("username", "") or "Sistema")
        material["inspection_note_number"] = str(note.get("numero", "") or "").strip()
        material["inspection_supplier_id"] = str(note.get("fornecedor_id", "") or material.get("fornecedor_id", "") or "").strip()
        material["inspection_supplier_name"] = str(note.get("fornecedor", "") or material.get("fornecedor", "") or "").strip()
        material["inspection_guia"] = str(guia or "").strip()
        material["inspection_fatura"] = str(fatura or "").strip()
        line["inspection_status"] = status
        line["inspection_defect"] = defect
        line["inspection_decision"] = decision
        line["quality_status"] = material.get("quality_status", status)

        if not blocked:
            return None
        existing_nc_id = str(material.get("quality_nc_id", "") or material.get("supplier_claim_id", "") or "").strip()
        if existing_nc_id:
            line["quality_nc_id"] = existing_nc_id
            return {"id": existing_nc_id}
        material_id = str(material.get("id", "") or "").strip()
        lote = str(material.get("lote_fornecedor", "") or line.get("lote_fornecedor", "") or "").strip()
        supplier_name = str(note.get("fornecedor", "") or material.get("fornecedor", "") or line.get("fornecedor", "") or "").strip()
        supplier_id = str(note.get("fornecedor_id", "") or material.get("fornecedor_id", "") or "").strip()
        ref_doc = str(note.get("numero", "") or "").strip()
        if lote:
            ref_doc = f"{ref_doc} / lote {lote}" if ref_doc else f"lote {lote}"
        description = (
            f"Receção de material marcada como {status}. "
            f"Material: {material.get('material', '')} {material.get('espessura', '')}; "
            f"formato: {material.get('formato', '')}; lote: {lote or '-'}; quantidade: {self._fmt(quantity)}. "
            f"Fornecedor: {supplier_name or supplier_id or '-'}."
        )
        if defect:
            description = f"{description} Defeito/observação: {defect}."
        try:
            nc = self.quality_nc_save(
                {
                    "origem": "Receção fornecedor",
                    "referencia": ref_doc,
                    "entidade_tipo": "Material",
                    "entidade_id": material_id,
                    "tipo": "Fornecedor",
                    "gravidade": "Alta" if status in {"Bloqueado", "Reclamado", "Rejeitado"} else "Media",
                    "estado": "Aberta",
                    "responsavel": "Qualidade",
                    "descricao": description,
                    "causa": "A apurar com fornecedor/receção.",
                    "acao": decision or "Bloquear stock e tratar reclamação ao fornecedor.",
                    "fornecedor_id": supplier_id,
                    "fornecedor_nome": supplier_name,
                    "material_id": material_id,
                    "lote_fornecedor": lote,
                    "ne_numero": str(note.get("numero", "") or "").strip(),
                    "guia": str(guia or "").strip(),
                    "fatura": str(fatura or "").strip(),
                    "decisao": decision,
                }
            )
        except Exception:
            nc = None
        if isinstance(nc, dict) and str(nc.get("id", "") or "").strip():
            material["quality_nc_id"] = str(nc.get("id", "") or "").strip()
            material["supplier_claim_id"] = str(nc.get("id", "") or "").strip()
            line["quality_nc_id"] = str(nc.get("id", "") or "").strip()
        return nc

    def _apply_product_quality_from_delivery(
        self,
        product: dict[str, Any] | None,
        line: dict[str, Any],
        note: dict[str, Any],
        inspection: dict[str, str],
        *,
        quantity: float,
        guia: str = "",
        fatura: str = "",
    ) -> dict[str, Any] | None:
        status = str(inspection.get("status", "") or "Aprovado").strip() or "Aprovado"
        defect = str(inspection.get("defect", "") or "").strip()
        decision = str(inspection.get("decision", "") or "").strip()
        line["inspection_status"] = status
        line["inspection_defect"] = defect
        line["inspection_decision"] = decision
        line["quality_status"] = "" if status == "Aprovado" else status
        if status == "Aprovado":
            if isinstance(product, dict):
                product["quality_status"] = "Aprovado"
                product["quality_blocked"] = False
            return None
        if isinstance(product, dict):
            product["quality_status"] = status
            product["quality_blocked"] = True
            product["inspection_defect"] = defect
            product["inspection_decision"] = decision
            product["inspection_note_number"] = str(note.get("numero", "") or "").strip()
        product_code = str((product or {}).get("codigo", "") or line.get("ref", "") or "").strip()
        supplier_name = str(note.get("fornecedor", "") or line.get("fornecedor_linha", "") or "").strip()
        supplier_id = str(note.get("fornecedor_id", "") or "").strip()
        description = (
            f"Receção de produto marcada como {status}. "
            f"Produto: {product_code or '-'} | {str(line.get('descricao', '') or (product or {}).get('descricao', '') or '').strip()}; "
            f"quantidade: {self._fmt(quantity)}. Fornecedor: {supplier_name or supplier_id or '-'}."
        )
        if defect:
            description = f"{description} Defeito/observação: {defect}."
        try:
            nc = self.quality_nc_save(
                {
                    "origem": "Receção fornecedor",
                    "referencia": str(note.get("numero", "") or "").strip(),
                    "entidade_tipo": "Fornecedor",
                    "entidade_id": supplier_id or supplier_name,
                    "tipo": "Fornecedor",
                    "gravidade": "Alta" if status in {"Bloqueado", "Reclamado", "Rejeitado"} else "Media",
                    "estado": "Aberta",
                    "responsavel": "Qualidade",
                    "descricao": description,
                    "causa": "A apurar com fornecedor/receção.",
                    "acao": decision or "Tratar reclamação ao fornecedor.",
                    "fornecedor_id": supplier_id,
                    "fornecedor_nome": supplier_name,
                    "ne_numero": str(note.get("numero", "") or "").strip(),
                    "guia": str(guia or "").strip(),
                    "fatura": str(fatura or "").strip(),
                    "decisao": decision,
                }
            )
        except Exception:
            nc = None
        if isinstance(nc, dict) and str(nc.get("id", "") or "").strip():
            line["quality_nc_id"] = str(nc.get("id", "") or "").strip()
            if isinstance(product, dict):
                product["quality_nc_id"] = str(nc.get("id", "") or "").strip()
        return nc

    def _create_product_placeholder_from_note_line(
        self,
        line: dict[str, Any],
        note_number: str,
        *,
        quantity: float = 0.0,
    ) -> dict[str, Any]:
        data = self.ensure_data()
        ref = str(line.get("ref", "") or "").strip() or str(self.desktop_main.next_produto_numero(data))
        descricao = str(line.get("descricao", "") or "").strip()
        if not descricao:
            raise ValueError("Descricao do produto em falta.")
        product = {
            "codigo": ref,
            "descricao": descricao,
            "categoria": str(line.get("categoria", "") or "Comercial").strip() or "Comercial",
            "subcat": str(line.get("subcat", "") or "").strip(),
            "tipo": str(line.get("tipo", "") or "Stock").strip() or "Stock",
            "dimensoes": str(line.get("dimensoes", "") or "").strip(),
            "comprimento": self._parse_float(line.get("comprimento", 0), 0),
            "largura": self._parse_float(line.get("largura", 0), 0),
            "espessura": self._parse_float(line.get("espessura", 0), 0),
            "metros_unidade": self._parse_float(line.get("metros_unidade", line.get("metros", 0)), 0),
            "metros": self._parse_float(line.get("metros_unidade", line.get("metros", 0)), 0),
            "peso_unid": self._parse_float(line.get("peso_unid", 0), 0),
            "fabricante": str(line.get("fabricante", "") or "").strip(),
            "modelo": str(line.get("modelo", "") or "").strip(),
            "unid": str(line.get("unid", "") or "UN").strip() or "UN",
            "qty": max(0.0, self._parse_float(quantity, 0)),
            "alerta": 0.0,
            "p_compra": self._parse_float(line.get("preco", line.get("p_compra", 0)), 0),
            "pvp1": 0.0,
            "pvp2": 0.0,
            "obs": f"Criado automaticamente pela NE {note_number}",
            "atualizado_em": self.desktop_main.now_iso(),
        }
        product["preco_unid"] = round(self._parse_float(self.desktop_main.produto_preco_unitario(product), 0), 4)
        data.setdefault("produtos", []).append(product)
        self.desktop_main.ensure_produto_seq(data, ref)
        if product["qty"] > 0:
            self.desktop_main.add_produto_mov(
                data,
                tipo="CRIAR_NE",
                operador=str((self.user or {}).get("username", "") or "Sistema"),
                codigo=ref,
                descricao=descricao,
                qtd=product["qty"],
                antes=0.0,
                depois=product["qty"],
                obs=f"NE {note_number}",
                origem="Notas Encomenda",
                ref_doc=note_number,
            )
        return product

    def _ne_normalize_line(self, payload: dict[str, Any]) -> dict[str, Any]:
        origem = str(payload.get("origem", "Produto") or "Produto").strip() or "Produto"
        ref = str(payload.get("ref", "") or "").strip()
        descricao = str(payload.get("descricao", "") or "").strip()
        fornecedor_linha = str(payload.get("fornecedor_linha", "") or "").strip()
        unid = str(payload.get("unid", "UN") or "UN").strip() or "UN"
        qtd = self._parse_float(payload.get("qtd", 0), 0)
        preco = self._parse_float(payload.get("preco", 0), 0)
        desconto = max(0.0, min(100.0, self._parse_float(payload.get("desconto", 0), 0)))
        iva = max(0.0, min(100.0, self._parse_float(payload.get("iva", 23), 23)))
        if not descricao:
            raise ValueError("Descrição da linha obrigatória.")
        if qtd <= 0:
            raise ValueError("Quantidade da linha inválida.")
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
            "qtd_entregue": self._parse_float(payload.get("qtd_entregue", qtd if payload.get("entregue") else 0), 0),
        }
        if self.desktop_main.origem_is_materia(origem):
            material = self.material_by_id(ref)
            if material:
                metrics = self.material_geometry_preview(material)
                line.update(
                    {
                        "material": material.get("material", ""),
                        "espessura": material.get("espessura", ""),
                        "comprimento": self._parse_dimension_mm(metrics.get("comprimento", material.get("comprimento", 0)), 0),
                        "largura": self._parse_dimension_mm(metrics.get("largura", material.get("largura", 0)), 0),
                        "altura": self._parse_dimension_mm(metrics.get("altura", material.get("altura", 0)), 0),
                        "diametro": self._parse_dimension_mm(metrics.get("diametro", material.get("diametro", 0)), 0),
                        "dimensao": str(material.get("dimensao", material.get("dimensoes", "")) or "").strip(),
                        "dimensoes": str(material.get("dimensao", material.get("dimensoes", "")) or "").strip(),
                        "metros": self._parse_float(metrics.get("metros", material.get("metros", 0)), 0),
                        "kg_m": self._parse_float(metrics.get("kg_m", material.get("kg_m", 0)), 0),
                        "localizacao": self._localizacao(material),
                        "lote_fornecedor": material.get("lote_fornecedor", ""),
                        "peso_unid": self._parse_float(metrics.get("peso_unid", material.get("peso_unid", 0)), 0),
                        "p_compra": self._parse_float(material.get("p_compra", 0), 0),
                        "formato": material.get("formato", self.desktop_main.detect_materia_formato(material)),
                        "secao_tipo": str(metrics.get("secao_tipo", material.get("secao_tipo", "")) or "").strip(),
                        "material_familia": str(material.get("material_familia", "") or "").strip(),
                        "_material_pending_create": False,
                        "_material_manual": False,
                    }
                )
            else:
                payload = self._infer_purchase_material_line(payload)
                formato_txt = str(payload.get("formato", "") or self.desktop_main.detect_materia_formato(payload) or "Chapa").strip() or "Chapa"
                material_txt = str(payload.get("material", "") or "").strip()
                esp_txt = str(payload.get("espessura", "") or "").strip()
                if not material_txt:
                    raise ValueError("Qualidade da matéria-prima obrigatória.")
                if formato_txt in {"Chapa", "Tubo", "Cantoneira", "Varão nervurado"} and not esp_txt:
                    raise ValueError("Espessura obrigatória para chapa, tubo, cantoneira e varão nervurado.")
                metrics = self.material_geometry_preview(payload)
                line.update(
                    {
                        "material": material_txt,
                        "espessura": esp_txt,
                        "comprimento": self._parse_dimension_mm(metrics.get("comprimento", payload.get("comprimento", 0)), 0),
                        "largura": self._parse_dimension_mm(metrics.get("largura", payload.get("largura", 0)), 0),
                        "altura": self._parse_dimension_mm(metrics.get("altura", payload.get("altura", 0)), 0),
                        "diametro": self._parse_dimension_mm(metrics.get("diametro", payload.get("diametro", 0)), 0),
                        "dimensao": str(payload.get("dimensao", payload.get("dimensoes", "")) or "").strip(),
                        "dimensoes": str(payload.get("dimensao", payload.get("dimensoes", "")) or "").strip(),
                        "metros": self._parse_float(metrics.get("metros", payload.get("metros", 0)), 0),
                        "kg_m": self._parse_float(metrics.get("kg_m", payload.get("kg_m", 0)), 0),
                        "localizacao": str(payload.get("localizacao", "") or "").strip(),
                        "lote_fornecedor": str(payload.get("lote_fornecedor", "") or "").strip(),
                        "peso_unid": self._parse_float(metrics.get("peso_unid", payload.get("peso_unid", 0)), 0),
                        "p_compra": self._parse_float(payload.get("p_compra", payload.get("preco", 0)), 0),
                        "formato": formato_txt,
                        "secao_tipo": str(metrics.get("secao_tipo", payload.get("secao_tipo", "")) or "").strip(),
                        "material_familia": str(payload.get("material_familia", "") or "").strip(),
                        "_material_pending_create": bool(payload.get("_material_pending_create", True)),
                        "_material_manual": bool(payload.get("_material_manual", True)),
                    }
                )
        else:
            product = next(
                (
                    row
                    for row in list(self.ensure_data().get("produtos", []) or [])
                    if str(row.get("codigo", "") or "").strip() == ref
                ),
                None,
            )
            line.update(
                {
                    "categoria": str(payload.get("categoria", (product or {}).get("categoria", "")) or "").strip(),
                    "tipo": str(payload.get("tipo", (product or {}).get("tipo", "")) or "").strip(),
                    "dimensoes": str(payload.get("dimensoes", (product or {}).get("dimensoes", "")) or "").strip(),
                    "peso_unid": self._parse_float(payload.get("peso_unid", (product or {}).get("peso_unid", 0)), 0),
                    "metros_unidade": self._parse_float(
                        payload.get("metros_unidade", (product or {}).get("metros_unidade", 0)),
                        0,
                    ),
                    "price_basis": str(payload.get("price_basis", (product or {}).get("price_basis", "")) or "").strip(),
                    "_product_pending_create": bool(payload.get("_product_pending_create", product is None)),
                }
            )
        return line

    def ne_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        data = self.ensure_data()
        numero = str(payload.get("numero", "") or "").strip() or self.desktop_main.next_ne_numero(data)
        fornecedor_id = str(payload.get("fornecedor_id", "") or "").strip()
        fornecedor = str(payload.get("fornecedor", "") or "").strip()
        contacto = str(payload.get("contacto", "") or "").strip()
        data_entrega = str(payload.get("data_entrega", "") or "").strip()
        obs = str(payload.get("obs", "") or "").strip()
        local_descarga = str(payload.get("local_descarga", "") or "").strip()
        meio_transporte = str(payload.get("meio_transporte", "") or "").strip()
        lines_payload = list(payload.get("lines", []) or [])
        normalized_lines = [self._ne_normalize_line(line) for line in lines_payload]
        resolved_supplier_id, resolved_supplier_text, resolved_contact = self._normalize_supplier_reference(fornecedor_id, fornecedor)
        fornecedor_id = resolved_supplier_id
        fornecedor = resolved_supplier_text
        if not contacto and resolved_contact:
            contacto = resolved_contact
        for line in normalized_lines:
            line_supplier = str(line.get("fornecedor_linha", "") or "").strip()
            if not line_supplier:
                continue
            _, resolved_line_supplier, _ = self._normalize_supplier_reference("", line_supplier)
            if resolved_line_supplier:
                line["fornecedor_linha"] = resolved_line_supplier
        lst = data.setdefault("notas_encomenda", [])
        existing = next((row for row in lst if str(row.get("numero", "") or "").strip() == numero), None)
        old_lines = list(existing.get("linhas", []) or []) if isinstance(existing, dict) else []
        if not fornecedor and normalized_lines:
            unique_suppliers = {
                str(line.get("fornecedor_linha", "") or "").strip()
                for line in normalized_lines
                if str(line.get("fornecedor_linha", "") or "").strip()
            }
            if len(unique_suppliers) == 1:
                fornecedor = next(iter(unique_suppliers))
                fornecedor_id, fornecedor, inferred_contact = self._resolve_supplier(fornecedor)
                if not contacto:
                    contacto = inferred_contact
        note = {
            "numero": numero,
            "fornecedor": fornecedor,
            "fornecedor_id": fornecedor_id,
            "contacto": contacto,
            "data_entrega": data_entrega,
            "obs": obs,
            "local_descarga": local_descarga,
            "meio_transporte": meio_transporte,
            "linhas": normalized_lines,
            "estado": str((existing or {}).get("estado", "Em edicao") or "Em edicao").strip() or "Em edicao",
            "oculta": bool((existing or {}).get("oculta", False)),
            "_draft": False,
            "origem_cotacao": str((existing or {}).get("origem_cotacao", "") or "").strip(),
            "ne_geradas": list((existing or {}).get("ne_geradas", []) or []),
            "entregas": list((existing or {}).get("entregas", []) or []),
            "documentos": list((existing or {}).get("documentos", []) or []),
            "guia_ultima": str((existing or {}).get("guia_ultima", "") or "").strip(),
            "fatura_ultima": str((existing or {}).get("fatura_ultima", "") or "").strip(),
            "fatura_caminho_ultima": str((existing or {}).get("fatura_caminho_ultima", "") or "").strip(),
            "data_doc_ultima": str((existing or {}).get("data_doc_ultima", "") or "").strip(),
            "data_ultima_entrega": str((existing or {}).get("data_ultima_entrega", "") or "").strip(),
            "data_entregue": str((existing or {}).get("data_entregue", "") or "").strip(),
            "data_aprovacao": str((existing or {}).get("data_aprovacao", "") or "").strip(),
        }
        for index, line in enumerate(note["linhas"]):
            if index >= len(old_lines):
                if line.get("entregue"):
                    line["qtd_entregue"] = self._parse_float(line.get("qtd", 0), 0)
                continue
            old = old_lines[index]
            qtd_tot = self._parse_float(line.get("qtd", 0), 0)
            qtd_old = self._parse_float(old.get("qtd_entregue", old.get("qtd", 0) if old.get("entregue") else 0), 0)
            qtd_old = max(0.0, min(qtd_tot, qtd_old))
            line["qtd_entregue"] = qtd_old
            if old.get("entregue") or (qtd_tot > 0 and qtd_old >= (qtd_tot - 1e-9)):
                line["entregue"] = True
            if old.get("_stock_in") and qtd_old > 0:
                line["_stock_in"] = True
            for key in ("guia_entrega", "fatura_entrega", "data_doc_entrega", "data_entrega_real", "obs_entrega", "entregas_linha"):
                if old.get(key):
                    line[key] = old.get(key)
        note_kind = self._note_kind(note)
        if note_kind == "rfq":
            note["fornecedor"] = ""
            note["fornecedor_id"] = ""
            note["contacto"] = ""
        elif not str(note.get("fornecedor_id", "") or "").strip() and str(note.get("fornecedor", "") or "").strip():
            raise ValueError("Seleciona um fornecedor válido da ficha de fornecedores.")
        product_changed = False
        material_changed = False
        for line in list(note.get("linhas", []) or []):
            if self.desktop_main.origem_is_materia(line.get("origem", "")):
                material_changed = self._update_materia_preco_from_unit(line.get("ref", ""), line.get("preco", 0)) or material_changed
            else:
                product_changed = self._update_produto_preco_from_unit(line.get("ref", ""), line.get("preco", 0)) or product_changed
        if material_changed:
            self._sync_ne_from_materia()
        if product_changed:
            self._sync_ne_from_products()
        self._recalculate_note_totals(note)
        if existing:
            existing.update(note)
        else:
            lst.append(note)
        self._save(force=True)
        return note

    def ne_create_draft(self) -> dict[str, Any]:
        data = self.ensure_data()
        numero = str(self.desktop_main.next_ne_numero(data))
        note = {
            "numero": numero,
            "fornecedor": "",
            "fornecedor_id": "",
            "contacto": "",
            "data_entrega": "",
            "obs": "",
            "local_descarga": "",
            "meio_transporte": "",
            "linhas": [],
            "total": 0.0,
            "estado": "Em edicao",
            "oculta": False,
            "_draft": True,
            "entregas": [],
            "documentos": [],
            "guia_ultima": "",
            "fatura_ultima": "",
            "fatura_caminho_ultima": "",
            "data_doc_ultima": "",
            "data_ultima_entrega": "",
        }
        data.setdefault("notas_encomenda", []).append(note)
        self._save(force=True)
        return note

    def ne_remove(self, numero: str) -> None:
        data = self.ensure_data()
        numero = str(numero or "").strip()
        before = len(list(data.get("notas_encomenda", []) or []))
        data["notas_encomenda"] = [row for row in list(data.get("notas_encomenda", []) or []) if str(row.get("numero", "") or "").strip() != numero]
        if len(data["notas_encomenda"]) == before:
            raise ValueError("Nota de Encomenda n?o encontrada.")
        self._save(force=True)

    def ne_approve(self, numero: str) -> dict[str, Any]:
        numero = str(numero or "").strip()
        note = next((row for row in self.ensure_data().get("notas_encomenda", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if note is None:
            raise ValueError("Nota de Encomenda n?o encontrada.")
        if not list(note.get("linhas", []) or []):
            raise ValueError("A nota n?o tem linhas.")
        note_kind = self._note_kind(note)
        note["estado"] = "Aprovada" if note_kind == "purchase_note" else "Cotacao aprovada"
        note["data_aprovacao"] = self.desktop_main.now_iso()
        note["_draft"] = False
        self._save(force=True)
        return note

    def ne_mark_sent(self, numero: str) -> dict[str, Any]:
        numero = str(numero or "").strip()
        note = next((row for row in self.ensure_data().get("notas_encomenda", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if note is None:
            raise ValueError("Nota de Encomenda n?o encontrada.")
        note["estado"] = "Enviada"
        note["data_envio"] = self.desktop_main.now_iso()
        note["_draft"] = False
        self._save(force=True)
        return note

    def ne_generate_supplier_orders(self, numero: str) -> list[dict[str, Any]]:
        data = self.ensure_data()
        number = str(numero or "").strip()
        note = next((row for row in list(data.get("notas_encomenda", []) or []) if str(row.get("numero", "") or "").strip() == number), None)
        if note is None:
            raise ValueError("Nota de Encomenda n?o encontrada.")
        lines = list(note.get("linhas", []) or [])
        if not lines:
            raise ValueError("A nota n?o tem linhas.")
        groups: dict[str, dict[str, Any]] = {}
        missing: list[str] = []
        for line in lines:
            target_supplier = str(line.get("fornecedor_linha", "") or note.get("fornecedor", "") or "").strip()
            supplier_id, supplier_text, supplier_contact = self._resolve_supplier(target_supplier)
            if not supplier_text:
                missing.append(str(line.get("ref", "") or "").strip() or str(line.get("descricao", "") or "").strip())
                continue
            key = supplier_id or supplier_text
            if key not in groups:
                groups[key] = {
                    "fornecedor_id": supplier_id,
                    "fornecedor": supplier_text,
                    "contacto": supplier_contact,
                    "linhas": [],
                }
            new_line = dict(line)
            new_line["fornecedor_linha"] = supplier_text
            new_line["entregue"] = False
            new_line["qtd_entregue"] = 0.0
            new_line["_stock_in"] = False
            for transient_key in ("guia_entrega", "fatura_entrega", "data_doc_entrega", "data_entrega_real", "obs_entrega", "entregas_linha"):
                new_line.pop(transient_key, None)
            groups[key]["linhas"].append(new_line)
        if missing:
            raise ValueError("Existem linhas sem fornecedor adjudicado: " + ", ".join(sorted(set(item for item in missing if item))))
        if not groups:
            raise ValueError("Nao existem fornecedores adjudicados para gerar NEs.")
        created: list[dict[str, Any]] = []
        notes = data.setdefault("notas_encomenda", [])
        for group in groups.values():
            new_number = str(self.desktop_main.next_ne_numero(data))
            new_note = {
                "numero": new_number,
                "fornecedor": group.get("fornecedor", ""),
                "fornecedor_id": group.get("fornecedor_id", ""),
                "contacto": group.get("contacto", ""),
                "data_entrega": str(note.get("data_entrega", "") or "").strip(),
                "obs": f"Gerada de {note.get('numero', '')}".strip(),
                "local_descarga": str(note.get("local_descarga", "") or "").strip(),
                "meio_transporte": str(note.get("meio_transporte", "") or "").strip(),
                "linhas": list(group.get("linhas", []) or []),
                "estado": "Aprovada",
                "_draft": False,
                "oculta": False,
                "origem_cotacao": str(note.get("numero", "") or "").strip(),
                "ne_geradas": [],
                "entregas": [],
                "documentos": [],
                "guia_ultima": "",
                "fatura_ultima": "",
                "fatura_caminho_ultima": "",
                "data_doc_ultima": "",
                "data_ultima_entrega": "",
            }
            self._recalculate_note_totals(new_note)
            notes.append(new_note)
            created.append({"numero": new_number, "fornecedor": new_note["fornecedor"], "total": new_note["total"]})
        note["estado"] = "Convertida"
        note["oculta"] = True
        note["_draft"] = False
        note["ne_geradas"] = [row["numero"] for row in created]
        self._save(force=True)
        return created

    def ne_render_pdf(self, numero: str, quote: bool = False, output_path: str | Path = "") -> Path:
        numero = str(numero or "").strip()
        note = next((row for row in self.ensure_data().get("notas_encomenda", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if note is None:
            raise ValueError("Nota de Encomenda n?o encontrada.")
        suffix = "_cotacao" if quote else ""
        path = Path(output_path) if str(output_path or "").strip() else Path(tempfile.gettempdir()) / f"lugest_ne_{numero}{suffix}.pdf"
        if quote:
            self.ne_expedicao_actions.render_ne_cotacao_pdf(self, str(path), note)
        else:
            self.ne_expedicao_actions.render_ne_pdf(self, str(path), note)
        return path

    def ne_open_pdf(self, numero: str, quote: bool = False) -> Path:
        path = self.ne_render_pdf(numero, quote=quote)
        os.startfile(str(path))
        return path

    def ne_documents(self, numero: str) -> list[dict[str, Any]]:
        numero = str(numero or "").strip()
        note = next((row for row in self.ensure_data().get("notas_encomenda", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if note is None:
            raise ValueError("Nota de Encomenda n?o encontrada.")
        return self._ne_document_rows(note)

    def ne_add_document(self, numero: str, payload: dict[str, Any]) -> dict[str, Any]:
        numero = str(numero or "").strip()
        note = next((row for row in self.ensure_data().get("notas_encomenda", []) if str(row.get("numero", "") or "").strip() == numero), None)
        if note is None:
            raise ValueError("Nota de Encomenda n?o encontrada.")
        raw = dict(payload or {})
        if not any(str(raw.get(key, "") or "").strip() for key in ("titulo", "guia", "fatura", "caminho", "obs")):
            raise ValueError("Indica pelo menos titulo, guia, fatura, caminho ou observacao.")
        apply_to_lines = bool(raw.get("apply_to_lines"))
        register_history = bool(raw.get("register_history", True))
        doc = self._ne_normalize_document(
            {
                "data_registo": raw.get("data_registo") or self.desktop_main.now_iso(),
                "tipo": raw.get("tipo", ""),
                "titulo": raw.get("titulo", ""),
                "caminho": raw.get("caminho", ""),
                "guia": raw.get("guia", ""),
                "fatura": raw.get("fatura", ""),
                "data_entrega": raw.get("data_entrega", ""),
                "data_documento": raw.get("data_documento", ""),
                "obs": raw.get("obs", ""),
            }
        )
        stored_doc = {
            key: doc.get(key, "")
            for key in ("data_registo", "tipo", "titulo", "caminho", "guia", "fatura", "data_entrega", "data_documento", "obs")
        }
        note.setdefault("documentos", []).append(stored_doc)
        if apply_to_lines:
            for line in list(note.get("linhas", []) or []):
                qtd_total = max(0.0, self._parse_float(line.get("qtd", 0), 0))
                qtd_ent = max(0.0, self._parse_float(line.get("qtd_entregue", qtd_total if line.get("entregue") else 0), 0))
                if qtd_ent <= 0 and not bool(line.get("entregue")):
                    continue
                if doc.get("guia"):
                    line["guia_entrega"] = doc["guia"]
                if doc.get("fatura"):
                    line["fatura_entrega"] = doc["fatura"]
                if doc.get("data_documento"):
                    line["data_doc_entrega"] = doc["data_documento"]
                if doc.get("data_entrega"):
                    line["data_entrega_real"] = doc["data_entrega"]
                if doc.get("obs"):
                    line["obs_entrega"] = doc["obs"]
        if register_history:
            note.setdefault("entregas", []).append(
                {
                    "data_registo": doc.get("data_registo", ""),
                    "data_entrega": doc.get("data_entrega", ""),
                    "guia": doc.get("guia", ""),
                    "fatura": doc.get("fatura", ""),
                    "data_documento": doc.get("data_documento", ""),
                    "obs": doc.get("obs", ""),
                    "linhas": [],
                    "quantidade_linhas": 0,
                    "quantidade_total": 0,
                    "tipo": doc.get("tipo", "DOCUMENTO"),
                    "titulo": doc.get("titulo", ""),
                    "caminho": doc.get("caminho", ""),
                }
            )
        if doc.get("guia"):
            note["guia_ultima"] = doc["guia"]
        if doc.get("fatura"):
            note["fatura_ultima"] = doc["fatura"]
        if doc.get("data_documento"):
            note["data_doc_ultima"] = doc["data_documento"]
        if doc.get("data_entrega"):
            note["data_ultima_entrega"] = doc["data_entrega"]
        if doc.get("caminho") and (doc.get("fatura") or "FATURA" in str(doc.get("tipo", "") or "").upper()):
            note["fatura_caminho_ultima"] = doc["caminho"]
        self._save(force=True)
        return self.ne_detail(numero)

    def ne_register_delivery(self, numero: str, payload: dict[str, Any]) -> dict[str, Any]:
        note = next((row for row in self.ensure_data().get("notas_encomenda", []) if str(row.get("numero", "") or "").strip() == str(numero or "").strip()), None)
        if note is None:
            raise ValueError("Nota de Encomenda n?o encontrada.")
        note_lines = list(note.get("linhas", []) or [])
        line_updates_by_index: dict[int, dict[str, Any]] = {}
        unresolved_legacy_items: list[dict[str, Any]] = []
        for item in list(payload.get("lines", []) or []):
            if not isinstance(item, dict):
                continue
            try:
                line_index = int(item.get("index"))
            except Exception:
                line_index = -1
            if line_index >= 0:
                line_updates_by_index[line_index] = dict(item)
                continue
            unresolved_legacy_items.append(dict(item))
        if unresolved_legacy_items:
            used_indexes = set(line_updates_by_index.keys())
            for item in unresolved_legacy_items:
                ref = str(item.get("ref", "") or "").strip()
                descricao = str(item.get("descricao", "") or "").strip().casefold()
                origem = str(item.get("origem", "") or "").strip().casefold()
                if not ref and not descricao:
                    continue
                matches: list[int] = []
                for idx, line in enumerate(note_lines):
                    if idx in used_indexes:
                        continue
                    line_ref = str(line.get("ref", "") or "").strip()
                    line_desc = str(line.get("descricao", "") or "").strip().casefold()
                    line_origin = str(line.get("origem", "") or "").strip().casefold()
                    if ref and line_ref != ref:
                        continue
                    if descricao and line_desc != descricao:
                        continue
                    if origem and line_origin != origem:
                        continue
                    matches.append(idx)
                if len(matches) == 1:
                    match_index = matches[0]
                    line_updates_by_index[match_index] = dict(item) | {"index": match_index}
                    used_indexes.add(match_index)
        if not line_updates_by_index:
            raise ValueError("Seleciona pelo menos uma linha para entregar.")
        data_entrega = str(payload.get("data_entrega", "") or "").strip()
        data_documento = str(payload.get("data_documento", "") or "").strip()
        guia = str(payload.get("guia", "") or "").strip()
        fatura = str(payload.get("fatura", "") or "").strip()
        titulo = str(payload.get("titulo", "") or "").strip()
        caminho = str(payload.get("caminho", "") or "").strip()
        obs = str(payload.get("obs", "") or "").strip()
        registo_ts = self.desktop_main.now_iso()
        any_delivery = False
        delivered_lines: list[str] = []
        total_qtd = 0.0
        for line_index, line in enumerate(note_lines):
            update = line_updates_by_index.get(line_index)
            if update is None:
                continue
            ref = str(line.get("ref", "") or "").strip()
            add_qtd = max(0.0, self._parse_float(update.get("qtd", 0), 0))
            if add_qtd <= 0:
                continue
            lote_override = str(update.get("lote_fornecedor", "") or "").strip()
            local_override = str(update.get("localizacao", "") or "").strip()
            qtd_total = max(0.0, self._parse_float(line.get("qtd", 0), 0))
            qtd_old = max(0.0, self._parse_float(line.get("qtd_entregue", qtd_total if line.get("entregue") else 0), 0))
            qtd_left = max(0.0, qtd_total - qtd_old)
            qty_apply = min(qtd_left, add_qtd)
            if qty_apply <= 0:
                continue
            working_line = dict(line)
            # Lote e localizacao da entrega sao sempre decididos nesta operacao.
            working_line["lote_fornecedor"] = lote_override
            working_line["localizacao"] = local_override
            working_line["fornecedor"] = str(line.get("fornecedor_linha", "") or note.get("fornecedor", "") or "").strip()
            working_line["fornecedor_id"] = str(note.get("fornecedor_id", "") or "").strip()
            if self.desktop_main.origem_is_materia(line.get("origem", "")):
                working_line = self._infer_purchase_material_line(working_line)
            qtd_new = qtd_old + qty_apply
            line["qtd_entregue"] = qtd_new
            line["entregue"] = qtd_new >= (qtd_total - 1e-9)
            line["logistic_status"] = "RECEBIDO" if line["entregue"] else "PENDENTE"
            line["quality_status"] = "EM_INSPECAO"
            line["inspection_status"] = "EM_INSPECAO"
            line["data_entrega_real"] = data_entrega
            line["data_doc_entrega"] = data_documento
            line["guia_entrega"] = guia
            line["fatura_entrega"] = fatura
            line["obs_entrega"] = obs
            line.setdefault("entregas_linha", []).append(
                {
                    "data_registo": registo_ts,
                    "data_entrega": data_entrega,
                    "data_documento": data_documento,
                    "guia": guia,
                    "fatura": fatura,
                    "obs": obs,
                    "qtd": qty_apply,
                    "lote_fornecedor": lote_override,
                    "localizacao": local_override,
                    "logistic_status": "RECEBIDO",
                    "quality_status": "EM_INSPECAO",
                    "inspection_status": "EM_INSPECAO",
                }
            )
            if self.desktop_main.origem_is_materia(line.get("origem", "")):
                material = self._find_existing_material_from_note_line(working_line)
                if material is None:
                    material = self._create_material_placeholder_from_note_line(
                        working_line,
                        str(note.get("numero", "") or "").strip(),
                        quantity=qty_apply,
                        lote_override=lote_override,
                        localizacao_override=local_override,
                    )
                    if material is None:
                        raise ValueError(
                            "A linha de matéria-prima não tem dados suficientes para criar stock: "
                            + (str(line.get("descricao", "") or "").strip() or str(line.get("material", "") or "").strip() or "-")
                            + "."
                        )
                else:
                    material["quantidade"] = self._parse_float(material.get("quantidade", 0), 0) + qty_apply
                    if lote_override:
                        material["lote_fornecedor"] = lote_override
                    if local_override:
                        material["Localização"] = local_override
                        material["Localizacao"] = local_override
                    material["atualizado_em"] = registo_ts
                if material is not None:
                    material["logistic_status"] = "RECEBIDO"
                    material["quality_status"] = "EM_INSPECAO"
                    material["inspection_status"] = "EM_INSPECAO"
                    material["quality_blocked"] = True
                    material["inspection_at"] = registo_ts
                    material["inspection_by"] = ""
                    material["inspection_note_number"] = str(note.get("numero", "") or "").strip()
                    material["inspection_supplier_id"] = str(note.get("fornecedor_id", "") or material.get("fornecedor_id", "") or "").strip()
                    material["inspection_supplier_name"] = str(note.get("fornecedor", "") or material.get("fornecedor", "") or "").strip()
                    material["inspection_guia"] = str(guia or "").strip()
                    material["inspection_fatura"] = str(fatura or "").strip()
                    material["atualizado_em"] = registo_ts
                    ref = str(material.get("id", "") or "").strip()
                    line["ref"] = ref
                    line["material"] = str(material.get("material", "") or "").strip()
                    line["espessura"] = str(material.get("espessura", "") or "").strip()
                    line["comprimento"] = self._parse_dimension_mm(material.get("comprimento", 0), 0)
                    line["largura"] = self._parse_dimension_mm(material.get("largura", 0), 0)
                    line["dimensao"] = str(material.get("dimensao", material.get("dimensoes", "")) or "").strip()
                    line["dimensoes"] = str(material.get("dimensao", material.get("dimensoes", "")) or "").strip()
                    line["metros"] = self._parse_float(material.get("metros", 0), 0)
                    line["localizacao"] = self._localizacao(material)
                    line["lote_fornecedor"] = str(material.get("lote_fornecedor", "") or "").strip()
                    line["peso_unid"] = self._parse_float(material.get("peso_unid", 0), 0)
                    line["p_compra"] = self._parse_float(material.get("p_compra", 0), 0)
                    line["formato"] = str(material.get("formato") or self.desktop_main.detect_materia_formato(material) or "").strip()
                    line["quality_status"] = str(material.get("quality_status", "") or "").strip()
                    line["quality_nc_id"] = str(material.get("quality_nc_id", "") or "").strip()
                    if line.get("entregas_linha"):
                        line["entregas_linha"][-1]["stock_ref"] = ref
                        line["entregas_linha"][-1]["quality_status"] = line["quality_status"]
                        line["entregas_linha"][-1]["quality_nc_id"] = line["quality_nc_id"]
                    line["_material_pending_create"] = False
                    line["_material_manual"] = False
            else:
                product = next((row for row in list(self.ensure_data().get("produtos", []) or []) if str(row.get("codigo", "") or "").strip() == ref), None)
                if product is None:
                    product = self._create_product_placeholder_from_note_line(
                        working_line,
                        str(note.get("numero", "") or "").strip(),
                        quantity=qty_apply,
                    )
                    ref = str(product.get("codigo", "") or "").strip()
                    line["ref"] = ref
                    line["_product_pending_create"] = False
                else:
                    before = self._parse_float(product.get("qty", 0), 0)
                    product["qty"] = before + qty_apply
                    product["atualizado_em"] = registo_ts
                    self.desktop_main.add_produto_mov(
                        self.ensure_data(),
                        tipo="Entrada",
                        operador=str((self.user or {}).get("username", "") or "Sistema"),
                        codigo=ref,
                        descricao=str(product.get("descricao", "") or "").strip(),
                        qtd=qty_apply,
                        antes=before,
                        depois=product["qty"],
                        obs=f"NE {note.get('numero', '')}",
                        origem="Notas Encomenda",
                        ref_doc=str(note.get("numero", "") or "").strip(),
                    )
                if product is not None:
                    product["logistic_status"] = "RECEBIDO"
                    product["quality_status"] = "EM_INSPECAO"
                    product["quality_blocked"] = True
                    product["inspection_note_number"] = str(note.get("numero", "") or "").strip()
                    product["inspection_supplier_id"] = str(note.get("fornecedor_id", "") or "").strip()
                    product["inspection_supplier_name"] = str(note.get("fornecedor", "") or "").strip()
                    product["inspection_guia"] = str(guia or "").strip()
                    product["inspection_fatura"] = str(fatura or "").strip()
                    product["atualizado_em"] = registo_ts
                    line["quality_status"] = "EM_INSPECAO"
                    line["descricao"] = str(product.get("descricao", "") or "").strip()
                    line["unid"] = str(product.get("unid", line.get("unid", "UN")) or "UN").strip() or "UN"
                    line["categoria"] = str(product.get("categoria", "") or "").strip()
                    line["tipo"] = str(product.get("tipo", "") or "").strip()
                    line["dimensoes"] = str(product.get("dimensoes", "") or "").strip()
                    line["peso_unid"] = self._parse_float(product.get("peso_unid", 0), 0)
                    line["metros_unidade"] = self._parse_float(product.get("metros_unidade", product.get("metros", 0)), 0)
                    if line.get("entregas_linha"):
                        line["entregas_linha"][-1]["quality_status"] = str(line.get("quality_status", "") or "").strip()
                        line["entregas_linha"][-1]["quality_nc_id"] = str(line.get("quality_nc_id", "") or "").strip()
            line["_stock_in"] = True
            any_delivery = True
            total_qtd += qty_apply
            delivered_lines.append(f"{(ref or str(line.get('descricao', '') or '-').strip())} ({self._fmt(qty_apply)})")
        if not any_delivery:
            raise ValueError("Nao foi possivel registar entrega para as quantidades indicadas.")
        note.setdefault("entregas", []).append(
            {
                "data_registo": registo_ts,
                "data_entrega": data_entrega,
                "guia": guia,
                "fatura": fatura,
                "data_documento": data_documento,
                "obs": obs,
                "linhas": delivered_lines,
                "quantidade_linhas": len(delivered_lines),
                "quantidade_total": total_qtd,
            }
        )
        for line in list(note.get("linhas", []) or []):
            qtd_total = max(0.0, self._parse_float(line.get("qtd", 0), 0))
            delivered_qty = 0.0
            for movement in list(line.get("entregas_linha", []) or []):
                delivered_qty += max(0.0, self._parse_float(movement.get("qtd", 0), 0))
            if delivered_qty > qtd_total > 0:
                delivered_qty = qtd_total
            line["qtd_entregue"] = delivered_qty
            line["entregue"] = bool(qtd_total > 0 and delivered_qty >= (qtd_total - 1e-9))
            line["_stock_in"] = bool(delivered_qty > 0)
            line["logistic_status"] = "RECEBIDO" if line["entregue"] else ("PENDENTE" if delivered_qty <= 0 else "PENDENTE")
        note["guia_ultima"] = guia
        note["fatura_ultima"] = fatura
        note["data_doc_ultima"] = data_documento
        note["data_ultima_entrega"] = data_entrega
        if caminho and fatura:
            note["fatura_caminho_ultima"] = caminho
        if any(str(value or "").strip() for value in (titulo, caminho, guia, fatura, obs)):
            note.setdefault("documentos", []).append(
                {
                    "data_registo": registo_ts,
                    "tipo": "ENTREGA",
                    "titulo": self._ne_document_title(
                        {
                            "titulo": titulo,
                            "guia": guia,
                            "fatura": fatura,
                            "caminho": caminho,
                            "data_entrega": data_entrega,
                            "data_documento": data_documento,
                        },
                        doc_type="ENTREGA",
                    ),
                    "caminho": caminho,
                    "guia": guia,
                    "fatura": fatura,
                    "data_entrega": data_entrega,
                    "data_documento": data_documento,
                    "obs": obs,
                }
            )
        all_lines = list(note.get("linhas", []) or [])
        if all(bool(line.get("entregue")) for line in all_lines):
            note["estado"] = "Entregue"
            note["data_entregue"] = data_entrega
        elif any(self._parse_float(line.get("qtd_entregue", 0), 0) > 0 for line in all_lines):
            note["estado"] = "Parcial"
        else:
            note["estado"] = "Aprovada"
        self._save(force=True)
        return self.ne_detail(str(note.get("numero", "") or ""))

