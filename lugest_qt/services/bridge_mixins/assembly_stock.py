from __future__ import annotations

import re
from typing import Any


class AssemblyStockBackendMixin:
    """Legacy adapter for assembly stock; see BACKEND_GUIDE.md."""

    def _preferred_supplier_for_product(self, produto_codigo: str) -> dict[str, str]:
        code = str(produto_codigo or "").strip()
        if not code:
            return {"fornecedor_id": "", "fornecedor": "", "contacto": "", "origem": ""}
        for note in reversed(list(self.ensure_data().get("notas_encomenda", []) or [])):
            note_supplier = str(note.get("fornecedor", "") or "").strip()
            note_contact = str(note.get("contacto", "") or "").strip()
            note_number = str(note.get("numero", "") or "").strip()
            for line in list(note.get("linhas", []) or []):
                if self.desktop_main.origem_is_materia(line.get("origem", "")):
                    continue
                if str(line.get("ref", "") or "").strip() != code:
                    continue
                raw_supplier = str(line.get("fornecedor_linha", "") or note_supplier or "").strip()
                if not raw_supplier:
                    continue
                supplier_id, supplier_text, supplier_contact = self._resolve_supplier(raw_supplier)
                return {
                    "fornecedor_id": supplier_id,
                    "fornecedor": supplier_text or raw_supplier,
                    "contacto": supplier_contact or note_contact,
                    "origem": note_number,
                }
        return {"fornecedor_id": "", "fornecedor": "", "contacto": "", "origem": ""}

    def _order_montagem_shortages(self, enc: dict[str, Any]) -> list[dict[str, Any]]:
        product_map = {
            str(prod.get("codigo", "") or "").strip(): prod
            for prod in list(self.ensure_data().get("produtos", []) or [])
            if str(prod.get("codigo", "") or "").strip()
        }
        shortages: list[dict[str, Any]] = []
        for item in list((enc or {}).get("montagem_itens", []) or []):
            item_type = self.desktop_main.normalize_orc_line_type(item.get("tipo_item"))
            code = str(item.get("produto_codigo", "") or "").strip()
            plan = round(self._parse_float(item.get("qtd_planeada", item.get("qtd", 0)), 0), 2)
            consumed = round(self._parse_float(item.get("qtd_consumida", 0), 0), 2)
            pending = round(max(0.0, plan - consumed), 2)
            if pending <= 1e-9:
                continue
            if self._montagem_item_is_raw_material(item):
                stock_id = str(item.get("stock_material_id", "") or "").strip()
                material_record = self.material_by_id(stock_id) if stock_id else None
                if material_record is not None:
                    available = max(
                        0.0,
                        self._parse_float(material_record.get("quantidade", 0), 0)
                        - self._parse_float(material_record.get("reservado", 0), 0),
                    )
                    material_txt = str(material_record.get("material", item.get("material", "")) or "").strip()
                    esp_txt = str(material_record.get("espessura", item.get("espessura", "")) or "").strip()
                    unit_price = self._parse_float(material_record.get("preco_unid", material_record.get("p_compra", item.get("preco_unit", 0))), 0)
                else:
                    material_txt = str(item.get("material", "") or "").strip()
                    esp_txt = str(item.get("espessura", "") or "").strip()
                    candidates = self.material_candidates(material_txt, esp_txt) if material_txt and esp_txt else []
                    available = round(sum(self._parse_float(row.get("disponivel", 0), 0) for row in candidates), 2)
                    unit_price = self._parse_float(item.get("preco_unit", item.get("price_base_value", 0)), 0)
                missing = round(max(0.0, pending - available), 2)
                if missing > 1e-9:
                    item_key = self._montagem_item_key(item)
                    shortages.append(
                        {
                            "kind": "material",
                            "item_key": item_key,
                            "produto_codigo": "",
                            "descricao": str(item.get("descricao", "") or material_txt or "").strip(),
                            "produto_unid": str(item.get("produto_unid", "") or "UN").strip() or "UN",
                            "qtd_pendente": pending,
                            "qtd_disponivel": round(available, 2),
                            "qtd_em_falta": missing,
                            "produto_encontrado": material_record is not None or available > 0,
                            "preco_unit": round(unit_price, 4),
                            "material": material_txt,
                            "espessura": esp_txt,
                            "dimensao": str(item.get("dimensao", item.get("dimensoes", "")) or "").strip(),
                            "stock_material_id": stock_id,
                            "fornecedor_id": "",
                            "fornecedor_sugerido": "",
                            "fornecedor_contacto": "",
                            "fornecedor_origem": "",
                        }
                    )
                continue
            if item_type != self.desktop_main.ORC_LINE_TYPE_PRODUCT:
                continue
            product = product_map.get(code)
            supplier_meta = self._preferred_supplier_for_product(code)
            available = round(self._parse_float((product or {}).get("qty", 0), 0), 2)
            missing = round(max(0.0, pending - available), 2)
            if product is None or missing > 1e-9:
                unit_price = round(
                    self._parse_float(
                        self.desktop_main.produto_preco_unitario(product or {}) if product is not None else item.get("preco_unit", 0),
                        0,
                    ),
                    4,
                )
                shortages.append(
                    {
                        "produto_codigo": code,
                        "kind": "product",
                        "item_key": self._montagem_item_key(item),
                        "descricao": str(item.get("descricao", "") or (product or {}).get("descricao", "") or "").strip(),
                        "produto_unid": str((product or {}).get("unid", "") or item.get("produto_unid", "") or "UN").strip() or "UN",
                        "qtd_pendente": pending,
                        "qtd_disponivel": available,
                        "qtd_em_falta": missing if product is not None else pending,
                        "produto_encontrado": product is not None,
                        "preco_unit": unit_price,
                        "fornecedor_id": str(supplier_meta.get("fornecedor_id", "") or "").strip(),
                        "fornecedor_sugerido": str(supplier_meta.get("fornecedor", "") or "").strip(),
                        "fornecedor_contacto": str(supplier_meta.get("contacto", "") or "").strip(),
                        "fornecedor_origem": str(supplier_meta.get("origem", "") or "").strip(),
                    }
                )
        shortages.sort(key=lambda row: (-self._parse_float(row.get("qtd_em_falta", 0), 0), str(row.get("produto_codigo", "") or row.get("descricao", "") or "")))
        return shortages

    def _montagem_item_is_raw_material(self, item: dict[str, Any] | None) -> bool:
        row = dict(item or {})
        if self.desktop_main.normalize_orc_line_type(row.get("tipo_item")) != self.desktop_main.ORC_LINE_TYPE_PIECE:
            return False
        raw_fields = self._montagem_item_raw_material_fields(row)
        if str(row.get("stock_item_kind", "") or "").strip() == "raw_material":
            return True
        if str(row.get("stock_material_id", "") or "").strip():
            return True
        if str(row.get("desenho", "") or "").strip():
            return False
        if self._parse_float(row.get("tempo_peca_min", row.get("tempo_pecas_min", row.get("tempo_total_min", 0))), 0) > 0:
            return False
        operacao = str(row.get("operacao", row.get("operacoes", "")) or "").strip()
        if operacao:
            op_norm = self.desktop_main.norm_text(operacao)
            if op_norm not in {"stockmp", "materia prima", "materia-prima"}:
                return False
        marker_text = " ".join(
            str(row.get(key, "") or "")
            for key in (
                "ref_externa",
                "descricao",
                "material",
                "material_family",
                "material_subtype",
                "calc_mode",
                "dimensao",
                "dimensoes",
                "stock_item_kind",
            )
        )
        marker_norm = self.desktop_main.norm_text(marker_text)
        raw_tokens = (
            "perfil",
            "ipe",
            "ipn",
            "hea",
            "heb",
            "upn",
            "barra",
            "chapa",
            "tubo",
            "cantoneira",
            "varao",
            "ferro nervurado",
            "stock mp",
            "stockmp",
        )
        has_material = bool(str(raw_fields.get("material", "") or row.get("material", "") or "").strip())
        has_size = bool(
            str(
                raw_fields.get("espessura", "")
                or raw_fields.get("dimensao", "")
                or row.get("espessura", "")
                or row.get("dimensao", "")
                or row.get("dimensoes", "")
                or ""
            ).strip()
        )
        return has_material and has_size and any(token in marker_norm for token in raw_tokens)

    def _montagem_item_raw_material_fields(self, item: dict[str, Any] | None) -> dict[str, str]:
        row = dict(item or {})
        material = str(row.get("material", "") or "").strip()
        espessura = str(row.get("espessura", "") or "").strip()
        dimensao = str(row.get("dimensao", row.get("dimensoes", "")) or "").strip()
        descricao = str(row.get("descricao", "") or "").strip()
        if material and (espessura or dimensao):
            return {"material": material, "espessura": espessura, "dimensao": dimensao}
        match = re.search(
            r"\b(tubo|chapa|perfil|barra|cantoneira|var[aã]o)\s+([A-Z0-9._/-]+)\s+([0-9]+(?:[,.][0-9]+)?(?:\s*/\s*[0-9]+(?:[,.][0-9]+)?)?(?:\s*x\s*[0-9]+(?:[,.][0-9]+)?(?:\s*/\s*[0-9]+(?:[,.][0-9]+)?){0,1}){0,3})\s*mm\b",
            descricao,
            flags=re.IGNORECASE,
        )
        parsed_material = ""
        parsed_dim = ""
        if match:
            parsed_material = str(match.group(2) or "").strip()
            parsed_dim = re.sub(r"\s+", "", str(match.group(3) or "").strip().replace(",", "."))
        if not parsed_dim:
            dim_match = re.search(
                r"\b([0-9]+(?:[,.][0-9]+)?(?:\s*/\s*[0-9]+(?:[,.][0-9]+)?)?(?:\s*x\s*[0-9]+(?:[,.][0-9]+)?(?:\s*/\s*[0-9]+(?:[,.][0-9]+)?){0,1}){1,3})\s*mm\b",
                descricao,
                flags=re.IGNORECASE,
            )
            if dim_match:
                parsed_dim = re.sub(r"\s+", "", str(dim_match.group(1) or "").strip().replace(",", "."))
        if not parsed_material:
            quality_match = re.search(r"\b(S\d{3,4}[A-Z0-9]*|DD\d{2,3}|DC\d{2,3}|AISI\s*\d{3,4}|INOX(?:\s*\d{3,4})?)\b", descricao, flags=re.IGNORECASE)
            if quality_match:
                parsed_material = re.sub(r"\s+", " ", str(quality_match.group(1) or "").strip()).upper()
        if not parsed_dim:
            profile_match = re.search(r"\b(IPE|IPN|HEA|HEB|UPN|UNP)\s*([0-9]+(?:[,.][0-9]+)?)\b", descricao, flags=re.IGNORECASE)
            if profile_match:
                family = str(profile_match.group(1) or "").upper()
                size = str(profile_match.group(2) or "").replace(",", ".").strip()
                parsed_dim = f"{family} {size}".strip()
                if not espessura:
                    espessura = size
        if not parsed_dim:
            return {"material": material, "espessura": espessura, "dimensao": dimensao}
        if not material:
            material = parsed_material or self._montagem_item_kind_from_description(descricao)
        if not dimensao:
            dimensao = f"{parsed_dim} mm" if parsed_dim else ""
        if not espessura and parsed_dim:
            parts = [part for part in re.split(r"[xX]", parsed_dim) if part.strip()]
            if parts:
                espessura = parts[-1].strip()
        return {"material": material, "espessura": espessura, "dimensao": dimensao}

    def _montagem_item_kind_from_description(self, descricao: str) -> str:
        text = str(descricao or "").strip()
        match = re.search(r"\b(chapa\s+gota|barra\s+chata|cantoneira(?:\s+abas\s+iguais)?|perfil|tubo|barra|var[aã]o)\b", text, flags=re.IGNORECASE)
        if not match:
            return ""
        return re.sub(r"\s+", " ", str(match.group(1) or "").strip()).title()

    def _montagem_item_dimension_numbers(self, item: dict[str, Any] | None) -> list[float]:
        fields = self._montagem_item_raw_material_fields(item)
        text = str(fields.get("dimensao", "") or (item or {}).get("descricao", "") or "").replace(",", ".")
        match = re.search(r"([0-9]+(?:\.[0-9]+)?(?:\s*x\s*[0-9]+(?:\.[0-9]+)?){1,3})", text, flags=re.IGNORECASE)
        if not match:
            return []
        values: list[float] = []
        for token in re.split(r"[xX]", str(match.group(1) or "")):
            try:
                number = float(str(token or "").strip())
            except Exception:
                continue
            if number > 0:
                values.append(round(number, 3))
        return values

    def _montagem_material_candidate_matches_dimension(self, item: dict[str, Any] | None, candidate: dict[str, Any]) -> bool:
        fields = self._montagem_item_raw_material_fields(item)
        expected_material = str(fields.get("material", "") or (item or {}).get("material", "") or "").strip()
        candidate_material = str(candidate.get("material", "") or "").strip()
        if expected_material and candidate_material:
            expected_norm = self.encomendas_actions._norm_material(expected_material)
            candidate_norm = self.encomendas_actions._norm_material(candidate_material)
            if expected_norm != candidate_norm and expected_norm not in candidate_norm and candidate_norm not in expected_norm:
                return False
        expected_esp = str(fields.get("espessura", "") or (item or {}).get("espessura", "") or "").strip()
        candidate_esp = str(candidate.get("espessura", "") or "").strip()
        if expected_esp and candidate_esp:
            if self.encomendas_actions._norm_espessura(candidate_esp) != self.encomendas_actions._norm_espessura(expected_esp):
                return False
        expected = self._montagem_item_dimension_numbers(item)
        if len(expected) < 2:
            return True
        candidate_dims = [
            round(self._parse_float(candidate.get("comprimento", 0), 0), 3),
            round(self._parse_float(candidate.get("largura", 0), 0), 3),
        ]
        candidate_dims = [value for value in candidate_dims if value > 0]
        if len(candidate_dims) < 2:
            return False
        expected_pair = sorted(expected[:2])
        candidate_pair = sorted(candidate_dims[:2])
        return all(abs(exp - got) <= 0.5 for exp, got in zip(expected_pair, candidate_pair))

    def _montagem_material_candidates_for_item(self, item: dict[str, Any] | None) -> list[dict[str, Any]]:
        fields = self._montagem_item_raw_material_fields(item)
        material_txt = str(fields.get("material", "") or (item or {}).get("material", "") or "").strip()
        esp_txt = str(fields.get("espessura", "") or (item or {}).get("espessura", "") or "").strip()
        candidates = self.material_candidates(material_txt, esp_txt) if material_txt and esp_txt else []
        matched = [row for row in candidates if self._montagem_material_candidate_matches_dimension(item, row)]
        if matched:
            return matched
        fallback: list[dict[str, Any]] = []
        for stock in list(self.ensure_data().get("materiais", []) or []):
            if self._material_quality_is_blocked(stock):
                continue
            total_qty = self._parse_float(stock.get("quantidade", 0), 0)
            reserved = self._parse_float(stock.get("reservado", 0), 0)
            disponivel = max(0.0, total_qty - reserved)
            if disponivel <= 0:
                continue
            candidate = {
                "material_id": str(stock.get("id", "") or "").strip(),
                "material": str(stock.get("material", "") or "").strip(),
                "espessura": str(stock.get("espessura", "") or "").strip(),
                "comprimento": round(self._parse_float(stock.get("comprimento", 0), 0), 2),
                "largura": round(self._parse_float(stock.get("largura", 0), 0), 2),
                "disponivel": round(disponivel, 2),
                "quantidade_total": round(total_qty, 2),
                "reservado": round(reserved, 2),
                "local": self._localizacao(stock),
                "lote": str(stock.get("lote_interno", "") or stock.get("lote_fornecedor", "") or "").strip(),
                "lote_fornecedor": str(stock.get("lote_fornecedor", "") or "").strip(),
                "peso_unid": round(self._parse_float(stock.get("peso_unid", 0), 0), 3),
                "p_compra": round(self._parse_float(stock.get("p_compra", 0), 0), 6),
                "is_retalho": bool(stock.get("is_sobra")),
            }
            candidate["dimensao"] = "x".join(
                part
                for part in (self._fmt(candidate["comprimento"]), self._fmt(candidate["largura"]))
                if str(part).strip() and str(part).strip() != "0"
            ) or "-"
            if self._montagem_material_candidate_matches_dimension(item, candidate):
                fallback.append(candidate)
        fallback.sort(
            key=lambda row: (
                0 if bool(row.get("is_retalho")) else 1,
                float(row.get("disponivel", 0) or 0),
                str(row.get("material_id", "") or ""),
            )
        )
        return fallback

    def _montagem_item_key(self, item: dict[str, Any] | None) -> str:
        row = dict(item or {})
        for key in ("linha_ordem", "grupo_uuid", "stock_material_id", "produto_codigo"):
            value = str(row.get(key, "") or "").strip()
            if value:
                return f"{key}:{value}"
        parts = [
            str(row.get("tipo_item", "") or "").strip(),
            str(row.get("descricao", "") or "").strip(),
            str(row.get("material", "") or "").strip(),
            str(row.get("espessura", "") or "").strip(),
        ]
        return "raw:" + "|".join(parts)

    def montagem_purchase_needs(self, order_numbers: list[str] | None = None) -> list[dict[str, Any]]:
        selected = {str(value or "").strip() for value in list(order_numbers or []) if str(value or "").strip()}
        data = self.ensure_data()
        grouped: dict[str, dict[str, Any]] = {}
        for enc in list(data.get("encomendas", []) or []):
            numero = str(enc.get("numero", "") or "").strip()
            if selected and numero not in selected:
                continue
            shortages = self._order_montagem_shortages(enc)
            if not shortages:
                continue
            client_code = str(enc.get("cliente", "") or "").strip()
            client_name = next(
                (
                    str(row.get("nome", "") or "").strip()
                    for row in list(data.get("clientes", []) or [])
                    if str(row.get("codigo", "") or "").strip() == client_code
                ),
                "",
            )
            client_label = " - ".join([value for value in (client_code, client_name) if value]).strip(" -")
            delivery_date = str(enc.get("data_entrega", "") or "").strip()
            for shortage in shortages:
                kind = str(shortage.get("kind", "product") or "product").strip()
                code = str(shortage.get("produto_codigo", "") or "").strip()
                key = f"{kind}:{code or shortage.get('stock_material_id', '') or shortage.get('material', '')}|{shortage.get('espessura', '')}|{shortage.get('descricao', '')}"
                entry = grouped.setdefault(
                    key,
                    {
                        "kind": kind,
                        "produto_codigo": code,
                        "descricao": str(shortage.get("descricao", "") or "").strip(),
                        "produto_unid": str(shortage.get("produto_unid", "") or "UN").strip() or "UN",
                        "preco_unit": round(self._parse_float(shortage.get("preco_unit", 0), 0), 4),
                        "material": str(shortage.get("material", "") or "").strip(),
                        "espessura": str(shortage.get("espessura", "") or "").strip(),
                        "dimensao": str(shortage.get("dimensao", "") or "").strip(),
                        "stock_material_id": str(shortage.get("stock_material_id", "") or "").strip(),
                        "qtd_em_falta": 0.0,
                        "produto_encontrado": bool(shortage.get("produto_encontrado")),
                        "fornecedor_id": str(shortage.get("fornecedor_id", "") or "").strip(),
                        "fornecedor": str(shortage.get("fornecedor_sugerido", "") or "").strip(),
                        "fornecedor_contacto": str(shortage.get("fornecedor_contacto", "") or "").strip(),
                        "fornecedor_origem": str(shortage.get("fornecedor_origem", "") or "").strip(),
                        "encomendas": [],
                        "clientes": [],
                        "_datas_entrega": [],
                    },
                )
                entry["qtd_em_falta"] = round(
                    self._parse_float(entry.get("qtd_em_falta", 0), 0) + self._parse_float(shortage.get("qtd_em_falta", 0), 0),
                    2,
                )
                if not str(entry.get("descricao", "") or "").strip():
                    entry["descricao"] = str(shortage.get("descricao", "") or "").strip()
                if not str(entry.get("fornecedor", "") or "").strip() and str(shortage.get("fornecedor_sugerido", "") or "").strip():
                    entry["fornecedor_id"] = str(shortage.get("fornecedor_id", "") or "").strip()
                    entry["fornecedor"] = str(shortage.get("fornecedor_sugerido", "") or "").strip()
                    entry["fornecedor_contacto"] = str(shortage.get("fornecedor_contacto", "") or "").strip()
                    entry["fornecedor_origem"] = str(shortage.get("fornecedor_origem", "") or "").strip()
                if numero and numero not in entry["encomendas"]:
                    entry["encomendas"].append(numero)
                if client_label and client_label not in entry["clientes"]:
                    entry["clientes"].append(client_label)
                if delivery_date and delivery_date not in entry["_datas_entrega"]:
                    entry["_datas_entrega"].append(delivery_date)
        rows = list(grouped.values())
        for row in rows:
            delivery_dates = sorted(str(value or "").strip() for value in list(row.pop("_datas_entrega", []) or []) if str(value or "").strip())
            row["data_entrega"] = delivery_dates[0] if delivery_dates else ""
            advice_line = {
                "ref": str(row.get("stock_material_id", "") or row.get("produto_codigo", "") or "").strip(),
                "descricao": str(row.get("descricao", "") or "").strip(),
                "origem": "Materia-prima" if str(row.get("kind", "") or "").strip() == "material" else "Produto",
                "qtd": self._parse_float(row.get("qtd_em_falta", 0), 0),
                "unid": str(row.get("produto_unid", "") or "UN").strip() or "UN",
                "preco": self._parse_float(row.get("preco_unit", 0), 0),
                "material": str(row.get("material", "") or "").strip(),
                "espessura": str(row.get("espessura", "") or "").strip(),
                "dimensao": str(row.get("dimensao", "") or "").strip(),
            }
            advice_getter = getattr(self, "purchase_advice_for_line", None)
            advice = dict(advice_getter(advice_line) or {}) if callable(advice_getter) else {}
            row["purchase_advice"] = advice
            if not str(row.get("fornecedor", "") or "").strip() and str(advice.get("habitual_supplier", "") or "").strip():
                row["fornecedor_id"] = str(advice.get("supplier_id", "") or "").strip()
                row["fornecedor"] = str(advice.get("habitual_supplier", "") or "").strip()
                row["fornecedor_contacto"] = str(advice.get("contact", "") or "").strip()
                row["fornecedor_origem"] = str(advice.get("last_purchase", "") or "").strip()
            if self._parse_float(row.get("preco_unit", 0), 0) <= 0 and self._parse_float(advice.get("avg_price", 0), 0) > 0:
                row["preco_unit"] = round(self._parse_float(advice.get("avg_price", 0), 0), 4)
            row["ultima_compra"] = str(advice.get("last_purchase", "") or "").strip()
            row["preco_medio"] = round(self._parse_float(advice.get("avg_price", 0), 0), 4)
            row["prazo_dias"] = self._parse_float(advice.get("lead_days", 0), 0)
            row["qtd_minima"] = round(self._parse_float(advice.get("min_qty", 0), 0), 4)
            row["alternativa_stock"] = str(advice.get("stock_alternative_txt", "") or "").strip()
        rows.sort(
            key=lambda row: (
                not bool(str(row.get("fornecedor", "") or "").strip()),
                str(row.get("data_entrega", "") or "9999-99-99"),
                -self._parse_float(row.get("qtd_em_falta", 0), 0),
                str(row.get("produto_codigo", "") or row.get("descricao", "") or ""),
            )
        )
        return rows

    def ne_create_from_montagem_shortages(self, order_numbers: list[str] | None = None) -> dict[str, Any]:
        needs = list(self.montagem_purchase_needs(order_numbers))
        if not needs:
            raise ValueError("Nao existem faltas de stock de montagem para gerar nota.")
        all_orders = sorted(
            {
                str(order or "").strip()
                for need in needs
                for order in list(need.get("encomendas", []) or [])
                if str(order or "").strip()
            }
        )
        delivery_dates = sorted(
            {
                str(need.get("data_entrega", "") or "").strip()
                for need in needs
                if str(need.get("data_entrega", "") or "").strip()
            }
        )
        unique_suppliers = []
        missing_supplier = []
        for need in needs:
            supplier_txt = str(need.get("fornecedor", "") or "").strip()
            if supplier_txt and supplier_txt.lower() not in [value.lower() for value in unique_suppliers]:
                unique_suppliers.append(supplier_txt)
            if not supplier_txt:
                missing_supplier.append(str(need.get("produto_codigo", "") or need.get("descricao", "") or "").strip())
        supplier_id = ""
        supplier_text = ""
        supplier_contact = ""
        if len(unique_suppliers) == 1:
            supplier_id, supplier_text, supplier_contact = self._resolve_supplier(unique_suppliers[0])
        obs_parts = ["Reposicao automatica de montagem"]
        if all_orders:
            obs_parts.append("Encomendas: " + ", ".join(all_orders))
        if missing_supplier:
            obs_parts.append("Fornecedor por validar: " + ", ".join(sorted(set(item for item in missing_supplier if item))))
        note = self.ne_save(
            {
                "fornecedor": supplier_text,
                "fornecedor_id": supplier_id,
                "contacto": supplier_contact,
                "data_entrega": delivery_dates[0] if delivery_dates else "",
                "obs": " | ".join(obs_parts),
                "lines": [
                    (
                        {
                            "ref": str(need.get("stock_material_id", "") or "").strip(),
                            "descricao": str(need.get("descricao", "") or need.get("material", "") or "").strip(),
                            "fornecedor_linha": str(need.get("fornecedor", "") or "").strip(),
                            "origem": "Materia-prima",
                            "qtd": round(self._parse_float(need.get("qtd_em_falta", 0), 0), 2),
                            "unid": str(need.get("produto_unid", "") or "UN").strip() or "UN",
                            "preco": round(self._parse_float(need.get("preco_unit", 0), 0), 4),
                            "desconto": 0.0,
                            "iva": 23.0,
                            "material": str(need.get("material", "") or "").strip(),
                            "espessura": str(need.get("espessura", "") or "").strip(),
                            "dimensao": str(need.get("dimensao", "") or "").strip(),
                            "dimensoes": str(need.get("dimensao", "") or "").strip(),
                            "purchase_advice": dict(need.get("purchase_advice", {}) or {}),
                        }
                        if str(need.get("kind", "") or "").strip() == "material"
                        else {
                            "ref": str(need.get("produto_codigo", "") or "").strip(),
                            "descricao": str(need.get("descricao", "") or "").strip(),
                            "fornecedor_linha": str(need.get("fornecedor", "") or "").strip(),
                            "origem": "Produto",
                            "qtd": round(self._parse_float(need.get("qtd_em_falta", 0), 0), 2),
                            "unid": str(need.get("produto_unid", "") or "UN").strip() or "UN",
                            "preco": round(self._parse_float(need.get("preco_unit", 0), 0), 4),
                            "desconto": 0.0,
                            "iva": 23.0,
                            "purchase_advice": dict(need.get("purchase_advice", {}) or {}),
                        }
                    )
                    for need in needs
                    if self._parse_float(need.get("qtd_em_falta", 0), 0) > 0
                ],
            }
        )
        note_number = str(note.get("numero", "") or "").strip()
        return {
            "numero": note_number,
            "orders": all_orders,
            "line_count": len(list(note.get("linhas", []) or [])),
            "missing_supplier": sorted(set(item for item in missing_supplier if item)),
            "detail": self.ne_detail(note_number),
        }

    def operator_montagem_stock_group(self, numero: str) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            return {}
        data = self.ensure_data()
        enc_num = str(enc.get("numero", "") or numero or "").strip()
        if not enc_num:
            return {}
        cliente_codigo = str(enc.get("cliente", "") or "").strip()
        cliente_nome = next(
            (
                str(row.get("nome", "") or "").strip()
                for row in list(data.get("clientes", []) or [])
                if isinstance(row, dict) and str(row.get("codigo", "") or "").strip() == cliente_codigo
            ),
            "",
        )
        raw_items = list(enc.get("montagem_itens", []) or [])
        rows: list[dict[str, Any]] = []
        for index, item in enumerate(raw_items):
            item_type = self.desktop_main.normalize_orc_line_type(item.get("tipo_item"))
            is_raw = self._montagem_item_is_raw_material(item)
            is_product = item_type == self.desktop_main.ORC_LINE_TYPE_PRODUCT
            if not is_product and not is_raw:
                continue
            raw_fields = self._montagem_item_raw_material_fields(item) if is_raw else {}
            plan = round(self._parse_float(item.get("qtd_planeada", 0), 0), 2)
            done = round(self._parse_float(item.get("qtd_consumida", 0), 0), 2)
            pending = round(max(0.0, plan - done), 2)
            shortage = 0.0
            if is_product:
                code = str(item.get("produto_codigo", "") or "").strip()
                product = next(
                    (
                        prod
                        for prod in list(self.ensure_data().get("produtos", []) or [])
                        if str(prod.get("codigo", "") or "").strip() == code
                    ),
                    None,
                )
                available = self._parse_float((product or {}).get("qty", 0), 0)
                shortage = round(max(0.0, pending - available), 2)
            elif is_raw:
                stock = self.material_by_id(str(item.get("stock_material_id", "") or "").strip())
                if stock is not None:
                    stock_candidate = {
                        "material": stock.get("material", ""),
                        "espessura": stock.get("espessura", ""),
                        "comprimento": stock.get("comprimento", 0),
                        "largura": stock.get("largura", 0),
                    }
                    if self._montagem_material_candidate_matches_dimension(item, stock_candidate):
                        available = self._parse_float((stock or {}).get("quantidade", 0), 0) - self._parse_float((stock or {}).get("reservado", 0), 0)
                    else:
                        available = 0.0
                else:
                    candidates = self._montagem_material_candidates_for_item(item)
                    available = sum(self._parse_float(row.get("disponivel", 0), 0) for row in candidates)
                shortage = round(max(0.0, pending - available), 2)
            rows.append(
                {
                    "id": f"COMP::{enc_num}::{index}",
                    "grupo_operador": "matéria-prima" if is_raw else "componentes",
                    "espessura": str(raw_fields.get("espessura", "") or item.get("espessura", "") or "").strip(),
                    "codigo": str(item.get("produto_codigo", "") or item.get("stock_material_id", "") or item.get("item_key", "") or "-").strip() or "-",
                    "descricao": str(item.get("descricao", "") or item.get("material", "") or "-").strip() or "-",
                    "tipo_item": item_type,
                    "tipo_label": "Matéria-prima" if is_raw else str(item.get("tipo_label", "") or "Componente/stock").strip(),
                    "material": str(raw_fields.get("material", "") or item.get("material", "") or "").strip(),
                    "dimensao": str(raw_fields.get("dimensao", "") or item.get("dimensao", item.get("dimensoes", "")) or "").strip(),
                    "unidade": str(item.get("produto_unid", "") or "UN").strip() or "UN",
                    "qtd_planeada": plan,
                    "qtd_consumida": done,
                    "qtd_pendente": pending,
                    "falta": shortage,
                    "estado": "Consumido" if pending <= 1e-9 else ("Sem stock" if shortage > 0 else "Pendente"),
                }
            )
        if not rows:
            return {}
        rows.sort(
            key=lambda row: (
                1 if str(row.get("grupo_operador", "") or "") == "matéria-prima" else 0,
                str(row.get("codigo", "") or ""),
                str(row.get("descricao", "") or ""),
            )
        )
        component_count = sum(1 for row in rows if str(row.get("grupo_operador", "") or "") != "matéria-prima")
        raw_count = sum(1 for row in rows if str(row.get("grupo_operador", "") or "") == "matéria-prima")
        plan_total = round(sum(float(row.get("qtd_planeada", 0) or 0) for row in rows), 2)
        done_total = round(sum(float(row.get("qtd_consumida", 0) or 0) for row in rows), 2)
        pending_total = round(max(0.0, plan_total - done_total), 2)
        progress = 0.0 if plan_total <= 0 else round(min(100.0, (done_total / plan_total) * 100.0), 1)
        state = "Consumido" if pending_total <= 1e-9 else ("Sem stock" if any(float(row.get("falta", 0) or 0) > 0 for row in rows) else "Pendente")
        return {
            "is_montagem_stock_group": True,
            "encomenda": enc_num,
            "cliente": f"{cliente_codigo or '-'} - {cliente_nome}".strip(" -"),
            "cliente_codigo": cliente_codigo,
            "cliente_nome": cliente_nome,
            "estado": state,
            "estado_espessura": state,
            "material": "Componentes/Matéria-prima",
            "espessura": "Stock",
            "tempo_plan_min": float(self.desktop_main.encomenda_montagem_tempo_min(enc) or 0),
            "tempo_real_min": 0.0,
            "desvio_min": 0.0,
            "progress_pct": progress,
            "pieces": [],
            "montagem_items": rows,
            "componentes_count": component_count,
            "materia_prima_count": raw_count,
            "can_consume_montagem": pending_total > 1e-9,
        }

    def operator_montagem_stock_groups(self, numero: str) -> list[dict[str, Any]]:
        base_group = dict(self.operator_montagem_stock_group(numero) or {})
        if not base_group:
            return []
        rows = [dict(row or {}) for row in list(base_group.get("montagem_items", []) or [])]
        if not rows:
            return []

        def build_group(kind: str, label: str, material_label: str, subset: list[dict[str, Any]]) -> dict[str, Any]:
            plan_total = round(sum(float(row.get("qtd_planeada", 0) or 0) for row in subset), 2)
            done_total = round(sum(float(row.get("qtd_consumida", 0) or 0) for row in subset), 2)
            pending_total = round(max(0.0, plan_total - done_total), 2)
            progress = 0.0 if plan_total <= 0 else round(min(100.0, (done_total / plan_total) * 100.0), 1)
            state = "Consumido" if pending_total <= 1e-9 else ("Sem stock" if any(float(row.get("falta", 0) or 0) > 0 for row in subset) else "Pendente")
            group = dict(base_group)
            group.update(
                {
                    "montagem_group_kind": kind,
                    "montagem_group_label": label,
                    "estado": state,
                    "estado_espessura": state,
                    "material": material_label,
                    "espessura": "Stock",
                    "progress_pct": progress,
                    "montagem_items": subset,
                    "componentes_count": sum(1 for row in subset if str(row.get("grupo_operador", "") or "") != "matéria-prima"),
                    "materia_prima_count": sum(1 for row in subset if str(row.get("grupo_operador", "") or "") == "matéria-prima"),
                    "can_consume_montagem": pending_total > 1e-9,
                }
            )
            return group

        component_rows = [row for row in rows if str(row.get("grupo_operador", "") or "") != "matéria-prima"]
        raw_rows = [row for row in rows if str(row.get("grupo_operador", "") or "") == "matéria-prima"]
        groups: list[dict[str, Any]] = []
        if component_rows:
            groups.append(build_group("componentes", "Componentes", "Componentes", component_rows))
        if raw_rows:
            groups.append(build_group("materia_prima", "Matéria-prima", "Matéria-prima", raw_rows))
        return groups

    def operator_montagem_stock_options(self, numero: str, item_id: str) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda não encontrada.")
        target_id = str(item_id or "").strip()
        for item_index, item in enumerate(list(enc.get("montagem_itens", []) or [])):
            current_id = f"COMP::{numero}::{item_index}"
            if current_id != target_id:
                continue
            if not self._montagem_item_is_raw_material(item):
                return {"item_id": current_id, "is_raw_material": False, "options": []}
            plan = self._parse_float(item.get("qtd_planeada", item.get("qtd", 0)), 0)
            done = self._parse_float(item.get("qtd_consumida", 0), 0)
            pending = round(max(0.0, plan - done), 4)
            raw_fields = self._montagem_item_raw_material_fields(item)
            options: list[dict[str, Any]] = []
            for candidate in self._montagem_material_candidates_for_item(item):
                material_id = str(candidate.get("material_id", "") or "").strip()
                stock = self.material_by_id(material_id)
                preview = dict(self.material_price_preview(stock) or {}) if isinstance(stock, dict) else {}
                formato = str(preview.get("formato", (stock or {}).get("formato", "")) or "").strip()
                metros = round(self._parse_float((stock or {}).get("metros", 0), 0), 3) if isinstance(stock, dict) else 0.0
                options.append(
                    {
                        **dict(candidate or {}),
                        "material_id": material_id,
                        "formato": formato,
                        "metros": metros,
                        "kg_m": round(self._parse_float((stock or {}).get("kg_m", 0), 0), 4) if isinstance(stock, dict) else 0.0,
                        "lote_interno": str((stock or {}).get("lote_interno", "") or "").strip() if isinstance(stock, dict) else "",
                        "label": (
                            f"{material_id} | {candidate.get('dimensao', '-')}"
                            f" | disp. {self._fmt(candidate.get('disponivel', 0))}"
                            f"{' un' if formato.lower() == 'chapa' else ''}"
                            f"{f' | {self._fmt(metros)} m/un' if metros > 0 else ''}"
                        ),
                    }
                )
            return {
                "item_id": current_id,
                "is_raw_material": True,
                "descricao": str(item.get("descricao", "") or "").strip(),
                "material": str(raw_fields.get("material", "") or item.get("material", "") or "").strip(),
                "espessura": str(raw_fields.get("espessura", "") or item.get("espessura", "") or "").strip(),
                "dimensao": str(raw_fields.get("dimensao", "") or item.get("dimensao", item.get("dimensoes", "")) or "").strip(),
                "qtd_pendente": pending,
                "options": options,
            }
        raise ValueError("Linha de matéria-prima não encontrada.")

    def operator_consume_montagem_stock(
        self,
        numero: str,
        operador: str = "",
        item_ids: list[str] | None = None,
        material_allocations: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda não encontrada.")
        items = list(enc.get("montagem_itens", []) or [])
        if not items:
            raise ValueError("Esta encomenda nao tem componentes de montagem.")
        selected_ids = {str(value or "").strip() for value in list(item_ids or []) if str(value or "").strip()}
        allocation_by_item = {
            str(key or "").strip(): dict(value or {})
            for key, value in dict(material_allocations or {}).items()
            if str(key or "").strip()
        }
        actor = str(operador or (self.user or {}).get("username", "") or "Sistema").strip() or "Sistema"
        product_map = {
            str(prod.get("codigo", "") or "").strip(): prod
            for prod in list(self.ensure_data().get("produtos", []) or [])
            if str(prod.get("codigo", "") or "").strip()
        }
        consumable = []
        shortages: list[str] = []
        for item_index, item in enumerate(items):
            item_id = f"COMP::{numero}::{item_index}"
            if selected_ids and item_id not in selected_ids:
                continue
            item_type = self.desktop_main.normalize_orc_line_type(item.get("tipo_item"))
            is_raw = self._montagem_item_is_raw_material(item)
            if item_type != self.desktop_main.ORC_LINE_TYPE_PRODUCT and not is_raw:
                continue
            plan = self._parse_float(item.get("qtd_planeada", item.get("qtd", 0)), 0)
            done = self._parse_float(item.get("qtd_consumida", 0), 0)
            pending = max(0.0, plan - done)
            if pending <= 1e-9:
                continue
            consumable.append(item)
            if item_type == self.desktop_main.ORC_LINE_TYPE_PRODUCT:
                code = str(item.get("produto_codigo", "") or "").strip()
                product = product_map.get(code)
                if product is None:
                    shortages.append(f"{code or '-'}: produto nao encontrado")
                    continue
                if bool(product.get("quality_blocked")) or (str(product.get("quality_status", "") or "").strip() and not self._quality_status_is_available(product.get("quality_status", ""))):
                    shortages.append(f"{code}: bloqueado pela qualidade")
                    continue
                available = self._parse_float(product.get("qty", 0), 0)
                if pending > available + 1e-9:
                    shortages.append(f"{code}: faltam {pending - available:.2f} ({available:.2f} disponivel)")
            elif is_raw:
                stock_id = str(item.get("stock_material_id", "") or "").strip()
                raw_fields = self._montagem_item_raw_material_fields(item)
                allocation_payload = allocation_by_item.get(item_id, {})
                if allocation_payload:
                    available = 0.0
                    for alloc in list(allocation_payload.get("allocations", []) or []):
                        alloc_id = str((alloc or {}).get("material_id", "") or "").strip()
                        alloc_qty = self._parse_float((alloc or {}).get("quantidade", 0), 0)
                        if alloc_qty <= 0:
                            continue
                        material = self.material_by_id(alloc_id)
                        if material is None:
                            shortages.append(f"{alloc_id}: matéria-prima nao encontrada")
                            continue
                        if self._material_quality_is_blocked(material):
                            shortages.append(f"{alloc_id}: bloqueado pela qualidade")
                            continue
                        stock_candidate = {
                            "material": material.get("material", ""),
                            "espessura": material.get("espessura", ""),
                            "comprimento": material.get("comprimento", 0),
                            "largura": material.get("largura", 0),
                        }
                        if not self._montagem_material_candidate_matches_dimension(item, stock_candidate):
                            shortages.append(f"{alloc_id}: não corresponde à dimensão pedida")
                            continue
                        stock_available = self._parse_float(material.get("quantidade", 0), 0) - self._parse_float(material.get("reservado", 0), 0)
                        if alloc_qty > stock_available + 1e-9:
                            shortages.append(f"{alloc_id}: faltam {alloc_qty - stock_available:.2f} ({stock_available:.2f} disponivel)")
                            continue
                        available += alloc_qty
                elif stock_id:
                    material = self.material_by_id(stock_id)
                    if material is None:
                        shortages.append(f"{stock_id}: matéria-prima nao encontrada")
                        continue
                    if self._material_quality_is_blocked(material):
                        shortages.append(f"{stock_id}: bloqueado pela qualidade")
                        continue
                    stock_candidate = {
                        "material": material.get("material", ""),
                        "espessura": material.get("espessura", ""),
                        "comprimento": material.get("comprimento", 0),
                        "largura": material.get("largura", 0),
                    }
                    if self._montagem_material_candidate_matches_dimension(item, stock_candidate):
                        available = self._parse_float(material.get("quantidade", 0), 0) - self._parse_float(material.get("reservado", 0), 0)
                    else:
                        available = 0.0
                else:
                    shortages.append(f"{item.get('descricao', '-')}: seleciona a unidade física a consumir")
                    available = 0.0
                if pending > available + 1e-9:
                    shortages.append(f"{stock_id or item.get('descricao', '-')}: faltam {pending - available:.2f} ({available:.2f} disponivel)")
        if not consumable:
            raise ValueError("Nao existem produtos/componentes de stock pendentes para consumir.")
        if shortages:
            raise ValueError("Stock insuficiente para consumir componentes:\n" + "\n".join(shortages))
        now_txt = self.desktop_main.now_iso()
        for item in consumable:
            item_type = self.desktop_main.normalize_orc_line_type(item.get("tipo_item"))
            plan = self._parse_float(item.get("qtd_planeada", item.get("qtd", 0)), 0)
            done = self._parse_float(item.get("qtd_consumida", 0), 0)
            pending = max(0.0, plan - done)
            if item_type == self.desktop_main.ORC_LINE_TYPE_PRODUCT:
                code = str(item.get("produto_codigo", "") or "").strip()
                product = product_map.get(code)
                if product is None:
                    continue
                if bool(product.get("quality_blocked")) or (str(product.get("quality_status", "") or "").strip() and not self._quality_status_is_available(product.get("quality_status", ""))):
                    raise ValueError(f"Produto {code} bloqueado pela qualidade.")
                before = self._parse_float(product.get("qty", 0), 0)
                product["qty"] = max(0.0, before - pending)
                product["atualizado_em"] = now_txt
                self.desktop_main.add_produto_mov(
                    self.ensure_data(),
                    tipo="BAIXA_MONTAGEM",
                    operador=actor,
                    codigo=code,
                    descricao=str(item.get("descricao", "") or product.get("descricao", "") or "").strip(),
                    qtd=pending,
                    antes=before,
                    depois=product["qty"],
                    obs=f"Montagem da encomenda {numero}",
                    origem="OPERADOR_MONTAGEM",
                    ref_doc=numero,
                )
            elif self._montagem_item_is_raw_material(item):
                stock_id = str(item.get("stock_material_id", "") or "").strip()
                item_id = f"COMP::{numero}::{items.index(item)}"
                allocation_payload = allocation_by_item.get(item_id, {})
                allocations: list[dict[str, Any]] = []
                if allocation_payload:
                    allocations = [
                        {
                            "material_id": str((row or {}).get("material_id", "") or "").strip(),
                            "quantidade": self._parse_float((row or {}).get("quantidade", 0), 0),
                        }
                        for row in list(allocation_payload.get("allocations", []) or [])
                        if str((row or {}).get("material_id", "") or "").strip() and self._parse_float((row or {}).get("quantidade", 0), 0) > 0
                    ]
                    allocated_total = round(sum(self._parse_float(row.get("quantidade", 0), 0) for row in allocations), 4)
                    if abs(allocated_total - pending) > 0.0001:
                        raise ValueError(f"{item.get('descricao', '-')}: a seleção física tem de totalizar {pending:.2f}.")
                    result = self.consume_material_allocations(
                        allocations,
                        retalho=dict(allocation_payload.get("retalho", {}) or {}),
                        source_material_id=str(allocation_payload.get("source_material_id", "") or "").strip(),
                        reason=f"operador_montagem_{numero}",
                    )
                elif stock_id:
                    material = self.material_by_id(stock_id)
                    if material is None:
                        raise ValueError(f"Matéria-prima {stock_id} não encontrada.")
                    if self._material_quality_is_blocked(material):
                        raise ValueError(f"Matéria-prima {stock_id} bloqueada pela qualidade.")
                    stock_candidate = {
                        "material": material.get("material", ""),
                        "espessura": material.get("espessura", ""),
                        "comprimento": material.get("comprimento", 0),
                        "largura": material.get("largura", 0),
                    }
                    if not self._montagem_material_candidate_matches_dimension(item, stock_candidate):
                        raise ValueError(f"Matéria-prima {stock_id} não corresponde à dimensão pedida.")
                    allocations.append({"material_id": stock_id, "quantidade": pending})
                    result = self.consume_material_allocations(allocations, reason=f"operador_montagem_{numero}")
                else:
                    raise ValueError(f"{item.get('descricao', '-')}: seleciona a unidade física a consumir.")
                if not str(item.get("stock_material_id", "") or "").strip() and allocations:
                    item["stock_material_id"] = str(allocations[0].get("material_id", "") or "").strip()
                item["stock_consumption"] = {
                    "consumed_total": round(self._parse_float(result.get("consumed_total", pending), pending), 2),
                    "used_lots": list(result.get("used_lots", []) or []),
                }
            item["qtd_consumida"] = round(plan, 2)
            item["estado"] = "Consumido"
            item["consumed_at"] = now_txt
            item["consumed_by"] = actor
        self.desktop_main.update_estado_encomenda_por_espessuras(enc)
        self._save(force=True)
        return self.order_detail(numero)

    def operator_scan_code(self, code: str, current_posto: str = "Geral") -> dict[str, Any]:
        def _norm_of_scan(value: str) -> str:
            text = str(value or "").strip()
            match = re.fullmatch(r"\s*OF\s*[-_'/\\\s]+\s*(\d{4})\s*[-_'/\\\s]+\s*(\d{1,8})\s*", text, flags=re.IGNORECASE)
            if match:
                return f"OF-{match.group(1)}-{match.group(2).zfill(4)[-4:]}"
            return text

        def _norm_opp_scan(value: str) -> str:
            text = str(value or "").strip()
            match = re.fullmatch(r"\s*OPP\s*[-_'/\\\s]+\s*(\d{4})\s*[-_'/\\\s]+\s*(\d{1,8})\s*[-_'/\\\s]+\s*(\d{1,4})\s*", text, flags=re.IGNORECASE)
            if match:
                return f"OPP-{match.group(1)}-{match.group(2).zfill(4)[-4:]}-{match.group(3).zfill(2)[-2:]}"
            return text

        def _parse_grp_scan(value: str) -> tuple[str, str, str] | None:
            text = str(value or "").strip()
            if not text.upper().startswith("GRP"):
                return None
            compact = re.sub(r"\s+", "", text)
            match = re.search(
                r"GRP.*?O?F[^0-9]*(\d{4})[^0-9]+(\d{1,8})(.*)$",
                compact,
                flags=re.IGNORECASE,
            )
            if not match:
                return None
            of_code = f"OF-{match.group(1)}-{match.group(2).zfill(4)[-4:]}"
            tail = str(match.group(3) or "").strip()
            tail = tail.replace("Ô", "|").replace("ô", "|").replace("^", "|").replace("¦", "|").replace(";", "|")
            tail = re.sub(r"^[^A-Za-z0-9]+", "", tail)
            tokens = [token.strip() for token in re.split(r"[^A-Za-z0-9.,_-]+", tail) if token.strip()]
            if len(tokens) < 2:
                return None
            return of_code, tokens[0], tokens[1]

        def _parse_opr_scan(value: str) -> tuple[str, str] | None:
            text = str(value or "").strip()
            if not text.upper().startswith("OPR"):
                return None
            compact = re.sub(r"\s+", "", text)
            match = re.search(
                r"OPR.*?O?PP[^0-9]*(\d{4})[^0-9]+(\d{1,8})[^0-9]+(\d{1,4})(.*)$",
                compact,
                flags=re.IGNORECASE,
            )
            if not match:
                return None
            opp_code = f"OPP-{match.group(1)}-{match.group(2).zfill(4)[-4:]}-{match.group(3).zfill(2)[-2:]}"
            tail = str(match.group(4) or "").strip()
            tail = re.sub(r"^[Êê]+", "E", tail)
            tail = tail.replace("Ô", "|").replace("ô", "|").replace("^", "|").replace("¦", "|").replace(";", "|")
            tail = re.sub(r"^[^A-Za-z0-9]+", "", tail)
            tokens = [token.strip() for token in re.split(r"[^A-Za-z0-9.,_-]+", tail) if token.strip()]
            if not tokens:
                return None
            return opp_code, tokens[0]

        code_txt = (
            str(code or "")
            .replace("\r", "")
            .replace("\n", "")
            .replace("\t", "")
            .strip()
        )
        if not code_txt:
            raise ValueError("Código vazio.")
        grp_scan = _parse_grp_scan(code_txt)
        opr_scan = _parse_opr_scan(code_txt)
        if grp_scan is not None:
            code_txt = f"GRP|{grp_scan[0]}|{grp_scan[1]}|{grp_scan[2]}"
        elif opr_scan is not None:
            code_txt = f"OPR|{opr_scan[0]}|{opr_scan[1]}"
        else:
            prefix = code_txt[:3].upper()
            if prefix == "GRP":
                raise ValueError("Código de espessura inválido.")
            if prefix == "OPR":
                raise ValueError("Código de operação inválido.")
        code_txt = code_txt.replace("Ô", "|").replace("ô", "|").replace("^", "|")
        code_txt = re.sub(r"\s*\|\s*", "|", code_txt)
        code_txt = code_txt.replace("¦", "|").replace(";", "|")
        if code_txt.upper().startswith("GRP") and not code_txt.upper().startswith("GRP|"):
            of_pos_match = re.search(r"OF\s*[-_'/\\\s]+\s*\d{4}\s*[-_'/\\\s]+\s*\d{1,8}", code_txt, flags=re.IGNORECASE)
            if of_pos_match:
                code_txt = "GRP|" + code_txt[of_pos_match.start():]
            else:
                f_pos_match = re.search(r"F\s*[-_'/\\\s]+\s*\d{4}\s*[-_'/\\\s]+\s*\d{1,8}", code_txt, flags=re.IGNORECASE)
                if f_pos_match:
                    code_txt = "GRP|O" + code_txt[f_pos_match.start():]
        if code_txt.upper().startswith("GRP|"):
            parts = code_txt.split("|")
            if len(parts) >= 4:
                parts[1] = _norm_of_scan(parts[1])
                code_txt = "|".join(parts)
        elif code_txt.upper().startswith("COMP|") or code_txt.upper().startswith("CPI|"):
            parts = code_txt.split("|")
            if len(parts) >= 2:
                parts[1] = _norm_of_scan(parts[1])
                code_txt = "|".join(parts)
        elif code_txt.upper().startswith("OPR|"):
            parts = code_txt.split("|")
            if len(parts) >= 2:
                parts[1] = _norm_opp_scan(parts[1])
                code_txt = "|".join(parts)
        else:
            code_txt = _norm_opp_scan(_norm_of_scan(code_txt))
        posto_txt = str(current_posto or "").strip() or "Geral"
        if code_txt.upper().startswith("OPR|"):
            parts = code_txt.split("|")
            if len(parts) < 3:
                raise ValueError("Código de operação inválido.")
            opp_code = str(parts[1] or "").strip()
            op_token = str(parts[2] or "").strip()
            enc, piece = self._find_piece_by_opp(opp_code)
            ctx = self.operator_piece_context(str(enc.get("numero", "") or ""), str(piece.get("id", "") or ""))
            pending_ops = list(ctx.get("pending_ops", []) or [])
            all_ops = [
                str(self.desktop_main.normalize_operacao_nome(op.get("nome", "")) or op.get("nome", "") or "").strip()
                for op in list(self.desktop_main.ensure_peca_operacoes(piece) or [])
                if str(op.get("nome", "") or "").strip()
            ]
            candidates = all_ops or pending_ops
            selected_op = ""
            token_norm = self.desktop_main.norm_text(op_token).casefold()
            token_norm = {
                "cl": "cl",
                "q": "q",
                "mb": "emb",
                "mb.": "emb",
                "emb": "emb",
                "emb.": "emb",
            }.get(token_norm, token_norm)
            explicit_tokens = {
                "cl": {"corte laser", "laser"},
                "q": {"quinagem"},
                "emb": {"embalamento"},
            }
            for op_name in candidates:
                abbrev_norm = self.desktop_main.norm_text(self._operation_pdf_abbrev(op_name)).casefold()
                abbrev_norm = abbrev_norm.rstrip(".")
                op_norm = self.desktop_main.norm_text(op_name).casefold()
                token_key = token_norm.rstrip(".")
                if token_key in explicit_tokens and op_norm in explicit_tokens[token_key]:
                    selected_op = op_name
                    break
                if token_key and token_key in {abbrev_norm, op_norm}:
                    selected_op = op_name
                    break
            if not selected_op:
                raise ValueError("Operação do código de barras não encontrada nesta OPP.")
            return {
                "tipo": "OPR",
                "encomenda_numero": str(enc.get("numero", "") or "").strip(),
                "piece_id": str(piece.get("id", "") or "").strip(),
                "opp": str(piece.get("opp", "") or "").strip(),
                "of": str(piece.get("of", "") or "").strip(),
                "posto": posto_txt,
                "operacao": selected_op,
                "pending_ops": pending_ops,
                "context": ctx,
            }
        if code_txt.upper().startswith("GRP|"):
            parts = code_txt.split("|")
            if len(parts) >= 4:
                of_code = parts[1].strip()
                material = parts[2].strip()
                espessura = parts[3].strip()
                enc = next((row for row in list(self.ensure_data().get("encomendas", []) or []) if self._order_of_code(row, create=False) == of_code), None)
                if enc is None:
                    raise ValueError("Grupo/espessura não encontrado.")
                return {
                    "tipo": "GRP",
                    "encomenda_numero": str(enc.get("numero", "") or "").strip(),
                    "of": of_code,
                    "material": material,
                    "espessura": espessura,
                }
            raise ValueError("Código de espessura inválido.")
        if code_txt.upper().startswith("COMP|"):
            parts = code_txt.split("|", 1)
            of_code = parts[1].strip() if len(parts) > 1 else ""
            enc = next((row for row in list(self.ensure_data().get("encomendas", []) or []) if self._order_of_code(row, create=False) == of_code), None)
            if enc is None:
                raise ValueError("Grupo de componentes não encontrado.")
            group = self.operator_montagem_stock_group(str(enc.get("numero", "") or ""))
            if not group:
                raise ValueError("Esta OF nao tem componentes de stock associados.")
            return {
                "tipo": "COMP",
                "encomenda_numero": str(enc.get("numero", "") or "").strip(),
                "of": of_code,
                "material": str(group.get("material", "") or "").strip(),
                "espessura": str(group.get("espessura", "") or "").strip(),
            }
        if code_txt.upper().startswith("CPI|"):
            parts = code_txt.split("|")
            if len(parts) < 3:
                raise ValueError("Código individual de componente inválido.")
            of_code = parts[1].strip()
            item_index = int(self._parse_float(parts[2], -1))
            enc = next((row for row in list(self.ensure_data().get("encomendas", []) or []) if self._order_of_code(row, create=False) == of_code), None)
            if enc is None:
                raise ValueError("Componente não encontrado.")
            group = self.operator_montagem_stock_group(str(enc.get("numero", "") or ""))
            if not group:
                raise ValueError("Esta OF nao tem componentes de stock associados.")
            rows = list(group.get("montagem_items", []) or [])
            if item_index < 0 or item_index >= len(rows):
                raise ValueError("Item de componente não encontrado.")
            item = dict(rows[item_index] or {})
            return {
                "tipo": "CPI",
                "encomenda_numero": str(enc.get("numero", "") or "").strip(),
                "of": of_code,
                "item_id": str(item.get("id", "") or "").strip(),
                "codigo": str(item.get("codigo", "") or "").strip(),
                "descricao": str(item.get("descricao", "") or "").strip(),
                "material": str(group.get("material", "") or "").strip(),
                "espessura": str(item.get("espessura", group.get("espessura", "")) or "").strip(),
            }
        if code_txt.upper().startswith("OF-"):
            enc = next((row for row in list(self.ensure_data().get("encomendas", []) or []) if self._order_of_code(row, create=False) == code_txt), None)
            if enc is None:
                raise ValueError("OF não encontrada.")
            detail = self.order_detail(str(enc.get("numero", "") or ""))
            return {"tipo": "OF", "encomenda": detail, "pieces": list(detail.get("pieces", []) or [])}
        enc, piece = self._find_piece_by_opp(code_txt)
        ctx = self.operator_piece_context(str(enc.get("numero", "") or ""), str(piece.get("id", "") or ""))
        pending_ops = list(ctx.get("pending_ops", []) or [])
        selected_op = ""
        posto_norm = self.desktop_main.norm_text(posto_txt)
        group_for_resource = self.workcenter_group_for_resource(posto_txt)
        for op_name in pending_ops:
            op_posto = self._operator_posto_for_operation(op_name)
            candidates = {self.desktop_main.norm_text(op_posto), self.desktop_main.norm_text(self.workcenter_group_for_resource(posto_txt, op_name)), self.desktop_main.norm_text(group_for_resource)}
            if posto_norm in candidates or any(token and token in self.desktop_main.norm_text(op_name) for token in posto_norm.split()):
                selected_op = op_name
                break
        if not selected_op and pending_ops:
            selected_op = pending_ops[0]
        return {
            "tipo": "OPP",
            "encomenda_numero": str(enc.get("numero", "") or "").strip(),
            "piece_id": str(piece.get("id", "") or "").strip(),
            "opp": str(piece.get("opp", "") or "").strip(),
            "of": str(piece.get("of", "") or "").strip(),
            "posto": posto_txt,
            "operacao": selected_op,
            "pending_ops": pending_ops,
            "context": ctx,
        }

    def order_consume_montagem(self, numero: str, operador: str = "") -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda não encontrada.")
        items = list(enc.get("montagem_itens", []) or [])
        if not items:
            raise ValueError("Esta encomenda nao tem itens de montagem.")
        actor = str(operador or (self.user or {}).get("username", "") or "Sistema").strip() or "Sistema"
        product_map = {
            str(prod.get("codigo", "") or "").strip(): prod
            for prod in list(self.ensure_data().get("produtos", []) or [])
            if str(prod.get("codigo", "") or "").strip()
        }
        shortages: list[str] = []
        for item in items:
            code = str(item.get("produto_codigo", "") or "").strip()
            plan = self._parse_float(item.get("qtd_planeada", item.get("qtd", 0)), 0)
            done = self._parse_float(item.get("qtd_consumida", 0), 0)
            pending = max(0.0, plan - done)
            if pending <= 1e-9:
                continue
            if self._montagem_item_is_raw_material(item):
                stock_id = str(item.get("stock_material_id", "") or "").strip()
                if stock_id:
                    material = self.material_by_id(stock_id)
                    if material is None:
                        shortages.append(f"{stock_id}: matéria-prima nao encontrada")
                        continue
                    if self._material_quality_is_blocked(material):
                        shortages.append(f"{stock_id}: bloqueado pela qualidade")
                        continue
                    available = self._parse_float(material.get("quantidade", 0), 0) - self._parse_float(material.get("reservado", 0), 0)
                else:
                    material_txt = str(item.get("material", "") or "").strip()
                    esp_txt = str(item.get("espessura", "") or "").strip()
                    candidates = self.material_candidates(material_txt, esp_txt) if material_txt and esp_txt else []
                    available = sum(self._parse_float(row.get("disponivel", 0), 0) for row in candidates)
                if pending > available + 1e-9:
                    label = stock_id or str(item.get("descricao", "") or item.get("material", "") or "-").strip()
                    shortages.append(f"{label}: faltam {pending - available:.2f} ({available:.2f} disponivel)")
                continue
            if self.desktop_main.normalize_orc_line_type(item.get("tipo_item")) != self.desktop_main.ORC_LINE_TYPE_PRODUCT:
                continue
            product = product_map.get(code)
            if product is None:
                shortages.append(f"{code or '-'}: produto nao encontrado")
                continue
            if bool(product.get("quality_blocked")) or (str(product.get("quality_status", "") or "").strip() and not self._quality_status_is_available(product.get("quality_status", ""))):
                shortages.append(f"{code}: bloqueado pela qualidade")
                continue
            available = self._parse_float(product.get("qty", 0), 0)
            if pending > available + 1e-9:
                shortages.append(f"{code}: faltam {pending - available:.2f} ({available:.2f} disponivel)")
        if shortages:
            raise ValueError("Stock insuficiente para concluir a montagem:\n" + "\n".join(shortages))
        changed = False
        now_txt = self.desktop_main.now_iso()
        for item in items:
            item_type = self.desktop_main.normalize_orc_line_type(item.get("tipo_item"))
            plan = self._parse_float(item.get("qtd_planeada", item.get("qtd", 0)), 0)
            done = self._parse_float(item.get("qtd_consumida", 0), 0)
            pending = max(0.0, plan - done)
            if pending <= 1e-9:
                continue
            if item_type == self.desktop_main.ORC_LINE_TYPE_PRODUCT:
                code = str(item.get("produto_codigo", "") or "").strip()
                product = product_map.get(code)
                if product is None:
                    continue
                if bool(product.get("quality_blocked")) or (str(product.get("quality_status", "") or "").strip() and not self._quality_status_is_available(product.get("quality_status", ""))):
                    raise ValueError(f"Produto {code} bloqueado pela qualidade.")
                before = self._parse_float(product.get("qty", 0), 0)
                after = max(0.0, before - pending)
                product["qty"] = after
                product["atualizado_em"] = now_txt
                self.desktop_main.add_produto_mov(
                    self.ensure_data(),
                    tipo="BAIXA_MONTAGEM",
                    operador=actor,
                    codigo=code,
                    descricao=str(item.get("descricao", "") or product.get("descricao", "") or "").strip(),
                    qtd=pending,
                    antes=before,
                    depois=after,
                    obs=f"Montagem da encomenda {numero}",
                    origem="MONTAGEM",
                    ref_doc=numero,
                )
                item["qtd_consumida"] = round(plan, 2)
                item["estado"] = "Consumido"
                item["consumed_at"] = now_txt
                item["consumed_by"] = actor
                changed = True
            elif self._montagem_item_is_raw_material(item):
                stock_id = str(item.get("stock_material_id", "") or "").strip()
                allocations: list[dict[str, Any]] = []
                remaining = pending
                if stock_id:
                    allocations.append({"material_id": stock_id, "quantidade": remaining})
                else:
                    material_txt = str(item.get("material", "") or "").strip()
                    esp_txt = str(item.get("espessura", "") or "").strip()
                    for candidate in self.material_candidates(material_txt, esp_txt) if material_txt and esp_txt else []:
                        if remaining <= 1e-9:
                            break
                        available = self._parse_float(candidate.get("disponivel", 0), 0)
                        qty = min(available, remaining)
                        if qty <= 1e-9:
                            continue
                        allocations.append({"material_id": str(candidate.get("material_id", "") or "").strip(), "quantidade": qty})
                        remaining = round(remaining - qty, 6)
                result = self.consume_material_allocations(
                    allocations,
                    reason=f"montagem_{numero}_{str(item.get('descricao', '') or item.get('material', '') or '').strip()}",
                )
                item["qtd_consumida"] = round(plan, 2)
                item["estado"] = "Consumido"
                item["consumed_at"] = now_txt
                item["consumed_by"] = actor
                if not str(item.get("stock_material_id", "") or "").strip() and allocations:
                    item["stock_material_id"] = str(allocations[0].get("material_id", "") or "").strip()
                item["stock_consumption"] = {
                    "consumed_total": round(self._parse_float(result.get("consumed_total", pending), pending), 2),
                    "used_lots": list(result.get("used_lots", []) or []),
                }
                changed = True
            elif item_type == self.desktop_main.ORC_LINE_TYPE_SERVICE:
                item["qtd_consumida"] = round(plan, 2)
                item["estado"] = "Concluido"
                item["consumed_at"] = now_txt
                item["consumed_by"] = actor
                changed = True
            else:
                # Componentes de conjunto sem produto nao consomem stock nem fecham aqui.
                continue
        if not changed:
            raise ValueError("Nao existem consumos pendentes de montagem.")
        self.desktop_main.update_estado_encomenda_por_espessuras(enc)
        self._save(force=True)
        return self.order_detail(numero)
