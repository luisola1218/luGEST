from __future__ import annotations

import re
from datetime import datetime
from lugest_core.laser.quote_engine import analyze_dxf_geometry
from pathlib import Path
from typing import Any


class OrdersBackendMixin:
    """Legacy adapter for orders; see BACKEND_GUIDE.md."""

    def _order_reserved_sheet(self, numero: str, material: str = "", espessura: str = "") -> str:
        enc = self.get_encomenda_by_numero(numero)
        if not enc or not list(enc.get("reservas", []) or []):
            return "-"
        mat_norm = str(material or "").strip().lower()
        esp_norm = self._planning_norm_esp(espessura)
        for row in list(enc.get("reservas", []) or []):
            if mat_norm and str(row.get("material", "") or "").strip().lower() != mat_norm:
                continue
            if esp_norm and self._planning_norm_esp(row.get("espessura", "")) != esp_norm:
                continue
            return f"{row.get('material', '')} {row.get('espessura', '')} ({row.get('quantidade', 0)})"
        return "-"

    def get_encomenda_by_numero(self, numero: str) -> dict[str, Any] | None:
        numero = str(numero or "").strip()
        enc = next((e for e in self.ensure_data().get("encomendas", []) if str(e.get("numero", "")).strip() == numero), None)
        if enc is not None:
            enc.setdefault("montagem_itens", [])
            self._ensure_unique_order_piece_refs(enc)
        return enc

    def _ensure_unique_order_piece_refs(self, enc: dict[str, Any]) -> bool:
        data = self.ensure_data()
        cliente_codigo = str((enc or {}).get("cliente", "") or "").strip()
        if not cliente_codigo:
            return False
        seen: set[str] = set()
        changed = False
        pieces = list(self.desktop_main.encomenda_pecas(enc) or [])
        for piece in pieces:
            current_ref = str(piece.get("ref_interna", "") or "").strip()
            needs_new = (not current_ref) or (current_ref in seen)
            if needs_new:
                new_ref = str(self.desktop_main.next_ref_interna_unique(data, cliente_codigo, list(seen)))
                piece["ref_interna"] = new_ref
                current_ref = new_ref
                changed = True
            seen.add(current_ref)
            if str(piece.get("ref_externa", "") or "").strip():
                try:
                    self.desktop_main.update_refs(data, current_ref, str(piece.get("ref_externa", "") or "").strip())
                except Exception:
                    pass
        if changed:
            self.desktop_main.update_estado_encomenda_por_espessuras(enc)
            self._save(force=True)
        return changed

    def order_detail(self, numero: str) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda não encontrada.")
        data = self.ensure_data()
        if self._ensure_order_fabrication_order(enc):
            self._save(force=True)
        cliente_codigo = str(enc.get("cliente", "") or "").strip()
        cliente_obj = {}
        find_cliente_fn = getattr(self.desktop_main, "find_cliente", None)
        if callable(find_cliente_fn) and cliente_codigo:
            cliente_obj = find_cliente_fn(data, cliente_codigo) or {}
        pieces = []
        for piece in self.desktop_main.encomenda_pecas(enc):
            ops = self.desktop_main.ensure_peca_operacoes(piece)
            qty_plan = self._parse_float(piece.get("quantidade_pedida", 0), 0)
            qty_prod = (
                self._parse_float(piece.get("produzido_ok", 0), 0)
                + self._parse_float(piece.get("produzido_nok", 0), 0)
                + self._parse_float(piece.get("produzido_qualidade", 0), 0)
            )
            pieces.append(
                {
                    "id": str(piece.get("id", "")).strip(),
                    "ref_interna": str(piece.get("ref_interna", "")).strip(),
                    "ref_externa": str(piece.get("ref_externa", "")).strip(),
                    "material": str(piece.get("material", "")).strip(),
                    "tipo_material": str(piece.get("tipo_material", "") or "CHAPA").strip(),
                    "subtipo_material": str(piece.get("subtipo_material", "") or piece.get("material", "")).strip(),
                    "espessura": str(piece.get("espessura", "")).strip(),
                    "dimensao": str(piece.get("dimensao", piece.get("dimensoes", "")) or "").strip(),
                    "perfil_tipo": str(piece.get("perfil_tipo", "") or "").strip(),
                    "perfil_tamanho": str(piece.get("perfil_tamanho", "") or "").strip(),
                    "comprimento_mm": self._parse_float(piece.get("comprimento_mm", 0), 0),
                    "tubo_forma": str(piece.get("tubo_forma", "") or "").strip(),
                    "lado_a": self._parse_float(piece.get("lado_a", 0), 0),
                    "lado_b": self._parse_float(piece.get("lado_b", 0), 0),
                    "tubo_espessura": self._parse_float(piece.get("tubo_espessura", 0), 0),
                    "diametro": self._parse_float(piece.get("diametro", 0), 0),
                    "estado": str(piece.get("estado", "")).strip(),
                    "qtd_plan": self._fmt(qty_plan),
                    "qtd_prod": self._fmt(qty_prod),
                    "descricao": str(piece.get("descricao", "") or piece.get("Observacoes", "") or "").strip(),
                    "of": str(piece.get("of", "") or "").strip(),
                    "opp": str(piece.get("opp", "") or "").strip(),
                    "peso_unid": round(self._parse_float(piece.get("peso_unid", 0), 0), 4),
                    "tempo_peca_min": round(self._parse_float(piece.get("tempo_peca_min", piece.get("tempo_pecas_min", 0)), 0), 3),
                    "revisao": str(piece.get("revisao", piece.get("rev", "")) or "").strip(),
                    "prioridade": str(piece.get("prioridade", piece.get("priority", "")) or "").strip(),
                    "conjunto_codigo": str(piece.get("conjunto_codigo", "") or "").strip(),
                    "conjunto_nome": str(piece.get("conjunto_nome", "") or "").strip(),
                    "grupo_uuid": str(piece.get("grupo_uuid", "") or "").strip(),
                    "operacoes": " + ".join(
                        [self.desktop_main.normalize_operacao_nome(op.get("nome", "")) for op in list(ops or []) if str(op.get("nome", "")).strip()]
                    ),
                    "desenho": bool(str(piece.get("desenho", "") or piece.get("desenho_path", "") or "").strip()),
                    "desenho_path": str(piece.get("desenho", "") or piece.get("desenho_path", "") or "").strip(),
                    "desenho_pdf": str(piece.get("desenho_pdf", "") or "").strip(),
                    "desenhos_pdf": [
                        str(item or "").strip()
                        for item in list(piece.get("desenhos_pdf", []) or [])
                        if str(item or "").strip()
                    ],
                    "ficheiros": [
                        str(item or "").strip()
                        for item in list(piece.get("ficheiros", []) or [])
                        if str(item or "").strip()
                    ],
                }
            )
        materials = []
        materials_tree = []
        for mat in list(enc.get("materiais", []) or []):
            esp_rows = []
            for esp in list(mat.get("espessuras", []) or []):
                op_times = self._planning_operation_times_map(esp)
                machine_map = self._order_esp_machine_map(esp)
                planning_ops = [op for op in self._planning_ops_from_esp_obj(esp) if op != "Montagem"]
                other_ops = [op for op in planning_ops if op != "Corte Laser"]
                ops_summary = []
                for op_name in other_ops:
                    op_value = str(op_times.get(op_name, "") or "").strip()
                    if op_value:
                        ops_summary.append(f"{op_name}: {self._fmt(op_value)} min")
                    else:
                        ops_summary.append(op_name)
                resource_summary = []
                for op_name in planning_ops:
                    resource_txt = str(machine_map.get(op_name, "") or "").strip()
                    if resource_txt:
                        resource_summary.append(f"{op_name}: {resource_txt}")
                esp_row = {
                    "material": str(mat.get("material", "")).strip(),
                    "espessura": str(esp.get("espessura", "")).strip(),
                    "estado": str(esp.get("estado", "")).strip(),
                    "tempo_min": self._fmt(esp.get("tempo_min", 0)),
                    "tempos_operacao": {op: self._fmt(value) for op, value in op_times.items() if str(value or "").strip()},
                    "operacoes_planeamento": planning_ops,
                    "tempo_operacoes_txt": " | ".join(ops_summary) or "-",
                    "maquinas_operacao": machine_map,
                    "recursos_operacao_txt": " | ".join(resource_summary) or "-",
                    "pecas": len(list(esp.get("pecas", []) or [])),
                }
                materials.append(esp_row)
                esp_rows.append(esp_row)
            materials_tree.append(
                {
                    "material": str(mat.get("material", "")).strip(),
                    "estado": str(mat.get("estado", "")).strip(),
                    "espessuras": esp_rows,
                }
            )
        montagem_items = []
        for item in list(enc.get("montagem_itens", []) or []):
            item_type = self.desktop_main.normalize_orc_line_type(item.get("tipo_item"))
            plan = round(self._parse_float(item.get("qtd_planeada", item.get("qtd", 0)), 0), 2)
            consumed = round(self._parse_float(item.get("qtd_consumida", 0), 0), 2)
            montagem_items.append(
                {
                    "tipo_item": item_type,
                    "stock_item_kind": str(item.get("stock_item_kind", "") or "").strip(),
                    "tipo_label": "Matéria-prima" if self._montagem_item_is_raw_material(item) else self.desktop_main.orc_line_type_label(item_type),
                    "item_key": self._montagem_item_key(item),
                    "descricao": str(item.get("descricao", "") or "").strip(),
                    "produto_codigo": str(item.get("produto_codigo", "") or "").strip(),
                    "produto_unid": str(item.get("produto_unid", "") or "").strip(),
                    "material": str(item.get("material", "") or "").strip(),
                    "espessura": str(item.get("espessura", "") or "").strip(),
                    "dimensao": str(item.get("dimensao", item.get("dimensoes", "")) or "").strip(),
                    "stock_material_id": str(item.get("stock_material_id", "") or "").strip(),
                    "qtd_planeada": plan,
                    "qtd_consumida": consumed,
                    "qtd_pendente": round(max(0.0, plan - consumed), 2),
                    "tempo_total_min": round(self._parse_float(item.get("tempo_total_min", item.get("tempo_min", 0)), 0), 2),
                    "preco_unit": round(self._parse_float(item.get("preco_unit", 0), 0), 4),
                    "conjunto_codigo": str(item.get("conjunto_codigo", "") or "").strip(),
                    "conjunto_nome": str(item.get("conjunto_nome", "") or "").strip(),
                    "grupo_uuid": str(item.get("grupo_uuid", "") or "").strip(),
                    "estado": str(item.get("estado", "") or "").strip(),
                    "created_at": str(item.get("created_at", "") or "").strip(),
                    "consumed_at": str(item.get("consumed_at", "") or "").strip(),
                    "consumed_by": str(item.get("consumed_by", "") or "").strip(),
                }
            )
        montagem_estado = str(self.desktop_main.encomenda_montagem_estado(enc) or "Nao aplicavel")
        montagem_shortages = self._order_montagem_shortages(enc)
        montagem_tempo_min = round(self._parse_float(self.desktop_main.encomenda_montagem_tempo_min(enc), 0), 2)
        montagem_resumo = str(self.desktop_main.encomenda_montagem_resumo(enc) or "").strip()
        return {
            "numero": str(enc.get("numero", "")).strip(),
            "tipo_encomenda": str(enc.get("tipo_encomenda", "") or "Cliente").strip(),
            "of_codigo": str(enc.get("of_codigo", "") or "").strip(),
            "ordem_fabrico": dict(enc.get("ordem_fabrico", {}) or {}),
            "cliente": cliente_codigo,
            "cliente_nome": str(cliente_obj.get("nome", "") or "").strip(),
            "posto_trabalho": self._order_workcenter(enc),
            "estado": str(enc.get("estado", "")).strip(),
            "data_entrega": str(enc.get("data_entrega", "")).strip(),
            "zona_transporte": self._transport_zone_for_order(enc, cliente_obj),
            "local_descarga": str(enc.get("local_descarga", "") or cliente_obj.get("morada", "") or "").strip(),
            "nota_cliente": str(enc.get("nota_cliente", "")).strip(),
            "nota_transporte": str(enc.get("nota_transporte", "") or "").strip(),
            "preco_transporte": round(self._parse_float(enc.get("preco_transporte", 0), 0), 2),
            "custo_transporte": round(self._parse_float(enc.get("custo_transporte", 0), 0), 2),
            "paletes": round(self._parse_float(enc.get("paletes", 0), 0), 2),
            "peso_bruto_kg": round(self._parse_float(enc.get("peso_bruto_kg", 0), 0), 2),
            "volume_m3": round(self._parse_float(enc.get("volume_m3", 0), 0), 3),
            "transportadora_id": str(enc.get("transportadora_id", "") or "").strip(),
            "transportadora_nome": str(enc.get("transportadora_nome", "") or "").strip(),
            "referencia_transporte": str(enc.get("referencia_transporte", "") or "").strip(),
            "transporte_numero": str(enc.get("transporte_numero", "") or "").strip(),
            "estado_transporte": str(enc.get("estado_transporte", "") or "").strip(),
            "tempo_estimado": self._parse_float(enc.get("tempo_estimado", enc.get("tempo", 0)), 0),
            "observacoes": str(enc.get("Observacoes", "") or enc.get("Observa??es", "") or "").strip(),
            "cativar": bool(enc.get("cativar")),
            "numero_orcamento": str(enc.get("numero_orcamento", "") or "").strip(),
            "can_edit_structure": not bool(str(enc.get("numero_orcamento", "") or "").strip()),
            "montagem_estado": montagem_estado,
            "montagem_tempo_min": montagem_tempo_min,
            "montagem_resumo": montagem_resumo,
            "montagem_stock_ready": not bool(montagem_shortages),
            "montagem_shortages": montagem_shortages,
            "montagem_items": montagem_items,
            "produto_fichas": [
                dict(row or {})
                for row in list(enc.get("produto_fichas", []) or [])
                if isinstance(row, dict)
            ],
            "can_consume_montagem": any(
                (
                    self.desktop_main.normalize_orc_line_type(row.get("tipo_item")) in {self.desktop_main.ORC_LINE_TYPE_PRODUCT, self.desktop_main.ORC_LINE_TYPE_SERVICE}
                    or self._montagem_item_is_raw_material(row)
                )
                and self._parse_float(row.get("qtd_planeada", 0), 0) > self._parse_float(row.get("qtd_consumida", 0), 0)
                for row in montagem_items
            ),
            "reservas": [
                {
                    "material_id": str(row.get("material_id", "") or "").strip(),
                    "material": str(row.get("material", "") or "").strip(),
                    "espessura": str(row.get("espessura", "") or "").strip(),
                    "quantidade": self._fmt(row.get("quantidade", 0)),
                }
                for row in list(enc.get("reservas", []) or [])
            ],
            "obs_interrupcao": str(enc.get("obs_interrupcao", "")).strip(),
            "pieces": pieces,
            "materials": materials,
            "materials_tree": materials_tree,
        }

    def _order_is_orc_based(self, enc: dict[str, Any]) -> bool:
        return bool(str((enc or {}).get("numero_orcamento", "") or "").strip())

    def _order_find_material(self, enc: dict[str, Any], material: str) -> dict[str, Any] | None:
        material_txt = str(material or "").strip().lower()
        for row in list(enc.get("materiais", []) or []):
            if str(row.get("material", "") or "").strip().lower() == material_txt:
                return row
        return None

    def _order_find_espessura(self, enc: dict[str, Any], material: str, espessura: str) -> dict[str, Any] | None:
        mat = self._order_find_material(enc, material)
        if mat is None:
            return None
        esp_txt = str(espessura or "").strip()
        for row in list(mat.get("espessuras", []) or []):
            if str(row.get("espessura", "") or "").strip() == esp_txt:
                return row
        return None

    def _order_find_piece(self, enc: dict[str, Any], ref_interna: str, ref_externa: str = "") -> tuple[dict[str, Any] | None, dict[str, Any] | None, dict[str, Any] | None]:
        ref_int = str(ref_interna or "").strip()
        ref_ext = str(ref_externa or "").strip()
        for mat in list(enc.get("materiais", []) or []):
            for esp in list(mat.get("espessuras", []) or []):
                for piece in list(esp.get("pecas", []) or []):
                    piece_ref_int = str(piece.get("ref_interna", "") or "").strip()
                    piece_ref_ext = str(piece.get("ref_externa", "") or "").strip()
                    if ref_int and piece_ref_int == ref_int:
                        return mat, esp, piece
                    if (not ref_int) and ref_ext and piece_ref_ext == ref_ext:
                        return mat, esp, piece
        return None, None, None

    def _sync_order_piece_documents_from_quote(self, enc: dict[str, Any]) -> bool:
        """Backfill technical PDFs into existing order pieces from their source quote."""
        if not isinstance(enc, dict):
            return False
        data = self.ensure_data()
        enc_numero = str(enc.get("numero", "") or "").strip()
        orc_numero = str(enc.get("numero_orcamento", "") or "").strip()
        source_quotes = []
        for orc in list(data.get("orcamentos", []) or []):
            if not isinstance(orc, dict):
                continue
            if orc_numero and str(orc.get("numero", "") or "").strip() == orc_numero:
                source_quotes.append(orc)
                continue
            if enc_numero and str(orc.get("numero_encomenda", "") or "").strip() == enc_numero:
                source_quotes.append(orc)
        if not source_quotes:
            return False
        digest_cache: dict[str, str] = {}

        def norm(value: Any) -> str:
            text = str(value or "").strip()
            norm_fn = getattr(self.desktop_main, "norm_text", None)
            if callable(norm_fn):
                return str(norm_fn(text) or "").strip()
            return text.casefold()

        def doc_list(row: dict[str, Any]) -> list[str]:
            return self._piece_pdf_references(row, digest_cache=digest_cache)

        def keys_for(row: dict[str, Any]) -> list[tuple[str, ...]]:
            ref_int = norm(row.get("ref_interna", ""))
            ref_ext = norm(row.get("ref_externa", ""))
            material = norm(row.get("material", "") or row.get("material_subtype", ""))
            espessura = norm(row.get("espessura", ""))
            descricao = norm(row.get("descricao", "") or row.get("Observacoes", ""))
            keys: list[tuple[str, ...]] = []
            if ref_int:
                keys.append(("int", ref_int))
            if ref_ext:
                keys.append(("ext", ref_ext))
            if ref_ext and material and espessura:
                keys.append(("extmat", ref_ext, material, espessura))
            if descricao and material and espessura:
                keys.append(("desc", descricao, material, espessura))
            return keys

        quote_docs: dict[tuple[str, ...], list[str]] = {}
        for orc in source_quotes:
            for line in list(orc.get("linhas", []) or []):
                if not isinstance(line, dict):
                    continue
                docs = doc_list(line)
                if not docs:
                    continue
                for key in keys_for(line):
                    bucket = quote_docs.setdefault(key, [])
                    for doc in docs:
                        if doc not in bucket:
                            bucket.append(doc)
        if not quote_docs:
            return False

        changed = False
        for piece in self.desktop_main.encomenda_pecas(enc):
            if not isinstance(piece, dict):
                continue
            existing_docs = doc_list(piece)
            matched_docs: list[str] = []
            for key in keys_for(piece):
                for doc in quote_docs.get(key, []):
                    if doc not in matched_docs:
                        matched_docs.append(doc)
            merged_docs = list(existing_docs)
            for doc in matched_docs:
                if doc not in merged_docs:
                    merged_docs.append(doc)
            if not matched_docs:
                if self._apply_piece_pdf_references(piece, existing_docs, digest_cache=digest_cache):
                    changed = True
                else:
                    continue
            if self._apply_piece_pdf_references(piece, merged_docs, digest_cache=digest_cache):
                changed = True
            else:
                continue
        return changed

    def order_rows(self, filter_text: str = "", estado: str = "Ativas", ano: str = "Todos", cliente: str = "Todos") -> list[dict[str, Any]]:
        data = self.ensure_data()
        query = str(filter_text or "").strip().lower()
        estado_filter = str(estado or "Ativas").strip().lower()
        ano_filter = str(ano or "Todos").strip()
        cliente_filter = str(cliente or "Todos").strip()
        clientes_nome = {
            str(c.get("codigo", "") or "").strip(): str(c.get("nome", "") or "").strip()
            for c in list(data.get("clientes", []) or [])
            if isinstance(c, dict)
        }
        rows = []
        for enc in data.get("encomendas", []):
            pieces = list(self.desktop_main.encomenda_pecas(enc))
            planeado = sum(self._parse_float(p.get("quantidade_pedida", 0), 0) for p in pieces)
            produzido = sum(
                self._parse_float(p.get("produzido_ok", 0), 0)
                + self._parse_float(p.get("produzido_nok", 0), 0)
                + self._parse_float(p.get("produzido_qualidade", 0), 0)
                for p in pieces
            )
            montagem_estado = str(self.desktop_main.encomenda_montagem_estado(enc) or "")
            progress = round((produzido / planeado) * 100.0, 1) if planeado > 0 else (100.0 if montagem_estado == "Consumida" else 0.0)
            estado_txt = str(enc.get("estado", "") or "").strip()
            estado_norm = self.desktop_main.norm_text(estado_txt)
            enc_year = ""
            try:
                enc_year = str(
                    self.desktop_main._enc_extract_year(
                        enc.get("data_criacao", ""),
                        enc.get("data_entrega", ""),
                        enc.get("numero", ""),
                        enc.get("ano"),
                    )
                    or ""
                ).strip()
            except Exception:
                enc_year = ""
            cli_code = str(enc.get("cliente", "") or "").strip()
            cli_display = f"{cli_code} - {clientes_nome.get(cli_code, '')}".strip(" -")
            if ano_filter.lower() not in ("todos", "todas", "all", "") and enc_year != ano_filter:
                continue
            if cliente_filter.lower() not in ("todos", "todas", "all", "") and cli_code != cliente_filter.split(" - ", 1)[0].strip():
                continue
            if estado_filter not in ("todos", "todas", "all", ""):
                if "ativ" in estado_filter and "concl" in estado_norm:
                    continue
                if "prepar" in estado_filter and "prepar" not in estado_norm:
                    continue
                if "montag" in estado_filter and "montag" not in estado_norm:
                    continue
                if "produ" in estado_filter and "produ" not in estado_norm:
                    continue
                if "concl" in estado_filter and "concl" not in estado_norm:
                    continue
            row = {
                "numero": str(enc.get("numero", "")).strip(),
                "of": self._order_of_code(enc),
                "nota_cliente": str(enc.get("nota_cliente", "") or "").strip(),
                "cliente": cli_display or cli_code or "-",
                "cliente_codigo": cli_code,
                "posto_trabalho": self._order_workcenter(enc),
                "data_criacao": str(enc.get("data_criacao", "") or "").strip(),
                "data_entrega": str(enc.get("data_entrega", "")).strip(),
                "tempo": self._fmt(enc.get("tempo_estimado", 0)),
                "estado": estado_txt,
                "cativar": "SIM" if bool(enc.get("cativar")) else "NAO",
                "pecas": len(pieces),
                "materiais": len(enc.get("materiais", []) or []),
                "planeado": self._fmt(planeado),
                "produzido": self._fmt(produzido),
                "progress": progress,
                "ano": enc_year,
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
        rows.sort(key=lambda item: item["numero"])
        return rows

    def order_presets(self) -> dict[str, Any]:
        data = self.ensure_data()
        materiais = list(
            dict.fromkeys(
                list(self.desktop_main.MATERIAIS_PRESET)
                + list(data.get("materiais_hist", []) or [])
                + [str(row.get("material", "") or "").strip() for row in list(data.get("materiais", []) or []) if str(row.get("material", "") or "").strip()]
            )
        )
        espessuras = list(
            dict.fromkeys(
                [self._fmt(v) for v in list(self.desktop_main.ESPESSURAS_PRESET)]
                + [str(value).strip() for value in list(data.get("espessuras_hist", []) or []) if str(value).strip()]
                + [str(row.get("espessura", "") or "").strip() for row in list(data.get("materiais", []) or []) if str(row.get("espessura", "") or "").strip()]
            )
        )
        return {
            "materiais": materiais,
            "espessuras": espessuras,
            "operacoes": list(dict.fromkeys(list(self.desktop_main.OFF_OPERACOES_DISPONIVEIS) + list(self.planning_operation_options()))),
            "operacao_default": str(self.desktop_main.OFF_OPERACAO_OBRIGATORIA),
        }

    def _order_sequence_from_numero(self, numero: str) -> str:
        numero_txt = str(numero or "").strip()
        match = re.search(r"(\d+)$", numero_txt)
        if not match:
            return ""
        return match.group(1).zfill(4)[-4:]

    def _order_expected_of_code(self, enc: dict[str, Any]) -> str:
        seq = self._order_sequence_from_numero(str((enc or {}).get("numero", "") or ""))
        if not seq:
            return ""
        year_txt = str((enc or {}).get("data_criacao", "") or self.desktop_main.now_iso()).strip()[:4]
        if not (len(year_txt) == 4 and year_txt.isdigit()):
            year_txt = str(datetime.now().year)
        return f"OF-{year_txt}-{seq}"

    def _order_of_code(self, enc: dict[str, Any], *, create: bool = True) -> str:
        expected_code = self._order_expected_of_code(enc)
        code = str(enc.get("of_codigo", "") or "").strip()
        if not code:
            ordem = enc.get("ordem_fabrico", {})
            if isinstance(ordem, dict):
                code = str(ordem.get("id", "") or ordem.get("codigo", "") or "").strip()
        if not code:
            for piece in list(self.desktop_main.encomenda_pecas(enc) or []):
                code = str(piece.get("of", "") or "").strip()
                if code:
                    break
        if expected_code and (not code or re.fullmatch(r"OF-\d{4}-\d{4,}", code, flags=re.IGNORECASE)):
            code = expected_code
        if not code and create:
            code = expected_code or str(self.desktop_main.next_of_numero(self.ensure_data()) or "").strip()
        if code:
            enc["of_codigo"] = code
            enc["ordem_fabrico"] = {
                "id": code,
                "encomenda_id": str(enc.get("numero", "") or "").strip(),
                "estado": str(enc.get("estado", "") or "Preparacao").strip() or "Preparacao",
                "data": str(enc.get("data_criacao", "") or self.desktop_main.now_iso()).strip()[:10],
            }
        return code

    def _next_order_opp_codigo(self, enc: dict[str, Any]) -> str:
        of_code = self._order_of_code(enc, create=True)
        prefix = ""
        parts = of_code.split("-")
        if len(parts) >= 3 and parts[0].upper() == "OF":
            prefix = f"OPP-{parts[1]}-{parts[2]}"
        max_seq = 0
        for piece in list(self.desktop_main.encomenda_pecas(enc) or []):
            opp = str(piece.get("opp", "") or "").strip()
            if prefix and opp.startswith(prefix + "-"):
                suffix = opp.rsplit("-", 1)[-1]
                if suffix.isdigit():
                    max_seq = max(max_seq, int(suffix))
        if prefix:
            return f"{prefix}-{max_seq + 1:02d}"
        return str(self.desktop_main.next_opp_numero(self.ensure_data()) or "").strip()

    def _ensure_order_fabrication_order(self, enc: dict[str, Any], *, sync_existing: bool = False) -> bool:
        changed = False
        previous_of = str(enc.get("of_codigo", "") or "").strip()
        of_code = self._order_of_code(enc, create=True)
        if of_code and previous_of != of_code:
            changed = True
        if of_code and str(enc.get("of_codigo", "") or "").strip() != of_code:
            enc["of_codigo"] = of_code
            changed = True
        if of_code:
            ordem = {
                "id": of_code,
                "encomenda_id": str(enc.get("numero", "") or "").strip(),
                "estado": str(enc.get("estado", "") or "Preparacao").strip() or "Preparacao",
                "data": str(enc.get("data_criacao", "") or self.desktop_main.now_iso()).strip()[:10],
            }
            if dict(enc.get("ordem_fabrico", {}) or {}) != ordem:
                enc["ordem_fabrico"] = ordem
                changed = True
        previous_opp_prefix = ""
        if previous_of and previous_of != of_code and previous_of.startswith("OF-"):
            previous_opp_prefix = "OPP-" + previous_of.split("-", 1)[1]
        for piece in list(self.desktop_main.encomenda_pecas(enc) or []):
            piece_of = str(piece.get("of", "") or "").strip()
            should_sync_piece_of = sync_existing or not piece_of or (previous_of and piece_of == previous_of)
            if of_code and should_sync_piece_of and piece_of != of_code:
                piece["of"] = of_code
                changed = True
            piece_opp = str(piece.get("opp", "") or "").strip()
            should_regen_opp = not piece_opp or bool(previous_opp_prefix and piece_opp.startswith(previous_opp_prefix + "-"))
            if should_regen_opp:
                piece["opp"] = self._next_order_opp_codigo(enc)
                changed = True
        return changed

    def _sync_order_piece_metrics_from_quote(self, enc: dict[str, Any]) -> bool:
        """Backfill unit time and weight into old orders created before these fields were copied."""
        if not isinstance(enc, dict):
            return False
        data = self.ensure_data()
        enc_numero = str(enc.get("numero", "") or "").strip()
        orc_numero = str(enc.get("numero_orcamento", "") or "").strip()
        source_quotes: list[dict[str, Any]] = []
        for orc in list(data.get("orcamentos", []) or []):
            if not isinstance(orc, dict):
                continue
            if orc_numero and str(orc.get("numero", "") or "").strip() == orc_numero:
                source_quotes.append(orc)
                continue
            if enc_numero and str(orc.get("numero_encomenda", "") or "").strip() == enc_numero:
                source_quotes.append(orc)
        if not source_quotes:
            return False

        def norm(value: Any) -> str:
            text = str(value or "").strip()
            norm_fn = getattr(self.desktop_main, "norm_text", None)
            return str(norm_fn(text) if callable(norm_fn) else text.casefold()).strip()

        def keys_for(row: dict[str, Any]) -> list[tuple[str, ...]]:
            ref_int = norm(row.get("ref_interna", ""))
            ref_ext = norm(row.get("ref_externa", ""))
            material = norm(row.get("material", "") or row.get("material_subtype", ""))
            espessura = norm(row.get("espessura", ""))
            descricao = norm(row.get("descricao", "") or row.get("Observacoes", ""))
            keys: list[tuple[str, ...]] = []
            if ref_int:
                keys.append(("int", ref_int))
            if ref_ext:
                keys.append(("ext", ref_ext))
            if ref_ext and material and espessura:
                keys.append(("extmat", ref_ext, material, espessura))
            if descricao and material and espessura:
                keys.append(("desc", descricao, material, espessura))
            return keys

        analysis_cache: dict[str, float] = {}

        def line_weight(line: dict[str, Any]) -> float:
            for key in ("peso_unid", "stock_metric_value", "peso_unitario", "peso"):
                value = self._parse_float(line.get(key, 0), 0)
                if value > 0:
                    return round(value, 4)
            path_txt = str(line.get("desenho", "") or line.get("desenho_path", "") or "").strip()
            if not path_txt:
                return 0.0
            resolved = self._resolve_file_reference(path_txt) or Path(path_txt)
            try:
                cache_key = str(resolved.resolve())
            except Exception:
                cache_key = str(resolved)
            if cache_key in analysis_cache:
                return analysis_cache[cache_key]
            weight = 0.0
            try:
                if resolved.exists() and resolved.suffix.lower() == ".dxf":
                    geometry = analyze_dxf_geometry(resolved)
                    metrics = dict(geometry.get("metrics", {}) or {})
                    net_area_m2 = self._parse_float(metrics.get("net_area_m2", geometry.get("net_area_m2", 0)), 0)
                    thickness_m = self._parse_float(line.get("espessura", 0), 0) / 1000.0
                    if net_area_m2 > 0 and thickness_m > 0:
                        weight = round(net_area_m2 * thickness_m * 7850.0, 4)
            except Exception:
                weight = 0.0
            analysis_cache[cache_key] = weight
            return weight

        quote_metrics: dict[tuple[str, ...], dict[str, float]] = {}
        for orc in source_quotes:
            for line in list(orc.get("linhas", []) or []):
                if not isinstance(line, dict):
                    continue
                tempo = round(self._parse_float(line.get("tempo_peca_min", line.get("tempo_pecas_min", 0)), 0), 4)
                peso = line_weight(line)
                if tempo <= 0 and peso <= 0:
                    continue
                payload = {"tempo_peca_min": tempo, "peso_unid": peso}
                for key in keys_for(line):
                    quote_metrics.setdefault(key, payload)
        if not quote_metrics:
            return False

        changed = False
        for piece in self.desktop_main.encomenda_pecas(enc):
            if not isinstance(piece, dict):
                continue
            metric: dict[str, float] | None = None
            for key in keys_for(piece):
                metric = quote_metrics.get(key)
                if metric:
                    break
            if not metric:
                continue
            if self._parse_float(piece.get("tempo_peca_min", piece.get("tempo_pecas_min", 0)), 0) <= 0 and metric.get("tempo_peca_min", 0) > 0:
                piece["tempo_peca_min"] = metric["tempo_peca_min"]
                changed = True
            if self._parse_float(piece.get("peso_unid", 0), 0) <= 0 and metric.get("peso_unid", 0) > 0:
                piece["peso_unid"] = metric["peso_unid"]
                changed = True
        return changed

    def order_suggest_ref_interna(
        self,
        numero: str = "",
        cliente: str = "",
        existing_refs: list[str] | tuple[str, ...] | set[str] | None = None,
    ) -> str:
        enc = self.get_encomenda_by_numero(numero) if str(numero or "").strip() else None
        cliente_codigo = str(cliente or (enc or {}).get("cliente", "") or "").strip()
        if not cliente_codigo:
            return ""
        existing = list(existing_refs or [])
        if enc is not None:
            existing.extend(str(piece.get("ref_interna", "") or "").strip() for piece in list(self.desktop_main.encomenda_pecas(enc)))
        return self._suggest_ref_interna_for_client(cliente_codigo, existing)

    def orc_suggest_ref_interna(
        self,
        cliente: str = "",
        existing_refs: list[str] | tuple[str, ...] | set[str] | None = None,
        numero: str = "",
    ) -> str:
        return self._suggest_ref_interna_for_client(cliente, existing_refs, exclude_orc_numero=numero)

    def order_create_or_update(self, payload: dict[str, Any]) -> dict[str, Any]:
        data = self.ensure_data()
        numero = str(payload.get("numero", "") or "").strip()
        enc = self.get_encomenda_by_numero(numero) if numero else None
        is_new = enc is None
        cliente = str(payload.get("cliente", "") or "").strip()
        if not cliente:
            raise ValueError("Cliente obrigatorio.")
        try:
            tempo_estimado = self._parse_float(payload.get("tempo_estimado", 0), 0)
        except Exception as exc:
            raise ValueError("Tempo estimado inválido.") from exc
        if enc is None:
            numero = self.desktop_main.next_encomenda_numero(data)
            enc = {
                "id": f"ENC{len(list(data.get('encomendas', []) or [])) + 1:05d}",
                "numero": numero,
                "cliente": cliente,
                "posto_trabalho": "",
                "nota_cliente": "",
                "nota_transporte": "",
                "preco_transporte": 0.0,
                "custo_transporte": 0.0,
                "paletes": 0.0,
                "peso_bruto_kg": 0.0,
                "volume_m3": 0.0,
                "transportadora_id": "",
                "transportadora_nome": "",
                "referencia_transporte": "",
                "zona_transporte": "",
                "local_descarga": "",
                "transporte_numero": "",
                "estado_transporte": "",
                "data_criacao": self.desktop_main.now_iso(),
                "data_entrega": "",
                "tempo_estimado": 0.0,
                "tempo": 0.0,
                "cativar": False,
                "Observacoes": "",
                "Observações": "",
                "estado": "Preparacao",
                "materiais": [],
                "reservas": [],
                "montagem_itens": [],
                "espessuras": [],
                "numero_orcamento": "",
                "tipo_encomenda": "Cliente",
                "of_codigo": "",
                "ordem_fabrico": {},
            }
            of_code = self._order_of_code(enc, create=True)
            data.setdefault("encomendas", []).append(enc)

        enc["cliente"] = cliente
        tipo_txt = str(payload.get("tipo_encomenda", enc.get("tipo_encomenda", "Cliente")) or "Cliente").strip()
        enc["tipo_encomenda"] = "Interna (produção)" if "intern" in self.desktop_main.norm_text(tipo_txt) else "Cliente"
        enc["posto_trabalho"] = self._normalize_workcenter_value(payload.get("posto_trabalho", "") or enc.get("posto_trabalho", ""))
        enc["nota_cliente"] = str(payload.get("nota_cliente", "") or "").strip()
        enc["nota_transporte"] = str(payload.get("nota_transporte", "") or enc.get("nota_transporte", "") or "").strip()
        enc["preco_transporte"] = round(self._parse_float(payload.get("preco_transporte", enc.get("preco_transporte", 0)), 0), 2)
        enc["custo_transporte"] = round(self._parse_float(payload.get("custo_transporte", enc.get("custo_transporte", 0)), 0), 2)
        enc["paletes"] = round(self._parse_float(payload.get("paletes", enc.get("paletes", 0)), 0), 2)
        enc["peso_bruto_kg"] = round(self._parse_float(payload.get("peso_bruto_kg", enc.get("peso_bruto_kg", 0)), 0), 2)
        enc["volume_m3"] = round(self._parse_float(payload.get("volume_m3", enc.get("volume_m3", 0)), 0), 3)
        transportadora_id, transportadora_nome, _transportadora_contacto = self._normalize_supplier_reference(
            payload.get("transportadora_id", enc.get("transportadora_id", "")),
            payload.get("transportadora_nome", enc.get("transportadora_nome", "")),
        )
        enc["transportadora_id"] = transportadora_id
        enc["transportadora_nome"] = transportadora_nome
        enc["referencia_transporte"] = str(payload.get("referencia_transporte", enc.get("referencia_transporte", "")) or "").strip()
        enc["zona_transporte"] = str(payload.get("zona_transporte", enc.get("zona_transporte", "")) or "").strip()
        enc["local_descarga"] = str(payload.get("local_descarga", "") or enc.get("local_descarga", "") or "").strip()
        enc["data_entrega"] = str(payload.get("data_entrega", "") or "").strip()
        enc["tempo_estimado"] = tempo_estimado
        enc["tempo"] = tempo_estimado
        obs_txt = str(payload.get("observacoes", "") or "").strip()
        enc["Observacoes"] = obs_txt
        enc["Observações"] = obs_txt
        requested_cativar = bool(payload.get("cativar"))
        if not requested_cativar and list(enc.get("reservas", []) or []):
            self.desktop_main.aplicar_reserva_em_stock(data, list(enc.get("reservas", []) or []), -1)
            enc["reservas"] = []
        enc.setdefault("montagem_itens", [])
        enc["cativar"] = requested_cativar and bool(enc.get("reservas"))
        self._ensure_order_fabrication_order(enc)
        self.desktop_main.update_estado_encomenda_por_espessuras(enc)
        self._save(force=True)
        return self.order_detail(numero)

    def order_remove(self, numero: str) -> None:
        numero = str(numero or "").strip()
        data = self.ensure_data()
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda não encontrada.")
        if list(enc.get("reservas", []) or []):
            self.desktop_main.aplicar_reserva_em_stock(data, list(enc.get("reservas", []) or []), -1)
        data["encomendas"] = [row for row in list(data.get("encomendas", []) or []) if str(row.get("numero", "") or "").strip() != numero]
        delete_order_fn = getattr(self.operador_actions, "_mysql_ops_delete_order", None)
        if callable(delete_order_fn):
            try:
                delete_order_fn(numero)
            except Exception:
                pass
        try:
            cache = getattr(self, "_op_mysql_ops_status_cache", None)
            if isinstance(cache, dict):
                cache.pop(numero, None)
        except Exception:
            pass
        for trip in list(data.get("transportes", []) or []):
            if not isinstance(trip, dict):
                continue
            trip["paragens"] = [
                row
                for row in list(trip.get("paragens", []) or [])
                if str((row or {}).get("encomenda_numero", (row or {}).get("encomenda", "")) or "").strip() != numero
            ]
            self._transport_reindex_stops(trip)
        self._transport_sync_order_links()
        self._save(force=True)

    def order_material_add(self, numero: str, material: str) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda não encontrada.")
        if self._order_is_orc_based(enc):
            raise ValueError("Encomenda originada de orcamento: material bloqueado.")
        material_txt = str(material or "").strip()
        if not material_txt:
            raise ValueError("Material obrigatorio.")
        if self._order_find_material(enc, material_txt) is not None:
            raise ValueError("Material ja existe.")
        enc.setdefault("materiais", []).append({"material": material_txt, "estado": "Preparacao", "espessuras": []})
        self.desktop_main.update_estado_encomenda_por_espessuras(enc)
        self._save(force=True)
        return self.order_detail(numero)

    def order_material_remove(self, numero: str, material: str) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda não encontrada.")
        if self._order_is_orc_based(enc):
            raise ValueError("Encomenda originada de orcamento: material bloqueado.")
        material_txt = str(material or "").strip().lower()
        enc["materiais"] = [row for row in list(enc.get("materiais", []) or []) if str(row.get("material", "") or "").strip().lower() != material_txt]
        self.desktop_main.update_estado_encomenda_por_espessuras(enc)
        self._save(force=True)
        return self.order_detail(numero)

    def order_espessura_add(self, numero: str, material: str, espessura: str) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda não encontrada.")
        if self._order_is_orc_based(enc):
            raise ValueError("Encomenda originada de orcamento: espessuras bloqueadas.")
        mat = self._order_find_material(enc, material)
        if mat is None:
            raise ValueError("Material não encontrado.")
        esp_txt = str(espessura or "").strip()
        if not esp_txt:
            raise ValueError("Espessura obrigatoria.")
        if self._order_find_espessura(enc, material, esp_txt) is not None:
            raise ValueError("Espessura ja existe.")
        mat.setdefault("espessuras", []).append({"espessura": esp_txt, "tempo_min": "", "tempos_operacao": {}, "maquinas_operacao": {}, "estado": "Preparacao", "pecas": []})
        self.desktop_main.update_estado_encomenda_por_espessuras(enc)
        self._save(force=True)
        return self.order_detail(numero)

    def order_espessura_remove(self, numero: str, material: str, espessura: str) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda não encontrada.")
        if self._order_is_orc_based(enc):
            raise ValueError("Encomenda originada de orcamento: espessuras bloqueadas.")
        mat = self._order_find_material(enc, material)
        if mat is None:
            raise ValueError("Material não encontrado.")
        esp_txt = str(espessura or "").strip()
        mat["espessuras"] = [row for row in list(mat.get("espessuras", []) or []) if str(row.get("espessura", "") or "").strip() != esp_txt]
        self.desktop_main.update_estado_encomenda_por_espessuras(enc)
        self._save(force=True)
        return self.order_detail(numero)

    def order_espessura_set_time(self, numero: str, material: str, espessura: str, tempo_min: Any) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda não encontrada.")
        esp = self._order_find_espessura(enc, material, espessura)
        if esp is None:
            raise ValueError("Espessura nao encontrada.")
        raw = str(tempo_min if tempo_min is not None else "").strip()
        if raw:
            try:
                int(raw)
            except Exception as exc:
                raise ValueError("Tempo inválido (minutos inteiros).") from exc
        esp["tempo_min"] = raw
        esp.setdefault("tempos_operacao", {})
        if raw:
            esp["tempos_operacao"]["Corte Laser"] = raw
        else:
            esp["tempos_operacao"].pop("Corte Laser", None)
        self.desktop_main.update_estado_encomenda_por_espessuras(enc)
        self._save(force=True)
        return self.order_detail(numero)

    def order_espessura_set_operation_times(
        self,
        numero: str,
        material: str,
        espessura: str,
        tempos_operacao: dict[str, Any] | None = None,
        maquinas_operacao: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda não encontrada.")
        esp = self._order_find_espessura(enc, material, espessura)
        if esp is None:
            raise ValueError("Espessura não encontrada.")
        cleaned: dict[str, str] = {}
        cleaned_resources: dict[str, str] = {}
        for op_name, raw_value in dict(tempos_operacao or {}).items():
            op_txt = self._planning_normalize_operation(op_name)
            if op_txt not in self.planning_operation_options():
                continue
            value_txt = str(raw_value if raw_value is not None else "").strip()
            if value_txt:
                try:
                    int(float(value_txt))
                except Exception as exc:
                    raise ValueError(f"Tempo inválido em {op_txt} (minutos inteiros).") from exc
                cleaned[op_txt] = value_txt
        for op_name, raw_value in dict(maquinas_operacao or {}).items():
            op_txt = self._planning_normalize_operation(op_name)
            if op_txt not in cleaned:
                continue
            resource_txt = self._sanitize_operation_resource(op_txt, raw_value)
            available_resources = [str(value or "").strip() for value in list(self.workcenter_resource_options(op_txt) or []) if str(value or "").strip()]
            if not resource_txt and len(available_resources) == 1:
                resource_txt = str(available_resources[0] or "").strip()
            if not resource_txt:
                continue
            if available_resources and all(resource_txt.lower() != value.lower() for value in available_resources):
                raise ValueError(f"O recurso '{resource_txt}' não pertence à operação {op_txt}.")
            cleaned_resources[op_txt] = resource_txt
        missing_resource = [op_name for op_name in cleaned if not str(cleaned_resources.get(op_name, "") or "").strip()]
        if missing_resource:
            raise ValueError(f"Seleciona o recurso/máquina para: {', '.join(missing_resource)}.")
        esp["tempo_min"] = cleaned.get("Corte Laser", str(esp.get("tempo_min", "") or "").strip())
        if not cleaned.get("Corte Laser"):
            esp["tempo_min"] = ""
        esp["tempos_operacao"] = cleaned
        esp["maquinas_operacao"] = cleaned_resources
        self.desktop_main.update_estado_encomenda_por_espessuras(enc)
        self._save(force=True)
        return self.order_detail(numero)

    def _next_order_piece_id(self, enc: dict[str, Any]) -> str:
        highest = 0
        for row in list(self.desktop_main.encomenda_pecas(enc)):
            try:
                highest = max(highest, int(str(row.get("id", "") or "").replace("PEC", "")))
            except Exception:
                continue
        return f"PEC{highest + 1:05d}"

    def order_piece_create_or_update(self, numero: str, payload: dict[str, Any], current_ref_interna: str = "") -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda não encontrada.")
        if self._order_is_orc_based(enc):
            raise ValueError("Encomenda originada de orcamento: pecas bloqueadas.")

        current_ref = str(current_ref_interna or "").strip()
        ref_int = str(payload.get("ref_interna", "") or "").strip()
        ref_ext = str(payload.get("ref_externa", "") or "").strip()
        tipo_material = str(payload.get("tipo_material", "") or "").strip().upper()
        if tipo_material not in {"CHAPA", "PERFIL", "TUBO", "OUTROS"}:
            tipo_material = "CHAPA"
        material = str(payload.get("material", "") or "").strip()
        subtipo_material = str(payload.get("subtipo_material", "") or material).strip()
        espessura = str(payload.get("espessura", "") or "").strip()
        dimensao = str(payload.get("dimensao", payload.get("dimensoes", "")) or "").strip()
        perfil_tipo = str(payload.get("perfil_tipo", "") or "").strip()
        perfil_tamanho = str(payload.get("perfil_tamanho", "") or "").strip()
        comprimento_mm = self._parse_float(payload.get("comprimento_mm", 0), 0)
        tubo_forma = str(payload.get("tubo_forma", "") or "").strip()
        lado_a = self._parse_float(payload.get("lado_a", 0), 0)
        lado_b = self._parse_float(payload.get("lado_b", 0), 0)
        tubo_espessura = self._parse_float(payload.get("tubo_espessura", 0), 0)
        diametro = self._parse_float(payload.get("diametro", 0), 0)
        descricao = str(payload.get("descricao", "") or "").strip()
        desenho = str(payload.get("desenho", "") or "").strip()
        ficheiros = [
            str(item or "").strip()
            for item in list(payload.get("ficheiros", []) or [])
            if str(item or "").strip()
        ]
        operacoes = " + ".join(self.desktop_main.parse_operacoes_lista(payload.get("operacoes", "")))
        if not operacoes:
            operacoes = str(self.desktop_main.OFF_OPERACAO_OBRIGATORIA)
        quantidade = self._parse_float(payload.get("quantidade_pedida", 0), 0)
        preco_unit = self._parse_float(payload.get("preco_unit", 0), 0)
        tempos_operacao = dict(payload.get("tempos_operacao", {}) or {})
        custos_operacao = dict(payload.get("custos_operacao", {}) or {})
        operacoes_detalhe = [dict(item or {}) for item in list(payload.get("operacoes_detalhe", []) or []) if isinstance(item, dict)]
        guardar_ref = bool(payload.get("guardar_ref", True))

        if not material:
            raise ValueError("Material obrigatorio.")
        if tipo_material == "CHAPA" and not espessura:
            raise ValueError("Espessura obrigatoria para chapa.")
        if tipo_material == "TUBO" and (not dimensao or not espessura):
            raise ValueError("Dimensao e espessura obrigatorias para tubo.")
        if tipo_material == "PERFIL" and not dimensao:
            raise ValueError("Dimensao obrigatoria para perfil.")
        if not espessura and tipo_material in {"PERFIL", "OUTROS"}:
            espessura = "-"
        if quantidade <= 0:
            raise ValueError("Quantidade invalida.")

        existing_refs = {
            str(piece.get("ref_interna", "") or "").strip()
            for piece in list(self.desktop_main.encomenda_pecas(enc))
            if str(piece.get("ref_interna", "") or "").strip()
        }
        if current_ref:
            existing_refs.discard(current_ref)
        if not ref_int:
            ref_int = self.desktop_main.next_ref_interna_unique(self.ensure_data(), enc.get("cliente", ""), list(existing_refs))
        if ref_int and ref_int in existing_refs:
            suggested = self.desktop_main.next_ref_interna_unique(self.ensure_data(), enc.get("cliente", ""), list(existing_refs))
            raise ValueError(f"Referencia interna ja existe nesta encomenda. Nova sugerida: {suggested}")

        mat = self._order_find_material(enc, material)
        if mat is None:
            mat = {"material": material, "estado": "Preparacao", "espessuras": []}
            enc.setdefault("materiais", []).append(mat)
        esp = self._order_find_espessura(enc, material, espessura)
        if esp is None:
            esp = {"espessura": espessura, "tempo_min": "", "tempos_operacao": {}, "maquinas_operacao": {}, "estado": "Preparacao", "pecas": []}
            mat.setdefault("espessuras", []).append(esp)

        _, old_esp, piece = self._order_find_piece(enc, current_ref, "")
        if piece is None:
            of_code = self._order_of_code(enc, create=True)
            piece = {
                "id": self._next_order_piece_id(enc),
                "of": of_code,
                "opp": self._next_order_opp_codigo(enc),
                "estado": "Preparacao",
                "produzido_ok": 0.0,
                "produzido_nok": 0.0,
                "produzido_qualidade": 0.0,
                "inicio_producao": "",
                "fim_producao": "",
                "hist": [],
                "qtd_expedida": 0.0,
                "expedicoes": [],
            }
            esp.setdefault("pecas", []).append(piece)
        elif old_esp is not esp:
            old_esp["pecas"] = [row for row in list(old_esp.get("pecas", []) or []) if row is not piece]
            esp.setdefault("pecas", []).append(piece)

        piece["ref_interna"] = ref_int
        piece["ref_externa"] = ref_ext
        piece["material"] = material
        piece["tipo_material"] = tipo_material
        piece["subtipo_material"] = subtipo_material
        piece["espessura"] = espessura
        piece["dimensao"] = dimensao
        piece["dimensoes"] = dimensao
        piece["perfil_tipo"] = perfil_tipo
        piece["perfil_tamanho"] = perfil_tamanho
        piece["comprimento_mm"] = comprimento_mm
        piece["tubo_forma"] = tubo_forma
        piece["lado_a"] = lado_a
        piece["lado_b"] = lado_b
        piece["tubo_espessura"] = tubo_espessura
        piece["diametro"] = diametro
        piece["descricao"] = descricao
        piece["quantidade_pedida"] = quantidade
        piece["Operacoes"] = operacoes
        piece["Observacoes"] = descricao
        piece["conjunto_codigo"] = str(payload.get("conjunto_codigo", piece.get("conjunto_codigo", "")) or "").strip()
        piece["conjunto_nome"] = str(payload.get("conjunto_nome", piece.get("conjunto_nome", "")) or "").strip()
        piece["grupo_uuid"] = str(payload.get("grupo_uuid", piece.get("grupo_uuid", "")) or "").strip()
        piece["desenho"] = desenho
        piece["desenho_path"] = desenho
        piece["ficheiros"] = ficheiros
        piece["tempos_operacao"] = dict(tempos_operacao)
        piece["custos_operacao"] = dict(custos_operacao)
        piece["operacoes_detalhe"] = list(operacoes_detalhe)
        if "of" not in piece or not str(piece.get("of", "")).strip():
            piece["of"] = self._order_of_code(enc, create=True)
        if "opp" not in piece or not str(piece.get("opp", "")).strip():
            piece["opp"] = self._next_order_opp_codigo(enc)
        piece["operacoes_fluxo"] = self.desktop_main.build_operacoes_fluxo(operacoes, piece.get("operacoes_fluxo"))
        self.desktop_main.ensure_peca_operacoes(piece)
        self.desktop_main.atualizar_estado_peca(piece)

        self.desktop_main.push_unique(self.ensure_data().setdefault("materiais_hist", []), material)
        self.desktop_main.push_unique(self.ensure_data().setdefault("espessuras_hist", []), espessura)

        if ref_ext:
            self.ensure_data().setdefault("peca_hist", {})[ref_ext] = {
                "ref_interna": ref_int,
                "descricao": descricao,
                "material": material,
                "espessura": espessura,
                "tipo_material": tipo_material,
                "subtipo_material": subtipo_material,
                "dimensao": dimensao,
                "Operacoes": operacoes,
                "Observacoes": descricao,
                "desenho": desenho,
                "ficheiros": list(ficheiros),
            }
            if guardar_ref:
                self.ensure_data().setdefault("orc_refs", {})[ref_ext] = {
                    "ref_interna": ref_int,
                    "ref_externa": ref_ext,
                    "descricao": descricao,
                    "material": material,
                    "espessura": espessura,
                    "preco_unit": preco_unit,
                    "operacao": operacoes,
                    "tempo_peca_min": round(self._parse_float(payload.get("tempo_peca_min", 0), 0), 3),
                    "operacoes_detalhe": list(operacoes_detalhe),
                    "tempos_operacao": dict(tempos_operacao),
                    "custos_operacao": dict(custos_operacao),
                    "desenho": desenho,
                }
                try:
                    self.desktop_main.mysql_upsert_orc_referencia(
                        ref_externa=ref_ext,
                        ref_interna=ref_int,
                        descricao=descricao,
                        material=material,
                        espessura=espessura,
                        preco_unit=preco_unit,
                        operacao=operacoes,
                        desenho_path=desenho,
                        tempo_peca_min=self._parse_float(payload.get("tempo_peca_min", 0), 0),
                        operacoes_lista=list(self.desktop_main.parse_operacoes_lista(operacoes)),
                        operacoes_detalhe=list(operacoes_detalhe),
                        tempos_operacao=dict(tempos_operacao),
                        custos_operacao=dict(custos_operacao),
                    )
                except Exception:
                    pass

        self.desktop_main.update_refs(self.ensure_data(), ref_int, ref_ext)
        self._ensure_order_fabrication_order(enc)
        self.desktop_main.update_estado_encomenda_por_espessuras(enc)
        self._save(force=True)
        return self.order_detail(numero)

    def order_model_options(self, filter_text: str = "") -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        rows: list[dict[str, Any]] = []
        for source, getter in (("modelo", getattr(self, "assembly_model_rows", None)), ("conjunto", getattr(self, "conjunto_rows", None))):
            if not callable(getter):
                continue
            for row in list(getter(filter_text) or []):
                if not bool(row.get("ativo", True)):
                    continue
                item = dict(row)
                item["origem_tipo"] = source
                item["label"] = f"{item.get('codigo', '')} | {item.get('descricao', '')}".strip(" |")
                if query and not any(query in str(value).lower() for value in item.values()):
                    continue
                rows.append(item)
        rows.sort(key=lambda row: (str(row.get("origem_tipo", "")), str(row.get("codigo", ""))))
        return rows

    def order_import_model(self, numero: str, codigo: str, quantity: Any = 1, source: str = "modelo") -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda não encontrada.")
        if self._order_is_orc_based(enc):
            raise ValueError("Encomenda originada de orçamento: estrutura bloqueada.")
        code = str(codigo or "").strip()
        if not code:
            raise ValueError("Seleciona um modelo/conjunto.")
        source_norm = str(source or "").strip().lower()
        expand_fn = self.conjunto_expand if source_norm == "conjunto" else self.assembly_model_expand
        detail_fn = self.conjunto_detail if source_norm == "conjunto" else self.assembly_model_detail
        source_detail = dict(detail_fn(code) or {})
        rows = list(expand_fn(code, quantity) or [])
        if not rows:
            raise ValueError("O modelo não tem linhas para importar.")
        imported_pieces = 0
        imported_items = 0
        for line in rows:
            line_type = self.desktop_main.normalize_orc_line_type(line.get("tipo_item"))
            if self.desktop_main.orc_line_is_piece(line):
                self.order_piece_create_or_update(
                    numero,
                    {
                        "ref_interna": "",
                        "ref_externa": str(line.get("ref_externa", "") or "").strip(),
                        "descricao": str(line.get("descricao", "") or "").strip(),
                        "tipo_material": str(line.get("tipo_material", "") or line.get("material_family", "") or "CHAPA").strip().upper(),
                        "material": str(line.get("material", "") or line.get("material_subtype", "") or "").strip(),
                        "subtipo_material": str(line.get("material_subtype", "") or line.get("material", "") or "").strip(),
                        "espessura": str(line.get("espessura", "") or "").strip(),
                        "dimensao": str(line.get("dimensao", line.get("dimensoes", "")) or line.get("profile_size", "") or line.get("tube_section", "") or "").strip(),
                        "operacoes": str(line.get("operacao", "") or "Embalamento").strip(),
                        "quantidade_pedida": self._parse_float(line.get("qtd", 0), 0),
                        "preco_unit": self._parse_float(line.get("preco_unit", 0), 0),
                        "tempo_peca_min": self._parse_float(line.get("tempo_peca_min", 0), 0),
                        "desenho": str(line.get("desenho", "") or "").strip(),
                        "guardar_ref": True,
                        "conjunto_codigo": str(line.get("conjunto_codigo", code) or "").strip(),
                        "conjunto_nome": str(line.get("conjunto_nome", source_detail.get("descricao", "")) or "").strip(),
                        "grupo_uuid": str(line.get("grupo_uuid", "") or "").strip(),
                    },
                )
                imported_pieces += 1
                continue
            enc.setdefault("montagem_itens", []).append(
                {
                    "linha_ordem": len(list(enc.get("montagem_itens", []) or [])) + 1,
                    "tipo_item": line_type,
                    "stock_item_kind": str(line.get("stock_item_kind", "") or "").strip(),
                    "descricao": str(line.get("descricao", "") or "").strip(),
                    "dimensao": str(line.get("dimensao", line.get("dimensoes", "")) or "").strip(),
                    "material": str(line.get("material", "") or "").strip(),
                    "espessura": str(line.get("espessura", "") or "").strip(),
                    "stock_material_id": str(line.get("stock_material_id", "") or "").strip(),
                    "produto_codigo": str(line.get("produto_codigo", "") or "").strip(),
                    "produto_unid": str(line.get("produto_unid", "") or "").strip() or ("SV" if self.desktop_main.orc_line_is_service(line) else "UN"),
                    "qtd_planeada": round(self._parse_float(line.get("qtd", 0), 0), 2),
                    "qtd_consumida": 0.0,
                    "preco_unit": round(self._parse_float(line.get("preco_unit", 0), 0), 4),
                    "conjunto_codigo": str(line.get("conjunto_codigo", code) or "").strip(),
                    "conjunto_nome": str(line.get("conjunto_nome", "") or "").strip(),
                    "grupo_uuid": str(line.get("grupo_uuid", "") or "").strip(),
                    "estado": "Pendente",
                    "obs": str(line.get("operacao", "") or "").strip(),
                    "created_at": self.desktop_main.now_iso(),
                    "consumed_at": "",
                    "consumed_by": "",
                }
            )
            imported_items += 1
        order_sheets = enc.setdefault("produto_fichas", [])
        existing_sheet = next(
            (row for row in order_sheets if str(row.get("codigo", "") or "").strip() == code),
            None,
        )
        snapshot = {
            "codigo": code,
            "param_codigo": str(source_detail.get("param_codigo", "") or "").strip(),
            "descricao": str(source_detail.get("descricao", "") or code).strip(),
            "notas": str(source_detail.get("notas", "") or "").strip(),
            "ficha_tecnica": dict(source_detail.get("ficha_tecnica", {}) or {}),
            "quantidade_conjuntos": round(self._parse_float(quantity, 1), 2),
            "total_custo": round(self._parse_float(source_detail.get("total_custo", 0), 0), 2),
            "total_final": round(self._parse_float(source_detail.get("total_final", 0), 0), 2),
            "margem_perc": round(self._parse_float(source_detail.get("margem_perc", 0), 0), 2),
        }
        if existing_sheet is None:
            order_sheets.append(snapshot)
        else:
            snapshot["quantidade_conjuntos"] = round(
                self._parse_float(existing_sheet.get("quantidade_conjuntos", 0), 0)
                + self._parse_float(quantity, 1),
                2,
            )
            existing_sheet.update(snapshot)
        enc["valor_adjudicado"] = round(
            sum(
                self._parse_float(row.get("total_final", 0), 0)
                * self._parse_float(row.get("quantidade_conjuntos", 0), 0)
                for row in order_sheets
                if isinstance(row, dict)
            ),
            2,
        )
        self._ensure_order_fabrication_order(enc)
        self.desktop_main.update_estado_encomenda_por_espessuras(enc)
        self._save(force=True)
        detail = self.order_detail(numero)
        detail["imported_pieces"] = imported_pieces
        detail["imported_items"] = imported_items
        return detail

    def order_create_with_models(
        self,
        payload: dict[str, Any],
        imports: list[dict[str, Any]] | None = None,
    ) -> dict[str, Any]:
        """Create an order and seed its manufacturing structure in one workflow."""
        if str(payload.get("numero", "") or "").strip():
            raise ValueError("Este fluxo destina-se apenas a novas encomendas.")

        normalized_imports: list[dict[str, Any]] = []
        for raw in list(imports or []):
            source = str(raw.get("source", raw.get("origem_tipo", "conjunto")) or "conjunto").strip().lower()
            source = "conjunto" if source == "conjunto" else "modelo"
            code = str(raw.get("codigo", "") or "").strip()
            quantity = round(self._parse_float(raw.get("quantity", raw.get("quantidade", 1)), 0), 2)
            if not code or quantity <= 0:
                raise ValueError("Conjunto e quantidade são obrigatórios.")
            detail_fn = self.conjunto_detail if source == "conjunto" else self.assembly_model_detail
            expand_fn = self.conjunto_expand if source == "conjunto" else self.assembly_model_expand
            detail_fn(code)
            if not list(expand_fn(code, quantity) or []):
                raise ValueError(f"O conjunto {code} não tem linhas para importar.")
            normalized_imports.append({"codigo": code, "quantity": quantity, "source": source})

        created = self.order_create_or_update(dict(payload or {}))
        numero = str(created.get("numero", "") or "").strip()
        imported_pieces = 0
        imported_items = 0
        try:
            for entry in normalized_imports:
                result = self.order_import_model(
                    numero,
                    entry["codigo"],
                    entry["quantity"],
                    entry["source"],
                )
                imported_pieces += int(result.get("imported_pieces", 0) or 0)
                imported_items += int(result.get("imported_items", 0) or 0)
        except Exception:
            try:
                self.order_remove(numero)
            except Exception:
                pass
            raise

        detail = self.order_detail(numero)
        detail["imported_pieces"] = imported_pieces
        detail["imported_items"] = imported_items
        detail["initial_imports"] = [dict(row) for row in normalized_imports]
        return detail

    def order_piece_remove(self, numero: str, ref_interna: str) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda não encontrada.")
        if self._order_is_orc_based(enc):
            raise ValueError("Encomenda originada de orcamento: pecas bloqueadas.")
        ref_int = str(ref_interna or "").strip()
        found = False
        for mat in list(enc.get("materiais", []) or []):
            for esp in list(mat.get("espessuras", []) or []):
                before = len(list(esp.get("pecas", []) or []))
                esp["pecas"] = [row for row in list(esp.get("pecas", []) or []) if str(row.get("ref_interna", "") or "").strip() != ref_int]
                if len(list(esp.get("pecas", []) or [])) != before:
                    found = True
        if not found:
            raise ValueError("Peça não encontrada.")
        self.desktop_main.update_estado_encomenda_por_espessuras(enc)
        self._save(force=True)
        return self.order_detail(numero)

    def order_stock_candidates(self, numero: str, material: str, espessura: str) -> list[dict[str, Any]]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda não encontrada.")
        material_norm = self.encomendas_actions._norm_material(material)
        esp_norm = self.encomendas_actions._norm_espessura(espessura)
        rows = []
        for stock in list(self.ensure_data().get("materiais", []) or []):
            if self._material_quality_is_blocked(stock):
                continue
            disponivel = self._parse_float(stock.get("quantidade", 0), 0) - self._parse_float(stock.get("reservado", 0), 0)
            if disponivel <= 0:
                continue
            if self.encomendas_actions._norm_material(stock.get("material")) != material_norm:
                continue
            if self.encomendas_actions._norm_espessura(stock.get("espessura")) != esp_norm:
                continue
            rows.append(
                {
                    "material_id": str(stock.get("id", "") or "").strip(),
                    "dimensao": f"{stock.get('comprimento', '')}x{stock.get('largura', '')}",
                    "disponivel": round(disponivel, 2),
                    "local": self._localizacao(stock),
                    "lote": str(stock.get("lote_interno", "") or stock.get("lote_fornecedor", "") or "").strip(),
                    "lote_fornecedor": str(stock.get("lote_fornecedor", "") or "").strip(),
                }
            )
        return rows

    def order_reserve_stock(self, numero: str, material: str, espessura: str, allocations: list[dict[str, Any]]) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda não encontrada.")
        if not str(material or "").strip() or not str(espessura or "").strip():
            raise ValueError("Selecione material e espessura antes de cativar.")
        any_saved = False
        for row in allocations or []:
            material_id = str((row or {}).get("material_id", "") or "").strip()
            quantidade = self._parse_float((row or {}).get("quantidade", 0), 0)
            if not material_id or quantidade <= 0:
                continue
            stock = next((m for m in list(self.ensure_data().get("materiais", []) or []) if str(m.get("id", "") or "").strip() == material_id), None)
            if stock is None:
                raise ValueError(f"Material não encontrado: {material_id}")
            if self._material_quality_is_blocked(stock):
                raise ValueError(f"Material {material_id} bloqueado pela qualidade.")
            disponivel = self._parse_float(stock.get("quantidade", 0), 0) - self._parse_float(stock.get("reservado", 0), 0)
            if quantidade > disponivel:
                raise ValueError(f"Quantidade maior que o disponivel para {material_id}")
            stock["reservado"] = self._parse_float(stock.get("reservado", 0), 0) + quantidade
            stock["atualizado_em"] = self.desktop_main.now_iso()
            enc.setdefault("reservas", []).append(
                {
                    "material_id": material_id,
                    "material": stock.get("material"),
                    "espessura": stock.get("espessura"),
                    "quantidade": quantidade,
                }
            )
            self.desktop_main.log_stock(
                self.ensure_data(),
                "CATIVAR",
                f"{material_id} qtd={quantidade} encomenda={enc.get('numero', '')}",
                operador=self._current_user_label(),
            )
            any_saved = True
        if not any_saved:
            raise ValueError("Nenhuma quantidade definida.")
        enc["cativar"] = True
        self._save(force=True)
        return self.order_detail(numero)

    def order_release_stock(self, numero: str, material: str, espessura: str) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(numero)
        if enc is None:
            raise ValueError("Encomenda não encontrada.")
        target = []
        keep = []
        for row in list(enc.get("reservas", []) or []):
            if self.encomendas_actions._match_material(row.get("material"), material) and self.encomendas_actions._norm_espessura(row.get("espessura")) == self.encomendas_actions._norm_espessura(espessura):
                target.append(row)
            else:
                keep.append(row)
        if not target:
            raise ValueError(f"Sem reservas para {material} esp. {espessura}.")
        for row in target:
            self.desktop_main.log_stock(
                self.ensure_data(),
                "LIBERTAR",
                f"{row.get('material_id', '')} qtd={row.get('quantidade', 0)} encomenda={enc.get('numero', '')}",
                operador=self._current_user_label(),
            )
        self.desktop_main.aplicar_reserva_em_stock(self.ensure_data(), target, -1)
        enc["reservas"] = keep
        enc["cativar"] = bool(enc.get("reservas"))
        self._save(force=True)
        return self.order_detail(numero)
