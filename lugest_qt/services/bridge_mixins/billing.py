from __future__ import annotations

import json
import os
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any


class BillingBridgeMixin:
    """Billing and fiscal export operations for the Qt bridge."""

    def _billing_records(self) -> list[dict[str, Any]]:
        return list(self.ensure_data().get("faturacao", []) or [])

    def _billing_find_record(self, numero: str) -> dict[str, Any] | None:
        target = str(numero or "").strip()
        if not target:
            return None
        return next(
            (
                row
                for row in self._billing_records()
                if str((row or {}).get("numero", "") or "").strip() == target
            ),
            None,
        )

    def _billing_find_source_record(self, orcamento_numero: str = "", encomenda_numero: str = "") -> dict[str, Any] | None:
        orc_num = str(orcamento_numero or "").strip()
        enc_num = str(encomenda_numero or "").strip()
        for row in self._billing_records():
            if not isinstance(row, dict):
                continue
            row_orc = str(row.get("orcamento_numero", "") or "").strip()
            row_enc = str(row.get("encomenda_numero", "") or "").strip()
            if orc_num and row_orc == orc_num:
                return row
            if enc_num and row_enc == enc_num:
                return row
        return None

    def _billing_next_number(self) -> str:
        year = str(self.desktop_main.datetime.now().year)
        highest = 0
        for row in self._billing_records():
            raw = str((row or {}).get("numero", "") or "").strip().upper()
            if not raw:
                continue
            digits = "".join(ch for ch in raw if ch.isdigit())
            if len(digits) >= 8 and digits.startswith(year):
                try:
                    highest = max(highest, int(digits[-4:]))
                    continue
                except Exception:
                    pass
            if digits:
                try:
                    highest = max(highest, int(digits[-4:]))
                except Exception:
                    continue
        return f"FAT-{year}-{highest + 1:04d}"

    def _billing_quote_by_number(self, numero: str) -> dict[str, Any] | None:
        target = str(numero or "").strip()
        if not target:
            return None
        return next(
            (
                row
                for row in list(self.ensure_data().get("orcamentos", []) or [])
                if str((row or {}).get("numero", "") or "").strip() == target
            ),
            None,
        )

    def _billing_order_by_number(self, numero: str) -> dict[str, Any] | None:
        return self.get_encomenda_by_numero(str(numero or "").strip())

    def _billing_quote_is_sold(self, quote: dict[str, Any]) -> bool:
        estado = self.desktop_main.norm_text((quote or {}).get("estado", ""))
        return bool(
            "aprov" in estado
            or "convert" in estado
            or str((quote or {}).get("numero_encomenda", "") or "").strip()
        )

    def _billing_client_info(
        self,
        *,
        quote: dict[str, Any] | None = None,
        order: dict[str, Any] | None = None,
        record: dict[str, Any] | None = None,
    ) -> dict[str, str]:
        if isinstance(quote, dict):
            client = self._normalize_orc_client(quote.get("cliente", {}))
            code = str(client.get("codigo", "") or "").strip()
            name = str(client.get("nome", "") or client.get("empresa", "") or "").strip()
            if code or name:
                return {"codigo": code, "nome": name, "label": f"{code} - {name}".strip(" -")}
        if isinstance(order, dict):
            code = str(order.get("cliente", "") or "").strip()
            name = ""
            if code:
                try:
                    name = str((self.desktop_main.find_cliente(self.ensure_data(), code) or {}).get("nome", "") or "").strip()
                except Exception:
                    name = ""
            if code or name:
                return {"codigo": code, "nome": name, "label": f"{code} - {name}".strip(" -")}
        if isinstance(record, dict):
            code = str(record.get("cliente_codigo", "") or "").strip()
            name = str(record.get("cliente_nome", "") or "").strip()
            if code or name:
                return {"codigo": code, "nome": name, "label": f"{code} - {name}".strip(" -")}
        return {"codigo": "", "nome": "", "label": "-"}

    def _billing_guides_for_order(self, encomenda_numero: str) -> list[dict[str, Any]]:
        enc_num = str(encomenda_numero or "").strip()
        if not enc_num:
            return []
        rows: list[dict[str, Any]] = []
        for row in list(self.ensure_data().get("expedicoes", []) or []):
            if not isinstance(row, dict):
                continue
            if bool(row.get("anulada")):
                continue
            order_number = str(row.get("encomenda", "") or row.get("encomenda_numero", "") or "").strip()
            if order_number != enc_num:
                continue
            rows.append(
                {
                    "numero": str(row.get("numero", "") or "").strip(),
                    "data_emissao": str(row.get("data_emissao", "") or "").replace("T", " ")[:19],
                    "destinatario": str(row.get("destinatario", "") or "").strip(),
                }
            )
        rows.sort(key=lambda item: str(item.get("data_emissao", "") or ""), reverse=True)
        return rows

    def _billing_default_serie_id(self, issue_date: str = "") -> str:
        raw_issue_date = str(issue_date or self.desktop_main.now_iso()).strip() or self.desktop_main.now_iso()
        default_fn = getattr(self.desktop_main, "_exp_default_serie_id", None)
        if callable(default_fn):
            try:
                return str(default_fn("FT", raw_issue_date) or "").strip() or f"FT{str(raw_issue_date)[:4]}"
            except Exception:
                pass
        return f"FT{str(raw_issue_date)[:4]}"

    def _billing_due_days_from_text(self, value: str) -> int:
        digits = "".join(ch if ch.isdigit() else " " for ch in str(value or ""))
        values = [chunk for chunk in digits.split() if chunk.isdigit()]
        if not values:
            return 30
        try:
            return max(0, min(365, int(values[0])))
        except Exception:
            return 30

    def _billing_client_snapshot(
        self,
        *,
        quote: dict[str, Any] | None = None,
        order: dict[str, Any] | None = None,
        record: dict[str, Any] | None = None,
    ) -> dict[str, str]:
        base = {
            "codigo": "",
            "nome": "",
            "nif": "",
            "morada": "",
            "contacto": "",
            "email": "",
            "cond_pagamento": "",
        }
        if isinstance(quote, dict):
            qclient = dict(self._normalize_orc_client(quote.get("cliente", {})) or {})
            base.update(
                {
                    "codigo": str(qclient.get("codigo", "") or "").strip(),
                    "nome": str(qclient.get("nome", "") or qclient.get("empresa", "") or "").strip(),
                    "nif": str(qclient.get("nif", "") or "").strip(),
                    "morada": str(qclient.get("morada", "") or "").strip(),
                    "contacto": str(qclient.get("contacto", "") or "").strip(),
                    "email": str(qclient.get("email", "") or "").strip(),
                }
            )
        if isinstance(order, dict) and not base.get("codigo"):
            base["codigo"] = str(order.get("cliente", "") or "").strip()
        if isinstance(record, dict):
            if not base.get("codigo"):
                base["codigo"] = str(record.get("cliente_codigo", "") or "").strip()
            if not base.get("nome"):
                base["nome"] = str(record.get("cliente_nome", "") or "").strip()

        client_ref = None
        client_code = str(base.get("codigo", "") or "").strip()
        for row in list(self.ensure_data().get("clientes", []) or []):
            if not isinstance(row, dict):
                continue
            row_code = str(row.get("codigo", "") or "").strip()
            if client_code and row_code == client_code:
                client_ref = row
                break
        if client_ref is None:
            for row in list(self.ensure_data().get("clientes", []) or []):
                if not isinstance(row, dict):
                    continue
                if base.get("nif") and str(row.get("nif", "") or "").strip() == base["nif"]:
                    client_ref = row
                    break
                if base.get("nome") and str(row.get("nome", "") or "").strip() == base["nome"]:
                    client_ref = row
                    break
        if isinstance(client_ref, dict):
            base["codigo"] = str(client_ref.get("codigo", "") or base.get("codigo", "") or "").strip()
            base["nome"] = str(base.get("nome", "") or client_ref.get("nome", "") or "").strip()
            base["nif"] = str(base.get("nif", "") or client_ref.get("nif", "") or "").strip()
            base["morada"] = str(base.get("morada", "") or client_ref.get("morada", "") or "").strip()
            base["contacto"] = str(base.get("contacto", "") or client_ref.get("contacto", "") or "").strip()
            base["email"] = str(base.get("email", "") or client_ref.get("email", "") or "").strip()
            base["cond_pagamento"] = str(client_ref.get("cond_pagamento", "") or "").strip()
        return base

    def _billing_actor(self) -> str:
        return str((self.user or {}).get("username", "") or "Sistema").strip() or "Sistema"

    def _billing_software_cert_number(self) -> str:
        branding_cfg = {}
        try:
            branding_cfg = dict(self.desktop_main.get_branding_config() or {})
        except Exception:
            branding_cfg = {}
        return str(
            os.getenv("LUGEST_SOFTWARE_CERT_NUMBER", "")
            or os.getenv("LUGEST_SOFTWARE_CERT", "")
            or branding_cfg.get("software_cert", "")
            or branding_cfg.get("software_cert_number", "")
            or ""
        ).strip()

    def _billing_software_producer_info(self, issuer: dict[str, Any] | None = None) -> dict[str, str]:
        issuer_row = dict(issuer or getattr(self.desktop_main, "get_guia_emitente_info", lambda: {})() or {})
        branding_cfg = {}
        try:
            branding_cfg = dict(self.desktop_main.get_branding_config() or {})
        except Exception:
            branding_cfg = {}
        producer_name = str(
            os.getenv("LUGEST_SOFTWARE_PRODUCER_NAME", "")
            or branding_cfg.get("software_producer_name", "")
            or issuer_row.get("nome", "")
            or "LuGEST"
        ).strip() or "LuGEST"
        producer_nif = str(
            os.getenv("LUGEST_SOFTWARE_PRODUCER_NIF", "")
            or branding_cfg.get("software_producer_nif", "")
            or issuer_row.get("nif", "")
            or "999999990"
        ).strip() or "999999990"
        product_id = str(
            os.getenv("LUGEST_SOFTWARE_PRODUCT_ID", "")
            or branding_cfg.get("software_product_id", "")
            or self.tax_compliance.DEFAULT_PRODUCT_ID
        ).strip() or self.tax_compliance.DEFAULT_PRODUCT_ID
        product_version = str(
            os.getenv("LUGEST_SOFTWARE_PRODUCT_VERSION", "")
            or branding_cfg.get("software_product_version", "")
            or self.tax_compliance.DEFAULT_PRODUCT_VERSION
        ).strip() or self.tax_compliance.DEFAULT_PRODUCT_VERSION
        hash_control = str(
            os.getenv("LUGEST_HASH_CONTROL", "")
            or branding_cfg.get("hash_control", "")
            or self.tax_compliance.DEFAULT_HASH_CONTROL
        ).strip() or self.tax_compliance.DEFAULT_HASH_CONTROL
        return {
            "producer_name": producer_name,
            "producer_nif": producer_nif,
            "product_id": product_id,
            "product_version": product_version,
            "hash_control": hash_control,
        }

    def _billing_signing_material(self) -> dict[str, str]:
        return self.tax_compliance.load_or_create_signing_material(self.base_dir)

    def _billing_invoice_snapshot(self, invoice: dict[str, Any] | None) -> dict[str, Any]:
        if not isinstance(invoice, dict):
            return {}
        return self.tax_compliance.deserialize_snapshot(invoice.get("document_snapshot_json", ""))

    def _billing_store_invoice_snapshot(self, invoice: dict[str, Any], document: dict[str, Any]) -> None:
        invoice["document_snapshot_json"] = self.tax_compliance.serialize_snapshot(document)

    def _billing_legal_invoice_no(self, invoice: dict[str, Any]) -> str:
        return self.tax_compliance.legal_document_number(
            invoice.get("doc_type", "FT"),
            invoice.get("serie_id") or invoice.get("serie"),
            invoice.get("seq_num", 0),
            invoice.get("numero_fatura", ""),
        )

    def _billing_status_source_id(self, invoice: dict[str, Any]) -> str:
        return str(
            invoice.get("status_source_id", "")
            or invoice.get("source_id", "")
            or self._billing_actor()
        ).strip() or "Sistema"

    def _billing_invoice_core_fields(self, invoice: dict[str, Any]) -> tuple[str, ...]:
        return (
            str(invoice.get("doc_type", "") or "").strip(),
            str(invoice.get("numero_fatura", "") or "").strip(),
            str(invoice.get("serie_id", "") or invoice.get("serie", "") or "").strip(),
            str(int(self._parse_float(invoice.get("seq_num", 0), 0) or 0)),
            str(invoice.get("atcud", "") or "").strip(),
            str(invoice.get("guia_numero", "") or "").strip(),
            str(invoice.get("data_emissao", "") or "").strip()[:10],
            f"{round(self._parse_float(invoice.get('valor_total', 0), 0), 2):.2f}",
        )

    def _billing_invoice_locked(self, invoice: dict[str, Any] | None) -> bool:
        if not isinstance(invoice, dict):
            return False
        return any(
            str(invoice.get(key, "") or "").strip()
            for key in ("system_entry_date", "hash", "legal_invoice_no", "document_snapshot_json")
        )

    def _billing_previous_signed_hash(self, current_invoice: dict[str, Any]) -> str:
        current_id = str(current_invoice.get("id", "") or "").strip()
        current_doc_type = str(current_invoice.get("doc_type", "") or "FT").strip().upper() or "FT"
        current_series = str(current_invoice.get("serie_id", "") or current_invoice.get("serie", "") or "").strip()
        current_seq = int(self._parse_float(current_invoice.get("seq_num", 0), 0) or 0)
        current_date = str(current_invoice.get("data_emissao", "") or "").strip()[:10]
        previous: dict[str, Any] | None = None
        for record in self._billing_records():
            if not isinstance(record, dict):
                continue
            for row in list(record.get("faturas", []) or []):
                if not isinstance(row, dict):
                    continue
                row_id = str(row.get("id", "") or "").strip()
                if current_id and row_id == current_id:
                    continue
                row_doc_type = str(row.get("doc_type", "") or "FT").strip().upper() or "FT"
                row_series = str(row.get("serie_id", "") or row.get("serie", "") or "").strip()
                if row_doc_type != current_doc_type or row_series != current_series:
                    continue
                row_hash = str(row.get("hash", "") or "").strip()
                if not row_hash:
                    continue
                row_seq = int(self._parse_float(row.get("seq_num", 0), 0) or 0)
                if current_seq > 0 and row_seq > 0:
                    if row_seq >= current_seq:
                        continue
                    if previous is None or row_seq > int(self._parse_float(previous.get("seq_num", 0), 0) or 0):
                        previous = row
                    continue
                row_entry = str(row.get("system_entry_date", "") or row.get("created_at", "") or "").strip()
                current_entry = str(current_invoice.get("system_entry_date", "") or current_invoice.get("created_at", "") or "").strip()
                if row_entry and current_entry and row_entry >= current_entry:
                    continue
                if previous is None:
                    previous = row
                    continue
                prev_entry = str(previous.get("system_entry_date", "") or previous.get("created_at", "") or "").strip()
                if row_entry > prev_entry or (row_entry == prev_entry and str(row.get("data_emissao", "") or "") >= str(previous.get("data_emissao", "") or "")):
                    previous = row
        return str((previous or {}).get("hash", "") or "").strip()

    def _billing_saft_hash_value(self, invoice: dict[str, Any]) -> str:
        if not self._billing_software_cert_number():
            return "0"
        if str(invoice.get("source_billing", "") or "").strip().upper() != "P":
            return "0"
        return str(invoice.get("hash", "") or "").strip() or "0"

    def _billing_saft_hash_control(self, invoice: dict[str, Any]) -> str:
        if not self._billing_software_cert_number():
            return "0"
        if str(invoice.get("source_billing", "") or "").strip().upper() != "P":
            return "0"
        return str(invoice.get("hash_control", "") or "").strip() or "0"

    def _billing_ensure_invoice_compliance(
        self,
        record: dict[str, Any],
        invoice: dict[str, Any],
        *,
        actor: str = "",
        force_snapshot: bool = False,
    ) -> dict[str, Any]:
        actor_txt = str(actor or self._billing_actor()).strip() or "Sistema"
        invoice["system_entry_date"] = str(invoice.get("system_entry_date", "") or invoice.get("created_at", "") or self.desktop_main.now_iso()).strip()
        invoice["source_id"] = str(invoice.get("source_id", "") or actor_txt).strip() or "Sistema"
        invoice["status_source_id"] = str(invoice.get("status_source_id", "") or invoice.get("source_id", "") or actor_txt).strip() or "Sistema"
        fallback_source = "M" if (str(invoice.get("caminho", "") or "").strip() and int(self._parse_float(invoice.get("seq_num", 0), 0) or 0) <= 0) else "P"
        invoice["source_billing"] = self.tax_compliance.normalize_source_billing(invoice.get("source_billing", ""), fallback=fallback_source)
        invoice["legal_invoice_no"] = str(invoice.get("legal_invoice_no", "") or self._billing_legal_invoice_no(invoice)).strip()
        producer = self._billing_software_producer_info()
        invoice["hash_control"] = str(invoice.get("hash_control", "") or producer.get("hash_control", self.tax_compliance.DEFAULT_HASH_CONTROL)).strip()
        if invoice["source_billing"] == "P":
            invoice["previous_hash"] = str(invoice.get("previous_hash", "") or self._billing_previous_signed_hash(invoice)).strip()
            if not str(invoice.get("hash", "") or "").strip():
                signing = self._billing_signing_material()
                seed = self.tax_compliance.build_invoice_hash_message(
                    invoice_date=invoice.get("data_emissao", ""),
                    system_entry_date=invoice.get("system_entry_date", ""),
                    invoice_no=invoice.get("legal_invoice_no", "") or invoice.get("numero_fatura", ""),
                    gross_total=invoice.get("valor_total", 0),
                    previous_hash=invoice.get("previous_hash", ""),
                )
                invoice["hash"] = self.tax_compliance.sign_message_pkcs1_sha1(seed, signing["private_key_pem"])
        else:
            invoice["previous_hash"] = ""
        invoice["communication_status"] = str(invoice.get("communication_status", "") or "Por comunicar").strip() or "Por comunicar"
        needs_snapshot = force_snapshot or not self._billing_invoice_snapshot(invoice)
        if needs_snapshot:
            document = self._billing_build_invoice_document(str(record.get("numero", "") or "").strip(), invoice, prefer_snapshot=False)
            self._billing_store_invoice_snapshot(invoice, document)
            return document
        return self._billing_build_invoice_document(str(record.get("numero", "") or "").strip(), invoice)

    def _billing_next_invoice_identifiers(
        self,
        *,
        issue_date: str = "",
        serie_id: str = "",
        validation_code_hint: str = "",
        reserve: bool = False,
    ) -> dict[str, Any]:
        issue_txt = str(issue_date or self.desktop_main.now_iso()).strip() or self.desktop_main.now_iso()
        sid = str(serie_id or "").strip() or self._billing_default_serie_id(issue_txt)
        ensure_series_fn = getattr(self.desktop_main, "ensure_at_series_record", None)
        if callable(ensure_series_fn):
            serie_obj = ensure_series_fn(
                self.ensure_data(),
                doc_type="FT",
                serie_id=sid,
                issue_date=issue_txt,
                validation_code_hint=str(validation_code_hint or "").strip(),
            )
        else:
            serie_obj = {
                "doc_type": "FT",
                "serie_id": sid,
                "inicio_sequencia": 1,
                "next_seq": 1,
                "validation_code": str(validation_code_hint or "").strip(),
            }
        start_seq = max(1, int(self._parse_float(serie_obj.get("inicio_sequencia", 1), 1) or 1))
        seq = max(start_seq, int(self._parse_float(serie_obj.get("next_seq", start_seq), start_seq) or start_seq))
        used_numbers: set[str] = set()
        used_seq: set[tuple[str, int]] = set()
        year = issue_txt[:4] if len(issue_txt) >= 4 and issue_txt[:4].isdigit() else str(self.desktop_main.datetime.now().year)
        for record in self._billing_records():
            if not isinstance(record, dict):
                continue
            for row in list(record.get("faturas", []) or []):
                if not isinstance(row, dict):
                    continue
                number = str(row.get("numero_fatura", "") or "").strip()
                if number:
                    used_numbers.add(number)
                row_sid = str(row.get("serie_id", "") or row.get("serie", "") or "").strip()
                row_seq = int(self._parse_float(row.get("seq_num", 0), 0) or 0)
                if row_sid and row_seq > 0:
                    used_seq.add((row_sid, row_seq))
        while True:
            number = f"FT-{year}-{seq:04d}"
            if number not in used_numbers and (sid, seq) not in used_seq:
                break
            seq += 1
        validation_code = str(serie_obj.get("validation_code", "") or "").strip()
        if reserve:
            serie_obj["next_seq"] = seq + 1
            if validation_code:
                serie_obj["status"] = "REGISTADA"
            serie_obj["updated_at"] = self.desktop_main.now_iso()
        return {
            "doc_type": "FT",
            "numero_fatura": number,
            "serie": sid,
            "serie_id": sid,
            "seq_num": seq,
            "at_validation_code": validation_code,
            "atcud": f"{validation_code}-{seq}" if validation_code else "",
        }

    def _billing_quote_source(self, quote: dict[str, Any]) -> dict[str, Any]:
        detail = self.orc_detail(str(quote.get("numero", "") or "").strip())
        iva_perc = round(self._parse_float(detail.get("iva_perc", 23), 23), 2)
        lines: list[dict[str, Any]] = []
        for row in list(detail.get("linhas", []) or []):
            qty = round(self._parse_float(row.get("qtd", 0), 0), 2)
            if qty <= 0:
                continue
            line_type = self.desktop_main.normalize_orc_line_type(row.get("tipo_item"))
            reference = (
                str(row.get("ref_externa", "") or "").strip()
                or str(row.get("ref_interna", "") or "").strip()
                or str(row.get("produto_codigo", "") or "").strip()
                or str(row.get("conjunto_codigo", "") or "").strip()
            )
            description = str(row.get("descricao", "") or "").strip() or reference or "Artigo"
            material = str(row.get("material", "") or "").strip()
            espessura = str(row.get("espessura", "") or "").strip()
            if line_type == self.desktop_main.ORC_LINE_TYPE_PIECE and material and espessura:
                description = f"{description} | {material} {espessura} mm"
            unit_price = round(self._parse_float(row.get("preco_unit", 0), 0), 4)
            subtotal = round(qty * unit_price, 2)
            tax_value = round(subtotal * (iva_perc / 100.0), 2)
            lines.append(
                {
                    "reference": reference or "-",
                    "description": description or "-",
                    "quantity": qty,
                    "unit": str(row.get("produto_unid", "") or "UN").strip() or "UN",
                    "unit_price": unit_price,
                    "iva_perc": iva_perc,
                    "subtotal": subtotal,
                    "valor_iva": tax_value,
                    "total": round(subtotal + tax_value, 2),
                    "ref_interna": str(row.get("ref_interna", "") or "").strip(),
                    "ref_externa": str(row.get("ref_externa", "") or "").strip(),
                    "peca_id": "",
                }
            )
        transport = round(self._parse_float(detail.get("preco_transporte", 0), 0), 2)
        if transport > 0:
            transport_tax = round(transport * (iva_perc / 100.0), 2)
            lines.append(
                {
                    "reference": "TRANSP",
                    "description": "Transporte",
                    "quantity": 1.0,
                    "unit": "SV",
                    "unit_price": transport,
                    "iva_perc": iva_perc,
                    "subtotal": transport,
                    "valor_iva": transport_tax,
                    "total": round(transport + transport_tax, 2),
                    "ref_interna": "",
                    "ref_externa": "",
                    "peca_id": "",
                }
            )
        return {
            "iva_perc": iva_perc,
            "subtotal": round(sum(self._parse_float(row.get("subtotal", 0), 0) for row in lines), 2),
            "valor_iva": round(sum(self._parse_float(row.get("valor_iva", 0), 0) for row in lines), 2),
            "total": round(sum(self._parse_float(row.get("total", 0), 0) for row in lines), 2),
            "lines": lines,
        }

    def _billing_order_source(self, order: dict[str, Any]) -> dict[str, Any]:
        enc = self.get_encomenda_by_numero(str(order.get("numero", "") or "").strip()) or dict(order or {})
        iva_perc = 23.0
        lines: list[dict[str, Any]] = []
        for piece in list(self.desktop_main.encomenda_pecas(enc) or []):
            qty = round(self._parse_float(piece.get("quantidade_pedida", 0), 0), 2)
            if qty <= 0:
                continue
            unit_price = round(self._parse_float(piece.get("preco_unit", 0), 0), 4)
            subtotal = round(qty * unit_price, 2)
            tax_value = round(subtotal * (iva_perc / 100.0), 2)
            material = str(piece.get("material", "") or "").strip()
            espessura = str(piece.get("espessura", "") or "").strip()
            description = str(piece.get("descricao", "") or piece.get("Observacoes", "") or "").strip() or "Peca"
            if material and espessura:
                description = f"{description} | {material} {espessura} mm"
            lines.append(
                {
                    "reference": str(piece.get("ref_externa", "") or piece.get("ref_interna", "") or piece.get("id", "")).strip() or "-",
                    "description": description,
                    "quantity": qty,
                    "unit": "UN",
                    "unit_price": unit_price,
                    "iva_perc": iva_perc,
                    "subtotal": subtotal,
                    "valor_iva": tax_value,
                    "total": round(subtotal + tax_value, 2),
                    "ref_interna": str(piece.get("ref_interna", "") or "").strip(),
                    "ref_externa": str(piece.get("ref_externa", "") or "").strip(),
                    "peca_id": str(piece.get("id", "") or "").strip(),
                }
            )
        for item in list(enc.get("montagem_itens", []) or []):
            qty = round(self._parse_float(item.get("qtd_planeada", item.get("qtd", 0)), 0), 2)
            if qty <= 0:
                continue
            unit_price = round(self._parse_float(item.get("preco_unit", 0), 0), 4)
            subtotal = round(qty * unit_price, 2)
            tax_value = round(subtotal * (iva_perc / 100.0), 2)
            lines.append(
                {
                    "reference": str(item.get("produto_codigo", "") or item.get("conjunto_codigo", "") or "").strip() or "ITEM",
                    "description": str(item.get("descricao", "") or item.get("conjunto_nome", "") or "Item de montagem").strip(),
                    "quantity": qty,
                    "unit": str(item.get("produto_unid", "") or "UN").strip() or "UN",
                    "unit_price": unit_price,
                    "iva_perc": iva_perc,
                    "subtotal": subtotal,
                    "valor_iva": tax_value,
                    "total": round(subtotal + tax_value, 2),
                    "ref_interna": "",
                    "ref_externa": "",
                    "peca_id": "",
                }
            )
        return {
            "iva_perc": iva_perc,
            "subtotal": round(sum(self._parse_float(row.get("subtotal", 0), 0) for row in lines), 2),
            "valor_iva": round(sum(self._parse_float(row.get("valor_iva", 0), 0) for row in lines), 2),
            "total": round(sum(self._parse_float(row.get("total", 0), 0) for row in lines), 2),
            "lines": lines,
        }

    def _billing_apply_guide_filter(self, lines: list[dict[str, Any]], guide_number: str) -> list[dict[str, Any]]:
        guide_num = str(guide_number or "").strip()
        if not guide_num:
            return [dict(row) for row in lines]
        try:
            guide = self.expedicao_detail(guide_num)
        except Exception:
            return [dict(row) for row in lines]
        qty_by_piece: dict[str, float] = {}
        qty_by_ref_int: dict[str, float] = {}
        qty_by_ref_ext: dict[str, float] = {}
        for row in list(guide.get("lines", []) or []):
            qty = round(self._parse_float(row.get("qtd", 0), 0), 2)
            if qty <= 0:
                continue
            piece_id = str(row.get("peca_id", "") or "").strip()
            ref_int = str(row.get("ref_interna", "") or "").strip()
            ref_ext = str(row.get("ref_externa", "") or "").strip()
            if piece_id:
                qty_by_piece[piece_id] = round(qty_by_piece.get(piece_id, 0.0) + qty, 2)
            if ref_int:
                qty_by_ref_int[ref_int] = round(qty_by_ref_int.get(ref_int, 0.0) + qty, 2)
            if ref_ext:
                qty_by_ref_ext[ref_ext] = round(qty_by_ref_ext.get(ref_ext, 0.0) + qty, 2)
        filtered: list[dict[str, Any]] = []
        for row in lines:
            qty = 0.0
            if str(row.get("peca_id", "") or "").strip():
                qty = qty_by_piece.get(str(row.get("peca_id", "") or "").strip(), 0.0)
            if qty <= 0 and str(row.get("ref_interna", "") or "").strip():
                qty = qty_by_ref_int.get(str(row.get("ref_interna", "") or "").strip(), 0.0)
            if qty <= 0 and str(row.get("ref_externa", "") or "").strip():
                qty = qty_by_ref_ext.get(str(row.get("ref_externa", "") or "").strip(), 0.0)
            if qty <= 0:
                continue
            new_row = dict(row)
            new_row["quantity"] = min(qty, round(self._parse_float(row.get("quantity", 0), 0), 2))
            new_row["subtotal"] = round(new_row["quantity"] * self._parse_float(new_row.get("unit_price", 0), 0), 2)
            new_row["valor_iva"] = round(new_row["subtotal"] * (self._parse_float(new_row.get("iva_perc", 0), 0) / 100.0), 2)
            new_row["total"] = round(new_row["subtotal"] + new_row["valor_iva"], 2)
            filtered.append(new_row)
        return filtered or [dict(row) for row in lines]

    def _billing_recalculate_lines(self, lines: list[dict[str, Any]]) -> dict[str, Any]:
        subtotal = round(sum(self._parse_float(row.get("subtotal", 0), 0) for row in lines), 2)
        tax_value = round(sum(self._parse_float(row.get("valor_iva", 0), 0) for row in lines), 2)
        return {
            "subtotal": subtotal,
            "valor_iva": tax_value,
            "total": round(subtotal + tax_value, 2),
            "lines": lines,
        }

    def _billing_source_snapshot(
        self,
        *,
        record: dict[str, Any],
        quote: dict[str, Any] | None = None,
        order: dict[str, Any] | None = None,
        guide_number: str = "",
    ) -> dict[str, Any]:
        if isinstance(quote, dict):
            source = self._billing_quote_source(quote)
        elif isinstance(order, dict):
            source = self._billing_order_source(order)
        else:
            source = {"iva_perc": 23.0, "subtotal": 0.0, "valor_iva": 0.0, "total": 0.0, "lines": []}
        source["lines"] = self._billing_apply_guide_filter(list(source.get("lines", []) or []), guide_number)
        return self._billing_recalculate_lines(list(source.get("lines", []) or [])) | {"iva_perc": round(self._parse_float(source.get("iva_perc", 23), 23), 2)}

    def _billing_adjust_document_total(
        self,
        lines: list[dict[str, Any]],
        *,
        target_total: float,
        default_iva: float,
    ) -> list[dict[str, Any]]:
        current_total = round(sum(self._parse_float(row.get("total", 0), 0) for row in lines), 2)
        diff = round(target_total - current_total, 2)
        if abs(diff) <= 0.02:
            return lines
        iva_perc = round(self._parse_float(default_iva, 23), 2)
        if iva_perc <= -100:
            adj_subtotal = diff
        else:
            adj_subtotal = round(diff / (1.0 + (iva_perc / 100.0)), 2)
        adj_tax = round(diff - adj_subtotal, 2)
        lines = list(lines or [])
        lines.append(
            {
                "reference": "AJUSTE",
                "description": "Ajuste de faturacao",
                "quantity": 1.0,
                "unit": "SV",
                "unit_price": adj_subtotal,
                "iva_perc": iva_perc,
                "subtotal": adj_subtotal,
                "valor_iva": adj_tax,
                "total": round(adj_subtotal + adj_tax, 2),
                "ref_interna": "",
                "ref_externa": "",
                "peca_id": "",
            }
        )
        return lines

    def _billing_build_invoice_document(self, record_number: str, invoice: dict[str, Any], *, prefer_snapshot: bool = True) -> dict[str, Any]:
        reg_num = str(record_number or "").strip()
        record = self._billing_find_record(reg_num)
        if record is None:
            raise ValueError("Registo de faturacao nao encontrado.")
        software_cert = self._billing_software_cert_number()
        invoice_id = str(invoice.get("id", "") or "").strip()
        if prefer_snapshot:
            snapshot = self._billing_invoice_snapshot(invoice)
            if snapshot:
                received = round(
                    sum(
                        self._parse_float(row.get("valor", 0), 0)
                        for row in list(record.get("pagamentos", []) or [])
                        if str(row.get("fatura_id", "") or "").strip() == invoice_id
                    ),
                    2,
                )
                total_amount = round(self._parse_float(snapshot.get("valor_total", 0), 0), 2)
                document = dict(snapshot)
                document["software_cert"] = software_cert
                document["valor_recebido"] = received
                document["saldo"] = round(max(0.0, total_amount - received), 2)
                return document
        quote_num = str(record.get("orcamento_numero", "") or "").strip()
        order_num = str(record.get("encomenda_numero", "") or "").strip()
        quote = self._billing_quote_by_number(quote_num) if quote_num else None
        order = self._billing_order_by_number(order_num) if order_num else None
        client = self._billing_client_snapshot(quote=quote, order=order, record=record)
        issuer = dict(getattr(self.desktop_main, "get_guia_emitente_info", lambda: {})() or {})
        source = self._billing_source_snapshot(
            record=record,
            quote=quote,
            order=order,
            guide_number=str(invoice.get("guia_numero", "") or "").strip(),
        )
        lines = list(source.get("lines", []) or [])
        target_total = round(self._parse_float(invoice.get("valor_total", source.get("total", 0)), 0), 2)
        default_iva = round(self._parse_float(invoice.get("iva_perc", source.get("iva_perc", 23)), 23), 2)
        if not lines:
            manual_subtotal = round(target_total / (1.0 + (default_iva / 100.0)), 2) if default_iva > -100 else target_total
            manual_tax = round(target_total - manual_subtotal, 2)
            lines = [
                {
                    "reference": "SERVICO",
                    "description": f"Venda associada ao registo {reg_num}",
                    "quantity": 1.0,
                    "unit": "SV",
                    "unit_price": manual_subtotal,
                    "iva_perc": default_iva,
                    "subtotal": manual_subtotal,
                    "valor_iva": manual_tax,
                    "total": round(manual_subtotal + manual_tax, 2),
                    "ref_interna": "",
                    "ref_externa": "",
                    "peca_id": "",
                }
            ]
        elif target_total > 0:
            lines = self._billing_adjust_document_total(lines, target_total=target_total, default_iva=default_iva)
        totals = self._billing_recalculate_lines(lines)
        recebido = round(
            sum(
                self._parse_float(row.get("valor", 0), 0)
                for row in list(record.get("pagamentos", []) or [])
                if str(row.get("fatura_id", "") or "").strip() == invoice_id
            ),
            2,
        )
        saldo = round(max(0.0, totals["total"] - recebido), 2)
        tax_summary: dict[float, dict[str, Any]] = {}
        for row in lines:
            rate = round(self._parse_float(row.get("iva_perc", default_iva), default_iva), 2)
            bucket = tax_summary.setdefault(rate, {"rate": rate, "base": 0.0, "tax": 0.0})
            bucket["base"] = round(bucket["base"] + self._parse_float(row.get("subtotal", 0), 0), 2)
            bucket["tax"] = round(bucket["tax"] + self._parse_float(row.get("valor_iva", 0), 0), 2)
        tax_summary_rows = [
            {
                "rate": rate,
                "rate_label": f"{self._fmt(rate)}%",
                "base": values["base"],
                "tax": values["tax"],
                "label": f"Base {self._fmt(rate)}%",
            }
            for rate, values in sorted(tax_summary.items(), key=lambda item: item[0])
        ]
        return {
            "titulo": "Fatura",
            "subtitulo": "Documento comercial e fiscal",
            "doc_type": str(invoice.get("doc_type", "") or "FT").strip() or "FT",
            "numero_fatura": str(invoice.get("numero_fatura", "") or "").strip(),
            "serie": str(invoice.get("serie", "") or invoice.get("serie_id", "") or "").strip(),
            "serie_id": str(invoice.get("serie_id", "") or invoice.get("serie", "") or "").strip(),
            "seq_num": int(self._parse_float(invoice.get("seq_num", 0), 0) or 0),
            "at_validation_code": str(invoice.get("at_validation_code", "") or "").strip(),
            "atcud": str(invoice.get("atcud", "") or "").strip(),
            "legal_invoice_no": str(invoice.get("legal_invoice_no", "") or self._billing_legal_invoice_no(invoice)).strip(),
            "system_entry_date": str(invoice.get("system_entry_date", "") or invoice.get("created_at", "") or self.desktop_main.now_iso()).strip(),
            "source_id": str(invoice.get("source_id", "") or self._billing_actor()).strip() or "Sistema",
            "status_source_id": self._billing_status_source_id(invoice),
            "source_billing": self.tax_compliance.normalize_source_billing(invoice.get("source_billing", ""), fallback="P"),
            "hash": str(invoice.get("hash", "") or "").strip(),
            "hash_control": str(invoice.get("hash_control", "") or "").strip(),
            "previous_hash": str(invoice.get("previous_hash", "") or "").strip(),
            "data_emissao": str(invoice.get("data_emissao", "") or self.desktop_main.now_iso())[:10],
            "data_vencimento": str(invoice.get("data_vencimento", "") or record.get("data_vencimento", "") or "").strip()[:10],
            "moeda": str(invoice.get("moeda", "") or "EUR").strip() or "EUR",
            "issuer": {
                "nome": str(issuer.get("nome", "") or "").strip(),
                "nif": str(issuer.get("nif", "") or "").strip(),
                "morada": str(issuer.get("morada", "") or "").strip(),
                "extra": "Portugal",
            },
            "customer": client,
            "references": {
                "registo": reg_num,
                "orcamento": quote_num,
                "encomenda": order_num,
                "guia": str(invoice.get("guia_numero", "") or "").strip(),
            },
            "subtotal": totals["subtotal"],
            "valor_iva": totals["valor_iva"],
            "valor_total": totals["total"],
            "valor_recebido": recebido,
            "saldo": saldo,
            "tax_summary": tax_summary_rows,
            "lines": lines,
            "obs": str(invoice.get("obs", "") or record.get("obs", "") or "").strip(),
            "software_cert": software_cert,
            "qr_payload": (
                f"ATCUD:{str(invoice.get('atcud', '') or '').strip() or '-'}"
                f"|DOC:{str(invoice.get('numero_fatura', '') or '').strip() or '-'}"
                f"|DT:{str(invoice.get('data_emissao', '') or '').strip()[:10] or '-'}"
                f"|A:{str(issuer.get('nif', '') or '').strip() or '-'}"
                f"|B:{str(client.get('nif', '') or '').strip() or '-'}"
                f"|TOT:{totals['total']:.2f}"
            ),
        }

    def billing_invoice_defaults(self, numero: str, invoice_id: str = "") -> dict[str, Any]:
        reg_num = str(numero or "").strip()
        record = self._billing_find_record(reg_num)
        if record is None:
            raise ValueError("Registo de faturacao nao encontrado.")
        row_id = str(invoice_id or "").strip()
        existing = next((row for row in list(record.get("faturas", []) or []) if str(row.get("id", "") or "").strip() == row_id), None) if row_id else None
        issue_date = str((existing or {}).get("data_emissao", "") or self.desktop_main.now_iso())[:10]
        base_data = dict(existing or {})
        if existing is None:
            base_data.update(self._billing_next_invoice_identifiers(issue_date=issue_date, reserve=False))
        quote = self._billing_quote_by_number(str(record.get("orcamento_numero", "") or "").strip()) if str(record.get("orcamento_numero", "") or "").strip() else None
        order = self._billing_order_by_number(str(record.get("encomenda_numero", "") or "").strip()) if str(record.get("encomenda_numero", "") or "").strip() else None
        quote, order = self._billing_sync_record_source(record, quote=quote, order=order, persist=True)
        client = self._billing_client_snapshot(
            quote=quote,
            order=order,
            record=record,
        )
        due_date = str((existing or {}).get("data_vencimento", "") or record.get("data_vencimento", "") or "").strip()[:10]
        if not due_date:
            try:
                base_date = self.desktop_main.date.fromisoformat(issue_date)
                due_date = (base_date + self.desktop_main.timedelta(days=self._billing_due_days_from_text(client.get("cond_pagamento", "")))).isoformat()
            except Exception:
                due_date = issue_date
        guides = self._billing_guides_for_order(str(record.get("encomenda_numero", "") or "").strip())
        guide_default = str((existing or {}).get("guia_numero", "") or "").strip()
        if not guide_default:
            guide_default = str((guides[0] if guides else {}).get("numero", "") or "").strip()
        base_data.setdefault("doc_type", "FT")
        base_data.setdefault("guia_numero", guide_default)
        base_data.setdefault("data_emissao", issue_date)
        base_data.setdefault("data_vencimento", due_date)
        base_data.setdefault("moeda", "EUR")
        estimated_total = round(self._billing_record_sale_total(record, quote), 2)
        if self._parse_float(base_data.get("valor_total", 0), 0) <= 0 and estimated_total > 0:
            base_data["valor_total"] = estimated_total
        base_data["_allow_zero_total"] = True
        provisional = self._billing_normalize_invoice(base_data, existing)
        doc = self._billing_build_invoice_document(reg_num, provisional)
        provisional["iva_perc"] = round(self._parse_float((existing or {}).get("iva_perc", doc.get("tax_summary", [{}])[0].get("rate", 23) if list(doc.get("tax_summary", []) or []) else 23), 23), 2)
        provisional["subtotal"] = round(self._parse_float((existing or {}).get("subtotal", doc.get("subtotal", 0)), doc.get("subtotal", 0)), 2)
        provisional["valor_iva"] = round(self._parse_float((existing or {}).get("valor_iva", doc.get("valor_iva", 0)), doc.get("valor_iva", 0)), 2)
        provisional["valor_total"] = round(self._parse_float((existing or {}).get("valor_total", doc.get("valor_total", 0)), doc.get("valor_total", 0)), 2)
        provisional["guide_options"] = [
            {
                "numero": str(row.get("numero", "") or "").strip(),
                "label": f"{str(row.get('numero', '') or '').strip()} | {str(row.get('data_emissao', '') or '').strip()[:10]}",
            }
            for row in guides
        ]
        return provisional

    def _billing_normalize_invoice(self, payload: dict[str, Any], existing: dict[str, Any] | None = None) -> dict[str, Any]:
        row = dict(existing or {})
        row_id = str(payload.get("id", "") or row.get("id", "") or self.desktop_main.uuid.uuid4().hex[:12].upper()).strip()
        doc_type = str(payload.get("doc_type", "") or row.get("doc_type", "") or "FT").strip() or "FT"
        numero_fatura = str(payload.get("numero_fatura", "") or row.get("numero_fatura", "") or "").strip()
        serie = str(payload.get("serie", "") or row.get("serie", "") or "").strip()
        serie_id = str(payload.get("serie_id", "") or row.get("serie_id", "") or serie).strip()
        seq_num = int(self._parse_float(payload.get("seq_num", row.get("seq_num", 0)), 0) or 0)
        at_validation_code = str(payload.get("at_validation_code", "") or row.get("at_validation_code", "") or "").strip()
        atcud = str(payload.get("atcud", "") or row.get("atcud", "") or "").strip()
        guia_numero = str(payload.get("guia_numero", "") or row.get("guia_numero", "") or "").strip()
        data_emissao = str(payload.get("data_emissao", "") or row.get("data_emissao", "") or "").strip()[:10]
        data_vencimento = str(payload.get("data_vencimento", "") or row.get("data_vencimento", "") or "").strip()[:10]
        moeda = str(payload.get("moeda", "") or row.get("moeda", "") or "EUR").strip() or "EUR"
        iva_perc = round(self._parse_float(payload.get("iva_perc", row.get("iva_perc", 23)), 23), 2)
        subtotal = round(self._parse_float(payload.get("subtotal", row.get("subtotal", 0)), 0), 2)
        valor_iva = round(self._parse_float(payload.get("valor_iva", row.get("valor_iva", 0)), 0), 2)
        valor_total = round(self._parse_float(payload.get("valor_total", row.get("valor_total", 0)), 0), 2)
        caminho = str(payload.get("caminho", "") or row.get("caminho", "") or "").strip()
        obs = str(payload.get("obs", "") or row.get("obs", "") or "").strip()
        estado = str(payload.get("estado", "") or row.get("estado", "") or "Emitida").strip() or "Emitida"
        anulada = bool(payload.get("anulada", row.get("anulada", False)))
        anulada_motivo = str(payload.get("anulada_motivo", "") or row.get("anulada_motivo", "") or "").strip()
        anulada_at = str(payload.get("anulada_at", "") or row.get("anulada_at", "") or "").strip()
        legal_invoice_no = str(payload.get("legal_invoice_no", "") or row.get("legal_invoice_no", "") or "").strip()
        system_entry_date = str(payload.get("system_entry_date", "") or row.get("system_entry_date", "") or "").strip()
        source_id = str(payload.get("source_id", "") or row.get("source_id", "") or "").strip()
        source_billing = str(payload.get("source_billing", "") or row.get("source_billing", "") or "").strip()
        status_source_id = str(payload.get("status_source_id", "") or row.get("status_source_id", "") or "").strip()
        hash_value = str(payload.get("hash", "") or row.get("hash", "") or "").strip()
        hash_control = str(payload.get("hash_control", "") or row.get("hash_control", "") or "").strip()
        previous_hash = str(payload.get("previous_hash", "") or row.get("previous_hash", "") or "").strip()
        document_snapshot_json = str(payload.get("document_snapshot_json", "") or row.get("document_snapshot_json", "") or "")
        communication_status = str(payload.get("communication_status", "") or row.get("communication_status", "") or "").strip()
        communication_filename = str(payload.get("communication_filename", "") or row.get("communication_filename", "") or "").strip()
        communication_error = str(payload.get("communication_error", "") or row.get("communication_error", "") or "").strip()
        communicated_at = str(payload.get("communicated_at", "") or row.get("communicated_at", "") or "").strip()
        communication_batch_id = str(payload.get("communication_batch_id", "") or row.get("communication_batch_id", "") or "").strip()
        if anulada or "anulad" in self.desktop_main.norm_text(estado):
            anulada = True
            estado = "Anulada"
        allow_zero_total = bool(payload.get("_allow_zero_total"))
        if not numero_fatura and not caminho:
            raise ValueError("Indica o numero da fatura ou associa o ficheiro.")
        if valor_total <= 0 and not allow_zero_total:
            raise ValueError("Valor da fatura invalido.")
        if not data_emissao:
            data_emissao = str(self.desktop_main.now_iso())[:10]
        return {
            "id": row_id,
            "doc_type": doc_type,
            "numero_fatura": numero_fatura,
            "serie": serie,
            "serie_id": serie_id,
            "seq_num": seq_num,
            "at_validation_code": at_validation_code,
            "atcud": atcud,
            "guia_numero": guia_numero,
            "data_emissao": data_emissao,
            "data_vencimento": data_vencimento,
            "moeda": moeda,
            "iva_perc": iva_perc,
            "subtotal": subtotal,
            "valor_iva": valor_iva,
            "valor_total": valor_total,
            "caminho": caminho,
            "obs": obs,
            "estado": estado,
            "anulada": anulada,
            "anulada_motivo": anulada_motivo,
            "anulada_at": anulada_at,
            "legal_invoice_no": legal_invoice_no,
            "system_entry_date": system_entry_date,
            "source_id": source_id,
            "source_billing": source_billing,
            "status_source_id": status_source_id,
            "hash": hash_value,
            "hash_control": hash_control,
            "previous_hash": previous_hash,
            "document_snapshot_json": document_snapshot_json,
            "communication_status": communication_status,
            "communication_filename": communication_filename,
            "communication_error": communication_error,
            "communicated_at": communicated_at,
            "communication_batch_id": communication_batch_id,
            "created_at": str(row.get("created_at", "") or self.desktop_main.now_iso()),
        }

    def _billing_invoice_is_void(self, invoice: dict[str, Any] | None) -> bool:
        if not isinstance(invoice, dict):
            return False
        if bool(invoice.get("anulada")):
            return True
        return "anulad" in self.desktop_main.norm_text(invoice.get("estado", ""))

    def _billing_active_invoices(self, record: dict[str, Any] | None) -> list[dict[str, Any]]:
        if not isinstance(record, dict):
            return []
        return [
            row
            for row in list(record.get("faturas", []) or [])
            if isinstance(row, dict) and not self._billing_invoice_is_void(row)
        ]

    def _billing_effective_payments(self, record: dict[str, Any] | None) -> list[dict[str, Any]]:
        if not isinstance(record, dict):
            return []
        invoices_by_id = {
            str(row.get("id", "") or "").strip(): row
            for row in list(record.get("faturas", []) or [])
            if isinstance(row, dict)
        }
        effective: list[dict[str, Any]] = []
        for row in list(record.get("pagamentos", []) or []):
            if not isinstance(row, dict):
                continue
            invoice_id = str(row.get("fatura_id", "") or "").strip()
            if invoice_id and self._billing_invoice_is_void(invoices_by_id.get(invoice_id)):
                continue
            effective.append(row)
        return effective

    def _billing_normalize_payment(self, payload: dict[str, Any], existing: dict[str, Any] | None = None) -> dict[str, Any]:
        row = dict(existing or {})
        row_id = str(payload.get("id", "") or row.get("id", "") or self.desktop_main.uuid.uuid4().hex[:12].upper()).strip()
        data_pagamento = str(payload.get("data_pagamento", "") or row.get("data_pagamento", "") or "").strip()[:10]
        valor = round(self._parse_float(payload.get("valor", row.get("valor", 0)), 0), 2)
        metodo = str(payload.get("metodo", "") or row.get("metodo", "") or "").strip()
        referencia = str(payload.get("referencia", "") or row.get("referencia", "") or "").strip()
        titulo = str(payload.get("titulo_comprovativo", "") or row.get("titulo_comprovativo", "") or "").strip()
        caminho = str(payload.get("caminho_comprovativo", "") or row.get("caminho_comprovativo", "") or "").strip()
        fatura_id = str(payload.get("fatura_id", "") or row.get("fatura_id", "") or "").strip()
        obs = str(payload.get("obs", "") or row.get("obs", "") or "").strip()
        if valor <= 0:
            raise ValueError("Valor do pagamento invalido.")
        if not data_pagamento:
            data_pagamento = str(self.desktop_main.now_iso())[:10]
        return {
            "id": row_id,
            "fatura_id": fatura_id,
            "data_pagamento": data_pagamento,
            "valor": valor,
            "metodo": metodo,
            "referencia": referencia,
            "titulo_comprovativo": titulo,
            "caminho_comprovativo": caminho,
            "obs": obs,
            "created_at": str(row.get("created_at", "") or self.desktop_main.now_iso()),
        }

    def _billing_record_sale_total(self, record: dict[str, Any], quote: dict[str, Any] | None = None) -> float:
        manual = round(self._parse_float(record.get("valor_venda_manual", 0), 0), 2)
        if manual > 0:
            return manual
        if isinstance(quote, dict):
            total_quote = round(self._parse_float(quote.get("total", 0), 0), 2)
            if total_quote > 0:
                return total_quote
        invoices_total = round(
            sum(self._parse_float(row.get("valor_total", 0), 0) for row in self._billing_active_invoices(record)),
            2,
        )
        return invoices_total

    def _billing_unpaid_due_date(self, record: dict[str, Any]) -> str:
        invoices = self._billing_active_invoices(record)
        payments_total = round(
            sum(self._parse_float(row.get("valor", 0), 0) for row in self._billing_effective_payments(record)),
            2,
        )
        invoices_total = round(sum(self._parse_float(row.get("valor_total", 0), 0) for row in invoices), 2)
        if invoices_total <= 0 or payments_total >= (invoices_total - 0.009):
            return ""
        due_candidates = [
            str(row.get("data_vencimento", "") or "").strip()[:10]
            for row in invoices
            if str(row.get("data_vencimento", "") or "").strip()
        ]
        if due_candidates:
            return sorted(due_candidates)[0]
        return str(record.get("data_vencimento", "") or "").strip()[:10]

    def _billing_invoice_status(self, record: dict[str, Any], quote: dict[str, Any] | None = None) -> str:
        sale_total = self._billing_record_sale_total(record, quote)
        invoices_total = round(sum(self._parse_float(row.get("valor_total", 0), 0) for row in self._billing_active_invoices(record)), 2)
        if invoices_total <= 0:
            return "Por faturar"
        if sale_total > 0 and invoices_total < (sale_total - 0.009):
            return "Faturada parcial"
        return "Faturada"

    def _billing_payment_status(self, record: dict[str, Any]) -> str:
        manual = str(record.get("estado_pagamento_manual", "") or "").strip()
        if manual and self.desktop_main.norm_text(manual) not in {"auto", "automatico", "automático"}:
            return manual
        invoices_total = round(sum(self._parse_float(row.get("valor_total", 0), 0) for row in self._billing_active_invoices(record)), 2)
        payments_total = round(
            sum(self._parse_float(row.get("valor", 0), 0) for row in self._billing_effective_payments(record)),
            2,
        )
        if invoices_total <= 0:
            return "Sem faturação"
        if payments_total >= (invoices_total - 0.009):
            return "Paga"
        if payments_total > 0:
            return "Parcial"
        due_date = self._billing_unpaid_due_date(record)
        today = self.desktop_main.now_iso()[:10]
        if due_date and due_date < today:
            return "Atrasada"
        return "Pendente"

    def _billing_sync_record_source(
        self,
        record: dict[str, Any] | None,
        *,
        quote: dict[str, Any] | None = None,
        order: dict[str, Any] | None = None,
        persist: bool = False,
    ) -> tuple[dict[str, Any] | None, dict[str, Any] | None]:
        if not isinstance(record, dict):
            return quote, order
        quote_num = str(record.get("orcamento_numero", "") or "").strip()
        order_num = str(record.get("encomenda_numero", "") or "").strip()
        if quote is None and quote_num:
            quote = self._billing_quote_by_number(quote_num)
        if order is None and order_num:
            order = self._billing_order_by_number(order_num)
        if order is None and isinstance(quote, dict):
            linked_order_num = str(quote.get("numero_encomenda", "") or "").strip()
            if linked_order_num:
                order = self._billing_order_by_number(linked_order_num)
        if quote is None and isinstance(order, dict):
            linked_quote_num = str(order.get("numero_orcamento", "") or "").strip()
            if linked_quote_num:
                quote = self._billing_quote_by_number(linked_quote_num)
        changed = False
        if isinstance(quote, dict):
            synced_quote_num = str(quote.get("numero", "") or "").strip()
            if synced_quote_num and synced_quote_num != quote_num:
                record["orcamento_numero"] = synced_quote_num
                quote_num = synced_quote_num
                changed = True
        if isinstance(order, dict):
            synced_order_num = str(order.get("numero", "") or "").strip()
            if synced_order_num and synced_order_num != order_num:
                record["encomenda_numero"] = synced_order_num
                order_num = synced_order_num
                changed = True
        client = self._billing_client_info(quote=quote, order=order, record=record)
        client_code = str(client.get("codigo", "") or "").strip()
        client_name = str(client.get("nome", "") or "").strip()
        if client_code and client_code != str(record.get("cliente_codigo", "") or "").strip():
            record["cliente_codigo"] = client_code
            changed = True
        if client_name and client_name != str(record.get("cliente_nome", "") or "").strip():
            record["cliente_nome"] = client_name
            changed = True
        if not str(record.get("data_venda", "") or "").strip():
            sale_date = (
                str((quote or {}).get("data", "") or "").strip()[:10]
                or str((order or {}).get("data_criacao", "") or "").strip()[:10]
            )
            if sale_date:
                record["data_venda"] = sale_date
                changed = True
        if persist and changed:
            record["updated_at"] = self.desktop_main.now_iso()
            self._save(force=True)
        return quote, order

    def _billing_build_row(
        self,
        *,
        source_type: str,
        quote: dict[str, Any] | None = None,
        order: dict[str, Any] | None = None,
        record: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        rec = dict(record or {})
        quote_num = str((quote or {}).get("numero", "") or rec.get("orcamento_numero", "") or "").strip()
        order_num = str((order or {}).get("numero", "") or rec.get("encomenda_numero", "") or "").strip()
        if order is None and order_num:
            order = self._billing_order_by_number(order_num)
        if quote is None and quote_num:
            quote = self._billing_quote_by_number(quote_num)
        client = self._billing_client_info(quote=quote, order=order, record=rec)
        if isinstance(order, dict):
            try:
                self.desktop_main.update_estado_expedicao_encomenda(order)
            except Exception:
                pass
        sale_total = self._billing_record_sale_total(rec, quote)
        invoices_total = round(sum(self._parse_float(row.get("valor_total", 0), 0) for row in self._billing_active_invoices(rec)), 2)
        payments_total = round(sum(self._parse_float(row.get("valor", 0), 0) for row in self._billing_effective_payments(rec)), 2)
        balance = round(max(0.0, invoices_total - payments_total), 2)
        guides = self._billing_guides_for_order(order_num)
        payment_status = self._billing_payment_status(rec) if rec else ("Sem faturação" if invoices_total <= 0 else "Pendente")
        invoice_status = self._billing_invoice_status(rec, quote) if rec else "Por faturar"
        sale_date = (
            str(rec.get("data_venda", "") or "").strip()[:10]
            or str((quote or {}).get("data", "") or "").strip()[:10]
            or str((order or {}).get("data_criacao", "") or "").strip()[:10]
        )
        year = sale_date[:4] if len(sale_date) >= 4 and sale_date[:4].isdigit() else str(self.desktop_main.datetime.now().year)
        latest_invoice = ""
        latest_invoice_date = ""
        if list(rec.get("faturas", []) or []):
            ordered_invoices = sorted(
                list(rec.get("faturas", []) or []),
                key=lambda row: (
                    str((row or {}).get("data_emissao", "") or ""),
                    str((row or {}).get("numero_fatura", "") or ""),
                ),
                reverse=True,
            )
            latest_invoice = str(ordered_invoices[0].get("numero_fatura", "") or "").strip()
            latest_invoice_date = str(ordered_invoices[0].get("data_emissao", "") or "").strip()[:10]
        source_label = "Orçamento vendido" if source_type == "quote" else ("Encomenda direta" if source_type == "order" else "Registo manual")
        return {
            "record_number": str(rec.get("numero", "") or "").strip(),
            "source_type": source_type,
            "source_number": quote_num if source_type == "quote" else (order_num or str(rec.get("numero", "") or "").strip()),
            "orcamento_numero": quote_num,
            "encomenda_numero": order_num,
            "cliente": client.get("label", "-") or "-",
            "cliente_codigo": client.get("codigo", ""),
            "cliente_nome": client.get("nome", ""),
            "origem": source_label,
            "estado_encomenda": str((order or {}).get("estado", "") or ("Sem encomenda" if quote_num else "Sem encomenda")).strip() or "Sem encomenda",
            "estado_expedicao": str((order or {}).get("estado_expedicao", "") or ("Sem encomenda" if quote_num else "Sem encomenda")).strip() or "Sem encomenda",
            "estado_faturacao": invoice_status,
            "estado_pagamento": payment_status,
            "vendido": sale_total,
            "faturado": invoices_total,
            "recebido": payments_total,
            "saldo": balance,
            "ultima_fatura": latest_invoice,
            "data_ultima_fatura": latest_invoice_date,
            "ultima_guia": str((guides[0] if guides else {}).get("numero", "") or "").strip(),
            "guias": len(guides),
            "data_venda": sale_date,
            "ano": year,
        }

    def billing_available_years(self) -> list[str]:
        years: set[str] = {str(self.desktop_main.datetime.now().year)}
        for row in self.billing_rows("", "Todas", "Todos"):
            year = str(row.get("ano", "") or "").strip()
            if year:
                years.add(year)
        return sorted(years, key=lambda value: int(value) if value.isdigit() else 0, reverse=True)

    def billing_rows(self, filter_text: str = "", state_filter: str = "Ativas", year: str = "Todos") -> list[dict[str, Any]]:
        data = self.ensure_data()
        query = str(filter_text or "").strip().lower()
        state_raw = str(state_filter or "Ativas").strip().lower()
        year_raw = str(year or "Todos").strip()
        rows: list[dict[str, Any]] = []
        seen_keys: set[str] = set()
        records = {str(row.get("numero", "") or "").strip(): dict(row or {}) for row in self._billing_records() if isinstance(row, dict)}

        def append_row(row: dict[str, Any], key: str) -> None:
            if key in seen_keys:
                return
            if year_raw and year_raw.lower() not in {"todos", "todas", "all", ""} and str(row.get("ano", "") or "").strip() != year_raw:
                return
            estado_faturacao = self.desktop_main.norm_text(row.get("estado_faturacao", ""))
            estado_pagamento = self.desktop_main.norm_text(row.get("estado_pagamento", ""))
            if state_raw and state_raw not in {"todos", "todas", "all", ""}:
                if "ativ" in state_raw and ("paga" in estado_pagamento and "atras" not in estado_pagamento):
                    return
                if "faturar" in state_raw and "por faturar" not in estado_faturacao:
                    return
                if "cobrar" in state_raw and not any(token in estado_pagamento for token in ("pendente", "parcial", "atras")):
                    return
                if "paga" in state_raw and "paga" not in estado_pagamento:
                    return
                if "atras" in state_raw and "atras" not in estado_pagamento:
                    return
            if query and not any(query in str(value).lower() for value in row.values()):
                return
            seen_keys.add(key)
            rows.append(row)

        for quote in list(data.get("orcamentos", []) or []):
            if not isinstance(quote, dict) or not self._billing_quote_is_sold(quote):
                continue
            quote_num = str(quote.get("numero", "") or "").strip()
            order_num = str(quote.get("numero_encomenda", "") or "").strip()
            record = self._billing_find_source_record(quote_num, order_num)
            order = self._billing_order_by_number(order_num) if order_num else None
            row = self._billing_build_row(source_type="quote", quote=quote, order=order, record=record)
            append_row(row, f"quote:{quote_num}")

        for order in list(data.get("encomendas", []) or []):
            if not isinstance(order, dict):
                continue
            order_num = str(order.get("numero", "") or "").strip()
            quote_num = str(order.get("numero_orcamento", "") or "").strip()
            if quote_num:
                continue
            record = self._billing_find_source_record("", order_num)
            row = self._billing_build_row(source_type="order", order=order, record=record)
            append_row(row, f"order:{order_num}")

        for record_num, record in records.items():
            quote_num = str(record.get("orcamento_numero", "") or "").strip()
            order_num = str(record.get("encomenda_numero", "") or "").strip()
            key = f"quote:{quote_num}" if quote_num else (f"order:{order_num}" if order_num else f"record:{record_num}")
            if key in seen_keys:
                continue
            row = self._billing_build_row(
                source_type="record" if not quote_num and not order_num else ("quote" if quote_num else "order"),
                quote=self._billing_quote_by_number(quote_num) if quote_num else None,
                order=self._billing_order_by_number(order_num) if order_num else None,
                record=record,
            )
            append_row(row, key)

        rows.sort(
            key=lambda item: (
                str(item.get("data_venda", "") or "0000-00-00"),
                str(item.get("record_number", "") or ""),
                str(item.get("orcamento_numero", "") or item.get("encomenda_numero", "") or ""),
            ),
            reverse=True,
        )
        return rows

    def billing_dashboard(self) -> dict[str, Any]:
        rows = self.billing_rows("", "Todas", "Todos")
        sold_total = round(sum(self._parse_float(row.get("vendido", 0), 0) for row in rows), 2)
        invoiced_total = round(sum(self._parse_float(row.get("faturado", 0), 0) for row in rows), 2)
        received_total = round(sum(self._parse_float(row.get("recebido", 0), 0) for row in rows), 2)
        balance_total = round(sum(self._parse_float(row.get("saldo", 0), 0) for row in rows), 2)
        return {
            "sold_total": sold_total,
            "invoiced_total": invoiced_total,
            "received_total": received_total,
            "balance_total": balance_total,
            "pending_invoice_count": sum(1 for row in rows if "por faturar" in self.desktop_main.norm_text(row.get("estado_faturacao", ""))),
            "open_payment_count": sum(1 for row in rows if any(token in self.desktop_main.norm_text(row.get("estado_pagamento", "")) for token in ("pendente", "parcial", "atras"))),
            "overdue_count": sum(1 for row in rows if "atras" in self.desktop_main.norm_text(row.get("estado_pagamento", ""))),
            "completed_orders": sum(1 for row in rows if "concl" in self.desktop_main.norm_text(row.get("estado_encomenda", ""))),
            "open_orders": sum(1 for row in rows if row.get("estado_encomenda") and "concl" not in self.desktop_main.norm_text(row.get("estado_encomenda", "")) and "sem encomenda" not in self.desktop_main.norm_text(row.get("estado_encomenda", ""))),
            "row_count": len(rows),
        }

    def _billing_create_record_from_source(self, source_type: str, source_number: str) -> dict[str, Any]:
        source_type_txt = str(source_type or "").strip().lower()
        source_number_txt = str(source_number or "").strip()
        quote = self._billing_quote_by_number(source_number_txt) if source_type_txt == "quote" else None
        order = self._billing_order_by_number(source_number_txt) if source_type_txt == "order" else None
        if quote is None and source_type_txt == "quote":
            raise ValueError("Orçamento não encontrado.")
        if order is None and source_type_txt == "order":
            raise ValueError("Encomenda não encontrada.")
        if quote is not None and not self._billing_quote_is_sold(quote):
            raise ValueError("O orçamento ainda não está marcado como vendido/aprovado.")
        order_num = str((quote or {}).get("numero_encomenda", "") or (order or {}).get("numero", "") or "").strip()
        quote_num = str((quote or {}).get("numero", "") or (order or {}).get("numero_orcamento", "") or "").strip()
        existing = self._billing_find_source_record(quote_num, order_num)
        if existing is not None:
            return existing
        client = self._billing_client_info(quote=quote, order=order)
        sale_date = (
            str((quote or {}).get("data", "") or "").strip()[:10]
            or str((order or {}).get("data_criacao", "") or "").strip()[:10]
            or str(self.desktop_main.now_iso())[:10]
        )
        due_date = ""
        try:
            due_date = (datetime.strptime(sale_date, "%Y-%m-%d") + timedelta(days=30)).strftime("%Y-%m-%d")
        except Exception:
            due_date = ""
        record = {
            "numero": self._billing_next_number(),
            "origem": "Orçamento" if quote_num else "Encomenda",
            "orcamento_numero": quote_num,
            "encomenda_numero": order_num,
            "cliente_codigo": client.get("codigo", ""),
            "cliente_nome": client.get("nome", ""),
            "data_venda": sale_date,
            "data_vencimento": due_date,
            "valor_venda_manual": 0.0,
            "estado_pagamento_manual": "",
            "obs": "",
            "created_at": self.desktop_main.now_iso(),
            "updated_at": self.desktop_main.now_iso(),
            "faturas": [],
            "pagamentos": [],
        }
        self.ensure_data().setdefault("faturacao", []).append(record)
        self._save(force=True)
        return record

    def billing_open_record(self, *, source_type: str = "", source_number: str = "", record_number: str = "") -> dict[str, Any]:
        reg_num = str(record_number or "").strip()
        if reg_num:
            return self.billing_detail(reg_num)
        record = self._billing_create_record_from_source(source_type, source_number)
        return self.billing_detail(str(record.get("numero", "") or "").strip())

    def billing_detail(self, numero: str) -> dict[str, Any]:
        record = self._billing_find_record(numero)
        if record is None:
            raise ValueError("Registo de faturação não encontrado.")
        quote_num = str(record.get("orcamento_numero", "") or "").strip()
        order_num = str(record.get("encomenda_numero", "") or "").strip()
        quote = self._billing_quote_by_number(quote_num) if quote_num else None
        order = self._billing_order_by_number(order_num) if order_num else None
        quote, order = self._billing_sync_record_source(record, quote=quote, order=order, persist=True)
        quote_num = str(record.get("orcamento_numero", "") or "").strip()
        order_num = str(record.get("encomenda_numero", "") or "").strip()
        if isinstance(order, dict):
            try:
                self.desktop_main.update_estado_expedicao_encomenda(order)
            except Exception:
                pass
        client = self._billing_client_info(quote=quote, order=order, record=record)
        client_fiscal = self._billing_client_snapshot(quote=quote, order=order, record=record)
        issuer = dict(getattr(self.desktop_main, "get_guia_emitente_info", lambda: {})() or {})
        guides = self._billing_guides_for_order(order_num)
        invoices = [dict(row) for row in list(record.get("faturas", []) or [])]
        payments = [dict(row) for row in list(record.get("pagamentos", []) or [])]
        invoice_map = {str(row.get("id", "") or "").strip(): row for row in invoices}
        payment_sums: dict[str, float] = {}
        effective_payments = self._billing_effective_payments(record)
        for pay in effective_payments:
            invoice_id = str(pay.get("fatura_id", "") or "").strip()
            if invoice_id:
                payment_sums[invoice_id] = round(payment_sums.get(invoice_id, 0.0) + self._parse_float(pay.get("valor", 0), 0), 2)
        for inv in invoices:
            inv_id = str(inv.get("id", "") or "").strip()
            paid_amount = round(payment_sums.get(inv_id, 0.0), 2)
            total_amount = round(self._parse_float(inv.get("valor_total", 0), 0), 2)
            inv["legal_invoice_no"] = str(inv.get("legal_invoice_no", "") or self._billing_legal_invoice_no(inv)).strip()
            inv["system_entry_date"] = str(inv.get("system_entry_date", "") or inv.get("created_at", "") or "").strip()
            inv["source_billing"] = self.tax_compliance.normalize_source_billing(inv.get("source_billing", ""), fallback="P")
            inv["hash_control"] = str(inv.get("hash_control", "") or self.tax_compliance.DEFAULT_HASH_CONTROL).strip()
            inv["communication_status"] = str(inv.get("communication_status", "") or "Por comunicar").strip() or "Por comunicar"
            if self._billing_invoice_is_void(inv):
                inv["anulada"] = True
                inv["estado"] = "Anulada"
                inv["recebido"] = 0.0
                inv["saldo"] = 0.0
                continue
            inv["recebido"] = paid_amount
            inv["saldo"] = round(max(0.0, total_amount - paid_amount), 2)
            if paid_amount >= (total_amount - 0.009):
                inv["estado"] = "Paga"
            elif paid_amount > 0:
                inv["estado"] = "Parcial"
            else:
                due_date = str(inv.get("data_vencimento", "") or "").strip()
                inv["estado"] = "Atrasada" if (due_date and due_date < self.desktop_main.now_iso()[:10]) else "Pendente"
        sold_total = self._billing_record_sale_total(record, quote)
        invoices_total = round(sum(self._parse_float(row.get("valor_total", 0), 0) for row in invoices if not self._billing_invoice_is_void(row)), 2)
        payments_total = round(sum(self._parse_float(row.get("valor", 0), 0) for row in effective_payments), 2)
        balance = round(max(0.0, invoices_total - payments_total), 2)
        last_invoice = ""
        if invoices:
            last_invoice = str(sorted(invoices, key=lambda row: (str(row.get("data_emissao", "") or ""), str(row.get("numero_fatura", "") or "")), reverse=True)[0].get("numero_fatura", "") or "").strip()
        fiscal_invoice = {}
        if invoices:
            fiscal_invoice = dict(
                sorted(
                    invoices,
                    key=lambda row: (
                        str(row.get("data_emissao", "") or ""),
                        str(row.get("numero_fatura", "") or ""),
                    ),
                    reverse=True,
                )[0]
            )
        return {
            "numero": str(record.get("numero", "") or "").strip(),
            "origem": str(record.get("origem", "") or "").strip(),
            "orcamento_numero": quote_num,
            "encomenda_numero": order_num,
            "cliente_codigo": client.get("codigo", ""),
            "cliente_nome": client.get("nome", ""),
            "cliente_label": client.get("label", "-") or "-",
            "cliente_nif": str(client_fiscal.get("nif", "") or "").strip(),
            "cliente_morada": str(client_fiscal.get("morada", "") or "").strip(),
            "cliente_contacto": str(client_fiscal.get("contacto", "") or "").strip(),
            "cliente_email": str(client_fiscal.get("email", "") or "").strip(),
            "emitente_nome": str(issuer.get("nome", "") or "").strip(),
            "emitente_nif": str(issuer.get("nif", "") or "").strip(),
            "emitente_morada": str(issuer.get("morada", "") or "").strip(),
            "data_venda": str(record.get("data_venda", "") or "").strip()[:10],
            "data_vencimento": str(record.get("data_vencimento", "") or "").strip()[:10],
            "valor_venda": sold_total,
            "valor_venda_manual": round(self._parse_float(record.get("valor_venda_manual", 0), 0), 2),
            "estado_faturacao": self._billing_invoice_status(record, quote),
            "estado_pagamento": self._billing_payment_status(record),
            "estado_pagamento_manual": str(record.get("estado_pagamento_manual", "") or "").strip(),
            "valor_faturado": invoices_total,
            "valor_recebido": payments_total,
            "saldo": balance,
            "por_faturar": round(max(0.0, sold_total - invoices_total), 2),
            "obs": str(record.get("obs", "") or "").strip(),
            "order_status": str((order or {}).get("estado", "") or "Sem encomenda").strip() or "Sem encomenda",
            "shipping_status": str((order or {}).get("estado_expedicao", "") or "Sem encomenda").strip() or "Sem encomenda",
            "quote_status": str((quote or {}).get("estado", "") or "").strip(),
            "guide_count": len(guides),
            "last_guide": str((guides[0] if guides else {}).get("numero", "") or "").strip(),
            "last_invoice": last_invoice,
            "fiscal_software_cert": self._billing_software_cert_number() or "Nao configurado",
            "fiscal_legal_invoice_no": str(fiscal_invoice.get("legal_invoice_no", "") or fiscal_invoice.get("numero_fatura", "") or "-").strip() or "-",
            "fiscal_source_billing": str(fiscal_invoice.get("source_billing", "") or "-").strip() or "-",
            "fiscal_system_entry_date": str(fiscal_invoice.get("system_entry_date", "") or "-").strip() or "-",
            "fiscal_hash_control": str(fiscal_invoice.get("hash_control", "") or "-").strip() or "-",
            "fiscal_hash": str(fiscal_invoice.get("hash", "") or "").strip(),
            "fiscal_communication_status": str(fiscal_invoice.get("communication_status", "") or "-").strip() or "-",
            "fiscal_communication_file": str(fiscal_invoice.get("communication_filename", "") or "").strip(),
            "guides": guides,
            "guide_options": [
                {
                    "numero": str(row.get("numero", "") or "").strip(),
                    "label": f"{str(row.get('numero', '') or '').strip()} | {str(row.get('data_emissao', '') or '').strip()[:10]}",
                }
                for row in guides
            ],
            "invoices": invoices,
            "invoice_options": [
                {
                    "id": str(row.get("id", "") or "").strip(),
                    "label": f"{str(row.get('numero_fatura', '') or row.get('id', '')).strip()} | {self._fmt(row.get('valor_total', 0))}",
                }
                for row in invoices
                if not self._billing_invoice_is_void(row)
            ],
            "payments": [
                {
                    **row,
                    "fatura_label": str((invoice_map.get(str(row.get('fatura_id', '') or '').strip()) or {}).get("numero_fatura", "") or "").strip(),
                }
                for row in payments
            ],
        }

    def billing_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        numero = str(payload.get("numero", "") or "").strip()
        record = self._billing_find_record(numero)
        if record is None:
            raise ValueError("Registo de faturação não encontrado.")
        record["data_venda"] = str(payload.get("data_venda", "") or record.get("data_venda", "") or "").strip()[:10]
        record["data_vencimento"] = str(payload.get("data_vencimento", "") or record.get("data_vencimento", "") or "").strip()[:10]
        record["valor_venda_manual"] = round(self._parse_float(payload.get("valor_venda_manual", record.get("valor_venda_manual", 0)), 0), 2)
        record["estado_pagamento_manual"] = str(payload.get("estado_pagamento_manual", record.get("estado_pagamento_manual", "") or "") or "").strip()
        record["obs"] = str(payload.get("obs", record.get("obs", "") or "") or "").strip()
        record["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return self.billing_detail(numero)

    def billing_remove(self, numero: str) -> None:
        reg_num = str(numero or "").strip()
        record = self._billing_find_record(reg_num)
        if record is None:
            raise ValueError("Registo de faturação não encontrado.")
        if list(record.get("faturas", []) or []) or list(record.get("pagamentos", []) or []):
            raise ValueError("Nao e possivel remover um registo de faturacao com faturas ou pagamentos. Use a anulacao dos documentos e mantenha o historico.")
        rows = list(self._billing_records())
        filtered = [row for row in rows if str((row or {}).get("numero", "") or "").strip() != reg_num]
        if len(filtered) == len(rows):
            raise ValueError("Registo de faturação não encontrado.")
        self.ensure_data()["faturacao"] = filtered
        self._save(force=True)

    def billing_add_invoice(self, numero: str, payload: dict[str, Any]) -> dict[str, Any]:
        reg_num = str(numero or "").strip()
        record = self._billing_find_record(reg_num)
        if record is None:
            raise ValueError("Registo de faturação não encontrado.")
        payload_dict = dict(payload or {})
        row_id = str(payload_dict.get("id", "") or "").strip()
        existing = next((row for row in list(record.get("faturas", []) or []) if str(row.get("id", "") or "").strip() == row_id), None) if row_id else None
        if existing is not None and self._billing_invoice_is_void(existing):
            raise ValueError("Nao e possivel editar uma fatura anulada.")
        invoice = self._billing_normalize_invoice(payload_dict, existing)
        if existing is not None and self._billing_invoice_locked(existing) and self._billing_invoice_core_fields(existing) != self._billing_invoice_core_fields(invoice):
            raise ValueError("Nao e possivel alterar os dados fiscais de uma fatura ja emitida.")
        document = self._billing_ensure_invoice_compliance(record, invoice, actor=self._billing_actor(), force_snapshot=existing is None)
        invoice["subtotal"] = round(self._parse_float(document.get("subtotal", invoice.get("subtotal", 0)), 0), 2)
        invoice["valor_iva"] = round(self._parse_float(document.get("valor_iva", invoice.get("valor_iva", 0)), 0), 2)
        invoice["valor_total"] = round(self._parse_float(document.get("valor_total", invoice.get("valor_total", 0)), 0), 2)
        if not invoice.get("iva_perc") and list(document.get("tax_summary", []) or []):
            invoice["iva_perc"] = round(self._parse_float((document.get("tax_summary", [{}])[0] or {}).get("rate", 23), 23), 2)
        if existing is None:
            record.setdefault("faturas", []).append(invoice)
        else:
            existing.update(invoice)
        record["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return self.billing_detail(reg_num)

    def billing_generate_invoice_pdf(self, numero: str, payload: dict[str, Any] | None = None) -> dict[str, Any]:
        reg_num = str(numero or "").strip()
        record = self._billing_find_record(reg_num)
        if record is None:
            raise ValueError("Registo de faturacao nao encontrado.")
        self._billing_sync_record_source(record, persist=True)
        payload_dict = dict(payload or {})
        row_id = str(payload_dict.get("id", "") or "").strip()
        existing = next((row for row in list(record.get("faturas", []) or []) if str(row.get("id", "") or "").strip() == row_id), None) if row_id else None
        if existing is not None and self._billing_invoice_is_void(existing):
            raise ValueError("Nao e possivel gerar PDF para uma fatura anulada.")
        if existing is None:
            raw_num = str(payload_dict.get("numero_fatura", "") or "").strip()
            raw_seq = int(self._parse_float(payload_dict.get("seq_num", 0), 0) or 0)
            manual_number = bool(raw_num) and raw_seq <= 0
            if not manual_number:
                payload_dict.update(
                    self._billing_next_invoice_identifiers(
                        issue_date=str(payload_dict.get("data_emissao", "") or self.desktop_main.now_iso())[:10],
                        serie_id=str(payload_dict.get("serie_id", "") or payload_dict.get("serie", "") or "").strip(),
                        validation_code_hint=str(payload_dict.get("at_validation_code", "") or "").strip(),
                        reserve=True,
                    )
                )
        invoice = self._billing_normalize_invoice(payload_dict, existing)
        if existing is not None and self._billing_invoice_locked(existing) and self._billing_invoice_core_fields(existing) != self._billing_invoice_core_fields(invoice):
            raise ValueError("Nao e possivel alterar os dados fiscais de uma fatura ja emitida.")
        document = self._billing_ensure_invoice_compliance(record, invoice, actor=self._billing_actor(), force_snapshot=existing is None)
        invoice["subtotal"] = round(self._parse_float(document.get("subtotal", 0), 0), 2)
        invoice["valor_iva"] = round(self._parse_float(document.get("valor_iva", 0), 0), 2)
        invoice["valor_total"] = round(self._parse_float(document.get("valor_total", invoice.get("valor_total", 0)), 0), 2)
        if not invoice.get("iva_perc") and list(document.get("tax_summary", []) or []):
            invoice["iva_perc"] = round(self._parse_float((document.get("tax_summary", [{}])[0] or {}).get("rate", 23), 23), 2)
        safe_number = "".join(ch if (ch.isalnum() or ch in ("-", "_")) else "_" for ch in str(invoice.get("numero_fatura", "") or "fatura").strip())
        output_hint = str(payload_dict.get("output_path", "") or payload_dict.get("caminho", "") or "").strip()
        output_path = Path(output_hint) if output_hint else self._storage_output_path("billing/invoices", f"{safe_number}.pdf")
        rendered = self.billing_pdf_actions.render_invoice_pdf(self, output_path, document)
        invoice["caminho"] = self._store_shared_file(rendered, "billing/invoices", preferred_name=f"{safe_number}.pdf")
        if existing is None:
            record.setdefault("faturas", []).append(invoice)
        else:
            existing.update(invoice)
        record["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return self.billing_detail(reg_num)

    def billing_remove_invoice(self, numero: str, invoice_id: str) -> dict[str, Any]:
        return self.billing_cancel_invoice(numero, invoice_id, "Anulada pelo utilizador.")

    def billing_cancel_invoice(self, numero: str, invoice_id: str, reason: str = "") -> dict[str, Any]:
        reg_num = str(numero or "").strip()
        row_id = str(invoice_id or "").strip()
        record = self._billing_find_record(reg_num)
        if record is None:
            raise ValueError("Registo de faturação não encontrado.")
        existing = next((row for row in list(record.get("faturas", []) or []) if str(row.get("id", "") or "").strip() == row_id), None)
        if existing is None:
            raise ValueError("Fatura não encontrada.")
        if self._billing_invoice_is_void(existing):
            raise ValueError("A fatura já está anulada.")
        linked_payments = [
            row
            for row in list(record.get("pagamentos", []) or [])
            if str(row.get("fatura_id", "") or "").strip() == row_id and self._parse_float(row.get("valor", 0), 0) > 0
        ]
        if linked_payments:
            raise ValueError("Nao e possivel anular uma fatura com pagamentos associados. Regulariza primeiro os pagamentos.")
        existing["anulada"] = True
        existing["estado"] = "Anulada"
        existing["anulada_motivo"] = str(reason or "").strip() or "Anulada pelo utilizador."
        existing["anulada_at"] = self.desktop_main.now_iso()
        existing["status_source_id"] = self._billing_actor()
        existing["communication_status"] = "Por comunicar"
        existing["communication_error"] = ""
        record["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return self.billing_detail(reg_num)

    def billing_add_payment(self, numero: str, payload: dict[str, Any]) -> dict[str, Any]:
        reg_num = str(numero or "").strip()
        record = self._billing_find_record(reg_num)
        if record is None:
            raise ValueError("Registo de faturação não encontrado.")
        payload_dict = dict(payload or {})
        row_id = str(payload_dict.get("id", "") or "").strip()
        invoice_id = str(payload_dict.get("fatura_id", "") or "").strip()
        if invoice_id and not any(str(row.get("id", "") or "").strip() == invoice_id for row in list(record.get("faturas", []) or [])):
            raise ValueError("A fatura associada ao pagamento não existe neste registo.")
        if invoice_id:
            invoice = next((row for row in list(record.get("faturas", []) or []) if str(row.get("id", "") or "").strip() == invoice_id), None)
            if self._billing_invoice_is_void(invoice):
                raise ValueError("Nao e possivel associar pagamentos a uma fatura anulada.")
        existing = next((row for row in list(record.get("pagamentos", []) or []) if str(row.get("id", "") or "").strip() == row_id), None) if row_id else None
        payment = self._billing_normalize_payment(payload_dict, existing)
        if existing is None:
            record.setdefault("pagamentos", []).append(payment)
        else:
            existing.update(payment)
        record["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return self.billing_detail(reg_num)

    def billing_remove_payment(self, numero: str, payment_id: str) -> dict[str, Any]:
        reg_num = str(numero or "").strip()
        row_id = str(payment_id or "").strip()
        record = self._billing_find_record(reg_num)
        if record is None:
            raise ValueError("Registo de faturação não encontrado.")
        before = len(list(record.get("pagamentos", []) or []))
        record["pagamentos"] = [row for row in list(record.get("pagamentos", []) or []) if str(row.get("id", "") or "").strip() != row_id]
        if len(record["pagamentos"]) == before:
            raise ValueError("Pagamento não encontrado.")
        record["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return self.billing_detail(reg_num)

    def _billing_export_invoice_rows(self, start_date: str = "", end_date: str = "") -> list[tuple[dict[str, Any], dict[str, Any], dict[str, Any]]]:
        start_txt = str(start_date or "").strip()[:10]
        end_txt = str(end_date or "").strip()[:10]
        changed = False
        rows: list[tuple[dict[str, Any], dict[str, Any], dict[str, Any]]] = []
        for record in self._billing_records():
            if not isinstance(record, dict):
                continue
            for invoice in list(record.get("faturas", []) or []):
                if not isinstance(invoice, dict):
                    continue
                issue_date = str(invoice.get("data_emissao", "") or "").strip()[:10]
                if start_txt and issue_date and issue_date < start_txt:
                    continue
                if end_txt and issue_date and issue_date > end_txt:
                    continue
                before = json.dumps(invoice, ensure_ascii=False, sort_keys=True, default=str)
                document = self._billing_ensure_invoice_compliance(record, invoice, actor=self._billing_actor(), force_snapshot=False)
                after = json.dumps(invoice, ensure_ascii=False, sort_keys=True, default=str)
                if before != after:
                    changed = True
                rows.append((record, invoice, document))
        if changed:
            self._save(force=True)
        return rows

    def _billing_export_payload_from_rows(
        self,
        export_rows: list[tuple[dict[str, Any], dict[str, Any], dict[str, Any]]],
        *,
        start_date: str = "",
        end_date: str = "",
    ) -> dict[str, Any]:
        issuer = dict(getattr(self.desktop_main, "get_guia_emitente_info", lambda: {})() or {})
        producer = self._billing_software_producer_info(issuer)
        customers_map: dict[str, dict[str, Any]] = {}
        products_map: dict[str, dict[str, Any]] = {}
        tax_map: dict[tuple[str, str, str], dict[str, Any]] = {}
        invoices_payload: list[dict[str, Any]] = []
        issue_dates = [str(invoice.get("data_emissao", "") or "").strip()[:10] for _, invoice, _ in export_rows if str(invoice.get("data_emissao", "") or "").strip()[:10]]
        computed_start = start_date[:10] if start_date else (min(issue_dates) if issue_dates else str(self.desktop_main.now_iso())[:10])
        computed_end = end_date[:10] if end_date else (max(issue_dates) if issue_dates else str(self.desktop_main.now_iso())[:10])
        for _record, invoice, document in export_rows:
            customer = dict(document.get("customer", {}) or {})
            customer_id = str(customer.get("codigo", "") or customer.get("nif", "") or "CONSUMIDOR-FINAL").strip() or "CONSUMIDOR-FINAL"
            customer_tax_id = str(customer.get("nif", "") or "999999990").strip() or "999999990"
            customers_map.setdefault(
                customer_id,
                {
                    "customer_id": customer_id,
                    "account_id": customer_id,
                    "tax_id": customer_tax_id,
                    "name": str(customer.get("nome", "") or "Cliente").strip() or "Cliente",
                    "address_detail": str(customer.get("morada", "") or "-").strip() or "-",
                    "city": "-",
                    "postal_code": "0000-000",
                    "country": "PT",
                },
            )
            line_payloads: list[dict[str, Any]] = []
            for line in list(document.get("lines", []) or []):
                product_code = str(line.get("reference", "") or line.get("ref_externa", "") or line.get("ref_interna", "") or "ITEM").strip() or "ITEM"
                product_type = self.tax_compliance.product_type_from_line(line)
                products_map.setdefault(
                    product_code,
                    {
                        "product_type": product_type,
                        "product_code": product_code,
                        "product_group": "SERVICOS" if product_type == "S" else "ARTIGOS",
                        "product_description": str(line.get("description", "") or product_code).strip() or product_code,
                        "product_number_code": product_code,
                    },
                )
                tax_rate = round(self._parse_float(line.get("iva_perc", 0), 0), 2)
                tax_code = self.tax_compliance.tax_code_from_rate(tax_rate)
                tax_map.setdefault(
                    ("IVA", "PT", tax_code),
                    {
                        "tax_type": "IVA",
                        "tax_country_region": "PT",
                        "tax_code": tax_code,
                        "tax_percentage": tax_rate,
                        "description": "IVA" if tax_rate > 0 else "Nao sujeito",
                    },
                )
                line_payloads.append(
                    {
                        "product_code": product_code,
                        "product_description": str(line.get("description", "") or product_code).strip() or product_code,
                        "quantity": round(self._parse_float(line.get("quantity", 0), 0), 3),
                        "unit_of_measure": str(line.get("unit", "") or "UN").strip() or "UN",
                        "unit_price": round(self._parse_float(line.get("unit_price", 0), 0), 2),
                        "description": str(line.get("description", "") or product_code).strip() or product_code,
                        "credit_amount": round(self._parse_float(line.get("subtotal", 0), 0), 2),
                        "tax_type": "IVA",
                        "tax_country_region": "PT",
                        "tax_code": tax_code,
                        "tax_percentage": tax_rate,
                    }
                )
            invoice_status_date = str(invoice.get("anulada_at", "") or invoice.get("system_entry_date", "") or invoice.get("created_at", "") or self.desktop_main.now_iso()).strip()
            invoices_payload.append(
                {
                    "invoice_no": str(invoice.get("legal_invoice_no", "") or self._billing_legal_invoice_no(invoice)).strip(),
                    "invoice_status": self.tax_compliance.invoice_status_code(is_void=self._billing_invoice_is_void(invoice)),
                    "invoice_status_date": invoice_status_date,
                    "status_source_id": self._billing_status_source_id(invoice),
                    "source_billing": self.tax_compliance.normalize_source_billing(invoice.get("source_billing", ""), fallback="P"),
                    "hash": self._billing_saft_hash_value(invoice),
                    "hash_control": self._billing_saft_hash_control(invoice),
                    "period": int((str(invoice.get("data_emissao", "") or "")[5:7] or "0")) if str(invoice.get("data_emissao", "") or "").strip()[:7] else 0,
                    "invoice_date": str(invoice.get("data_emissao", "") or "").strip()[:10],
                    "invoice_type": str(invoice.get("doc_type", "") or "FT").strip() or "FT",
                    "source_id": str(invoice.get("source_id", "") or self._billing_actor()).strip() or "Sistema",
                    "system_entry_date": str(invoice.get("system_entry_date", "") or invoice.get("created_at", "") or self.desktop_main.now_iso()).strip(),
                    "customer_id": customer_id,
                    "lines": line_payloads,
                    "tax_payable": round(self._parse_float(document.get("valor_iva", 0), 0), 2),
                    "net_total": round(self._parse_float(document.get("subtotal", 0), 0), 2),
                    "gross_total": round(self._parse_float(document.get("valor_total", 0), 0), 2),
                }
            )
        return {
            "header": {
                "audit_file_version": self.tax_compliance.SAFT_PT_AUDIT_FILE_VERSION,
                "company_id": str(issuer.get("nif", "") or producer.get("producer_nif", "999999990")).strip() or "999999990",
                "tax_registration_number": str(issuer.get("nif", "") or "999999990").strip() or "999999990",
                "tax_accounting_basis": "F",
                "company_name": str(issuer.get("nome", "") or "LuGEST").strip() or "LuGEST",
                "business_name": str(issuer.get("nome", "") or "LuGEST").strip() or "LuGEST",
                "company_address_detail": str(issuer.get("morada", "") or "-").strip() or "-",
                "company_city": "-",
                "company_postal_code": "0000-000",
                "company_country": "PT",
                "fiscal_year": (computed_start[:4] if len(computed_start) >= 4 else str(self.desktop_main.datetime.now().year)),
                "start_date": computed_start,
                "end_date": computed_end,
                "currency_code": "EUR",
                "date_created": str(self.desktop_main.now_iso())[:10],
                "tax_entity": "Global",
                "product_company_tax_id": producer.get("producer_nif", "999999990"),
                "software_certificate_number": self._billing_software_cert_number() or "0",
                "product_id": producer.get("product_id", self.tax_compliance.DEFAULT_PRODUCT_ID),
                "product_version": producer.get("product_version", self.tax_compliance.DEFAULT_PRODUCT_VERSION),
                "header_comment": "Exportação interna LuGEST SAF-T(PT) - faturação.",
            },
            "customers": list(customers_map.values()),
            "products": list(products_map.values()),
            "tax_table": list(tax_map.values()),
            "invoices": invoices_payload,
        }

    def _billing_prepare_at_payload(
        self,
        export_rows: list[tuple[dict[str, Any], dict[str, Any], dict[str, Any]]],
    ) -> tuple[dict[str, Any], list[tuple[dict[str, Any], dict[str, Any], dict[str, Any]]], str]:
        pending: list[tuple[dict[str, Any], dict[str, Any], dict[str, Any]]] = []
        for record, invoice, document in export_rows:
            if self.tax_compliance.normalize_source_billing(invoice.get("source_billing", ""), fallback="P") != "P":
                continue
            status = str(invoice.get("communication_status", "") or "").strip().lower()
            if status == "comunicada":
                continue
            pending.append((record, invoice, document))
        if not pending:
            raise ValueError("Nao existem faturas pendentes para preparar comunicacao AT.")
        issuer = dict(getattr(self.desktop_main, "get_guia_emitente_info", lambda: {})() or {})
        producer = self._billing_software_producer_info(issuer)
        batch_id = self.desktop_main.uuid.uuid4().hex[:12].upper()
        payload = {
            "header": {
                "generated_at": self.desktop_main.now_iso(),
                "issuer_name": str(issuer.get("nome", "") or "LuGEST").strip() or "LuGEST",
                "issuer_tax_id": str(issuer.get("nif", "") or "999999990").strip() or "999999990",
                "software_certificate_number": self._billing_software_cert_number() or "0",
                "product_id": producer.get("product_id", self.tax_compliance.DEFAULT_PRODUCT_ID),
                "product_version": producer.get("product_version", self.tax_compliance.DEFAULT_PRODUCT_VERSION),
                "preparation_mode": "manual",
            },
            "documents": [
                {
                    "document_id": str(invoice.get("id", "") or "").strip(),
                    "invoice_no": str(invoice.get("legal_invoice_no", "") or self._billing_legal_invoice_no(invoice)).strip(),
                    "invoice_date": str(invoice.get("data_emissao", "") or "").strip()[:10],
                    "invoice_type": str(invoice.get("doc_type", "") or "FT").strip() or "FT",
                    "atcud": str(invoice.get("atcud", "") or "").strip(),
                    "hash": self._billing_saft_hash_value(invoice),
                    "hash_control": self._billing_saft_hash_control(invoice),
                    "customer_tax_id": str((document.get("customer", {}) or {}).get("nif", "") or "999999990").strip() or "999999990",
                    "gross_total": round(self._parse_float(document.get("valor_total", 0), 0), 2),
                    "status": str(invoice.get("communication_status", "") or "Por comunicar").strip() or "Por comunicar",
                    "source_billing": self.tax_compliance.normalize_source_billing(invoice.get("source_billing", ""), fallback="P"),
                    "system_entry_date": str(invoice.get("system_entry_date", "") or invoice.get("created_at", "") or self.desktop_main.now_iso()).strip(),
                }
                for _, invoice, document in pending
            ],
        }
        return payload, pending, batch_id

    def _billing_record_export_rows(self, numero: str) -> list[tuple[dict[str, Any], dict[str, Any], dict[str, Any]]]:
        reg_num = str(numero or "").strip()
        record = self._billing_find_record(reg_num)
        if record is None:
            raise ValueError("Registo de faturação não encontrado.")
        changed = False
        rows: list[tuple[dict[str, Any], dict[str, Any], dict[str, Any]]] = []
        for invoice in list(record.get("faturas", []) or []):
            if not isinstance(invoice, dict):
                continue
            before = json.dumps(invoice, ensure_ascii=False, sort_keys=True, default=str)
            document = self._billing_ensure_invoice_compliance(record, invoice, actor=self._billing_actor(), force_snapshot=False)
            after = json.dumps(invoice, ensure_ascii=False, sort_keys=True, default=str)
            if before != after:
                changed = True
            rows.append((record, invoice, document))
        if not rows:
            raise ValueError("Este registo ainda não tem faturas para exportar.")
        if changed:
            self._save(force=True)
        return rows

    def billing_export_saft_pt(self, start_date: str = "", end_date: str = "", output_path: str = "") -> str:
        export_rows = self._billing_export_invoice_rows(start_date, end_date)
        if not export_rows:
            raise ValueError("Nao existem faturas no intervalo indicado para exportar SAF-T(PT).")
        output_target = Path(output_path) if str(output_path or "").strip() else (self.base_dir / "generated" / "compliance" / "saft" / f"saft_pt_{self.desktop_main.datetime.now().strftime('%Y%m%d_%H%M%S')}.xml")
        rendered = self.tax_compliance.render_saft_pt_xml(self._billing_export_payload_from_rows(export_rows, start_date=start_date, end_date=end_date), output_target)
        return str(rendered)

    def billing_prepare_at_communication_batch(self, start_date: str = "", end_date: str = "", output_path: str = "") -> str:
        export_rows = self._billing_export_invoice_rows(start_date, end_date)
        payload, pending, batch_id = self._billing_prepare_at_payload(export_rows)
        output_target = Path(output_path) if str(output_path or "").strip() else (self.base_dir / "generated" / "compliance" / "at" / f"at_preparacao_{batch_id}.xml")
        rendered = self.tax_compliance.render_at_communication_preparation_xml(payload, output_target)
        for record, invoice, _document in pending:
            invoice["communication_status"] = "Preparada"
            invoice["communication_filename"] = str(rendered)
            invoice["communication_batch_id"] = batch_id
            invoice["communication_error"] = ""
            record["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return str(rendered)

    def billing_export_record_saft_pt(self, numero: str, output_path: str = "") -> str:
        export_rows = self._billing_record_export_rows(numero)
        issue_dates = [str(invoice.get("data_emissao", "") or "").strip()[:10] for _, invoice, _ in export_rows if str(invoice.get("data_emissao", "") or "").strip()[:10]]
        output_target = Path(output_path) if str(output_path or "").strip() else self._storage_output_path("billing/compliance", f"saft_pt_{str(numero or '').strip() or 'registo'}.xml")
        rendered = self.tax_compliance.render_saft_pt_xml(
            self._billing_export_payload_from_rows(
                export_rows,
                start_date=min(issue_dates) if issue_dates else "",
                end_date=max(issue_dates) if issue_dates else "",
            ),
            output_target,
        )
        return str(self._store_shared_file(rendered, "billing/compliance", preferred_name=Path(str(rendered)).name))

    def billing_prepare_record_at_communication_batch(self, numero: str, output_path: str = "") -> str:
        export_rows = self._billing_record_export_rows(numero)
        payload, pending, batch_id = self._billing_prepare_at_payload(export_rows)
        output_target = Path(output_path) if str(output_path or "").strip() else self._storage_output_path("billing/compliance", f"at_preparacao_{str(numero or '').strip() or batch_id}.xml")
        rendered = self.tax_compliance.render_at_communication_preparation_xml(payload, output_target)
        rendered_ref = self._store_shared_file(rendered, "billing/compliance", preferred_name=Path(str(rendered)).name)
        for record, invoice, _document in pending:
            invoice["communication_status"] = "Preparada"
            invoice["communication_filename"] = rendered_ref
            invoice["communication_batch_id"] = batch_id
            invoice["communication_error"] = ""
            record["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return rendered_ref

    def billing_open_path(self, path: str) -> Path:
        return self.open_file_reference(path)

