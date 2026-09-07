from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from typing import Any


class PurchasingDocumentsBackendMixin:
    """Legacy adapter for purchasing documents; see BACKEND_GUIDE.md."""

    def _sync_ne_from_materia(self) -> None:
        data = self.ensure_data()
        holder = SimpleNamespace(data=data)
        changed = False
        for ne in data.get("notas_encomenda", []):
            if self.ne_expedicao_actions._sync_ne_linhas_with_materia(holder, ne):
                changed = True
        if changed:
            self._save(force=True)

    def _sync_ne_from_products(self) -> None:
        data = self.ensure_data()
        changed = False
        for note in list(data.get("notas_encomenda", []) or []):
            if self._sync_note_lines_with_products(note):
                self._recalculate_note_totals(note)
                changed = True
        if changed:
            self._save(force=True)

    def _sync_note_lines_with_products(self, note: dict[str, Any]) -> bool:
        changed = False
        product_map = {str(row.get("codigo", "") or "").strip(): row for row in list(self.ensure_data().get("produtos", []) or [])}
        for line in list(note.get("linhas", []) or []):
            if self.desktop_main.origem_is_materia(line.get("origem", "Produto")):
                continue
            product = product_map.get(str(line.get("ref", "") or "").strip())
            if not product:
                continue
            new_price = round(self._parse_float(self.desktop_main.produto_preco_unitario(product), 0), 6)
            old_price = self._parse_float(line.get("preco", 0), 0)
            if abs(new_price - old_price) > 1e-9:
                line["preco"] = new_price
                qty = self._parse_float(line.get("qtd", 0), 0)
                discount = max(0.0, min(100.0, self._parse_float(line.get("desconto", 0), 0)))
                iva = max(0.0, min(100.0, self._parse_float(line.get("iva", 23), 23)))
                base = (qty * new_price) * (1.0 - (discount / 100.0))
                line["total"] = round(base + (base * iva / 100.0), 4)
                changed = True
            new_desc = str(product.get("descricao", "") or "").strip()
            if new_desc and str(line.get("descricao", "") or "").strip() != new_desc:
                line["descricao"] = new_desc
                changed = True
        return changed

    def _update_materia_preco_from_unit(self, materia_id: str, preco_unit: Any) -> bool:
        material = self.material_by_id(str(materia_id or "").strip())
        if material is None:
            return False
        price_line = self._parse_float(preco_unit, 0)
        old = self._parse_float(material.get("p_compra", 0), 0)
        new_value = old
        formato = str(material.get("formato") or self.desktop_main.detect_materia_formato(material) or "").strip()
        if formato == "Tubo":
            metros = self._parse_float(material.get("metros", 0), 0)
            if metros > 0:
                new_value = round(price_line / metros, 6)
        elif formato in ("Chapa", "Perfil"):
            peso = self._parse_float(material.get("peso_unid", 0), 0)
            if peso > 0:
                new_value = round(price_line / peso, 6)
        else:
            new_value = round(price_line, 6)
        if abs(new_value - old) <= 1e-9:
            return False
        material["p_compra"] = new_value
        material["atualizado_em"] = self.desktop_main.now_iso()
        refresh_model = getattr(self, "_conjunto_refresh_model_prices", None)
        if callable(refresh_model):
            for model in list(self.ensure_data().get("conjuntos", []) or []):
                if isinstance(model, dict):
                    refresh_model(model)
        return True

    def _update_produto_preco_from_unit(self, produto_codigo: str, preco_unit: Any) -> bool:
        code = str(produto_codigo or "").strip()
        product = next((row for row in list(self.ensure_data().get("produtos", []) or []) if str(row.get("codigo", "") or "").strip() == code), None)
        if product is None:
            return False
        price_line = self._parse_float(preco_unit, 0)
        old = self._parse_float(product.get("p_compra", 0), 0)
        new_value = old
        modo = self.desktop_main.produto_modo_preco(product.get("categoria", ""), product.get("tipo", ""))
        if modo == "peso":
            peso = self._parse_float(product.get("peso_unid", 0), 0)
            if peso > 0:
                new_value = round(price_line / peso, 6)
        elif modo == "metros":
            metros = self._parse_float(product.get("metros_unidade", product.get("metros", 0)), 0)
            if metros > 0:
                new_value = round(price_line / metros, 6)
        else:
            new_value = round(price_line, 6)
        if abs(new_value - old) <= 1e-9:
            return False
        product["p_compra"] = new_value
        product["atualizado_em"] = self.desktop_main.now_iso()
        refresh_model = getattr(self, "_conjunto_refresh_model_prices", None)
        if callable(refresh_model):
            for model in list(self.ensure_data().get("conjuntos", []) or []):
                if isinstance(model, dict):
                    refresh_model(model)
        return True

    def _resolve_supplier(self, raw_value: str) -> tuple[str, str, str]:
        raw = str(raw_value or "").strip()
        if not raw:
            return "", "", ""
        supplier_id = ""
        supplier_name = raw
        if " - " in raw:
            supplier_id, supplier_name = [part.strip() for part in raw.split(" - ", 1)]
        supplier = None
        if supplier_id:
            supplier = next((row for row in list(self.ensure_data().get("fornecedores", []) or []) if str(row.get("id", "") or "").strip() == supplier_id), None)
        if supplier is None and supplier_name:
            supplier = next((row for row in list(self.ensure_data().get("fornecedores", []) or []) if str(row.get("nome", "") or "").strip().lower() == supplier_name.lower()), None)
        if supplier is None:
            return supplier_id, raw, ""
        supplier_id = str(supplier.get("id", "") or "").strip()
        supplier_name = str(supplier.get("nome", "") or "").strip()
        return supplier_id, supplier_name, str(supplier.get("contacto", "") or "").strip()

    def _normalize_supplier_reference(self, supplier_id_value: Any, supplier_text_value: Any) -> tuple[str, str, str]:
        supplier_id_raw = str(supplier_id_value or "").strip()
        supplier_text_raw = str(supplier_text_value or "").strip()
        candidates: list[str] = []
        if supplier_id_raw and supplier_text_raw:
            combined = supplier_text_raw if " - " in supplier_text_raw else f"{supplier_id_raw} - {supplier_text_raw}"
            candidates.append(combined.strip())
        if supplier_id_raw:
            candidates.append(supplier_id_raw)
        if supplier_text_raw:
            candidates.append(supplier_text_raw)
        seen: set[str] = set()
        for candidate in candidates:
            key = candidate.lower()
            if not candidate or key in seen:
                continue
            seen.add(key)
            resolved_id, resolved_text, resolved_contact = self._resolve_supplier(candidate)
            if resolved_id:
                return resolved_id, resolved_text, resolved_contact
        return "", supplier_text_raw, ""

    def _recalculate_note_totals(self, note: dict[str, Any]) -> None:
        note["total"] = round(sum(self._parse_float(line.get("total", 0), 0) for line in list(note.get("linhas", []) or [])), 2)

    def _note_kind(self, note: dict[str, Any]) -> str:
        if str(note.get("origem_cotacao", "") or "").strip():
            return "supplier_order"
        suppliers = {
            str(line.get("fornecedor_linha", "") or "").strip().lower()
            for line in list(note.get("linhas", []) or [])
            if str(line.get("fornecedor_linha", "") or "").strip()
        }
        if len(suppliers) > 1:
            return "rfq"
        if len(suppliers) == 1 and not str(note.get("fornecedor", "") or "").strip():
            return "rfq"
        if list(note.get("ne_geradas", []) or []):
            return "rfq"
        return "purchase_note"

    def _ne_document_type(self, payload: dict[str, Any] | None, fallback: str = "") -> str:
        raw = str((payload or {}).get("tipo", "") or fallback or "").strip().upper()
        raw = raw.replace("+", "_").replace("-", "_").replace(" ", "_")
        normalized = "".join(ch for ch in raw if ch.isalnum() or ch == "_").strip("_")
        if normalized in {"GUIA", "FATURA", "GUIA_FATURA", "ENTREGA", "OUTRO", "DOCUMENTO"}:
            return "DOCUMENTO" if normalized == "OUTRO" else normalized
        guia = str((payload or {}).get("guia", "") or "").strip()
        fatura = str((payload or {}).get("fatura", "") or "").strip()
        if guia and fatura:
            return "GUIA_FATURA"
        if fatura:
            return "FATURA"
        if guia:
            return "GUIA"
        return normalized or "DOCUMENTO"

    def _ne_document_type_label(self, doc_type: str) -> str:
        return {
            "ENTREGA": "Entrega",
            "GUIA": "Guia",
            "FATURA": "Fatura",
            "GUIA_FATURA": "Guia + Fatura",
            "DOCUMENTO": "Documento",
        }.get(str(doc_type or "").strip().upper(), "Documento")

    def _ne_document_title(self, payload: dict[str, Any] | None, doc_type: str = "") -> str:
        raw = payload or {}
        explicit = str(raw.get("titulo", "") or "").strip()
        if explicit:
            return explicit
        guia = str(raw.get("guia", "") or "").strip()
        fatura = str(raw.get("fatura", "") or "").strip()
        caminho = str(raw.get("caminho", "") or "").strip()
        data_documento = str(raw.get("data_documento", "") or "").strip()
        data_entrega = str(raw.get("data_entrega", "") or "").strip()
        resolved_type = self._ne_document_type(raw, fallback=doc_type)
        if guia and fatura:
            return f"Guia {guia} / Fatura {fatura}"
        if fatura:
            return f"Fatura {fatura}"
        if guia:
            return f"Guia {guia}"
        if resolved_type == "ENTREGA":
            date_txt = data_entrega or data_documento
            return f"Entrega {date_txt}".strip()
        if caminho:
            try:
                return Path(caminho).name or self._ne_document_type_label(resolved_type)
            except Exception:
                return self._ne_document_type_label(resolved_type)
        return self._ne_document_type_label(resolved_type)

    def _ne_document_signature(self, payload: dict[str, Any]) -> tuple[str, str, str, str, str, str, str]:
        return (
            str(payload.get("data_registo", "") or "").strip()[:19],
            str(payload.get("tipo", "") or "").strip().upper(),
            str(payload.get("guia", "") or "").strip(),
            str(payload.get("fatura", "") or "").strip(),
            str(payload.get("data_entrega", "") or "").strip()[:10],
            str(payload.get("data_documento", "") or "").strip()[:10],
            str(payload.get("obs", "") or "").strip(),
        )

    def _ne_normalize_document(self, payload: dict[str, Any] | None, default_type: str = "DOCUMENTO") -> dict[str, Any]:
        raw = dict(payload or {})
        doc_type = self._ne_document_type(raw, fallback=default_type)
        data_registo = str(raw.get("data_registo", "") or "").strip() or self.desktop_main.now_iso()
        normalized = {
            "data_registo": data_registo,
            "tipo": doc_type,
            "titulo": self._ne_document_title(raw, doc_type=doc_type),
            "caminho": str(raw.get("caminho", "") or "").strip(),
            "guia": str(raw.get("guia", "") or "").strip(),
            "fatura": str(raw.get("fatura", "") or "").strip(),
            "data_entrega": str(raw.get("data_entrega", "") or "").strip()[:10],
            "data_documento": str(raw.get("data_documento", "") or "").strip()[:10],
            "obs": str(raw.get("obs", "") or "").strip(),
        }
        normalized["tipo_label"] = self._ne_document_type_label(doc_type)
        normalized["has_path"] = bool(normalized["caminho"])
        return normalized

    def _ne_validate_document_payload(
        self,
        payload: dict[str, Any] | None,
        *,
        allow_delivery: bool = False,
    ) -> dict[str, Any]:
        """Validate a new purchasing document without rewriting legacy records."""
        raw = dict(payload or {})
        doc_type = self._ne_document_type(raw, fallback="ENTREGA" if allow_delivery else "DOCUMENTO")
        guia = str(raw.get("guia", "") or "").strip()
        fatura = str(raw.get("fatura", "") or "").strip()
        title = str(raw.get("titulo", "") or "").strip()
        path = str(raw.get("caminho", "") or "").strip()
        obs = str(raw.get("obs", "") or "").strip()

        if doc_type == "GUIA":
            if not guia:
                raise ValueError("Indica o número da guia.")
            if fatura:
                raise ValueError("O tipo Guia não pode guardar um número de fatura. Seleciona 'Guia + Fatura'.")
        elif doc_type == "FATURA":
            if not fatura:
                raise ValueError("Indica o número da fatura.")
        elif doc_type == "GUIA_FATURA":
            if not guia or not fatura:
                raise ValueError("Indica os números da guia e da fatura.")
        elif doc_type == "ENTREGA":
            if not allow_delivery:
                raise ValueError("O tipo Entrega só pode ser criado durante uma receção.")
            if not guia and not fatura:
                raise ValueError("A receção tem de ficar associada a uma guia ou a uma fatura.")
            doc_type = "GUIA_FATURA" if guia and fatura else ("FATURA" if fatura else "GUIA")
        elif doc_type == "DOCUMENTO":
            if not title and not path:
                raise ValueError("Num documento geral indica o título ou seleciona um ficheiro.")
            if guia or fatura:
                raise ValueError("Para registar números de guia ou fatura seleciona o tipo documental correto.")

        if not any((title, guia, fatura, path, obs)):
            raise ValueError("Indica os dados ou o ficheiro do documento.")
        raw["tipo"] = doc_type
        raw["guia"] = guia
        raw["fatura"] = fatura
        return raw

    def _ne_assert_document_unique(self, note: dict[str, Any], document: dict[str, Any]) -> None:
        doc_type = str(document.get("tipo", "") or "").strip().upper()
        guia = str(document.get("guia", "") or "").strip().casefold()
        fatura = str(document.get("fatura", "") or "").strip().casefold()
        for existing in self._ne_document_rows(note):
            existing_type = str(existing.get("tipo", "") or "").strip().upper()
            existing_guia = str(existing.get("guia", "") or "").strip().casefold()
            existing_fatura = str(existing.get("fatura", "") or "").strip().casefold()
            if fatura and existing_fatura == fatura:
                raise ValueError(f"A fatura {document.get('fatura')} já está registada nesta nota.")
            if doc_type in {"GUIA", "GUIA_FATURA"} and guia and existing_guia == guia and existing_type in {"GUIA", "GUIA_FATURA"}:
                raise ValueError(f"A guia {document.get('guia')} já está registada nesta nota.")

    def _ne_document_rows(self, note: dict[str, Any]) -> list[dict[str, Any]]:
        docs: list[dict[str, Any]] = []
        seen: set[tuple[str, str, str, str, str, str, str]] = set()
        for raw_index, raw_doc in enumerate(list(note.get("documentos", []) or [])):
            if not isinstance(raw_doc, dict):
                continue
            doc = self._ne_normalize_document(raw_doc)
            doc["source"] = "documento"
            doc["source_index"] = raw_index
            signature = self._ne_document_signature(doc)
            seen.add(signature)
            docs.append(doc)
        for raw_index, entrega in enumerate(list(note.get("entregas", []) or [])):
            if not isinstance(entrega, dict):
                continue
            doc = self._ne_normalize_document(
                {
                    "data_registo": entrega.get("data_registo"),
                    "tipo": entrega.get("tipo") or "ENTREGA",
                    "titulo": entrega.get("titulo", ""),
                    "caminho": entrega.get("caminho", ""),
                    "guia": entrega.get("guia", ""),
                    "fatura": entrega.get("fatura", ""),
                    "data_entrega": entrega.get("data_entrega", ""),
                    "data_documento": entrega.get("data_documento", ""),
                    "obs": entrega.get("obs", ""),
                },
                default_type="ENTREGA",
            )
            signature = self._ne_document_signature(doc)
            if signature in seen:
                continue
            doc["source"] = "entrega"
            doc["source_index"] = raw_index
            doc["derived"] = True
            docs.append(doc)
        docs.sort(
            key=lambda item: (
                str(item.get("data_registo", "") or ""),
                str(item.get("data_documento", "") or ""),
                str(item.get("data_entrega", "") or ""),
                str(item.get("titulo", "") or ""),
            ),
            reverse=True,
        )
        for index, doc in enumerate(docs):
            doc["index"] = index
        return docs
