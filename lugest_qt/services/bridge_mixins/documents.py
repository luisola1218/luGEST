from __future__ import annotations

import hashlib
import os
import re
from lugest_infra.storage import files as lugest_storage
from pathlib import Path
from typing import Any


class DocumentsBackendMixin:
    """Legacy adapter for documents; see BACKEND_GUIDE.md."""

    def _file_reference_name(self, raw: Any, fallback: str = "ficheiro") -> str:
        current = str(raw or "").strip()
        if not current:
            return fallback
        resolved = lugest_storage.resolve_file_reference(current, base_dir=self.base_dir)
        if resolved is not None:
            name = str(resolved.name or "").strip()
            if name:
                return name
        try:
            return Path(current).name or fallback
        except Exception:
            return fallback

    def _storage_output_path(self, category: str, filename: str) -> Path:
        return lugest_storage.allocate_storage_output_path(category, filename, base_dir=self.base_dir)

    def _store_shared_file(self, raw: Any, category: str, preferred_name: str = "") -> str:
        return lugest_storage.import_file_to_storage(
            raw,
            category,
            base_dir=self.base_dir,
            preferred_name=preferred_name,
        )

    def _resolve_file_reference(self, raw: Any) -> Path | None:
        return lugest_storage.resolve_file_reference(raw, base_dir=self.base_dir)

    def _document_reference_key(self, raw: Any, digest_cache: dict[str, str] | None = None) -> str:
        text = str(raw or "").strip()
        if not text:
            return ""
        resolved = self._resolve_file_reference(text)
        path_obj = resolved or Path(text)
        try:
            normalized = str(path_obj.resolve())
        except Exception:
            normalized = str(path_obj)
        if path_obj.exists() and path_obj.is_file():
            digest = ""
            if digest_cache is not None:
                digest = str(digest_cache.get(normalized, "") or "")
            if not digest:
                hasher = hashlib.sha1()
                try:
                    with path_obj.open("rb") as handle:
                        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                            if not chunk:
                                break
                            hasher.update(chunk)
                    digest = hasher.hexdigest()
                except Exception:
                    digest = ""
                if digest_cache is not None:
                    digest_cache[normalized] = digest
            if digest:
                try:
                    stat = path_obj.stat()
                    return f"sha1:{stat.st_size}:{digest}"
                except Exception:
                    return f"sha1:{digest}"
        return f"path:{normalized.casefold()}"

    def _document_reference_priority(self, raw: Any, preferred_category: str = "technical_pdfs") -> tuple[int, int, int, str]:
        text = str(raw or "").strip()
        if not text:
            return (99, 99, 99, "")
        normalized = text.replace("\\", "/").casefold()
        preferred = preferred_category.replace("\\", "/").strip("/").casefold()
        is_preferred = f"/{preferred}/" in f"/{normalized}/"
        exists_score = 0 if (self._resolve_file_reference(text) or Path(text)).exists() else 1
        suffix_match = re.search(r"_(\d+)(?=\.pdf$)", Path(text).name, re.IGNORECASE)
        suffix_value = int(suffix_match.group(1)) if suffix_match else 0
        return (0 if is_preferred else 1, exists_score, suffix_value, normalized)

    def _dedupe_document_references(
        self,
        values: list[Any],
        preferred_category: str = "technical_pdfs",
        digest_cache: dict[str, str] | None = None,
    ) -> list[str]:
        chosen: dict[str, str] = {}
        order: list[str] = []
        for item in list(values or []):
            item_txt = str(item or "").strip()
            if not item_txt:
                continue
            key = self._document_reference_key(item_txt, digest_cache=digest_cache)
            if not key:
                continue
            if key not in order:
                order.append(key)
                chosen[key] = item_txt
                continue
            current = chosen.get(key, "")
            if self._document_reference_priority(item_txt, preferred_category) < self._document_reference_priority(current, preferred_category):
                chosen[key] = item_txt
        return [chosen[key] for key in order if str(chosen.get(key, "") or "").strip()]

    def _piece_pdf_references(
        self,
        row: dict[str, Any],
        preferred_category: str = "technical_pdfs",
        digest_cache: dict[str, str] | None = None,
    ) -> list[str]:
        if not isinstance(row, dict):
            return []
        return self._dedupe_document_references(
            [
                row.get("desenho_pdf", ""),
                *list(row.get("desenhos_pdf", []) or []),
                *[
                    item
                    for item in list(row.get("ficheiros", []) or [])
                    if str(item or "").strip().lower().endswith(".pdf")
                ],
            ],
            preferred_category=preferred_category,
            digest_cache=digest_cache,
        )

    def _apply_piece_pdf_references(
        self,
        row: dict[str, Any],
        docs: list[Any],
        preferred_category: str = "technical_pdfs",
        digest_cache: dict[str, str] | None = None,
    ) -> bool:
        if not isinstance(row, dict):
            return False
        canonical_docs = self._dedupe_document_references(
            list(docs or []),
            preferred_category=preferred_category,
            digest_cache=digest_cache,
        )
        current_main = str(row.get("desenho_pdf", "") or "").strip()
        current_list = [str(item or "").strip() for item in list(row.get("desenhos_pdf", []) or []) if str(item or "").strip()]
        current_files = [str(item or "").strip() for item in list(row.get("ficheiros", []) or []) if str(item or "").strip()]
        non_pdf_files = [item for item in current_files if not item.lower().endswith(".pdf")]
        merged_files = non_pdf_files + canonical_docs
        deduped_files: list[str] = []
        for item in merged_files:
            if item not in deduped_files:
                deduped_files.append(item)
        changed = False
        new_main = canonical_docs[0] if canonical_docs else ""
        if current_main != new_main:
            row["desenho_pdf"] = new_main
            changed = True
        if current_list != canonical_docs:
            row["desenhos_pdf"] = canonical_docs
            changed = True
        if current_files != deduped_files:
            row["ficheiros"] = deduped_files
            changed = True
        return changed

    def open_file_reference(self, raw: str) -> Path:
        target = self._resolve_file_reference(raw)
        if target is None:
            raise ValueError("Ficheiro não indicado.")
        target = target.resolve()
        if not target.exists():
            raise ValueError(f"Ficheiro não encontrado: {target}")
        os.startfile(str(target))
        return target

    def _normalize_storage_paths_for_save(self, changed_keys: list[str] | None = None) -> None:
        data = self.ensure_data()
        digest_cache: dict[str, str] = {}
        selected_keys = {
            str(key or "")
            for key in list(changed_keys or [])
            if str(key or "") and not str(key or "").startswith("__")
        }

        def normalize_drawings(node: Any) -> None:
            if isinstance(node, dict):
                pdf_cache: dict[str, str] = {}

                def store_pdf(value: Any) -> str:
                    raw = str(value or "").strip()
                    if not raw:
                        return ""
                    if raw in pdf_cache:
                        return pdf_cache[raw]
                    pdf_cache[raw] = self._store_shared_file(
                        raw,
                        "technical_pdfs",
                        preferred_name=self._file_reference_name(raw, "desenho.pdf"),
                    )
                    return pdf_cache[raw]

                def dedupe(values: list[str]) -> list[str]:
                    clean: list[str] = []
                    for item in values:
                        item_txt = str(item or "").strip()
                        if item_txt and item_txt not in clean:
                            clean.append(item_txt)
                    return clean

                for key, value in list(node.items()):
                    if key in {"desenho", "desenho_path"}:
                        node[key] = self._store_shared_file(
                            value,
                            "drawings",
                            preferred_name=self._file_reference_name(value, "desenho"),
                        )
                    elif key == "desenho_pdf":
                        node[key] = store_pdf(value)
                    elif key == "desenhos_pdf" and isinstance(value, list):
                        node[key] = dedupe(
                            [
                                store_pdf(item)
                                for item in value
                                if str(item or "").strip()
                            ]
                        )
                    elif key == "ficheiros" and isinstance(value, list):
                        node[key] = dedupe(
                            [
                                (
                                    store_pdf(item)
                                    if str(item or "").strip().lower().endswith(".pdf")
                                    else self._store_shared_file(
                                        item,
                                        "technical_files",
                                        preferred_name=self._file_reference_name(item, "ficheiro"),
                                    )
                                )
                                for item in value
                                if str(item or "").strip()
                            ]
                        )
                    else:
                        normalize_drawings(value)
                if str(node.get("desenho_pdf", "") or "").strip():
                    node["desenhos_pdf"] = dedupe(
                        [
                            str(node.get("desenho_pdf", "") or "").strip(),
                            *list(node.get("desenhos_pdf", []) or []),
                        ]
                    )
                if isinstance(node.get("ficheiros"), list):
                    node["ficheiros"] = dedupe(
                        [
                            *list(node.get("ficheiros", []) or []),
                            *[
                                item
                                for item in [node.get("desenho_pdf", ""), *list(node.get("desenhos_pdf", []) or [])]
                                if str(item or "").strip()
                            ],
                        ]
                    )
                self._apply_piece_pdf_references(node, self._piece_pdf_references(node, digest_cache=digest_cache), digest_cache=digest_cache)
            elif isinstance(node, list):
                for item in node:
                    normalize_drawings(item)

        if changed_keys is None:
            normalize_drawings(data)
        else:
            for key in selected_keys:
                normalize_drawings(data.get(key))

        notes = data.get("notas_encomenda", []) if changed_keys is None or "notas_encomenda" in selected_keys else []
        for note in list(notes or []):
            if not isinstance(note, dict):
                continue
            note["fatura_caminho_ultima"] = self._store_shared_file(
                note.get("fatura_caminho_ultima", ""),
                "notas_encomenda/documentos",
                preferred_name=self._file_reference_name(note.get("fatura_caminho_ultima", ""), "fatura"),
            )
            for key in ("documentos", "entregas"):
                for row in list(note.get(key, []) or []):
                    if not isinstance(row, dict):
                        continue
                    row["caminho"] = self._store_shared_file(
                        row.get("caminho", ""),
                        "notas_encomenda/documentos",
                        preferred_name=self._file_reference_name(row.get("caminho", ""), row.get("titulo", "") or "documento"),
                    )

        billing_rows = data.get("faturacao", []) if changed_keys is None or "faturacao" in selected_keys else []
        for record in list(billing_rows or []):
            if not isinstance(record, dict):
                continue
            for invoice in list(record.get("faturas", []) or []):
                if not isinstance(invoice, dict):
                    continue
                invoice_number = str(invoice.get("numero_fatura", "") or invoice.get("id", "") or "fatura").strip() or "fatura"
                invoice["caminho"] = self._store_shared_file(
                    invoice.get("caminho", ""),
                    "billing/invoices",
                    preferred_name=self._file_reference_name(invoice.get("caminho", ""), f"{invoice_number}.pdf"),
                )
                invoice["communication_filename"] = self._store_shared_file(
                    invoice.get("communication_filename", ""),
                    "billing/compliance",
                    preferred_name=self._file_reference_name(invoice.get("communication_filename", ""), f"{invoice_number}_at.xml"),
                )
            for payment in list(record.get("pagamentos", []) or []):
                if not isinstance(payment, dict):
                    continue
                payment_name = str(payment.get("titulo_comprovativo", "") or payment.get("referencia", "") or payment.get("id", "") or "comprovativo").strip()
                payment["caminho_comprovativo"] = self._store_shared_file(
                    payment.get("caminho_comprovativo", ""),
                    "billing/payments",
                    preferred_name=self._file_reference_name(payment.get("caminho_comprovativo", ""), payment_name or "comprovativo"),
                )
