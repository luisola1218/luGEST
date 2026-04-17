from __future__ import annotations

import sys
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import lugest_storage
from lugest_qt.services.main_bridge import LegacyBackend


def _must_find(node: ET.Element, path: str) -> ET.Element:
    found = node.find(path)
    if found is None:
        raise RuntimeError(f"Elemento XML em falta: {path}")
    return found


def main() -> int:
    backend = LegacyBackend()
    quote_num = ""
    order_num = ""
    record_num = ""
    created_paths: list[Path] = []
    try:
        quote = backend.orc_save(
            {
                "cliente": {"codigo": "CL0001", "nome": "Verto"},
                "estado": "Aprovado",
                "linhas": [
                    {
                        "tipo_item": "Peca",
                        "ref_externa": "VERIFY-SAFT-001",
                        "descricao": "VERIFY TAX COMPLIANCE",
                        "material": "S235JR",
                        "espessura": "2",
                        "operacao": "Corte Laser + Embalamento",
                        "qtd": 8,
                        "preco_unit": 18.5,
                        "tempo_peca_min": 1,
                    }
                ],
            }
        )
        quote_num = str(quote.get("numero", "") or "").strip()
        backend.orc_set_state(quote_num, "Aprovado")
        converted = backend.orc_convert_to_order(quote_num, "VERIFY TAX")
        order_num = str((converted.get("encomenda") or {}).get("numero", "") or "").strip()

        detail = backend.billing_open_record(source_type="quote", source_number=quote_num)
        record_num = str(detail.get("numero", "") or "").strip()
        if not record_num:
            raise RuntimeError(f"Registo de faturação sem número: {detail}")

        payload = dict(backend.billing_invoice_defaults(record_num))
        payload["data_emissao"] = "2026-03-25"
        payload["data_vencimento"] = "2026-04-24"
        detail = backend.billing_generate_invoice_pdf(record_num, payload)
        invoice = dict((detail.get("invoices") or [{}])[0] or {})
        invoice_id = str(invoice.get("id", "") or "").strip()
        if not invoice_id:
            raise RuntimeError(f"Fatura não criada: {detail}")
        for field in ("system_entry_date", "source_id", "legal_invoice_no", "hash", "document_snapshot_json"):
            if not str(invoice.get(field, "") or "").strip():
                raise RuntimeError(f"Campo de conformidade em falta ({field}): {invoice}")
        if str(invoice.get("communication_status", "") or "").strip() != "Por comunicar":
            raise RuntimeError(f"Estado de comunicação inicial inválido: {invoice}")

        try:
            backend.billing_add_invoice(
                record_num,
                {
                    "id": invoice_id,
                    "numero_fatura": "FT-ALTERADA-9999",
                    "valor_total": invoice.get("valor_total", 0),
                },
            )
        except Exception as exc:
            if "emitida" not in str(exc).lower():
                raise RuntimeError(f"Bloqueio de edição fiscal devolveu erro inesperado: {exc}") from exc
        else:
            raise RuntimeError("Foi possível alterar os dados fiscais de uma fatura já emitida.")

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            saft_path = Path(backend.billing_export_saft_pt(output_path=temp_path / "saft_pt.xml"))
            batch_path = Path(backend.billing_prepare_at_communication_batch(output_path=temp_path / "at_preparacao.xml"))
            created_paths.extend([saft_path, batch_path])

            saft_root = ET.parse(saft_path).getroot()
            if saft_root.tag != "AuditFile":
                raise RuntimeError(f"Raiz SAF-T inválida: {saft_root.tag}")
            invoice_nodes = saft_root.findall("./SourceDocuments/SalesInvoices/Invoice")
            invoice_node = next(
                (
                    node
                    for node in invoice_nodes
                    if (_must_find(node, "InvoiceNo").text or "").strip() == str(invoice.get("legal_invoice_no", "") or "").strip()
                ),
                None,
            )
            if invoice_node is None:
                raise RuntimeError("A fatura de teste não apareceu no SAF-T exportado.")
            hash_text = (_must_find(invoice_node, "Hash").text or "").strip()
            cert_number = (_must_find(saft_root, "./Header/SoftwareCertificateNumber").text or "").strip()
            if cert_number == "0":
                if hash_text != "0":
                    raise RuntimeError(f"Hash do SAF-T devia sair a 0 sem nº AT configurado: {hash_text}")
            elif hash_text != str(invoice.get("hash", "") or "").strip():
                raise RuntimeError(f"Hash do SAF-T não corresponde ao documento: {hash_text} vs {invoice.get('hash')}")
            if not (_must_find(invoice_node, "SystemEntryDate").text or "").strip():
                raise RuntimeError("SystemEntryDate não saiu no SAF-T.")

            batch_root = ET.parse(batch_path).getroot()
            if batch_root.tag != "ATCommunicationPreparation":
                raise RuntimeError(f"Raiz preparação AT inválida: {batch_root.tag}")
            doc_nodes = batch_root.findall("./Documents/Document")
            doc_node = next(
                (
                    node
                    for node in doc_nodes
                    if (_must_find(node, "InvoiceNo").text or "").strip() == str(invoice.get("legal_invoice_no", "") or "").strip()
                ),
                None,
            )
            if doc_node is None:
                raise RuntimeError("Lote AT sem InvoiceNo esperado.")

        refreshed = backend.billing_detail(record_num)
        refreshed_invoice = dict((refreshed.get("invoices") or [{}])[0] or {})
        if str(refreshed_invoice.get("communication_status", "") or "").strip() != "Preparada":
            raise RuntimeError(f"Estado de comunicação não ficou Preparada: {refreshed_invoice}")

        print("tax-compliance-exports-ok", quote_num, order_num, record_num, refreshed_invoice.get("legal_invoice_no", ""))
        return 0
    finally:
        for path in created_paths:
            try:
                if path.exists():
                    path.unlink()
            except Exception:
                pass
        data = backend.ensure_data()
        if record_num:
            for row in list(data.get("faturacao", []) or []):
                if str((row or {}).get("numero", "") or "").strip() == record_num:
                    for inv in list((row or {}).get("faturas", []) or []):
                        try:
                            output = lugest_storage.resolve_file_reference(str(inv.get("caminho", "") or "").strip(), base_dir=ROOT)
                            if output.exists():
                                output.unlink()
                        except Exception:
                            pass
            data["faturacao"] = [
                row
                for row in list(data.get("faturacao", []) or [])
                if str((row or {}).get("numero", "") or "").strip() != record_num
            ]
        if order_num:
            try:
                backend.order_remove(order_num)
            except Exception:
                pass
        if quote_num:
            try:
                backend.orc_remove(quote_num)
            except Exception:
                pass
        backend._save(force=True)


if __name__ == "__main__":
    raise SystemExit(main())
