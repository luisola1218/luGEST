from __future__ import annotations

import copy
import sys
import tempfile
import uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_qt.services.main_bridge import LegacyBackend
from pypdf import PdfReader


def _assert(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def _pick_client(backend: LegacyBackend, suffix: str) -> dict[str, str]:
    for row in list(backend.ensure_data().get("clientes", []) or []):
        if not isinstance(row, dict):
            continue
        code = str(row.get("codigo", "") or "").strip()
        if code:
            return {
                "codigo": code,
                "nome": str(row.get("nome", "") or "").strip(),
                "nif": str(row.get("nif", "") or "").strip(),
                "morada": str(row.get("morada", "") or "").strip(),
                "contacto": str(row.get("contacto", "") or "").strip(),
                "email": str(row.get("email", "") or "").strip(),
            }
    return {
        "codigo": f"CLPDF{suffix[:4]}",
        "nome": f"Cliente PDF Vertical {suffix}",
        "nif": "",
        "morada": "Rua de Teste, 1",
        "contacto": "",
        "email": "",
    }


def main() -> int:
    backend = LegacyBackend()
    data = backend.ensure_data()
    token = uuid.uuid4().hex[:8].upper()
    quote_num = f"ORC-VERIFY-PDF-{token}"
    pdf_path = Path(tempfile.gettempdir()) / f"lugest_verify_quote_vertical_{token}.pdf"
    orc_seq_snapshot = data.get("orc_seq", 1)
    seq_snapshot = copy.deepcopy(dict(data.get("seq", {}) or {}))

    try:
        client = _pick_client(backend, token)
        lines = []
        for idx in range(22):
            is_dxf = idx % 3 == 0
            lines.append(
                {
                    "tipo_item": backend.desktop_main.ORC_LINE_TYPE_PIECE,
                    "ref_externa": f"REF-EXT-PDF-{token}-{idx:02d}",
                    "ref_interna": f"INT-{token}-{idx:02d}",
                    "descricao": (
                        "Conjunto vertical de teste com descricao longa para validar quebra de texto "
                        f"e recorte de celulas na linha {idx + 1}"
                    ),
                    "material": "AISI 304L",
                    "espessura": "2",
                    "operacao": "Corte Laser / Quinagem" if is_dxf else "",
                    "desenho": f"peca_{idx:02d}.dxf" if is_dxf else "",
                    "qtd": 2 + (idx % 4),
                    "preco_unit": 12.5 + idx,
                    "desconto_perc": 0,
                }
            )

        quote = backend.orc_save(
            {
                "numero": quote_num,
                "estado": "Aprovado",
                "cliente": client,
                "linhas": lines,
                "iva_perc": 23,
                "incremento_preco_perc": 30,
                "desconto_perc": 10,
                "desconto_modo": "total",
                "nota_transporte": "Transporte a Nosso Cargo",
                "preco_transporte": 123.45,
                "nota_cliente": "VERIFY_PDF_VERTICAL",
                "executado_por": "VERIFY",
                "nota_transporte": "Entrega acordada",
                "prazo_entrega_texto": "A combinar",
            }
        )
        _assert(str(quote.get("numero", "") or "").strip() == quote_num, f"Orcamento nao guardado: {quote}")

        rendered = backend.orc_render_pdf(quote_num, pdf_path)
        _assert(rendered.exists(), f"PDF nao foi criado: {rendered}")
        _assert(rendered.stat().st_size > 3500, f"PDF demasiado pequeno: {rendered.stat().st_size}")

        reader = PdfReader(str(rendered))
        extracted_text = "\n".join((page.extract_text() or "") for page in reader.pages)
        first_page = reader.pages[0]
        page_w = float(first_page.mediabox.width)
        page_h = float(first_page.mediabox.height)
        _assert(page_w < page_h, f"PDF nao esta em A4 vertical: {page_w}x{page_h}.")
        _assert(len(reader.pages) >= 2, f"PDF vertical devia paginar em multiplas paginas: {len(reader.pages)}")
        for token_text in ("Orcamento", quote_num, "Valor sem IVA", "Ao valor sem IVA acresce o IVA", "Transporte", "123.45 EUR"):
            _assert(token_text in extracted_text, f"PDF sem o token esperado `{token_text}`.")
        for token_text in ("Descricao", "Serra"):
            _assert(token_text in extracted_text, f"PDF sem o token esperado `{token_text}`.")
        _assert("Incremento" not in extracted_text, "PDF do cliente nao deve expor o incremento interno de precos.")

        print("quote-vertical-pdf-ok", quote_num, rendered)
        return 0
    finally:
        try:
            data["orcamentos"] = [
                row
                for row in list(data.get("orcamentos", []) or [])
                if str((row or {}).get("numero", "") or "").strip() != quote_num
            ]
            data["orc_seq"] = orc_seq_snapshot
            data["seq"] = seq_snapshot
            backend._save(force=True)
        finally:
            if pdf_path.exists():
                try:
                    pdf_path.unlink()
                except Exception:
                    pass


if __name__ == "__main__":
    raise SystemExit(main())
