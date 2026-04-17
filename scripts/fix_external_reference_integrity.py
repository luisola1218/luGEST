from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_qt.services.main_bridge import LegacyBackend


def _piece_safe(data: dict[str, Any], piece: dict[str, Any]) -> bool:
    piece_id = str(piece.get("id", "") or "").strip()
    ref_int = str(piece.get("ref_interna", "") or "").strip()
    if not piece_id:
        return False
    if float(piece.get("produzido_ok", 0) or 0) > 0:
        return False
    if float(piece.get("produzido_nok", 0) or 0) > 0:
        return False
    if float(piece.get("produzido_qualidade", 0) or 0) > 0:
        return False
    if float(piece.get("qtd_expedida", 0) or 0) > 0:
        return False
    for ev in list(data.get("op_eventos", []) or []):
        if str(ev.get("peca_id", "") or "").strip() == piece_id or str(ev.get("ref_interna", "") or "").strip() == ref_int:
            return False
    for row in list(data.get("plano", []) or []):
        if str(row.get("peca_id", "") or "").strip() == piece_id or str(row.get("ref_interna", "") or "").strip() == ref_int:
            return False
    for ex in list(data.get("expedicoes", []) or []):
        for line in list(ex.get("linhas", []) or []):
            if str(line.get("peca_id", "") or "").strip() == piece_id or str(line.get("ref_interna", "") or "").strip() == ref_int:
                return False
    return True


def _next_test_ref(existing: set[str], cliente: str) -> str:
    regex = re.compile(rf"^{re.escape(cliente)}-TEST-(\d+)$", re.I)
    highest = 0
    for ref in existing:
        match = regex.match(str(ref or "").strip())
        if not match:
            continue
        try:
            highest = max(highest, int(match.group(1)))
        except Exception:
            continue
    while True:
        highest += 1
        candidate = f"{cliente}-TEST-{highest:02d}"
        if candidate not in existing:
            return candidate


def main() -> int:
    backend = LegacyBackend()
    data = backend.ensure_data()
    desktop = backend.desktop_main

    active_exts = {
        str(line.get("ref_externa", "") or "").strip()
        for quote in list(data.get("orcamentos", []) or [])
        for line in list(quote.get("linhas", []) or [])
        if str(line.get("ref_externa", "") or "").strip()
    }
    active_exts.update(
        {
            str(piece.get("ref_externa", "") or "").strip()
            for enc in list(data.get("encomendas", []) or [])
            for piece in list(desktop.encomenda_pecas(enc))
            if str(piece.get("ref_externa", "") or "").strip()
        }
    )

    changed: list[tuple[str, str, str]] = []

    for quote in list(data.get("orcamentos", []) or []):
        cliente = str(((quote or {}).get("cliente") or {}).get("codigo", "") or "").strip()
        if not cliente:
            continue
        order_num = str(quote.get("numero_encomenda", "") or "").strip()
        order = backend.get_encomenda_by_numero(order_num) if order_num else None
        order_pieces = list(desktop.encomenda_pecas(order)) if order else []

        seen_in_quote: set[str] = set()
        for line in list(quote.get("linhas", []) or []):
            ref_ext = str(line.get("ref_externa", "") or "").strip()
            material = str(line.get("material", "") or "").strip()
            esp = str(line.get("espessura", "") or "").strip()
            matching_piece = None
            if order_pieces:
                candidates = [
                    piece
                    for piece in order_pieces
                    if str(piece.get("ref_interna", "") or "").strip() == str(line.get("ref_interna", "") or "").strip()
                    and str(piece.get("material", "") or "").strip() == material
                    and str(piece.get("espessura", "") or "").strip() == esp
                ]
                if len(candidates) == 1:
                    matching_piece = candidates[0]
            if matching_piece and not _piece_safe(data, matching_piece):
                if ref_ext:
                    seen_in_quote.add(ref_ext)
                continue

            needs_fix = False
            if not ref_ext:
                needs_fix = True
            elif ref_ext in seen_in_quote:
                needs_fix = True
            elif ref_ext.upper().startswith("CL") and not ref_ext.upper().startswith(f"{cliente.upper()}-"):
                needs_fix = True

            if not needs_fix:
                seen_in_quote.add(ref_ext)
                continue

            new_ref = _next_test_ref(active_exts | seen_in_quote, cliente)
            old_ref = ref_ext
            line["ref_externa"] = new_ref
            if matching_piece:
                matching_piece["ref_externa"] = new_ref
            payload = (data.get("orc_refs", {}) or {}).get(old_ref)
            if isinstance(payload, dict):
                data.setdefault("orc_refs", {})[new_ref] = dict(payload)
                data["orc_refs"][new_ref]["ref_externa"] = new_ref
                data["orc_refs"][new_ref]["ref_interna"] = str(line.get("ref_interna", "") or "").strip()
                data["orc_refs"][new_ref]["material"] = material
                data["orc_refs"][new_ref]["espessura"] = esp
                data["orc_refs"][new_ref]["descricao"] = str(line.get("descricao", "") or "").strip()
            active_exts.add(new_ref)
            seen_in_quote.add(new_ref)
            changed.append((str(quote.get("numero", "") or "").strip(), old_ref, new_ref))

    if changed:
        refs: list[str] = []
        seen_refs: set[str] = set()
        for enc in list(data.get("encomendas", []) or []):
            for piece in list(desktop.encomenda_pecas(enc)):
                for value in (piece.get("ref_interna"), piece.get("ref_externa")):
                    txt = str(value or "").strip()
                    if txt and txt not in seen_refs:
                        seen_refs.add(txt)
                        refs.append(txt)
        for key, payload in list((data.get("orc_refs", {}) or {}).items()):
            for value in (key, (payload or {}).get("ref_interna"), (payload or {}).get("ref_externa")):
                txt = str(value or "").strip()
                if txt and txt not in seen_refs:
                    seen_refs.add(txt)
                    refs.append(txt)
        data["refs"] = refs
        backend._save(force=True)

    for quote_num, old_ref, new_ref in changed:
        print(f"{quote_num}: {old_ref or '<vazio>'} -> {new_ref}")
    print(f"external-reference-fix-ok {len(changed)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
