from __future__ import annotations

import sys
from collections import Counter
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_qt.services.main_bridge import LegacyBackend


def _piece_safe_to_rename(data: dict[str, Any], piece: dict[str, Any]) -> bool:
    ref = str(piece.get("ref_interna", "") or "").strip()
    piece_id = str(piece.get("id", "") or "").strip()
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
    if str(piece.get("estado", "") or "").strip().lower() not in ("", "preparacao", "preparação"):
        return False
    for ev in list(data.get("op_eventos", []) or []):
        if str(ev.get("ref_interna", "") or "").strip() == ref or str(ev.get("peca_id", "") or "").strip() == piece_id:
            return False
    for row in list(data.get("plano", []) or []):
        if str(row.get("ref_interna", "") or "").strip() == ref or str(row.get("peca_id", "") or "").strip() == piece_id:
            return False
    for ex in list(data.get("expedicoes", []) or []):
        for line in list(ex.get("linhas", []) or []):
            if str(line.get("ref_interna", "") or "").strip() == ref or str(line.get("peca_id", "") or "").strip() == piece_id:
                return False
    return True


def _rebuild_refs(data: dict[str, Any], backend: LegacyBackend) -> None:
    refs: list[str] = []
    seen: set[str] = set()

    def push(value: Any) -> None:
        txt = str(value or "").strip()
        if not txt or txt in seen:
            return
        seen.add(txt)
        refs.append(txt)

    for enc in list(data.get("encomendas", []) or []):
        for piece in list(backend.desktop_main.encomenda_pecas(enc)):
            push(piece.get("ref_interna"))
            push(piece.get("ref_externa"))
    for key, payload in list((data.get("orc_refs", {}) or {}).items()):
        push(key)
        if isinstance(payload, dict):
            push(payload.get("ref_interna"))
            push(payload.get("ref_externa"))
    data["refs"] = refs


def main() -> int:
    backend = LegacyBackend()
    data = backend.ensure_data()
    desktop = backend.desktop_main
    changed: list[dict[str, str]] = []

    for quote in list(data.get("orcamentos", []) or []):
        cliente = str(((quote or {}).get("cliente") or {}).get("codigo", "") or "").strip()
        if not cliente:
            continue
        order_num = str(quote.get("numero_encomenda", "") or "").strip()
        order = backend.get_encomenda_by_numero(order_num) if order_num else None
        order_pieces = list(desktop.encomenda_pecas(order)) if order else []

        quote_seen: Counter[str] = Counter()
        for line in list(quote.get("linhas", []) or []):
            ref_int = str(line.get("ref_interna", "") or "").strip()
            if ref_int:
                quote_seen[ref_int] += 1

        for line in list(quote.get("linhas", []) or []):
            ref_int = str(line.get("ref_interna", "") or "").strip()
            ref_ext = str(line.get("ref_externa", "") or "").strip()
            material = str(line.get("material", "") or "").strip()
            esp = str(line.get("espessura", "") or "").strip()

            matching_piece = None
            if order_pieces:
                candidates = [
                    piece
                    for piece in order_pieces
                    if str(piece.get("ref_externa", "") or "").strip() == ref_ext
                    and str(piece.get("material", "") or "").strip() == material
                    and str(piece.get("espessura", "") or "").strip() == esp
                ]
                if len(candidates) == 1:
                    matching_piece = candidates[0]

            piece_ref = str((matching_piece or {}).get("ref_interna", "") or "").strip()
            piece_safe = bool(matching_piece) and _piece_safe_to_rename(data, matching_piece)
            mismatch = not ref_int.upper().startswith(f"{cliente.upper()}-") if ref_int else True
            duplicate_in_quote = bool(ref_int) and quote_seen[ref_int] > 1

            if matching_piece and piece_ref.upper().startswith(f"{cliente.upper()}-") and not duplicate_in_quote and ref_int != piece_ref:
                old_ref = ref_int
                line["ref_interna"] = piece_ref
                changed.append(
                    {
                        "scope": "quote-sync",
                        "documento": str(quote.get("numero", "") or "").strip(),
                        "ref_antiga": old_ref,
                        "ref_nova": piece_ref,
                    }
                )
                continue

            if not mismatch and not duplicate_in_quote:
                continue
            if matching_piece and not piece_safe:
                continue

            existing = [
                str(piece.get("ref_interna", "") or "").strip()
                for enc in list(data.get("encomendas", []) or [])
                if str(enc.get("cliente", "") or "").strip() == cliente
                for piece in list(desktop.encomenda_pecas(enc))
            ]
            existing += [
                str(l.get("ref_interna", "") or "").strip()
                for orc in list(data.get("orcamentos", []) or [])
                for l in list(orc.get("linhas", []) or [])
                if str((((orc or {}).get("cliente") or {}).get("codigo", "") or "").strip()) == cliente
            ]
            for remove_ref in filter(None, [ref_int, piece_ref]):
                while remove_ref in existing:
                    existing.remove(remove_ref)
            new_ref = desktop.next_ref_interna_unique(data, cliente, existing)
            if not new_ref:
                continue

            old_ref = ref_int
            line["ref_interna"] = new_ref
            if matching_piece:
                matching_piece["ref_interna"] = new_ref
            if ref_ext and Counter(str(l.get("ref_externa", "") or "").strip() for l in list(quote.get("linhas", []) or []))[ref_ext] == 1:
                payload = (data.get("orc_refs", {}) or {}).get(ref_ext)
                if isinstance(payload, dict):
                    payload["ref_interna"] = new_ref
            changed.append(
                {
                    "scope": "quote-order-fix" if matching_piece else "quote-fix",
                    "documento": str(quote.get("numero", "") or "").strip(),
                    "ref_antiga": old_ref,
                    "ref_nova": new_ref,
                }
            )

    if changed:
        _rebuild_refs(data, backend)
        backend._save(force=True)

    for row in changed:
        print(f"{row['scope']}: {row['documento']} {row['ref_antiga']} -> {row['ref_nova']}")
    print(f"reference-fix-ok {len(changed)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
