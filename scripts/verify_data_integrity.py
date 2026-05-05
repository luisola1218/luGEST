from __future__ import annotations

import json
import re
from collections import Counter
from copy import deepcopy
from datetime import datetime
from pathlib import Path
from typing import Any

import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_qt.services.main_bridge import LegacyBackend


def _num(value: Any, default: float = 0.0) -> float:
    try:
        return float(str(value).replace(",", "."))
    except Exception:
        return default


def _line_total(line: dict[str, Any]) -> float:
    stored = _num(line.get("total", 0), 0.0)
    if stored > 0:
        return stored
    qty = _num(line.get("quantidade", line.get("qtd", 0)), 0.0)
    preco = _num(line.get("preco_unit", line.get("preco", 0)), 0.0)
    desconto = max(0.0, min(100.0, _num(line.get("desconto", 0), 0.0)))
    iva = max(0.0, min(100.0, _num(line.get("iva", 23), 23.0)))
    base = qty * preco
    base *= 1.0 - (desconto / 100.0)
    return base * (1.0 + iva / 100.0)


def _detect_retalho_like(record: dict[str, Any]) -> bool:
    loc = str(record.get("Localizacao", record.get("Localização", "")) or "").strip().upper()
    tipo = str(record.get("tipo", "") or "").strip().lower()
    if bool(record.get("is_sobra")):
        return True
    if loc == "RETALHO":
        return True
    return "retalho" in tipo


def _rebuild_refs(data: dict[str, Any], backend: LegacyBackend) -> list[str]:
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
    for raw in list(data.get("orc_refs", {}) or {}):
        push(raw)
    for payload in list((data.get("orc_refs", {}) or {}).values()):
        if not isinstance(payload, dict):
            continue
        push(payload.get("ref_interna"))
        push(payload.get("ref_externa"))
    return refs


def _max_numeric_suffix(rows: list[dict[str, Any]], key: str, pattern: str) -> int:
    highest = 0
    regex = re.compile(pattern)
    for row in rows:
        value = str((row or {}).get(key, "") or "").strip()
        match = regex.search(value)
        if not match:
            continue
        try:
            highest = max(highest, int(match.group(1)))
        except Exception:
            continue
    return highest


def audit(fix_safe: bool = False) -> dict[str, Any]:
    backend = LegacyBackend()
    data = backend.ensure_data()
    desktop = backend.desktop_main
    now_iso = getattr(desktop, "now_iso", lambda: datetime.now().isoformat())

    pieces: list[dict[str, Any]] = []
    for enc in list(data.get("encomendas", []) or []):
        for piece in list(desktop.encomenda_pecas(enc)):
            pieces.append(
                {
                    "encomenda": str(enc.get("numero", "") or "").strip(),
                    "cliente": str(enc.get("cliente", "") or "").strip(),
                    "id": str(piece.get("id", "") or "").strip(),
                    "ref_interna": str(piece.get("ref_interna", "") or "").strip(),
                    "ref_externa": str(piece.get("ref_externa", "") or "").strip(),
                    "of": str(piece.get("of", "") or "").strip(),
                    "opp": str(piece.get("opp", "") or "").strip(),
                    "material": str(piece.get("material", "") or "").strip(),
                    "espessura": str(piece.get("espessura", "") or "").strip(),
                    "piece": piece,
                }
            )

    ref_groups: dict[str, list[dict[str, Any]]] = {}
    for row in pieces:
        ref_int = str(row.get("ref_interna", "") or "").strip()
        if not ref_int:
            continue
        ref_groups.setdefault(ref_int, []).append(row)
    refs_dup: list[dict[str, Any]] = []
    refs_reused_same_article: list[dict[str, Any]] = []
    for ref_int, group in sorted(ref_groups.items()):
        if len(group) <= 1:
            continue
        signatures = sorted(
            {
                (
                    str(row.get("cliente", "") or "").strip().upper(),
                    str(row.get("ref_externa", "") or "").strip().upper(),
                    str(row.get("material", "") or "").strip().upper(),
                    str(row.get("espessura", "") or "").strip(),
                )
                for row in group
            }
        )
        detail_rows = [
            {
                "encomenda": str(row.get("encomenda", "") or "").strip(),
                "id": str(row.get("id", "") or "").strip(),
                "ref_externa": str(row.get("ref_externa", "") or "").strip(),
                "material": str(row.get("material", "") or "").strip(),
                "espessura": str(row.get("espessura", "") or "").strip(),
                "of": str(row.get("of", "") or "").strip(),
                "opp": str(row.get("opp", "") or "").strip(),
            }
            for row in group
        ]
        payload = {
            "ref_interna": ref_int,
            "rows": detail_rows,
        }
        if len(signatures) <= 1:
            refs_reused_same_article.append(payload)
        else:
            refs_dup.append(payload)
    opp_dup = sorted([key for key, count in Counter(row["opp"] for row in pieces if row["opp"]).items() if count > 1])
    of_orders: dict[str, set[str]] = {}
    for row in pieces:
        of_txt = str(row.get("of", "") or "").strip()
        enc_txt = str(row.get("encomenda", "") or "").strip()
        if of_txt:
            of_orders.setdefault(of_txt, set()).add(enc_txt)
    of_dup = sorted([key for key, orders in of_orders.items() if len({order for order in orders if order}) > 1])

    prefix_mismatch: list[dict[str, Any]] = []
    duplicated_ref_externa: list[dict[str, Any]] = []
    ref_ext_counter = Counter((row["encomenda"], row["ref_externa"]) for row in pieces if row["ref_externa"])
    for row in pieces:
        cliente = row["cliente"]
        ref_int = row["ref_interna"]
        ref_ext = row["ref_externa"]
        if cliente and ref_int and not ref_int.upper().startswith(f"{cliente.upper()}-"):
            prefix_mismatch.append(
                {
                    "encomenda": row["encomenda"],
                    "cliente": cliente,
                    "id": row["id"],
                    "ref_interna": ref_int,
                    "ref_externa": ref_ext,
                }
            )
        if ref_ext and ref_ext_counter[(row["encomenda"], ref_ext)] > 1:
            duplicated_ref_externa.append(
                {
                    "encomenda": row["encomenda"],
                    "id": row["id"],
                    "ref_externa": ref_ext,
                    "ref_interna": ref_int,
                }
            )

    materials = list(data.get("materiais", []) or [])
    stock_issues: list[dict[str, Any]] = []
    retalho_fixed = 0
    retalho_issues: list[dict[str, Any]] = []
    for material in materials:
        qty = _num(material.get("quantidade", 0), 0.0)
        reserved = _num(material.get("reservado", 0), 0.0)
        if qty < 0 or reserved < 0 or reserved > qty + 1e-9:
            stock_issues.append(
                {
                    "id": str(material.get("id", "") or "").strip(),
                    "quantidade": qty,
                    "reservado": reserved,
                    "material": str(material.get("material", "") or "").strip(),
                    "espessura": str(material.get("espessura", "") or "").strip(),
                }
            )
        if not _detect_retalho_like(material):
            continue
        before = (
            round(_num(material.get("peso_unid", 0), 0.0), 6),
            round(_num(material.get("preco_unid", 0), 0.0), 6),
            str(material.get("origem_lote", "") or "").strip(),
            tuple(material.get("origem_lotes_baixa", []) or []),
        )
        target = material if fix_safe else deepcopy(material)
        backend.materia_actions._hydrate_retalho_record(data, target)
        lote = str(target.get("origem_lote", "") or target.get("lote_fornecedor", "") or "").strip()
        if lote and not str(target.get("origem_lote", "") or "").strip():
            target["origem_lote"] = lote
        if lote and not list(target.get("origem_lotes_baixa", []) or []):
            target["origem_lotes_baixa"] = [lote]
        target["atualizado_em"] = now_iso()
        after = (
            round(_num(target.get("peso_unid", 0), 0.0), 6),
            round(_num(target.get("preco_unid", 0), 0.0), 6),
            str(target.get("origem_lote", "") or "").strip(),
            tuple(target.get("origem_lotes_baixa", []) or []),
        )
        if before != after:
            retalho_fixed += 1
        if after[0] <= 0 or not after[2]:
            retalho_issues.append(
                {
                    "id": str(material.get("id", "") or "").strip(),
                    "material": str(material.get("material", "") or "").strip(),
                    "espessura": str(material.get("espessura", "") or "").strip(),
                    "lote": after[2],
                    "peso_unid": after[0],
                }
            )

    enc_nums = {str(enc.get("numero", "") or "").strip() for enc in list(data.get("encomendas", []) or [])}
    orphan_plan: list[dict[str, Any]] = []
    orphan_plan_ids: list[str] = []
    for row in list(data.get("plano", []) or []):
        enc_num = str(row.get("encomenda", "") or "").strip()
        if enc_num and enc_num not in enc_nums:
            orphan_plan.append({"id": str(row.get("id", "") or "").strip(), "encomenda": enc_num})
            orphan_plan_ids.append(str(row.get("id", "") or "").strip())

    expedicoes = list(data.get("expedicoes", []) or [])
    expedition_issues: list[dict[str, Any]] = []
    for row in expedicoes:
        enc_num = str(row.get("encomenda", "") or "").strip()
        if enc_num and enc_num not in enc_nums:
            expedition_issues.append({"numero": str(row.get("numero", "") or "").strip(), "tipo": "missing_order", "encomenda": enc_num})

    notes = list(data.get("notas_encomenda", []) or [])
    notes_fixed = 0
    note_issues: list[dict[str, Any]] = []
    for note in notes:
        lines = list(note.get("linhas", []) or [])
        calc_total = round(sum(_line_total(line) for line in lines), 2)
        stored_total = round(_num(note.get("total", 0), 0.0), 2)
        if abs(calc_total - stored_total) > 0.009:
            note_issues.append(
                {
                    "numero": str(note.get("numero", "") or "").strip(),
                    "guardado": stored_total,
                    "calculado": calc_total,
                    "linhas": len(lines),
                }
            )
            if fix_safe:
                note["total"] = calc_total
                notes_fixed += 1

    changed = False
    if fix_safe:
        if retalho_fixed > 0 or notes_fixed > 0 or orphan_plan_ids:
            changed = True
        if orphan_plan_ids:
            data["plano"] = [row for row in list(data.get("plano", []) or []) if str(row.get("id", "") or "").strip() not in set(orphan_plan_ids)]
        new_refs = _rebuild_refs(data, backend)
        if new_refs != list(data.get("refs", []) or []):
            data["refs"] = new_refs
            changed = True
        data.setdefault("seq", {})
        data["of_seq"] = max(
            _max_numeric_suffix([{"of": row["of"]} for row in pieces], "of", r"OF-\d{4}-(\d{4})$") + 1,
            int(_num(data.get("of_seq", 1), 1)),
        )
        data["opp_seq"] = max(
            _max_numeric_suffix([{"opp": row["opp"]} for row in pieces], "opp", r"OPP-\d{4}-(\d{4})$") + 1,
            int(_num(data.get("opp_seq", 1), 1)),
        )
        data["orc_seq"] = max(
            _max_numeric_suffix(list(data.get("orcamentos", []) or []), "numero", r"ORC-\d{4}-(\d{4})$") + 1,
            int(_num(data.get("orc_seq", 1), 1)),
        )
        data["seq"]["ne"] = max(
            _max_numeric_suffix(list(data.get("notas_encomenda", []) or []), "numero", r"NE-\d{4}-(\d{4})$") + 1,
            int(_num(data.get("seq", {}).get("ne", 1), 1)),
        )
        data["exp_seq"] = max(
            _max_numeric_suffix(expedicoes, "numero", r"GT-\d{4}-(\d{1,})$") + 1,
            int(_num(data.get("exp_seq", 1), 1)),
        )
        if changed:
            backend._save(force=True)

    report = {
        "counts": {
            "encomendas": len(list(data.get("encomendas", []) or [])),
            "pecas": len(pieces),
            "materiais": len(materials),
            "notas_encomenda": len(notes),
            "expedicoes": len(expedicoes),
            "plano": len(list(data.get("plano", []) or [])),
        },
        "safe_fixes": {
            "retalhos_rehidratados": retalho_fixed,
            "notas_total_recalculado": notes_fixed,
            "blocos_orfaos_removidos": len(orphan_plan_ids) if fix_safe else 0,
        },
        "issues_safe": {
            "stock": stock_issues,
            "retalhos": retalho_issues,
            "notas": note_issues,
            "plano_orfao": orphan_plan,
            "expedicoes": expedition_issues,
            "ref_interna_reutilizada": refs_reused_same_article,
        },
        "issues_risky": {
            "ref_interna_prefixo_errado": prefix_mismatch,
            "ref_externa_duplicada": duplicated_ref_externa,
            "ref_interna_duplicada": refs_dup,
            "opp_duplicada": opp_dup,
            "of_duplicada": of_dup,
        },
    }
    return report


def main() -> int:
    import argparse

    parser = argparse.ArgumentParser(description="Audit and optionally repair safe data integrity issues.")
    parser.add_argument("--fix-safe", action="store_true", help="Apply only safe repairs.")
    args = parser.parse_args()
    report = audit(fix_safe=args.fix_safe)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    print("data-integrity-ok")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
