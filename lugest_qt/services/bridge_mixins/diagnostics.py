from __future__ import annotations

import copy
import re
import time
from collections import Counter
from datetime import datetime
from typing import Any


class DiagnosticsBackendMixin:
    """Legacy adapter for diagnostics; see BACKEND_GUIDE.md."""

    def _diagnostic_line_total(self, line: dict[str, Any]) -> float:
        stored = self._parse_float(line.get("total", 0), 0.0)
        if stored > 0:
            return stored
        qty = self._parse_float(line.get("quantidade", line.get("qtd", 0)), 0.0)
        price = self._parse_float(line.get("preco_unit", line.get("preco", 0)), 0.0)
        discount = max(0.0, min(100.0, self._parse_float(line.get("desconto", 0), 0.0)))
        vat = max(0.0, min(100.0, self._parse_float(line.get("iva", 23), 23.0)))
        base = qty * price * (1.0 - (discount / 100.0))
        return base * (1.0 + vat / 100.0)

    def _diagnostic_is_retalho(self, record: dict[str, Any]) -> bool:
        location = str(record.get("Localizacao", record.get("Localização", "")) or "").strip().upper()
        kind = str(record.get("tipo", "") or "").strip().lower()
        return bool(record.get("is_sobra")) or location == "RETALHO" or "retalho" in kind

    def _diagnostic_rebuild_refs(self, data: dict[str, Any]) -> list[str]:
        refs: list[str] = []
        seen: set[str] = set()

        def push(value: Any) -> None:
            text = str(value or "").strip()
            if text and text not in seen:
                seen.add(text)
                refs.append(text)

        for enc in list(data.get("encomendas", []) or []):
            for piece in list(self.desktop_main.encomenda_pecas(enc) or []):
                push(piece.get("ref_interna"))
                push(piece.get("ref_externa"))
        for raw in list(data.get("orc_refs", {}) or {}):
            push(raw)
        for payload in list((data.get("orc_refs", {}) or {}).values()):
            if isinstance(payload, dict):
                push(payload.get("ref_interna"))
                push(payload.get("ref_externa"))
        return refs

    def _diagnostic_max_numeric_suffix(self, rows: list[dict[str, Any]], key: str, pattern: str) -> int:
        highest = 0
        regex = re.compile(pattern)
        for row in list(rows or []):
            value = str((row or {}).get(key, "") or "").strip()
            match = regex.search(value)
            if not match:
                continue
            try:
                highest = max(highest, int(match.group(1)))
            except Exception:
                continue
        return highest

    def system_diagnostics_report(self, fix_safe: bool = False) -> dict[str, Any]:
        data = self.ensure_data()
        now_iso = getattr(self.desktop_main, "now_iso", lambda: datetime.now().isoformat())

        pieces: list[dict[str, Any]] = []
        for enc in list(data.get("encomendas", []) or []):
            for piece in list(self.desktop_main.encomenda_pecas(enc) or []):
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
                    }
                )

        ref_groups: dict[str, list[dict[str, Any]]] = {}
        for row in pieces:
            ref_int = str(row.get("ref_interna", "") or "").strip()
            if ref_int:
                ref_groups.setdefault(ref_int, []).append(row)
        refs_dup: list[dict[str, Any]] = []
        refs_reused_same_article: list[dict[str, Any]] = []
        for ref_int, group in sorted(ref_groups.items()):
            if len(group) <= 1:
                continue
            signatures = {
                (
                    str(row.get("cliente", "") or "").strip().upper(),
                    str(row.get("ref_externa", "") or "").strip().upper(),
                    str(row.get("material", "") or "").strip().upper(),
                    str(row.get("espessura", "") or "").strip(),
                )
                for row in group
            }
            payload = {
                "ref_interna": ref_int,
                "linhas": len(group),
                "encomendas": ", ".join(sorted({str(row.get("encomenda", "") or "").strip() for row in group if row.get("encomenda")})),
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
            cliente = str(row.get("cliente", "") or "").strip()
            ref_int = str(row.get("ref_interna", "") or "").strip()
            ref_ext = str(row.get("ref_externa", "") or "").strip()
            if cliente and ref_int and not ref_int.upper().startswith(f"{cliente.upper()}-"):
                prefix_mismatch.append({"encomenda": row["encomenda"], "cliente": cliente, "ref_interna": ref_int})
            if ref_ext and ref_ext_counter[(row["encomenda"], ref_ext)] > 1:
                duplicated_ref_externa.append({"encomenda": row["encomenda"], "ref_externa": ref_ext, "ref_interna": ref_int})

        materials = list(data.get("materiais", []) or [])
        stock_issues: list[dict[str, Any]] = []
        retalho_fixed = 0
        retalho_issues: list[dict[str, Any]] = []
        for material in materials:
            qty = self._parse_float(material.get("quantidade", 0), 0.0)
            reserved = self._parse_float(material.get("reservado", 0), 0.0)
            if qty < 0 or reserved < 0 or reserved > qty + 1e-9:
                stock_issues.append(
                    {
                        "id": str(material.get("id", "") or "").strip(),
                        "material": str(material.get("material", "") or "").strip(),
                        "espessura": str(material.get("espessura", "") or "").strip(),
                        "quantidade": qty,
                        "reservado": reserved,
                    }
                )
            if not self._diagnostic_is_retalho(material):
                continue
            before = (
                round(self._parse_float(material.get("peso_unid", 0), 0.0), 6),
                round(self._parse_float(material.get("preco_unid", 0), 0.0), 6),
                str(material.get("origem_lote", "") or "").strip(),
                tuple(material.get("origem_lotes_baixa", []) or []),
            )
            target = material if fix_safe else copy.deepcopy(material)
            self.materia_actions._hydrate_retalho_record(data, target)
            lote = str(target.get("origem_lote", "") or target.get("lote_fornecedor", "") or "").strip()
            if lote and not str(target.get("origem_lote", "") or "").strip():
                target["origem_lote"] = lote
            if lote and not list(target.get("origem_lotes_baixa", []) or []):
                target["origem_lotes_baixa"] = [lote]
            if fix_safe:
                target["atualizado_em"] = now_iso()
            after = (
                round(self._parse_float(target.get("peso_unid", 0), 0.0), 6),
                round(self._parse_float(target.get("preco_unid", 0), 0.0), 6),
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
        orphan_plan_ids: set[str] = set()
        for row in list(data.get("plano", []) or []):
            enc_num = str(row.get("encomenda", "") or "").strip()
            if enc_num and enc_num not in enc_nums:
                row_id = str(row.get("id", "") or "").strip()
                orphan_plan.append({"id": row_id, "encomenda": enc_num})
                orphan_plan_ids.add(row_id)

        expedicoes = list(data.get("expedicoes", []) or [])
        expedition_issues = [
            {"numero": str(row.get("numero", "") or "").strip(), "encomenda": str(row.get("encomenda", "") or "").strip()}
            for row in expedicoes
            if str(row.get("encomenda", "") or "").strip() and str(row.get("encomenda", "") or "").strip() not in enc_nums
        ]

        notes = list(data.get("notas_encomenda", []) or [])
        notes_fixed = 0
        note_issues: list[dict[str, Any]] = []
        for note in notes:
            lines = list(note.get("linhas", []) or [])
            calc_total = round(sum(self._diagnostic_line_total(line) for line in lines), 2)
            stored_total = round(self._parse_float(note.get("total", 0), 0.0), 2)
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
                data["plano"] = [row for row in list(data.get("plano", []) or []) if str(row.get("id", "") or "").strip() not in orphan_plan_ids]
            new_refs = self._diagnostic_rebuild_refs(data)
            if new_refs != list(data.get("refs", []) or []):
                data["refs"] = new_refs
                changed = True
            data.setdefault("seq", {})
            seq_updates = {
                "of_seq": max(self._diagnostic_max_numeric_suffix([{"of": row["of"]} for row in pieces], "of", r"OF-\d{4}-(\d{4})$") + 1, int(self._parse_float(data.get("of_seq", 1), 1))),
                "opp_seq": max(self._diagnostic_max_numeric_suffix([{"opp": row["opp"]} for row in pieces], "opp", r"OPP-\d{4}-(\d{4})$") + 1, int(self._parse_float(data.get("opp_seq", 1), 1))),
                "orc_seq": max(self._diagnostic_max_numeric_suffix(list(data.get("orcamentos", []) or []), "numero", r"ORC-\d{4}-(\d{4})$") + 1, int(self._parse_float(data.get("orc_seq", 1), 1))),
                "exp_seq": max(self._diagnostic_max_numeric_suffix(expedicoes, "numero", r"GT-\d{4}-(\d{1,})$") + 1, int(self._parse_float(data.get("exp_seq", 1), 1))),
            }
            for key, value in seq_updates.items():
                if int(self._parse_float(data.get(key, 1), 1)) != int(value):
                    data[key] = int(value)
                    changed = True
            ne_seq = max(self._diagnostic_max_numeric_suffix(list(data.get("notas_encomenda", []) or []), "numero", r"NE-\d{4}-(\d{4})$") + 1, int(self._parse_float(data.get("seq", {}).get("ne", 1), 1)))
            if int(self._parse_float(data.get("seq", {}).get("ne", 1), 1)) != int(ne_seq):
                data["seq"]["ne"] = int(ne_seq)
                changed = True
            if changed:
                self._save(force=True, blocking=True)

        save_state = self.save_runtime_state()
        issues_safe = {
            "stock": stock_issues,
            "retalhos": retalho_issues,
            "notas": note_issues,
            "plano_orfao": orphan_plan,
            "expedicoes": expedition_issues,
            "ref_interna_reutilizada": refs_reused_same_article,
        }
        issues_risky = {
            "ref_interna_prefixo_errado": prefix_mismatch,
            "ref_externa_duplicada": duplicated_ref_externa,
            "ref_interna_duplicada": refs_dup,
            "opp_duplicada": opp_dup,
            "of_duplicada": of_dup,
        }
        critical_count = len(stock_issues) + (1 if save_state.get("last_error") else 0)
        warning_count = sum(len(list(value or [])) for value in issues_safe.values()) + sum(len(list(value or [])) for value in issues_risky.values())
        status = "critical" if critical_count else ("warning" if warning_count else "ok")
        return {
            "generated_at": now_iso(),
            "status": status,
            "critical_count": critical_count,
            "warning_count": warning_count,
            "counts": {
                "encomendas": len(list(data.get("encomendas", []) or [])),
                "pecas": len(pieces),
                "materiais": len(materials),
                "produtos": len(list(data.get("produtos", []) or [])),
                "notas_encomenda": len(notes),
                "expedicoes": len(expedicoes),
                "plano": len(list(data.get("plano", []) or [])),
            },
            "safe_fixes": {
                "retalhos_rehidratados": retalho_fixed,
                "notas_total_recalculado": notes_fixed,
                "blocos_orfaos_removidos": len(orphan_plan_ids) if fix_safe else 0,
                "gravado": bool(changed),
            },
            "issues_safe": issues_safe,
            "issues_risky": issues_risky,
            "runtime": {
                **save_state,
                "cache_age_sec": round(max(0.0, time.time() - float(self._data_loaded_at or 0.0)), 2) if self._data_loaded_at else 0.0,
            },
        }

    def system_diagnostics_fix_safe(self) -> dict[str, Any]:
        return self.system_diagnostics_report(fix_safe=True)

    def audit_rows(self, filter_text: str = "", limit: int = 500) -> list[dict[str, Any]]:
        query = str(filter_text or "").strip().lower()
        rows: list[dict[str, Any]] = []
        for raw in reversed(list(self.ensure_data().get("audit_log", []) or [])):
            if not isinstance(raw, dict):
                continue
            row = {
                "created_at": str(raw.get("created_at", "") or "").strip(),
                "user": str(raw.get("user", "") or "").strip(),
                "action": str(raw.get("action", "") or "").strip(),
                "entity_type": str(raw.get("entity_type", "") or "").strip(),
                "entity_id": str(raw.get("entity_id", "") or "").strip(),
                "summary": str(raw.get("summary", "") or "").strip(),
            }
            if query and not any(query in str(value).lower() for value in row.values()):
                continue
            rows.append(row)
            if len(rows) >= max(1, int(limit or 500)):
                break
        return rows
