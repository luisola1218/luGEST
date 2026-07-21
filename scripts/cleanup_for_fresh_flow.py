from __future__ import annotations

import argparse
import json
import sys
from copy import deepcopy
from datetime import datetime
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_qt.services.legacy_backend import LegacyBackend


LIST_BUCKETS_TO_CLEAR = (
    "audit_log",
    "clientes",
    "encomendas",
    "expedicoes",
    "faturacao",
    "faturacao_registos",
    "fornecedores",
    "notas_encomenda",
    "op_eventos",
    "op_paragens",
    "orcamentos",
    "plano",
    "plano_bloqueios",
    "plano_hist",
    "produtos_mov",
    "qualidade",
    "quality_documents",
    "quality_nonconformities",
    "rejeitadas_hist",
    "stock_log",
    "transportes",
)

DICT_BUCKETS_TO_CLEAR = (
    "orc_refs",
    "peca_hist",
    "refs",
)

QUALITY_PREFIXES = ("inspection_", "quality_")
QUALITY_FIELDS = {
    "logistic_status",
    "origem_encomenda",
    "supplier_claim_id",
}


def _snapshot(data: dict[str, Any]) -> dict[str, Any]:
    return json.loads(json.dumps(data, ensure_ascii=False))


def _counts(data: dict[str, Any]) -> dict[str, int]:
    return {
        key: len(value)
        for key, value in sorted(data.items())
        if isinstance(value, (list, dict))
    }


def _clean_stock_records(rows: list[dict[str, Any]], quantity_key: str) -> None:
    for row in rows:
        if not isinstance(row, dict):
            continue
        row["reservado"] = 0.0
        for key in list(row):
            key_txt = str(key or "")
            if key_txt.startswith(QUALITY_PREFIXES) or key_txt in QUALITY_FIELDS:
                row.pop(key, None)
        # Keep the physical quantity and all commercial/technical master data untouched.
        row[quantity_key] = row.get(quantity_key, 0)


def _backup_path() -> Path:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    release_backup = Path.home() / "Desktop" / "App LuisGEST - Revisão Final" / "Base de Dados" / "Backups"
    release_backup.mkdir(parents=True, exist_ok=True)
    return release_backup / f"antes_limpeza_fluxo_{stamp}.json"


def main() -> int:
    parser = argparse.ArgumentParser(description="Limpa transações e preserva stock/conjuntos para um fluxo novo.")
    parser.add_argument("--apply", action="store_true", help="Grava a limpeza. Sem esta opção apenas apresenta a previsão.")
    args = parser.parse_args()

    backend = LegacyBackend()
    current = _snapshot(backend.ensure_data())
    before = _counts(current)
    preserved = {
        "materiais": len(list(current.get("materiais", []) or [])),
        "produtos": len(list(current.get("produtos", []) or [])),
        "conjuntos": len(list(current.get("conjuntos", []) or [])),
        "conjuntos_modelo": len(list(current.get("conjuntos_modelo", []) or [])),
        "users": len(list(current.get("users", []) or [])),
    }

    cleaned = deepcopy(current)
    for key in LIST_BUCKETS_TO_CLEAR:
        cleaned[key] = []
    for key in DICT_BUCKETS_TO_CLEAR:
        cleaned[key] = {}

    _clean_stock_records(list(cleaned.get("materiais", []) or []), "quantidade")
    _clean_stock_records(list(cleaned.get("produtos", []) or []), "qty")

    seq = dict(cleaned.get("seq", {}) or {})
    seq["encomenda"] = 1
    seq["cliente"] = 1
    seq["ne"] = 1
    seq["fornecedor"] = 1
    seq["ref_interna"] = {}
    cleaned["seq"] = seq
    cleaned["orc_seq"] = 1
    cleaned["of_seq"] = 1
    cleaned["opp_seq"] = 1
    cleaned["exp_seq"] = 1

    preview = {
        "modo": "aplicar" if args.apply else "previsao",
        "preservado": preserved,
        "remover": {
            key: before.get(key, 0)
            for key in (*LIST_BUCKETS_TO_CLEAR, *DICT_BUCKETS_TO_CLEAR)
            if before.get(key, 0)
        },
    }
    if not args.apply:
        print(json.dumps(preview, ensure_ascii=False, indent=2))
        return 0

    backup = _backup_path()
    backup.write_text(json.dumps(current, ensure_ascii=False, indent=2), encoding="utf-8")
    backend.data = cleaned
    # Maintenance operations must bypass the normal asynchronous save queue.
    backend._save(force=True, audit=False, blocking=True)

    check_backend = LegacyBackend()
    check_backend.reload(force=True)
    persisted = check_backend.ensure_data()
    after = _counts(persisted)
    for key, expected in preserved.items():
        if after.get(key, 0) != expected:
            raise RuntimeError(f"Falha de preservação em {key}: esperado {expected}, obtido {after.get(key, 0)}")
    for key in (*LIST_BUCKETS_TO_CLEAR, *DICT_BUCKETS_TO_CLEAR):
        if after.get(key, 0) != 0:
            raise RuntimeError(f"A limpeza não persistiu em {key}: {after.get(key, 0)} registos")
    if any(float(row.get("reservado", 0) or 0) != 0 for row in list(persisted.get("materiais", []) or [])):
        raise RuntimeError("Ainda existem reservas de matéria-prima.")
    if any(float(row.get("reservado", 0) or 0) != 0 for row in list(persisted.get("produtos", []) or [])):
        raise RuntimeError("Ainda existem reservas de produto.")

    print(json.dumps({**preview, "backup": str(backup), "resultado": _counts(persisted)}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
