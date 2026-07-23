from __future__ import annotations

import argparse
import copy
import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_qt.services.legacy_backend import LegacyBackend


DOCUMENT_LIST_KEYS = (
    "orcamentos",
    "notas_encomenda",
    "expedicoes",
    "faturacao",
    "faturacao_registos",
    "faturas",
    "pagamentos",
    "plano",
    "planeamento_blocos",
    "avarias",
    "refs",
)


def _serialized(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, default=str).upper()


def _is_verification_record(value: Any, order_numbers: set[str] | None = None) -> bool:
    raw = _serialized(value)
    if "VERIFY" in raw:
        return True
    return any(number and number.upper() in raw for number in set(order_numbers or set()))


def _backup_path() -> Path:
    desktop = Path.home() / "Desktop"
    release_dir = next(iter(sorted(desktop.glob("App LuisGEST - Revis*o Final"))), None)
    target_dir = (release_dir / "Base de Dados" / "Backups") if release_dir else (ROOT / "backups")
    target_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return target_dir / f"antes_limpeza_verify_{stamp}.json"


def _summary(data: dict[str, Any]) -> dict[str, int]:
    orders = [
        row
        for row in list(data.get("encomendas", []) or [])
        if _is_verification_record(row)
    ]
    order_numbers = {
        str(row.get("numero", "") or "").strip()
        for row in orders
        if str(row.get("numero", "") or "").strip()
    }
    existing_orders = {
        str(row.get("numero", "") or "").strip()
        for row in list(data.get("encomendas", []) or [])
        if isinstance(row, dict) and str(row.get("numero", "") or "").strip()
    }
    result = {"encomendas": len(orders)}
    for key in DOCUMENT_LIST_KEYS:
        result[key] = sum(
            1
            for row in list(data.get(key, []) or [])
            if _is_verification_record(row, order_numbers)
            or (
                key in {"plano", "planeamento_blocos"}
                and str((row or {}).get("encomenda", "") or "").strip()
                and str((row or {}).get("encomenda", "") or "").strip() not in existing_orders
            )
        )
    return result


def cleanup(*, write: bool) -> dict[str, Any]:
    backend = LegacyBackend()
    data = backend.ensure_data()
    before = _summary(data)
    if not write:
        return {"modo": "preview", "antes": before}

    backup = _backup_path()
    backup.write_text(
        json.dumps(copy.deepcopy(data), ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )
    backend.flush_pending_save(force=True)
    backend.drain_async_saves(timeout_sec=20.0)

    verification_orders = [
        row
        for row in list(data.get("encomendas", []) or [])
        if _is_verification_record(row)
    ]
    order_numbers = {
        str(row.get("numero", "") or "").strip()
        for row in verification_orders
        if str(row.get("numero", "") or "").strip()
    }
    existing_orders = {
        str(row.get("numero", "") or "").strip()
        for row in list(data.get("encomendas", []) or [])
        if isinstance(row, dict) and str(row.get("numero", "") or "").strip()
    }

    for key in DOCUMENT_LIST_KEYS:
        data[key] = [
            row
            for row in list(data.get(key, []) or [])
            if not (
                _is_verification_record(row, order_numbers)
                or (
                    key in {"plano", "planeamento_blocos"}
                    and str((row or {}).get("encomenda", "") or "").strip()
                    and str((row or {}).get("encomenda", "") or "").strip() not in existing_orders
                )
            )
        ]

    for number in sorted(order_numbers):
        try:
            backend.order_remove(number)
        except Exception:
            current = backend.ensure_data()
            current["encomendas"] = [
                row
                for row in list(current.get("encomendas", []) or [])
                if str(row.get("numero", "") or "").strip() != number
            ]

    data = backend.ensure_data()
    existing_orders = {
        str(row.get("numero", "") or "").strip()
        for row in list(data.get("encomendas", []) or [])
        if isinstance(row, dict) and str(row.get("numero", "") or "").strip()
    }
    for key in DOCUMENT_LIST_KEYS:
        data[key] = [
            row
            for row in list(data.get(key, []) or [])
            if not (
                _is_verification_record(row, order_numbers)
                or (
                    key in {"plano", "planeamento_blocos"}
                    and str((row or {}).get("encomenda", "") or "").strip()
                    and str((row or {}).get("encomenda", "") or "").strip() not in existing_orders
                )
            )
        ]
    data["encomendas"] = [
        row
        for row in list(data.get("encomendas", []) or [])
        if str(row.get("numero", "") or "").strip() not in order_numbers
    ]
    backend._save(force=True, audit=False, blocking=True)
    backend.drain_async_saves(timeout_sec=20.0)
    backend.reload(force=True)
    after = _summary(backend.ensure_data())
    if any(after.values()):
        raise RuntimeError(f"A limpeza deixou artefactos de verificacao: {after}")
    return {"modo": "write", "backup": str(backup), "antes": before, "depois": after}


def main() -> int:
    parser = argparse.ArgumentParser(description="Remove apenas artefactos tecnicos VERIFY da base LuisGEST.")
    parser.add_argument("--write", action="store_true", help="Aplica a limpeza. Sem esta opcao apenas mostra a previsao.")
    args = parser.parse_args()
    print(json.dumps(cleanup(write=bool(args.write)), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
