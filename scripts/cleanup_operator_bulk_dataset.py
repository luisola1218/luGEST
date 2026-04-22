from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_qt.services.main_bridge import LegacyBackend


MANIFEST_PATH = ROOT / "generated" / "seeds" / "operator_bulk_manifest.json"


def _load_manifest() -> dict[str, object] | None:
    if not MANIFEST_PATH.exists():
        return None
    try:
        return json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
    except Exception:
        return None


def main() -> int:
    manifest = _load_manifest()
    if not manifest:
        print("cleanup: sem manifest para remover")
        return 0

    backend = LegacyBackend()
    client_codes = {
        str(value).strip() for value in list(manifest.get("client_codes", []) or []) if str(value).strip()
    }
    quote_numbers = {
        str(value).strip() for value in list(manifest.get("quote_numbers", []) or []) if str(value).strip()
    }
    order_numbers = {
        str(value).strip() for value in list(manifest.get("order_numbers", []) or []) if str(value).strip()
    }
    ref_externas = {
        str(value).strip() for value in list(manifest.get("ref_externas", []) or []) if str(value).strip()
    }

    data = backend.ensure_data()
    data["plano"] = [
        row for row in list(data.get("plano", []) or [])
        if str(row.get("encomenda", row.get("encomenda_numero", "")) or "").strip() not in order_numbers
    ]
    data["plano_hist"] = [
        row for row in list(data.get("plano_hist", []) or [])
        if str(row.get("encomenda", row.get("encomenda_numero", "")) or "").strip() not in order_numbers
    ]
    data["expedicoes"] = [
        row for row in list(data.get("expedicoes", []) or [])
        if str(row.get("encomenda_numero", "") or "").strip() not in order_numbers
    ]
    data["encomendas"] = [
        row for row in list(data.get("encomendas", []) or [])
        if str(row.get("numero", "") or "").strip() not in order_numbers
    ]
    data["orcamentos"] = [
        row for row in list(data.get("orcamentos", []) or [])
        if str(row.get("numero", "") or "").strip() not in quote_numbers
    ]
    data["clientes"] = [
        row for row in list(data.get("clientes", []) or [])
        if str(row.get("codigo", "") or "").strip() not in client_codes
    ]

    peca_hist = dict(data.get("peca_hist", {}) or {})
    for ref in ref_externas:
        peca_hist.pop(ref, None)
    data["peca_hist"] = peca_hist

    orc_refs = dict(data.get("orc_refs", {}) or {})
    for ref in ref_externas:
        orc_refs.pop(ref, None)
    data["orc_refs"] = orc_refs

    data["faturacao_registos"] = [
        row for row in list(data.get("faturacao_registos", []) or [])
        if str(row.get("orcamento_numero", "") or "").strip() not in quote_numbers
        and str(row.get("encomenda_numero", "") or "").strip() not in order_numbers
    ]

    backend._save(force=True)
    if MANIFEST_PATH.exists():
        MANIFEST_PATH.unlink()

    print(
        json.dumps(
            {
                "clientes_removidos": len(client_codes),
                "orcamentos_removidos": len(quote_numbers),
                "encomendas_removidas": len(order_numbers),
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
