from __future__ import annotations

import copy
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from lugest_qt.services.legacy_backend import LegacyBackend


def main() -> int:
    backend = LegacyBackend()
    live_data = backend.ensure_data()
    backend._replace_data_cache(copy.deepcopy(live_data))
    backend._save = lambda *args, **kwargs: None
    backend.ensure_inventory_scan_codes = lambda **kwargs: {"changed": False, "entries": 0, "mysql_synced": False}

    material_ids = [
        str(row.get("id", "") or "").strip()
        for row in backend.ensure_data().get("materiais", [])
        if str(row.get("id", "") or "").strip()
    ][:2]
    product_codes = [
        str(row.get("codigo", "") or "").strip()
        for row in backend.ensure_data().get("produtos", [])
        if str(row.get("codigo", "") or "").strip()
    ][:2]
    assert len(material_ids) == 2 and len(product_codes) == 2

    material_total = len(backend.ensure_data().get("materiais", []))
    product_total = len(backend.ensure_data().get("produtos", []))
    assert backend.remove_materials(material_ids) == 2
    assert backend.product_remove_many(product_codes) == 2

    remaining_materials = {
        str(row.get("id", "") or "").strip() for row in backend.ensure_data().get("materiais", [])
    }
    remaining_products = {
        str(row.get("codigo", "") or "").strip() for row in backend.ensure_data().get("produtos", [])
    }
    assert not set(material_ids) & remaining_materials
    assert not set(product_codes) & remaining_products
    assert len(remaining_materials) == material_total - 2
    assert len(remaining_products) == product_total - 2
    print("inventory-bulk-delete-ok materials=2 products=2 live-data-untouched")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
