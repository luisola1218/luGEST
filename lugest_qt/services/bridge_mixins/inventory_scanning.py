from __future__ import annotations

from typing import Any


class InventoryScanningBackendMixin:
    """Legacy adapter for inventory scanning; see BACKEND_GUIDE.md."""

    def inventory_scan_code(self, entity_type: str, entity_id: Any) -> str:
        kind = str(entity_type or "").strip().upper()
        identifier = str(entity_id or "").strip().upper()
        if kind in {"MATERIAL", "MATERIA", "MATERIA_PRIMA", "MP"}:
            kind = "MAT"
        elif kind in {"PRODUCT", "PRODUTO", "PRODUCTS"}:
            kind = "PRD"
        if kind not in {"MAT", "PRD"} or not identifier:
            raise ValueError("Tipo ou identificador de picagem invalido.")
        return f"{kind}|{identifier}"

    def _sync_inventory_scan_mysql(self, entries: list[tuple[str, str, str]]) -> bool:
        connect = getattr(self.desktop_main, "_mysql_connect", None)
        if not callable(connect):
            return False
        conn = None
        try:
            conn = connect()
            with conn.cursor() as cur:
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS inventory_scan_codes (
                        scan_code VARCHAR(96) PRIMARY KEY,
                        entity_type VARCHAR(16) NOT NULL,
                        entity_id VARCHAR(80) NOT NULL,
                        updated_at DATETIME NULL,
                        UNIQUE KEY uq_inventory_scan_entity (entity_type, entity_id)
                    )
                    """
                )
                cur.execute("SELECT scan_code, entity_type, entity_id FROM inventory_scan_codes WHERE entity_type IN ('MAT', 'PRD')")
                existing_rows = list(cur.fetchall() or [])
                existing = {
                    str(row.get("scan_code", "") if isinstance(row, dict) else row[0]): (
                        str(row.get("entity_type", "") if isinstance(row, dict) else row[1]),
                        str(row.get("entity_id", "") if isinstance(row, dict) else row[2]),
                    )
                    for row in existing_rows
                }
                for scan_code, entity_type, entity_id in entries:
                    if existing.get(scan_code) == (entity_type, entity_id):
                        continue
                    cur.execute(
                        """
                        INSERT INTO inventory_scan_codes (scan_code, entity_type, entity_id, updated_at)
                        VALUES (%s, %s, %s, NOW())
                        ON DUPLICATE KEY UPDATE
                            scan_code=VALUES(scan_code),
                            updated_at=VALUES(updated_at)
                        """,
                        (scan_code, entity_type, entity_id),
                    )
                current_codes = {scan_code for scan_code, _entity_type, _entity_id in entries}
                stale_codes = [
                    scan_code for scan_code in existing if scan_code not in current_codes
                ]
                if stale_codes:
                    cur.executemany("DELETE FROM inventory_scan_codes WHERE scan_code=%s", [(code,) for code in stale_codes])
            conn.commit()
            return True
        except Exception:
            return False
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass

    def _inventory_scan_mysql_lookup(self, scan_code: str) -> dict[str, str]:
        connect = getattr(self.desktop_main, "_mysql_connect", None)
        if not callable(connect):
            return {}
        conn = None
        try:
            conn = connect()
            with conn.cursor() as cur:
                cur.execute(
                    "SELECT scan_code, entity_type, entity_id FROM inventory_scan_codes WHERE scan_code=%s LIMIT 1",
                    (str(scan_code or "").strip().upper(),),
                )
                row = cur.fetchone()
            if not row:
                return {}
            if isinstance(row, dict):
                return {key: str(row.get(key, "") or "") for key in ("scan_code", "entity_type", "entity_id")}
            return {"scan_code": str(row[0] or ""), "entity_type": str(row[1] or ""), "entity_id": str(row[2] or "")}
        except Exception:
            return {}
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass

    def ensure_inventory_scan_codes(self, *, persist: bool = True) -> dict[str, Any]:
        data = self.ensure_data()
        entries: list[tuple[str, str, str]] = []
        for bucket, id_field, kind in (("materiais", "id", "MAT"), ("produtos", "codigo", "PRD")):
            for record in list(data.get(bucket, []) or []):
                if not isinstance(record, dict):
                    continue
                entity_id = str(record.get(id_field, "") or "").strip().upper()
                if not entity_id:
                    continue
                scan_code = self.inventory_scan_code(kind, entity_id)
                entries.append((scan_code, kind, entity_id))
        mysql_synced = self._sync_inventory_scan_mysql(entries) if persist else False
        return {"changed": False, "entries": len(entries), "mysql_synced": mysql_synced}

    def inventory_scan_lookup(self, value: Any, expected_type: str = "") -> dict[str, Any]:
        raw = str(value or "").strip().upper()
        if not raw:
            raise ValueError("Codigo de picagem vazio.")
        if raw.startswith("LUG|"):
            raw = raw[4:]
        expected = str(expected_type or "").strip().upper()
        if expected in {"MATERIAL", "MATERIA", "MATERIA_PRIMA", "MP"}:
            expected = "MAT"
        elif expected in {"PRODUCT", "PRODUTO", "PRODUCTS"}:
            expected = "PRD"

        kind = ""
        entity_id = ""
        mysql_row = self._inventory_scan_mysql_lookup(raw)
        if mysql_row:
            kind = str(mysql_row.get("entity_type", "") or "").strip().upper()
            entity_id = str(mysql_row.get("entity_id", "") or "").strip().upper()
        parts = raw.split("|", 1)
        if not kind and len(parts) == 2 and parts[0] in {"MAT", "PRD"}:
            kind, entity_id = parts[0], parts[1].strip()
        elif not kind:
            index_row = dict((self.ensure_data().get("inventory_scan_index", {}) or {}).get(raw, {}) or {})
            kind = str(index_row.get("entity_type", "") or "").strip().upper()
            entity_id = str(index_row.get("entity_id", "") or "").strip().upper()
            if not kind:
                if raw.startswith("MAT"):
                    kind, entity_id = "MAT", raw
                elif raw.startswith("PRD"):
                    kind, entity_id = "PRD", raw
        if kind not in {"MAT", "PRD"} or not entity_id:
            raise ValueError(f"Codigo de picagem desconhecido: {raw}")
        if expected in {"MAT", "PRD"} and kind != expected:
            label = "materia-prima" if expected == "MAT" else "produto"
            raise ValueError(f"O codigo lido nao pertence ao stock de {label}.")
        if kind == "MAT":
            record = self.material_by_id(entity_id)
        else:
            record = next(
                (row for row in list(self.ensure_data().get("produtos", []) or []) if str(row.get("codigo", "") or "").strip().upper() == entity_id),
                None,
            )
        if record is None:
            raise ValueError(f"Registo nao encontrado para o codigo {raw}.")
        return {
            "entity_type": kind,
            "entity_id": entity_id,
            "scan_code": self.inventory_scan_code(kind, entity_id),
            "record": record,
        }
