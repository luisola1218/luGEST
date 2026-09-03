from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from mysql.export_current_schema_sql import TARGET_CHARSET, TARGET_COLLATION, _show_create_table


class _Cursor:
    def execute(self, _sql: str) -> None:
        return None

    def fetchone(self):
        return {
            "Table": "legacy_table",
            "Create Table": "CREATE TABLE `legacy_table` (`texto` text) ENGINE=InnoDB DEFAULT CHARSET=utf8",
        }


def main() -> int:
    schema_path = ROOT / "mysql" / "lugest.sql"
    schema = schema_path.read_text(encoding="utf-8-sig")
    assert not re.search(r"\b(?:DEFAULT )?CHARSET=utf8(?:mb3)?\b", schema, flags=re.IGNORECASE)
    assert not re.search(r"\bCHARACTER SET utf8(?:mb3)?\b", schema, flags=re.IGNORECASE)
    assert not re.search(r"\bCOLLATE[= ]+utf8(?:mb3)?_", schema, flags=re.IGNORECASE)
    table_count = len(re.findall(r"^CREATE TABLE", schema, flags=re.IGNORECASE | re.MULTILINE))
    assert table_count >= 50
    assert schema.count(f"DEFAULT CHARSET={TARGET_CHARSET}") == table_count

    normalized = _show_create_table(_Cursor(), "legacy_table")
    assert f"DEFAULT CHARSET={TARGET_CHARSET}" in normalized
    assert f"COLLATE={TARGET_COLLATION}" in normalized
    assert "DEFAULT CHARSET=utf8;" not in normalized
    print(f"mysql-schema-definition-ok tables={table_count} charset={TARGET_CHARSET}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
