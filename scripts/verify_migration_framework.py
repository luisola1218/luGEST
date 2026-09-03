from __future__ import annotations

import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from mysql.migrate_lugest_mysql import discover_migrations, pending_migrations


def main() -> int:
    with tempfile.TemporaryDirectory(prefix="lugest-migrations-") as tmp:
        root = Path(tmp)
        (root / "0001_primeira.sql").write_text("CREATE TABLE exemplo (id INT);\n", encoding="utf-8")
        (root / "0002_indice.sql").write_text("CREATE INDEX idx_exemplo ON exemplo (id);\n", encoding="utf-8")
        migrations = discover_migrations(root)
        assert [item.key for item in migrations] == ["0001_primeira", "0002_indice"]
        assert len(pending_migrations(migrations, {})) == 2
        assert len(pending_migrations(migrations, {migrations[0].key: migrations[0].checksum_sha256})) == 1
        try:
            pending_migrations(migrations, {migrations[0].key: "0" * 64})
        except RuntimeError:
            pass
        else:
            raise AssertionError("Uma migração aplicada e alterada não foi bloqueada.")

        (root / "3_nome_invalido.sql").write_text("SELECT 1;", encoding="utf-8")
        try:
            discover_migrations(root)
        except ValueError:
            pass
        else:
            raise AssertionError("Um nome de migração inválido foi aceite.")

    print("migration-framework-ok ordering=yes checksum=locked naming=validated")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
