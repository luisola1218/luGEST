from __future__ import annotations

import argparse
import hashlib
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Mapping

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import main
from mysql.install_lugest_mysql import split_sql_statements


MIGRATIONS_DIR = Path(__file__).resolve().parent / "migrations"
MIGRATION_NAME = re.compile(r"^(\d{4})_([a-z0-9][a-z0-9_]*)\.sql$")
LOCK_NAME = "lugest_schema_migrations_v1"


@dataclass(frozen=True)
class Migration:
    key: str
    path: Path
    checksum_sha256: str
    sql: str


def discover_migrations(directory: Path = MIGRATIONS_DIR) -> list[Migration]:
    migrations: list[Migration] = []
    seen_numbers: set[str] = set()
    if not directory.exists():
        return migrations
    for path in sorted(directory.glob("*.sql"), key=lambda item: item.name.casefold()):
        match = MIGRATION_NAME.fullmatch(path.name)
        if not match:
            raise ValueError(f"Nome de migração inválido: {path.name}")
        sequence = match.group(1)
        if sequence in seen_numbers:
            raise ValueError(f"Número de migração duplicado: {sequence}")
        seen_numbers.add(sequence)
        raw = path.read_bytes()
        sql = raw.decode("utf-8-sig").strip()
        if not sql:
            raise ValueError(f"Migração vazia: {path.name}")
        migrations.append(
            Migration(
                key=path.stem,
                path=path,
                checksum_sha256=hashlib.sha256(raw).hexdigest(),
                sql=sql,
            )
        )
    return migrations


def pending_migrations(
    migrations: list[Migration],
    applied_checksums: Mapping[str, str],
) -> list[Migration]:
    pending: list[Migration] = []
    for migration in migrations:
        applied_checksum = str(applied_checksums.get(migration.key, "") or "").strip().lower()
        if not applied_checksum:
            pending.append(migration)
            continue
        if applied_checksum != migration.checksum_sha256.lower():
            raise RuntimeError(
                f"A migração já aplicada {migration.key} foi alterada. "
                "Cria uma nova migração em vez de editar o histórico."
            )
    return pending


def _migration_table_exists(cursor) -> bool:
    cursor.execute(
        """
        SELECT COUNT(*) AS n
        FROM information_schema.tables
        WHERE table_schema = DATABASE() AND table_name = 'schema_migrations'
        """
    )
    row = cursor.fetchone()
    value = row.get("n", 0) if isinstance(row, dict) else (row[0] if row else 0)
    return int(value or 0) > 0


def _applied_checksums(cursor) -> dict[str, str]:
    if not _migration_table_exists(cursor):
        return {}
    cursor.execute("SELECT migration_key, checksum_sha256 FROM schema_migrations ORDER BY id")
    result: dict[str, str] = {}
    for row in cursor.fetchall() or []:
        if isinstance(row, dict):
            key = str(row.get("migration_key", "") or "").strip()
            checksum = str(row.get("checksum_sha256", "") or "").strip()
        else:
            key = str(row[0] if row else "").strip()
            checksum = str(row[1] if row and len(row) > 1 else "").strip()
        if key:
            result[key] = checksum
    return result


def _ensure_migration_table(cursor) -> None:
    cursor.execute(
        """
        CREATE TABLE IF NOT EXISTS schema_migrations (
          id BIGINT NOT NULL AUTO_INCREMENT,
          migration_key VARCHAR(191) COLLATE utf8mb4_unicode_ci NOT NULL,
          checksum_sha256 VARCHAR(64) COLLATE utf8mb4_unicode_ci NOT NULL,
          applied_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
          applied_by VARCHAR(120) COLLATE utf8mb4_unicode_ci NOT NULL DEFAULT '',
          notes TEXT COLLATE utf8mb4_unicode_ci,
          PRIMARY KEY (id),
          UNIQUE KEY uq_schema_migrations_key (migration_key)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
        """
    )


def _apply_migrations(connection, migrations: list[Migration], applied_by: str) -> None:
    with connection.cursor() as cursor:
        cursor.execute("SELECT GET_LOCK(%s, 30) AS acquired", (LOCK_NAME,))
        row = cursor.fetchone()
        acquired = row.get("acquired", 0) if isinstance(row, dict) else (row[0] if row else 0)
        if int(acquired or 0) != 1:
            raise RuntimeError("Não foi possível obter o bloqueio exclusivo de migrações.")
        try:
            _ensure_migration_table(cursor)
            connection.commit()
            for migration in migrations:
                print(f"APPLY {migration.key}", flush=True)
                for statement in split_sql_statements(migration.sql):
                    cursor.execute(statement)
                cursor.execute(
                    """
                    INSERT INTO schema_migrations
                      (migration_key, checksum_sha256, applied_by, notes)
                    VALUES (%s, %s, %s, %s)
                    """,
                    (
                        migration.key,
                        migration.checksum_sha256,
                        str(applied_by or "").strip()[:120],
                        migration.path.name,
                    ),
                )
                connection.commit()
        except Exception:
            connection.rollback()
            raise
        finally:
            try:
                cursor.execute("SELECT RELEASE_LOCK(%s)", (LOCK_NAME,))
            except Exception:
                pass


def _validated_backup(path_value: str) -> Path:
    path = Path(str(path_value or "").strip()).expanduser()
    if not path.is_file() or path.stat().st_size <= 0:
        raise ValueError("--backup-confirmed tem de apontar para um backup existente e não vazio.")
    return path.resolve()


def main_entry(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Planeia ou aplica migrações versionadas do luGEST.")
    parser.add_argument("--apply", action="store_true", help="Aplica as migrações pendentes.")
    parser.add_argument("--backup-confirmed", default="", help="Backup validado obrigatório com --apply.")
    parser.add_argument("--applied-by", default="", help="Responsável registado no histórico.")
    parser.add_argument("--directory", type=Path, default=MIGRATIONS_DIR)
    args = parser.parse_args(argv)

    if args.apply:
        backup_path = _validated_backup(args.backup_confirmed)
        print(f"Backup confirmado: {backup_path}")

    migrations = discover_migrations(args.directory)
    connection = main._mysql_connect()
    try:
        with connection.cursor() as cursor:
            applied = _applied_checksums(cursor)
        pending = pending_migrations(migrations, applied)
        print(f"Migrações encontradas: {len(migrations)} | aplicadas: {len(migrations) - len(pending)} | pendentes: {len(pending)}")
        for migration in pending:
            print(f"PENDING {migration.key} {migration.checksum_sha256[:12]}")
        if not args.apply:
            print("migration-plan-ok mode=read-only")
            return 0
        if pending:
            _apply_migrations(connection, pending, args.applied_by)
        print(f"migration-apply-ok applied={len(pending)}")
        return 0
    finally:
        connection.close()


if __name__ == "__main__":
    raise SystemExit(main_entry())
