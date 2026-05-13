from __future__ import annotations

import argparse
import hashlib
import os
import socket
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
ROOT = BASE_DIR.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from install_lugest_mysql import connect_mysql, execute_sql_batch, split_sql_statements, validate_schema
from mysql_tooling import env_int, load_env_candidates


load_env_candidates(
    [
        ROOT / "lugest.env",
        ROOT.parent / "Desktop App" / "lugest.env",
    ]
)


MIGRATION_FILES = sorted([*BASE_DIR.glob("patch_*.sql"), *(BASE_DIR / "Migracoes").glob("patch_*.sql")])
MIGRATION_TABLE_SQL = """
CREATE TABLE IF NOT EXISTS `schema_migrations` (
  `id` BIGINT NOT NULL AUTO_INCREMENT,
  `migration_key` VARCHAR(191) NOT NULL,
  `checksum_sha256` VARCHAR(64) NOT NULL,
  `applied_at` DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
  `applied_by` VARCHAR(120) NOT NULL DEFAULT '',
  `notes` TEXT NULL,
  PRIMARY KEY (`id`),
  UNIQUE KEY `uq_schema_migrations_key` (`migration_key`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
""".strip()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Aplica e versiona migracoes MySQL do LuGEST em clientes em producao."
    )
    parser.add_argument("--host", default=str(os.environ.get("LUGEST_MYSQL_ADMIN_HOST", os.environ.get("LUGEST_DB_HOST", "127.0.0.1")) or "127.0.0.1"))
    parser.add_argument("--port", type=int, default=env_int("LUGEST_MYSQL_ADMIN_PORT", env_int("LUGEST_DB_PORT", 3306)))
    parser.add_argument("--admin-user", default=str(os.environ.get("LUGEST_MYSQL_ADMIN_USER", os.environ.get("LUGEST_DB_USER", "root")) or "root"))
    parser.add_argument("--admin-password", default=str(os.environ.get("LUGEST_MYSQL_ADMIN_PASSWORD", os.environ.get("LUGEST_DB_PASS", "")) or ""))
    parser.add_argument("--database", default=str(os.environ.get("LUGEST_DB_NAME", "lugest") or "lugest"))
    parser.add_argument("--status", action="store_true", help="Mostra o estado das migracoes e termina.")
    parser.add_argument(
        "--baseline-current",
        action="store_true",
        help="Regista todas as migracoes atuais como ja aplicadas, sem executar SQL.",
    )
    parser.add_argument(
        "--legacy-apply-all",
        action="store_true",
        help="Aplica todos os patch_*.sql a uma base antiga sem historico de migracoes e regista o resultado.",
    )
    parser.add_argument("--skip-validation", action="store_true", help="Nao valida o schema final.")
    parser.add_argument("--actor", default="", help="Identificador de quem aplicou a migracao.")
    parser.add_argument("--dry-run", action="store_true", help="Mostra o plano sem alterar a base.")
    return parser.parse_args()


def default_actor() -> str:
    username = str(os.environ.get("USERNAME", os.environ.get("USER", "utilizador")) or "utilizador").strip() or "utilizador"
    machine = str(os.environ.get("COMPUTERNAME", socket.gethostname()) or "posto").strip() or "posto"
    return f"{machine}\\{username}"


def sha256_file(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def ensure_migration_table(cur, *, dry_run: bool) -> None:
    if dry_run:
        print("[DRY-RUN] Garantir tabela schema_migrations")
        return
    cur.execute(MIGRATION_TABLE_SQL)


def fetch_applied_migrations(cur) -> dict[str, dict[str, str]]:
    cur.execute(
        """
        SELECT migration_key, checksum_sha256, applied_at, applied_by, notes
        FROM schema_migrations
        ORDER BY migration_key
        """
    )
    applied: dict[str, dict[str, str]] = {}
    for row in cur.fetchall() or []:
        if isinstance(row, dict):
            key = str(row.get("migration_key") or "").strip()
            checksum = str(row.get("checksum_sha256") or "").strip()
            applied_at = str(row.get("applied_at") or "").strip()
            applied_by = str(row.get("applied_by") or "").strip()
            notes = str(row.get("notes") or "").strip()
        else:
            key = str(row[0] or "").strip()
            checksum = str(row[1] or "").strip()
            applied_at = str(row[2] or "").strip()
            applied_by = str(row[3] or "").strip()
            notes = str(row[4] or "").strip()
        if key:
            applied[key] = {
                "checksum": checksum,
                "applied_at": applied_at,
                "applied_by": applied_by,
                "notes": notes,
            }
    return applied


def record_migration(cur, *, migration_key: str, checksum: str, actor: str, notes: str, dry_run: bool) -> None:
    if dry_run:
        print(f"[DRY-RUN] Registar migracao {migration_key} ({notes})")
        return
    cur.execute(
        """
        INSERT INTO schema_migrations (migration_key, checksum_sha256, applied_by, notes)
        VALUES (%s, %s, %s, %s)
        ON DUPLICATE KEY UPDATE
            checksum_sha256 = VALUES(checksum_sha256),
            applied_by = VALUES(applied_by),
            notes = VALUES(notes)
        """,
        (migration_key, checksum, actor, notes),
    )


def table_exists(cur, table_name: str) -> bool:
    cur.execute(
        """
        SELECT COUNT(*)
        FROM information_schema.tables
        WHERE table_schema = DATABASE() AND table_name = %s
        """,
        (table_name,),
    )
    row = cur.fetchone()
    if isinstance(row, dict):
        try:
            return int(next(iter(row.values()))) > 0
        except Exception:
            return False
    if row:
        try:
            return int(row[0]) > 0
        except Exception:
            return False
    return False


def count_business_tables(cur) -> int:
    cur.execute(
        """
        SELECT COUNT(*)
        FROM information_schema.tables
        WHERE table_schema = DATABASE()
          AND table_type = 'BASE TABLE'
          AND table_name <> 'schema_migrations'
        """
    )
    row = cur.fetchone()
    if isinstance(row, dict):
        try:
            return int(next(iter(row.values())))
        except Exception:
            return 0
    if row:
        try:
            return int(row[0])
        except Exception:
            return 0
    return 0


def print_status(cur) -> int:
    tracked = table_exists(cur, "schema_migrations")
    print("== Estado das migracoes LuGEST ==")
    print(f"Tabela schema_migrations: {'OK' if tracked else 'EM FALTA'}")
    if not tracked:
        print(f"Migracoes conhecidas....: {len(MIGRATION_FILES)}")
        print("Migracoes aplicadas....: 0")
        print("Migracoes pendentes....: " + ", ".join(path.name for path in MIGRATION_FILES))
        return 0

    applied = fetch_applied_migrations(cur)
    pending = [path.name for path in MIGRATION_FILES if path.name not in applied]
    print(f"Migracoes conhecidas....: {len(MIGRATION_FILES)}")
    print(f"Migracoes aplicadas....: {len(applied)}")
    print(f"Migracoes pendentes....: {len(pending)}")
    if pending:
        print("Lista pendente.........: " + ", ".join(pending))
    checksum_drift: list[str] = []
    for path in MIGRATION_FILES:
        entry = applied.get(path.name)
        if not entry:
            continue
        current_checksum = sha256_file(path)
        if entry.get("checksum") and entry["checksum"] != current_checksum:
            checksum_drift.append(path.name)
    if checksum_drift:
        print("Atencao checksum drift.: " + ", ".join(checksum_drift))
        return 1
    print("Estado.................: OK")
    return 0


def baseline_current(cur, *, actor: str, dry_run: bool) -> tuple[int, int]:
    ensure_migration_table(cur, dry_run=dry_run)
    applied = fetch_applied_migrations(cur) if not dry_run else {}
    added = 0
    skipped = 0
    for path in MIGRATION_FILES:
        checksum = sha256_file(path)
        current = applied.get(path.name)
        if current:
            if current.get("checksum") and current["checksum"] != checksum:
                raise RuntimeError(
                    f"Migracao {path.name} ja registada com checksum diferente. "
                    "Revê a pasta de scripts antes de continuar."
                )
            skipped += 1
            continue
        record_migration(
            cur,
            migration_key=path.name,
            checksum=checksum,
            actor=actor,
            notes="baseline-current",
            dry_run=dry_run,
        )
        added += 1
    return added, skipped


def apply_pending(cur, *, actor: str, dry_run: bool) -> tuple[int, int, int]:
    ensure_migration_table(cur, dry_run=dry_run)
    applied = fetch_applied_migrations(cur) if not dry_run else {}
    if not applied and count_business_tables(cur) > 0:
        raise RuntimeError(
            "Base sem historico de migracoes. Usa --baseline-current se a base ja estiver alinhada "
            "ou --legacy-apply-all se precisas aplicar patches antigos."
        )

    applied_count = 0
    skipped_count = 0
    statements_count = 0
    for path in MIGRATION_FILES:
        checksum = sha256_file(path)
        current = applied.get(path.name)
        if current:
            if current.get("checksum") and current["checksum"] != checksum:
                raise RuntimeError(
                    f"Migracao {path.name} ja registada com checksum diferente. "
                    "Nao e seguro continuar sem rever os scripts."
                )
            skipped_count += 1
            continue

        statements = split_sql_statements(path.read_text(encoding="utf-8"))
        executed, skipped = execute_sql_batch(
            cur,
            statements,
            tolerant=True,
            dry_run=dry_run,
            label=path.name,
        )
        record_migration(
            cur,
            migration_key=path.name,
            checksum=checksum,
            actor=actor,
            notes="apply-pending",
            dry_run=dry_run,
        )
        applied_count += 1
        statements_count += executed
        skipped_count += skipped
    return applied_count, statements_count, skipped_count


def legacy_apply_all(cur, *, actor: str, dry_run: bool) -> tuple[int, int, int]:
    ensure_migration_table(cur, dry_run=dry_run)
    applied = fetch_applied_migrations(cur) if not dry_run else {}
    applied_count = 0
    statements_count = 0
    skipped_count = 0
    for path in MIGRATION_FILES:
        checksum = sha256_file(path)
        current = applied.get(path.name)
        if current:
            if current.get("checksum") and current["checksum"] != checksum:
                raise RuntimeError(
                    f"Migracao {path.name} ja registada com checksum diferente. "
                    "Nao e seguro continuar sem rever os scripts."
                )
            continue
        statements = split_sql_statements(path.read_text(encoding="utf-8"))
        executed, skipped = execute_sql_batch(
            cur,
            statements,
            tolerant=True,
            dry_run=dry_run,
            label=path.name,
        )
        record_migration(
            cur,
            migration_key=path.name,
            checksum=checksum,
            actor=actor,
            notes="legacy-apply-all",
            dry_run=dry_run,
        )
        applied_count += 1
        statements_count += executed
        skipped_count += skipped
    return applied_count, statements_count, skipped_count


def main() -> int:
    args = parse_args()
    actor = str(args.actor or "").strip() or default_actor()
    selected_modes = sum(bool(flag) for flag in (args.status, args.baseline_current, args.legacy_apply_all))
    if selected_modes > 1:
        print("Escolhe apenas um destes modos: --status, --baseline-current ou --legacy-apply-all.")
        return 2

    if not args.admin_user or not args.admin_password:
        print("Define --admin-user e --admin-password ou configura LUGEST_MYSQL_ADMIN_USER/LUGEST_MYSQL_ADMIN_PASSWORD.")
        return 2

    try:
        import pymysql
    except Exception as exc:
        print(f"PyMySQL nao esta instalado neste Python: {exc}")
        return 2

    print("== LuGEST MySQL Migrations ==")
    print(f"Servidor.....: {args.host}:{args.port}")
    print(f"Base alvo....: {args.database}")
    print(f"Actor........: {actor}")
    print(f"Migracoes....: {len(MIGRATION_FILES)} conhecidas")
    if args.dry_run:
        print("DryRun.......: ativo")

    conn = None
    try:
        conn = connect_mysql(pymysql, args.host, args.port, args.admin_user, args.admin_password, args.database)
        with conn.cursor() as cur:
            if args.status:
                result = print_status(cur)
                conn.rollback()
                return result

            if args.baseline_current:
                added, skipped = baseline_current(cur, actor=actor, dry_run=args.dry_run)
                if not args.dry_run:
                    conn.commit()
                print(f"baseline-current: registadas={added} existentes={skipped}")
            elif args.legacy_apply_all:
                applied, statements, skipped = legacy_apply_all(cur, actor=actor, dry_run=args.dry_run)
                if not args.dry_run:
                    conn.commit()
                print(f"legacy-apply-all: migracoes={applied} statements={statements} skipped={skipped}")
            else:
                applied, statements, skipped = apply_pending(cur, actor=actor, dry_run=args.dry_run)
                if not args.dry_run:
                    conn.commit()
                print(f"apply-pending: migracoes={applied} statements={statements} skipped={skipped}")

            if not args.skip_validation:
                if args.dry_run:
                    print("[DRY-RUN] Validacao final do schema")
                else:
                    issues = validate_schema(cur)
                    if issues:
                        print("mysql-migrations-validate-failed")
                        for issue in issues:
                            print(f"- {issue}")
                        return 1
                    print("mysql-migrations-validate-ok")
    except Exception as exc:
        if conn is not None and not args.dry_run:
            try:
                conn.rollback()
            except Exception:
                pass
        print(f"mysql-migrations-failed: {exc}")
        return 1
    finally:
        if conn is not None:
            try:
                conn.close()
            except Exception:
                pass

    print("mysql-migrations-ok")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
