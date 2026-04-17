from __future__ import annotations

import argparse
import gzip
import os
import subprocess
import sys
import tempfile
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
ROOT = BASE_DIR.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from install_lugest_mysql import connect_mysql, validate_schema
from mysql_tooling import build_mysql_env, env_int, load_env_candidates, resolve_mysql_binary


load_env_candidates(
    [
        ROOT / "lugest.env",
        ROOT / "impulse_mobile_api" / ".env",
        ROOT.parent / "Desktop App" / "lugest.env",
        ROOT.parent / "Mobile API" / ".env",
        ROOT.parent / "impulse_mobile_api" / ".env",
    ]
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Restaura um dump MySQL do LuGEST.")
    parser.add_argument("--host", default=str(os.environ.get("LUGEST_MYSQL_ADMIN_HOST", os.environ.get("LUGEST_DB_HOST", "127.0.0.1")) or "127.0.0.1"))
    parser.add_argument("--port", type=int, default=env_int("LUGEST_MYSQL_ADMIN_PORT", env_int("LUGEST_DB_PORT", 3306)))
    parser.add_argument("--admin-user", default=str(os.environ.get("LUGEST_MYSQL_ADMIN_USER", os.environ.get("LUGEST_DB_USER", "")) or ""))
    parser.add_argument("--admin-password", default=str(os.environ.get("LUGEST_MYSQL_ADMIN_PASSWORD", os.environ.get("LUGEST_DB_PASS", "")) or ""))
    parser.add_argument("--database", default=str(os.environ.get("LUGEST_DB_NAME", "lugest") or "lugest"))
    parser.add_argument("--input", default="", help="Ficheiro .sql/.sql.gz ou pasta de backup.")
    parser.add_argument("--backup-root", default=str(ROOT / "backups" / "mysql"))
    parser.add_argument("--mysql-path", default="")
    parser.add_argument("--reset-database", action="store_true", help="Apaga e recria a base antes do restauro.")
    parser.add_argument("--skip-validation", action="store_true")
    parser.add_argument("--dry-run", action="store_true")
    return parser.parse_args()


def resolve_backup_input(raw_input: str, backup_root: Path) -> Path:
    direct = Path(str(raw_input or "").strip()).expanduser()
    if raw_input and direct.exists():
        if direct.is_file():
            return direct
        dump_files = sorted(list(direct.glob("*.sql")) + list(direct.glob("*.sql.gz")), key=lambda p: p.stat().st_mtime, reverse=True)
        if dump_files:
            return dump_files[0]
    if backup_root.exists():
        dump_files = sorted(backup_root.rglob("*.sql"), key=lambda p: p.stat().st_mtime, reverse=True)
        dump_files += sorted(backup_root.rglob("*.sql.gz"), key=lambda p: p.stat().st_mtime, reverse=True)
        if dump_files:
            return sorted(dump_files, key=lambda p: p.stat().st_mtime, reverse=True)[0]
    raise FileNotFoundError("Nao foi encontrado nenhum dump .sql ou .sql.gz para restaurar.")


def main() -> int:
    args = parse_args()
    mysql_bin = resolve_mysql_binary(args.mysql_path, "MYSQL_PATH", ["mysql.exe", "mysql"])
    backup_root = Path(args.backup_root).expanduser()

    dump_path: Path | None = None
    try:
        dump_path = resolve_backup_input(args.input, backup_root)
    except Exception as exc:
        if args.dry_run:
            print(f"[DRY-RUN] Dump indisponivel neste posto: {exc}")
        else:
            print(f"mysql-restore-failed: {exc}")
            return 2

    print("== LuGEST MySQL Restore ==")
    print(f"Servidor....: {args.host}:{args.port}")
    print(f"Base........: {args.database}")
    print(f"Dump........: {dump_path or '<nao encontrado neste posto>'}")
    if args.dry_run:
        print("DryRun......: ativo")

    if not mysql_bin:
        message = "mysql client nao foi encontrado. Define --mysql-path ou MYSQL_PATH."
        if args.dry_run:
            print(f"[DRY-RUN] {message}")
            return 0
        print(message)
        return 2

    if not args.admin_user or not args.admin_password:
        print("Define --admin-user e --admin-password ou configura LUGEST_MYSQL_ADMIN_USER/LUGEST_MYSQL_ADMIN_PASSWORD.")
        return 2

    print("Comando.....: mysql <credenciais ocultas> < dump")
    if args.dry_run:
        return 0

    if dump_path is None:
        print("mysql-restore-failed: Nao foi encontrado nenhum dump para restaurar.")
        return 2

    try:
        import pymysql
    except Exception as exc:
        print(f"mysql-restore-failed: PyMySQL nao disponivel: {exc}")
        return 2

    env = build_mysql_env(args.admin_password)
    temp_sql_path: Path | None = None

    try:
        server_conn = connect_mysql(pymysql, args.host, args.port, args.admin_user, args.admin_password, None)
        try:
            with server_conn.cursor() as cur:
                if args.reset_database:
                    cur.execute(f"DROP DATABASE IF EXISTS `{args.database}`")
                cur.execute(f"CREATE DATABASE IF NOT EXISTS `{args.database}`")
            server_conn.commit()
        finally:
            server_conn.close()

        source_path = dump_path
        if dump_path.suffix.lower() == ".gz":
            with gzip.open(dump_path, "rb") as source, tempfile.NamedTemporaryFile(delete=False, suffix=".sql") as tmp:
                tmp.write(source.read())
                temp_sql_path = Path(tmp.name)
            source_path = temp_sql_path

        cmd = [
            mysql_bin,
            f"--host={args.host}",
            f"--port={args.port}",
            f"--user={args.admin_user}",
            args.database,
        ]
        with source_path.open("rb") as stdin_handle:
            proc = subprocess.run(cmd, env=env, stdin=stdin_handle, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False)
        if proc.returncode != 0:
            stderr = proc.stderr.decode("utf-8", errors="replace")
            raise RuntimeError(stderr.strip() or f"mysql terminou com exit code {proc.returncode}")

        if not args.skip_validation:
            conn = connect_mysql(pymysql, args.host, args.port, args.admin_user, args.admin_password, args.database)
            try:
                with conn.cursor() as cur:
                    issues = validate_schema(cur)
            finally:
                conn.close()
            if issues:
                print("mysql-restore-validate-failed")
                for issue in issues:
                    print(f"- {issue}")
                return 1
    except Exception as exc:
        print(f"mysql-restore-failed: {exc}")
        return 1
    finally:
        if temp_sql_path is not None:
            temp_sql_path.unlink(missing_ok=True)

    print("mysql-restore-ok")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
