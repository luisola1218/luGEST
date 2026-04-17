from __future__ import annotations

import argparse
import gzip
import json
import os
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
ROOT = BASE_DIR.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

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


def timestamp_token() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Cria um backup SQL da base MySQL do LuGEST.")
    parser.add_argument("--host", default=str(os.environ.get("LUGEST_DB_HOST", "127.0.0.1") or "127.0.0.1"))
    parser.add_argument("--port", type=int, default=env_int("LUGEST_DB_PORT", 3306))
    parser.add_argument("--user", default=str(os.environ.get("LUGEST_DB_USER", "") or ""))
    parser.add_argument("--password", default=str(os.environ.get("LUGEST_DB_PASS", "") or ""))
    parser.add_argument("--database", default=str(os.environ.get("LUGEST_DB_NAME", "lugest") or "lugest"))
    parser.add_argument("--label", default="", help="Etiqueta opcional para identificar o backup.")
    parser.add_argument("--output-root", default=str(ROOT / "backups" / "mysql"))
    parser.add_argument("--mysqldump-path", default="")
    parser.add_argument("--plain-sql", action="store_true", help="Guarda o dump em .sql sem compressao gzip.")
    parser.add_argument("--dry-run", action="store_true", help="Mostra o plano sem criar ficheiros.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    dump_bin = resolve_mysql_binary(args.mysqldump_path, "MYSQLDUMP_PATH", ["mysqldump.exe", "mysqldump"])
    token = timestamp_token()
    label = str(args.label or "").strip()
    suffix = f"_{label}" if label else ""
    backup_dir = Path(args.output_root).expanduser() / f"{args.database}_{token}{suffix}"
    sql_name = f"{args.database}_{token}{suffix}.sql"
    sql_path = backup_dir / sql_name
    archive_path = Path(str(sql_path) + ".gz")
    metadata_path = backup_dir / "metadata.json"

    print("== LuGEST MySQL Backup ==")
    print(f"Servidor....: {args.host}:{args.port}")
    print(f"Base........: {args.database}")
    print(f"Utilizador..: {args.user or '<vazio>'}")
    print(f"Destino.....: {backup_dir}")
    print(f"Formato.....: {'sql' if args.plain_sql else 'sql.gz'}")
    if args.dry_run:
        print("DryRun......: ativo")

    if not dump_bin:
        message = "mysqldump nao foi encontrado. Define --mysqldump-path ou MYSQLDUMP_PATH."
        if args.dry_run:
            print(f"[DRY-RUN] {message}")
            return 0
        print(message)
        return 2

    if not args.user or not args.password:
        print("Define --user e --password ou configura LUGEST_DB_USER/LUGEST_DB_PASS.")
        return 2

    cmd = [
        dump_bin,
        f"--host={args.host}",
        f"--port={args.port}",
        f"--user={args.user}",
        "--single-transaction",
        "--quick",
        "--skip-lock-tables",
        "--default-character-set=utf8mb4",
        "--routines",
        "--events",
        args.database,
    ]
    print("Comando.....: mysqldump <credenciais ocultas> --single-transaction --quick ...")

    if args.dry_run:
        return 0

    backup_dir.mkdir(parents=True, exist_ok=True)
    env = build_mysql_env(args.password)

    try:
        with sql_path.open("wb") as handle:
            proc = subprocess.run(cmd, env=env, stdout=handle, stderr=subprocess.PIPE, check=False)
        if proc.returncode != 0:
            stderr = proc.stderr.decode("utf-8", errors="replace")
            raise RuntimeError(stderr.strip() or f"mysqldump terminou com exit code {proc.returncode}")

        final_path = sql_path
        if not args.plain_sql:
            with sql_path.open("rb") as source, gzip.open(archive_path, "wb", compresslevel=6) as target:
                shutil.copyfileobj(source, target)
            sql_path.unlink(missing_ok=True)
            final_path = archive_path

        metadata = {
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "host": args.host,
            "port": args.port,
            "database": args.database,
            "user": args.user,
            "label": label,
            "dump_tool": dump_bin,
            "format": "sql" if args.plain_sql else "sql.gz",
            "artifact": final_path.name,
        }
        metadata_path.write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as exc:
        print(f"mysql-backup-failed: {exc}")
        return 1

    print("mysql-backup-ok")
    print(f"Arquivo.....: {final_path}")
    print(f"Metadata....: {metadata_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
