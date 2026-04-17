import argparse
import os
import sys
from pathlib import Path


def split_sql_statements(sql_text: str):
    stmts = []
    buf = []
    in_single = False
    in_double = False
    in_backtick = False
    i = 0
    while i < len(sql_text):
        ch = sql_text[i]
        nxt = sql_text[i + 1] if i + 1 < len(sql_text) else ""

        # line comments
        if not in_single and not in_double and not in_backtick and ch == "-" and nxt == "-":
            while i < len(sql_text) and sql_text[i] != "\n":
                i += 1
            continue
        if not in_single and not in_double and not in_backtick and ch == "#":
            while i < len(sql_text) and sql_text[i] != "\n":
                i += 1
            continue

        if ch == "'" and not in_double and not in_backtick:
            in_single = not in_single
            buf.append(ch)
            i += 1
            continue
        if ch == '"' and not in_single and not in_backtick:
            in_double = not in_double
            buf.append(ch)
            i += 1
            continue
        if ch == "`" and not in_single and not in_double:
            in_backtick = not in_backtick
            buf.append(ch)
            i += 1
            continue

        if ch == ";" and not in_single and not in_double and not in_backtick:
            stmt = "".join(buf).strip()
            if stmt:
                stmts.append(stmt)
            buf = []
            i += 1
            continue

        buf.append(ch)
        i += 1

    tail = "".join(buf).strip()
    if tail:
        stmts.append(tail)
    return stmts


def main():
    parser = argparse.ArgumentParser(description="Apply LU-GEST MySQL update script.")
    parser.add_argument("--host", default=str(os.environ.get("LUGEST_DB_HOST", "127.0.0.1") or "127.0.0.1"))
    parser.add_argument("--port", type=int, default=int(os.environ.get("LUGEST_DB_PORT", "3306") or "3306"))
    parser.add_argument("--user", default=str(os.environ.get("LUGEST_DB_USER", "root") or "root"))
    parser.add_argument("--password", default=str(os.environ.get("LUGEST_DB_PASS", "") or ""))
    parser.add_argument("--database", default=str(os.environ.get("LUGEST_DB_NAME", "lugest") or "lugest"))
    parser.add_argument(
        "--file",
        default="mysql/lugest_update_schema_mysql56_heidi_safe.sql",
        help="SQL file path",
    )
    args = parser.parse_args()

    try:
        import pymysql
    except Exception:
        print("PyMySQL não está instalado. Instale com: python -m pip install PyMySQL")
        return 2

    sql_path = Path(args.file)
    if not sql_path.exists():
        print(f"Ficheiro SQL não encontrado: {sql_path}")
        return 2

    sql_text = sql_path.read_text(encoding="utf-8")
    statements = split_sql_statements(sql_text)
    if not statements:
        print("Sem statements SQL para executar.")
        return 1

    conn = pymysql.connect(
        host=args.host,
        port=args.port,
        user=args.user,
        password=args.password,
        database=args.database,
        charset="utf8mb4",
        autocommit=False,
        connect_timeout=8,
        read_timeout=30,
        write_timeout=30,
    )

    ok = 0
    failed = 0
    try:
        with conn.cursor() as cur:
            for idx, stmt in enumerate(statements, start=1):
                try:
                    cur.execute(stmt)
                    ok += 1
                except Exception as ex:
                    failed += 1
                    print(f"[ERRO {idx}] {ex}")
                    print(f"SQL: {stmt[:220]}{'...' if len(stmt) > 220 else ''}")
            conn.commit()
    finally:
        conn.close()

    print(f"Concluído. OK={ok}, ERROS={failed}")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
