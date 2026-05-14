from __future__ import annotations

import argparse
import os
import re
import sys
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
BASE_SCHEMA_PATH = BASE_DIR / "lugest.sql"
PATCH_FILES = sorted([*BASE_DIR.glob("patch_*.sql"), *(BASE_DIR / "Migracoes").glob("patch_*.sql")])

REQUIRED = {
    "app_config": {"ckey", "cvalue"},
    "clientes": {"codigo", "nome"},
    "conjuntos_modelo": {"codigo", "descricao", "ativo", "updated_at"},
    "conjuntos_modelo_itens": {"conjunto_codigo", "linha_ordem", "tipo_item", "qtd", "preco_unit"},
    "fornecedores": {"id", "nome"},
    "materiais": {"id", "material", "espessura", "quantidade", "reservado", "lote_interno", "lote_fornecedor"},
    "encomendas": {"numero", "cliente_codigo", "estado", "data_entrega"},
    "encomenda_espessuras": {"encomenda_numero", "material", "espessura", "tempo_min", "estado"},
    "encomenda_montagem_itens": {"encomenda_numero", "linha_ordem", "tipo_item", "qtd_planeada", "qtd_consumida", "estado"},
    "pecas": {"id", "encomenda_numero", "ref_interna", "ref_externa", "of_codigo", "opp_codigo", "estado"},
    "orcamentos": {"numero", "cliente_codigo", "estado"},
    "orcamento_linhas": {
        "orcamento_numero",
        "ref_interna",
        "ref_externa",
        "material",
        "espessura",
        "tempo_peca_min",
        "tipo_item",
        "produto_codigo",
        "produto_unid",
        "conjunto_codigo",
        "conjunto_nome",
        "grupo_uuid",
        "qtd_base",
    },
    "orc_referencias_historico": {"ref_externa", "ref_interna", "material", "espessura"},
    "notas_encomenda": {"numero", "fornecedor_id", "estado", "total"},
    "notas_encomenda_linhas": {"ne_numero", "ref_material", "origem", "qtd", "preco", "total"},
    "notas_encomenda_entregas": {"ne_numero", "data_entrega", "guia", "fatura"},
    "notas_encomenda_linha_entregas": {"ne_numero", "linha_ordem", "qtd"},
    "notas_encomenda_documentos": {"ne_numero", "tipo", "titulo", "caminho", "guia", "fatura", "data_entrega", "data_documento"},
    "expedicoes": {"numero", "encomenda_numero", "estado"},
    "expedicao_linhas": {"expedicao_numero", "ref_interna", "qtd"},
    "faturacao_registos": {"numero", "orcamento_numero", "encomenda_numero", "cliente_codigo", "data_venda", "data_vencimento"},
    "faturacao_faturas": {
        "registo_numero",
        "documento_id",
        "numero_fatura",
        "data_emissao",
        "valor_total",
        "legal_invoice_no",
        "system_entry_date",
        "source_id",
        "source_billing",
        "hash",
        "hash_control",
        "communication_status",
    },
    "faturacao_pagamentos": {"registo_numero", "pagamento_id", "fatura_documento_id", "data_pagamento", "valor"},
    "plano": {"bloco_id", "encomenda_numero", "data_planeada", "inicio", "duracao_min"},
    "produtos": {"codigo", "descricao", "qty"},
    "produtos_mov": {"codigo", "tipo", "qtd", "data"},
    "stock_log": {"acao", "data"},
    "users": {"username", "password", "role"},
}

IDEMPOTENT_ERROR_CODES = {
    1007,  # database exists
    1050,  # table exists
    1060,  # duplicate column
    1061,  # duplicate key
    1091,  # can't drop missing key/column
    1826,  # duplicate foreign key constraint name
}

IDEMPOTENT_ERROR_FRAGMENTS = (
    "already exists",
    "duplicate column name",
    "duplicate key name",
    "duplicate entry",
    "can't drop",
    "duplicate foreign key constraint name",
)


def split_sql_statements(sql_text: str) -> list[str]:
    statements: list[str] = []
    buf: list[str] = []
    in_single = False
    in_double = False
    in_backtick = False
    in_block_comment = False
    i = 0
    while i < len(sql_text):
        ch = sql_text[i]
        nxt = sql_text[i + 1] if i + 1 < len(sql_text) else ""

        if in_block_comment:
            if ch == "*" and nxt == "/":
                in_block_comment = False
                i += 2
                continue
            i += 1
            continue

        if not in_single and not in_double and not in_backtick and ch == "/" and nxt == "*":
            in_block_comment = True
            i += 2
            continue
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
                statements.append(stmt)
            buf = []
            i += 1
            continue

        buf.append(ch)
        i += 1

    tail = "".join(buf).strip()
    if tail:
        statements.append(tail)
    return statements


def sanitize_base_schema(sql_text: str) -> list[str]:
    filtered: list[str] = []
    for stmt in split_sql_statements(sql_text):
        normalized = re.sub(r"\s+", " ", stmt.strip()).upper()
        if normalized.startswith("DROP DATABASE "):
            continue
        if normalized.startswith("CREATE DATABASE "):
            continue
        if normalized.startswith("USE "):
            continue
        filtered.append(stmt)
    return filtered


def _error_code(exc: Exception) -> int | None:
    if not getattr(exc, "args", None):
        return None
    try:
        return int(exc.args[0])
    except Exception:
        return None


def is_idempotent_error(exc: Exception) -> bool:
    code = _error_code(exc)
    if code in IDEMPOTENT_ERROR_CODES:
        return True
    message = str(exc).lower()
    return any(fragment in message for fragment in IDEMPOTENT_ERROR_FRAGMENTS)


def connect_mysql(pymysql_module, host: str, port: int, user: str, password: str, database: str | None = None):
    kwargs = {
        "host": host,
        "port": port,
        "user": user,
        "password": password,
        "charset": "utf8mb4",
        "autocommit": False,
        "connect_timeout": 10,
        "read_timeout": 30,
        "write_timeout": 30,
    }
    if database:
        kwargs["database"] = database
    return pymysql_module.connect(**kwargs)


def _row_lookup(row, *keys):
    if isinstance(row, dict):
        lowered = {str(key or "").strip().lower(): value for key, value in row.items()}
        for key in keys:
            key_txt = str(key or "").strip()
            if not key_txt:
                continue
            if key_txt in row:
                return row[key_txt]
            found = lowered.get(key_txt.lower())
            if found is not None:
                return found
        return None
    if row:
        try:
            return row[0]
        except Exception:
            return None
    return None


def validate_schema(cur) -> list[str]:
    cur.execute("SELECT table_name FROM information_schema.tables WHERE table_schema = DATABASE()")
    existing_tables = {
        str(_row_lookup(row, "table_name", "TABLE_NAME") or "").strip()
        for row in (cur.fetchall() or [])
        if str(_row_lookup(row, "table_name", "TABLE_NAME") or "").strip()
    }
    issues: list[str] = []
    missing_tables = sorted(table for table in REQUIRED if table not in existing_tables)
    if missing_tables:
        issues.append("Tabelas em falta: " + ", ".join(missing_tables))
        return issues

    for table, expected_columns in REQUIRED.items():
        cur.execute(
            """
            SELECT column_name
            FROM information_schema.columns
            WHERE table_schema = DATABASE() AND table_name = %s
            """,
            (table,),
        )
        rows = cur.fetchall() or []
        existing_columns = {
            str(_row_lookup(row, "column_name", "COLUMN_NAME") or "").strip()
            for row in rows
            if str(_row_lookup(row, "column_name", "COLUMN_NAME") or "").strip()
        }
        missing = sorted(col for col in expected_columns if col not in existing_columns)
        if missing:
            issues.append(f"{table}: {', '.join(missing)}")
    return issues


def execute_sql_batch(cur, statements: list[str], *, tolerant: bool, dry_run: bool, label: str) -> tuple[int, int]:
    if dry_run:
        print(f"[DRY-RUN] {label}: {len(statements)} statement(s)")
        return len(statements), 0

    executed = 0
    skipped = 0
    for idx, stmt in enumerate(statements, start=1):
        try:
            cur.execute(stmt)
            executed += 1
        except Exception as exc:
            if tolerant and is_idempotent_error(exc):
                skipped += 1
                continue
            print(f"[ERRO] {label} statement {idx}: {exc}")
            raise
    return executed, skipped


def create_or_update_app_user(cur, *, database: str, app_user: str, app_password: str, app_host: str, dry_run: bool) -> None:
    create_stmt = f"CREATE USER IF NOT EXISTS `{app_user}`@`{app_host}` IDENTIFIED BY %s"
    alter_stmt = f"ALTER USER `{app_user}`@`{app_host}` IDENTIFIED BY %s"
    grant_stmt = f"GRANT ALL PRIVILEGES ON `{database}`.* TO `{app_user}`@`{app_host}`"
    flush_stmt = "FLUSH PRIVILEGES"

    if dry_run:
        print(f"[DRY-RUN] CREATE/ALTER USER `{app_user}`@`{app_host}` e GRANT ALL em `{database}`.*")
        return

    cur.execute(create_stmt, (app_password,))
    cur.execute(alter_stmt, (app_password,))
    cur.execute(grant_stmt)
    cur.execute(flush_stmt)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Provisiona ou valida a base de dados MySQL do LuGEST.")
    parser.add_argument("--host", default=str(os.environ.get("LUGEST_MYSQL_ADMIN_HOST", "127.0.0.1") or "127.0.0.1"))
    parser.add_argument("--port", type=int, default=int(os.environ.get("LUGEST_MYSQL_ADMIN_PORT", "3306") or "3306"))
    parser.add_argument("--admin-user", default=str(os.environ.get("LUGEST_MYSQL_ADMIN_USER", "root") or "root"))
    parser.add_argument("--admin-password", default=str(os.environ.get("LUGEST_MYSQL_ADMIN_PASSWORD", "") or ""))
    parser.add_argument("--database", default=str(os.environ.get("LUGEST_DB_NAME", "lugest") or "lugest"))
    parser.add_argument("--app-user", default=str(os.environ.get("LUGEST_DB_USER", "") or ""))
    parser.add_argument("--app-password", default=str(os.environ.get("LUGEST_DB_PASS", "") or ""))
    parser.add_argument("--app-host", default=str(os.environ.get("LUGEST_DB_APP_HOST", "localhost") or "localhost"))
    parser.add_argument("--skip-base-schema", action="store_true", help="Nao importa o lugest.sql.")
    parser.add_argument("--skip-patches", action="store_true", help="Nao aplica os patch_*.sql.")
    parser.add_argument("--skip-validation", action="store_true", help="Nao valida o schema final.")
    parser.add_argument("--validate-only", action="store_true", help="Apenas valida o schema da base existente.")
    parser.add_argument("--reset-database", action="store_true", help="Apaga e recria a base antes de importar o schema.")
    parser.add_argument("--dry-run", action="store_true", help="Mostra o plano sem aplicar alteracoes.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    if args.validate_only and (args.reset_database or not args.skip_base_schema or not args.skip_patches):
        args.reset_database = False
        args.skip_base_schema = True
        args.skip_patches = True

    if not args.validate_only and bool(args.app_user) != bool(args.app_password):
        print("Define ambos --app-user e --app-password para criar/atualizar o utilizador dedicado.")
        return 2

    try:
        import pymysql
    except Exception as exc:
        print(f"PyMySQL nao esta instalado neste Python: {exc}")
        return 2

    print("== LuGEST MySQL Installer ==")
    print(f"Servidor admin: {args.host}:{args.port}")
    print(f"Base alvo.....: {args.database}")
    if args.app_user:
        print(f"Utilizador app: {args.app_user}@{args.app_host}")
    else:
        print("Utilizador app: <nao configurado neste passo>")
    print(f"Modo..........: {'VALIDACAO' if args.validate_only else 'INSTALACAO'}")
    if args.dry_run:
        print("DryRun........: ativo")

    if args.validate_only:
        conn = connect_mysql(pymysql, args.host, args.port, args.admin_user, args.admin_password, args.database)
        try:
            with conn.cursor() as cur:
                issues = validate_schema(cur)
        finally:
            conn.close()
        if issues:
            print("mysql-validate-failed")
            for issue in issues:
                print(f"- {issue}")
            return 1
        print("mysql-validate-ok")
        return 0

    if not BASE_SCHEMA_PATH.exists() and not args.skip_base_schema:
        print(f"Schema base em falta: {BASE_SCHEMA_PATH}")
        return 2

    server_conn = None
    db_conn = None
    try:
        if not args.dry_run:
            server_conn = connect_mysql(pymysql, args.host, args.port, args.admin_user, args.admin_password, None)
            with server_conn.cursor() as cur:
                if args.reset_database:
                    cur.execute(f"DROP DATABASE IF EXISTS `{args.database}`")
                cur.execute(f"CREATE DATABASE IF NOT EXISTS `{args.database}`")
                if args.app_user:
                    create_or_update_app_user(
                        cur,
                        database=args.database,
                        app_user=args.app_user,
                        app_password=args.app_password,
                        app_host=args.app_host,
                        dry_run=False,
                    )
            server_conn.commit()
        else:
            print(f"[DRY-RUN] CREATE DATABASE IF NOT EXISTS `{args.database}`")
            if args.reset_database:
                print(f"[DRY-RUN] DROP DATABASE IF EXISTS `{args.database}`")
            if args.app_user:
                create_or_update_app_user(
                    None,
                    database=args.database,
                    app_user=args.app_user,
                    app_password=args.app_password,
                    app_host=args.app_host,
                    dry_run=True,
                )

        if not args.dry_run:
            db_conn = connect_mysql(pymysql, args.host, args.port, args.admin_user, args.admin_password, args.database)
            cur = db_conn.cursor()
        else:
            cur = None

        try:
            if not args.skip_base_schema:
                schema_text = BASE_SCHEMA_PATH.read_text(encoding="utf-8")
                schema_statements = split_sql_statements(schema_text) if args.reset_database else sanitize_base_schema(schema_text)
                executed, skipped = execute_sql_batch(
                    cur,
                    schema_statements,
                    tolerant=not args.reset_database,
                    dry_run=args.dry_run,
                    label="schema-base",
                )
                print(f"schema-base: executed={executed} skipped={skipped}")

            if not args.skip_patches:
                for patch_path in PATCH_FILES:
                    patch_statements = split_sql_statements(patch_path.read_text(encoding="utf-8"))
                    executed, skipped = execute_sql_batch(
                        cur,
                        patch_statements,
                        tolerant=True,
                        dry_run=args.dry_run,
                        label=patch_path.name,
                    )
                    print(f"{patch_path.name}: executed={executed} skipped={skipped}")

            if not args.dry_run and db_conn is not None:
                db_conn.commit()

            if not args.skip_validation:
                if args.dry_run:
                    print("[DRY-RUN] Validacao final do schema")
                else:
                    issues = validate_schema(cur)
                    if issues:
                        print("mysql-install-validate-failed")
                        for issue in issues:
                            print(f"- {issue}")
                        return 1
                    print("mysql-install-validate-ok")
        finally:
            if cur is not None:
                try:
                    cur.close()
                except Exception:
                    pass

    except Exception as exc:
        if db_conn is not None and not args.dry_run:
            try:
                db_conn.rollback()
            except Exception:
                pass
        if server_conn is not None and not args.dry_run:
            try:
                server_conn.rollback()
            except Exception:
                pass
        print(f"mysql-install-failed: {exc}")
        return 1
    finally:
        if db_conn is not None:
            try:
                db_conn.close()
            except Exception:
                pass
        if server_conn is not None:
            try:
                server_conn.close()
            except Exception:
                pass

    print("mysql-install-ok")
    print("Resumo env aplicacao:")
    print(f"LUGEST_DB_HOST={args.host}")
    print(f"LUGEST_DB_PORT={args.port}")
    if args.app_user:
        print(f"LUGEST_DB_USER={args.app_user}")
        print("LUGEST_DB_PASS=<password definido no passo de instalacao>")
    print(f"LUGEST_DB_NAME={args.database}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
