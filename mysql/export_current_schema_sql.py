from __future__ import annotations

import argparse
import re
import sys
from collections import deque
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import main


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_OUTPUT = BASE_DIR / "lugest.sql"
FULL_INSTALL_OUTPUT = BASE_DIR / "lugest_instalacao_unica.sql"

STARTER_USERS = [
    ("admin", "Trocar#Admin2026", "Admin"),
    ("operador", "Trocar#Operador2026", "Operador"),
    ("orcamentista", "Trocar#Orc2026", "Orcamentista"),
    ("planeamento", "Trocar#Planeamento2026", "Planeamento"),
]

STARTER_OPERADORES = ["Operador 1"]
STARTER_ORCAMENTISTAS = ["Orcamentista 1"]


def _row_lookup(row, *keys):
    if isinstance(row, dict):
        normalized = {}
        for key, value in row.items():
            key_txt = str(key or "").strip()
            if not key_txt:
                continue
            normalized.setdefault(key_txt, value)
            normalized.setdefault(_normalize_key(key_txt), value)
        for key in keys:
            key_txt = str(key or "").strip()
            if not key_txt:
                continue
            if key_txt in normalized:
                return normalized[key_txt]
            alt_key = _normalize_key(key_txt)
            if alt_key in normalized:
                return normalized[alt_key]
        return None
    if row:
        try:
            return row[0]
        except Exception:
            return None
    return None


def _normalize_key(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().lower())


def _database_options(cur) -> tuple[str, str]:
    cur.execute(
        """
        SELECT DEFAULT_CHARACTER_SET_NAME, DEFAULT_COLLATION_NAME
        FROM information_schema.SCHEMATA
        WHERE SCHEMA_NAME = DATABASE()
        """
    )
    row = cur.fetchone()
    charset = str(_row_lookup(row, "DEFAULT_CHARACTER_SET_NAME", "default_character_set_name") or "utf8mb4").strip() or "utf8mb4"
    collation = str(_row_lookup(row, "DEFAULT_COLLATION_NAME", "default_collation_name") or "utf8mb4_unicode_ci").strip() or "utf8mb4_unicode_ci"
    return charset, collation


def _table_names(cur) -> list[str]:
    cur.execute(
        """
        SELECT table_name
        FROM information_schema.tables
        WHERE table_schema = DATABASE() AND table_type = 'BASE TABLE'
        ORDER BY table_name
        """
    )
    rows = cur.fetchall() or []
    names = []
    for row in rows:
        table_name = str(_row_lookup(row, "table_name", "TABLE_NAME") or "").strip()
        if table_name:
            names.append(table_name)
    return names


def _dependencies(cur, tables: list[str]) -> dict[str, set[str]]:
    known = set(tables)
    deps = {table: set() for table in tables}
    cur.execute(
        """
        SELECT table_name, referenced_table_name
        FROM information_schema.key_column_usage
        WHERE table_schema = DATABASE()
          AND referenced_table_name IS NOT NULL
        """
    )
    for row in cur.fetchall() or []:
        table_name = str(_row_lookup(row, "table_name", "TABLE_NAME") or "").strip()
        ref_table = str(_row_lookup(row, "referenced_table_name", "REFERENCED_TABLE_NAME") or "").strip()
        if table_name in known and ref_table in known and table_name != ref_table:
            deps.setdefault(table_name, set()).add(ref_table)
    return deps


def _topological_order(tables: list[str], deps: dict[str, set[str]]) -> list[str]:
    in_degree = {table: len(deps.get(table, set())) for table in tables}
    reverse: dict[str, set[str]] = {table: set() for table in tables}
    for table, refs in deps.items():
        for ref in refs:
            if ref in reverse:
                reverse[ref].add(table)
    queue = deque(sorted(table for table, degree in in_degree.items() if degree == 0))
    ordered: list[str] = []
    while queue:
        table = queue.popleft()
        ordered.append(table)
        for follower in sorted(reverse.get(table, set())):
            in_degree[follower] -= 1
            if in_degree[follower] == 0:
                queue.append(follower)
    if len(ordered) != len(tables):
        remaining = [table for table in tables if table not in ordered]
        ordered.extend(sorted(remaining))
    return ordered


def _show_create_table(cur, table_name: str) -> str:
    cur.execute(f"SHOW CREATE TABLE `{table_name}`")
    row = cur.fetchone()
    create_sql = _row_lookup(row, "Create Table", "create_table")
    if not create_sql:
        raise RuntimeError(f"Nao foi possivel obter o CREATE TABLE de `{table_name}`.")
    statement = str(create_sql).strip().rstrip(";")
    statement = re.sub(r"^CREATE TABLE\s+`", "CREATE TABLE IF NOT EXISTS `", statement, count=1, flags=re.IGNORECASE)
    statement = re.sub(r"\sAUTO_INCREMENT=\d+\b", "", statement)
    return statement + ";"


def export_schema(output_path: Path) -> Path:
    conn = main._mysql_connect()
    try:
        with conn.cursor() as cur:
            charset, collation = _database_options(cur)
            tables = _table_names(cur)
            deps = _dependencies(cur, tables)
            ordered_tables = _topological_order(tables, deps)
            statements = [_show_create_table(cur, table_name) for table_name in ordered_tables]
    finally:
        conn.close()

    lines = [
        "-- =====================================================",
        "-- LuGEST CURRENT FULL SCHEMA",
        "-- Gerado a partir da base MySQL atual",
        "-- =====================================================",
        "",
        "-- Opcional: para reiniciar tudo, descomenta a linha seguinte.",
        "-- DROP DATABASE IF EXISTS `lugest`;",
        f"CREATE DATABASE IF NOT EXISTS `lugest` CHARACTER SET {charset} COLLATE {collation};",
        "USE `lugest`;",
        "SET NAMES utf8mb4;",
        "SET FOREIGN_KEY_CHECKS=0;",
        "",
        "-- =====================================================",
        "-- TABELAS",
        "-- =====================================================",
        "",
    ]
    for statement in statements:
        lines.append(statement)
        lines.append("")
    lines.extend(
        [
            "SET FOREIGN_KEY_CHECKS=1;",
            "",
            "-- Fim do schema atual consolidado.",
            "",
        ]
    )

    output_path.write_text("\n".join(lines), encoding="utf-8")
    return output_path


def _sql_quote(value: str) -> str:
    return "'" + str(value or "").replace("\\", "\\\\").replace("'", "''") + "'"


def _starter_seed_lines() -> list[str]:
    lines = [
        "-- =====================================================",
        "-- UTILIZADORES INICIAIS",
        "-- =====================================================",
        "-- Primeiro login temporario:",
    ]
    for username, password, role in STARTER_USERS:
        lines.append(f"--   {username} / {password} ({role})")
    lines.extend(
        [
            "-- Troca estas passwords logo apos a instalacao.",
            "",
        ]
    )

    user_values = []
    for username, password, role in STARTER_USERS:
        hashed = main.normalize_password_for_storage(username, password, require_strong=False)
        user_values.append(f"({_sql_quote(username)}, {_sql_quote(hashed)}, {_sql_quote(role)})")
    lines.append("INSERT IGNORE INTO `users` (`username`, `password`, `role`) VALUES")
    lines.append("  " + ",\n  ".join(user_values) + ";")
    lines.append("")

    operador_values = ", ".join(f"({_sql_quote(name)})" for name in STARTER_OPERADORES)
    if operador_values:
        lines.append(f"INSERT IGNORE INTO `operadores` (`nome`) VALUES {operador_values};")
        lines.append("")

    orc_values = ", ".join(f"({_sql_quote(name)})" for name in STARTER_ORCAMENTISTAS)
    if orc_values:
        lines.append(f"INSERT IGNORE INTO `orcamentistas` (`nome`) VALUES {orc_values};")
        lines.append("")

    return lines


def export_full_install(output_path: Path) -> Path:
    export_schema(output_path)
    base_text = output_path.read_text(encoding="utf-8")
    seed_lines = _starter_seed_lines()
    final_text = base_text.rstrip() + "\n\n" + "\n".join(seed_lines) + "\n"
    output_path.write_text(final_text, encoding="utf-8")
    return output_path


def main_entry(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Exporta o schema atual do LuGEST para um SQL unico.")
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT), help="Caminho do ficheiro SQL de saida.")
    parser.add_argument(
        "--with-starter-users",
        action="store_true",
        help="Inclui utilizadores iniciais minimos e listas base para arranque imediato.",
    )
    args = parser.parse_args(argv)
    output_default = FULL_INSTALL_OUTPUT if args.with_starter_users and str(args.output) == str(DEFAULT_OUTPUT) else Path(args.output)
    output_path = output_default.resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if args.with_starter_users:
        final_path = export_full_install(output_path)
    else:
        final_path = export_schema(output_path)
    print(final_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main_entry())
