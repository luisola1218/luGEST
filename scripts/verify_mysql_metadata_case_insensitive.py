from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import main


class FakeCursor:
    def __init__(self, rows):
        self._rows = list(rows)

    def execute(self, _query, _params=None):
        return None

    def fetchall(self):
        return list(self._rows)


def main_entry() -> int:
    main._mysql_schema_cache_reset()

    tables = main._mysql_existing_tables(FakeCursor([{"TABLE_NAME": "clientes"}, {"table_name": "users"}]), force=True)
    if tables != {"clientes", "users"}:
        raise RuntimeError(f"Tabelas lidas incorretamente: {tables}")

    main._mysql_schema_cache_reset()
    columns = main._mysql_table_columns(
        FakeCursor([{"COLUMN_NAME": "Codigo"}, {"column_name": "Nome"}, {"CoLuMn_NaMe": "Email"}]),
        "clientes",
        force=True,
    )
    if columns != {"codigo", "nome", "email"}:
        raise RuntimeError(f"Colunas lidas incorretamente: {columns}")

    main._mysql_schema_cache_reset()
    indexes = main._mysql_table_indexes(
        FakeCursor([{"INDEX_NAME": "idx_clientes_nome"}, {"index_name": "PRIMARY"}, {"InDeX_NaMe": "idx_clientes_email"}]),
        "clientes",
        force=True,
    )
    if indexes != {"idx_clientes_nome", "primary", "idx_clientes_email"}:
        raise RuntimeError(f"Indices lidos incorretamente: {indexes}")

    print("mysql-metadata-case-insensitive-ok")
    return 0


if __name__ == "__main__":
    raise SystemExit(main_entry())
