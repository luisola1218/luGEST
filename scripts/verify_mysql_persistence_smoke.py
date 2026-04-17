from __future__ import annotations

import json
import sys
import uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import main


def _client_exists(cur, code: str) -> dict | None:
    cur.execute(
        """
        SELECT codigo, nome
        FROM clientes
        WHERE codigo = %s
        LIMIT 1
        """,
        (code,),
    )
    row = cur.fetchone()
    return dict(row or {}) if isinstance(row, dict) else ({"codigo": row[0], "nome": row[1]} if row else None)


def main_entry() -> int:
    conn = main._mysql_connect()
    try:
        with conn.cursor() as cur:
            tables = main._mysql_existing_tables(cur, force=True)
            if "clientes" not in tables:
                raise RuntimeError("Tabela `clientes` nao encontrada via information_schema.")

            columns = main._mysql_table_columns(cur, "clientes", force=True)
            required_columns = {"codigo", "nome"}
            if not required_columns.issubset(columns):
                raise RuntimeError(f"Colunas em falta em `clientes`: {sorted(required_columns - set(columns))}")

            indexes = main._mysql_table_indexes(cur, "clientes", force=True)
            if not indexes:
                raise RuntimeError("Nao foi possivel ler indices de `clientes` via information_schema.")
    finally:
        conn.close()

    test_code = "ZZMETA" + uuid.uuid4().hex[:6].upper()
    test_name = f"Cliente Metadata {test_code}"
    cleanup_done = False

    try:
        data = main.load_data()
        data["clientes"] = [row for row in list(data.get("clientes", []) or []) if str(row.get("codigo", "") or "").strip() != test_code]
        data.setdefault("clientes", []).append(
            {
                "codigo": test_code,
                "nome": test_name,
                "empresa": test_name,
                "nif": "",
                "morada": "",
                "contacto": "",
                "email": "",
            }
        )
        main.save_data(data, force=True)

        conn = main._mysql_connect()
        try:
            with conn.cursor() as cur:
                row = _client_exists(cur, test_code)
        finally:
            conn.close()
        if not row:
            raise RuntimeError("Cliente de teste nao ficou persistido diretamente na tabela `clientes`.")

        cleanup = main.load_data()
        cleanup["clientes"] = [
            row
            for row in list(cleanup.get("clientes", []) or [])
            if str(row.get("codigo", "") or "").strip() != test_code
        ]
        main.save_data(cleanup, force=True)
        cleanup_done = True

        conn = main._mysql_connect()
        try:
            with conn.cursor() as cur:
                row_after = _client_exists(cur, test_code)
        finally:
            conn.close()
        if row_after:
            raise RuntimeError("Cliente de teste nao foi removido da tabela `clientes` durante a limpeza.")

    finally:
        if not cleanup_done:
            try:
                cleanup = main.load_data()
                cleanup["clientes"] = [
                    row
                    for row in list(cleanup.get("clientes", []) or [])
                    if str(row.get("codigo", "") or "").strip() != test_code
                ]
                main.save_data(cleanup, force=True)
            except Exception:
                pass

    print(
        json.dumps(
            {
                "metadata_tables_ok": True,
                "metadata_columns_ok": True,
                "metadata_indexes_ok": True,
                "client_created_and_removed": True,
                "test_code": test_code,
            },
            ensure_ascii=False,
        )
    )
    print("mysql-persistence-smoke-ok")
    return 0


if __name__ == "__main__":
    raise SystemExit(main_entry())
