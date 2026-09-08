"""Explicitly authorized integration tests on the configured database, rolled back.

Run one flow per process. Only transactional InnoDB tables are accepted. DDL,
SQL transaction control and secondary connections are blocked. AUTO_INCREMENT
allocations may leave gaps even after rollback; table-row hashes exclude these.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import runpy
import sys
import tempfile
import time
import traceback

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
FLOWS = ("verify_purchase_flow", "verify_conjuntos_montagem_flow",
         "verify_fabrication_order_flow", "verify_planning_flow", "verify_billing_flow",
         "verify_quote_nesting_flow", "verify_inventory_flow")


def _inventory_flow():
    """Temporary product exercised only inside this runner's transaction."""
    from copy import deepcopy
    from uuid import uuid4
    from lugest_qt.services.legacy_backend import LegacyBackend
    backend = LegacyBackend()
    code = 'TST-ST-' + uuid4().hex[:10]
    product = backend.product_save({'codigo': code, 'descricao': 'Parafuso teste transacional',
                                    'qty': 10, 'p_compra': 3.5, 'pvp1': 5, 'unid': 'UN'})
    assert product['qty'] == 10
    before = deepcopy(backend.ensure_data())
    try:
        backend.product_consume(code, 2, issue_mode='operator')
    except ValueError:
        pass
    else:
        raise AssertionError('Missing operator accepted')
    assert backend.ensure_data() == before
    result = backend.product_consume(code, 2, issue_mode='operator', target_operator='TEST')
    assert result['qty'] == 8
    backend.reload(force=True)
    assert backend.product_detail(code)['qty'] == 8
    summary = backend.product_issue_summary(operator_name='TEST', codigo=code)
    assert summary['linhas'] == 1 and summary['qtd_total'] == 2
    backend.product_save({**backend.product_detail(code), 'qty': 6})
    backend.reload(force=True)
    assert backend.product_detail(code)['qty'] == 6
    assert any(row['tipo'] == 'AJUSTE_STOCK' for row in backend.product_movements(code))
    assert backend.product_remove_many([code, code]) == 1
    backend.reload(force=True)
    assert not any(row['codigo'] == code for row in backend.product_rows())
    print('inventory-db-ok create=yes edit=yes delete=yes invalid-no-mutation=yes issue=yes reload=yes movements=yes', flush=True)


def _quote_nesting_flow():
    """Only called after the transaction guard is installed by this runner."""
    from uuid import uuid4
    from lugest_qt.services.legacy_backend import LegacyBackend
    backend = LegacyBackend()
    token = uuid4().hex[:12]
    client = backend.client_save({'codigo': 'CT-' + token, 'nome': 'Transactional test'})
    number = 'ORC-TEST-' + token
    backend.orc_save({'numero': number, 'cliente': client, 'linhas': [{
        'tipo_item': backend.desktop_main.ORC_LINE_TYPE_SERVICE,
        'descricao': 'Transactional test', 'qtd': 1, 'preco_unit': 1,
        'operacao': 'Montagem',
    }]})
    key = 'rollback-test-' + token
    saved = backend.orc_save_nesting_study(number, {
        'group_key': key, 'group_label': 'Transactional integration test',
        'quote_bridge': {'quantity': 2},
    })
    assert saved['quote_number'] == number
    assert backend._mysql_orc_nesting_studies(number)[key]['quote_bridge']['quantity'] == 2
    created = saved['created_at']
    saved['quote_bridge']['quantity'] = 3
    updated = backend.orc_save_nesting_study(number, saved)
    assert updated['created_at'] == created
    assert backend.orc_nesting_studies(number)[key]['quote_bridge']['quantity'] == 3
    backend._mysql_delete_orc_nesting_studies(number, key)
    assert key not in backend._mysql_orc_nesting_studies(number)
    print('quote-nesting-db-ok create=yes update=yes sql-mirror=yes delete=yes', flush=True)


def table_hashes(connection, tables):
    result = {}
    with connection.cursor() as cursor:
        for table in tables:
            cursor.execute(f"SELECT * FROM `{table}`")
            rows = sorted(json.dumps(row, sort_keys=True, default=str, ensure_ascii=False) for row in cursor.fetchall())
            result[table] = {"count": len(rows), "sha256": hashlib.sha256("\n".join(rows).encode()).hexdigest()}
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("flow", choices=FLOWS)
    args = parser.parse_args()
    import main as runtime
    import pymysql

    report = {"flow": args.flow, "status": "failed", "rollback": False}
    connect = pymysql.connect
    connection = connect(host=runtime.MYSQL_HOST, port=runtime.MYSQL_PORT,
                         user=runtime.MYSQL_USER, password=runtime.MYSQL_PASSWORD,
                         database=runtime.MYSQL_DB_NAME, charset="utf8mb4", autocommit=False,
                         cursorclass=pymysql.cursors.DictCursor, connect_timeout=5,
                         read_timeout=30, write_timeout=30)
    before = {}
    tables = []
    started = time.monotonic()
    try:
        with connection.cursor() as cursor:
            cursor.execute("SELECT TABLE_NAME, ENGINE FROM information_schema.TABLES WHERE TABLE_SCHEMA=%s AND TABLE_TYPE='BASE TABLE'", (runtime.MYSQL_DB_NAME,))
            metadata = cursor.fetchall()
            if not metadata or any(row["ENGINE"] != "InnoDB" for row in metadata):
                raise RuntimeError("Rollback tests require all tables to use InnoDB")
            tables = sorted(row["TABLE_NAME"] for row in metadata)
            if any(not re.fullmatch(r"[A-Za-z0-9_]+", name) for name in tables):
                raise RuntimeError("Unexpected table identifier")
            cursor.execute("SET SESSION innodb_lock_wait_timeout=5")
        before = table_hashes(connection, tables)
        report["tables"] = len(tables)
        allowed_tables = {name.casefold() for name in tables}

        class GuardCursor:
            def __init__(self, cursor): self.cursor = cursor
            def __enter__(self): return self
            def __exit__(self, *args): self.cursor.close()
            def __getattr__(self, name): return getattr(self.cursor, name)
            def execute(self, sql, parameters=None):
                text = str(sql).strip()
                normalized = re.sub(r"\s+", " ", text).upper()
                existing = re.match(r"CREATE TABLE IF NOT EXISTS `?([A-Za-z0-9_]+)`?\s*\(", text, re.I)
                if existing and existing.group(1).casefold() in allowed_tables:
                    self.cursor.rowcount = 0
                    return 0
                verb = normalized.split(" ", 1)[0]
                if verb not in {"SELECT", "SHOW", "DESCRIBE", "INSERT", "UPDATE", "DELETE", "REPLACE", "SET"}:
                    raise RuntimeError("SQL blocked in rollback test: " + verb)
                if verb == "SET" and re.search(r"AUTOCOMMIT|TRANSACTION|GLOBAL", normalized):
                    raise RuntimeError("SQL transaction/session change blocked")
                return self.cursor.execute(sql, parameters)
            def executemany(self, sql, rows):
                total = 0
                for row in rows:
                    total += self.execute(sql, row)
                return total

        class GuardConnection:
            def cursor(self, cursorclass=None):
                return GuardCursor(connection.cursor(cursorclass))
            def commit(self): pass
            def rollback(self): pass
            def close(self): pass
            def ping(self, reconnect=False):
                return connection.ping(reconnect=False)
            def autocommit(self, value):
                if value:
                    raise RuntimeError("Autocommit blocked")
            def __enter__(self): return self
            def __exit__(self, *args): pass

        proxy = GuardConnection()

        def guarded_connect(*positional, **kwargs):
            if positional or kwargs.get("database", kwargs.get("db")) != runtime.MYSQL_DB_NAME:
                raise RuntimeError("Different database blocked")
            if kwargs.get("host") != runtime.MYSQL_HOST or kwargs.get("autocommit", False):
                raise RuntimeError("Different server/autocommit blocked")
            return proxy

        pymysql.connect = guarded_connect
        runtime._MYSQL_SCHEMA_SYNCED = True
        runtime._ASYNC_SAVE_ENABLED = False
        with tempfile.TemporaryDirectory(prefix="lugest-rollback-") as temporary:
            runtime.BASE_DIR = temporary
            for key, folder in (("LUGEST_USER_DATA_DIR", "user"), ("LUGEST_MACHINE_DATA_DIR", "machine"),
                                ("LUGEST_SHARED_STORAGE_ROOT", "shared")):
                os.environ[key] = str(Path(temporary) / folder)
            if args.flow == 'verify_quote_nesting_flow':
                code = _quote_nesting_flow()
            elif args.flow == 'verify_inventory_flow':
                code = _inventory_flow()
            else:
                entry = runpy.run_path(str(ROOT / "scripts" / (args.flow + ".py")))
                code = entry["main"]()
            if code not in (0, None):
                raise RuntimeError(f"Flow returned {code}")
        report["status"] = "passed"
    except Exception as exc:
        report.update(error=str(exc), traceback=traceback.format_exc())
        print("Flow failed:", str(exc), flush=True)
    finally:
        pymysql.connect = connect
        try:
            connection.rollback()
            report["rollback"] = True
            if before:
                after = table_hashes(connection, tables)
                report["rows_unchanged"] = before == after
                report["changed_tables"] = [name for name in tables if before[name] != after[name]]
        finally:
            connection.rollback()
            connection.close()
        report["seconds"] = round(time.monotonic() - started, 3)
        output = ROOT / "reports/integration" / (args.flow + "_rollback.json")
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
        print(json.dumps({key: value for key, value in report.items() if key != "traceback"}, ensure_ascii=False), flush=True)
    return 0 if report["status"] == "passed" and report.get("rows_unchanged") else 1


if __name__ == "__main__":
    raise SystemExit(main())
