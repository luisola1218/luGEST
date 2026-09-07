"""Run integration flows in a newly created disposable database.

Uses configured server credentials but never reuses or clones the operational
database. Runtime writes and generated files are redirected to a temporary
directory. Run explicitly; this is not part of SafeOnly.
"""
from __future__ import annotations

import importlib.util
import json
import os
from pathlib import Path
import runpy
import sys
import tempfile
import time
import traceback
import uuid

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
FLOWS = (
    "verify_purchase_flow", "verify_conjuntos_montagem_flow",
    "verify_fabrication_order_flow", "verify_planning_flow", "verify_billing_flow",
)


def main() -> int:
    import main as runtime
    import pymysql

    database = "lugest_stress_" + uuid.uuid4().hex
    original_database = runtime.MYSQL_DB_NAME
    assert database != original_database and database.startswith("lugest_stress_")
    connect = pymysql.connect
    connection_class = pymysql.connections.Connection
    original_select_db = connection_class.select_db
    connection_options = dict(host=runtime.MYSQL_HOST, port=runtime.MYSQL_PORT,
                              user=runtime.MYSQL_USER, password=runtime.MYSQL_PASSWORD,
                              charset="utf8mb4", connect_timeout=5, read_timeout=60, write_timeout=60)
    report = {"database": database, "flows": [], "removed": False}
    created = False
    admin = None
    previous_environment = dict(os.environ)
    previous_base = runtime.BASE_DIR
    try:
        admin = connect(**connection_options)
        with admin.cursor() as cursor:
            # No IF NOT EXISTS: an existing database must never be reused.
            cursor.execute(f"CREATE DATABASE `{database}` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci")
        created = True
        print("Created disposable database:", database, flush=True)

        def guarded_connect(*args, **kwargs):
            if args or kwargs.get("database", kwargs.get("db")) != database:
                raise RuntimeError("Integration test blocked a connection outside its disposable database")
            if kwargs.get("host", runtime.MYSQL_HOST) != connection_options["host"]:
                raise RuntimeError("Integration test blocked a different server")
            return connect(**kwargs)

        def guarded_select_db(connection, name):
            if isinstance(name, bytes):
                name = name.decode("utf-8")
            if name != database:
                raise RuntimeError("Integration test blocked a database switch")
            return original_select_db(connection, name)

        pymysql.connect = guarded_connect
        connection_class.select_db = guarded_select_db
        installer = runpy.run_path(str(ROOT / "mysql/install_lugest_mysql.py"))
        connection = guarded_connect(database=database, **connection_options)
        try:
            with connection.cursor() as cursor:
                for sql in installer["sanitize_base_schema"]((ROOT / "mysql/lugest.sql").read_text(encoding="utf-8-sig")):
                    cursor.execute(sql)
            connection.commit()
        finally:
            connection.close()
        with tempfile.TemporaryDirectory(prefix="lugest-integration-") as temporary:
            work = Path(temporary)
            runtime.BASE_DIR = str(work)
            runtime.MYSQL_DB_NAME = database
            runtime._MYSQL_SCHEMA_SYNCED = False
            runtime._ASYNC_SAVE_ENABLED = False
            for key, value in {
                "LUGEST_DB_NAME": database,
                "LUGEST_USER_DATA_DIR": str(work / "user"),
                "LUGEST_MACHINE_DATA_DIR": str(work / "machine"),
                "LUGEST_SHARED_STORAGE_ROOT": str(work / "shared"),
                "QT_QPA_PLATFORM": "offscreen",
            }.items():
                os.environ[key] = value
            # Reuse the existing synthetic-data builder; it reads only this database.
            spec = importlib.util.spec_from_file_location("isolated_seed_driver", ROOT / "scripts/stress_test_isolated.py")
            driver = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(driver)
            args = driver.parse_args(["--clients", "4", "--orders", "6", "--completed", "2",
                                      "--purchase-notes", "2", "--products", "6", "--materials", "6",
                                      "--expeditions", "2", "--transports", "2", "--invoices", "2"])
            data, _ = driver._generate_dataset(args, "INTEGRATION")
            runtime.save_data(data, force=True)
            for name in FLOWS:
                started = time.monotonic()
                result = {"name": name}
                print("Running", name, flush=True)
                try:
                    entry = runpy.run_path(str(ROOT / "scripts" / (name + ".py")))
                    code = entry["main"]()
                    if code not in (0, None):
                        raise RuntimeError(f"Flow returned {code}")
                    result["status"] = "passed"
                except Exception as exc:
                    result.update(status="failed", error=str(exc), traceback=traceback.format_exc())
                    print(name, "FAILED:", str(exc), flush=True)
                result["seconds"] = round(time.monotonic() - started, 3)
                report["flows"].append(result)
    except Exception as exc:
        report["setup_error"] = str(exc)
        print("Isolated integration setup failed:", str(exc), flush=True)
    finally:
        pymysql.connect = connect
        connection_class.select_db = original_select_db
        runtime.MYSQL_DB_NAME = original_database
        runtime.BASE_DIR = previous_base
        os.environ.clear()
        os.environ.update(previous_environment)
        if created:
            try:
                with admin.cursor() as cursor:
                    cursor.execute(f"DROP DATABASE `{database}`")
                report["removed"] = True
            except Exception as exc:
                report["cleanup_error"] = str(exc)
        if admin is not None:
            admin.close()
        target = ROOT / "reports/integration"
        target.mkdir(parents=True, exist_ok=True)
        output = target / (database + ".json")
        output.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
        print("Report:", output, "removed:", report["removed"], flush=True)
    return 0 if len(report["flows"]) == len(FLOWS) and all(row["status"] == "passed" for row in report["flows"]) and report["removed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
