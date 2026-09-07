"""Client page -> backend -> service -> repository with isolated persistence.

No main.py import, real MySQL, external files, messages or network access.
"""
from __future__ import annotations

from copy import deepcopy
from dataclasses import fields
import os
from pathlib import Path
import sys
import tempfile
from types import SimpleNamespace
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")


def main():
    from PySide6.QtWidgets import QApplication, QMessageBox
    from lugest_qt.services.legacy_runtime import LegacyRuntime
    from lugest_qt.services.main_bridge import LegacyBackend
    from lugest_qt.ui.pages.partners_pages import ClientsPage

    app = QApplication.instance() or QApplication(["verify-clients-integration"])
    with tempfile.TemporaryDirectory(prefix="lugest-clients-") as temporary:
        persisted = {"clientes": [], "encomendas": [], "orcamentos": []}
        failure = [False]
        writes = []

        def save_data(data, **kwargs):
            if failure[0]:
                raise OSError("simulated database failure")
            persisted.clear()
            persisted.update(deepcopy(data))
            writes.append(kwargs)

        def next_code(data):
            used = {row["codigo"] for row in data.get("clientes", [])}
            index = 1
            while f"CL{index:04d}" in used:
                index += 1
            return f"CL{index:04d}"

        def normalize(value, data):
            return value if isinstance(value, dict) else {"codigo": str(value)}

        desktop = SimpleNamespace(BASE_DIR=temporary, load_data=lambda: deepcopy(persisted),
                                  save_data=save_data, next_cliente_codigo=next_code,
                                  _normalize_orc_cliente=normalize, now_iso=lambda: "2026-09-07T12:00:00",
                                  _ASYNC_SAVE_ENABLED=False)
        dependencies = {field.name: SimpleNamespace() for field in fields(LegacyRuntime)}
        dependencies["desktop_main"] = desktop
        backend = LegacyBackend(runtime=LegacyRuntime(**dependencies))
        errors = []
        with patch("socket.create_connection", side_effect=AssertionError("Unexpected network")), \
             patch.object(QMessageBox, "critical", side_effect=lambda *args: errors.append(args[-1])), \
             patch.object(QMessageBox, "question", return_value=QMessageBox.Yes):
            page = ClientsPage(backend)
            page.client_name_edit.setText("  Construções Árvore  ")
            page.client_email_edit.setText("teste@example.invalid")
            page._save_client()
            assert not errors and len(persisted["clientes"]) == 1
            assert persisted["clientes"][0]["nome"] == "Construções Árvore"
            assert len(backend.client_rows("arvore construcoes")) == 1
            assert page.table.rowCount() == 1
            assert backend.order_clients()[0]["codigo"] == "CL0001"
            page.client_name_edit.setText("Alterado")
            page._save_client()
            assert persisted["clientes"][0]["nome"] == "Alterado"
            # A rejected save must not appear as successful in the data cache.
            failure[0] = True
            page.client_name_edit.setText("Nao gravado")
            page._save_client()
            assert errors and backend.client_rows()[0]["nome"] == "Alterado"
            failure[0] = False
            errors.clear()
            for bucket, value in (("encomendas", "CL0001"), ("orcamentos", {"codigo": "CL0001"})):
                backend.ensure_data()[bucket] = [{"cliente": value}]
                try:
                    backend.client_remove("CL0001")
                except ValueError:
                    pass
                else:
                    raise AssertionError(f"Client in {bucket} was removed")
                backend.ensure_data()[bucket] = []
            page._remove_client()
            assert not errors and not persisted["clientes"] and page.table.rowCount() == 0
            before = deepcopy(backend.ensure_data())
            try:
                backend.client_save({"nome": "   "})
            except ValueError:
                pass
            else:
                raise AssertionError("Blank name accepted")
            assert backend.ensure_data() == before
            # Synchronous direct-upsert path updates baseline only on success.
            def upsert(row):
                if failure[0]:
                    raise OSError("simulated upsert failure")
            desktop.mysql_upsert_cliente = upsert
            saved = backend.client_save({"codigo": "CL0099", "nome": "Direct"})
            saved["nome"] = "mutated result"
            baseline = deepcopy(backend._base_data_snapshot)
            failure[0] = True
            try:
                backend.client_save({"codigo": "CL0099", "nome": "Rejected"})
            except OSError:
                pass
            else:
                raise AssertionError("Upsert failure suppressed")
            assert backend.client_rows()[0]["nome"] == "Direct"
            assert backend._base_data_snapshot == baseline
            # Async mode retains the existing queue acceptance contract.
            failure[0] = False
            desktop._ASYNC_SAVE_ENABLED = True
            backend.client_save({"codigo": "CL0100", "nome": "Queued"})
            assert writes[-1]["force"] is False
            page.close()
            app.processEvents()
        assert "main" not in sys.modules
    print("clients-integration-ok ui-crud=yes search=yes references=yes failure-cache=yes async-dispatch=yes no-main=yes")


if __name__ == "__main__":
    main()
