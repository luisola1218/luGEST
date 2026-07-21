from __future__ import annotations

import copy
import atexit
import os
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_qt.services.legacy_backend import LegacyBackend
from lugest_qt.ui.pages.opp_page import OppPage
from PySide6.QtWidgets import QApplication


def main() -> int:
    backend = LegacyBackend()
    live_snapshot = copy.deepcopy(backend.ensure_data())
    original_save = backend._save
    restored = False

    def restore_live_data() -> None:
        nonlocal restored
        if restored:
            return
        restored = True
        backend._replace_data_cache(copy.deepcopy(live_snapshot))
        backend._save = original_save
        backend._save(force=True, audit=False, blocking=True)

    atexit.register(restore_live_data)
    sandbox = copy.deepcopy(live_snapshot)
    backend._replace_data_cache(sandbox)
    backend._save = lambda *args, **kwargs: None

    sandbox.setdefault("clientes", []).append(
        {"codigo": "CLPORT", "nome": "Cliente Portfolio Industrial"}
    )
    order = backend.order_create_or_update(
        {
            "cliente": "CLPORT",
            "nota_cliente": "VERIFY_OPP_PORTFOLIO",
            "data_entrega": "2026-09-30",
            "tempo_estimado": 60,
        }
    )
    order_number = str(order.get("numero", "") or "").strip()
    backend.order_piece_create_or_update(
        order_number,
        {
            "material": "S235JR",
            "espessura": "3",
            "descricao": "Estrutura de teste portfolio",
            "ref_externa": "PORT-001",
            "quantidade_pedida": 10,
            "preco_unit": 100,
            "operacoes": "Corte Laser + Quinagem",
            "guardar_ref": False,
        },
    )
    stored_order = backend.get_encomenda_by_numero(order_number)
    if not isinstance(stored_order, dict):
        raise RuntimeError("Encomenda de teste nao encontrada.")
    stored_order["numero_orcamento"] = "ORC-PORT-001"
    pieces = list(backend.desktop_main.encomenda_pecas(stored_order) or [])
    if len(pieces) != 1:
        raise RuntimeError(f"Pecas de teste inesperadas: {pieces}")
    pieces[0]["produzido_ok"] = 6
    pieces[0]["qtd_expedida"] = 4
    pieces[0]["estado"] = "Em producao"

    sandbox.setdefault("orcamentos", []).append(
        {
            "numero": "ORC-PORT-001",
            "numero_encomenda": order_number,
            "estado": "Aprovado",
            "data": "2026-07-21",
            "total": 1230.0,
            "cliente": {"codigo": "CLPORT", "nome": "Cliente Portfolio Industrial"},
            "linhas": [],
        }
    )
    sandbox.setdefault("faturacao", []).append(
        {
            "numero": "FAT-PORT-001",
            "origem": "Orcamento",
            "orcamento_numero": "ORC-PORT-001",
            "encomenda_numero": order_number,
            "cliente_codigo": "CLPORT",
            "cliente_nome": "Cliente Portfolio Industrial",
            "data_venda": "2026-07-21",
            "valor_venda_manual": 0,
            "faturas": [
                {
                    "id": "INV-PORT-001",
                    "numero_fatura": "FT PORT/1",
                    "valor_total": 615.0,
                    "data_emissao": "2026-07-21",
                    "estado": "Emitida",
                }
            ],
            "pagamentos": [
                {
                    "id": "PAY-PORT-001",
                    "fatura_id": "INV-PORT-001",
                    "valor": 200.0,
                    "data_pagamento": "2026-07-21",
                }
            ],
        }
    )

    payload = backend.opp_client_portfolio("CLPORT", "2026", "")
    totals = dict(payload.get("totals", {}) or {})
    orders = list(payload.get("orders", []) or [])
    if len(orders) != 1:
        raise RuntimeError(f"Portfolio devia conter uma encomenda: {orders}")
    row = orders[0]
    expected = {
        "adjudicado": 1230.0,
        "faturado": 615.0,
        "recebido": 200.0,
        "por_faturar": 615.0,
        "saldo_receber": 415.0,
        "qtd_plan": 10.0,
        "qtd_prod": 6.0,
        "qtd_exp": 4.0,
        "progress": 60.0,
    }
    for field, value in expected.items():
        actual = round(float(row.get(field, 0) or 0), 2)
        if actual != value:
            raise RuntimeError(f"{field} incorreto: esperado {value}, obtido {actual}")
    for field in ("adjudicado", "faturado", "recebido", "por_faturar", "saldo_receber"):
        if round(float(totals.get(field, 0) or 0), 2) != expected[field]:
            raise RuntimeError(f"Total {field} incorreto: {totals}")
    if len(list(row.get("pieces", []) or [])) != 1:
        raise RuntimeError("Detalhe produtivo da encomenda nao inclui a OPP.")
    if backend.opp_client_portfolio("CLPORT", "2025", "").get("orders"):
        raise RuntimeError("Filtro anual incluiu uma encomenda do ano errado.")
    if len(backend.opp_client_portfolio("CLPORT", "2026", "PORT-001").get("orders", [])) != 1:
        raise RuntimeError("Pesquisa por referencia nao encontrou a encomenda.")

    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    app = QApplication.instance() or QApplication(["verify-opp-client-portfolio"])
    page = OppPage(backend)
    page.resize(1600, 900)
    page.show()
    page.refresh()
    page.portfolio_client_combo.setCurrentIndex(page.portfolio_client_combo.findData("CLPORT"))
    app.processEvents()
    if page.portfolio_orders_table.rowCount() != 1:
        raise RuntimeError("A grelha de encomendas nao apresentou a encomenda do cliente.")
    if page.portfolio_pieces_table.rowCount() != 1:
        raise RuntimeError("A grelha de producao nao apresentou a OPP da encomenda.")
    screenshot_path = Path(tempfile.gettempdir()) / "lugest_opp_client_portfolio.png"
    if not page.grab().save(str(screenshot_path)):
        raise RuntimeError("Nao foi possivel guardar a captura de validacao da carteira OPP.")
    page.close()

    restore_live_data()
    atexit.unregister(restore_live_data)
    print("opp-client-portfolio-ok", order_number, totals, screenshot_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
