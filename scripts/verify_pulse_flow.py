from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from impulse_mobile_api.app.services import pulse_runtime
from lugest_qt.services.main_bridge import LegacyBackend


def _assert_no_data_scope() -> None:
    data = pulse_runtime.get_dashboard(period="7 dias", year="2099", encomenda="Todas", visao="Todas", origem="Ambos")
    summary = dict(data.get("summary", {}) or {})
    if any(float(summary.get(key, 0) or 0) != 0.0 for key in ("oee", "disponibilidade", "performance", "qualidade")):
        raise RuntimeError(f"Resumo sem dados devia estar a zero: {summary}")
    andon = dict(summary.get("andon", {}) or {})
    if any(int(andon.get(key, 0) or 0) != 0 for key in ("prod", "setup", "espera", "stop")):
        raise RuntimeError(f"Andon sem dados devia estar a zero: {andon}")
    alerts = str(summary.get("alerts", "") or "")
    if "Sem dados operacionais" not in alerts or "Qualidade abaixo da meta" in alerts:
        raise RuntimeError(f"Alertas sem dados incoerentes: {alerts}")


def _assert_running_scope() -> None:
    backend = LegacyBackend()
    created_number = ""
    try:
        enc = backend.order_create_or_update(
            {
                "cliente": "CL0002",
                "nota_cliente": "VERIFY_PULSE_FLOW",
                "data_entrega": "2026-03-31",
                "tempo_estimado": 30,
            }
        )
        created_number = str(enc.get("numero", "") or "").strip()
        backend.order_piece_create_or_update(
            created_number,
            {
                "material": "S275JR",
                "espessura": "8",
                "descricao": "VERIFY PULSE FLOW",
                "ref_externa": "VERIFY-PULSE-001",
                "quantidade_pedida": 5,
                "operacoes": "Corte Laser + Embalamento",
                "guardar_ref": False,
            },
        )
        detail = backend.order_detail(created_number)
        piece = detail["pieces"][0]
        piece_id = str(piece.get("id", "") or "").strip()

        backend.operator_start_piece(created_number, piece_id, "admin", "Corte Laser", "Geral")
        data = pulse_runtime.get_dashboard(period="7 dias", year="2026", encomenda=created_number, visao="Todas", origem="Ambos")
        summary = dict(data.get("summary", {}) or {})
        running = list(data.get("running", []) or [])
        if not running or int(summary.get("pecas_em_curso", 0) or 0) <= 0:
            raise RuntimeError(f"Pulse nao refletiu a operacao em curso: {data}")
    finally:
        if created_number:
            try:
                backend.order_remove(created_number)
            except Exception:
                pass


def _assert_plan_delay_scope() -> None:
    backend = LegacyBackend()
    created_number = ""
    created_block_id = ""
    item_key = ""
    try:
        enc = backend.order_create_or_update(
            {
                "cliente": "CL0002",
                "nota_cliente": "VERIFY_PLAN_DELAY",
                "data_entrega": "2026-04-03",
                "tempo_estimado": 30,
            }
        )
        created_number = str(enc.get("numero", "") or "").strip()
        backend.order_piece_create_or_update(
            created_number,
            {
                "material": "S235JR",
                "espessura": "5",
                "descricao": "VERIFY PLAN DELAY",
                "ref_externa": "VERIFY-PLAN-DELAY-001",
                "quantidade_pedida": 3,
                "operacoes": "Corte Laser + Embalamento",
                "guardar_ref": False,
            },
        )
        backend.order_espessura_set_time(created_number, "S235JR", "5", 30)
        block = backend._planning_make_block(
            created_number,
            "S235JR",
            "5",
            "Corte Laser",
            date.today().isoformat(),
            8 * 60,
            30,
            posto="Maquina 3030",
        )
        created_block_id = str(block.get("id", "") or "").strip()
        backend.ensure_data().setdefault("plano", []).append(block)
        backend._save(force=True)

        data = pulse_runtime.get_dashboard(
            period="Hoje",
            year="2026",
            encomenda=created_number,
            visao="Todas",
            origem="Ambos",
        )
        plan_delay = dict(data.get("plan_delay", {}) or {})
        if int(plan_delay.get("open_count", 0) or 0) <= 0:
            raise RuntimeError(f"Pulse nao sinalizou o atraso face ao planeamento: {data}")
        items = list(plan_delay.get("items", []) or [])
        if not items:
            raise RuntimeError(f"Sem linhas de atraso face ao planeamento: {data}")
        item_key = str(items[0].get("item_key", "") or "").strip()
        if not item_key:
            raise RuntimeError(f"Sem item_key no atraso ao planeamento: {items[0]}")

        pulse_runtime.set_plan_delay_reason(item_key, "Mudança de prioridade / urgência")
        justified = pulse_runtime.get_dashboard(
            period="Hoje",
            year="2026",
            encomenda=created_number,
            visao="Todas",
            origem="Ambos",
        )
        justified_delay = dict(justified.get("plan_delay", {}) or {})
        if int(justified_delay.get("open_count", 0) or 0) != 0:
            raise RuntimeError(f"O atraso devia ficar retirado do estado ativo após justificar: {justified}")
        if int(justified_delay.get("acknowledged_count", 0) or 0) <= 0:
            raise RuntimeError(f"O atraso justificado nao ficou marcado: {justified}")

        pulse_runtime.clear_plan_delay_reason(item_key)
        restored = pulse_runtime.get_dashboard(
            period="Hoje",
            year="2026",
            encomenda=created_number,
            visao="Todas",
            origem="Ambos",
        )
        restored_delay = dict(restored.get("plan_delay", {}) or {})
        if int(restored_delay.get("open_count", 0) or 0) <= 0:
            raise RuntimeError(f"O atraso devia voltar a ficar pendente depois de reativar o aviso: {restored}")
    finally:
        if item_key:
            try:
                pulse_runtime.clear_plan_delay_reason(item_key)
            except Exception:
                pass
        if created_block_id:
            try:
                backend.planning_remove_block(created_block_id)
            except Exception:
                pass
        if created_number:
            try:
                backend.order_remove(created_number)
            except Exception:
                pass


def main() -> int:
    _assert_no_data_scope()
    _assert_running_scope()
    _assert_plan_delay_scope()
    print("pulse-flow-ok")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
