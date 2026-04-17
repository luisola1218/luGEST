from __future__ import annotations

import sys
from datetime import date, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_qt.services.main_bridge import LegacyBackend


def _create_planning_order(
    backend: LegacyBackend,
    *,
    client: str,
    note: str,
    delivery: str,
    material: str,
    esp: str,
    ref_ext: str,
    tempo_min: int,
) -> tuple[str, str, str]:
    enc = backend.order_create_or_update(
        {
            "cliente": client,
            "nota_cliente": note,
            "data_entrega": delivery,
            "tempo_estimado": tempo_min,
        }
    )
    numero = str(enc.get("numero", "") or "").strip()
    backend.order_piece_create_or_update(
        numero,
        {
            "material": material,
            "espessura": esp,
            "descricao": note,
            "ref_externa": ref_ext,
            "quantidade_pedida": 1,
            "operacoes": "Corte Laser + Embalamento",
            "guardar_ref": False,
        },
    )
    backend.order_espessura_set_time(numero, material, esp, tempo_min)
    return numero, material, esp


def _cleanup_order_blocks(backend: LegacyBackend, numero: str) -> None:
    for row in list(backend.ensure_data().get("plano", []) or []):
        if str(row.get("encomenda", "") or "").strip() != numero:
            continue
        block_id = str(row.get("id", "") or "").strip()
        if block_id:
            try:
                backend.planning_remove_block(block_id)
            except Exception:
                pass


def _cleanup_test_week(backend: LegacyBackend, week_start: date) -> None:
    week_days = {(week_start + timedelta(days=i)).isoformat() for i in range(6)}
    for row in list(backend.ensure_data().get("plano", []) or []):
        if str(row.get("data", "") or "").strip() not in week_days:
            continue
        block_id = str(row.get("id", "") or "").strip()
        if not block_id:
            continue
        try:
            backend.planning_remove_block(block_id)
        except Exception:
            pass


def main() -> int:
    backend = LegacyBackend()
    created_orders: list[str] = []
    manual_week_start = date(2032, 1, 5)
    auto_week_start = manual_week_start + timedelta(days=7)
    try:
        _cleanup_test_week(backend, manual_week_start)
        _cleanup_test_week(backend, auto_week_start)

        manual_num, manual_mat, manual_esp = _create_planning_order(
            backend,
            client="CL0001",
            note="VERIFY_PLANNING_MANUAL",
            delivery="2032-01-10",
            material="S275JR",
            esp="8",
            ref_ext="VERIFY-PLAN-MANUAL-001",
            tempo_min=60,
        )
        created_orders.append(manual_num)

        pending_rows = backend.planning_pending_rows(manual_num)
        manual_row = next((row for row in pending_rows if str(row.get("numero", "") or "").strip() == manual_num), None)
        if manual_row is None:
            raise RuntimeError(f"Encomenda manual nao apareceu no backlog: {pending_rows}")

        try:
            backend.planning_place_block(manual_num, manual_mat, manual_esp, manual_week_start.isoformat(), "12:30")
        except ValueError as exc:
            message = str(exc)
            if "bloquead" not in message.lower():
                raise RuntimeError(f"Bloqueio de almoco devolveu erro inesperado: {message}") from exc
        else:
            raise RuntimeError("Planeamento permitiu bloco dentro do almoco.")

        manual_block = backend.planning_place_block(manual_num, manual_mat, manual_esp, manual_week_start.isoformat(), "08:00")
        if str(manual_block.get("inicio", "") or "").strip() != "08:00":
            raise RuntimeError(f"Bloco manual colocado em hora errada: {manual_block}")

        try:
            backend.planning_shift_block(str(manual_block.get("id", "") or "").strip(), minutes_offset=270)
        except ValueError as exc:
            message = str(exc)
            if "bloquead" not in message.lower():
                raise RuntimeError(f"Deslocacao para almoco devolveu erro inesperado: {message}") from exc
        else:
            raise RuntimeError("Planeamento permitiu mover bloco para dentro do almoco.")

        auto_num_a, auto_mat_a, auto_esp_a = _create_planning_order(
            backend,
            client="CL0002",
            note="VERIFY_PLANNING_AUTO_A",
            delivery="2032-01-17",
            material="AISI 304L",
            esp="2",
            ref_ext="VERIFY-PLAN-AUTO-A",
            tempo_min=270,
        )
        auto_num_b, auto_mat_b, auto_esp_b = _create_planning_order(
            backend,
            client="CL0002",
            note="VERIFY_PLANNING_AUTO_B",
            delivery="2032-01-17",
            material="S235JR",
            esp="10",
            ref_ext="VERIFY-PLAN-AUTO-B",
            tempo_min=60,
        )
        created_orders.extend([auto_num_a, auto_num_b])

        ordered = []
        for numero, material, esp in (
            (auto_num_a, auto_mat_a, auto_esp_a),
            (auto_num_b, auto_mat_b, auto_esp_b),
        ):
            row = next(
                (
                    item
                    for item in backend.planning_pending_rows(numero)
                    if str(item.get("numero", "") or "").strip() == numero
                    and str(item.get("material", "") or "").strip() == material
                    and str(item.get("espessura", "") or "").strip() == esp
                ),
                None,
            )
            if row is None:
                raise RuntimeError(f"Encomenda {numero} nao apareceu no backlog para auto-planeamento.")
            ordered.append(row)

        placed = backend.planning_auto_plan(ordered, auto_week_start)
        placed_by_order: dict[str, list[dict]] = {}
        for row in placed:
            key = str(row.get("encomenda", "") or "").strip()
            placed_by_order.setdefault(key, []).append(row)
        if len(placed_by_order) != 2:
            raise RuntimeError(f"Auto-planeamento devolveu blocos inesperados: {placed}")
        if len(placed_by_order[auto_num_a]) != 1 or str(placed_by_order[auto_num_a][0].get("inicio", "") or "").strip() != "08:00":
            raise RuntimeError(f"Primeiro bloco auto-planeado ficou em hora errada: {placed_by_order[auto_num_a]}")
        if len(placed_by_order[auto_num_b]) != 1 or str(placed_by_order[auto_num_b][0].get("inicio", "") or "").strip() != "14:00":
            raise RuntimeError(f"Segundo bloco nao saltou o almoco: {placed_by_order[auto_num_b]}")

        print("planning-flow-ok", manual_num, auto_num_a, auto_num_b)
        return 0
    finally:
        try:
            _cleanup_test_week(backend, manual_week_start)
        except Exception:
            pass
        try:
            _cleanup_test_week(backend, auto_week_start)
        except Exception:
            pass
        for numero in created_orders:
            try:
                _cleanup_order_blocks(backend, numero)
            except Exception:
                pass
        for numero in reversed(created_orders):
            try:
                backend.order_remove(numero)
            except Exception:
                pass


if __name__ == "__main__":
    raise SystemExit(main())
