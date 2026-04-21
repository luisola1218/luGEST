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
    resource: str = "",
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
    if str(resource or "").strip():
        backend.order_espessura_set_operation_times(
            numero,
            material,
            esp,
            {"Corte Laser": tempo_min},
            {"Corte Laser": resource},
        )
    else:
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
    resource_week_start = auto_week_start + timedelta(days=7)
    packaging_week_start = resource_week_start + timedelta(days=7)
    try:
        _cleanup_test_week(backend, manual_week_start)
        _cleanup_test_week(backend, auto_week_start)
        _cleanup_test_week(backend, resource_week_start)
        _cleanup_test_week(backend, packaging_week_start)

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

        resource_num, resource_mat, resource_esp = _create_planning_order(
            backend,
            client="CL0002",
            note="VERIFY_PLANNING_RESOURCE_FILTER",
            delivery="2032-01-24",
            material="S355JR",
            esp="6",
            ref_ext="VERIFY-PLAN-RESOURCE",
            tempo_min=60,
            resource="Maquina 3030",
        )
        created_orders.append(resource_num)
        backend.ensure_data().setdefault("plano", []).append(
            backend._planning_make_block(
                resource_num,
                resource_mat,
                resource_esp,
                "Corte Laser",
                resource_week_start.isoformat(),
                480,
                60,
                posto="Laser",
            )
        )
        backend._save(force=True)

        visible_in_resource = [
            row
            for row in list(
                backend.planning_overview_data(
                    week_start=resource_week_start.isoformat(),
                    operation="Corte Laser",
                    resource="Maquina 3030",
                ).get("active", [])
                or []
            )
            if str(row.get("encomenda", "") or "").strip() == resource_num
        ]
        if len(visible_in_resource) != 1:
            raise RuntimeError(
                f"Bloco legado sem maquina explicita nao apareceu no recurso correto: {visible_in_resource}"
            )

        wrong_resource_pending = [
            row
            for row in list(backend.planning_pending_rows(operation="Corte Laser", resource="Maquina 3030") or [])
            if str(row.get("numero", "") or "").strip() == resource_num
        ]
        if wrong_resource_pending:
            raise RuntimeError(
                f"Backlog por recurso mostrou item ja planeado noutro recurso: {wrong_resource_pending}"
            )

        resource_blocks = [
            row
            for row in list(backend.ensure_data().get("plano", []) or [])
            if str(row.get("encomenda", "") or "").strip() == resource_num
        ]
        if len(resource_blocks) != 1:
            raise RuntimeError(f"Bloco de teste por recurso inesperado: {resource_blocks}")
        backend.planning_remove_block(str(resource_blocks[0].get("id", "") or "").strip())

        restored_pending = [
            row
            for row in list(backend.planning_pending_rows(operation="Corte Laser", resource="Maquina 3030") or [])
            if str(row.get("numero", "") or "").strip() == resource_num
        ]
        if len(restored_pending) != 1:
            raise RuntimeError(f"Remover bloco nao devolveu a encomenda ao backlog correto: {restored_pending}")

        pack_num_late, pack_mat_late, pack_esp_late = _create_planning_order(
            backend,
            client="CL0002",
            note="VERIFY_PLANNING_PACK_LATE",
            delivery="2032-01-31",
            material="S235JR",
            esp="5",
            ref_ext="VERIFY-PLAN-PACK-LATE",
            tempo_min=30,
        )
        pack_num_early, pack_mat_early, pack_esp_early = _create_planning_order(
            backend,
            client="CL0002",
            note="VERIFY_PLANNING_PACK_EARLY",
            delivery="2032-01-25",
            material="AISI 304L",
            esp="3",
            ref_ext="VERIFY-PLAN-PACK-EARLY",
            tempo_min=30,
        )
        created_orders.extend([pack_num_late, pack_num_early])
        backend.order_espessura_set_operation_times(
            pack_num_late,
            pack_mat_late,
            pack_esp_late,
            {"Corte Laser": 30, "Embalamento": 30},
            {"Corte Laser": "Maquina 3030", "Embalamento": "Embalamento"},
        )
        backend.order_espessura_set_operation_times(
            pack_num_early,
            pack_mat_early,
            pack_esp_early,
            {"Corte Laser": 30, "Embalamento": 30},
            {"Corte Laser": "Maquina 5030", "Embalamento": "Embalamento"},
        )
        pack_dates = [packaging_week_start + timedelta(days=i) for i in range(6)]
        pack_anchor = backend._planning_slot_datetime(packaging_week_start.isoformat(), 480)
        pack_result = backend._planning_schedule_followup_jobs(
            [
                {
                    "numero": pack_num_late,
                    "material": pack_mat_late,
                    "espessura": pack_esp_late,
                    "sequence": ["Embalamento"],
                    "index": 0,
                    "anchor_dt": pack_anchor,
                    "resource": "Embalamento",
                    "data_entrega": "2032-01-31",
                },
                {
                    "numero": pack_num_early,
                    "material": pack_mat_early,
                    "espessura": pack_esp_early,
                    "sequence": ["Embalamento"],
                    "index": 0,
                    "anchor_dt": pack_anchor,
                    "resource": "Embalamento",
                    "data_entrega": "2032-01-25",
                },
            ],
            pack_dates,
            initial_cursor_dt=pack_anchor,
        )
        pack_blocks = [
            row
            for row in list(pack_result.get("placed", []) or [])
            if str(row.get("operacao", "") or "").strip() == "Embalamento"
        ]
        if len(pack_blocks) < 2:
            raise RuntimeError(f"Fila de embalamento nao gerou blocos suficientes: {pack_blocks}")
        first_pack = min(pack_blocks, key=lambda row: (str(row.get("data", "") or ""), str(row.get("inicio", "") or "")))
        if str(first_pack.get("encomenda", "") or "").strip() != pack_num_early:
            raise RuntimeError(f"Embalamento nao priorizou o prazo de entrega mais cedo: {pack_blocks}")

        print("planning-flow-ok", manual_num, auto_num_a, auto_num_b, resource_num)
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
        try:
            _cleanup_test_week(backend, resource_week_start)
        except Exception:
            pass
        try:
            _cleanup_test_week(backend, packaging_week_start)
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
