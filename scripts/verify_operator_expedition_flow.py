from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_qt.services.main_bridge import LegacyBackend


def _reset_piece_ops(backend: LegacyBackend, enc_num: str, piece_id: str) -> None:
    reset_fn = getattr(backend.operador_actions, "_mysql_ops_reset_piece", None)
    if callable(reset_fn):
        reset_fn(enc_num, piece_id)


def _start_with_retry(
    backend: LegacyBackend,
    enc_num: str,
    piece_id: str,
    operator_name: str,
    operation: str,
) -> None:
    try:
        backend.operator_start_piece(enc_num, piece_id, operator_name, operation, "Geral")
    except ValueError as exc:
        message = str(exc)
        if "Operacao ocupada" not in message:
            raise
        _reset_piece_ops(backend, enc_num, piece_id)
        backend.operator_start_piece(enc_num, piece_id, operator_name, operation, "Geral")


def _raw_piece(
    backend: LegacyBackend,
    enc_num: str,
    piece_id: str,
    ref_interna: str = "",
    ref_externa: str = "",
) -> dict:
    enc = backend.get_encomenda_by_numero(enc_num) or {}
    for piece in list(backend.desktop_main.encomenda_pecas(enc)):
        if str(piece.get("id", "") or "").strip() == str(piece_id or "").strip():
            return piece
        if ref_interna and str(piece.get("ref_interna", "") or "").strip() == str(ref_interna).strip():
            return piece
        if ref_externa and str(piece.get("ref_externa", "") or "").strip() == str(ref_externa).strip():
            return piece
    return {}


def _finish_with_retry(
    backend: LegacyBackend,
    enc_num: str,
    piece_id: str,
    operator_name: str,
    operation: str,
    ok_qty: float,
) -> None:
    last_exc: Exception | None = None
    for _attempt in range(3):
        try:
            backend.operator_finish_piece(enc_num, piece_id, operator_name, ok_qty, 0, 0, operation, "Geral")
            return
        except ValueError as exc:
            message = str(exc)
            if ("Estado atual: Livre" not in message) and ("Inicia primeiro a operacao" not in message):
                raise
            last_exc = exc
            _reset_piece_ops(backend, enc_num, piece_id)
            _start_with_retry(backend, enc_num, piece_id, operator_name, operation)
    if last_exc is not None:
        raise last_exc


def main() -> int:
    backend = LegacyBackend()
    backend.save_branding_settings(
        {
            "guia_serie_id": "GT2026",
            "guia_validation_code": "TESTE-GT2026",
        }
    )
    created_number = ""
    guide_number = ""
    try:
        enc = backend.order_create_or_update(
            {
                "cliente": "CL0002",
                "nota_cliente": "VERIFY_OPERATOR_EXP",
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
                "descricao": "VERIFY OPERATOR -> EXPEDITION",
                "ref_externa": "VERIFY-OPP-EXP-001",
                "quantidade_pedida": 100,
                "operacoes": "Corte Laser + Embalamento",
                "guardar_ref": False,
            },
        )
        detail = backend.order_detail(created_number)
        if len(detail.get("pieces", [])) != 1:
            raise RuntimeError(f"Encomenda de teste sem peça válida: {detail}")
        piece = detail["pieces"][0]
        piece_id = str(piece.get("id", "") or "").strip()
        ref_interna = str(piece.get("ref_interna", "") or "").strip()
        ref_externa = str(piece.get("ref_externa", "") or "").strip()
        _reset_piece_ops(backend, created_number, piece_id)
        if backend.expedicao_available_pieces(created_number):
            raise RuntimeError("Peça não deveria estar disponível antes da produção.")

        _start_with_retry(backend, created_number, piece_id, "admin", "Corte Laser")
        _finish_with_retry(backend, created_number, piece_id, "admin", "Corte Laser", 10)
        raw_piece = _raw_piece(backend, created_number, piece_id, ref_interna, ref_externa)
        laser_row = next((op for op in list(raw_piece.get("operacoes_fluxo", []) or []) if str(op.get("nome", "") or "").strip() == "Corte Laser"), {})
        if str(laser_row.get("estado", "") or "").strip() != "Incompleta":
            raise RuntimeError(f"Estado parcial da operação perdeu-se após fecho parcial: {laser_row}")
        backend.flush_pending_save(force=True)
        backend.drain_async_saves(timeout_sec=10.0)
        backend.reload()
        raw_piece = _raw_piece(backend, created_number, piece_id, ref_interna, ref_externa)
        piece_id = str(raw_piece.get("id", "") or piece_id).strip()
        laser_row = next((op for op in list(raw_piece.get("operacoes_fluxo", []) or []) if str(op.get("nome", "") or "").strip() == "Corte Laser"), {})
        if str(laser_row.get("estado", "") or "").strip() != "Incompleta":
            raise RuntimeError(f"Estado parcial da operação não sobreviveu ao reload: {laser_row}")
        if backend.expedicao_available_pieces(created_number):
            raise RuntimeError("Peça não deveria estar disponível antes do embalamento.")

        _start_with_retry(backend, created_number, piece_id, "admin", "Embalamento")
        _finish_with_retry(backend, created_number, piece_id, "admin", "Embalamento", 10)

        available = backend.expedicao_available_pieces(created_number)
        if len(available) != 1:
            raise RuntimeError(f"Peça não ficou disponível para expedição: {available}")
        if abs(float(available[0].get("disponivel_num", 0) or 0) - 10.0) > 1e-9:
            raise RuntimeError(f"Quantidade disponível incorreta: {available}")

        pending_orders = backend.expedicao_pending_orders(created_number, "Todas")
        if not any(str(row.get("numero", "") or "").strip() == created_number for row in pending_orders):
            raise RuntimeError(f"Encomenda não apareceu na lista de expedição: {pending_orders}")

        line = dict(available[0])
        line["peca_id"] = piece_id
        line["qtd"] = 10.0
        line["peso"] = 0.0
        line["unid"] = "UN"
        defaults = backend.expedicao_defaults_for_order(created_number)
        detail = backend.expedicao_emit_off(created_number, [line], defaults)
        guide_number = str(detail.get("numero", "") or "").strip()
        if not guide_number:
            raise RuntimeError(f"Guia parcial sem número: {detail}")

        enc = backend.get_encomenda_by_numero(created_number) or {}
        if str(enc.get("estado_expedicao", "") or "").strip() != "Parcialmente expedida":
            raise RuntimeError(f"Estado de expedição deveria ser parcial após expedir 10/100: {enc.get('estado_expedicao')}")
        if backend.expedicao_available_pieces(created_number):
            raise RuntimeError("Não deveria sobrar disponibilidade depois de expedir toda a tranche embalada.")

        _start_with_retry(backend, created_number, piece_id, "admin", "Corte Laser")
        _finish_with_retry(backend, created_number, piece_id, "admin", "Corte Laser", 5)
        if backend.expedicao_available_pieces(created_number):
            raise RuntimeError("A nova tranche não deveria ficar disponível antes de passar no embalamento.")

        _start_with_retry(backend, created_number, piece_id, "admin", "Embalamento")
        _finish_with_retry(backend, created_number, piece_id, "admin", "Embalamento", 5)
        available = backend.expedicao_available_pieces(created_number)
        if len(available) != 1 or abs(float(available[0].get("disponivel_num", 0) or 0) - 5.0) > 1e-9:
            raise RuntimeError(f"A segunda tranche parcial não ficou disponível corretamente: {available}")

        print("operator-expedition-ok", created_number, piece_id, available[0].get("disponivel_num"))
        return 0
    finally:
        data = backend.ensure_data()
        if guide_number:
            data["expedicoes"] = [
                row for row in list(data.get("expedicoes", []) or [])
                if str(row.get("numero", "") or "").strip() != guide_number
            ]
        if created_number:
            try:
                backend.order_remove(created_number)
            except Exception:
                pass
        backend._save(force=True)


if __name__ == "__main__":
    raise SystemExit(main())
