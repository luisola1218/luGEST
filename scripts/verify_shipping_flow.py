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
        if "Operacao ocupada" not in str(exc):
            raise
        _reset_piece_ops(backend, enc_num, piece_id)
        backend.operator_start_piece(enc_num, piece_id, operator_name, operation, "Geral")


def _finish_with_retry(
    backend: LegacyBackend,
    enc_num: str,
    piece_id: str,
    operator_name: str,
    operation: str,
    ok_qty: float,
) -> None:
    try:
        backend.operator_finish_piece(enc_num, piece_id, operator_name, ok_qty, 0, 0, operation, "Geral")
    except ValueError as exc:
        message = str(exc)
        if ("Estado atual: Livre" not in message) and ("Inicia primeiro a operacao" not in message):
            raise
        _reset_piece_ops(backend, enc_num, piece_id)
        backend.operator_start_piece(enc_num, piece_id, operator_name, operation, "Geral")
        backend.operator_finish_piece(enc_num, piece_id, operator_name, ok_qty, 0, 0, operation, "Geral")


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
                "nota_cliente": "VERIFY_SHIPPING_FLOW",
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
                "descricao": "VERIFY SHIPPING FLOW",
                "ref_externa": "VERIFY-SHIP-001",
                "quantidade_pedida": 5,
                "operacoes": "Corte Laser + Embalamento",
                "guardar_ref": False,
            },
        )
        detail = backend.order_detail(created_number)
        piece = detail["pieces"][0]
        piece_id = str(piece.get("id", "") or "").strip()
        _reset_piece_ops(backend, created_number, piece_id)

        _start_with_retry(backend, created_number, piece_id, "admin", "Corte Laser")
        _finish_with_retry(backend, created_number, piece_id, "admin", "Corte Laser", 5)
        _start_with_retry(backend, created_number, piece_id, "admin", "Embalamento")
        _finish_with_retry(backend, created_number, piece_id, "admin", "Embalamento", 5)

        available = backend.expedicao_available_pieces(created_number)
        if len(available) != 1:
            raise RuntimeError(f"Peça não disponível para guia: {available}")
        line = dict(available[0])
        line["peca_id"] = piece_id
        line["qtd"] = 5.0
        line["peso"] = 0.0
        line["unid"] = "UN"
        defaults = backend.expedicao_defaults_for_order(created_number)
        detail = backend.expedicao_emit_off(created_number, [line], defaults)
        guide_number = str(detail.get("numero", "") or "").strip()
        if not guide_number:
            raise RuntimeError(f"Guia sem número: {detail}")

        history_rows = backend.expedicao_rows(created_number)
        if not any(str(row.get("numero", "") or "").strip() == guide_number for row in history_rows):
            raise RuntimeError(f"Guia não apareceu no histórico: {history_rows}")
        guide_detail = backend.expedicao_detail(guide_number)
        if len(guide_detail.get("lines", [])) != 1:
            raise RuntimeError(f"Detalhe da guia inválido: {guide_detail}")
        if backend.expedicao_available_pieces(created_number):
            raise RuntimeError("Peça continua disponível após emissão total da guia.")

        print("shipping-flow-ok", created_number, guide_number)
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
