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
    nok_qty: float,
) -> None:
    try:
        backend.operator_finish_piece(enc_num, piece_id, operator_name, ok_qty, nok_qty, 0, operation, "Geral")
    except ValueError as exc:
        message = str(exc)
        if ("Estado atual: Livre" not in message) and ("Inicia primeiro a operacao" not in message):
            raise
        _reset_piece_ops(backend, enc_num, piece_id)
        backend.operator_start_piece(enc_num, piece_id, operator_name, operation, "Geral")
        backend.operator_finish_piece(enc_num, piece_id, operator_name, ok_qty, nok_qty, 0, operation, "Geral")


def _run_operation(
    backend: LegacyBackend,
    enc_num: str,
    piece_id: str,
    operation: str,
    ok_qty: float,
    nok_qty: float,
) -> None:
    _start_with_retry(backend, enc_num, piece_id, "admin", operation)
    _finish_with_retry(backend, enc_num, piece_id, "admin", operation, ok_qty, nok_qty)


def _cleanup(backend: LegacyBackend, enc_num: str, guide_num: str) -> None:
    data = backend.ensure_data()
    if guide_num:
        data["expedicoes"] = [
            row for row in list(data.get("expedicoes", []) or [])
            if str(row.get("numero", "") or "").strip() != guide_num
        ]
    if enc_num:
        try:
            backend.order_remove(enc_num)
        except Exception:
            pass


def main() -> int:
    backend = LegacyBackend()
    backend.save_branding_settings(
        {
            "guia_serie_id": "GT2026",
            "guia_validation_code": "TESTE-GT2026",
        }
    )

    enc_nok = ""
    enc_mix = ""
    guide_mix = ""
    try:
        enc = backend.order_create_or_update(
            {
                "cliente": "CL0002",
                "nota_cliente": "VERIFY SHIPPING EDGE NOK",
                "data_entrega": "2026-03-31",
                "tempo_estimado": 30,
            }
        )
        enc_nok = str(enc.get("numero", "") or "").strip()
        backend.order_piece_create_or_update(
            enc_nok,
            {
                "material": "S275JR",
                "espessura": "8",
                "descricao": "VERIFY ALL NOK SHOULD NOT SHIP",
                "ref_externa": "VERIFY-SHIP-NOK-001",
                "quantidade_pedida": 10,
                "operacoes": "Corte Laser + Embalamento",
                "guardar_ref": False,
            },
        )
        piece = backend.order_detail(enc_nok)["pieces"][0]
        piece_id = str(piece.get("id", "") or "").strip()
        _reset_piece_ops(backend, enc_nok, piece_id)
        _run_operation(backend, enc_nok, piece_id, "Corte Laser", 0, 10)
        _run_operation(backend, enc_nok, piece_id, "Embalamento", 0, 10)
        available = backend.expedicao_available_pieces(enc_nok)
        if available:
            raise RuntimeError(f"Peça 100% NOK apareceu como expedível: {available}")

        enc = backend.order_create_or_update(
            {
                "cliente": "CL0002",
                "nota_cliente": "VERIFY SHIPPING EDGE MIX",
                "data_entrega": "2026-03-31",
                "tempo_estimado": 30,
            }
        )
        enc_mix = str(enc.get("numero", "") or "").strip()
        backend.order_piece_create_or_update(
            enc_mix,
            {
                "material": "S275JR",
                "espessura": "8",
                "descricao": "VERIFY MIX OK NOK",
                "ref_externa": "VERIFY-SHIP-MIX-001",
                "quantidade_pedida": 10,
                "operacoes": "Corte Laser + Embalamento",
                "guardar_ref": False,
            },
        )
        piece = backend.order_detail(enc_mix)["pieces"][0]
        piece_id = str(piece.get("id", "") or "").strip()
        _reset_piece_ops(backend, enc_mix, piece_id)
        _run_operation(backend, enc_mix, piece_id, "Corte Laser", 8, 2)
        _run_operation(backend, enc_mix, piece_id, "Embalamento", 8, 2)
        available = backend.expedicao_available_pieces(enc_mix)
        if len(available) != 1 or abs(float(available[0].get("disponivel_num", 0) or 0) - 8.0) > 1e-9:
            raise RuntimeError(f"Quantidade mista OK/NOK disponível incorreta: {available}")
        line = dict(available[0])
        line["peca_id"] = piece_id
        line["qtd"] = 8.0
        line["peso"] = 0.0
        line["unid"] = "UN"
        detail = backend.expedicao_emit_off(enc_mix, [line], backend.expedicao_defaults_for_order(enc_mix))
        guide_mix = str(detail.get("numero", "") or "").strip()
        enc_after = backend.get_encomenda_by_numero(enc_mix) or {}
        if str(enc_after.get("estado_expedicao", "") or "").strip() != "Totalmente expedida":
            raise RuntimeError(f"Estado de expedição incorreto após expedir todo o bom da peça: {enc_after}")

        print("shipping-edge-cases-ok", enc_nok, enc_mix, guide_mix)
        return 0
    finally:
        _cleanup(backend, enc_nok, "")
        _cleanup(backend, enc_mix, guide_mix)
        backend._save(force=True)


if __name__ == "__main__":
    raise SystemExit(main())
