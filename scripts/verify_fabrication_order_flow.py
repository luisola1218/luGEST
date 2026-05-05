from __future__ import annotations

import copy
import sys
import tempfile
import uuid
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_qt.services.main_bridge import LegacyBackend


def _assert(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def _client_code(backend: LegacyBackend, token: str) -> tuple[str, str]:
    data = backend.ensure_data()
    for row in list(data.get("clientes", []) or []):
        code = str((row or {}).get("codigo", "") or "").strip()
        if code:
            return code, ""
    code = f"CLMES{token[:4]}"
    data.setdefault("clientes", []).append({"codigo": code, "nome": f"Cliente MES {token}", "nif": "", "morada": ""})
    return code, code


def main() -> int:
    backend = LegacyBackend()
    data = backend.ensure_data()
    token = uuid.uuid4().hex[:8].upper()
    model_code = f"CJ-MES-{token}"
    order_num = ""
    created_client = ""
    pdf_path = Path(tempfile.gettempdir()) / f"lugest_verify_of_{token}.pdf"

    seq_snapshot = copy.deepcopy(dict(data.get("seq", {}) or {}))
    of_seq_snapshot = data.get("of_seq", 1)
    opp_seq_snapshot = data.get("opp_seq", 1)
    orc_seq_snapshot = data.get("orc_seq", 1)

    client_code, created_client = _client_code(backend, token)

    try:
        backend.assembly_model_save(
            {
                "codigo": model_code,
                "descricao": f"Conjunto MES {token}",
                "ativo": True,
                "itens": [
                    {
                        "tipo_item": "peca_fabricada",
                        "ref_externa": f"CHAPA-{token}",
                        "descricao": "Chapa laser",
                        "material": "S235JR",
                        "espessura": "5",
                        "operacao": "Corte Laser + Quinagem",
                        "qtd": 2,
                        "preco_unit": 10,
                    },
                    {
                        "tipo_item": "servico_montagem",
                        "descricao": "Montagem final",
                        "produto_unid": "SV",
                        "qtd": 1,
                        "preco_unit": 5,
                    },
                ],
            }
        )

        detail = backend.order_create_or_update(
            {
                "cliente": client_code,
                "tipo_encomenda": "Interna (produção)",
                "data_entrega": "2026-05-30",
                "tempo_estimado": 0,
                "nota_cliente": f"VERIFY_MES_{token}",
            }
        )
        order_num = str(detail.get("numero", "") or "").strip()
        _assert(order_num, "Encomenda nao foi criada.")
        _assert(str(detail.get("tipo_encomenda", "")) == "Interna (produção)", f"Tipo da encomenda incorreto: {detail}")
        _assert(str(detail.get("of_codigo", "")).startswith("OF-"), f"OF nao gerada: {detail}")

        detail = backend.order_import_model(order_num, model_code, 1)
        _assert(int(detail.get("imported_pieces", 0) or 0) == 1, f"Pecas importadas inesperadas: {detail}")
        _assert(int(detail.get("imported_items", 0) or 0) == 1, f"Itens de montagem inesperados: {detail}")

        backend.order_piece_create_or_update(
            order_num,
            {
                "ref_externa": f"PERFIL-{token}",
                "descricao": "Perfil cortado a medida",
                "tipo_material": "PERFIL",
                "material": "S355",
                "subtipo_material": "S355",
                "perfil_tipo": "IPE",
                "perfil_tamanho": "500",
                "dimensao": "IPE 500",
                "espessura": "",
                "operacoes": "Serralharia",
                "quantidade_pedida": 3,
                "ficheiros": [str(pdf_path.with_suffix(".step"))],
            },
        )

        detail = backend.order_detail(order_num)
        pieces = list(detail.get("pieces", []) or [])
        _assert(len(pieces) == 2, f"Numero de pecas invalido: {pieces}")
        of_code = str(detail.get("of_codigo", "") or "")
        _assert(of_code and all(str(row.get("of", "") or "") == of_code for row in pieces), f"OF por encomenda nao esta consistente: {pieces}")
        opps = [str(row.get("opp", "") or "") for row in pieces]
        _assert(len(set(opps)) == len(opps) and all(opp.startswith("OPP-") for opp in opps), f"OPP invalidas: {opps}")
        _assert(
            any(str(row.get("tipo_material", "")) == "PERFIL" and str(row.get("dimensao", "")) == "IPE 500" and str(row.get("perfil_tipo", "")) == "IPE" for row in pieces),
            "Perfil/dimensao nao foram guardados.",
        )

        rendered = backend.order_fabrication_pdf(order_num, pdf_path)
        _assert(rendered.exists() and rendered.stat().st_size > 0, f"PDF OF vazio: {rendered}")

        of_scan = backend.operator_scan_code(of_code, current_posto="Corte Laser")
        _assert(str(of_scan.get("tipo", "")) == "OF", f"Scan OF invalido: {of_scan}")
        opp_scan = backend.operator_scan_code(opps[0], current_posto="Corte Laser")
        _assert(str(opp_scan.get("tipo", "")) == "OPP", f"Scan OPP invalido: {opp_scan}")
        _assert(str(opp_scan.get("operacao", "") or ""), f"Scan OPP nao escolheu operacao: {opp_scan}")

        print(f"fabrication-order-flow-ok {order_num} {of_code} {' '.join(opps)}")
        return 0
    finally:
        cleanup_errors: list[str] = []
        try:
            if order_num:
                backend.order_remove(order_num)
        except Exception as exc:
            cleanup_errors.append(str(exc))
        try:
            backend.assembly_model_remove(model_code)
        except Exception as exc:
            cleanup_errors.append(str(exc))
        if created_client:
            data["clientes"] = [row for row in list(data.get("clientes", []) or []) if str((row or {}).get("codigo", "") or "") != created_client]
        data["seq"] = seq_snapshot
        data["of_seq"] = of_seq_snapshot
        data["opp_seq"] = opp_seq_snapshot
        data["orc_seq"] = orc_seq_snapshot
        try:
            backend._save(force=True)
        except Exception as exc:
            cleanup_errors.append(str(exc))
        try:
            if pdf_path.exists():
                pdf_path.unlink()
        except Exception:
            pass
        if cleanup_errors:
            print("cleanup-warnings " + " | ".join(cleanup_errors))


if __name__ == "__main__":
    raise SystemExit(main())
