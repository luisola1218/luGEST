from __future__ import annotations

from copy import deepcopy
from datetime import date, timedelta
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_qt.services.main_bridge import LegacyBackend


def _make_order(
    numero: str,
    cliente: str,
    material: str,
    espessura: str,
    entrega: str,
    lote: str,
    descricao: str,
    qtd: int,
    *,
    posto_trabalho: str = "Maquina 3030",
) -> dict:
    return {
        "numero": numero,
        "cliente": cliente,
        "posto_trabalho": posto_trabalho,
        "estado": "Preparacao",
        "data_entrega": entrega,
        "nota_cliente": "",
        "tempo_estimado": 60,
        "Observacoes": "",
        "cativar": False,
        "reservas": [],
        "numero_orcamento": "",
        "montagem_itens": [],
        "materiais": [
            {
                "material": material,
                "estado": "Preparacao",
                "espessuras": [
                    {
                        "espessura": espessura,
                        "estado": "Preparacao",
                        "tempo_min": 60,
                        "lote_baixa": lote,
                        "pecas": [
                            {
                                "id": f"{numero}-P1",
                                "ref_interna": f"{numero}-INT",
                                "ref_externa": f"{numero}-EXT",
                                "descricao": descricao,
                                "material": material,
                                "espessura": espessura,
                                "estado": "Preparacao",
                                "quantidade_pedida": qtd,
                                "ops": [{"nome": "Corte Laser", "estado": "Pendente"}],
                            }
                        ],
                    }
                ],
            }
        ],
    }


def _make_reserva(material_id: str, material: str, espessura: str, quantidade: float, lote: str = "") -> dict:
    row = {
        "material_id": str(material_id or "").strip(),
        "material": str(material or "").strip(),
        "espessura": str(espessura or "").strip(),
        "quantidade": float(quantidade or 0),
    }
    if lote:
        row["lote"] = str(lote or "").strip()
    return row


def main() -> int:
    backend = LegacyBackend()
    data = backend.ensure_data()
    original_encomendas = deepcopy(list(data.get("encomendas", []) or []))
    original_plano = deepcopy(list(data.get("plano", []) or []))
    original_materiais = deepcopy(list(data.get("materiais", []) or []))
    clients = list(data.get("clientes", []) or [])
    if not clients:
        raise RuntimeError("Sem clientes para validar o assistente de materia-prima.")

    client_code = str(dict(clients[0]).get("codigo", "") or "").strip()
    if not client_code:
        raise RuntimeError("Sem codigo de cliente para validar o assistente de materia-prima.")

    numero = "ASSIST-VERIFY-0001"
    numero_a = "ASSIST-VERIFY-A"
    numero_b = "ASSIST-VERIFY-B"
    numero_group = "ASSIST-VERIFY-GRP"
    material_name = "ASSIST-RETALHO"
    material_swap = "ASSIST-SWAP"
    espessura = "5"
    purge_orders = {numero, numero_a, numero_b, numero_group}
    purge_material_ids = {"ASSIST-RET-0001", "ASSIST-LOT-Z", "ASSIST-LOT-A", "ASSIST-LOT-B"}

    try:
        data["encomendas"] = [
            row for row in list(data.get("encomendas", []) or []) if str(row.get("numero", "") or "").strip() not in purge_orders
        ]
        data["plano"] = [
            row for row in list(data.get("plano", []) or []) if str(row.get("encomenda", "") or "").strip() not in purge_orders
        ]
        data["materiais"] = [
            row for row in list(data.get("materiais", []) or []) if str(row.get("id", "") or "").strip() not in purge_material_ids
        ]

        data.setdefault("materiais", []).extend(
            [
                {
                    "id": "ASSIST-LOT-Z",
                    "material": material_name,
                    "espessura": espessura,
                    "quantidade": 6.0,
                    "reservado": 0.0,
                    "comprimento": 3000,
                    "largura": 1500,
                    "Localizacao": "ARMAZEM-A",
                    "lote_fornecedor": "LOTE-Z",
                    "is_sobra": False,
                },
                {
                    "id": "ASSIST-RET-0001",
                    "material": material_name,
                    "espessura": espessura,
                    "quantidade": 1.0,
                    "reservado": 0.0,
                    "comprimento": 1200,
                    "largura": 800,
                    "Localizacao": "RETALHO",
                    "lote_fornecedor": "RET-VERIFY",
                    "is_sobra": True,
                    "origem_lote": "LOTE-BASE",
                    "origem_lotes_baixa": ["LOTE-BASE"],
                },
                {
                    "id": "ASSIST-LOT-A",
                    "material": material_swap,
                    "espessura": espessura,
                    "quantidade": 4.0,
                    "reservado": 0.0,
                    "comprimento": 3000,
                    "largura": 1500,
                    "Localizacao": "ARMAZEM-B",
                    "lote_fornecedor": "LOTE-A",
                    "is_sobra": False,
                },
                {
                    "id": "ASSIST-LOT-B",
                    "material": material_swap,
                    "espessura": espessura,
                    "quantidade": 4.0,
                    "reservado": 0.0,
                    "comprimento": 3000,
                    "largura": 1500,
                    "Localizacao": "ARMAZEM-C",
                    "lote_fornecedor": "LOTE-B",
                    "is_sobra": False,
                },
            ]
        )

        data.setdefault("encomendas", []).extend(
            [
                _make_order(
                    numero,
                    client_code,
                    material_name,
                    espessura,
                    (date.today() + timedelta(days=1)).isoformat(),
                    "LOTE-Z",
                    "Peca de teste para retalho",
                    2,
                    posto_trabalho="Maquina 3030",
                ),
                _make_order(
                    numero_a,
                    client_code,
                    material_swap,
                    espessura,
                    (date.today() + timedelta(days=16)).isoformat(),
                    "LOTE-A",
                    "Peca A para troca de urgencia",
                    4,
                    posto_trabalho="Maquina 5030",
                ),
                _make_order(
                    numero_b,
                    client_code,
                    material_swap,
                    espessura,
                    (date.today() + timedelta(days=15)).isoformat(),
                    "LOTE-B",
                    "Peca B para troca de urgencia",
                    4,
                    posto_trabalho="Maquina 5030",
                ),
                _make_order(
                    numero_group,
                    client_code,
                    material_name,
                    espessura,
                    (date.today() + timedelta(days=2)).isoformat(),
                    "LOTE-Z",
                    "Peca para validar agrupamento por formato",
                    2,
                    posto_trabalho="Maquina 3030",
                ),
            ]
        )
        grouped_order = next(
            row for row in list(data.get("encomendas", []) or []) if str(row.get("numero", "") or "").strip() == numero_group
        )
        grouped_order["cativar"] = True
        grouped_order["reservas"] = [
            _make_reserva("ASSIST-LOT-Z", material_name, espessura, 1, "LOTE-Z"),
            _make_reserva("ASSIST-RET-0001", material_name, espessura, 1, "RET-VERIFY"),
        ]

        tomorrow = (date.today() + timedelta(days=1)).isoformat()
        data.setdefault("plano", []).extend(
            [
                backend._planning_make_block(numero_b, material_swap, espessura, "Corte Laser", tomorrow, 8 * 60, 60),
                backend._planning_make_block(numero_a, material_swap, espessura, "Corte Laser", tomorrow, 10 * 60, 60),
            ]
        )

        snapshot = backend.material_assistant_snapshot(horizon_days=5)
        if not isinstance(snapshot, dict) or "suggestions" not in snapshot or "needs" not in snapshot:
            raise RuntimeError(f"Snapshot invalido do assistente MP: {snapshot}")

        needs = [row for row in list(snapshot.get("needs", []) or []) if str(row.get("numero", "") or "").strip() == numero]
        if not needs:
            raise RuntimeError(f"O assistente MP nao listou a encomenda sintetica: {snapshot}")

        if any(str(row.get("kind", "") or "").strip() == "retalho" for row in list(snapshot.get("suggestions", []) or [])):
            raise RuntimeError(f"O assistente MP nao devia promover retalho como sugestao principal: {snapshot}")

        base_need = needs[0]
        if int(base_need.get("retalho_count", 0) or 0) <= 0:
            raise RuntimeError(f"O assistente MP devia manter o alerta de retalho apenas no contexto da separacao: {base_need}")

        keep_ready_suggestion = next(
            (
                row
                for row in list(snapshot.get("suggestions", []) or [])
                if str(row.get("kind", "") or "").strip() == "keep_ready"
            ),
            None,
        )
        if keep_ready_suggestion is None:
            raise RuntimeError(f"O assistente MP devia sugerir 'nao arrumar' para material cativado no horizonte curto: {snapshot}")

        suggestion_id = str(keep_ready_suggestion.get("id", "") or "").strip()
        feedback = backend.material_assistant_set_feedback(suggestion_id, "accepted")
        if str(dict(feedback.get(suggestion_id, {}) or {}).get("decision", "") or "").strip() != "accepted":
            raise RuntimeError(f"Feedback de validacao nao foi guardado: {feedback}")

        refreshed = backend.material_assistant_snapshot(horizon_days=5)
        refreshed_row = next(
            row for row in list(refreshed.get("suggestions", []) or []) if str(row.get("id", "") or "").strip() == suggestion_id
        )
        if str(refreshed_row.get("status_label", "") or "").strip() != "Validada":
            raise RuntimeError(f"Estado visual da sugestao nao foi atualizado: {refreshed_row}")

        planned_need = next(
            row for row in list(refreshed.get("needs", []) or []) if str(row.get("numero", "") or "").strip() == numero_b
        )
        if str(planned_need.get("plan_origin", "") or "").strip() != "Planeamento":
            raise RuntimeError(f"A necessidade devia estar ancorada ao planeamento e nao ao prazo da encomenda: {planned_need}")
        if tomorrow not in str(planned_need.get("next_action_label", "") or ""):
            raise RuntimeError(f"A proxima acao devia refletir a data planeada e nao a entrega: {planned_need}")

        priority_swap = next(
            (
                row
                for row in list(refreshed.get("suggestions", []) or [])
                if str(row.get("numero", "") or "").strip() == numero_b and str(row.get("kind", "") or "").strip() == "fito_lot"
            ),
            None,
        )
        if priority_swap is None:
            raise RuntimeError(f"O assistente MP nao detetou a troca de lote por urgencia: {refreshed}")
        if "LOTE-A" not in str(priority_swap.get("recommendation", "") or ""):
            raise RuntimeError(f"A recomendacao de urgencia nao apontou para o lote esperado: {priority_swap}")
        if any(str(row.get("kind", "") or "").strip() == "separate_lot" for row in list(refreshed.get("suggestions", []) or [])):
            raise RuntimeError(f"A sugestao 'separate_lot' nao devia aparecer na lista inteligente: {refreshed}")

        separation_rows = list(backend.material_assistant_separation_rows(horizon_days=5) or [])
        separation_row = next((row for row in separation_rows if str(row.get("numero", "") or "").strip() == numero), None)
        if separation_row is None:
            raise RuntimeError(f"A lista de separacao MP nao incluiu a encomenda sintetica: {separation_rows}")
        if (
            not str(separation_row.get("dimensao", "") or "").strip()
            or not str(separation_row.get("visto_sep", "") or "").strip()
            or not str(separation_row.get("planeamento_turno", "") or "").strip()
            or not str(separation_row.get("group_label", "") or "").strip()
            or not str(separation_row.get("posto_trabalho", "") or "").strip()
        ):
            raise RuntimeError(f"A linha de separacao nao ficou completa: {separation_row}")
        grouped_rows = [row for row in separation_rows if str(row.get("numero", "") or "").strip() == numero_group]
        if len(grouped_rows) < 2:
            raise RuntimeError(f"A separacao nao abriu varias linhas para validar o agrupamento por formato: {grouped_rows}")
        grouped_keys = {str(row.get("group_key", "") or "").strip() for row in grouped_rows}
        if len(grouped_keys) < 2:
            raise RuntimeError(f"O agrupamento por formato nao segregou a mesma espessura em grupos distintos: {grouped_rows}")
        if not any("Formato" in str(row.get("group_label", "") or "") for row in grouped_rows):
            raise RuntimeError(f"O label do agrupamento nao evidenciou o formato: {grouped_rows}")
        backend.material_assistant_set_check(str(separation_row.get("check_key", "") or separation_row.get("need_key", "") or ""), "sep", True)
        checked_row = next(
            row
            for row in list(backend.material_assistant_separation_rows(horizon_days=5) or [])
            if str(row.get("numero", "") or "").strip() == numero
        )
        if not bool(checked_row.get("visto_sep_checked")):
            raise RuntimeError(f"O visto de separacao nao ficou gravado: {checked_row}")

        pdf_path = backend.material_assistant_render_separation_pdf(horizon_days=5)
        if not pdf_path.exists() or pdf_path.stat().st_size <= 1200:
            raise RuntimeError(f"PDF de separacao MP invalido: {pdf_path}")

        backend.material_assistant_set_feedback(suggestion_id, "clear")
        backend.material_assistant_set_check(str(separation_row.get("check_key", "") or separation_row.get("need_key", "") or ""), "sep", False)
        print("material-assistant-ok")
        return 0
    finally:
        data["encomendas"] = original_encomendas
        data["plano"] = original_plano
        data["materiais"] = original_materiais


if __name__ == "__main__":
    raise SystemExit(main())
