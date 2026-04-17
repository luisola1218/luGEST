from __future__ import annotations

from datetime import date, datetime, timedelta
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_qt.services.main_bridge import LegacyBackend


DEMO_NUMBERS = {"BARCELBAL1001", "BARCELBAL1002", "BARCELBAL1003"}
CLIENT_SPECS = [
    ("Atlas Planeamento", "500910001", "Braga"),
    ("Boreal Planeamento", "500910002", "Porto"),
    ("Celta Planeamento", "500910003", "Guimaraes"),
]


def _ensure_client(data: dict, backend: LegacyBackend, *, name: str, nif: str, city: str) -> str:
    for row in list(data.get("clientes", []) or []):
        if str(row.get("nome", "") or "").strip().lower() == name.strip().lower():
            return str(row.get("codigo", "") or "").strip()
    code = str(backend.desktop_main.next_cliente_codigo(data) or "").strip()
    data.setdefault("clientes", []).append(
        {
            "codigo": code,
            "nome": name,
            "nif": nif,
            "morada": city,
            "contacto": "253000000",
            "email": f"{name.lower().replace(' ','.')}@demo.local",
            "observacoes": "Cliente demo planeamento por operacao",
        }
    )
    return code


def _op_flow(desktop_main, ops: list[str], *, concluded: bool = False) -> list[dict]:
    flow = desktop_main.build_operacoes_fluxo(ops)
    if concluded:
        for row in flow:
            row["estado"] = "Concluida"
    return flow


def _piece(desktop_main, numero: str, idx: int, *, material: str, esp: str, qty: int, descricao: str, ops: list[str], concluded: bool = False) -> dict:
    piece = {
        "id": f"{numero}-P{idx:02d}",
        "ref_interna": f"{numero}-INT-{idx:02d}",
        "ref_externa": f"{numero}-EXT-{idx:02d}",
        "descricao": descricao,
        "material": material,
        "espessura": esp,
        "estado": "Concluida" if concluded else "Preparacao",
        "quantidade_pedida": int(qty or 1),
        "Operacoes": " + ".join([str(op or "").strip() for op in list(ops or []) if str(op or "").strip()]),
        "operacoes_fluxo": _op_flow(desktop_main, ops, concluded=concluded),
    }
    desktop_main.ensure_peca_operacoes(piece)
    if concluded:
        piece["produzido_ok"] = float(qty or 1)
    return piece


def _esp_bucket(*, esp: str, pieces: list[dict], tempos: dict[str, int | float | str]) -> dict:
    tempos_clean = {}
    for op_name, raw in dict(tempos or {}).items():
        value_txt = str(raw if raw is not None else "").strip()
        if value_txt:
            tempos_clean[str(op_name)] = value_txt
    all_done = all("concl" in str(piece.get("estado", "") or "").strip().lower() for piece in list(pieces or []))
    return {
        "espessura": str(esp or "").strip(),
        "tempo_min": str(tempos_clean.get("Corte Laser", "") or ""),
        "tempos_operacao": tempos_clean,
        "estado": "Concluida" if all_done else "Preparacao",
        "pecas": list(pieces or []),
    }


def _order(
    *,
    numero: str,
    cliente: str,
    nota: str,
    entrega: str,
    posto_trabalho: str,
    materiais: list[dict],
    montagem_itens: list[dict] | None = None,
) -> dict:
    tempo_estimado = 0.0
    for mat in list(materiais or []):
        for esp in list(mat.get("espessuras", []) or []):
            tempos = dict(esp.get("tempos_operacao", {}) or {})
            for raw in tempos.values():
                try:
                    tempo_estimado += float(raw or 0)
                except Exception:
                    continue
    for row in list(montagem_itens or []):
        try:
            tempo_estimado += float(row.get("tempo_total_min", 0) or 0)
        except Exception:
            continue
    return {
        "numero": numero,
        "cliente": cliente,
        "posto_trabalho": posto_trabalho,
        "estado": "Preparacao",
        "data_entrega": entrega,
        "nota_cliente": nota,
        "Observacoes": "Demo planeamento por operacao",
        "tempo_estimado": round(tempo_estimado, 2),
        "tempo": round(tempo_estimado, 2),
        "cativar": False,
        "reservas": [],
        "numero_orcamento": "",
        "materiais": list(materiais or []),
        "montagem_itens": list(montagem_itens or []),
    }


def _material(nome: str, espessuras: list[dict]) -> dict:
    return {
        "material": nome,
        "estado": "Preparacao",
        "espessuras": list(espessuras or []),
    }


def _montagem_item(idx: int, *, codigo: str, descricao: str, qty: float, tempo_total_min: float, conjunto: str) -> dict:
    return {
        "id": f"MONT-{idx:02d}",
        "linha_ordem": idx,
        "tipo_item": "produto_stock",
        "descricao": descricao,
        "produto_codigo": codigo,
        "produto_unid": "UN",
        "qtd_planeada": float(qty or 0),
        "qtd_consumida": 0.0,
        "tempo_total_min": float(tempo_total_min or 0),
        "preco_unit": 0.0,
        "conjunto_codigo": conjunto,
        "conjunto_nome": conjunto,
        "grupo_uuid": f"{conjunto}-{idx:02d}",
        "estado": "Pendente",
        "obs": "Demo planeamento montagem",
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "consumed_at": "",
        "consumed_by": "",
    }


def _week_start_today() -> date:
    today = date.today()
    return today - timedelta(days=today.weekday())


def _select_rows_for_numbers(rows: list[dict], numbers: set[str]) -> list[dict]:
    selected = []
    for row in list(rows or []):
        numero = str(row.get("numero", "") or "").strip()
        if numero in numbers:
            selected.append(dict(row))
    selected.sort(
        key=lambda row: (
            str(row.get("numero", "") or "").strip(),
            str(row.get("material", "") or "").strip(),
            float(LegacyBackend._parse_float(None, row.get("espessura", 0), 0) if False else 0),
        )
    )
    return selected


def main() -> int:
    backend = LegacyBackend()
    desktop_main = backend.desktop_main
    data = backend.ensure_data()

    data["encomendas"] = [
        row for row in list(data.get("encomendas", []) or [])
        if str(row.get("numero", "") or "").strip() not in DEMO_NUMBERS
    ]
    data["plano"] = [
        row for row in list(data.get("plano", []) or [])
        if str(row.get("encomenda", "") or "").strip() not in DEMO_NUMBERS
    ]
    data["plano_hist"] = [
        row for row in list(data.get("plano_hist", []) or [])
        if str(row.get("encomenda", "") or "").strip() not in DEMO_NUMBERS
    ]

    atlas_code = _ensure_client(data, backend, name=CLIENT_SPECS[0][0], nif=CLIENT_SPECS[0][1], city=CLIENT_SPECS[0][2])
    boreal_code = _ensure_client(data, backend, name=CLIENT_SPECS[1][0], nif=CLIENT_SPECS[1][1], city=CLIENT_SPECS[1][2])
    celta_code = _ensure_client(data, backend, name=CLIENT_SPECS[2][0], nif=CLIENT_SPECS[2][1], city=CLIENT_SPECS[2][2])

    entrega_base = _week_start_today() + timedelta(days=3)

    order_1001 = _order(
        numero="BARCELBAL1001",
        cliente=atlas_code,
        nota="Demo Atlas | Laser + Quinagem + Roscagem",
        entrega=(entrega_base + timedelta(days=1)).isoformat(),
        posto_trabalho="Maquina 3030",
        materiais=[
            _material(
                "S235JR",
                [
                    _esp_bucket(
                        esp="2",
                        pieces=[
                            _piece(
                                desktop_main,
                                "BARCELBAL1001",
                                1,
                                material="S235JR",
                                esp="2",
                                qty=4,
                                descricao="Chapa suporte 2 mm",
                                ops=["Corte Laser", "Quinagem"],
                            )
                        ],
                        tempos={"Corte Laser": 55, "Quinagem": 30},
                    ),
                    _esp_bucket(
                        esp="4",
                        pieces=[
                            _piece(
                                desktop_main,
                                "BARCELBAL1001",
                                2,
                                material="S235JR",
                                esp="4",
                                qty=2,
                                descricao="Base roscada 4 mm",
                                ops=["Corte Laser", "Quinagem", "Roscagem"],
                            )
                        ],
                        tempos={"Corte Laser": 40, "Quinagem": 25, "Roscagem": 15},
                    ),
                ],
            )
        ],
    )

    order_1002 = _order(
        numero="BARCELBAL1002",
        cliente=boreal_code,
        nota="Demo Boreal | Laser + Serralharia + Maquinacao + Lacagem",
        entrega=(entrega_base + timedelta(days=2)).isoformat(),
        posto_trabalho="Maquina 5030",
        materiais=[
            _material(
                "S355JR",
                [
                    _esp_bucket(
                        esp="6",
                        pieces=[
                            _piece(
                                desktop_main,
                                "BARCELBAL1002",
                                1,
                                material="S355JR",
                                esp="6",
                                qty=3,
                                descricao="Estrutura soldada 6 mm",
                                ops=["Corte Laser", "Serralharia", "Lacagem"],
                            )
                        ],
                        tempos={"Corte Laser": 60, "Serralharia": 45, "Lacagem": 35},
                    )
                ],
            ),
            _material(
                "C45E",
                [
                    _esp_bucket(
                        esp="20",
                        pieces=[
                            _piece(
                                desktop_main,
                                "BARCELBAL1002",
                                2,
                                material="C45E",
                                esp="20",
                                qty=1,
                                descricao="Bucha maquinada",
                                ops=["Maquinacao"],
                            )
                        ],
                        tempos={"Maquinacao": 75},
                    )
                ],
            ),
        ],
    )

    order_1003 = _order(
        numero="BARCELBAL1003",
        cliente=celta_code,
        nota="Demo Celta | Montagem final",
        entrega=(entrega_base + timedelta(days=3)).isoformat(),
        posto_trabalho="Maquina 5040",
        materiais=[
            _material(
                "S235JR",
                [
                    _esp_bucket(
                        esp="3",
                        pieces=[
                            _piece(
                                desktop_main,
                                "BARCELBAL1003",
                                1,
                                material="S235JR",
                                esp="3",
                                qty=1,
                                descricao="Subconjunto pronto para montagem",
                                ops=["Corte Laser", "Quinagem"],
                                concluded=True,
                            )
                        ],
                        tempos={"Corte Laser": 35, "Quinagem": 20},
                    )
                ],
            )
        ],
        montagem_itens=[
            _montagem_item(
                1,
                codigo="KIT-MONT-01",
                descricao="Kit montagem final",
                qty=1,
                tempo_total_min=70,
                conjunto="Conjunto final demo",
            )
        ],
    )

    demo_orders = [order_1001, order_1002, order_1003]
    for enc in demo_orders:
        desktop_main.update_estado_encomenda_por_espessuras(enc)
        data.setdefault("encomendas", []).append(enc)

    backend._save(force=True)
    backend.reload(force=True)

    week_start = _week_start_today()
    selected_numbers = {str(enc.get("numero", "") or "").strip() for enc in demo_orders}
    operation_sequence = [
        "Corte Laser",
        "Quinagem",
        "Serralharia",
        "Maquinacao",
        "Roscagem",
        "Lacagem",
        "Montagem",
    ]
    planned_counts: dict[str, int] = {}
    for operation in operation_sequence:
        rows = [dict(row) for row in list(backend.planning_pending_rows(operation=operation) or []) if str(row.get("numero", "") or "").strip() in selected_numbers]
        if not rows:
            planned_counts[operation] = 0
            continue
        placed = list(backend.planning_auto_plan(rows, week_start=week_start, operation=operation) or [])
        planned_counts[operation] = len(placed)

    backend.reload(force=True)

    print("planning-operations-demo-ok")
    print("orders:", ", ".join(sorted(selected_numbers)))
    for operation in operation_sequence:
        pending_rows = [row for row in list(backend.planning_pending_rows(operation=operation) or []) if str(row.get("numero", "") or "").strip() in selected_numbers]
        active_blocks = [
            row
            for row in list(backend.ensure_data().get("plano", []) or [])
            if str(row.get("encomenda", "") or "").strip() in selected_numbers
            and str(backend._planning_row_operation(row) or "").strip() == operation
        ]
        print(f"{operation}: backlog={len(pending_rows)} blocos={len(active_blocks)} colocados={planned_counts.get(operation, 0)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
