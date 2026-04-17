from __future__ import annotations

from datetime import date, timedelta
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_qt.services.main_bridge import LegacyBackend


PREFIX = "BARCELBAL"
VERIFY_PREFIX = "ASSIST-VERIFY-"
MAT_PREFIX = "AMD-"
POSTOS = ["Maquina 3030", "Maquina 5030", "Maquina 5040"]
DEMO_NUMBERS = {
    "BARCELBAL0001",
    "BARCELBAL0002",
    "BARCELBAL0003",
    "BARCELBAL0004",
    "BARCELBAL0005",
    "BARCELBAL0006",
    "980301",
    "980302",
    "980303",
    "980304",
    "980305",
    "980306",
}


def _mat_id(code: str, suffix: str) -> str:
    return f"{MAT_PREFIX}{code}-{suffix}"


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
            "email": f"{name.lower()}@demo.local",
            "observacoes": "Cliente demo Assistente MP",
        }
    )
    return code


def _make_material(
    material_id: str,
    material: str,
    espessura: str,
    lote: str,
    *,
    quantidade: float,
    comprimento: int,
    largura: int,
    local: str,
    is_sobra: bool = False,
    origem_lote: str = "",
    origem_encomenda: str = "",
) -> dict:
    row = {
        "id": material_id,
        "material": material,
        "espessura": espessura,
        "quantidade": float(quantidade),
        "reservado": 0.0,
        "comprimento": int(comprimento),
        "largura": int(largura),
        "Localizacao": local,
        "lote_fornecedor": lote,
        "is_sobra": bool(is_sobra),
    }
    if origem_lote:
        row["origem_lote"] = origem_lote
    if origem_encomenda:
        row["origem_encomenda"] = origem_encomenda
    return row


def _build_piece(numero: str, index: int, item: dict) -> dict:
    return {
        "id": f"{numero}-P{index:02d}",
        "ref_interna": f"{numero}-INT-{index:02d}",
        "ref_externa": f"{numero}-EXT-{index:02d}",
        "descricao": str(item.get("descricao", "") or "").strip(),
        "material": str(item.get("material", "") or "").strip(),
        "espessura": str(item.get("esp", "") or "").strip(),
        "estado": "Preparacao",
        "quantidade_pedida": int(item.get("qtd", 1) or 1),
        "ops": [{"nome": "Corte Laser", "estado": "Pendente"}],
    }


def _make_order(
    numero: str,
    cliente: str,
    entrega: str,
    descricao: str,
    items: list[dict],
    reservas: list[dict] | None = None,
    *,
    posto_trabalho: str = "",
) -> dict:
    materiais_map: dict[str, dict] = {}
    total_min = 0.0
    for index, item in enumerate(list(items or []), start=1):
        material = str(item.get("material", "") or "").strip()
        esp = str(item.get("esp", "") or "").strip()
        total_min += float(item.get("tempo_min", 0) or 0)
        material_bucket = materiais_map.setdefault(
            material,
            {
                "material": material,
                "estado": "Preparacao",
                "espessuras": [],
            },
        )
        material_bucket["espessuras"].append(
            {
                "espessura": esp,
                "estado": "Preparacao",
                "tempo_min": float(item.get("tempo_min", 0) or 0),
                "lote_baixa": str(item.get("lote", "") or "").strip(),
                "pecas": [_build_piece(numero, index, item)],
            }
        )
    return {
        "numero": numero,
        "cliente": cliente,
        "posto_trabalho": str(posto_trabalho or "").strip(),
        "estado": "Preparacao",
        "data_entrega": entrega,
        "nota_cliente": descricao,
        "tempo_estimado": total_min,
        "Observacoes": "",
        "cativar": bool(list(reservas or [])),
        "reservas": [dict(row or {}) for row in list(reservas or [])],
        "numero_orcamento": "",
        "montagem_itens": [],
        "materiais": list(materiais_map.values()),
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


def _is_demo_order(numero: str) -> bool:
    numero_txt = str(numero or "").strip()
    return (
        numero_txt.startswith(PREFIX)
        or numero_txt.startswith(VERIFY_PREFIX)
        or numero_txt in DEMO_NUMBERS
    )


def _clear_demo_config(backend: LegacyBackend) -> None:
    cfg = backend._load_qt_config()
    feedback = dict(cfg.get("material_assistant_feedback", {}) or {})
    checks = dict(cfg.get("material_assistant_checks", {}) or {})
    cfg["material_assistant_feedback"] = {
        key: value
        for key, value in feedback.items()
        if PREFIX not in str(key or "") and VERIFY_PREFIX not in str(key or "")
    }
    cfg["material_assistant_checks"] = {
        key: value
        for key, value in checks.items()
        if PREFIX not in str(key or "") and VERIFY_PREFIX not in str(key or "")
    }
    backend._save_qt_config(cfg)


def _pick_demo_day(data: dict, start_from: date) -> date:
    for offset in range(1, 15):
        day = start_from + timedelta(days=offset)
        if day.weekday() > 5:
            continue
        day_txt = day.isoformat()
        if not any(str(row.get("data", "") or "").strip() == day_txt for row in list(data.get("plano", []) or [])):
            return day
    return start_from + timedelta(days=1)


def _sequential_plan(backend: LegacyBackend, data: dict, flat_specs: list[dict], base_day: date) -> None:
    dates = [base_day + timedelta(days=i) for i in range(6)]
    start_min, end_min, slot = backend._planning_grid_metrics()
    cur_day_idx = 0
    cur_min = start_min

    for spec in flat_specs:
        duration = backend._planning_round_duration(spec.get("tempo_min", 0))
        if duration <= 0:
            raise RuntimeError(f"Tempo invalido para {spec.get('numero', '-')} / {spec.get('material', '-')} {spec.get('esp', '-')}")
        remaining = duration
        first_day = ""
        first_start = None
        while remaining > 0:
            next_day_idx, day_txt, segment_start, segment_end = backend._planning_next_free_segment(dates, cur_day_idx, cur_min)
            if not day_txt or segment_start is None or segment_end is None:
                raise RuntimeError(f"Sem espaco livre para planear {spec.get('numero', '-')}")
            free_minutes = max(0, int(segment_end - segment_start))
            chunk = min(remaining, free_minutes)
            if chunk % slot != 0:
                chunk = int((chunk + slot - 1) // slot) * slot
            if chunk > free_minutes:
                chunk = free_minutes
            if chunk <= 0:
                raise RuntimeError(f"Sem bloco utilizavel para {spec.get('numero', '-')}")
            data.setdefault("plano", []).append(
                backend._planning_make_block(
                    str(spec.get("numero", "") or "").strip(),
                    str(spec.get("material", "") or "").strip(),
                    str(spec.get("esp", "") or "").strip(),
                    day_txt,
                    segment_start,
                    chunk,
                    posto=str(spec.get("posto", "") or "").strip(),
                )
            )
            if first_start is None:
                first_day = day_txt
                first_start = int(segment_start)
            remaining -= chunk
            cur_day_idx = next_day_idx
            cur_min = int(segment_start + chunk)
            if cur_min >= end_min:
                cur_day_idx += 1
                cur_min = start_min
        spec["plan_day"] = first_day
        spec["start_min"] = int(first_start if first_start is not None else start_min)
        spec["duracao_planeada"] = duration


def main() -> int:
    backend = LegacyBackend()
    data = backend.ensure_data()

    data["encomendas"] = [
        row
        for row in list(data.get("encomendas", []) or [])
        if not _is_demo_order(str(row.get("numero", "") or "").strip())
    ]
    data["plano"] = [
        row
        for row in list(data.get("plano", []) or [])
        if not _is_demo_order(str(row.get("encomenda", "") or "").strip())
    ]
    data["materiais"] = [
        row
        for row in list(data.get("materiais", []) or [])
        if not str(row.get("id", "") or "").strip().startswith(MAT_PREFIX)
        and not str(row.get("id", "") or "").strip().startswith(PREFIX)
    ]
    data["postos_trabalho"] = list(POSTOS)
    _clear_demo_config(backend)

    atlas = _ensure_client(data, backend, name="Atlas", nif="500100001", city="Rua do Atlas, Braga")
    boreal = _ensure_client(data, backend, name="Boreal", nif="500100002", city="Zona Industrial, Viana")
    celta = _ensure_client(data, backend, name="Celta", nif="500100003", city="Parque Empresarial, Porto")
    delta = _ensure_client(data, backend, name="Delta", nif="500100004", city="Avenida Delta, Guimaraes")

    data.setdefault("materiais", []).extend(
        [
            _make_material(_mat_id("ATL02", "A"), "S235JR", "2", "ATL-02-A", quantidade=8, comprimento=3000, largura=1500, local="RACK-ATL-02A"),
            _make_material(_mat_id("ATL02", "B"), "S235JR", "2", "ATL-02-B", quantidade=6, comprimento=2500, largura=1250, local="RACK-ATL-02B"),
            _make_material(_mat_id("ATL02", "C"), "S235JR", "2", "ATL-02-C", quantidade=4, comprimento=2000, largura=1000, local="RACK-ATL-02C"),
            _make_material(_mat_id("ATL02", "R1"), "S235JR", "2", "RET-ATL-02A", quantidade=1, comprimento=1600, largura=900, local="RET-ATL-02A", is_sobra=True, origem_lote="ATL-02-A", origem_encomenda="BARCELBAL0001"),
            _make_material(_mat_id("ATL02", "R2"), "S235JR", "2", "RET-ATL-02B", quantidade=1, comprimento=1400, largura=750, local="RET-ATL-02B", is_sobra=True, origem_lote="ATL-02-A", origem_encomenda="BARCELBAL0001"),
            _make_material(_mat_id("ATL02", "R3"), "S235JR", "2", "RET-ATL-02C", quantidade=1, comprimento=1000, largura=600, local="RET-ATL-02C", is_sobra=True, origem_lote="ATL-02-A", origem_encomenda="BARCELBAL0001"),
            _make_material(_mat_id("ATL02", "R4"), "S235JR", "2", "RET-ATL-02D", quantidade=1, comprimento=800, largura=500, local="RET-ATL-02D", is_sobra=True, origem_lote="ATL-02-A", origem_encomenda="BARCELBAL0001"),
            _make_material(_mat_id("ATL03", "A"), "S235JR", "3", "ATL-03-A", quantidade=5, comprimento=3000, largura=1500, local="RACK-ATL-03"),
            _make_material(_mat_id("ATL03", "B"), "S235JR", "3", "ATL-03-B", quantidade=4, comprimento=2500, largura=1250, local="RACK-ATL-03B"),
            _make_material(_mat_id("ATL05", "A"), "S235JR", "5", "ATL-05-A", quantidade=6, comprimento=3000, largura=1500, local="RACK-ATL-05A"),
            _make_material(_mat_id("ATL05", "B"), "S235JR", "5", "ATL-05-B", quantidade=5, comprimento=2500, largura=1250, local="RACK-ATL-05B"),
            _make_material(_mat_id("ATL06", "A"), "S235JR", "6", "ATL-06-A", quantidade=4, comprimento=3000, largura=1500, local="RACK-ATL-06"),
            _make_material(_mat_id("ATL08", "A"), "S235JR", "8", "ATL-08-A", quantidade=4, comprimento=3000, largura=1500, local="RACK-ATL-08"),
            _make_material(_mat_id("S35504", "A"), "S355MC", "4", "S355-04-A", quantidade=5, comprimento=3000, largura=1500, local="RACK-S355-04A"),
            _make_material(_mat_id("S35504", "B"), "S355MC", "4", "S355-04-B", quantidade=4, comprimento=2500, largura=1250, local="RACK-S355-04B"),
            _make_material(_mat_id("BOR04", "A"), "S275JR", "4", "BOR-04-A", quantidade=5, comprimento=3000, largura=1500, local="RACK-BOR-04A"),
            _make_material(_mat_id("BOR04", "B"), "S275JR", "4", "BOR-04-B", quantidade=3, comprimento=2500, largura=1250, local="RACK-BOR-04B"),
            _make_material(_mat_id("BOR06", "A"), "S275JR", "6", "BOR-06-A", quantidade=5, comprimento=3000, largura=1500, local="RACK-BOR-06"),
            _make_material(_mat_id("BOR06", "B"), "S275JR", "6", "BOR-06-B", quantidade=3, comprimento=2500, largura=1250, local="RACK-BOR-06B"),
            _make_material(_mat_id("BOR10", "A"), "S275JR", "10", "BOR-10-A", quantidade=4, comprimento=3000, largura=1500, local="RACK-BOR-10"),
            _make_material(_mat_id("CEL15", "A"), "AISI 304L", "1.5", "CEL-15-A", quantidade=4, comprimento=3000, largura=1500, local="RACK-CEL-15"),
            _make_material(_mat_id("CEL20", "A"), "AISI 304L", "2", "CEL-20-A", quantidade=4, comprimento=3000, largura=1500, local="RACK-CEL-20"),
            _make_material(_mat_id("CEL20", "B"), "AISI 304L", "2", "CEL-20-B", quantidade=3, comprimento=2500, largura=1250, local="RACK-CEL-20B"),
        ]
    )

    base_day = _pick_demo_day(data, date.today())
    order_specs = [
        {
            "numero": "BARCELBAL0001",
            "cliente": atlas,
            "posto_trabalho": "Maquina 3030",
            "entrega": (base_day + timedelta(days=1)).isoformat(),
            "descricao": "Atlas | separacao multipla 2 mm",
            "items": [
                {"material": "S235JR", "esp": "2", "lote": "ATL-02-A", "descricao": "Atlas | S235JR 2 mm", "qtd": 4, "tempo_min": 60},
                {"material": "S235JR", "esp": "3", "lote": "ATL-03-A", "descricao": "Atlas | S235JR 3 mm", "qtd": 3, "tempo_min": 45},
                {"material": "S235JR", "esp": "5", "lote": "ATL-05-A", "descricao": "Atlas | S235JR 5 mm", "qtd": 2, "tempo_min": 60},
                {"material": "AISI 304L", "esp": "2", "lote": "CEL-20-A", "descricao": "Atlas | inox 2 mm", "qtd": 2, "tempo_min": 45},
            ],
            "reservas": [
                _make_reserva(_mat_id("ATL02", "A"), "S235JR", "2", 2, "ATL-02-A"),
                _make_reserva(_mat_id("ATL02", "B"), "S235JR", "2", 1, "ATL-02-B"),
                _make_reserva(_mat_id("ATL02", "R1"), "S235JR", "2", 1, "RET-ATL-02A"),
                _make_reserva(_mat_id("ATL02", "R2"), "S235JR", "2", 1, "RET-ATL-02B"),
                _make_reserva(_mat_id("CEL20", "A"), "AISI 304L", "2", 1, "CEL-20-A"),
            ],
        },
        {
            "numero": "BARCELBAL0002",
            "cliente": atlas,
            "posto_trabalho": "Maquina 5030",
            "entrega": (base_day + timedelta(days=2)).isoformat(),
            "descricao": "Atlas | formatos alternativos",
            "items": [
                {"material": "S355MC", "esp": "4", "lote": "S355-04-A", "descricao": "Atlas | S355MC 4 mm", "qtd": 3, "tempo_min": 60},
                {"material": "S235JR", "esp": "8", "lote": "ATL-08-A", "descricao": "Atlas | S235JR 8 mm", "qtd": 2, "tempo_min": 60},
                {"material": "S235JR", "esp": "6", "lote": "ATL-06-A", "descricao": "Atlas | S235JR 6 mm", "qtd": 2, "tempo_min": 45},
            ],
            "reservas": [
                _make_reserva(_mat_id("S35504", "B"), "S355MC", "4", 1, "S355-04-B"),
                _make_reserva(_mat_id("ATL08", "A"), "S235JR", "8", 1, "ATL-08-A"),
            ],
        },
        {
            "numero": "BARCELBAL0003",
            "cliente": boreal,
            "posto_trabalho": "Maquina 5040",
            "entrega": (base_day + timedelta(days=2)).isoformat(),
            "descricao": "Boreal | aco carbono",
            "items": [
                {"material": "S275JR", "esp": "4", "lote": "BOR-04-A", "descricao": "Boreal | S275JR 4 mm", "qtd": 3, "tempo_min": 45},
                {"material": "S275JR", "esp": "6", "lote": "BOR-06-A", "descricao": "Boreal | S275JR 6 mm", "qtd": 3, "tempo_min": 60},
                {"material": "S275JR", "esp": "10", "lote": "BOR-10-A", "descricao": "Boreal | S275JR 10 mm", "qtd": 2, "tempo_min": 60},
            ],
            "reservas": [
                _make_reserva(_mat_id("BOR04", "A"), "S275JR", "4", 1, "BOR-04-A"),
                _make_reserva(_mat_id("BOR04", "B"), "S275JR", "4", 1, "BOR-04-B"),
                _make_reserva(_mat_id("BOR06", "A"), "S275JR", "6", 2, "BOR-06-A"),
            ],
        },
        {
            "numero": "BARCELBAL0004",
            "cliente": celta,
            "posto_trabalho": "Maquina 3030",
            "entrega": (base_day + timedelta(days=3)).isoformat(),
            "descricao": "Celta | inox e carbono",
            "items": [
                {"material": "AISI 304L", "esp": "1.5", "lote": "CEL-15-A", "descricao": "Celta | inox 1.5 mm", "qtd": 2, "tempo_min": 45},
                {"material": "AISI 304L", "esp": "2", "lote": "CEL-20-A", "descricao": "Celta | inox 2 mm", "qtd": 2, "tempo_min": 45},
                {"material": "S355MC", "esp": "4", "lote": "S355-04-B", "descricao": "Celta | S355MC 4 mm", "qtd": 2, "tempo_min": 45},
            ],
        },
        {
            "numero": "BARCELBAL0005",
            "cliente": delta,
            "posto_trabalho": "Maquina 3030",
            "entrega": base_day.isoformat(),
            "descricao": "Delta | urgente com retalho",
            "items": [
                {"material": "S235JR", "esp": "2", "lote": "ATL-02-B", "descricao": "Delta | S235JR 2 mm urgente", "qtd": 2, "tempo_min": 30},
            ],
        },
        {
            "numero": "BARCELBAL0006",
            "cliente": boreal,
            "posto_trabalho": "Maquina 5030",
            "entrega": (base_day + timedelta(days=4)).isoformat(),
            "descricao": "Boreal | misto complementar",
            "items": [
                {"material": "S235JR", "esp": "5", "lote": "ATL-05-B", "descricao": "Boreal | S235JR 5 mm", "qtd": 2, "tempo_min": 45},
                {"material": "S355MC", "esp": "4", "lote": "S355-04-A", "descricao": "Boreal | S355MC 4 mm", "qtd": 2, "tempo_min": 45},
            ],
            "reservas": [
                _make_reserva(_mat_id("ATL05", "A"), "S235JR", "5", 1, "ATL-05-A"),
                _make_reserva(_mat_id("ATL05", "B"), "S235JR", "5", 1, "ATL-05-B"),
            ],
        },
    ]

    flat_specs: list[dict] = []
    for order in order_specs:
        data.setdefault("encomendas", []).append(
            _make_order(
                str(order.get("numero", "") or "").strip(),
                str(order.get("cliente", "") or "").strip(),
                str(order.get("entrega", "") or "").strip(),
                str(order.get("descricao", "") or "").strip(),
                list(order.get("items", []) or []),
                list(order.get("reservas", []) or []),
                posto_trabalho=str(order.get("posto_trabalho", "") or "").strip(),
            )
        )
        for item in list(order.get("items", []) or []):
            flat_specs.append(
                {
                    "numero": str(order.get("numero", "") or "").strip(),
                    "cliente": str(order.get("cliente", "") or "").strip(),
                    "material": str(item.get("material", "") or "").strip(),
                    "esp": str(item.get("esp", "") or "").strip(),
                    "entrega": str(order.get("entrega", "") or "").strip(),
                    "lote": str(item.get("lote", "") or "").strip(),
                    "descricao": str(item.get("descricao", "") or "").strip(),
                    "qtd": int(item.get("qtd", 1) or 1),
                    "tempo_min": int(item.get("tempo_min", 0) or 0),
                    "posto": str(order.get("posto_trabalho", "") or "Sem posto").strip() or "Sem posto",
                }
            )

    flat_specs.sort(
        key=lambda row: (
            str(row.get("entrega", "") or "9999-99-99"),
            0 if str(row.get("numero", "") or "").strip() == "BARCELBAL0005" else 1,
            str(row.get("posto", "") or ""),
            str(row.get("numero", "") or ""),
            str(row.get("material", "") or ""),
            str(row.get("esp", "") or ""),
        )
    )

    _sequential_plan(backend, data, flat_specs, base_day)

    for order in list(order_specs or []):
        reservas = [dict(row or {}) for row in list(order.get("reservas", []) or [])]
        if reservas:
            backend.desktop_main.aplicar_reserva_em_stock(data, reservas, 1)

    backend._normalize_storage_paths_for_save()
    payload, _changed = backend._merge_latest_for_save()
    backend.desktop_main._mysql_save_relational_data(payload)

    posto_map = {
        (
            str(spec.get("numero", "") or "").strip(),
            str(spec.get("material", "") or "").strip(),
            str(spec.get("esp", "") or "").strip(),
        ): str(spec.get("posto", "") or "").strip()
        for spec in flat_specs
        if str(spec.get("posto", "") or "").strip()
    }
    try:
        conn = backend.desktop_main._mysql_connect()
        cur = conn.cursor()
        for (numero, material, espessura), posto in posto_map.items():
            cur.execute(
                """
                UPDATE plano
                   SET posto=%s
                 WHERE encomenda_numero=%s
                   AND material=%s
                   AND espessura=%s
                """,
                (posto, numero, material, espessura),
            )
            cur.execute(
                """
                UPDATE plano_hist
                   SET posto=%s
                 WHERE encomenda_numero=%s
                   AND material=%s
                   AND espessura=%s
                """,
                (posto, numero, material, espessura),
            )
        conn.commit()
        cur.close()
        conn.close()
    except Exception:
        pass
    print(
        "material-assistant-demo-ok "
        f"orders={len(order_specs)} "
        f"needs={len(flat_specs)} "
        f"day={base_day.isoformat()} "
        f"minutes={sum(int(spec.get('duracao_planeada', 0) or 0) for spec in flat_specs)}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
