from __future__ import annotations

import json
import random
import sys
from datetime import date, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_qt.services.main_bridge import LegacyBackend


TAG = "OPERATOR_BULK_20260421"
MANIFEST_PATH = ROOT / "generated" / "seeds" / "operator_bulk_manifest.json"
TARGET_CLIENTS = 500
TARGET_QUOTES = 500
TARGET_CONVERTED = 300
TARGET_PLANNED = 30
TARGET_FINISHED = 15
LASER_RESOURCES = ("Maquina 3030", "Maquina 5030")
RNG = random.Random(20260421)

CLIENT_PREFIXES = [
    "Alfa",
    "Atlas",
    "Boreal",
    "Celta",
    "Delta",
    "Eixo",
    "Ferro",
    "Gamma",
    "Helix",
    "Iber",
    "Jota",
    "Kappa",
    "Lince",
    "Metro",
    "Norte",
    "Omega",
    "Ponte",
    "Quadrante",
    "Rumo",
    "Sigma",
    "Tavares",
    "Uni",
    "Vector",
    "West",
    "Zenite",
]

CLIENT_SUFFIXES = [
    "Metal",
    "Industrial",
    "Mecanica",
    "Solucoes",
    "Tecnica",
    "Factory",
    "Estruturas",
    "Precision",
    "Inox",
    "Sistemas",
]

MATERIAL_SPECS = [
    {"material": "INOX304 2B", "espessura": "2", "operations": "Corte Laser + Embalamento"},
    {"material": "S235JR", "espessura": "3", "operations": "Corte Laser + Embalamento"},
    {"material": "S235JR", "espessura": "5", "operations": "Corte Laser + Embalamento"},
    {"material": "S275JR", "espessura": "8", "operations": "Corte Laser + Embalamento"},
    {"material": "S355JR", "espessura": "10", "operations": "Corte Laser + Embalamento"},
    {"material": "AISI 304L GR 220 ESCOVADO", "espessura": "1,5", "operations": "Corte Laser + Embalamento"},
]


def _load_manifest() -> dict[str, object] | None:
    if not MANIFEST_PATH.exists():
        return None
    try:
        return json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
    except Exception:
        return None


def _save_manifest(payload: dict[str, object]) -> None:
    MANIFEST_PATH.parent.mkdir(parents=True, exist_ok=True)
    MANIFEST_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _week_start_today() -> date:
    today = date.today()
    return today - timedelta(days=today.weekday())


def _print_progress(label: str, current: int, total: int) -> None:
    print(f"{label}: {current}/{total}")


def _clean_generated_dataset(backend: LegacyBackend, manifest: dict[str, object] | None = None) -> None:
    manifest = manifest or _load_manifest() or {}
    data = backend.ensure_data()

    client_codes = {
        str(value).strip()
        for value in list(manifest.get("client_codes", []) or [])
        if str(value).strip()
    }
    quote_numbers = {
        str(value).strip()
        for value in list(manifest.get("quote_numbers", []) or [])
        if str(value).strip()
    }
    order_numbers = {
        str(value).strip()
        for value in list(manifest.get("order_numbers", []) or [])
        if str(value).strip()
    }
    ref_externas = {
        str(value).strip()
        for value in list(manifest.get("ref_externas", []) or [])
        if str(value).strip()
    }

    if not client_codes:
        client_codes = {
            str(row.get("codigo", "") or "").strip()
            for row in list(data.get("clientes", []) or [])
            if str(row.get("email", "") or "").strip().endswith("@seed-lugest.local")
            or TAG in str(row.get("obs_tecnicas", "") or "")
            or TAG in str(row.get("observacoes", "") or "")
        }
    if not quote_numbers:
        quote_numbers = {
            str(row.get("numero", "") or "").strip()
            for row in list(data.get("orcamentos", []) or [])
            if TAG in str(row.get("nota_cliente", "") or "")
        }
    if not order_numbers:
        order_numbers = {
            str(row.get("numero", "") or "").strip()
            for row in list(data.get("encomendas", []) or [])
            if TAG in str(row.get("nota_cliente", "") or "")
            or TAG in str(row.get("observacoes", "") or "")
        }
    if not ref_externas:
        ref_externas = {
            str(value).strip()
            for value in list((data.get("orc_refs", {}) or {}).keys())
            if str(value).strip().startswith("LGS-OPR-")
        }

    if order_numbers:
        data["plano"] = [
            row for row in list(data.get("plano", []) or [])
            if str(row.get("encomenda", row.get("encomenda_numero", "")) or "").strip() not in order_numbers
        ]
        data["plano_hist"] = [
            row for row in list(data.get("plano_hist", []) or [])
            if str(row.get("encomenda", row.get("encomenda_numero", "")) or "").strip() not in order_numbers
        ]
        data["expedicoes"] = [
            row for row in list(data.get("expedicoes", []) or [])
            if str(row.get("encomenda_numero", "") or "").strip() not in order_numbers
        ]
    if order_numbers:
        data["encomendas"] = [
            row for row in list(data.get("encomendas", []) or [])
            if str(row.get("numero", "") or "").strip() not in order_numbers
        ]
    if quote_numbers:
        data["orcamentos"] = [
            row for row in list(data.get("orcamentos", []) or [])
            if str(row.get("numero", "") or "").strip() not in quote_numbers
        ]
    if client_codes:
        data["clientes"] = [
            row for row in list(data.get("clientes", []) or [])
            if str(row.get("codigo", "") or "").strip() not in client_codes
        ]

    if ref_externas:
        peca_hist = dict(data.get("peca_hist", {}) or {})
        for ref in ref_externas:
            peca_hist.pop(ref, None)
        data["peca_hist"] = peca_hist

        orc_refs = dict(data.get("orc_refs", {}) or {})
        for ref in ref_externas:
            orc_refs.pop(ref, None)
        data["orc_refs"] = orc_refs

    if quote_numbers:
        data["faturacao_registos"] = [
            row for row in list(data.get("faturacao_registos", []) or [])
            if str(row.get("orcamento_numero", "") or "").strip() not in quote_numbers
        ]
    if order_numbers:
        data["faturacao_registos"] = [
            row for row in list(data.get("faturacao_registos", []) or [])
            if str(row.get("encomenda_numero", "") or "").strip() not in order_numbers
        ]

    backend._save(force=True)
    if MANIFEST_PATH.exists():
        MANIFEST_PATH.unlink()


def _start_with_retry(backend: LegacyBackend, enc_num: str, piece_id: str, operator_name: str, operation: str) -> None:
    try:
        backend.operator_start_piece(enc_num, piece_id, operator_name, operation, "Geral")
    except ValueError as exc:
        message = str(exc)
        if "Operacao ocupada" not in message:
            raise
        reset_fn = getattr(backend.operador_actions, "_mysql_ops_reset_piece", None)
        if callable(reset_fn):
            reset_fn(enc_num, piece_id)
        backend.operator_start_piece(enc_num, piece_id, operator_name, operation, "Geral")


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
            reset_fn = getattr(backend.operador_actions, "_mysql_ops_reset_piece", None)
            if callable(reset_fn):
                reset_fn(enc_num, piece_id)
            _start_with_retry(backend, enc_num, piece_id, operator_name, operation)
    if last_exc is not None:
        raise last_exc


def _client_name(index: int) -> str:
    prefix = CLIENT_PREFIXES[(index - 1) % len(CLIENT_PREFIXES)]
    suffix = CLIENT_SUFFIXES[((index - 1) // len(CLIENT_PREFIXES)) % len(CLIENT_SUFFIXES)]
    return f"{prefix} {suffix} {index:03d}"


def _create_clients(backend: LegacyBackend) -> list[dict[str, str]]:
    data = backend.ensure_data()
    created: list[dict[str, str]] = []
    for index in range(1, TARGET_CLIENTS + 1):
        name = _client_name(index)
        row = {
            "codigo": str(backend.desktop_main.next_cliente_codigo(data)),
            "nome": name,
            "nif": f"50{index:07d}"[:9],
            "morada": f"Zona Industrial {((index - 1) % 40) + 1}, Portugal",
            "contacto": f"91{index:07d}"[:9],
            "email": f"cliente.{index:03d}@seed-lugest.local",
            "prazo_entrega": f"{5 + (index % 10)} dias",
            "cond_pagamento": "30 dias",
            "obs_tecnicas": TAG,
            "observacoes": "Lote de simulacao para operador",
        }
        data.setdefault("clientes", []).append(row)
        created.append({"codigo": str(row.get("codigo", "") or "").strip(), "nome": str(row.get("nome", "") or "").strip()})
        if index % 50 == 0 or index == TARGET_CLIENTS:
            _print_progress("clientes", index, TARGET_CLIENTS)
    return created


def _quote_payload(client: dict[str, str], quote_index: int, *, machine: str | None = None, approved: bool = False) -> tuple[dict[str, object], str]:
    spec = MATERIAL_SPECS[(quote_index - 1) % len(MATERIAL_SPECS)]
    qty = float(2 + (quote_index % 9))
    price_unit = round(18.5 + ((quote_index * 1.37) % 85), 2)
    ext_ref = f"LGS-OPR-{quote_index:04d}-01"
    note = f"Lote operador {TAG} | cliente {client['codigo']} | quote {quote_index:04d}"
    state = "Aprovado" if approved else ("Enviado" if quote_index % 3 else "Em revisão")
    payload: dict[str, object] = {
        "cliente": {"codigo": client["codigo"], "nome": client["nome"]},
        "estado": state,
        "posto_trabalho": machine or "",
        "linhas": [
            {
                "tipo_item": "Peca",
                "ref_externa": ext_ref,
                "descricao": f"Conjunto corte {quote_index:04d}",
                "material": spec["material"],
                "espessura": spec["espessura"],
                "operacao": spec["operations"],
                "qtd": qty,
                "preco_unit": price_unit,
                "tempo_peca_min": round(2.0 + ((quote_index % 5) * 0.7), 2),
            }
        ],
        "iva_perc": 23,
        "preco_transporte": 0.0,
        "nota_cliente": note,
        "executado_por": "admin",
        "nota_transporte": "Simulacao interna",
    }
    return payload, ext_ref


def _apply_operation_times(
    backend: LegacyBackend,
    order_num: str,
    material: str,
    espessura: str,
    machine: str,
    quote_index: int,
) -> None:
    laser_min = 60 + ((quote_index - 1) % 5) * 30
    pack_min = 15 + ((quote_index - 1) % 4) * 15
    backend.order_espessura_set_operation_times(
        order_num,
        material,
        espessura,
        tempos_operacao={
            "Corte Laser": laser_min,
            "Embalamento": pack_min,
        },
        maquinas_operacao={
            "Corte Laser": machine,
            "Embalamento": "Embalamento",
        },
    )


def _direct_quote_row(
    backend: LegacyBackend,
    client: dict[str, str],
    quote_index: int,
    *,
    quote_num: str,
    order_num: str,
    ext_ref: str,
    ref_interna: str,
    machine: str,
    approved: bool,
) -> dict[str, object]:
    spec = MATERIAL_SPECS[(quote_index - 1) % len(MATERIAL_SPECS)]
    qty = float(2 + (quote_index % 9))
    price_unit = round(18.5 + ((quote_index * 1.37) % 85), 2)
    line_total = round(qty * price_unit, 2)
    client_payload = {
        "codigo": client["codigo"],
        "nome": client["nome"],
        "empresa": client["nome"],
        "nif": f"50{quote_index:07d}"[:9],
        "morada": f"Zona Industrial {((quote_index - 1) % 40) + 1}, Portugal",
        "contacto": f"91{quote_index:07d}"[:9],
        "email": f"cliente.{quote_index:03d}@seed-lugest.local",
    }
    line = {
        "tipo_item": backend.desktop_main.ORC_LINE_TYPE_PIECE,
        "ref_interna": ref_interna,
        "ref_externa": ext_ref,
        "descricao": f"Conjunto corte {quote_index:04d}",
        "material": spec["material"],
        "espessura": spec["espessura"],
        "operacao": spec["operations"],
        "qtd": qty,
        "preco_unit": price_unit,
        "tempo_peca_min": round(2.0 + ((quote_index % 5) * 0.7), 2),
        "total": line_total,
        "of": "",
        "desenho": "",
        "operacoes_lista": [part.strip() for part in str(spec["operations"]).split("+") if part.strip()],
    }
    return {
        "numero": quote_num,
        "data": backend.desktop_main.now_iso(),
        "estado": "Convertido em Encomenda" if approved else ("Enviado" if quote_index % 3 else "Em revisão"),
        "cliente": client_payload,
        "posto_trabalho": machine if approved else "",
        "linhas": [line],
        "iva_perc": 23.0,
        "desconto_perc": 0.0,
        "desconto_valor": 0.0,
        "preco_transporte": 0.0,
        "custo_transporte": 0.0,
        "paletes": 0.0,
        "peso_bruto_kg": 0.0,
        "volume_m3": 0.0,
        "transportadora_id": "",
        "transportadora_nome": "",
        "referencia_transporte": "",
        "zona_transporte": "",
        "subtotal_linhas": line_total,
        "subtotal_bruto": line_total,
        "subtotal": line_total,
        "total": round(line_total * 1.23, 2),
        "numero_encomenda": order_num if approved else "",
        "ano": date.today().year,
        "executado_por": "admin",
        "nota_transporte": "Simulacao interna",
        "notas_pdf": "",
        "nota_cliente": f"Lote operador {TAG} | cliente {client['codigo']} | quote {quote_index:04d}",
    }


def _direct_order_row(
    backend: LegacyBackend,
    data: dict[str, object],
    client: dict[str, str],
    quote_index: int,
    *,
    order_num: str,
    quote_num: str,
    ext_ref: str,
    ref_interna: str,
    machine: str,
) -> dict[str, object]:
    spec = MATERIAL_SPECS[(quote_index - 1) % len(MATERIAL_SPECS)]
    qty = float(2 + (quote_index % 9))
    laser_total = float(60 + ((quote_index - 1) % 5) * 30)
    pack_total = float(15 + ((quote_index - 1) % 4) * 15)
    piece = {
        "id": f"{order_num}-P01",
        "ref_interna": ref_interna,
        "ref_externa": ext_ref,
        "descricao": f"Conjunto corte {quote_index:04d}",
        "material": spec["material"],
        "espessura": spec["espessura"],
        "quantidade_pedida": qty,
        "Operacoes": spec["operations"],
        "Observacoes": f"{TAG} conjunto {quote_index:04d}",
        "desenho": "",
        "tempo_peca_min": round((laser_total + pack_total) / max(qty, 1.0), 2),
        "tempos_operacao": {
            "Corte Laser": round(laser_total, 2),
            "Embalamento": round(pack_total, 2),
        },
        "custos_operacao": {},
        "operacoes_detalhe": [],
        "of": backend.desktop_main.next_of_numero(data),
        "opp": backend.desktop_main.next_opp_numero(data),
        "estado": "Preparacao",
        "produzido_ok": 0.0,
        "produzido_nok": 0.0,
        "produzido_qualidade": 0.0,
        "inicio_producao": "",
        "fim_producao": "",
        "hist": [],
        "tempo_producao_min": 0.0,
        "lote_baixa": "",
        "qtd_expedida": 0.0,
        "expedicoes": [],
    }
    piece["operacoes_fluxo"] = backend.desktop_main.build_operacoes_fluxo(spec["operations"])
    backend.desktop_main.ensure_peca_operacoes(piece)
    backend.desktop_main.atualizar_estado_peca(piece)
    esp_bucket = {
        "espessura": spec["espessura"],
        "tempo_min": round(laser_total, 2),
        "tempos_operacao": {
            "Corte Laser": round(laser_total, 2),
            "Embalamento": round(pack_total, 2),
        },
        "maquinas_operacao": {
            "Corte Laser": machine,
            "Embalamento": "Embalamento",
        },
        "estado": "Preparacao",
        "pecas": [piece],
        "inicio_producao": "",
        "fim_producao": "",
        "tempo_producao_min": 0.0,
        "lote_baixa": "",
    }
    order = {
        "numero": order_num,
        "cliente": client["codigo"],
        "nota_cliente": f"{TAG} | operador backlog {quote_index:04d}",
        "nota_transporte": "",
        "preco_transporte": 0.0,
        "custo_transporte": 0.0,
        "paletes": 0.0,
        "peso_bruto_kg": 0.0,
        "volume_m3": 0.0,
        "transportadora_id": "",
        "transportadora_nome": "",
        "referencia_transporte": "",
        "zona_transporte": "",
        "local_descarga": f"Zona Industrial {((quote_index - 1) % 40) + 1}, Portugal",
        "transporte_numero": "",
        "estado_transporte": "",
        "data_criacao": backend.desktop_main.now_iso(),
        "data_entrega": (date.today() + timedelta(days=3 + (quote_index % 18))).isoformat(),
        "tempo": round((laser_total + pack_total) / 60.0, 2),
        "tempo_estimado": round(laser_total + pack_total, 2),
        "cativar": False,
        "posto_trabalho": machine,
        "observacoes": f"{TAG} | origem orcamento {quote_num}",
        "alerta_conversao": True,
        "estado": "Preparacao",
        "materiais": [
            {
                "material": spec["material"],
                "estado": "Preparacao",
                "espessuras": [esp_bucket],
            }
        ],
        "reservas": [],
        "montagem_itens": [],
        "numero_orcamento": quote_num,
        "estado_expedicao": "Nao expedida",
    }
    backend._ensure_unique_order_piece_refs(order)
    backend.desktop_main.update_refs(data, ref_interna, ext_ref)
    backend.desktop_main.update_estado_encomenda_por_espessuras(order)
    return order


def _create_quotes_and_orders(backend: LegacyBackend, clients: list[dict[str, str]]) -> dict[str, object]:
    data = backend.ensure_data()
    quote_numbers: list[str] = []
    order_numbers: list[str] = []
    ref_externas: list[str] = []
    planned_order_numbers: list[str] = []
    resource_plan_targets: dict[str, list[str]] = {resource: [] for resource in LASER_RESOURCES}

    for quote_index in range(1, TARGET_QUOTES + 1):
        client = clients[quote_index - 1]
        approved = quote_index <= TARGET_CONVERTED
        machine = LASER_RESOURCES[(quote_index - 1) % len(LASER_RESOURCES)] if approved else None
        quote_num = str(backend.desktop_main.next_orc_numero(data))
        order_num = str(backend.desktop_main.next_encomenda_numero(data)) if approved else ""
        ext_ref = f"LGS-OPR-{quote_index:04d}-01"
        ref_interna = str(backend.desktop_main.next_ref_interna_unique(data, client["codigo"], []))
        quote = _direct_quote_row(
            backend,
            client,
            quote_index,
            quote_num=quote_num,
            order_num=order_num,
            ext_ref=ext_ref,
            ref_interna=ref_interna,
            machine=machine or "",
            approved=approved,
        )
        data.setdefault("orcamentos", []).append(quote)
        quote_numbers.append(quote_num)
        ref_externas.append(ext_ref)
        data.setdefault("orc_refs", {})[ext_ref] = {
            "ref_interna": ref_interna,
            "ref_externa": ext_ref,
            "descricao": quote["linhas"][0]["descricao"],
            "material": quote["linhas"][0]["material"],
            "espessura": quote["linhas"][0]["espessura"],
            "preco_unit": quote["linhas"][0]["preco_unit"],
            "operacao": quote["linhas"][0]["operacao"],
            "tempo_peca_min": quote["linhas"][0]["tempo_peca_min"],
            "operacoes_detalhe": [],
            "tempos_operacao": {
                "Corte Laser": round(60 + ((quote_index - 1) % 5) * 30, 2),
                "Embalamento": round(15 + ((quote_index - 1) % 4) * 15, 2),
            },
            "custos_operacao": {},
            "desenho": "",
        }

        if approved:
            order = _direct_order_row(
                backend,
                data,
                client,
                quote_index,
                order_num=order_num,
                quote_num=quote_num,
                ext_ref=ext_ref,
                ref_interna=ref_interna,
                machine=machine or "Maquina 3030",
            )
            data.setdefault("encomendas", []).append(order)
            order_numbers.append(order_num)
            data.setdefault("peca_hist", {})[ext_ref] = {
                "ref_interna": ref_interna,
                "descricao": order["materiais"][0]["espessuras"][0]["pecas"][0]["descricao"],
                "material": order["materiais"][0]["material"],
                "espessura": order["materiais"][0]["espessuras"][0]["espessura"],
                "Operacoes": order["materiais"][0]["espessuras"][0]["pecas"][0]["Operacoes"],
                "Observacoes": order["materiais"][0]["espessuras"][0]["pecas"][0]["Observacoes"],
                "desenho": "",
            }

            if len(planned_order_numbers) < TARGET_PLANNED:
                planned_order_numbers.append(order_num)
                resource_plan_targets[machine or "Maquina 3030"].append(order_num)

        if quote_index % 25 == 0 or quote_index == TARGET_QUOTES:
            _print_progress("orcamentos", quote_index, TARGET_QUOTES)

    return {
        "quote_numbers": quote_numbers,
        "order_numbers": order_numbers,
        "ref_externas": ref_externas,
        "planned_order_numbers": planned_order_numbers,
        "resource_plan_targets": resource_plan_targets,
    }


def _plan_laser_orders(
    backend: LegacyBackend,
    resource_targets: dict[str, list[str]],
    *,
    target_total: int,
) -> tuple[dict[str, list[str]], list[str]]:
    week_start = _week_start_today()
    placed_by_resource: dict[str, list[str]] = {}
    placed_all: list[str] = []
    for resource, target_orders in resource_targets.items():
        pending = [
            dict(row)
            for row in list(backend.planning_pending_rows(operation="Corte Laser", resource=resource) or [])
            if str(row.get("numero", "") or "").strip() in set(target_orders)
        ]
        placed = list(backend.planning_auto_plan(pending, week_start=week_start, operation="Corte Laser") or [])
        placed_numbers = []
        for row in placed:
            number = str(row.get("encomenda", "") or "").strip()
            if number and number not in placed_numbers:
                placed_numbers.append(number)
        placed_by_resource[resource] = placed_numbers
        for number in placed_numbers:
            if number not in placed_all:
                placed_all.append(number)

    if len(placed_all) < target_total:
        missing = target_total - len(placed_all)
        for resource in LASER_RESOURCES:
            if missing <= 0:
                break
            current_placed = set(placed_by_resource.get(resource, []) or [])
            resource_target_set = set(resource_targets.get(resource, []) or [])
            extra_pending = [
                dict(row)
                for row in list(backend.planning_pending_rows(operation="Corte Laser", resource=resource) or [])
                if str(row.get("numero", "") or "").strip() not in current_placed
                and str(row.get("numero", "") or "").strip() not in resource_target_set
            ]
            extra_pending.sort(
                key=lambda row: (
                    str(row.get("data_entrega", "") or ""),
                    str(row.get("numero", "") or ""),
                )
            )
            for row in extra_pending:
                if missing <= 0:
                    break
                try:
                    placed = list(
                        backend.planning_auto_plan([row], week_start=week_start, operation="Corte Laser") or []
                    )
                except Exception:
                    continue
                if not placed:
                    continue
                number = str(placed[0].get("encomenda", "") or "").strip()
                if not number or number in placed_all:
                    continue
                placed_by_resource.setdefault(resource, []).append(number)
                placed_all.append(number)
                missing -= 1

    return placed_by_resource, placed_all


def _finish_orders(backend: LegacyBackend, order_numbers: list[str]) -> list[str]:
    finished: list[str] = []
    for index, order_num in enumerate(order_numbers, start=1):
        detail = backend.order_detail(order_num)
        pieces = list(detail.get("pieces", []) or [])
        if not pieces:
            continue
        piece = dict(pieces[0])
        piece_id = str(piece.get("id", "") or "").strip()
        qty_raw = piece.get("quantidade_pedida", piece.get("qtd_plan", piece.get("quantidade", 0)))
        try:
            qty = float(str(qty_raw).replace(",", "."))
        except Exception:
            qty = 0.0
        if qty <= 0:
            continue
        _start_with_retry(backend, order_num, piece_id, "admin", "Corte Laser")
        _finish_with_retry(backend, order_num, piece_id, "admin", "Corte Laser", qty)
        _start_with_retry(backend, order_num, piece_id, "admin", "Embalamento")
        _finish_with_retry(backend, order_num, piece_id, "admin", "Embalamento", qty)
        finished.append(order_num)
        if index % 5 == 0 or index == len(order_numbers):
            _print_progress("encomendas finalizadas", index, len(order_numbers))
    return finished


def main() -> int:
    backend = LegacyBackend()
    manifest = _load_manifest()
    if manifest and str(manifest.get("tag", "") or "").strip() == TAG:
        print("cleanup-anterior: sim")
        _clean_generated_dataset(backend, manifest)
        backend.reload(force=True)

    clients = _create_clients(backend)
    created = _create_quotes_and_orders(backend, clients)
    backend._save(force=True)
    backend.reload(force=True)

    resource_targets = dict(created["resource_plan_targets"])
    placed_by_resource, planned_order_numbers = _plan_laser_orders(
        backend,
        resource_targets,
        target_total=TARGET_PLANNED,
    )
    finish_targets = planned_order_numbers[:TARGET_FINISHED]
    finished_orders = _finish_orders(backend, finish_targets)

    week_start = _week_start_today().isoformat()
    manifest_payload = {
        "tag": TAG,
        "created_at": backend.desktop_main.now_iso(),
        "week_start": week_start,
        "client_codes": [row["codigo"] for row in clients],
        "quote_numbers": list(created["quote_numbers"]),
        "order_numbers": list(created["order_numbers"]),
        "planned_order_numbers": planned_order_numbers,
        "finished_order_numbers": finished_orders,
        "ref_externas": list(created["ref_externas"]),
        "resource_plan_targets": resource_targets,
        "placed_by_resource": placed_by_resource,
    }
    _save_manifest(manifest_payload)

    print(
        json.dumps(
            {
                "tag": TAG,
                "manifest": str(MANIFEST_PATH),
                "clientes_criados": len(clients),
                "orcamentos_criados": len(created["quote_numbers"]),
                "orcamentos_convertidos": len(created["order_numbers"]),
                "encomendas_planeadas": len(planned_order_numbers),
                "planeadas_3030": len(placed_by_resource.get("Maquina 3030", [])),
                "planeadas_5030": len(placed_by_resource.get("Maquina 5030", [])),
                "encomendas_finalizadas": len(finished_orders),
                "week_start": week_start,
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
