from __future__ import annotations

import json
import sys
from copy import deepcopy
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_qt.services.main_bridge import LegacyBackend


def _snapshot(data: dict) -> dict:
    return json.loads(json.dumps(data, ensure_ascii=False))


def _today_iso(backend: LegacyBackend) -> str:
    return str(backend.desktop_main.now_iso())[:10]


def _backup_runtime(backend: LegacyBackend) -> Path:
    data = _snapshot(backend.ensure_data())
    stamp = str(backend.desktop_main.now_iso()).replace(":", "").replace("-", "").replace("T", "_")[:15]
    out_dir = ROOT / "backups"
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / f"internal_runtime_before_reset_{stamp}.json"
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def _reset_runtime(backend: LegacyBackend) -> Path:
    backup_path = _backup_runtime(backend)
    current = _snapshot(backend.ensure_data())
    fresh = backend.desktop_main._copy_default_data()
    fresh["users"] = list(current.get("users", []) or [])
    fresh["operadores"] = list(
        dict.fromkeys(
            [str(value).strip() for value in list(current.get("operadores", []) or []) if str(value).strip()]
            + ["admin", "Operador 1", "Operador 2"]
        )
    )
    fresh["orcamentistas"] = list(
        dict.fromkeys(
            [str(value).strip() for value in list(current.get("orcamentistas", []) or []) if str(value).strip()]
            + ["admin", "Orçamentista 1"]
        )
    )
    today_iso = _today_iso(backend)
    fresh["at_series"] = [
        {
            "doc_type": "GT",
            "serie_id": f"GT{today_iso[:4]}",
            "inicio_sequencia": 1,
            "next_seq": 1,
            "data_inicio_prevista": today_iso,
            "validation_code": f"TESTE-GT{today_iso[:4]}",
            "status": "REGISTADA",
            "last_error": "",
            "last_sent_payload_hash": "",
            "updated_at": backend.desktop_main.now_iso(),
        },
        {
            "doc_type": "FT",
            "serie_id": f"FT{today_iso[:4]}",
            "inicio_sequencia": 1,
            "next_seq": 1,
            "data_inicio_prevista": today_iso,
            "validation_code": f"TESTE-FT{today_iso[:4]}",
            "status": "REGISTADA",
            "last_error": "",
            "last_sent_payload_hash": "",
            "updated_at": backend.desktop_main.now_iso(),
        },
    ]
    backend.data = fresh
    backend._save(force=True)
    return backup_path


def _ensure_branding(backend: LegacyBackend) -> None:
    current = backend.branding_settings()
    emit = dict(current.get("guia_emitente", {}) or {})
    if not str(emit.get("nome", "") or "").strip():
        emit["nome"] = "LuGEST Industrial, Lda."
    if not str(emit.get("nif", "") or "").strip():
        emit["nif"] = "509999990"
    if not str(emit.get("morada", "") or "").strip():
        emit["morada"] = "Rua da Industria 100, 4700-000 Braga"
    if not str(emit.get("local_carga", "") or "").strip():
        emit["local_carga"] = emit["morada"]
    rodape = list(current.get("empresa_info_rodape", []) or [])
    if not rodape:
        rodape = [
            "LuGEST Industrial, Lda.",
            "Rua da Industria 100, 4700-000 Braga",
            "NIF 509999990",
            "geral@lugest.local | +351 253 000 000",
        ]
    payload = dict(current)
    payload["guia_emitente"] = emit
    payload["empresa_info_rodape"] = rodape
    payload["guia_serie_id"] = f"GT{_today_iso(backend)[:4]}"
    payload["guia_validation_code"] = f"TESTE-GT{_today_iso(backend)[:4]}"
    backend.save_branding_settings(payload)


def _seed_clients(backend: LegacyBackend) -> list[dict]:
    rows = [
        {
            "nome": "Cliente Atlas",
            "nif": "500000001",
            "morada": "Rua do Atlas 10, Braga",
            "contacto": "253100001",
            "email": "atlas@cliente.pt",
            "cond_pagamento": "30 dias",
        },
        {
            "nome": "Cliente Boreal",
            "nif": "500000002",
            "morada": "Avenida Boreal 20, Porto",
            "contacto": "220100002",
            "email": "boreal@cliente.pt",
            "cond_pagamento": "30 dias",
        },
        {
            "nome": "Cliente Cobalto",
            "nif": "500000003",
            "morada": "Zona Industrial Cobalto 3, Aveiro",
            "contacto": "234100003",
            "email": "cobalto@cliente.pt",
            "cond_pagamento": "45 dias",
        },
        {
            "nome": "Cliente Delta",
            "nif": "500000004",
            "morada": "Parque Delta 44, Leiria",
            "contacto": "244100004",
            "email": "delta@cliente.pt",
            "cond_pagamento": "30 dias",
        },
        {
            "nome": "Cliente Épsilon",
            "nif": "500000005",
            "morada": "Rua Épsilon 55, Lisboa",
            "contacto": "211100005",
            "email": "epsilon@cliente.pt",
            "cond_pagamento": "60 dias",
        },
    ]
    created: list[dict] = []
    for row in rows:
        created.append(backend.client_save(row))
    return created


def _seed_materials(backend: LegacyBackend) -> list[dict]:
    rows = []
    material_names = ["S235JR", "S275JR", "S355JR", "INOX304", "INOX316", "AL5754", "AL1050", "GALV", "CORTEN", "HARDOX"]
    thicknesses = ["1.5", "2", "3", "4", "5", "6", "8", "10", "12", "15"]
    for index, (material, espessura) in enumerate(zip(material_names, thicknesses), start=1):
        rows.append(
            backend.add_material(
                {
                    "formato": "Chapa",
                    "material": material,
                    "espessura": espessura,
                    "comprimento": 3000 + (index * 50),
                    "largura": 1500,
                    "quantidade": 8 + index,
                    "reservado": 0,
                    "p_compra": 125 + (index * 7),
                    "local": f"A-{index:02d}",
                    "lote_fornecedor": f"L{index:03d}",
                }
            )
        )
    return rows


def _seed_products(backend: LegacyBackend) -> list[dict]:
    rows = []
    for index in range(1, 11):
        rows.append(
            backend.product_save(
                {
                    "descricao": f"Produto Stock {index:02d}",
                    "categoria": "Consumiveis" if index <= 5 else "Componentes",
                    "tipo": "Parafusaria" if index <= 5 else "Montagem",
                    "dimensoes": f"M{4 + index}x{12 + index}",
                    "unid": "UN",
                    "qty": 40 + index,
                    "alerta": 10,
                    "p_compra": 0.35 + (index * 0.11),
                    "pvp1": 0.65 + (index * 0.18),
                    "fabricante": "Interno",
                    "modelo": f"STK-{index:02d}",
                    "obs": "Seed interno ERP",
                }
            )
        )
    return rows


def _history_quote_payload(client: dict, quote_index: int, materials: list[dict], product_code: str) -> dict:
    material_a = materials[(quote_index - 1) % len(materials)]
    material_b = materials[(quote_index + 1) % len(materials)]
    lines = []
    for ref_index in range(1, 6):
        lines.append(
            {
                "tipo_item": "Peca",
                "ref_externa": f"HIST-{client['codigo']}-{quote_index:02d}-{ref_index:02d}",
                "descricao": f"Historico {client['nome']} {quote_index:02d}/{ref_index:02d}",
                "material": str(material_a.get("material", "") or ""),
                "espessura": str(material_a.get("espessura", "") or ""),
                "operacao": "Corte Laser + Quinagem + Embalamento" if ref_index % 2 == 0 else "Corte Laser + Embalamento",
                "produto_codigo": "",
                "produto_unid": "",
                "qtd": 4 + ref_index,
                "preco_unit": round(12.5 + (quote_index * 1.4) + (ref_index * 0.75), 2),
                "tempo_peca_min": 2 + ref_index,
                "desenho": "",
            }
        )
        material_a, material_b = material_b, material_a
    return {
        "cliente": {"codigo": client["codigo"], "nome": client["nome"]},
        "estado": "Aprovado" if quote_index % 2 == 0 else "Em edição",
        "linhas": lines,
        "iva_perc": 23,
        "preco_transporte": 18.0 if quote_index == 1 else 0.0,
        "nota_cliente": f"Histórico inicial {client['nome']}",
        "executado_por": "admin",
    }


def _seed_quote_history(backend: LegacyBackend, clients: list[dict], materials: list[dict], products: list[dict]) -> list[dict]:
    created: list[dict] = []
    product_code = str((products[0] if products else {}).get("codigo", "") or "").strip()
    for index, client in enumerate(clients, start=1):
        created.append(backend.orc_save(_history_quote_payload(client, index, materials, product_code)))
    return created


def main() -> int:
    backend = LegacyBackend()
    backup_path = _reset_runtime(backend)
    _ensure_branding(backend)
    clients = _seed_clients(backend)
    materials = _seed_materials(backend)
    products = _seed_products(backend)
    history_quotes = _seed_quote_history(backend, clients, materials, products)
    backend.reload()
    print(
        json.dumps(
            {
                "backup": str(backup_path),
                "clientes": len(clients),
                "materiais": len(materials),
                "produtos": len(products),
                "orcamentos_historico": len(history_quotes),
                "refs_historico": 5 * len(clients),
            },
            ensure_ascii=False,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
