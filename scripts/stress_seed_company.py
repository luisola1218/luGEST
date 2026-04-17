import json
import random
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import main

TAG = "STRESS_20260302"
TARGET_CLIENTES = 500
TARGET_ORCAMENTOS = 500
TARGET_ENCOMENDAS = 500
TARGET_ENCOMENDAS_FINALIZADAS = 100
TARGET_NE = 100
TARGET_PRODUTOS_NE = 140
TARGET_MATERIAIS = 60


def _now_date(days=0):
    return (datetime.now() + timedelta(days=days)).strftime("%Y-%m-%d")


def _next_mat_id(data):
    max_n = 0
    for m in data.get("materiais", []):
        mid = str(m.get("id", ""))
        if mid.startswith("MAT") and mid[3:].isdigit():
            max_n = max(max_n, int(mid[3:]))
    return f"MAT{max_n + 1:05d}"


def _tagged_clientes(data):
    return [c for c in data.get("clientes", []) if TAG in str(c.get("nome", ""))]


def _tagged_produtos(data):
    return [p for p in data.get("produtos", []) if TAG in str(p.get("descricao", ""))]


def _tagged_materiais(data):
    out = []
    for m in data.get("materiais", []):
        loc = str(m.get("Localizacao", "") or m.get("Localização", "") or "")
        lote = str(m.get("lote_fornecedor", "") or "")
        if TAG in loc or TAG in lote:
            out.append(m)
    return out


def _orc_is_tagged(o):
    if not isinstance(o, dict):
        return False
    if TAG in str(o.get("nota_cliente", "")):
        return True
    for l in o.get("linhas", []):
        if TAG in str(l.get("ref_externa", "")):
            return True
    return False


def _enc_is_tagged(e):
    if not isinstance(e, dict):
        return False
    if TAG in str(e.get("nota_cliente", "")):
        return True
    if TAG in str(e.get("Observações", "")) or TAG in str(e.get("observacoes", "")):
        return True
    for p in main.encomenda_pecas(e):
        if TAG in str(p.get("ref_externa", "")):
            return True
    return False


def _ne_is_tagged(ne):
    if not isinstance(ne, dict):
        return False
    if TAG in str(ne.get("obs", "")):
        return True
    if TAG in str(ne.get("origem_cotacao", "")):
        return True
    for l in ne.get("linhas", []):
        if TAG in str(l.get("descricao", "")):
            return True
    return False


def _ensure_fornecedor_stress(data):
    for f in data.get("fornecedores", []):
        if TAG in str(f.get("nome", "")):
            return f
    fid = main.next_fornecedor_numero(data)
    f = {
        "id": fid,
        "nome": f"{TAG} FORNECEDOR",
        "nif": "999999999",
        "morada": f"{TAG} MORADA",
        "contacto": "219999999",
        "email": "stress@empresa.local",
        "codigo_postal": "1000-000",
        "localidade": "Lisboa",
        "pais": "Portugal",
        "cond_pagamento": "30 dias",
        "prazo_entrega_dias": "7",
        "website": "",
        "obs": TAG,
    }
    data.setdefault("fornecedores", []).append(f)
    return f


def _ensure_clientes(data):
    existing = _tagged_clientes(data)
    need = max(0, TARGET_CLIENTES - len(existing))
    for _ in range(need):
        codigo = main.next_cliente_codigo(data)
        idx = len(existing) + 1
        c = {
            "codigo": codigo,
            "nome": f"{TAG} CLIENTE {idx:04d}",
            "nif": f"9{idx:08d}"[:9],
            "morada": f"{TAG} MORADA {idx:04d}",
            "contacto": f"91{idx:07d}"[:9],
            "email": f"stress_cliente_{idx:04d}@empresa.local",
            "prazo_entrega": "5 dias",
            "cond_pagamento": "30 dias",
            "obs_tecnicas": TAG,
        }
        data.setdefault("clientes", []).append(c)
        existing.append(c)
    return existing


def _ensure_produtos(data):
    existing = _tagged_produtos(data)
    need = max(0, TARGET_PRODUTOS_NE - len(existing))
    for _ in range(need):
        codigo = main.next_produto_numero(data)
        idx = len(existing) + 1
        p = {
            "codigo": codigo,
            "descricao": f"{TAG} PRODUTO {idx:04d}",
            "categoria": "Consumiveis Soldadura",
            "subcat": "Diversos",
            "tipo": "Outros",
            "unid": "UN",
            "qty": float(200 + (idx % 15)),
            "alerta": 10.0,
            "p_compra": round(5.0 + (idx % 17) * 0.35, 4),
            "atualizado_em": main.now_iso(),
        }
        data.setdefault("produtos", []).append(p)
        existing.append(p)
    return existing


def _ensure_materiais(data):
    existing = _tagged_materiais(data)
    need = max(0, TARGET_MATERIAIS - len(existing))
    formas = ["Chapa", "Tubo", "Perfil"]
    espessuras = ["1.5", "2", "3", "4", "5", "6", "8", "10"]
    for _ in range(need):
        idx = len(existing) + 1
        formato = formas[idx % len(formas)]
        esp = espessuras[idx % len(espessuras)]
        comprimento = 6000.0 if formato in ("Chapa", "Perfil") else 0.0
        largura = 1500.0 if formato == "Chapa" else (80.0 if formato == "Perfil" else 0.0)
        metros = 6.0 if formato == "Tubo" else 0.0
        m = {
            "id": _next_mat_id(data),
            "formato": formato,
            "material": f"{TAG}_MATERIAL_{(idx % 12) + 1:02d}",
            "espessura": esp,
            "comprimento": comprimento,
            "largura": largura,
            "metros": metros,
            "quantidade": float(3000 + idx * 7),
            "reservado": 0.0,
            "Localizacao": f"{TAG}_RACK_{(idx % 20) + 1:02d}",
            "lote_fornecedor": f"{TAG}_LOTE_{idx:04d}",
            "peso_unid": round(3.5 + (idx % 11) * 0.41, 3),
            "p_compra": round(1.7 + (idx % 13) * 0.19, 4),
            "is_sobra": False,
            "atualizado_em": main.now_iso(),
        }
        data.setdefault("materiais", []).append(m)
        existing.append(m)
        main.push_unique(data.setdefault("materiais_hist", []), m.get("material", ""))
        main.push_unique(data.setdefault("espessuras_hist", []), m.get("espessura", ""))
    return existing


def _create_orcamentos_encomendas(data, clientes, materiais):
    existing_orc = [o for o in data.get("orcamentos", []) if _orc_is_tagged(o)]
    existing_enc = [e for e in data.get("encomendas", []) if _enc_is_tagged(e)]
    need = max(0, TARGET_ORCAMENTOS - len(existing_orc))
    today = datetime.now()

    for i in range(need):
        idx = len(existing_orc) + 1
        cli = clientes[(idx - 1) % len(clientes)]
        mat = materiais[(idx - 1) % len(materiais)]
        qtd = float(5 + (idx % 12))
        preco_unit = round(95 + (idx % 37) * 1.8, 2)
        total_linha = round(qtd * preco_unit, 2)
        ref_int = main.next_ref_interna_unique(data, cli.get("codigo", ""))
        ref_ext = f"{TAG}-REFEXT-{idx:04d}"
        orc_num = main.next_orc_numero(data)
        of_cod = main.next_of_numero(data)
        opp_cod = main.next_opp_numero(data)
        prazo = today + timedelta(days=(idx % 9) + 1)

        linha = {
            "ref_interna": ref_int,
            "ref_externa": ref_ext,
            "descricao": f"{TAG} PECA {idx:04d}",
            "material": mat.get("material", ""),
            "espessura": str(mat.get("espessura", "")).strip(),
            "operacao": "Corte Laser + Quinagem + Embalamento",
            "of": of_cod,
            "qtd": qtd,
            "preco_unit": preco_unit,
            "total": total_linha,
            "desenho": "",
        }
        orc = {
            "numero": orc_num,
            "data": main.now_iso(),
            "estado": "Aprovado",
            "cliente": {
                "codigo": cli.get("codigo", ""),
                "nome": cli.get("nome", ""),
                "empresa": cli.get("nome", ""),
                "nif": cli.get("nif", ""),
                "morada": cli.get("morada", ""),
                "contacto": cli.get("contacto", ""),
                "email": cli.get("email", ""),
            },
            "iva_perc": 23.0,
            "subtotal": total_linha,
            "total": round(total_linha * 1.23, 2),
            "numero_encomenda": "",
            "nota_cliente": f"{TAG} Orcamento aprovado e convertido",
            "executado_por": "stress-bot",
            "nota_transporte": TAG,
            "notas_pdf": TAG,
            "linhas": [linha],
        }
        data.setdefault("orcamentos", []).append(orc)
        existing_orc.append(orc)

        enc_num = main.next_encomenda_numero(data)
        peca = {
            "id": f"{enc_num}-001",
            "ref_interna": ref_int,
            "ref_externa": ref_ext,
            "material": linha.get("material", ""),
            "espessura": linha.get("espessura", ""),
            "Operacoes": linha.get("operacao", ""),
            "quantidade_pedida": qtd,
            "of": of_cod,
            "opp": opp_cod,
            "estado": "Preparacao",
            "produzido_ok": 0.0,
            "produzido_nok": 0.0,
            "inicio_producao": "",
            "fim_producao": "",
            "tempo_producao_min": 0.0,
            "lote_baixa": "",
            "desenho": "",
            "hist": [],
            "qtd_expedida": 0.0,
        }
        main.ensure_peca_operacoes(peca)
        main.atualizar_estado_peca(peca)
        esp_obj = {
            "espessura": linha.get("espessura", ""),
            "tempo_min": 90.0,
            "estado": "Preparacao",
            "pecas": [peca],
            "inicio_producao": "",
            "fim_producao": "",
            "tempo_producao_min": 0.0,
            "lote_baixa": "",
        }
        mat_obj = {
            "material": linha.get("material", ""),
            "estado": "Preparacao",
            "espessuras": [esp_obj],
        }
        enc = {
            "id": f"ENCST{idx:05d}",
            "numero": enc_num,
            "cliente": cli.get("codigo", ""),
            "nota_cliente": f"{TAG} Encomenda gerada automaticamente",
            "data_criacao": main.now_iso(),
            "data_entrega": prazo.strftime("%Y-%m-%d"),
            "tempo_estimado": 120.0,
            "tempo": 120.0,
            "cativar": False,
            "Observações": f"{TAG} fluxo orc->enc",
            "observacoes": f"{TAG} fluxo orc->enc",
            "estado": "Preparacao",
            "espessuras": [],
            "reservas": [],
            "materiais": [mat_obj],
            "numero_orcamento": orc_num,
            "estado_expedicao": "Nao expedida",
        }
        main.update_estado_encomenda_por_espessuras(enc)
        data.setdefault("encomendas", []).append(enc)
        existing_enc.append(enc)
        orc["numero_encomenda"] = enc_num

    return existing_orc, existing_enc


def _plan_week(data, encomendas_tag):
    data["plano"] = [p for p in data.get("plano", []) if TAG not in str(p.get("id", ""))]
    base = datetime.now().date()
    slots = ["08:00", "09:30", "11:00", "13:30", "15:00", "16:30"]
    for i, enc in enumerate(encomendas_tag):
        pecas = main.encomenda_pecas(enc)
        if not pecas:
            continue
        p0 = pecas[0]
        item = {
            "id": f"{TAG}-BLK-{i + 1:04d}-{enc.get('numero', '')}",
            "encomenda": enc.get("numero", ""),
            "material": p0.get("material", ""),
            "espessura": p0.get("espessura", ""),
            "data": str(base + timedelta(days=i % 7)),
            "inicio": slots[i % len(slots)],
            "duracao_min": float(60 + (i % 4) * 30),
            "color": "#d9534f",
            "chapa": f"{TAG}-CHAPA-{(i % 50) + 1:03d}",
        }
        data.setdefault("plano", []).append(item)


def _finalizar_e_baixar(data, encomendas_tag, materiais_tag):
    final_list = encomendas_tag[:TARGET_ENCOMENDAS_FINALIZADAS]
    mat_index = {}
    for m in materiais_tag:
        k = (str(m.get("material", "")).strip(), str(m.get("espessura", "")).strip())
        mat_index.setdefault(k, []).append(m)

    for enc in final_list:
        for m in enc.get("materiais", []):
            m["estado"] = "Concluida"
            for e in m.get("espessuras", []):
                e["estado"] = "Concluida"
                e["tempo_producao_min"] = main.parse_float(e.get("tempo_min", 90), 90)
                e["inicio_producao"] = e.get("inicio_producao") or main.now_iso()
                e["fim_producao"] = main.now_iso()
                for p in e.get("pecas", []):
                    qtd = main.parse_float(p.get("quantidade_pedida", 0), 0)
                    p["inicio_producao"] = p.get("inicio_producao") or main.now_iso()
                    p["fim_producao"] = main.now_iso()
                    p["tempo_producao_min"] = max(30.0, main.parse_float(p.get("tempo_producao_min", 0), 0))
                    p["produzido_ok"] = qtd
                    p["produzido_nok"] = 0.0
                    p["estado"] = "Concluida"
                    main.ensure_peca_operacoes(p)
                    main.atualizar_estado_peca(p)

                    mk = (str(p.get("material", "")).strip(), str(p.get("espessura", "")).strip())
                    options = mat_index.get(mk) or []
                    if options:
                        sm = options[0]
                        before = main.parse_float(sm.get("quantidade", 0), 0)
                        after = max(0.0, before - qtd)
                        sm["quantidade"] = after
                        sm["atualizado_em"] = main.now_iso()
                        main.log_stock(
                            data,
                            "BAIXA_STRESS",
                            f"{TAG} enc={enc.get('numero','')} mat={sm.get('id','')} qtd={qtd}",
                        )

        main.update_estado_encomenda_por_espessuras(enc)
        enc["estado"] = "Concluida"


def _create_notas_encomenda(data, fornecedor, produtos_tag):
    notas_tag = [n for n in data.get("notas_encomenda", []) if _ne_is_tagged(n)]
    need = max(0, TARGET_NE - len(notas_tag))
    for _ in range(need):
        idx = len(notas_tag) + 1
        ne_num = main.next_ne_numero(data)
        linhas = []
        total = 0.0
        refs = []
        for j in range(3):
            prod = produtos_tag[(idx * 3 + j) % len(produtos_tag)]
            qtd = float(4 + ((idx + j) % 9))
            preco = max(0.01, main.produto_preco_unitario(prod))
            subtotal = round(qtd * preco, 2)
            total += subtotal
            before = main.parse_float(prod.get("qty", 0), 0)
            after = before + qtd
            prod["qty"] = after
            prod["atualizado_em"] = main.now_iso()
            guia = f"GT-{TAG}-{idx:03d}"
            fatura = f"FT-{TAG}-{idx:03d}"
            linha = {
                "ref": prod.get("codigo", ""),
                "descricao": f"{prod.get('descricao', '')} | {TAG}",
                "fornecedor_linha": f"{fornecedor.get('id','')} - {fornecedor.get('nome','')}".strip(" -"),
                "origem": "Produto",
                "qtd": qtd,
                "unid": prod.get("unid", "UN"),
                "preco": preco,
                "total": subtotal,
                "entregue": True,
                "qtd_entregue": qtd,
                "_stock_in": True,
                "guia_entrega": guia,
                "fatura_entrega": fatura,
                "data_doc_entrega": _now_date(),
                "data_entrega_real": _now_date(),
                "obs_entrega": TAG,
                "entregas_linha": [
                    {
                        "data_registo": main.now_iso(),
                        "data_entrega": _now_date(),
                        "data_documento": _now_date(),
                        "guia": guia,
                        "fatura": fatura,
                        "obs": TAG,
                        "qtd": qtd,
                    }
                ],
            }
            refs.append(linha["ref"])
            linhas.append(linha)
            data.setdefault("produtos_mov", []).append(
                {
                    "data": main.now_iso(),
                    "tipo": "Entrada",
                    "operador": "stress-bot",
                    "codigo": prod.get("codigo", ""),
                    "descricao": prod.get("descricao", ""),
                    "qtd": qtd,
                    "antes": before,
                    "depois": after,
                    "obs": TAG,
                    "origem": "NE_STRESS",
                    "ref_doc": ne_num,
                }
            )

        ne = {
            "numero": ne_num,
            "fornecedor": f"{fornecedor.get('id','')} - {fornecedor.get('nome','')}".strip(" -"),
            "fornecedor_id": fornecedor.get("id", ""),
            "contacto": fornecedor.get("contacto", ""),
            "data_entrega": _now_date(),
            "obs": f"{TAG} nota de encomenda automatica",
            "local_descarga": f"{TAG} armazem",
            "meio_transporte": "Transportadora",
            "linhas": linhas,
            "estado": "Entregue",
            "total": round(total, 2),
            "oculta": False,
            "_draft": False,
            "origem_cotacao": TAG,
            "ne_geradas": [],
            "data_ultima_entrega": _now_date(),
            "guia_ultima": f"GT-{TAG}-{idx:03d}",
            "fatura_ultima": f"FT-{TAG}-{idx:03d}",
            "fatura_caminho_ultima": "",
            "data_doc_ultima": _now_date(),
            "entregas": [
                {
                    "data_registo": main.now_iso(),
                    "data_entrega": _now_date(),
                    "guia": f"GT-{TAG}-{idx:03d}",
                    "fatura": f"FT-{TAG}-{idx:03d}",
                    "data_documento": _now_date(),
                    "obs": TAG,
                    "linhas": refs,
                    "quantidade_linhas": len(linhas),
                    "quantidade_total": sum(main.parse_float(l.get("qtd", 0), 0) for l in linhas),
                }
            ],
        }
        data.setdefault("notas_encomenda", []).append(ne)
        notas_tag.append(ne)


def _manifest_path():
    out_dir = Path("backups")
    out_dir.mkdir(parents=True, exist_ok=True)
    return out_dir / f"stress_manifest_{TAG}.json"


def main_seed():
    t0 = time.time()
    random.seed(26032026)

    data = main.load_data()
    baseline = {
        "clientes": len(data.get("clientes", [])),
        "orcamentos": len(data.get("orcamentos", [])),
        "encomendas": len(data.get("encomendas", [])),
        "plano": len(data.get("plano", [])),
        "notas_encomenda": len(data.get("notas_encomenda", [])),
        "materiais": len(data.get("materiais", [])),
        "produtos": len(data.get("produtos", [])),
        "produtos_mov": len(data.get("produtos_mov", [])),
    }

    clientes = _ensure_clientes(data)
    produtos = _ensure_produtos(data)
    materiais = _ensure_materiais(data)
    fornecedor = _ensure_fornecedor_stress(data)
    _create_orcamentos_encomendas(data, clientes, materiais)

    encomendas_tag = [e for e in data.get("encomendas", []) if _enc_is_tagged(e)]
    encomendas_tag.sort(key=lambda x: str(x.get("numero", "")))
    _plan_week(data, encomendas_tag[:TARGET_ENCOMENDAS])
    _finalizar_e_baixar(data, encomendas_tag, materiais)
    _create_notas_encomenda(data, fornecedor, produtos)

    save_start = time.time()
    main.save_data(data, force=True)
    save_elapsed = time.time() - save_start

    data2 = main.load_data()
    tagged = {
        "clientes_tag": len(_tagged_clientes(data2)),
        "produtos_tag": len(_tagged_produtos(data2)),
        "materiais_tag": len(_tagged_materiais(data2)),
        "orcamentos_tag": len([o for o in data2.get("orcamentos", []) if _orc_is_tagged(o)]),
        "encomendas_tag": len([e for e in data2.get("encomendas", []) if _enc_is_tagged(e)]),
        "encomendas_tag_concluidas": len(
            [
                e
                for e in data2.get("encomendas", [])
                if _enc_is_tagged(e) and "concl" in main.norm_text(e.get("estado", ""))
            ]
        ),
        "notas_encomenda_tag": len([n for n in data2.get("notas_encomenda", []) if _ne_is_tagged(n)]),
        "plano_tag": len([p for p in data2.get("plano", []) if TAG in str(p.get("id", ""))]),
        "produtos_mov_tag": len([m for m in data2.get("produtos_mov", []) if TAG in str(m.get("obs", ""))]),
    }

    manifest = {
        "tag": TAG,
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "baseline": baseline,
        "tagged_counts": tagged,
        "save_elapsed_sec": round(save_elapsed, 3),
        "total_elapsed_sec": round(time.time() - t0, 3),
        "notes": "Apagar apenas com script de cleanup usando o mesmo TAG.",
    }
    mp = _manifest_path()
    mp.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(manifest, ensure_ascii=False, indent=2))
    print(f"Manifest: {mp}")


if __name__ == "__main__":
    main_seed()
