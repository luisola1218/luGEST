import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import main

TAG_DEFAULT = "STRESS_20260302"


def _orc_is_tagged(o, tag):
    if not isinstance(o, dict):
        return False
    if tag in str(o.get("nota_cliente", "")):
        return True
    for l in o.get("linhas", []):
        if tag in str(l.get("ref_externa", "")):
            return True
    return False


def _enc_is_tagged(e, tag):
    if not isinstance(e, dict):
        return False
    if tag in str(e.get("nota_cliente", "")):
        return True
    if tag in str(e.get("Observações", "")) or tag in str(e.get("observacoes", "")):
        return True
    for p in main.encomenda_pecas(e):
        if tag in str(p.get("ref_externa", "")):
            return True
    return False


def _ne_is_tagged(ne, tag):
    if not isinstance(ne, dict):
        return False
    if tag in str(ne.get("obs", "")) or tag in str(ne.get("origem_cotacao", "")):
        return True
    for l in ne.get("linhas", []):
        if tag in str(l.get("descricao", "")):
            return True
    return False


def _count_tagged(data, tag):
    return {
        "clientes_tag": len([c for c in data.get("clientes", []) if tag in str(c.get("nome", ""))]),
        "produtos_tag": len([p for p in data.get("produtos", []) if tag in str(p.get("descricao", ""))]),
        "materiais_tag": len(
            [
                m
                for m in data.get("materiais", [])
                if tag in str(m.get("lote_fornecedor", ""))
                or tag in str(m.get("Localizacao", "") or m.get("Localização", ""))
            ]
        ),
        "orcamentos_tag": len([o for o in data.get("orcamentos", []) if _orc_is_tagged(o, tag)]),
        "encomendas_tag": len([e for e in data.get("encomendas", []) if _enc_is_tagged(e, tag)]),
        "notas_encomenda_tag": len([n for n in data.get("notas_encomenda", []) if _ne_is_tagged(n, tag)]),
        "plano_tag": len([p for p in data.get("plano", []) if tag in str(p.get("id", ""))]),
        "produtos_mov_tag": len([m for m in data.get("produtos_mov", []) if tag in str(m.get("obs", ""))]),
        "stock_log_tag": len([s for s in data.get("stock_log", []) if tag in str(s.get("detalhes", ""))]),
        "fornecedores_tag": len([f for f in data.get("fornecedores", []) if tag in str(f.get("nome", ""))]),
    }


def cleanup(tag):
    data = main.load_data()
    before = _count_tagged(data, tag)

    data["notas_encomenda"] = [n for n in data.get("notas_encomenda", []) if not _ne_is_tagged(n, tag)]
    data["plano"] = [p for p in data.get("plano", []) if tag not in str(p.get("id", ""))]
    data["plano_hist"] = [p for p in data.get("plano_hist", []) if tag not in str(p.get("id", ""))]
    data["encomendas"] = [e for e in data.get("encomendas", []) if not _enc_is_tagged(e, tag)]
    data["orcamentos"] = [o for o in data.get("orcamentos", []) if not _orc_is_tagged(o, tag)]
    data["produtos_mov"] = [m for m in data.get("produtos_mov", []) if tag not in str(m.get("obs", ""))]
    data["stock_log"] = [s for s in data.get("stock_log", []) if tag not in str(s.get("detalhes", ""))]
    data["produtos"] = [p for p in data.get("produtos", []) if tag not in str(p.get("descricao", ""))]
    data["materiais"] = [
        m
        for m in data.get("materiais", [])
        if tag not in str(m.get("lote_fornecedor", ""))
        and tag not in str(m.get("Localizacao", "") or m.get("Localização", ""))
    ]
    data["clientes"] = [c for c in data.get("clientes", []) if tag not in str(c.get("nome", ""))]
    data["fornecedores"] = [f for f in data.get("fornecedores", []) if tag not in str(f.get("nome", ""))]

    if isinstance(data.get("orc_refs"), dict):
        data["orc_refs"] = {
            k: v
            for k, v in data.get("orc_refs", {}).items()
            if tag not in str(k)
            and tag not in str((v or {}).get("descricao", ""))
            and tag not in str((v or {}).get("ref_interna", ""))
        }

    try:
        data["materiais_hist"] = [x for x in data.get("materiais_hist", []) if tag not in str(x)]
        data["espessuras_hist"] = [x for x in data.get("espessuras_hist", []) if tag not in str(x)]
    except Exception:
        pass

    try:
        main._rebuild_runtime_sequences(data)  # pylint: disable=protected-access
    except Exception:
        pass

    main.save_data(data, force=True)
    after_data = main.load_data()
    after = _count_tagged(after_data, tag)

    result = {
        "tag": tag,
        "executed_at": datetime.now().isoformat(timespec="seconds"),
        "before": before,
        "after": after,
    }
    out_dir = Path("backups")
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"stress_cleanup_{tag}.json"
    out_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    print(f"Report: {out_path}")


def cli():
    parser = argparse.ArgumentParser(description="Limpa dados de stress da base MySQL.")
    parser.add_argument("--tag", default=TAG_DEFAULT, help="Marcador dos dados de stress.")
    parser.add_argument(
        "--confirm",
        action="store_true",
        help="Obrigatorio para executar a limpeza.",
    )
    args = parser.parse_args()
    if not args.confirm:
        print("Use --confirm para executar a limpeza.")
        return
    cleanup(args.tag)


if __name__ == "__main__":
    cli()
