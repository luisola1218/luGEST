from __future__ import annotations

import copy
import re
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from pypdf import PdfReader

from lugest_qt.services.legacy_backend import LegacyBackend


def _assert(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def main() -> int:
    backend = LegacyBackend()
    data = copy.deepcopy(backend.ensure_data())
    models = [row for row in data.get("conjuntos", []) if isinstance(row, dict) and row.get("itens")]
    _assert(bool(models), "Nao existe um conjunto com componentes para validar.")

    model = max(models, key=lambda row: len(list(row.get("itens", []) or [])))
    model["param_codigo"] = str(model.get("param_codigo", "") or "0001").zfill(4)
    model["ficha_tecnica"] = {
        "familia": "Equipamento industrial",
        "aplicacao": "Integracao em linha de producao",
        "modelo": str(model.get("codigo", "") or ""),
        "versao": "REV-A",
        "estado_documento": "Libertado para producao",
        "responsavel_tecnico": "Engenharia",
        "normas": "Diretiva Maquinas / requisitos aplicaveis",
        "alimentacao": "Conforme projeto",
        "acabamento": "Conforme especificacao",
        "dimensoes_gerais": "Conforme desenho de conjunto",
        "peso_estimado": "Conforme BOM",
    }
    backend._replace_data_cache(data)
    backend._save = lambda *args, **kwargs: None

    expected_positions = len(list(model.get("itens", []) or []))
    output = Path(tempfile.gettempdir()) / "lugest_verify_conjunto_technical.pdf"
    try:
        rendered = backend.conjunto_sheet_pdf(str(model.get("codigo", "")), output)
        _assert(rendered.exists() and rendered.stat().st_size > 0, "PDF tecnico vazio ou inexistente.")
        reader = PdfReader(str(rendered))
        text = "\n".join(page.extract_text() or "" for page in reader.pages)
        normalized = " ".join(text.split())

        for required in (
            "DOSSIER TECNICO DE CONJUNTO",
            "BOM | BILL OF MATERIALS",
            "MATERIAL + QTD.",
            "PONTOS DE CONTROLO",
            f"CFG|{model['param_codigo']}|{model.get('codigo', '')}",
        ):
            _assert(required in normalized, f"Conteudo tecnico em falta: {required}")

        for forbidden in ("EUR", "Custo atual", "Final atual", "Preco atual", "Observacoes"):
            _assert(forbidden.lower() not in normalized.lower(), f"Conteudo proibido no PDF tecnico: {forbidden}")

        positions = {
            int(value)
            for value in re.findall(r"(?m)^\s*(\d{3})\s+(?:FAB\.|STOCK|OPER\.|M\.PRIMA)", text)
        }
        _assert(
            len(positions) == expected_positions,
            f"BOM incompleta: {len(positions)}/{expected_positions} posicoes.",
        )
        print(
            f"conjunto-technical-pdf-ok pages={len(reader.pages)} "
            f"positions={len(positions)} bytes={rendered.stat().st_size}"
        )
    finally:
        output.unlink(missing_ok=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
