from __future__ import annotations

import json
import sys
from pathlib import Path
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_core.intelligence.remote_product_ai import RemoteProductAIClient


class _FakeResponse:
    def __init__(self, payload: dict) -> None:
        self._raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")

    def __enter__(self):
        return self

    def __exit__(self, _exc_type, _exc, _tb) -> None:
        return None

    def read(self, _limit: int = -1) -> bytes:
        return self._raw


def main() -> None:
    product = {
        "categoria": "Informática",
        "subcat": "Periféricos",
        "tipo": "Rato",
        "descricao_normalizada": "Rato USB",
        "fabricante": "",
        "modelo": "",
        "dimensoes": "",
        "resumo": "Periférico apontador com ligação USB.",
        "justificacao": "A descrição identifica diretamente o produto.",
        "recomendacao": "Confirmar fabricante e modelo.",
        "atributos": {
            "material": "",
            "acabamento": "",
            "cor": "",
            "norma": "",
            "referencia": "",
        },
        "confidence": 0.94,
    }
    material = {
        "formato": "Chapa",
        "material": "S235JR",
        "material_familia": "steel",
        "secao_tipo": "",
        "espessura": "15",
        "comprimento": "3000",
        "largura": "1500",
        "altura": "",
        "diametro": "",
        "metros": "",
        "kg_m": "",
        "quantidade": "10",
        "reservado": "0",
        "local": "",
        "lote_fornecedor": "9288X20029",
        "p_compra": "",
        "resumo": "10 chapas S235JR de 15 mm.",
        "missing_fields": [],
        "confidence": 0.98,
    }
    payloads = [
        {"id": "resp_product", "output_text": json.dumps(product, ensure_ascii=False)},
        {"id": "resp_material", "output_text": json.dumps(material, ensure_ascii=False)},
    ]
    requests = []

    def fake_urlopen(request, timeout):
        requests.append((request, timeout, json.loads(request.data.decode("utf-8"))))
        return _FakeResponse(payloads.pop(0))

    client = RemoteProductAIClient(
        openai_api_key="test-only",
        openai_model="gpt-5.6-sol",
        timeout_seconds=12,
    )
    with patch(
        "lugest_core.intelligence.remote_product_ai.urllib.request.urlopen",
        side_effect=fake_urlopen,
    ):
        product_result = client.lookup(
            "Rato USB",
            taxonomy={"categories": [{"label": "Informática"}]},
        )
        material_result = client.material_stock_command(
            "Cria stock de chapa S235JR 15 mm 3000x1500, 10 unidades, lote 9288X20029",
            presets={"formatos": ["Chapa"]},
        )

    assert product_result["candidate"]["categoria"] == "Informática"
    assert product_result["engine"] == "openai-gpt-5.6-sol"
    assert material_result["candidate"]["quantidade"] == "10"
    assert material_result["candidate"]["lote_fornecedor"] == "9288X20029"
    assert material_result["engine"] == "openai-gpt-5.6-sol"
    assert len(requests) == 2
    for request, timeout, body in requests:
        assert request.full_url == "https://api.openai.com/v1/responses"
        assert request.headers["Authorization"] == "Bearer test-only"
        assert timeout == 12
        assert body["model"] == "gpt-5.6-sol"
        assert body["text"]["format"]["type"] == "json_schema"
        assert body["text"]["format"]["strict"] is True
    print("FOCUSED_AI_OK")


if __name__ == "__main__":
    main()
