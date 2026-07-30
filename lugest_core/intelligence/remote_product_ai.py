from __future__ import annotations

import json
import os
import urllib.error
import urllib.parse
import urllib.request
from typing import Any


class ProductAILookupError(RuntimeError):
    """Raised when the optional LuGEST product intelligence service fails."""


def _clean_text(value: Any, limit: int = 500) -> str:
    return " ".join(str(value or "").split())[:limit]


def _valid_source_rows(value: Any) -> list[dict[str, str]]:
    sources: list[dict[str, str]] = []
    for raw in list(value or [])[:8]:
        if not isinstance(raw, dict):
            continue
        url = _clean_text(raw.get("url"), 1000)
        parsed = urllib.parse.urlparse(url)
        if parsed.scheme not in {"http", "https"} or not parsed.netloc:
            continue
        sources.append(
            {
                "title": _clean_text(raw.get("title") or parsed.netloc, 180),
                "url": url,
            }
        )
    return sources


def sanitize_product_ai_response(payload: Any) -> dict[str, Any]:
    """Validate the small, auditable response contract used by the desktop."""

    if not isinstance(payload, dict):
        raise ProductAILookupError("A resposta do serviço de IA não é válida.")
    raw_candidate = payload.get("candidate", payload)
    if not isinstance(raw_candidate, dict):
        raise ProductAILookupError("A IA não devolveu uma ficha de produto válida.")
    candidate = {
        "categoria": _clean_text(raw_candidate.get("categoria") or raw_candidate.get("category"), 100),
        "subcat": _clean_text(raw_candidate.get("subcat") or raw_candidate.get("subcategory"), 100),
        "tipo": _clean_text(raw_candidate.get("tipo") or raw_candidate.get("type"), 120),
        "descricao_normalizada": _clean_text(
            raw_candidate.get("descricao_normalizada") or raw_candidate.get("normalized_description"),
            300,
        ),
        "fabricante": _clean_text(raw_candidate.get("fabricante") or raw_candidate.get("manufacturer"), 120),
        "modelo": _clean_text(raw_candidate.get("modelo") or raw_candidate.get("model"), 160),
        "dimensoes": _clean_text(raw_candidate.get("dimensoes") or raw_candidate.get("dimensions"), 120),
        "resumo": _clean_text(raw_candidate.get("resumo") or raw_candidate.get("summary"), 600),
        "justificacao": _clean_text(
            raw_candidate.get("justificacao") or raw_candidate.get("reasoning"),
            900,
        ),
        "recomendacao": _clean_text(
            raw_candidate.get("recomendacao") or raw_candidate.get("recommendation"),
            500,
        ),
    }
    raw_attributes = raw_candidate.get("atributos") or raw_candidate.get("attributes") or {}
    candidate["atributos"] = {
        _clean_text(key, 60): _clean_text(value, 180)
        for key, value in dict(raw_attributes).items()
        if _clean_text(key, 60) and _clean_text(value, 180)
    } if isinstance(raw_attributes, dict) else {}
    try:
        confidence = float(raw_candidate.get("confidence", payload.get("confidence", 0)) or 0)
    except (TypeError, ValueError):
        confidence = 0.0
    if confidence > 1:
        confidence /= 100.0
    candidate["confidence"] = max(0.0, min(1.0, confidence))
    sources = _valid_source_rows(payload.get("sources") or raw_candidate.get("sources"))
    if not candidate["descricao_normalizada"] and not candidate["modelo"]:
        raise ProductAILookupError("A IA não conseguiu identificar o produto com segurança.")
    return {
        "candidate": candidate,
        "sources": sources,
        "request_id": _clean_text(payload.get("request_id"), 100),
        "engine": _clean_text(payload.get("engine") or "remote-ai", 80),
    }


_OLLAMA_PRODUCT_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "categoria": {"type": "string"},
        "subcat": {"type": "string"},
        "tipo": {"type": "string"},
        "descricao_normalizada": {"type": "string"},
        "fabricante": {"type": "string"},
        "modelo": {"type": "string"},
        "dimensoes": {"type": "string"},
        "resumo": {"type": "string"},
        "justificacao": {"type": "string"},
        "recomendacao": {"type": "string"},
        "atributos": {
            "type": "object",
            "additionalProperties": {"type": "string"},
        },
        "confidence": {"type": "number", "minimum": 0, "maximum": 1},
    },
    "required": [
        "categoria",
        "subcat",
        "tipo",
        "descricao_normalizada",
        "fabricante",
        "modelo",
        "dimensoes",
        "resumo",
        "justificacao",
        "recomendacao",
        "atributos",
        "confidence",
    ],
    "additionalProperties": False,
}

_OLLAMA_MATERIAL_STOCK_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "formato": {"type": "string"},
        "material": {"type": "string"},
        "material_familia": {"type": "string"},
        "secao_tipo": {"type": "string"},
        "espessura": {"type": "string"},
        "comprimento": {"type": "string"},
        "largura": {"type": "string"},
        "altura": {"type": "string"},
        "diametro": {"type": "string"},
        "metros": {"type": "string"},
        "kg_m": {"type": "string"},
        "quantidade": {"type": "string"},
        "reservado": {"type": "string"},
        "local": {"type": "string"},
        "lote_fornecedor": {"type": "string"},
        "p_compra": {"type": "string"},
        "resumo": {"type": "string"},
        "missing_fields": {
            "type": "array",
            "items": {"type": "string"},
        },
        "confidence": {"type": "number", "minimum": 0, "maximum": 1},
    },
    "required": [
        "formato",
        "material",
        "material_familia",
        "secao_tipo",
        "espessura",
        "comprimento",
        "largura",
        "altura",
        "diametro",
        "metros",
        "kg_m",
        "quantidade",
        "reservado",
        "local",
        "lote_fornecedor",
        "p_compra",
        "resumo",
        "missing_fields",
        "confidence",
    ],
    "additionalProperties": False,
}

_OPENAI_PRODUCT_SCHEMA: dict[str, Any] = json.loads(json.dumps(_OLLAMA_PRODUCT_SCHEMA))
_OPENAI_PRODUCT_SCHEMA["properties"]["atributos"] = {
    "type": "object",
    "properties": {
        "material": {"type": "string"},
        "acabamento": {"type": "string"},
        "cor": {"type": "string"},
        "norma": {"type": "string"},
        "referencia": {"type": "string"},
    },
    "required": ["material", "acabamento", "cor", "norma", "referencia"],
    "additionalProperties": False,
}

def _sanitize_material_candidate(candidate: Any, *, engine: str) -> dict[str, Any]:
    if not isinstance(candidate, dict):
        raise ProductAILookupError("A IA não devolveu uma ficha de material válida.")
    clean_candidate = {
        key: _clean_text(candidate.get(key), 300)
        for key in (
            "formato",
            "material",
            "material_familia",
            "secao_tipo",
            "espessura",
            "comprimento",
            "largura",
            "altura",
            "diametro",
            "metros",
            "kg_m",
            "quantidade",
            "reservado",
            "local",
            "lote_fornecedor",
            "p_compra",
            "resumo",
        )
    }
    clean_candidate["missing_fields"] = [
        _clean_text(item, 80)
        for item in list(candidate.get("missing_fields", []) or [])[:12]
        if _clean_text(item, 80)
    ]
    try:
        confidence = float(candidate.get("confidence", 0) or 0)
    except (TypeError, ValueError):
        confidence = 0.0
    clean_candidate["confidence"] = max(0.0, min(1.0, confidence))
    return {"candidate": clean_candidate, "engine": engine}


def _json_object_from_text(value: Any) -> dict[str, Any]:
    text = str(value or "").strip()
    if text.startswith("```"):
        text = text.split("\n", 1)[-1]
        text = text.rsplit("```", 1)[0].strip()
    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        start = text.find("{")
        end = text.rfind("}")
        if start < 0 or end <= start:
            raise ProductAILookupError("A Google IA não devolveu uma ficha estruturada.")
        try:
            parsed = json.loads(text[start : end + 1])
        except json.JSONDecodeError as exc:
            raise ProductAILookupError("A Google IA devolveu uma ficha ilegível.") from exc
    if not isinstance(parsed, dict):
        raise ProductAILookupError("A Google IA não devolveu um objeto de produto.")
    return parsed


def sanitize_gemini_interaction_response(payload: Any) -> dict[str, Any]:
    """Convert a grounded Gemini Interactions response into our local contract."""

    if not isinstance(payload, dict):
        raise ProductAILookupError("A resposta da Google IA não é válida.")
    output_text = _clean_text(payload.get("output_text"), 20_000)
    sources: list[dict[str, str]] = []
    for step in list(payload.get("steps", []) or []):
        if not isinstance(step, dict) or step.get("type") != "model_output":
            continue
        for block in list(step.get("content", []) or []):
            if not isinstance(block, dict) or block.get("type") != "text":
                continue
            if not output_text:
                output_text = str(block.get("text", "") or "").strip()
            for annotation in list(block.get("annotations", []) or []):
                if not isinstance(annotation, dict) or annotation.get("type") != "url_citation":
                    continue
                sources.append(
                    {
                        "title": _clean_text(annotation.get("title"), 180),
                        "url": _clean_text(annotation.get("url"), 1000),
                    }
                )
    candidate = _json_object_from_text(output_text)
    return sanitize_product_ai_response(
        {
            "candidate": candidate,
            "sources": _valid_source_rows(sources),
            "request_id": payload.get("id", ""),
            "engine": "google-gemini-search",
        }
    )


class RemoteProductAIClient:
    """Client for a LuGEST-owned gateway.

    The OpenAI API key belongs on that gateway, never in the distributed
    desktop executable or its .env file.
    """

    def __init__(
        self,
        endpoint: str = "",
        access_token: str = "",
        gemini_api_key: str = "",
        gemini_model: str = "",
        timeout_seconds: float = 25.0,
        ollama_url: str = "",
        ollama_model: str = "",
        openai_api_key: str = "",
        openai_model: str = "",
        openai_base_url: str = "",
    ) -> None:
        self.endpoint = _clean_text(endpoint or os.getenv("LUGEST_AI_ENDPOINT"), 1000)
        self.access_token = _clean_text(
            access_token or os.getenv("LUGEST_AI_ACCESS_TOKEN"),
            1000,
        )
        self.gemini_api_key = _clean_text(
            gemini_api_key
            or os.getenv("LUGEST_GEMINI_API_KEY")
            or os.getenv("GEMINI_API_KEY"),
            1000,
        )
        self.gemini_model = _clean_text(
            gemini_model or os.getenv("LUGEST_GEMINI_MODEL") or "gemini-3.6-flash",
            100,
        )
        self.ollama_url = _clean_text(
            ollama_url
            or os.getenv("LUGEST_OLLAMA_URL")
            or "http://127.0.0.1:11434",
            1000,
        ).rstrip("/")
        self.ollama_model = _clean_text(
            ollama_model or os.getenv("LUGEST_OLLAMA_MODEL") or "qwen3:4b",
            100,
        )
        self.openai_api_key = _clean_text(
            openai_api_key
            or os.getenv("LUGEST_OPENAI_API_KEY")
            or os.getenv("OPENAI_API_KEY"),
            1000,
        )
        self.openai_model = _clean_text(
            openai_model or os.getenv("LUGEST_OPENAI_MODEL") or "gpt-5.6-sol",
            100,
        )
        self.openai_base_url = _clean_text(
            openai_base_url
            or os.getenv("LUGEST_OPENAI_BASE_URL")
            or "https://api.openai.com/v1",
            1000,
        ).rstrip("/")
        self.timeout_seconds = max(3.0, min(float(timeout_seconds), 180.0))

    @property
    def configured(self) -> bool:
        parsed = urllib.parse.urlparse(self.endpoint)
        ollama = urllib.parse.urlparse(self.ollama_url)
        return (
            (parsed.scheme == "https" and bool(parsed.netloc))
            or (ollama.scheme in {"http", "https"} and bool(ollama.netloc))
            or bool(self.gemini_api_key)
            or bool(self.openai_api_key)
        )

    def lookup(
        self,
        description: str,
        *,
        local_candidate: dict[str, Any] | None = None,
        taxonomy: dict[str, Any] | None = None,
        locale: str = "pt-PT",
        user_instruction: str = "",
        previous_candidate: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        description = _clean_text(description, 500)
        if not description:
            raise ProductAILookupError("Escreve primeiro a descrição do produto.")
        if not self.configured:
            raise ProductAILookupError(
                "A inteligência de produtos ainda não está configurada. "
                "Instala o Ollama com o modelo qwen3:4b ou define o serviço de IA LuGEST."
            )
        if self.endpoint:
            parsed_endpoint = urllib.parse.urlparse(self.endpoint)
            if parsed_endpoint.scheme != "https" or not parsed_endpoint.netloc:
                raise ProductAILookupError(
                    "LUGEST_AI_ENDPOINT tem de ser um endereço HTTPS válido."
                )
            return self._lookup_gateway(
                description,
                local_candidate=local_candidate,
                taxonomy=taxonomy,
                locale=locale,
                user_instruction=user_instruction,
                previous_candidate=previous_candidate,
            )
        provider_error: ProductAILookupError | None = None
        if self.openai_api_key:
            try:
                return self._lookup_openai(
                    description,
                    local_candidate=local_candidate,
                    taxonomy=taxonomy,
                    locale=locale,
                    user_instruction=user_instruction,
                    previous_candidate=previous_candidate,
                )
            except ProductAILookupError as exc:
                provider_error = exc
        if self.gemini_api_key:
            try:
                result = self._lookup_gemini(
                    description,
                    local_candidate=local_candidate,
                    taxonomy=taxonomy,
                    locale=locale,
                    user_instruction=user_instruction,
                    previous_candidate=previous_candidate,
                )
                if provider_error is not None:
                    result["fallback_reason"] = "openai_unavailable"
                return result
            except ProductAILookupError as exc:
                provider_error = provider_error or exc
        if self.ollama_url:
            try:
                result = self._lookup_ollama(
                    description,
                    local_candidate=local_candidate,
                    taxonomy=taxonomy,
                    locale=locale,
                    user_instruction=user_instruction,
                    previous_candidate=previous_candidate,
                )
                if provider_error is not None:
                    result["fallback_reason"] = "cloud_unavailable"
                return result
            except ProductAILookupError as exc:
                provider_error = provider_error or exc
        if provider_error is not None:
            raise provider_error
        raise ProductAILookupError("Não existe nenhum motor de IA configurado.")

    def _material_gateway(
        self,
        command: str,
        *,
        presets: dict[str, Any] | None,
        locale: str,
    ) -> dict[str, Any]:
        parsed_endpoint = urllib.parse.urlparse(self.endpoint)
        if parsed_endpoint.scheme != "https" or not parsed_endpoint.netloc:
            raise ProductAILookupError(
                "LUGEST_AI_ENDPOINT tem de ser um endereço HTTPS válido."
            )
        body = json.dumps(
            {
                "task": "material_stock_draft",
                "command": command,
                "locale": locale,
                "presets": dict(presets or {}),
                "requirements": {
                    "do_not_persist": True,
                    "do_not_invent": True,
                    "return_structured_candidate": True,
                },
            },
            ensure_ascii=False,
        ).encode("utf-8")
        headers = {
            "Accept": "application/json",
            "Content-Type": "application/json; charset=utf-8",
            "User-Agent": "LuGEST-Desktop/MaterialFocusedAI",
        }
        if self.access_token:
            headers["Authorization"] = f"Bearer {self.access_token}"
        request = urllib.request.Request(self.endpoint, data=body, headers=headers, method="POST")
        try:
            with urllib.request.urlopen(request, timeout=self.timeout_seconds) as response:
                raw = response.read(1_000_001)
            decoded = json.loads(raw.decode("utf-8"))
        except urllib.error.HTTPError as exc:
            raise ProductAILookupError(
                f"O serviço IA LuGEST respondeu com erro {exc.code}."
            ) from exc
        except urllib.error.URLError as exc:
            raise ProductAILookupError(
                "Não foi possível contactar o serviço IA LuGEST."
            ) from exc
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise ProductAILookupError(
                "O serviço IA LuGEST devolveu uma resposta ilegível."
            ) from exc
        return _sanitize_material_candidate(
            decoded.get("candidate", decoded),
            engine=_clean_text(decoded.get("engine") or "lugest-ai-gateway", 80),
        )

    def _openai_structured(
        self,
        *,
        schema_name: str,
        schema: dict[str, Any],
        instructions: str,
        user_input: str,
        max_output_tokens: int,
    ) -> tuple[dict[str, Any], str]:
        parsed_base = urllib.parse.urlparse(self.openai_base_url)
        is_local = parsed_base.hostname in {"127.0.0.1", "localhost"}
        allowed_schemes = {"http", "https"} if is_local else {"https"}
        if parsed_base.scheme not in allowed_schemes or not parsed_base.netloc:
            raise ProductAILookupError(
                "LUGEST_OPENAI_BASE_URL tem de ser um endereço HTTPS válido."
            )
        body = json.dumps(
            {
                "model": self.openai_model,
                "instructions": instructions,
                "input": user_input,
                "reasoning": {"effort": "low"},
                "max_output_tokens": max_output_tokens,
                "text": {
                    "format": {
                        "type": "json_schema",
                        "name": schema_name,
                        "strict": True,
                        "schema": schema,
                    }
                },
            },
            ensure_ascii=False,
        ).encode("utf-8")
        request = urllib.request.Request(
            f"{self.openai_base_url}/responses",
            data=body,
            headers={
                "Accept": "application/json",
                "Content-Type": "application/json; charset=utf-8",
                "Authorization": f"Bearer {self.openai_api_key}",
                "User-Agent": "LuGEST-Desktop/FocusedAI",
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=self.timeout_seconds) as response:
                raw = response.read(1_000_001)
        except urllib.error.HTTPError as exc:
            detail = ""
            try:
                decoded_error = json.loads(exc.read(5000).decode("utf-8", errors="replace"))
                detail = _clean_text(
                    dict(decoded_error.get("error", {}) or {}).get("message"),
                    450,
                )
            except Exception:
                pass
            if exc.code in {401, 403}:
                raise ProductAILookupError(
                    "A credencial OpenAI foi recusada. Confirma o serviço seguro configurado."
                ) from exc
            if exc.code == 429:
                raise ProductAILookupError(
                    "O serviço OpenAI atingiu o limite de utilização ou faturação."
                ) from exc
            suffix = f": {detail}" if detail else ""
            raise ProductAILookupError(
                f"O serviço OpenAI respondeu com erro {exc.code}{suffix}"
            ) from exc
        except urllib.error.URLError as exc:
            raise ProductAILookupError(
                "Não foi possível contactar o serviço OpenAI. Confirma a Internet."
            ) from exc
        if len(raw) > 1_000_000:
            raise ProductAILookupError("A resposta OpenAI excedeu o limite permitido.")
        try:
            decoded = json.loads(raw.decode("utf-8"))
            output_text = str(decoded.get("output_text", "") or "").strip()
            if not output_text:
                for output in list(decoded.get("output", []) or []):
                    for content in list(dict(output or {}).get("content", []) or []):
                        block = dict(content or {})
                        if block.get("type") in {"output_text", "text"} and block.get("text"):
                            output_text = str(block.get("text") or "").strip()
                            break
                    if output_text:
                        break
            if not output_text:
                raise ProductAILookupError(
                    "O modelo avançado não devolveu uma proposta estruturada."
                )
            return _json_object_from_text(output_text), _clean_text(decoded.get("id"), 100)
        except (UnicodeDecodeError, json.JSONDecodeError, AttributeError, TypeError) as exc:
            raise ProductAILookupError(
                "O serviço OpenAI devolveu uma resposta ilegível."
            ) from exc

    def _lookup_openai(
        self,
        description: str,
        *,
        local_candidate: dict[str, Any] | None,
        taxonomy: dict[str, Any] | None,
        locale: str,
        user_instruction: str,
        previous_candidate: dict[str, Any] | None,
    ) -> dict[str, Any]:
        instructions = (
            "És o especialista de catálogo industrial do ERP LuGEST. Identifica produtos com "
            "rigor e devolve exclusivamente a ficha pedida pelo esquema. Usa preferencialmente "
            "a taxonomia fornecida; só propõe uma categoria nova quando não existir encaixe "
            "profissional. Não inventes fabricante, modelo, norma ou dimensões. Campos "
            "desconhecidos ficam vazios e a confiança deve baixar quando houver ambiguidade. "
            "A justificação deve ser breve e verificável, sem raciocínio interno passo a passo."
        )
        user_input = json.dumps(
            {
                "locale": locale,
                "description": description,
                "user_instruction": _clean_text(user_instruction, 1000),
                "local_candidate": dict(local_candidate or {}),
                "previous_candidate": dict(previous_candidate or {}),
                "taxonomy": dict(taxonomy or {}),
            },
            ensure_ascii=False,
        )
        candidate, request_id = self._openai_structured(
            schema_name="lugest_product_candidate",
            schema=_OPENAI_PRODUCT_SCHEMA,
            instructions=instructions,
            user_input=user_input,
            max_output_tokens=1800,
        )
        return sanitize_product_ai_response(
            {
                "candidate": candidate,
                "request_id": request_id,
                "engine": f"openai-{self.openai_model}",
            }
        )

    def material_stock_command(
        self,
        command: str,
        *,
        presets: dict[str, Any] | None = None,
        locale: str = "pt-PT",
    ) -> dict[str, Any]:
        """Convert a natural-language stock command into an uncommitted material draft."""

        command = _clean_text(command, 1000)
        if not command:
            raise ProductAILookupError("Escreve primeiro o material que pretendes criar.")
        if not self.endpoint and not self.ollama_url and not self.openai_api_key:
            raise ProductAILookupError(
                "A criação assistida de matéria-prima necessita do Ollama no posto do cliente."
            )
        prompt = (
            "Atua como assistente de stock industrial do ERP LuGEST. Converte o pedido numa "
            "ficha de matéria-prima, mas não graves nada. Usa apenas valores expressamente "
            "indicados ou inequivocamente dedutíveis. Para chapa, formato='Chapa', comprimento "
            "e largura são milímetros, espessura é milímetros e quantidade é o número de chapas. "
            "Mapeia 'lote externo', 'lote fornecedor' e 'lote do fornecedor' para "
            "lote_fornecedor. material_familia deve ser uma destas chaves quando reconhecível: "
            "steel, stainless, aluminium, copper, brass, plastic; caso contrário fica vazio. "
            "Não confundas espessura com dimensões ou quantidade. Campos desconhecidos ficam "
            "vazios. reservado deve ser zero quando não indicado. missing_fields deve listar "
            "apenas dados obrigatórios em falta para guardar. Responde no esquema fornecido. "
            f"Idioma: {locale}. "
            f"Opções do posto: {json.dumps(presets or {}, ensure_ascii=False)}. "
            f"Pedido do utilizador: {command}"
        )
        cloud_error: ProductAILookupError | None = None
        if self.endpoint:
            try:
                return self._material_gateway(
                    command,
                    presets=presets,
                    locale=locale,
                )
            except ProductAILookupError as exc:
                cloud_error = exc
        if self.openai_api_key:
            try:
                candidate, _request_id = self._openai_structured(
                    schema_name="lugest_material_stock_draft",
                    schema=_OLLAMA_MATERIAL_STOCK_SCHEMA,
                    instructions=(
                        "És o especialista de matéria-prima do ERP industrial LuGEST. Converte "
                        "o pedido numa proposta de stock não gravada. Respeita unidades, séries "
                        "normalizadas e opções do posto; não inventes lotes, quantidades, preços "
                        "ou dimensões. Campos desconhecidos ficam vazios."
                    ),
                    user_input=prompt,
                    max_output_tokens=1600,
                )
                return _sanitize_material_candidate(
                    candidate,
                    engine=f"openai-{self.openai_model}",
                )
            except ProductAILookupError as exc:
                cloud_error = exc
        if not self.ollama_url:
            if cloud_error is not None:
                raise cloud_error
            raise ProductAILookupError("Não existe nenhum motor de IA configurado.")
        body = json.dumps(
            {
                "model": self.ollama_model,
                "prompt": prompt,
                "stream": False,
                "think": False,
                "format": _OLLAMA_MATERIAL_STOCK_SCHEMA,
                "keep_alive": "10m",
                "options": {
                    "temperature": 0.05,
                    "num_ctx": 8192,
                },
            },
            ensure_ascii=False,
        ).encode("utf-8")
        request = urllib.request.Request(
            f"{self.ollama_url}/api/generate",
            data=body,
            headers={
                "Accept": "application/json",
                "Content-Type": "application/json; charset=utf-8",
                "User-Agent": "LuGEST-Desktop/MaterialCopilot-Local",
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=self.timeout_seconds) as response:
                raw = response.read(300_001)
            decoded = json.loads(raw.decode("utf-8"))
            candidate = _json_object_from_text(decoded.get("response"))
        except urllib.error.HTTPError as exc:
            raise ProductAILookupError(f"O Ollama respondeu com erro {exc.code}.") from exc
        except urllib.error.URLError as exc:
            raise ProductAILookupError(
                "A IA local não está acessível. Inicia o Ollama e confirma o modelo configurado."
            ) from exc
        except (UnicodeDecodeError, json.JSONDecodeError, AttributeError) as exc:
            raise ProductAILookupError("O Ollama devolveu uma ficha de material ilegível.") from exc
        result = _sanitize_material_candidate(
            candidate,
            engine=f"ollama-{self.ollama_model}",
        )
        if cloud_error is not None:
            result["fallback_reason"] = "openai_unavailable"
        return result

    def _lookup_gateway(
        self,
        description: str,
        *,
        local_candidate: dict[str, Any] | None,
        taxonomy: dict[str, Any] | None,
        locale: str,
        user_instruction: str,
        previous_candidate: dict[str, Any] | None,
    ) -> dict[str, Any]:
        body = json.dumps(
            {
                "task": "product_identification",
                "description": description,
                "locale": _clean_text(locale, 20) or "pt-PT",
                "local_candidate": dict(local_candidate or {}),
                "taxonomy": dict(taxonomy or {}),
                "user_instruction": _clean_text(user_instruction, 1000),
                "previous_candidate": dict(previous_candidate or {}),
                "requirements": {
                    "use_web_research": True,
                    "require_sources": True,
                    "do_not_invent": True,
                    "return_structured_candidate": True,
                },
            },
            ensure_ascii=False,
        ).encode("utf-8")
        headers = {
            "Accept": "application/json",
            "Content-Type": "application/json; charset=utf-8",
            "User-Agent": "LuGEST-Desktop/ProductCopilot",
        }
        if self.access_token:
            headers["Authorization"] = f"Bearer {self.access_token}"
        request = urllib.request.Request(self.endpoint, data=body, headers=headers, method="POST")
        try:
            with urllib.request.urlopen(request, timeout=self.timeout_seconds) as response:
                raw = response.read(1_000_001)
        except urllib.error.HTTPError as exc:
            detail = ""
            try:
                detail = _clean_text(exc.read(800).decode("utf-8", errors="replace"), 300)
            except Exception:
                pass
            suffix = f": {detail}" if detail else ""
            raise ProductAILookupError(f"O serviço de IA respondeu com erro {exc.code}{suffix}") from exc
        except urllib.error.URLError as exc:
            raise ProductAILookupError(
                "Não foi possível contactar o serviço de IA. Confirma a Internet e tenta novamente."
            ) from exc
        if len(raw) > 1_000_000:
            raise ProductAILookupError("A resposta do serviço de IA excedeu o limite permitido.")
        try:
            decoded = json.loads(raw.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise ProductAILookupError("O serviço de IA devolveu uma resposta ilegível.") from exc
        return sanitize_product_ai_response(decoded)

    def _lookup_ollama(
        self,
        description: str,
        *,
        local_candidate: dict[str, Any] | None,
        taxonomy: dict[str, Any] | None,
        locale: str,
        user_instruction: str,
        previous_candidate: dict[str, Any] | None,
    ) -> dict[str, Any]:
        """Identify a product with a free Ollama model running locally or on the LAN."""

        question = _clean_text(user_instruction, 1000)
        if question:
            conversation_instruction = (
                "PEDIDO PRIORITÁRIO DO UTILIZADOR: "
                f"{question}\n"
                "Responde realmente à pergunta sobre o produto, usando o teu conhecimento geral. "
                "O campo resumo deve conter uma resposta útil em 2 a 5 frases completas. "
                "Não transformes a pergunta numa etiqueta, não respondas apenas 'interpretei X "
                "como X' e não copies o resumo anterior. Se a proposta anterior estiver errada, "
                "corrige também categoria, subcategoria, tipo e confiança. "
            )
        else:
            conversation_instruction = (
                "Não existe uma pergunta adicional; faz a identificação inicial do produto. "
            )
        prompt = (
            f"{conversation_instruction}"
            f"PRODUTO EM ANÁLISE: {description}.\n"
            "Atua como catalogador industrial do ERP LuGEST. Analisa a descrição sem inventar "
            "dados. Extrai fabricante, referência/modelo, material, acabamento, medida e outras "
            "características presentes. Prefere exatamente os nomes existentes na taxonomia "
            "LuGEST. Se a taxonomia não tiver um encaixe correto, propõe nomes profissionais "
            "curtos para categoria, subcategoria e tipo. 'Outros' é apenas um último recurso: "
            "não o uses quando o objeto tem uma família reconhecível. A análise local marcada "
            "como Outros é provisória e não deve condicionar a tua classificação. "
            "Mantém campos desconhecidos vazios e "
            "reduz confidence quando houver ambiguidade. Em resumo explica em linguagem simples "
            "o produto que entendeste. Em justificacao explica brevemente por que escolheste a "
            "classificação, sem revelar raciocínio interno passo a passo. Em recomendacao indica "
            "o que o utilizador deve confirmar antes de guardar. "
            f"Idioma: {locale}. "
            f"Descrição: {description}. "
            f"Análise local existente: {json.dumps(local_candidate or {}, ensure_ascii=False)}. "
            f"Taxonomia disponível: {json.dumps(taxonomy or {}, ensure_ascii=False)}. "
            f"Proposta anterior: {json.dumps(previous_candidate or {}, ensure_ascii=False)}. "
            f"Pergunta ou correção do utilizador: {question or 'Nenhuma'}. "
            "Revê todos os campos da proposta que devam mudar."
        )
        body = json.dumps(
            {
                "model": self.ollama_model,
                "prompt": prompt,
                "stream": False,
                "think": False,
                "format": _OLLAMA_PRODUCT_SCHEMA,
                "keep_alive": "10m",
                "options": {
                    "temperature": 0.2 if question else 0.1,
                    "num_ctx": 16384,
                },
            },
            ensure_ascii=False,
        ).encode("utf-8")
        request = urllib.request.Request(
            f"{self.ollama_url}/api/generate",
            data=body,
            headers={
                "Accept": "application/json",
                "Content-Type": "application/json; charset=utf-8",
                "User-Agent": "LuGEST-Desktop/ProductCopilot-Local",
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=self.timeout_seconds) as response:
                raw = response.read(1_000_001)
        except urllib.error.HTTPError as exc:
            detail = ""
            try:
                decoded_error = json.loads(exc.read(3000).decode("utf-8", errors="replace"))
                detail = _clean_text(decoded_error.get("error"), 350)
            except Exception:
                pass
            if exc.code == 404 or "not found" in detail.casefold():
                raise ProductAILookupError(
                    f"O modelo local '{self.ollama_model}' ainda não está instalado. "
                    f"Executa: ollama pull {self.ollama_model}"
                ) from exc
            suffix = f": {detail}" if detail else ""
            raise ProductAILookupError(f"O Ollama respondeu com erro {exc.code}{suffix}") from exc
        except urllib.error.URLError as exc:
            raise ProductAILookupError(
                "A IA local gratuita não está acessível. Inicia o Ollama e confirma "
                f"LUGEST_OLLAMA_URL={self.ollama_url}. Depois instala o modelo com "
                f"'ollama pull {self.ollama_model}'."
            ) from exc
        if len(raw) > 1_000_000:
            raise ProductAILookupError("A resposta do Ollama excedeu o limite permitido.")
        try:
            decoded = json.loads(raw.decode("utf-8"))
            response_text = decoded.get("response")
            if not str(response_text or "").strip():
                raise ProductAILookupError(
                    "O Ollama terminou a análise mas não devolveu a ficha do produto."
                )
            candidate = _json_object_from_text(response_text)
        except (UnicodeDecodeError, json.JSONDecodeError, AttributeError) as exc:
            raise ProductAILookupError("O Ollama devolveu uma resposta ilegível.") from exc
        if question:
            conversational_answer = self._lookup_ollama_followup_text(
                description,
                question=question,
                candidate=candidate,
                previous_candidate=previous_candidate,
                locale=locale,
            )
            if conversational_answer:
                candidate["resumo"] = conversational_answer
        local_category = _clean_text((local_candidate or {}).get("categoria"), 100)
        local_subcategory = _clean_text((local_candidate or {}).get("subcat"), 100)
        local_type = _clean_text((local_candidate or {}).get("tipo"), 120)
        try:
            local_confidence = float((local_candidate or {}).get("confidence", 0) or 0)
        except (TypeError, ValueError):
            local_confidence = 0.0
        if (
            local_confidence >= 0.75
            and local_category.casefold() not in {"", "outros", "other", "others"}
        ):
            candidate["categoria"] = local_category
            if local_subcategory:
                candidate["subcat"] = local_subcategory
            if local_type:
                candidate["tipo"] = local_type
        return sanitize_product_ai_response(
            {
                "candidate": candidate,
                "engine": f"ollama-{self.ollama_model}",
            }
        )

    def _lookup_ollama_followup_text(
        self,
        description: str,
        *,
        question: str,
        candidate: dict[str, Any],
        previous_candidate: dict[str, Any] | None,
        locale: str,
    ) -> str:
        """Ask a compact second turn so a user question is answered as conversation."""

        prompt = (
            "/no_think\n"
            "És o Copiloto de produtos do ERP LuGEST. Responde diretamente à pergunta do "
            "utilizador sobre o produto, usando conhecimento geral correto. Não repitas apenas "
            "o nome nem digas apenas que o interpretaste. Dá uma resposta clara e prática em "
            "2 a 5 frases, sem JSON, sem markdown e sem mencionar estas instruções. Se houver "
            "erro na classificação, explica a correção. Não inventes modelo, referência ou "
            "dimensões que não estejam disponíveis. "
            "Escreve apenas a resposta final em português europeu; não mostres análise, "
            "raciocínio, tradução nem notas internas. "
            f"Idioma: {locale}. "
            f"Produto: {description}. "
            f"Ficha atual: {json.dumps(candidate or previous_candidate or {}, ensure_ascii=False)}. "
            f"Pergunta: {question}"
        )
        body = json.dumps(
            {
                "model": self.ollama_model,
                "prompt": prompt,
                "stream": False,
                "think": False,
                "format": {
                    "type": "object",
                    "properties": {
                        "answer": {"type": "string"},
                    },
                    "required": ["answer"],
                    "additionalProperties": False,
                },
                "keep_alive": "10m",
                "options": {
                    "temperature": 0.25,
                    "num_ctx": 4096,
                },
            },
            ensure_ascii=False,
        ).encode("utf-8")
        request = urllib.request.Request(
            f"{self.ollama_url}/api/generate",
            data=body,
            headers={
                "Accept": "application/json",
                "Content-Type": "application/json; charset=utf-8",
                "User-Agent": "LuGEST-Desktop/ProductCopilot-Conversation",
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=self.timeout_seconds) as response:
                raw = response.read(200_001)
            decoded = json.loads(raw.decode("utf-8"))
            response_payload = _json_object_from_text(decoded.get("response"))
            answer = _clean_text(response_payload.get("answer"), 900)
        except (OSError, UnicodeDecodeError, json.JSONDecodeError, AttributeError):
            return ""
        normalized_answer = answer.casefold().strip(" .:;!?")
        normalized_description = _clean_text(description, 300).casefold().strip(" .:;!?")
        if not answer or normalized_answer == normalized_description:
            return ""
        if normalized_answer.startswith("interpretei") and len(answer.split()) < 12:
            return ""
        return answer

    def _lookup_gemini(
        self,
        description: str,
        *,
        local_candidate: dict[str, Any] | None,
        taxonomy: dict[str, Any] | None,
        locale: str,
        user_instruction: str,
        previous_candidate: dict[str, Any] | None,
    ) -> dict[str, Any]:
        """Developer/test path using Gemini grounded with Google Search."""

        prompt = (
            "Identifica com rigor este produto industrial usando pesquisa Google. "
            "Procura primeiro pelo fabricante e pela referência exata. Não inventes. "
            "Responde apenas com um único objeto JSON, sem markdown, com estas chaves: "
            "categoria, subcat, tipo, descricao_normalizada, fabricante, modelo, dimensoes, "
            "atributos (objeto) e confidence (número entre 0 e 1). "
            "Usa preferencialmente uma classificação existente na taxonomia fornecida. "
            f"Idioma: {locale}. "
            f"Descrição recebida: {description}. "
            f"Candidato local: {json.dumps(local_candidate or {}, ensure_ascii=False)}. "
            f"Taxonomia LuGEST: {json.dumps(taxonomy or {}, ensure_ascii=False)}. "
            f"Proposta anterior: {json.dumps(previous_candidate or {}, ensure_ascii=False)}. "
            f"Pergunta/correção do utilizador: {_clean_text(user_instruction, 1000) or 'Nenhuma'}."
        )
        body = json.dumps(
            {
                "model": self.gemini_model,
                "input": prompt,
                "tools": [{"type": "google_search"}],
            },
            ensure_ascii=False,
        ).encode("utf-8")
        request = urllib.request.Request(
            "https://generativelanguage.googleapis.com/v1beta/interactions",
            data=body,
            headers={
                "Accept": "application/json",
                "Content-Type": "application/json; charset=utf-8",
                "x-goog-api-key": self.gemini_api_key,
                "User-Agent": "LuGEST-Desktop/ProductCopilot-Dev",
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=self.timeout_seconds) as response:
                raw = response.read(1_000_001)
        except urllib.error.HTTPError as exc:
            detail = ""
            try:
                decoded_error = json.loads(exc.read(3000).decode("utf-8", errors="replace"))
                detail = _clean_text(
                    dict(decoded_error.get("error", {}) or {}).get("message"),
                    350,
                )
            except Exception:
                pass
            if exc.code == 429:
                try:
                    result = self._lookup_gemini_without_search(prompt)
                    result["fallback_reason"] = "google_search_quota"
                    return result
                except ProductAILookupError as fallback_exc:
                    raise ProductAILookupError(
                        "A Google IA está ligada, mas este projeto não tem quota disponível. "
                        "Ativa a faturação/quota no Google AI Studio ou aguarda pela reposição do limite. "
                        "O reconhecimento local do LuGEST continua disponível."
                    ) from fallback_exc
            if exc.code in {401, 403}:
                raise ProductAILookupError(
                    "A chave Google foi recusada. Confirma se pertence ao projeto correto "
                    "e se a Gemini API está autorizada."
                ) from exc
            suffix = f": {detail}" if detail else ""
            raise ProductAILookupError(f"A Google IA respondeu com erro {exc.code}{suffix}") from exc
        except urllib.error.URLError as exc:
            raise ProductAILookupError(
                "Não foi possível contactar a Google IA. Confirma a Internet e tenta novamente."
            ) from exc
        if len(raw) > 1_000_000:
            raise ProductAILookupError("A resposta da Google IA excedeu o limite permitido.")
        try:
            decoded = json.loads(raw.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise ProductAILookupError("A Google IA devolveu uma resposta ilegível.") from exc
        return sanitize_gemini_interaction_response(decoded)

    def _lookup_gemini_without_search(self, prompt: str) -> dict[str, Any]:
        """Retry with model knowledge when Google Search quota is unavailable."""

        body = json.dumps(
            {
                "model": self.gemini_model,
                "input": (
                    prompt
                    + " A pesquisa Google não está disponível neste pedido. "
                    "Usa apenas conhecimento próprio e reduz a confiança quando não tiveres certeza."
                ),
            },
            ensure_ascii=False,
        ).encode("utf-8")
        request = urllib.request.Request(
            "https://generativelanguage.googleapis.com/v1beta/interactions",
            data=body,
            headers={
                "Accept": "application/json",
                "Content-Type": "application/json; charset=utf-8",
                "x-goog-api-key": self.gemini_api_key,
                "User-Agent": "LuGEST-Desktop/ProductCopilot-Dev",
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=self.timeout_seconds) as response:
                raw = response.read(1_000_001)
        except urllib.error.HTTPError as exc:
            if exc.code == 429:
                raise ProductAILookupError("Quota Gemini indisponível.") from exc
            if exc.code in {401, 403}:
                raise ProductAILookupError("Chave Gemini recusada.") from exc
            raise ProductAILookupError(f"A Google IA respondeu com erro {exc.code}.") from exc
        except urllib.error.URLError as exc:
            raise ProductAILookupError("Não foi possível contactar a Google IA.") from exc
        try:
            decoded = json.loads(raw.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise ProductAILookupError("A Google IA devolveu uma resposta ilegível.") from exc
        result = sanitize_gemini_interaction_response(decoded)
        result["engine"] = "google-gemini-no-search"
        return result
