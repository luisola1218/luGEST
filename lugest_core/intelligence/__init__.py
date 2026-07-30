"""Local, auditable intelligence helpers used by the LuGEST copilots."""

from .product_copilot import (
    extract_product_attributes,
    normalize_product_description,
    product_similarity,
)
from .remote_product_ai import (
    ProductAILookupError,
    RemoteProductAIClient,
    sanitize_gemini_interaction_response,
    sanitize_product_ai_response,
)
from .erp_copilot import (
    ERPCopilot,
    MODULE_KNOWLEDGE,
    build_erp_snapshot,
    contextual_question,
    deterministic_answer,
)

__all__ = [
    "extract_product_attributes",
    "normalize_product_description",
    "product_similarity",
    "ProductAILookupError",
    "RemoteProductAIClient",
    "sanitize_gemini_interaction_response",
    "sanitize_product_ai_response",
    "ERPCopilot",
    "MODULE_KNOWLEDGE",
    "build_erp_snapshot",
    "contextual_question",
    "deterministic_answer",
]
