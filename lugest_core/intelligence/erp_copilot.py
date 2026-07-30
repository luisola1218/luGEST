from __future__ import annotations

import json
import os
import re
import unicodedata
import urllib.parse
import urllib.request
from datetime import date, datetime
from typing import Any


_COPILOT_ROUTE_SCHEMA: dict[str, Any] = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "intent": {
            "type": "string",
            "enum": [
                "weather",
                "erp_query",
                "navigate",
                "prepare_action",
                "general_chat",
            ],
        },
        "module": {"type": "string"},
        "topic": {"type": "string"},
        "location": {"type": "string"},
        "date_reference": {
            "type": "string",
            "enum": ["today", "tomorrow", "unspecified"],
        },
        "action_type": {"type": "string"},
        "command": {"type": "string"},
    },
    "required": [
        "intent",
        "module",
        "topic",
        "location",
        "date_reference",
        "action_type",
        "command",
    ],
}


MODULE_KNOWLEDGE: dict[str, dict[str, Any]] = {
    "stock_dashboard": {
        "nome": "Visão Executiva",
        "descricao": "Património, compras, vendas, faturação e leitura executiva da empresa.",
        "fluxos": ["consultar património", "comparar stock", "abrir encomendas e transportes", "exportar indicadores"],
    },
    "pulse": {
        "nome": "Pulse",
        "descricao": "OEE, disponibilidade, perdas, desvios e desempenho industrial.",
        "fluxos": ["analisar desempenho", "consultar atrasos", "gerar relatório"],
    },
    "materials": {
        "nome": "Matéria-Prima",
        "descricao": "Lotes, formatos, disponibilidade, reservas, valorização e rastreabilidade.",
        "fluxos": ["criar lote", "corrigir stock", "dar baixa", "calcular peso", "consultar histórico"],
    },
    "products": {
        "nome": "Produtos",
        "descricao": "Catálogo de produto acabado, stock, preços, classificação e movimentos.",
        "fluxos": ["criar produto", "classificar produto", "alterar stock e preços", "consultar movimentos"],
    },
    "customers": {
        "nome": "Clientes",
        "descricao": "Cadastro comercial, contactos, moradas, localização e histórico do cliente.",
        "fluxos": ["criar cliente", "editar dados", "consultar localização e histórico"],
    },
    "suppliers": {
        "nome": "Fornecedores",
        "descricao": "Cadastro, contactos, moradas, condições e histórico de fornecedores.",
        "fluxos": ["criar fornecedor", "editar dados", "consultar compras e localização"],
    },
    "orders": {
        "nome": "Encomendas",
        "descricao": "Ordens de fabrico, materiais, peças, operações, montagem e progresso.",
        "fluxos": ["criar ordem", "abrir ordem", "editar OF", "acompanhar progresso"],
    },
    "quotes": {
        "nome": "Orçamentos",
        "descricao": "Propostas, DXF/DWG, cálculo, nesting, condições e conversão em encomenda.",
        "fluxos": ["criar proposta", "orçamentar DXF", "calcular nesting", "aprovar ou rejeitar", "converter em encomenda"],
    },
    "planning": {
        "nome": "Planeamento",
        "descricao": "Carga semanal por máquina, blocos, prazos e sequência produtiva.",
        "fluxos": ["planear operação", "mover bloco", "bloquear período", "consultar carga"],
    },
    "transportes": {
        "nome": "Transportes",
        "descricao": "Viagens, destinos, rotas, custos, transportadoras e confirmação de entregas.",
        "fluxos": ["criar viagem", "adicionar destino", "abrir rota", "confirmar entrega"],
    },
    "material_assistant": {
        "nome": "Assistente MP",
        "descricao": "Decisões de separação e cativação de matéria-prima ligadas ao stock e planeamento.",
        "fluxos": ["validar recomendação", "consultar stock", "abrir plano de separação"],
    },
    "operator": {
        "nome": "Operador",
        "descricao": "Execução das operações, tempos, consumos, baixas e registo de avarias.",
        "fluxos": ["iniciar operação", "interromper", "finalizar", "consumir componentes", "registar avaria"],
    },
    "opp": {
        "nome": "OPP",
        "descricao": "Consulta e preparação das ordens de processo e produção.",
        "fluxos": ["localizar OPP", "abrir encomenda", "consultar operação"],
    },
    "shipping": {
        "nome": "Expedição",
        "descricao": "Preparação, documentação e saída de encomendas.",
        "fluxos": ["preparar expedição", "validar volumes", "emitir documentação"],
    },
    "billing": {
        "nome": "Faturação",
        "descricao": "Documentos comerciais, valores faturados, recebimentos e saldos.",
        "fluxos": ["emitir documento", "consultar faturação", "registar recebimento"],
    },
    "purchase_notes": {
        "nome": "Notas de Encomenda",
        "descricao": "Aprovisionamento, pedidos de orçamento, fornecedores, receções, guias e faturas.",
        "fluxos": ["criar pedido", "pedir orçamento", "aprovar", "enviar NE", "registar guia ou fatura"],
    },
    "quality": {
        "nome": "Qualidade",
        "descricao": "Inspeções, não conformidades, decisões e rastreabilidade da qualidade.",
        "fluxos": ["registar inspeção", "abrir não conformidade", "acompanhar decisão"],
    },
    "diagnostics": {
        "nome": "Diagnóstico",
        "descricao": "Verificações técnicas e operacionais do sistema.",
        "fluxos": ["executar diagnóstico", "consultar alertas"],
    },
    "avarias": {
        "nome": "Avarias",
        "descricao": "Registo, acompanhamento e encerramento de avarias e paragens.",
        "fluxos": ["registar avaria", "acompanhar paragem", "encerrar avaria"],
    },
}


def _text(value: Any, limit: int = 240) -> str:
    return " ".join(str(value or "").split())[:limit]


def _number(value: Any) -> float:
    try:
        return float(str(value or "0").replace(" ", "").replace(",", "."))
    except (TypeError, ValueError):
        return 0.0


def _normalized(value: Any) -> str:
    raw = unicodedata.normalize("NFKD", str(value or ""))
    raw = "".join(char for char in raw if not unicodedata.combining(char))
    raw = re.sub(r"[^a-zA-Z0-9]+", " ", raw.casefold())
    return re.sub(r"\s+", " ", raw).strip()


def _date_value(value: Any) -> date | None:
    raw = _text(value, 32)
    if not raw:
        return None
    try:
        return datetime.fromisoformat(raw[:10]).date()
    except (TypeError, ValueError):
        return None


def _allowed(permissions: dict[str, Any] | None, key: str) -> bool:
    return not permissions or key not in permissions or bool(permissions.get(key))


def _state(row: dict[str, Any]) -> str:
    return _normalized(
        row.get("estado") or row.get("status") or row.get("situacao") or row.get("fase")
    )


def _sample(
    rows: list[dict[str, Any]],
    *,
    identifiers: tuple[str, ...],
    description_fields: tuple[str, ...] = (),
    limit: int = 6,
) -> list[dict[str, str]]:
    output: list[dict[str, str]] = []
    for row in rows[:limit]:
        identifier = next((_text(row.get(key), 80) for key in identifiers if _text(row.get(key), 80)), "")
        description = next(
            (_text(row.get(key), 120) for key in description_fields if _text(row.get(key), 120)),
            "",
        )
        output.append(
            {
                "id": identifier or "-",
                "descricao": description,
                "estado": _text(row.get("estado") or row.get("status") or row.get("situacao"), 50),
            }
        )
    return output


def build_erp_snapshot(
    data: dict[str, Any] | None,
    *,
    permissions: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Create a bounded, privacy-conscious and read-only ERP context."""

    source = dict(data or {})
    today = date.today()
    snapshot: dict[str, Any] = {
        "gerado_em": datetime.now().isoformat(timespec="minutes"),
        "modo": "somente_leitura",
    }
    if _allowed(permissions, "materials"):
        rows = [dict(row or {}) for row in list(source.get("materiais", []) or [])]
        critical = [
            row for row in rows
            if _number(row.get("disponivel", row.get("quantidade", row.get("qtd", 0)))) <= 0
            or any(token in _state(row) for token in ("critico", "baixo", "sem stock"))
        ]
        snapshot["materia_prima"] = {
            "registos": len(rows),
            "quantidade_total": round(
                sum(_number(row.get("quantidade", row.get("qtd", 0))) for row in rows), 3
            ),
            "criticos": len(critical),
            "amostra_critica": _sample(
                critical,
                identifiers=("id", "codigo", "lote_interno"),
                description_fields=("material", "descricao", "formato"),
            ),
        }
    if _allowed(permissions, "products"):
        rows = [dict(row or {}) for row in list(source.get("produtos", []) or [])]
        def product_available(row: dict[str, Any]) -> float:
            for key in ("available_qty", "disponivel", "qty", "quantidade", "qtd"):
                if key in row and str(row.get(key, "") or "").strip():
                    return _number(row.get(key))
            return 0.0

        without_stock = []
        low = []
        for row in rows:
            available = product_available(row)
            alert = _number(row.get("alerta", row.get("stock_minimo", 0)))
            enriched = dict(row)
            enriched["estado"] = (
                "Sem stock"
                if available <= 0
                else ("Stock baixo" if alert > 0 and available <= alert else _text(row.get("estado"), 50))
            )
            if available <= 0:
                without_stock.append(enriched)
            if (
                available <= 0
                or (alert > 0 and available <= alert)
                or any(token in _state(row) for token in ("critico", "baixo", "sem stock"))
            ):
                low.append(enriched)
        snapshot["produtos"] = {
            "registos": len(rows),
            "sem_stock": len(without_stock),
            "sem_stock_ou_criticos": len(low),
            "amostra_sem_stock": _sample(
                without_stock,
                identifiers=("codigo", "id"),
                description_fields=("descricao", "nome"),
            ),
            "amostra": _sample(
                low,
                identifiers=("codigo", "id"),
                description_fields=("descricao", "nome"),
            ),
        }
    if _allowed(permissions, "orders"):
        rows = [dict(row or {}) for row in list(source.get("encomendas", []) or [])]
        active = [
            row for row in rows
            if not any(token in _state(row) for token in ("conclu", "fechad", "cancelad", "entregue"))
        ]
        overdue = []
        for row in active:
            due = _date_value(
                row.get("data_entrega") or row.get("entrega") or row.get("prazo")
                or row.get("data_prevista")
            )
            if due and due < today:
                overdue.append(row)
        snapshot["encomendas"] = {
            "total": len(rows),
            "ativas": len(active),
            "atrasadas": len(overdue),
            "amostra_atrasada": _sample(
                overdue,
                identifiers=("numero", "of", "id"),
                description_fields=("referencia_cliente", "descricao"),
            ),
        }
    if _allowed(permissions, "quotes"):
        rows = [dict(row or {}) for row in list(source.get("orcamentos", []) or [])]
        open_rows = [
            row for row in rows
            if not any(token in _state(row) for token in ("aprov", "rejeit", "convert", "cancel"))
        ]
        snapshot["orcamentos"] = {
            "total": len(rows),
            "em_aberto": len(open_rows),
            "valor_aberto": round(sum(
                _number(row.get("total", row.get("valor_total", row.get("preco_total", 0))))
                for row in open_rows
            ), 2),
            "amostra": _sample(
                open_rows,
                identifiers=("numero", "codigo", "id"),
                description_fields=("descricao", "referencia"),
            ),
        }
    if _allowed(permissions, "planning"):
        rows = [dict(row or {}) for row in list(source.get("plano", []) or [])]
        snapshot["planeamento"] = {
            "blocos": len(rows),
            "minutos_planeados": round(sum(
                _number(row.get("duracao_min") or row.get("tempo_min") or row.get("minutos")
                        or row.get("duracao"))
                for row in rows
            ), 1),
            "amostra": _sample(
                rows,
                identifiers=("encomenda", "id"),
                description_fields=("operacao", "maquina", "posto"),
            ),
        }
    if _allowed(permissions, "purchase_notes"):
        rows = [dict(row or {}) for row in list(source.get("notas_encomenda", []) or [])]
        pending = [
            row for row in rows
            if not any(token in _state(row) for token in ("recebid", "conclu", "cancel", "fechad"))
        ]
        snapshot["compras"] = {
            "notas_encomenda": len(rows),
            "pendentes": len(pending),
            "amostra": _sample(
                pending,
                identifiers=("numero", "id"),
                description_fields=("fornecedor", "descricao"),
            ),
        }
    if _allowed(permissions, "transportes"):
        raw = source.get("viagens", source.get("transportes", source.get("transportes_viagens", [])))
        rows = [dict(row or {}) for row in list(raw or [])]
        active = [
            row for row in rows
            if not any(token in _state(row) for token in ("entregue", "conclu", "cancel", "fechad"))
        ]
        snapshot["transportes"] = {
            "viagens": len(rows),
            "ativas_ou_planeadas": len(active),
            "amostra": _sample(
                active,
                identifiers=("numero", "id", "codigo"),
                description_fields=("motorista", "matricula", "descricao"),
            ),
        }
    if _allowed(permissions, "quality"):
        quality_raw = source.get(
            "quality_nonconformities",
            source.get("qualidade", []),
        )
        rows = [dict(row or {}) for row in list(quality_raw or [])]
        open_rows = [
            row for row in rows
            if not any(token in _state(row) for token in ("conclu", "fechad", "resolvid", "cancel"))
        ]
        snapshot["qualidade"] = {
            "registos": len(rows),
            "em_aberto": len(open_rows),
            "amostra": _sample(
                open_rows,
                identifiers=("id", "numero", "codigo"),
                description_fields=("titulo", "descricao", "tipo"),
            ),
        }
    if _allowed(permissions, "avarias"):
        failure_raw = source.get("op_paragens", source.get("avarias", []))
        rows = [dict(row or {}) for row in list(failure_raw or [])]
        active = [
            row for row in rows
            if not _text(row.get("fechada_at"), 40)
            if not any(token in _state(row) for token in ("conclu", "fechad", "resolvid", "fim"))
        ]
        snapshot["avarias"] = {
            "registos": len(rows),
            "ativas": len(active),
            "amostra": _sample(
                active,
                identifiers=("id", "encomenda_numero", "ref_interna"),
                description_fields=("causa", "detalhe", "posto", "descricao", "motivo", "maquina"),
            ),
        }
    return snapshot


def _format_samples(rows: list[dict[str, str]]) -> str:
    values = []
    for row in rows[:5]:
        label = row.get("id", "-")
        if row.get("descricao"):
            label += f" — {row['descricao']}"
        if row.get("estado"):
            label += f" ({row['estado']})"
        values.append(label)
    return "\n".join(f"• {value}" for value in values)


def _count_label(value: Any, singular: str, plural: str | None = None) -> str:
    count = int(_number(value))
    return f"{count} {singular if count == 1 else (plural or singular + 's')}"


def _format_eur(value: Any) -> str:
    amount = _number(value)
    return f"{amount:,.2f}".replace(",", "\u0000").replace(".", ",").replace("\u0000", ".") + " EUR"


def contextual_question(
    question: str,
    conversation: list[dict[str, Any]] | None = None,
) -> str:
    """Resolve short follow-ups without changing what is displayed to the user."""

    current = _text(question, 1000)
    normalized = _normalized(current)
    if not conversation:
        return current
    previous_user = ""
    for turn in reversed(conversation[-10:]):
        if str(turn.get("role", "")).casefold() == "user":
            previous_user = _text(turn.get("content"), 800)
            if previous_user:
                break
    if not previous_user or _normalized(previous_user) == normalized:
        return current
    words = normalized.split()
    follow_up = (
        len(words) <= 5
        or normalized.startswith(("e ", "em ", "quanto", "quais", "qual deles", "esses", "isso"))
    )
    if follow_up:
        return f"{previous_user}. Seguimento do utilizador: {current}"
    return current


def _matched_module(question: str, snapshot: dict[str, Any], current_page: str = "") -> str:
    normalized = _normalized(question)
    modules = dict(snapshot.get("modulos_lugest", {}) or {})
    strong_terms = (
        ("nota encomenda", "purchase_notes"),
        ("nao conformidade", "quality"),
        ("assistente mp", "material_assistant"),
        ("materia prima", "materials"),
        ("ordem fabrico", "orders"),
        ("viagem", "transportes"),
        ("rota", "transportes"),
        ("orcamento", "quotes"),
        ("planeamento", "planning"),
        ("faturacao", "billing"),
        ("expedicao", "shipping"),
    )
    for term, key in strong_terms:
        if term in normalized and key in modules:
            return key
    aliases = {
        "stock_dashboard": ("dashboard", "visao executiva", "cockpit", "patrimonio"),
        "materials": ("materia prima", "material", "lote", "chapa", "perfil"),
        "products": ("produto", "artigo"),
        "customers": ("cliente",),
        "suppliers": ("fornecedor",),
        "orders": ("encomenda", "ordem fabrico", "of"),
        "quotes": ("orcamento", "proposta", "cotacao"),
        "planning": ("planeamento", "plano", "carga maquina"),
        "transportes": ("transporte", "viagem", "rota"),
        "material_assistant": ("assistente mp", "separacao material", "cativacao"),
        "operator": ("operador", "operacao"),
        "opp": ("opp",),
        "shipping": ("expedicao",),
        "billing": ("faturacao", "fatura"),
        "purchase_notes": ("nota encomenda", "aprovisionamento", "compra"),
        "quality": ("qualidade", "nao conformidade"),
        "diagnostics": ("diagnostico",),
        "avarias": ("avaria", "paragem"),
        "pulse": ("pulse", "oee", "desempenho"),
    }
    candidates: list[tuple[int, str]] = []
    for key, info in modules.items():
        terms = set(aliases.get(key, ()))
        terms.add(_normalized(info.get("nome")))
        for term in terms:
            if term and term in normalized:
                candidates.append((len(term), key))
    if candidates:
        return max(candidates)[1]
    return current_page if current_page in modules else ""


def deterministic_answer(
    question: str,
    snapshot: dict[str, Any],
    *,
    current_page: str = "",
) -> dict[str, Any]:
    normalized = _normalized(question)
    normalized_page = _normalized(current_page)
    page = ""
    title = "Leitura operacional"
    answer = ""
    suggestions: list[str] = []
    intent = "consult"
    requires_confirmation = False
    proposed_action: dict[str, Any] = {}
    requested_module = _matched_module(question, snapshot, current_page)
    mutating_request = any(
        token in normalized
        for token in (
            "cria", "criar", "adicion", "regist", "altera", "editar", "remove",
            "elimina", "apaga", "aprova", "rejeita", "guarda", "finaliza", "inicia",
        )
    )
    navigation_request = any(
        token in normalized
        for token in ("abre", "abrir", "vai para", "ir para", "mostra o menu", "leva me")
    )
    finance = dict(snapshot.get("financeiro", {}) or {})
    financial_query = any(
        token in normalized
        for token in (
            "patrim", "euro", "valor do stock", "valor stock", "financeir",
            "faturado", "recebido", "saldo clientes", "vendas", "compras recebidas",
        )
    )
    stock_query = any(
        token in normalized
        for token in ("stock", "material", "materia prima", "produto", "disponivel")
    )
    product_focus = (
        "produto" in normalized
        or (normalized_page == "products" and stock_query)
    )
    material_focus = (
        any(token in normalized for token in ("material", "materia prima"))
        or (normalized_page == "materials" and stock_query)
    )
    if mutating_request and requested_module:
        info = dict(snapshot.get("modulos_lugest", {}).get(requested_module, {}) or {})
        title, page = "Ação preparada", requested_module
        intent = "prepare_action"
        requires_confirmation = True
        if requested_module == "materials" and any(
            token in normalized for token in ("cria", "criar", "adicion", "stock", "lote")
        ):
            proposed_action = {
                "type": "create_material_stock",
                "command": _text(question, 1000),
                "target": "materials",
                "mutates_data": True,
            }
            answer = (
                "Entendi. Vou interpretar o material, dimensões, quantidade e lote, "
                "e abrir uma proposta preenchida para reveres antes de criar o stock.\n"
                "Responde «sim»/«continua» ou clica em «Rever e confirmar»."
            )
            suggestions = ["Rever e confirmar", "Cancelar"]
        else:
            proposed_action = {
                "type": "open_workflow",
                "command": _text(question, 1000),
                "target": requested_module,
                "mutates_data": False,
            }
            answer = (
                f"Entendi. Este pedido pertence a {info.get('nome', requested_module)}. "
                "Vou abrir o fluxo correto; a gravação só será feita depois da tua confirmação."
            )
            suggestions = ["Preparar no menu relacionado", "Cancelar"]
    elif navigation_request and requested_module:
        info = dict(snapshot.get("modulos_lugest", {}).get(requested_module, {}) or {})
        title, page = "Navegação no luGEST", requested_module
        intent = "navigate"
        answer = f"Posso abrir o módulo {info.get('nome', requested_module)}."
        suggestions = ["Explica este módulo"]
    elif financial_query and any(
        token in normalized for token in ("patrim", "stock", "euro", "valor")
    ):
        title, page = "Património em stock", "stock_dashboard"
        if finance:
            answer = (
                f"O património total em stock é {_format_eur(finance.get('stock_total'))}.\n\n"
                f"• Matéria-prima: {_format_eur(finance.get('stock_materias'))}\n"
                f"• Produto acabado: {_format_eur(finance.get('stock_produtos'))}\n"
                f"• Disponível: {_format_eur(finance.get('stock_disponivel'))}\n"
                f"• Reservado: {_format_eur(finance.get('stock_reservado'))}"
            )
        else:
            answer = "O valor financeiro do stock não está acessível com as permissões atuais."
        suggestions = ["Analisar stock crítico", "Resumir compras pendentes", "Mostrar encomendas atrasadas"]
    elif financial_query:
        title, page = "Resumo financeiro", "stock_dashboard"
        if finance:
            answer = (
                f"Vendas registadas: {_format_eur(finance.get('vendido_total'))}.\n"
                f"Faturado: {_format_eur(finance.get('faturado_total'))}.\n"
                f"Recebido: {_format_eur(finance.get('recebido_total'))}.\n"
                f"Saldo de clientes: {_format_eur(finance.get('saldo_clientes'))}.\n"
                f"Compras recebidas: {_format_eur(finance.get('compras_total'))}.\n"
                f"Compromissos abertos: {_format_eur(finance.get('compromissos_total'))}."
            )
        else:
            answer = "O resumo financeiro não está acessível com as permissões atuais."
        suggestions = ["Qual é o património em stock?", "Analisar stock crítico"]
    elif any(token in normalized for token in ("como funciona", "como posso", "como faco", "explica", "o que faz")):
        modules = dict(snapshot.get("modulos_lugest", {}) or {})
        matched_key = next(
            (
                key for key, info in modules.items()
                if _normalized(info.get("nome")) in normalized
                or any(part in normalized for part in _normalized(info.get("nome")).split())
            ),
            current_page if current_page in modules else "",
        )
        info = dict(modules.get(matched_key, {}) or {})
        if info:
            title, page = f"Como funciona: {info.get('nome')}", matched_key
            flows = "\n".join(f"• {item}" for item in list(info.get("fluxos", []) or []))
            answer = f"{info.get('descricao', '')}\n\nFluxos disponíveis:\n{flows}"
        else:
            answer = (
                "Posso explicar os menus e os fluxos do luGEST. Indica o módulo ou a tarefa "
                "que queres realizar, por exemplo: «Como crio uma viagem?»."
            )
        suggestions = ["Como funciona o Planeamento?", "Como crio uma nota de encomenda?"]
    elif stock_query and product_focus and not material_focus:
        products = dict(snapshot.get("produtos", {}) or {})
        title, page = "Stock de produtos", "products"
        answer = (
            f"Existem {_count_label(products.get('registos', 0), 'produto registado', 'produtos registados')}. "
            f"Há {_count_label(products.get('sem_stock', 0), 'produto')} sem stock e "
            f"{_count_label(products.get('sem_stock_ou_criticos', 0), 'produto')} "
            "sem stock ou abaixo do alerta."
        )
        sample_key = "amostra_sem_stock" if "sem stock" in normalized else "amostra"
        sample = _format_samples(list(products.get(sample_key, []) or []))
        if sample:
            answer += "\n\nProdutos a rever:\n" + sample
        elif products.get("sem_stock_ou_criticos", 0):
            answer += "\n\nExistem produtos a rever, mas a lista detalhada não está acessível."
        else:
            answer += "\n\nNão foram detetados produtos sem stock ou abaixo do alerta."
        suggestions = ["Analisar matéria-prima crítica", "Mostrar encomendas atrasadas"]
    elif stock_query and material_focus and not product_focus:
        materials = dict(snapshot.get("materia_prima", {}) or {})
        title, page = "Stock de matéria-prima", "materials"
        answer = (
            f"Existem {_count_label(materials.get('registos', 0), 'lote de matéria-prima', 'lotes de matéria-prima')} e "
            f"há {_count_label(materials.get('criticos', 0), 'lote crítico ou sem disponibilidade', 'lotes críticos ou sem disponibilidade')}."
        )
        sample = _format_samples(list(materials.get("amostra_critica", []) or []))
        if sample:
            answer += "\n\nLotes a rever:\n" + sample
        else:
            answer += "\n\nNão foram detetados lotes críticos."
        suggestions = ["Analisar stock de produtos", "Mostrar encomendas atrasadas"]
    elif stock_query:
        materials = dict(snapshot.get("materia_prima", {}) or {})
        products = dict(snapshot.get("produtos", {}) or {})
        title, page = "Stock e disponibilidade", "materials"
        answer = (
            f"Matéria-prima: {materials.get('registos', 0)} registos, "
            f"{materials.get('criticos', 0)} críticos ou sem disponibilidade. "
            f"Produtos: {products.get('registos', 0)} registos, "
            f"{products.get('sem_stock_ou_criticos', 0)} críticos ou sem stock."
        )
        material_sample = _format_samples(list(materials.get("amostra_critica", []) or []))
        product_sample = _format_samples(list(products.get("amostra", []) or []))
        if material_sample:
            answer += "\n\nMatéria-prima a rever:\n" + material_sample
        if product_sample:
            answer += "\n\nProdutos a rever:\n" + product_sample
        suggestions = ["Mostrar encomendas atrasadas", "Resumir compras pendentes"]
    elif (
        any(token in normalized for token in ("atras", "encomenda", "ordem", "of"))
        and not ("nota" in normalized and "encomenda" in normalized)
    ):
        rows = dict(snapshot.get("encomendas", {}) or {})
        title, page = "Encomendas e prazos", "orders"
        answer = (
            f"Há {_count_label(rows.get('ativas', 0), 'encomenda ativa', 'encomendas ativas')} "
            f"em {_count_label(rows.get('total', 0), 'registo')}. "
            f"Detetei {_count_label(rows.get('atrasadas', 0), 'encomenda')} com prazo anterior a hoje."
        )
        sample = _format_samples(list(rows.get("amostra_atrasada", []) or []))
        if sample:
            answer += "\n\nEncomendas a rever:\n" + sample
        suggestions = ["Analisar stock crítico", "Resumir planeamento"]
    elif any(token in normalized for token in ("compra", "nota", "fornecedor", "aprovision")):
        rows = dict(snapshot.get("compras", {}) or {})
        title, page = "Compras e aprovisionamento", "purchase_notes"
        answer = (
            f"Há {_count_label(rows.get('notas_encomenda', 0), 'nota de encomenda', 'notas de encomenda')}; "
            f"{_count_label(rows.get('pendentes', 0), 'nota está pendente', 'notas estão pendentes')}."
        )
        sample = _format_samples(list(rows.get("amostra", []) or []))
        if sample:
            answer += "\n\nRegistos a acompanhar:\n" + sample
        suggestions = ["Analisar stock crítico", "Mostrar encomendas atrasadas"]
    elif any(token in normalized for token in ("plane", "carga", "maquina", "capacidade")):
        rows = dict(snapshot.get("planeamento", {}) or {})
        title, page = "Planeamento", "planning"
        answer = (
            f"O plano contém {_count_label(rows.get('blocos', 0), 'bloco')} e cerca de "
            f"{rows.get('minutos_planeados', 0):g} minutos registados."
        )
        suggestions = ["Mostrar encomendas atrasadas", "Resumir avarias"]
    elif any(token in normalized for token in ("transporte", "viagem", "entrega", "rota")):
        rows = dict(snapshot.get("transportes", {}) or {})
        title, page = "Transportes", "transportes"
        answer = (
            f"Estão registadas {_count_label(rows.get('viagens', 0), 'viagem', 'viagens')}; "
            f"{_count_label(rows.get('ativas_ou_planeadas', 0), 'está ativa ou planeada', 'estão ativas ou planeadas')}."
        )
        suggestions = ["Mostrar encomendas atrasadas", "Resumir compras pendentes"]
    elif any(token in normalized for token in ("avaria", "paragem", "incidente")):
        rows = dict(snapshot.get("avarias", {}) or {})
        title, page = "Avarias", "avarias"
        answer = (
            f"Há {_count_label(rows.get('registos', 0), 'registo de avaria', 'registos de avaria')} e "
            f"{_count_label(rows.get('ativas', 0), 'situação ainda ativa', 'situações ainda ativas')}."
        )
        suggestions = ["Resumir planeamento", "Mostrar encomendas atrasadas"]
    elif any(token in normalized for token in ("qualidade", "nao conform", "rejei")):
        rows = dict(snapshot.get("qualidade", {}) or {})
        title, page = "Qualidade", "quality"
        answer = (
            f"Há {_count_label(rows.get('registos', 0), 'registo de qualidade', 'registos de qualidade')} e "
            f"{_count_label(rows.get('em_aberto', 0), 'está em aberto', 'estão em aberto')}."
        )
        suggestions = ["Resumir avarias", "Mostrar encomendas atrasadas"]
    elif any(token in normalized for token in ("orcamento", "proposta", "cotacao")):
        rows = dict(snapshot.get("orcamentos", {}) or {})
        title, page = "Orçamentos", "quotes"
        answer = (
            f"Há {_count_label(rows.get('total', 0), 'orçamento')}; "
            f"{_count_label(rows.get('em_aberto', 0), 'está em aberto', 'estão em aberto')}. "
            f"Valor aberto registado: {rows.get('valor_aberto', 0):,.2f} EUR."
        )
        suggestions = ["Mostrar encomendas atrasadas", "Analisar stock crítico"]
    else:
        sections = []
        for key, label, metric in (
            ("encomendas", "encomendas ativas", "ativas"),
            ("materia_prima", "matérias-primas críticas", "criticos"),
            ("compras", "compras pendentes", "pendentes"),
            ("avarias", "avarias ativas", "ativas"),
        ):
            rows = dict(snapshot.get(key, {}) or {})
            if rows:
                sections.append(f"{rows.get(metric, 0)} {label}")
        answer = "Resumo disponível: " + (", ".join(sections) if sections else "sem dados acessíveis.")
        answer += (
            "\n\nPosso analisar stock, encomendas e atrasos, planeamento, compras, "
            "transportes, qualidade, avarias ou orçamentos."
        )
        suggestions = ["Analisar stock crítico", "Mostrar encomendas atrasadas", "Resumir planeamento"]
    return {
        "title": title,
        "answer": answer,
        "engine": "regras-locais",
        "navigation_target": page,
        "suggestions": suggestions[:3],
        "read_only": True,
        "intent": intent,
        "requires_confirmation": requires_confirmation,
        "proposed_action": proposed_action,
    }


def _is_verified_operational_question(question: str) -> bool:
    """Keep direct ERP facts deterministic instead of letting a model rewrite them."""

    normalized = _normalized(question)
    if any(
        token in normalized
        for token in (
            "patrim", "euro", "financeir", "faturado", "recebido",
            "saldo clientes", "valor stock",
        )
    ):
        return True
    if any(token in normalized for token in ("stock", "disponib")):
        return True
    if any(token in normalized for token in ("materia", "material", "produto")) and any(
        token in normalized for token in ("critic", "sem", "baixo", "alerta")
    ):
        return True
    if any(token in normalized for token in ("encomend", "ordem", "prazo")) and "atras" in normalized:
        return True
    if any(token in normalized for token in ("compra", "aprovision", "nota")) and any(
        token in normalized for token in ("pendent", "abert", "resum", "estado")
    ):
        return True
    topic_stems = (
        ("plane", "carga", "capacidade"),
        ("transport", "viag", "entrega", "rota"),
        ("avari", "parag", "incident"),
        ("qualidade", "nao conform", "rejei"),
        ("orcament", "proposta", "cotacao"),
    )
    factual_stems = ("resum", "quant", "list", "mostr", "qual", "estado", "abert", "pendent", "ativ")
    return any(token in normalized for group in topic_stems for token in group) and any(
        token in normalized for token in factual_stems
    )


def _valid_chat_answer(answer: str, question: str) -> bool:
    """Reject model echoes and internal/meta answers before they reach the UI."""

    answer_norm = _normalized(answer)
    question_norm = _normalized(question)
    if not answer_norm or answer_norm == question_norm or len(answer_norm) < 4:
        return False
    if question_norm and (
        answer_norm in {f"tu {question_norm}", f"utilizador {question_norm}"}
        or (
            len(answer_norm.split()) <= len(question_norm.split()) + 2
            and question_norm in answer_norm
        )
    ):
        return False
    forbidden = (
        "resumo json fornecido",
        "instrucoes do sistema",
        "system prompt",
        "nao posso fazer algo que de para perceber",
    )
    return not any(token in answer_norm for token in forbidden)


_WEATHER_CODES = {
    0: "céu limpo",
    1: "predominantemente limpo",
    2: "parcialmente nublado",
    3: "nublado",
    45: "nevoeiro",
    48: "nevoeiro com geada",
    51: "chuvisco fraco",
    53: "chuvisco",
    55: "chuvisco forte",
    61: "chuva fraca",
    63: "chuva",
    65: "chuva forte",
    71: "neve fraca",
    73: "neve",
    75: "neve forte",
    80: "aguaceiros fracos",
    81: "aguaceiros",
    82: "aguaceiros fortes",
    95: "trovoada",
    96: "trovoada com granizo",
    99: "trovoada forte com granizo",
}


def _is_weather_question(question: str) -> bool:
    normalized = _normalized(question)
    return any(
        token in normalized
        for token in (
            "meteorologia", "previsao do tempo", "estado do tempo", "vai chover",
            "temperatura", "que tempo", "como esta o tempo", "como vai estar o tempo",
            "tempo hoje", "tempo amanha", "tempo em ", "tempo no ", "tempo na ",
        )
    )


def _weather_location(question: str) -> str:
    text = re.sub(r"[?!.]+$", "", str(question or "").strip())
    matches = re.findall(r"\b(?:em|no|na|para)\s+([\wÀ-ÿ .'-]+)", text, flags=re.IGNORECASE)
    if matches:
        candidate = matches[-1].strip()
        candidate = re.sub(
            r"^(?:hoje|amanhã|amanha)\s+(?:em|no|na)\s+",
            "",
            candidate,
            flags=re.IGNORECASE,
        )
        candidate = re.sub(
            r"\s+(?:hoje|amanhã|amanha|agora|esta\s+manhã|esta\s+manha|"
            r"esta\s+tarde|esta\s+noite)$",
            "",
            candidate,
            flags=re.IGNORECASE,
        ).strip()
        if candidate and _normalized(candidate) not in {"hoje", "amanha"}:
            return candidate
    return "Lisboa, Portugal"


def _online_weather_answer(
    question: str,
    *,
    location: str = "",
    date_reference: str = "",
    timeout_seconds: float = 8.0,
) -> dict[str, Any]:
    location = _text(location, 180) or _weather_location(question)
    geocode_url = (
        "https://geocoding-api.open-meteo.com/v1/search?"
        + urllib.parse.urlencode(
            {"name": location, "count": 1, "language": "pt", "format": "json"}
        )
    )
    headers = {"Accept": "application/json", "User-Agent": "luGEST-ERP-Copilot"}
    with urllib.request.urlopen(
        urllib.request.Request(geocode_url, headers=headers),
        timeout=max(2.0, timeout_seconds),
    ) as response:
        geocode = json.loads(response.read(200_001).decode("utf-8"))
    results = list(geocode.get("results", []) or [])
    if not results:
        nominatim_url = (
            "https://nominatim.openstreetmap.org/search?"
            + urllib.parse.urlencode(
                {
                    "q": location,
                    "format": "jsonv2",
                    "limit": 1,
                    "countrycodes": "pt",
                    "accept-language": "pt",
                }
            )
        )
        with urllib.request.urlopen(
            urllib.request.Request(nominatim_url, headers=headers),
            timeout=max(2.0, timeout_seconds),
        ) as response:
            alternative = json.loads(response.read(200_001).decode("utf-8"))
        if not alternative:
            raise ValueError(f"Não encontrei a localização «{location}».")
        row = dict(alternative[0] or {})
        display_parts = [
            part.strip() for part in str(row.get("display_name", "") or "").split(",") if part.strip()
        ]
        place = {
            "name": display_parts[0] if display_parts else location,
            "country": display_parts[-1] if display_parts else "Portugal",
            "latitude": row.get("lat"),
            "longitude": row.get("lon"),
        }
    else:
        place = dict(results[0] or {})
    forecast_url = (
        "https://api.open-meteo.com/v1/forecast?"
        + urllib.parse.urlencode(
            {
                "latitude": place.get("latitude"),
                "longitude": place.get("longitude"),
                "current": (
                    "temperature_2m,apparent_temperature,precipitation,"
                    "weather_code,wind_speed_10m"
                ),
                "daily": (
                    "weather_code,temperature_2m_max,temperature_2m_min,"
                    "precipitation_probability_max"
                ),
                "timezone": "Europe/Lisbon",
                "forecast_days": 3,
            }
        )
    )
    with urllib.request.urlopen(
        urllib.request.Request(forecast_url, headers=headers),
        timeout=max(2.0, timeout_seconds),
    ) as response:
        forecast = json.loads(response.read(300_001).decode("utf-8"))
    current = dict(forecast.get("current", {}) or {})
    daily = dict(forecast.get("daily", {}) or {})
    tomorrow = date_reference == "tomorrow" or (
        date_reference not in {"today", "tomorrow"} and "amanha" in _normalized(question)
    )
    index = 1 if tomorrow else 0
    label = "Amanhã" if tomorrow else "Hoje"
    codes = list(daily.get("weather_code", []) or [])
    minimums = list(daily.get("temperature_2m_min", []) or [])
    maximums = list(daily.get("temperature_2m_max", []) or [])
    rain = list(daily.get("precipitation_probability_max", []) or [])
    code = int(codes[index]) if len(codes) > index and codes[index] is not None else int(
        current.get("weather_code", -1) or -1
    )
    description = _WEATHER_CODES.get(code, "condições variáveis")
    place_name = ", ".join(
        part for part in (_text(place.get("name"), 80), _text(place.get("country"), 80)) if part
    )
    if tomorrow:
        answer = (
            f"{label} em {place_name}: {description}, "
            f"mínima de {float(minimums[index]):.0f} °C e máxima de {float(maximums[index]):.0f} °C. "
            f"Probabilidade máxima de precipitação: {int(rain[index] or 0)}%."
        )
    else:
        answer = (
            f"Agora em {place_name}: {description}, {float(current.get('temperature_2m', 0)):.1f} °C "
            f"(sensação {float(current.get('apparent_temperature', 0)):.1f} °C), "
            f"vento {float(current.get('wind_speed_10m', 0)):.0f} km/h. "
            f"Hoje: {float(minimums[index]):.0f}–{float(maximums[index]):.0f} °C, "
            f"precipitação até {int(rain[index] or 0)}%."
        )
    return {
        "title": "Meteorologia",
        "answer": answer,
        "engine": "Open-Meteo · dados atuais",
        "navigation_target": "",
        "suggestions": ["E amanhã?", "Que tempo está no Porto?"],
        "read_only": True,
        "intent": "consult",
        "requires_confirmation": False,
        "proposed_action": {},
    }


class ERPCopilot:
    def __init__(self, *, ollama_url: str = "", ollama_model: str = "", timeout_seconds: float = 45.0) -> None:
        self.ollama_url = _text(
            ollama_url or os.getenv("LUGEST_OLLAMA_URL") or "http://127.0.0.1:11434", 300
        ).rstrip("/")
        self.ollama_model = _text(
            ollama_model or os.getenv("LUGEST_OLLAMA_MODEL") or "qwen3:4b", 100
        )
        self.openai_api_key = _text(
            os.getenv("LUGEST_OPENAI_API_KEY") or os.getenv("OPENAI_API_KEY"), 500
        )
        self.openai_model = _text(
            os.getenv("LUGEST_OPENAI_MODEL") or "gpt-5.6-sol", 100
        )
        self.openai_base_url = _text(
            os.getenv("LUGEST_OPENAI_BASE_URL") or "https://api.openai.com/v1", 300
        ).rstrip("/")
        self.timeout_seconds = max(2.0, min(float(timeout_seconds or 45), 180.0))

    @staticmethod
    def _response_output_text(payload: dict[str, Any]) -> str:
        chunks: list[str] = []
        for item in list(payload.get("output", []) or []):
            if not isinstance(item, dict) or item.get("type") != "message":
                continue
            for content in list(item.get("content", []) or []):
                if isinstance(content, dict) and content.get("type") == "output_text":
                    chunks.append(str(content.get("text", "") or ""))
        return "\n".join(part.strip() for part in chunks if part.strip()).strip()

    def _openai_response(self, body: dict[str, Any]) -> dict[str, Any]:
        if not self.openai_api_key:
            raise RuntimeError("OpenAI não configurada")
        request = urllib.request.Request(
            f"{self.openai_base_url}/responses",
            data=json.dumps(body, ensure_ascii=False).encode("utf-8"),
            headers={
                "Accept": "application/json",
                "Content-Type": "application/json; charset=utf-8",
                "Authorization": f"Bearer {self.openai_api_key}",
                "User-Agent": "luGEST-ERP-Copilot",
            },
            method="POST",
        )
        with urllib.request.urlopen(request, timeout=self.timeout_seconds) as response:
            return json.loads(response.read(1_500_001).decode("utf-8"))

    def _route_question(
        self,
        question: str,
        *,
        current_page: str = "",
        conversation: list[dict[str, Any]] | None = None,
    ) -> dict[str, str] | None:
        """Let a model select a typed host tool; no phrase-specific parsing is required."""

        recent = [
            {
                "role": str(turn.get("role", "") or ""),
                "content": _text(turn.get("content"), 400),
            }
            for turn in list(conversation or [])[-6:]
            if str(turn.get("role", "") or "").casefold() in {"user", "assistant"}
        ]
        instructions = (
            "És o router de ferramentas do ERP luGEST. Interpreta português europeu e devolve "
            "exclusivamente o objeto pedido pelo esquema. Não respondas à pergunta. "
            "weather: meteorologia; extrai apenas a localidade, sem palavras temporais. "
            "erp_query: consulta de dados internos do ERP. navigate: abrir um módulo. "
            "prepare_action: pedido que cria, altera, remove, aprova, inicia ou finaliza dados. "
            "general_chat: cultura geral, conversa ou qualquer questão fora do ERP. "
            "Para criar stock de matéria-prima usa action_type=create_material_stock. "
            "Módulos válidos: " + ", ".join(MODULE_KNOWLEDGE) + "."
        )
        input_text = json.dumps(
            {
                "menu_atual": current_page,
                "conversa_recente": recent,
                "pergunta": question,
            },
            ensure_ascii=False,
        )
        if self.openai_api_key:
            try:
                payload = self._openai_response(
                    {
                        "model": self.openai_model,
                        "instructions": instructions,
                        "input": input_text,
                        "text": {
                            "format": {
                                "type": "json_schema",
                                "name": "lugest_tool_route",
                                "strict": True,
                                "schema": _COPILOT_ROUTE_SCHEMA,
                            }
                        },
                        "store": False,
                    }
                )
                route = json.loads(self._response_output_text(payload))
                if isinstance(route, dict) and route.get("intent"):
                    return {key: _text(value, 1000) for key, value in route.items()}
            except Exception:
                pass

        body = json.dumps(
            {
                "model": self.ollama_model,
                "messages": [
                    {"role": "system", "content": instructions},
                    {"role": "user", "content": input_text},
                ],
                "format": _COPILOT_ROUTE_SCHEMA,
                "stream": False,
                "think": False,
                "keep_alive": "10m",
                "options": {"temperature": 0.0, "num_ctx": 4096},
            },
            ensure_ascii=False,
        ).encode("utf-8")
        request = urllib.request.Request(
            f"{self.ollama_url}/api/chat",
            data=body,
            headers={
                "Accept": "application/json",
                "Content-Type": "application/json; charset=utf-8",
                "User-Agent": "luGEST-ERP-Copilot",
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=self.timeout_seconds) as response:
                payload = json.loads(response.read(300_001).decode("utf-8"))
            message = payload.get("message", {})
            route = json.loads(
                str(message.get("content", "") if isinstance(message, dict) else "")
            )
            if isinstance(route, dict) and route.get("intent"):
                return {key: _text(value, 1000) for key, value in route.items()}
        except Exception:
            pass
        return None

    def _model_chat(
        self,
        question: str,
        *,
        system_prompt: str,
        history: list[dict[str, str]],
    ) -> tuple[str, str]:
        if self.openai_api_key:
            try:
                input_items = [
                    {"role": item["role"], "content": item["content"]}
                    for item in history
                ]
                input_items.append({"role": "user", "content": question})
                payload = self._openai_response(
                    {
                        "model": self.openai_model,
                        "instructions": system_prompt,
                        "input": input_items,
                        "store": False,
                    }
                )
                answer = _text(self._response_output_text(payload), 4000)
                if _valid_chat_answer(answer, question):
                    return answer, f"openai-{self.openai_model}"
            except Exception:
                pass
        body = json.dumps(
            {
                "model": self.ollama_model,
                "messages": [
                    {"role": "system", "content": system_prompt},
                    *history,
                    {"role": "user", "content": question},
                ],
                "stream": False,
                "think": False,
                "keep_alive": "10m",
                "options": {"temperature": 0.2, "num_ctx": 8192},
            },
            ensure_ascii=False,
        ).encode("utf-8")
        request = urllib.request.Request(
            f"{self.ollama_url}/api/chat",
            data=body,
            headers={
                "Accept": "application/json",
                "Content-Type": "application/json; charset=utf-8",
                "User-Agent": "luGEST-ERP-Copilot",
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=self.timeout_seconds) as response:
                payload = json.loads(response.read(500_001).decode("utf-8"))
            message = payload.get("message", {})
            raw = str(message.get("content", "") if isinstance(message, dict) else "").strip()
            if "</think>" in raw:
                raw = raw.split("</think>", 1)[1].strip()
            answer = _text(raw, 4000)
            if _valid_chat_answer(answer, question):
                return answer, f"ollama-{self.ollama_model}"
        except Exception:
            pass
        return "", ""

    def status(self, *, timeout_seconds: float = 1.2) -> dict[str, Any]:
        if self.openai_api_key:
            return {
                "available": True,
                "service_online": True,
                "model": self.openai_model,
                "models": [self.openai_model],
                "provider": "openai",
                "message": "IA OpenAI configurada; Ollama disponível como contingência.",
            }
        request = urllib.request.Request(
            f"{self.ollama_url}/api/tags",
            headers={"Accept": "application/json", "User-Agent": "LuGEST-ERP-Copilot"},
        )
        try:
            with urllib.request.urlopen(request, timeout=max(0.3, timeout_seconds)) as response:
                payload = json.loads(response.read(300_001).decode("utf-8"))
            models = [
                _text(row.get("name") or row.get("model"), 100)
                for row in list(payload.get("models", []) or [])
                if isinstance(row, dict)
            ]
            wanted = self.ollama_model.casefold()
            available = any(
                model.casefold() == wanted
                or model.casefold().split(":", 1)[0] == wanted.split(":", 1)[0]
                for model in models
            )
            return {
                "available": available,
                "service_online": True,
                "model": self.ollama_model,
                "models": models[:20],
                "message": "IA local disponível." if available
                else f"Ollama disponível, mas falta instalar {self.ollama_model}.",
            }
        except Exception:
            return {
                "available": False,
                "service_online": False,
                "model": self.ollama_model,
                "models": [],
                "message": "Modo sem IA: análises locais continuam disponíveis.",
            }

    def resolve_action_followup(
        self,
        question: str,
        action: dict[str, Any],
        *,
        conversation: list[dict[str, Any]] | None = None,
    ) -> dict[str, str]:
        """Understand a natural follow-up while an ERP action awaits confirmation."""

        text = _text(question, 700)
        normalized = _normalized(text)
        positive = (
            "sim", "ok", "confirma", "pode ser", "avanca", "faz", "segue",
            "siga", "entao", "continua", "procede", "faz isso", "forca",
        )
        negative = ("nao", "cancela", "cancelar", "esquece", "para", "deixa estar")
        if normalized in positive or any(
            normalized.startswith(f"{token} ") for token in positive
        ):
            return {"decision": "execute", "answer": ""}
        if normalized in negative or any(
            normalized.startswith(f"{token} ") for token in negative
        ):
            return {"decision": "cancel", "answer": "A ação foi cancelada. Nada foi alterado."}
        correction_terms = (
            "afinal", "corrige", "corrigir", "muda", "alterar para", "troca",
            "deve ser", "passa a", "o lote e", "a quantidade e", "a espessura e",
        )
        if any(token in normalized for token in correction_terms):
            return {
                "decision": "update",
                "correction": text,
                "answer": (
                    f"Associei esta correção à proposta: {text}. "
                    "A ficha será recalculada quando confirmares."
                ),
            }

        recent = []
        for turn in list(conversation or [])[-6:]:
            role = str(turn.get("role", "") or "").casefold()
            content = _text(turn.get("content"), 500)
            if role in {"user", "assistant"} and content:
                recent.append({"role": role, "content": content})
        prompt = (
            "Estás a interpretar a resposta do utilizador a uma ação pendente no ERP luGEST. "
            "Responde obrigatoriamente numa única linha começada por EXECUTAR, CANCELAR, AJUSTAR ou ESCLARECER. "
            "EXECUTAR quando quer avançar, mesmo de forma coloquial ou implícita. "
            "CANCELAR quando não quer a ação. ESCLARECER quando faz uma pergunta, corrige dados ou a intenção "
            "não é segura. AJUSTAR quando fornece ou corrige dados da ação; acrescenta após `|` um resumo curto "
            "da correção. Depois de ESCLARECER, acrescenta após `|` uma resposta curta e útil em português. "
            "Nunca inventes que a ação foi concluída.\n"
            f"Ação pendente: {json.dumps(action, ensure_ascii=False)}\n"
            f"Conversa: {json.dumps(recent, ensure_ascii=False)}\n"
            f"Nova frase: {text}"
        )
        body = json.dumps(
            {
                "model": self.ollama_model,
                "messages": [
                    {"role": "system", "content": "Classificador seguro de intenção do ERP luGEST."},
                    {"role": "user", "content": prompt},
                ],
                "stream": False,
                "think": False,
                "keep_alive": "10m",
                "options": {"temperature": 0.0, "num_ctx": 2048},
            },
            ensure_ascii=False,
        ).encode("utf-8")
        request = urllib.request.Request(
            f"{self.ollama_url}/api/chat",
            data=body,
            headers={
                "Accept": "application/json",
                "Content-Type": "application/json; charset=utf-8",
                "User-Agent": "LuGEST-ERP-Copilot",
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=self.timeout_seconds) as response:
                payload = json.loads(response.read(100_001).decode("utf-8"))
            message = payload.get("message", {})
            raw = _text(message.get("content") if isinstance(message, dict) else "", 1200)
            if "</think>" in raw:
                raw = raw.split("</think>", 1)[1].strip()
            upper = raw.upper()
            decision_match = re.search(
                r"\b(EXECUTAR|CANCELAR|AJUSTAR|ESCLARECER)\b",
                upper[:300],
            )
            decision_word = decision_match.group(1) if decision_match else ""
            if decision_word == "EXECUTAR":
                return {"decision": "execute", "answer": ""}
            if decision_word == "CANCELAR":
                return {"decision": "cancel", "answer": "A ação foi cancelada. Nada foi alterado."}
            if decision_word == "AJUSTAR":
                detail = raw.split("|", 1)[1].strip() if "|" in raw else text
                return {
                    "decision": "update",
                    "correction": text,
                    "answer": (
                        f"Atualizei a instrução com: {detail}. "
                        "A proposta será recalculada quando confirmares."
                    ),
                }
            if decision_word == "ESCLARECER":
                detail = raw.split("|", 1)[1].strip() if "|" in raw else ""
                return {
                    "decision": "clarify",
                    "answer": detail or (
                        "A ação continua preparada. Indica o que queres corrigir ou confirma para avançar."
                    ),
                }
        except Exception:
            pass
        return {
            "decision": "clarify",
            "answer": (
                "Não interpretei esta resposta com segurança. A ação continua preparada; "
                "responde «sim» para avançar, «não» para cancelar, ou indica a correção pretendida."
            ),
        }

    def ask(
        self,
        question: str,
        snapshot: dict[str, Any],
        *,
        current_page: str = "",
        conversation: list[dict[str, Any]] | None = None,
    ) -> dict[str, Any]:
        question = _text(question, 1000)
        if not question:
            raise ValueError("Escreve uma pergunta para o Copiloto.")
        resolved_question = contextual_question(question, conversation)
        route = self._route_question(
            question,
            current_page=current_page,
            conversation=conversation,
        )
        route_intent = str((route or {}).get("intent", "") or "")
        if route_intent == "weather" or (not route and _is_weather_question(resolved_question)):
            try:
                return _online_weather_answer(
                    resolved_question,
                    location=str((route or {}).get("location", "") or ""),
                    date_reference=str((route or {}).get("date_reference", "") or ""),
                    timeout_seconds=min(self.timeout_seconds, 10.0),
                )
            except Exception as exc:
                return {
                    "title": "Meteorologia indisponível",
                    "answer": (
                        "Não consegui consultar a meteorologia em direto. "
                        "Confirma a ligação à Internet e indica a cidade, por exemplo: "
                        "«Que tempo está no Porto?»"
                    ),
                    "engine": "Open-Meteo indisponível",
                    "navigation_target": "",
                    "suggestions": [],
                    "read_only": True,
                    "intent": "consult",
                    "requires_confirmation": False,
                    "proposed_action": {},
                    "error": _text(exc, 300),
                }
        if route_intent == "prepare_action":
            module = str((route or {}).get("module", "") or "")
            action_type = str((route or {}).get("action_type", "") or "")
            if action_type == "create_material_stock" or module == "materials":
                return {
                    "title": "Ação preparada",
                    "answer": (
                        "Interpretei o pedido como criação de stock de matéria-prima. "
                        "Vou abrir uma proposta preenchida para reveres todos os campos antes de gravar."
                    ),
                    "engine": (
                        f"openai-{self.openai_model}"
                        if self.openai_api_key
                        else f"ollama-{self.ollama_model} · ferramentas"
                    ),
                    "navigation_target": "materials",
                    "suggestions": ["Rever e confirmar", "Cancelar"],
                    "read_only": True,
                    "intent": "prepare_action",
                    "requires_confirmation": True,
                    "proposed_action": {
                        "type": "create_material_stock",
                        # Keep the complete original request. The router chooses the
                        # tool; the domain parser receives every material/detail.
                        "command": question,
                        "target": "materials",
                        "mutates_data": True,
                    },
                }
            if module in MODULE_KNOWLEDGE:
                return {
                    "title": "Ação preparada",
                    "answer": (
                        f"O pedido pertence ao módulo {MODULE_KNOWLEDGE[module]['nome']}. "
                        "Vou abrir o fluxo correto; nenhuma alteração é gravada sem revisão."
                    ),
                    "engine": "ferramentas luGEST",
                    "navigation_target": module,
                    "suggestions": ["Preparar no menu relacionado", "Cancelar"],
                    "read_only": True,
                    "intent": "prepare_action",
                    "requires_confirmation": True,
                    "proposed_action": {
                        "type": "open_workflow",
                        "command": str((route or {}).get("command", "") or question),
                        "target": module,
                        "mutates_data": False,
                    },
                }
        if route_intent == "navigate":
            module = str((route or {}).get("module", "") or "")
            if module in MODULE_KNOWLEDGE:
                return {
                    "title": "Navegação no luGEST",
                    "answer": f"Posso abrir o módulo {MODULE_KNOWLEDGE[module]['nome']}.",
                    "engine": "ferramentas luGEST",
                    "navigation_target": module,
                    "suggestions": ["Explica este módulo"],
                    "read_only": True,
                    "intent": "navigate",
                    "requires_confirmation": False,
                    "proposed_action": {},
                }
        fallback = deterministic_answer(
            resolved_question,
            snapshot,
            current_page=current_page,
        )
        if fallback.get("intent") in {"navigate", "prepare_action"}:
            return {
                **fallback,
                "engine": "fluxos luGEST verificados",
            }
        if _is_verified_operational_question(resolved_question):
            return {
                **fallback,
                "engine": "dados ERP verificados",
            }
        safe_history = []
        for turn in list(conversation or [])[-8:]:
            role = str(turn.get("role", "") or "").casefold()
            content = _text(turn.get("content"), 700)
            if role in {"user", "assistant"} and content:
                safe_history.append({"role": role, "content": content})
        system_prompt = (
            "És o Copiloto do ERP industrial luGEST. Conversa naturalmente em português europeu. "
            "Responde diretamente à intenção do utilizador, nunca repitas ou imites a pergunta. "
            "Mantém as respostas curtas (normalmente 2 a 6 linhas), concretas e fáceis de executar. "
            "Conheces os módulos, dados e fluxos presentes no contexto ERP abaixo. "
            "Podes responder normalmente a perguntas de cultura geral e usar o teu conhecimento do modelo. "
            "Quando a pergunta for de cultura geral, responde apenas ao tema e não menciones o ERP. "
            "A restrição seguinte aplica-se apenas a factos internos da empresa: usa somente o contexto ERP; "
            "se não existirem, diz claramente que o dado não está disponível. "
            "Não inventes pessoas da empresa, valores, datas, estados ou ações. "
            "O programa anfitrião trata das ferramentas e confirmações: quando for pedida uma alteração, "
            "explica brevemente o que será preparado, mas nunca afirmes que já foi gravado. "
            "Não menciones JSON, prompts, instruções internas ou limitações técnicas salvo se perguntado. "
            f"Menu atual: {_text(current_page, 60) or 'não indicado'}.\n"
            "Contexto ERP verificado:\n"
            f"{json.dumps(snapshot, ensure_ascii=False, separators=(',', ':'))}"
        )
        answer, engine = self._model_chat(
            question,
            system_prompt=system_prompt,
            history=safe_history,
        )
        if answer:
            return {
                **fallback,
                "title": "Análise do Copiloto",
                "answer": answer,
                "engine": engine,
            }
        return fallback
