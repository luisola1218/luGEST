from __future__ import annotations

import math
from datetime import datetime
from pathlib import Path
from typing import Any

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas as pdf_canvas

from .dossier_layout import Column, DossierLayout


def _branding(backend) -> tuple[dict[str, Any], Path | None, str]:
    branding = dict(backend.branding_settings() or {})
    logo_text = str(branding.get("logo_path", "") or "").strip()
    logo = Path(logo_text) if logo_text and Path(logo_text).exists() else None
    issued = str(backend.desktop_main.now_iso() or "").replace("T", " ")[:19]
    return branding, logo, issued


def _layout(backend, c, page_size, code: str, issued: str, logo: Path | None) -> DossierLayout:
    branding = backend.branding_settings()
    return DossierLayout(
        c,
        page_size,
        primary=branding.get("primary_color", "#00A6A6"),
        logo_path=logo,
        logo_draw=backend._draw_operator_logo_plate,
        issued_at=issued,
        document_code=code,
    )


def render_planning_deadlines(backend, path: str | Path, rows: list[dict[str, Any]]) -> Path:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    page_size = landscape(A4)
    _brand, logo, issued = _branding(backend)
    columns = [
        Column("Encomenda", 30 * mm), Column("Cliente", 52 * mm), Column("Entrega", 25 * mm),
        Column("Grupos", 18 * mm, "center"), Column("Planeado", 29 * mm),
        Column("Fim do fluxo", 38 * mm), Column("Estado", 48 * mm),
    ]
    per_page = 18
    total_pages = max(1, math.ceil(len(rows) / per_page))
    c = pdf_canvas.Canvas(str(target), pagesize=page_size)
    layout = _layout(backend, c, page_size, "PLN-PRAZOS", issued, logo)
    complete = sum(1 for row in rows if str(row.get("estado", "")) in {"Planeado completo", "Fluxo concluido", "Fluxo concluído"})
    partial = sum(1 for row in rows if str(row.get("estado", "")) == "Planeado parcial")
    pending = sum(1 for row in rows if str(row.get("estado", "")) == "Por planear")
    for page_index in range(total_pages):
        if page_index:
            c.showPage()
        page_no = page_index + 1
        y = layout.begin_page(
            "Prazo Final do Fluxo",
            "Previsao consolidada de conclusao das encomendas e operacoes seguintes",
            page_no,
            total_pages,
            section_label="PLANEAMENTO E CAPACIDADE",
        )
        if page_no == 1:
            y = layout.section(y, "01", "Resumo operacional", "Estado do planeamento atual")
            y = layout.metrics(y, [
                ("ENCOMENDAS", str(len(rows)), layout.palette["primary"]),
                ("FLUXO FECHADO", str(complete), layout.palette["green"]),
                ("PARCIAIS", str(partial), layout.palette["amber"]),
                ("POR PLANEAR", str(pending), layout.palette["steel"]),
            ])
        y = layout.section(y, "02", "Prazos por encomenda", "Entrega, planeamento e fim previsto")
        y = layout.table_header(y, columns)
        for local_index, row in enumerate(rows[page_index * per_page : (page_index + 1) * per_page]):
            state = str(row.get("estado", "") or "-")
            tone = "success" if state in {"Planeado completo", "Fluxo concluido", "Fluxo concluído"} else "warning" if state == "Planeado parcial" else "danger" if state == "Por planear" else ""
            y = layout.table_row(y, columns, [
                row.get("numero"), row.get("cliente"), row.get("data_entrega"), row.get("grupos_txt"),
                row.get("planeado_txt"), row.get("fim_txt"), state,
            ], height=9 * mm, index=local_index, tone=tone, font_size=6.2, max_lines=2)
        if not rows:
            layout.label_value(layout.margin, y, layout.inner_width, "ESTADO", "Sem encomendas com fluxo planeado neste momento.", height=16 * mm)
        layout.footer(page_no, total_pages, "PRAZOS DO FLUXO")
    c.save()
    return target


def render_planning_order(
    backend,
    path: str | Path,
    detail: dict[str, Any],
    flow_rows: list[dict[str, Any]],
    *,
    focus_material: str = "",
    focus_espessura: str = "",
) -> Path:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    page_size = landscape(A4)
    _brand, logo, issued = _branding(backend)
    materials = list(detail.get("materials", []) or [])
    flow_per_page, material_per_page = 16, 17
    flow_pages = max(1, math.ceil(len(flow_rows) / flow_per_page))
    material_pages = max(1, math.ceil(len(materials) / material_per_page))
    total_pages = flow_pages + material_pages
    numero = str(detail.get("numero", "") or "-")
    c = pdf_canvas.Canvas(str(target), pagesize=page_size)
    layout = _layout(backend, c, page_size, f"PLN-{numero}", issued, logo)
    flow_cols = [
        Column("Data", 25 * mm), Column("Inicio", 17 * mm), Column("Fim", 17 * mm),
        Column("Operacao", 33 * mm), Column("Recurso", 38 * mm), Column("Material", 38 * mm),
        Column("Esp.", 16 * mm, "center"), Column("Dur.", 17 * mm, "right"), Column("Origem", 39 * mm),
    ]
    material_cols = [
        Column("Material", 42 * mm), Column("Esp.", 19 * mm, "center"), Column("Operacoes", 57 * mm),
        Column("Recursos", 72 * mm), Column("Tempos", 50 * mm),
    ]
    active = sum(1 for row in flow_rows if str(row.get("source", "")) == "Ativo")
    history = len(flow_rows) - active
    total_min = sum(float(row.get("duracao_min", 0) or 0) for row in flow_rows)
    page_no = 0
    for chunk_index in range(flow_pages):
        page_no += 1
        if page_no > 1:
            c.showPage()
        y = layout.begin_page(
            f"Fluxo de Planeamento | {numero}",
            "Materiais, operacoes, recursos e blocos temporais da encomenda",
            page_no,
            total_pages,
            section_label="PLANEAMENTO DE ENCOMENDA",
        )
        if chunk_index == 0:
            y = layout.section(y, "01", "Identificacao e estado", "Resumo da encomenda")
            client = " - ".join(part for part in [str(detail.get("cliente", "") or ""), str(detail.get("cliente_nome", "") or "")] if part) or "-"
            gap = 3 * mm
            card_w = (layout.inner_width - 3 * gap) / 4
            for index, (label, value) in enumerate((("Cliente", client), ("Entrega", str(detail.get("data_entrega", "") or "-")), ("Estado", str(detail.get("estado", "") or "-")), ("Blocos", f"{active} ativos | {history} historico"))):
                layout.label_value(layout.margin + index * (card_w + gap), y, card_w, label, value, height=14 * mm)
            y -= 18 * mm
            y = layout.metrics(y, [
                ("BLOCOS", str(len(flow_rows)), layout.palette["primary"]),
                ("OPERACOES", str(len({str(r.get('operacao','')) for r in flow_rows if str(r.get('operacao',''))})), layout.palette["green"]),
                ("TEMPO TOTAL", f"{total_min:.0f} min", layout.palette["amber"]),
                ("MATERIAIS", str(len(materials)), layout.palette["steel"]),
            ])
        y = layout.section(y, "02", "Blocos de planeamento", "Sequencia operacional")
        y = layout.table_header(y, flow_cols)
        for local_index, row in enumerate(flow_rows[chunk_index * flow_per_page : (chunk_index + 1) * flow_per_page]):
            highlight = str(row.get("material", "") or "").strip() == str(focus_material or "").strip() and str(row.get("espessura", "") or "").strip() == str(focus_espessura or "").strip() and bool(focus_material)
            y = layout.table_row(y, flow_cols, [
                row.get("data"), row.get("inicio"), row.get("fim"), row.get("operacao"), row.get("recurso"),
                row.get("material"), row.get("espessura"), f"{float(row.get('duracao_min', 0) or 0):.0f}", row.get("source"),
            ], height=9 * mm, index=local_index, tone="warning" if highlight else "", font_size=5.9, max_lines=2)
        if not flow_rows:
            layout.label_value(layout.margin, y, layout.inner_width, "PLANEAMENTO", "Ainda nao existem blocos registados para esta encomenda.", height=16 * mm)
        layout.footer(page_no, total_pages, "BLOCOS DE PLANEAMENTO")
    for chunk_index in range(material_pages):
        page_no += 1
        c.showPage()
        y = layout.begin_page(f"Materiais e Operacoes | {numero}", "Recursos e tempos por grupo de material", page_no, total_pages, section_label="PLANEAMENTO DE ENCOMENDA")
        y = layout.section(y, "03", "Materiais e operacoes", "Recursos previstos e tempos")
        y = layout.table_header(y, material_cols)
        for local_index, row in enumerate(materials[chunk_index * material_per_page : (chunk_index + 1) * material_per_page]):
            highlight = str(row.get("material", "") or "").strip() == str(focus_material or "").strip() and str(row.get("espessura", "") or "").strip() == str(focus_espessura or "").strip() and bool(focus_material)
            y = layout.table_row(y, material_cols, [
                row.get("material"), row.get("espessura"), " + ".join(list(row.get("operacoes_planeamento", []) or [])) or "-",
                row.get("recursos_operacao_txt"), row.get("tempo_operacoes_txt"),
            ], height=10 * mm, index=local_index, tone="warning" if highlight else "", font_size=6.1, max_lines=2)
        if not materials:
            layout.label_value(layout.margin, y, layout.inner_width, "MATERIAIS", "Sem materiais associados.", height=16 * mm)
        layout.footer(page_no, total_pages, "MATERIAIS E OPERACOES")
    c.save()
    return target


def render_material_separation(backend, path: str | Path, rows: list[dict[str, Any]], horizon_days: int) -> Path:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    page_size = A4
    _brand, logo, issued = _branding(backend)
    alerts = list(backend.material_assistant_alert_rows(horizon_days=horizon_days) or [])

    def paginate(items: list[dict[str, Any]], first_capacity: int, next_capacity: int) -> list[list[dict[str, Any]]]:
        if not items:
            return [[]]
        pages = [items[:first_capacity]]
        offset = first_capacity
        while offset < len(items):
            pages.append(items[offset : offset + next_capacity])
            offset += next_capacity
        return pages

    separation_chunks = paginate(rows, 10, 14)
    alert_chunks = paginate(alerts, 14, 14)
    total_pages = len(separation_chunks) + len(alert_chunks)
    c = pdf_canvas.Canvas(str(target), pagesize=page_size)
    layout = _layout(backend, c, page_size, "MP-SEPARACAO", issued, logo)
    cols = [
        Column("Prio.", 14 * mm),
        Column("Posto", 24 * mm),
        Column("Encomenda", 24 * mm),
        Column("Material / lote", 35 * mm),
        Column("Qtd. / dimensao", 24 * mm),
        Column("Planeado", 27 * mm),
        Column("Acao sugerida", 40 * mm),
    ]
    alert_cols = [
        Column("Prioridade", 20 * mm),
        Column("Encomenda", 25 * mm),
        Column("Material", 36 * mm),
        Column("Necessario / disponivel", 30 * mm),
        Column("Acao recomendada", 46 * mm),
        Column("Prazo", 31 * mm),
    ]
    page_no = 0
    for chunk_index, page_rows in enumerate(separation_chunks):
        page_no += 1
        if page_no > 1:
            c.showPage()
        y = layout.begin_page(
            "Separacao de Materia-Prima",
            f"Horizonte operacional de {int(horizon_days)} dias uteis",
            page_no,
            total_pages,
            section_label="ASSISTENTE DE MATERIAL",
        )
        if chunk_index == 0:
            groups = len({(str(r.get("numero", "")), str(r.get("material_group", ""))) for r in rows})
            y = layout.section(y, "01", "Resumo de separacao", "Necessidades por posto e encomenda")
            y = layout.metrics(y, [
                ("LINHAS", str(len(rows)), layout.palette["primary"]),
                ("GRUPOS", str(groups), layout.palette["green"]),
                ("POSTOS", str(len({str(r.get('posto_trabalho','')) for r in rows})), layout.palette["amber"]),
                ("ALERTAS", str(len(alerts)), layout.palette["steel"]),
            ])
        y = layout.section(y, "02", "Lista de separacao", "Lote, quantidade e acao operacional")
        y = layout.table_header(y, cols)
        for local_index, row in enumerate(page_rows):
            tone = str(row.get("priority_tone", "") or "")
            material = str(row.get("material_group", row.get("material", "")) or "-").strip() or "-"
            lot = str(row.get("lote_sugerido", row.get("lote_atual", "")) or "").strip()
            material_lot = f"{material}\nLote {lot}" if lot and lot != "-" else material
            quantity = str(row.get("quantidade_label", backend._fmt(row.get("quantidade", 0))) or "-")
            dimension = str(row.get("dimensao", "") or "").strip()
            quantity_dimension = f"{quantity}\n{dimension}" if dimension and dimension != "-" else quantity
            planning_day = str(row.get("planeamento_dia", "") or "-").strip() or "-"
            planning_time = str(row.get("planeamento_hora", row.get("proxima_acao", "")) or "").strip()
            planning = f"{planning_day}\n{planning_time}" if planning_time and planning_time != "-" else planning_day
            y = layout.table_row(y, cols, [
                row.get("priority_label"),
                row.get("posto_trabalho"),
                row.get("numero"),
                material_lot,
                quantity_dimension,
                planning,
                row.get("acao_sugerida"),
            ], height=13 * mm, index=local_index, tone=tone if tone in {"danger", "warning", "success"} else "", font_size=5.8, max_lines=3)
        if not rows:
            layout.label_value(layout.margin, y, layout.inner_width, "SEPARACAO", "Sem necessidades no horizonte atual.", height=16 * mm)
        layout.footer(page_no, total_pages, "SEPARACAO DE MATERIA-PRIMA")
    for chunk_index, page_rows in enumerate(alert_chunks):
        page_no += 1
        c.showPage()
        y = layout.begin_page("Sugestoes e Alertas de Material", "Acoes de compra, cativacao e regularizacao", page_no, total_pages, section_label="ASSISTENTE DE MATERIAL")
        y = layout.section(y, "03", "Alertas e sugestoes", "Decisoes recomendadas")
        y = layout.table_header(y, alert_cols)
        for local_index, row in enumerate(page_rows):
            needed = str(row.get("quantidade_necessaria", row.get("necessidade", "-")) or "-")
            available = str(row.get("quantidade_disponivel", row.get("disponivel", "-")) or "-")
            y = layout.table_row(y, alert_cols, [
                row.get("priority_label", row.get("prioridade")),
                row.get("numero"),
                row.get("material_group", row.get("material")),
                f"Nec. {needed}\nDisp. {available}",
                row.get("acao_sugerida", row.get("sugestao")),
                row.get("delivery", row.get("prazo")),
            ], height=13 * mm, index=local_index, tone=str(row.get("priority_tone", "") or ""), font_size=5.9, max_lines=3)
        if not alerts:
            layout.label_value(layout.margin, y, layout.inner_width, "ALERTAS", "Sem alertas de material pendentes.", height=16 * mm)
        layout.footer(page_no, total_pages, "ALERTAS DE MATERIAL")
    c.save()
    return target


def render_material_history(backend, path: str | Path, rows: list[dict[str, Any]], title: str) -> Path:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    page_size = landscape(A4)
    _brand, logo, issued = _branding(backend)
    first_page_capacity = 11
    next_page_capacity = 14
    if len(rows) <= first_page_capacity:
        total_pages = 1
    else:
        total_pages = 1 + math.ceil((len(rows) - first_page_capacity) / next_page_capacity)
    columns = [
        Column("Data", 35 * mm), Column("Acao", 24 * mm), Column("Operador", 28 * mm),
        Column("Materia-prima", 44 * mm), Column("Esp.", 15 * mm, "center"),
        Column("Dimensao", 28 * mm), Column("Lote", 30 * mm), Column("Qtd.", 16 * mm, "right"),
        Column("Reserv.", 18 * mm, "right"), Column("Detalhes", 42 * mm),
    ]
    c = pdf_canvas.Canvas(str(target), pagesize=page_size)
    layout = _layout(backend, c, page_size, "MP-HISTORICO", issued, logo)
    row_offset = 0
    for page_index in range(total_pages):
        if page_index:
            c.showPage()
        page_no = page_index + 1
        y = layout.begin_page(title or "Historico de Materia-Prima", "Movimentos, reservas, lotes e rastreabilidade de stock", page_no, total_pages, section_label="CONTROLO DE STOCK")
        if page_index == 0:
            y = layout.section(y, "01", "Resumo do historico", "Registos de auditoria de stock")
            y = layout.metrics(y, [
                ("MOVIMENTOS", str(len(rows)), layout.palette["primary"]),
                ("OPERADORES", str(len({str(r.get('operador','')) for r in rows if str(r.get('operador',''))})), layout.palette["green"]),
                ("MATERIAIS", str(len({str(r.get('material_id','')) for r in rows if str(r.get('material_id',''))})), layout.palette["amber"]),
                ("LOTES", str(len({str(r.get('lote','')) for r in rows if str(r.get('lote','')) not in {'','-'}})), layout.palette["steel"]),
            ])
        y = layout.section(y, "02", "Movimentos de materia-prima", "Ordem cronologica decrescente")
        y = layout.table_header(y, columns)
        capacity = first_page_capacity if page_index == 0 else next_page_capacity
        page_rows = rows[row_offset : row_offset + capacity]
        row_offset += len(page_rows)
        for local_index, row in enumerate(page_rows):
            action = str(row.get("acao", "") or "")
            tone = "danger" if any(token in action.casefold() for token in ("baixa", "remov", "anul")) else "success" if any(token in action.casefold() for token in ("entrada", "rececao")) else ""
            y = layout.table_row(y, columns, [
                row.get("data"), action, row.get("operador"), row.get("material"), row.get("espessura"),
                row.get("dimensao"), row.get("lote"), row.get("qtd"), row.get("reservado"), row.get("detalhes"),
            ], height=9 * mm, index=local_index, tone=tone, font_size=5.7, max_lines=2)
        if not rows:
            layout.label_value(layout.margin, y, layout.inner_width, "HISTORICO", "Sem movimentos registados.", height=16 * mm)
        layout.footer(page_no, total_pages, "HISTORICO DE MATERIA-PRIMA")
    c.save()
    return target


def render_quality_dossier(backend, path: str | Path) -> Path:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    page_size = A4
    _brand, logo, issued = _branding(backend)
    summary = dict(backend.quality_summary() or {})
    checklist = list(backend.quality_iso_checklist() or [])
    nc_rows = list(backend.quality_nc_rows("", "Todos") or [])
    docs = list(backend.quality_document_rows("") or [])
    audit = list(backend.audit_rows("", limit=120) or [])
    sections = [
        ("Nao Conformidades", nc_rows, 14),
        ("Documentos Controlados", docs, 15),
        ("Auditoria e Rastreabilidade", audit, 14),
    ]
    total_pages = 1 + sum(max(1, math.ceil(len(rows) / capacity)) for _title, rows, capacity in sections)
    c = pdf_canvas.Canvas(str(target), pagesize=page_size)
    layout = _layout(backend, c, page_size, "QLT-ISO9001", issued, logo)
    y = layout.begin_page("Dossier da Qualidade ISO 9001", "Indicadores, conformidade, documentos e rastreabilidade", 1, total_pages, status="DOCUMENTO CONTROLADO", section_label="SISTEMA DE GESTAO DA QUALIDADE")
    y = layout.section(y, "01", "Indicadores do sistema", "Estado consolidado")
    y = layout.metrics(y, [
        ("NC ABERTAS", str(summary.get("open_nc", 0)), layout.palette["primary"]),
        ("FORA DE PRAZO", str(summary.get("overdue_nc", 0)), layout.palette["amber"]),
        ("DOCUMENTOS", str(summary.get("documents", 0)), layout.palette["green"]),
        ("AUDITORIA", str(summary.get("audit_events", 0)), layout.palette["steel"]),
    ])
    y = layout.section(y, "02", "Checklist ISO 9001", "Evidencia e estado")
    check_cols = [Column("Area", 52 * mm), Column("Estado", 25 * mm), Column("Evidencia", layout.inner_width - 77 * mm)]
    y = layout.table_header(y, check_cols)
    for index, row in enumerate(checklist):
        state = str(row.get("estado", "") or "-")
        tone = "success" if state == "OK" else "warning" if state in {"Atencao", "Pendente"} else ""
        y = layout.table_row(y, check_cols, [row.get("area"), state, row.get("evidencia")], height=13 * mm, index=index, tone=tone, font_size=6.0, max_lines=3)
    layout.footer(1, total_pages, "RESUMO E CHECKLIST")

    page_no = 1
    for section_index, (title, rows, capacity) in enumerate(sections, start=3):
        chunks = max(1, math.ceil(len(rows) / capacity))
        for chunk in range(chunks):
            page_no += 1
            c.showPage()
            y = layout.begin_page(title, "Registos ligados ao sistema de gestao da qualidade", page_no, total_pages, section_label="SISTEMA DE GESTAO DA QUALIDADE")
            y = layout.section(y, f"{section_index:02d}", title, f"Folha {chunk + 1}/{chunks}")
            if title == "Nao Conformidades":
                columns = [Column("ID", 28 * mm), Column("Estado", 25 * mm), Column("Gravidade", 25 * mm), Column("Entidade / Ref.", 48 * mm), Column("Descricao", layout.inner_width - 126 * mm)]
                values = lambda row: [row.get("id"), row.get("estado"), row.get("gravidade"), row.get("entidade_label", row.get("referencia")), row.get("descricao")]
            elif title == "Documentos Controlados":
                columns = [Column("ID", 27 * mm), Column("Tipo", 31 * mm), Column("Versao", 18 * mm), Column("Titulo", 66 * mm), Column("Ligacao", layout.inner_width - 142 * mm)]
                values = lambda row: [row.get("id"), row.get("tipo"), row.get("versao"), row.get("titulo"), f"{row.get('entidade', row.get('entidade_tipo',''))} {row.get('referencia', row.get('entidade_id',''))}"]
            else:
                columns = [Column("Data", 31 * mm), Column("Utilizador", 28 * mm), Column("Acao", 45 * mm), Column("Entidade", 34 * mm), Column("Resumo", layout.inner_width - 138 * mm)]
                values = lambda row: [str(row.get("created_at", ""))[:19], row.get("user"), row.get("action"), f"{row.get('entity_type','')} {row.get('entity_id','')}", row.get("summary")]
            y = layout.table_header(y, columns)
            for local_index, row in enumerate(rows[chunk * capacity : (chunk + 1) * capacity]):
                tone = "danger" if str(row.get("estado", "")).lower() in {"aberta", "vencida"} else ""
                y = layout.table_row(y, columns, values(row), height=13 * mm, index=local_index, tone=tone, font_size=5.7, max_lines=3)
            if not rows:
                layout.label_value(layout.margin, y, layout.inner_width, "REGISTOS", "Sem registos nesta seccao.", height=16 * mm)
            layout.footer(page_no, total_pages, title.upper())
    c.save()
    return target
