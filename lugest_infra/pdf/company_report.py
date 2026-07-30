from __future__ import annotations

import math
from pathlib import Path
from typing import Any

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas as pdf_canvas

from .dossier_layout import Column, DossierLayout


def _money(value: Any) -> str:
    try:
        amount = float(value or 0)
    except (TypeError, ValueError):
        amount = 0.0
    return f"{amount:,.2f} EUR".replace(",", "X").replace(".", ",").replace("X", ".")


def _qty(value: Any) -> str:
    try:
        amount = float(value or 0)
    except (TypeError, ValueError):
        amount = 0.0
    return f"{amount:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _chunks(rows: list[dict[str, Any]], size: int) -> list[list[dict[str, Any]]]:
    return [rows[index : index + size] for index in range(0, len(rows), size)] or [[]]


def _new_layout(backend, canvas_obj, page_size, issued: str, logo: Path | None) -> DossierLayout:
    branding = dict(backend.branding_settings() or {})
    return DossierLayout(
        canvas_obj,
        page_size,
        primary=branding.get("primary_color", "#00A6A6"),
        logo_path=logo,
        logo_draw=backend._draw_operator_logo_plate,
        issued_at=issued,
        document_code="REL-VALOR-EMPRESA",
    )


def render_company_value_report(backend, path: str | Path, payload: dict[str, Any]) -> Path:
    """Render the consolidated management report using the shared dossier geometry."""

    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    page_size = A4
    branding = dict(backend.branding_settings() or {})
    logo_text = str(branding.get("logo_path", "") or "").strip()
    logo = Path(logo_text) if logo_text and Path(logo_text).exists() else None
    issued = str(backend.desktop_main.now_iso() or "").replace("T", " ")[:19]

    materials = list(payload.get("top_materias", []) or [])
    products = list(payload.get("top_produtos", []) or [])
    # The dashboard keeps full detail in these keys; retain compatibility with older payloads.
    materials = list(payload.get("materias", materials) or materials)
    products = list(payload.get("produtos", products) or products)
    purchases = list(payload.get("compras", []) or [])
    commitments = list(payload.get("compromissos", []) or [])
    billing = list(payload.get("billing_rows", []) or [])

    pages: list[tuple[str, list[dict[str, Any]]]] = [("summary", []), ("analysis", [])]
    pages.extend(("materials", chunk) for chunk in _chunks(materials, 14))
    pages.extend(("products", chunk) for chunk in _chunks(products, 14))
    pages.extend(("purchases", chunk) for chunk in _chunks(purchases, 14))
    if commitments:
        pages.extend(("commitments", chunk) for chunk in _chunks(commitments, 14))
    if billing:
        pages.extend(("billing", chunk) for chunk in _chunks(billing, 14))

    total_pages = len(pages)
    canvas_obj = pdf_canvas.Canvas(str(target), pagesize=page_size)
    layout = _new_layout(backend, canvas_obj, page_size, issued, logo)
    summary = dict(payload.get("executive_summary", {}) or {})
    period = str(summary.get("periodo", payload.get("selected_year", "Todos os anos")) or "Todos os anos")

    for page_index, (kind, rows) in enumerate(pages):
        if page_index:
            canvas_obj.showPage()
        page_no = page_index + 1
        titles = {
            "summary": ("Relatório de Valor Empresarial", "Património, compras, compromissos, vendas e recebimentos"),
            "analysis": ("Análise de Compras e Cobertura", "Fornecedores, evolução mensal e qualidade da informação"),
            "materials": ("Inventário de Matéria-Prima", "Valorização física ao preço unitário atual"),
            "products": ("Inventário de Produto Acabado", "Valorização física ao preço unitário atual"),
            "purchases": ("Compras Recebidas", "Linhas efetivamente rececionadas no período selecionado"),
            "commitments": ("Compromissos de Compra", "Linhas aprovadas ainda por rececionar"),
            "billing": ("Vendas e Cobranças", "Vendido, faturado, recebido e saldo por registo"),
        }
        title, subtitle = titles[kind]
        y = layout.begin_page(
            title,
            f"{subtitle} | {period}",
            page_no,
            total_pages,
            section_label="GESTÃO E CONTROLO EMPRESARIAL",
        )

        if kind == "summary":
            y = layout.section(y, "01", "Posição executiva", "Valores registados no sistema")
            executive_metrics = [
                ("PATRIMÓNIO EM STOCK", _money(summary.get("stock_total")), layout.palette["primary"]),
                ("COMPRAS RECEBIDAS", _money(summary.get("compras_total")), layout.palette["amber"]),
                ("COMPROMISSOS ABERTOS", _money(summary.get("compromissos_total")), layout.palette["steel"]),
                ("SALDO DE CLIENTES", _money(summary.get("saldo_clientes")), layout.palette["danger"]),
                ("VENDAS REGISTADAS", _money(summary.get("vendido_total")), layout.palette["green"]),
                ("FATURADO", _money(summary.get("faturado_total")), layout.palette["primary"]),
                ("RECEBIDO", _money(summary.get("recebido_total")), layout.palette["green"]),
                ("STOCK RESERVADO", _money(summary.get("stock_reservado")), layout.palette["amber"]),
            ]
            for metric_index in range(0, len(executive_metrics), 2):
                y = layout.metrics(y, executive_metrics[metric_index : metric_index + 2], height=16 * mm)
            value_cols = [
                Column("Componente", 34 * mm), Column("Valor físico", 32 * mm, "right"),
                Column("Disponível", 32 * mm, "right"), Column("Reservado", 32 * mm, "right"),
                Column("Critério", 58 * mm),
            ]
            y = layout.section(y, "02", "Composição do património", "Stock atual, disponível e reservado")
            y = layout.table_header(y, value_cols)
            for index, row in enumerate(list(payload.get("value_rows", []) or [])):
                y = layout.table_row(y, value_cols, [
                    row.get("componente"), _money(row.get("valor")), _money(row.get("disponivel")),
                    _money(row.get("reservado")), row.get("criterio"),
                ], index=index, height=9 * mm, font_size=6.1)
            y = layout.section(y - 3 * mm, "03", "Âmbito da leitura", "Posição atual e movimento do período")
            layout.label_value(
                layout.margin,
                y,
                layout.inner_width,
                "INTERPRETAÇÃO",
                "O património representa o stock físico ao preço atual. Compras, vendas, faturação e recebimentos são movimentos do período selecionado e não são somados ao stock.",
                height=15 * mm,
                max_lines=2,
            )

        elif kind == "analysis":
            suppliers = list(payload.get("compras_por_fornecedor", []) or [])
            months = list(payload.get("compras_por_mes", []) or [])
            combined = max(len(suppliers), len(months), 1)
            flow_cols = [Column("Indicador", 38 * mm), Column("Valor", 32 * mm, "right"), Column("Leitura", 118 * mm)]
            y = layout.section(y, "03", "Fluxo comercial e financeiro", "Sem confundir stock com movimento do período")
            y = layout.table_header(y, flow_cols)
            for index, row in enumerate(list(payload.get("flow_rows", []) or [])):
                y = layout.table_row(y, flow_cols, [
                    row.get("indicador"), _money(row.get("valor")), row.get("leitura"),
                ], index=index, height=6.5 * mm, font_size=5.8)
            cols = [
                Column("Fornecedor", 58 * mm), Column("Total", 28 * mm, "right"),
                Column("Mês", 30 * mm), Column("Total", 28 * mm, "right"),
                Column("Peso no período", 44 * mm),
            ]
            y = layout.section(y - 3 * mm, "04", "Distribuição das compras", "Fornecedor e mês de receção")
            y = layout.table_header(y, cols)
            purchases_total = float(summary.get("compras_total", 0) or 0)
            for index in range(combined):
                supplier = suppliers[index] if index < len(suppliers) else {}
                month = months[index] if index < len(months) else {}
                supplier_total = float(supplier.get("total", 0) or 0)
                weight = (supplier_total / purchases_total * 100.0) if purchases_total > 0 else 0.0
                y = layout.table_row(y, cols, [
                    supplier.get("fornecedor"), _money(supplier_total) if supplier else "-",
                    month.get("mes"), _money(month.get("total")) if month else "-",
                    f"{weight:.1f}% das compras" if supplier else "-",
                ], index=index, height=8 * mm, font_size=6.0)
            y = layout.section(y - 3 * mm, "05", "Controlo da informação", "Critérios e pontos de atenção")
            gap = 3 * mm
            card_w = (layout.inner_width - 2 * gap) / 3
            controls = [
                ("REFERÊNCIAS SEM PREÇO", str(int(summary.get("referencias_sem_preco", 0) or 0)), "Devem ser valorizadas para fechar o património."),
                ("LINHAS POR RECEBER", str(len(commitments)), f"Compromisso aberto de {_money(summary.get('compromissos_total'))}."),
                ("PERÍODO DO RELATÓRIO", period, "O stock é atual; movimentos respeitam o filtro temporal."),
            ]
            for index, (label, value, note) in enumerate(controls):
                x = layout.margin + index * (card_w + gap)
                layout.label_value(x, y, card_w, label, f"{value} | {note}", height=21 * mm, max_lines=3)
            y -= 25 * mm
            layout.label_value(
                layout.margin,
                y,
                layout.inner_width,
                "METODOLOGIA",
                "Stock = quantidade física x preço unitário atual. Compras = quantidades rececionadas x preço da linha. "
                "Compromissos = quantidade aprovada ainda por receber x preço da linha. Valores comerciais provêm do módulo de faturação. "
                "Este documento é uma leitura de gestão e não substitui a contabilidade oficial.",
                height=18 * mm,
                max_lines=3,
            )

        elif kind == "materials":
            cols = [
                Column("ID", 21 * mm), Column("Material", 29 * mm), Column("Esp.", 12 * mm, "center"),
                Column("Tipo / dimensão", 38 * mm), Column("Físico", 17 * mm, "right"),
                Column("Reserv.", 17 * mm, "right"), Column("Dispon.", 17 * mm, "right"),
                Column("Preço", 18 * mm, "right"), Column("Valor", 19 * mm, "right"),
            ]
            y = layout.section(y, "06", "Matéria-prima valorizada", f"{len(materials)} referências com stock positivo")
            y = layout.table_header(y, cols)
            for index, row in enumerate(rows):
                dimension = " | ".join(value for value in [str(row.get("tipo", "") or ""), str(row.get("dimensao", "") or "")] if value) or "-"
                y = layout.table_row(y, cols, [
                    row.get("id"), row.get("material"), row.get("espessura"), dimension,
                    _qty(row.get("qty")), _qty(row.get("reservado")), _qty(row.get("disponivel")),
                    _money(row.get("preco_unid")), _money(row.get("valor")),
                ], index=index, height=9 * mm, font_size=5.8)

        elif kind == "products":
            cols = [
                Column("Código", 24 * mm), Column("Descrição", 52 * mm), Column("Categoria", 28 * mm),
                Column("Físico", 16 * mm, "right"), Column("Reserv.", 16 * mm, "right"),
                Column("Dispon.", 16 * mm, "right"), Column("Preço", 18 * mm, "right"), Column("Valor", 18 * mm, "right"),
            ]
            y = layout.section(y, "07", "Produto acabado valorizado", f"{len(products)} referências com stock positivo")
            y = layout.table_header(y, cols)
            for index, row in enumerate(rows):
                y = layout.table_row(y, cols, [
                    row.get("codigo"), row.get("descricao"), row.get("categoria"), _qty(row.get("qty")),
                    _qty(row.get("reservado")), _qty(row.get("disponivel")), _money(row.get("preco_unid")), _money(row.get("valor")),
                ], index=index, height=9 * mm, font_size=5.8)

        elif kind == "purchases":
            cols = [
                Column("Data", 20 * mm), Column("NE", 24 * mm), Column("Fornecedor", 34 * mm),
                Column("Artigo", 54 * mm), Column("Qtd.", 15 * mm, "right"),
                Column("Preço", 18 * mm, "right"), Column("Total", 23 * mm, "right"),
            ]
            y = layout.section(y, "08", "Histórico de receções", f"{len(purchases)} linhas no período")
            y = layout.table_header(y, cols)
            for index, row in enumerate(rows):
                y = layout.table_row(y, cols, [
                    row.get("data"), row.get("ne"), row.get("fornecedor"), row.get("artigo"),
                    _qty(row.get("qtd")), _money(row.get("preco")), _money(row.get("total")),
                ], index=index, height=9 * mm, font_size=5.8)

        elif kind == "commitments":
            cols = [
                Column("NE", 24 * mm), Column("Fornecedor", 38 * mm), Column("Artigo", 54 * mm),
                Column("Pendente", 18 * mm, "right"), Column("Preço", 18 * mm, "right"),
                Column("Compromisso", 22 * mm, "right"), Column("Entrega", 14 * mm),
            ]
            y = layout.section(y, "09", "Linhas aprovadas por receber", f"{len(commitments)} compromissos abertos")
            y = layout.table_header(y, cols)
            for index, row in enumerate(rows):
                y = layout.table_row(y, cols, [
                    row.get("ne"), row.get("fornecedor"), row.get("artigo"), _qty(row.get("qtd_pendente")),
                    _money(row.get("preco")), _money(row.get("total")), row.get("entrega"),
                ], index=index, height=9 * mm, font_size=5.8, tone="warning")

        elif kind == "billing":
            cols = [
                Column("Registo", 22 * mm), Column("Cliente", 43 * mm), Column("Data", 18 * mm),
                Column("Vendido", 22 * mm, "right"), Column("Faturado", 22 * mm, "right"),
                Column("Recebido", 22 * mm, "right"), Column("Saldo", 22 * mm, "right"),
                Column("Estado", 17 * mm),
            ]
            y = layout.section(y, "10", "Posição comercial por registo", f"{len(billing)} registos no período")
            y = layout.table_header(y, cols)
            for index, row in enumerate(rows):
                balance = float(row.get("saldo", 0) or 0)
                y = layout.table_row(y, cols, [
                    row.get("record_number", row.get("numero")), row.get("cliente", row.get("cliente_nome")),
                    row.get("data_venda"), _money(row.get("vendido")), _money(row.get("faturado")),
                    _money(row.get("recebido")), _money(balance), row.get("estado_pagamento"),
                ], index=index, height=9 * mm, font_size=5.7, tone="warning" if balance > 0 else "success")

        layout.footer(page_no, total_pages, title.upper())

    canvas_obj.save()
    return target


def render_pulse_performance_report(
    backend,
    path: str | Path,
    payload: dict[str, Any],
    scope: dict[str, Any] | None = None,
) -> Path:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    page_size = A4
    branding = dict(backend.branding_settings() or {})
    logo_text = str(branding.get("logo_path", "") or "").strip()
    logo = Path(logo_text) if logo_text and Path(logo_text).exists() else None
    issued = str(backend.desktop_main.now_iso() or "").replace("T", " ")[:19]
    summary = dict(payload.get("summary", {}) or {})
    running = list(payload.get("running", []) or [])
    stops = list(payload.get("top_stops", []) or [])
    history = list(payload.get("history", []) or [])
    scope = dict(scope or {})
    scope_text = " | ".join(
        f"{label} {scope.get(key)}"
        for key, label in (("period", "Período"), ("year", "Ano"), ("origin", "Origem"), ("view", "Visão"))
        if str(scope.get(key, "") or "").strip()
    ) or "Âmbito atual"
    pages: list[tuple[str, list[dict[str, Any]]]] = [("summary", [])]
    if running:
        pages.extend(("running", chunk) for chunk in _chunks(running, 16))
    if history:
        pages.extend(("history", chunk) for chunk in _chunks(history, 17))
    total_pages = len(pages)
    canvas_obj = pdf_canvas.Canvas(str(target), pagesize=page_size)
    layout = DossierLayout(
        canvas_obj,
        page_size,
        primary=branding.get("primary_color", "#00A6A6"),
        logo_path=logo,
        logo_draw=backend._draw_operator_logo_plate,
        issued_at=issued,
        document_code="PULSE-OEE",
    )
    for page_index, (kind, rows) in enumerate(pages):
        if page_index:
            canvas_obj.showPage()
        page_no = page_index + 1
        title = "Pulse | Desempenho Industrial" if kind == "summary" else ("Pulse | Peças em Curso" if kind == "running" else "Pulse | Histórico Consolidado")
        y = layout.begin_page(title, scope_text, page_no, total_pages, section_label="INDÚSTRIA 4.0 E OEE")
        if kind == "summary":
            y = layout.section(y, "01", "Indicadores de desempenho", "OEE e componentes do período")
            pulse_metrics = [
                ("OEE", f"{float(summary.get('oee', 0) or 0):.1f}%", layout.palette["primary"]),
                ("DISPONIBILIDADE", f"{float(summary.get('disponibilidade', 0) or 0):.1f}%", layout.palette["green"]),
                ("PERFORMANCE", f"{float(summary.get('performance', 0) or 0):.1f}%", layout.palette["amber"]),
                ("QUALIDADE", f"{float(summary.get('qualidade', 0) or 0):.1f}%", layout.palette["steel"]),
                ("PARAGENS", f"{float(summary.get('paragens_min', 0) or 0):.1f} min", layout.palette["danger"]),
                ("DESVIO MÁXIMO", f"{float(summary.get('desvio_max_min', 0) or 0):.1f} min", layout.palette["amber"]),
                ("PEÇAS EM CURSO", str(int(summary.get("pecas_em_curso", 0) or 0)), layout.palette["primary"]),
                ("FORA DE TEMPO", str(int(summary.get("pecas_fora_tempo", 0) or 0)), layout.palette["danger"]),
            ]
            for metric_index in range(0, len(pulse_metrics), 2):
                y = layout.metrics(y, pulse_metrics[metric_index : metric_index + 2], height=16 * mm)
            stop_cols = [
                Column("Causa", 50 * mm), Column("Encomenda", 28 * mm), Column("Operador", 35 * mm),
                Column("Ocorrências", 25 * mm, "right"), Column("Minutos", 25 * mm, "right"), Column("Peso", 25 * mm, "right"),
            ]
            y = layout.section(y, "02", "Principais causas de paragem", "Concentração de perdas registadas")
            y = layout.table_header(y, stop_cols)
            total_stop = sum(float(row.get("minutos", 0) or 0) for row in stops)
            for index, row in enumerate(stops[:8]):
                minutes = float(row.get("minutos", 0) or 0)
                y = layout.table_row(y, stop_cols, [
                    row.get("causa"), row.get("encomenda"), row.get("operador"), row.get("ocorrencias"),
                    f"{minutes:.1f}", f"{(minutes / total_stop * 100):.1f}%" if total_stop else "0.0%",
                ], index=index, height=7.5 * mm, font_size=6.0, tone="danger" if index == 0 and minutes > 0 else "")
            y = layout.section(y - 3 * mm, "03", "Leitura e ação", "Mensagem gerada pelo Pulse")
            layout.label_value(layout.margin, y, layout.inner_width, "ALERTAS", str(summary.get("alerts", "-") or "-"), height=16 * mm, max_lines=3)
        elif kind == "running":
            cols = [
                Column("Encomenda", 28 * mm), Column("Peça", 50 * mm), Column("Operação", 35 * mm),
                Column("Operador", 30 * mm), Column("Tempo", 22 * mm, "right"), Column("Desvio", 23 * mm, "right"),
            ]
            y = layout.section(y, "04", "Trabalho em execução", f"{len(running)} peças em curso")
            y = layout.table_header(y, cols)
            for index, row in enumerate(rows):
                delta = float(row.get("delta_min", 0) or 0)
                y = layout.table_row(y, cols, [
                    row.get("encomenda"), row.get("peca"), row.get("operacao"), row.get("operador"),
                    f"{float(row.get('elapsed_min', 0) or 0):.1f} min", f"{delta:.1f} min",
                ], index=index, height=9 * mm, font_size=6.0, tone="danger" if delta > 0 else "success")
        else:
            cols = [
                Column("Encomenda", 40 * mm), Column("Operações", 23 * mm, "right"),
                Column("Tempo real", 32 * mm, "right"), Column("Planeado", 32 * mm, "right"),
                Column("Desvio", 30 * mm, "right"), Column("Leitura", 31 * mm),
            ]
            y = layout.section(y, "05", "Histórico de execução", f"{len(history)} encomendas consolidadas")
            y = layout.table_header(y, cols)
            for index, row in enumerate(rows):
                delta = float(row.get("delta_min", 0) or 0)
                y = layout.table_row(y, cols, [
                    row.get("encomenda"), row.get("ops"), f"{float(row.get('elapsed_min', 0) or 0):.1f} min",
                    f"{float(row.get('plan_min', 0) or 0):.1f} min", f"{delta:.1f} min", "Desvio" if delta > 0 else "Dentro do plano",
                ], index=index, height=9 * mm, font_size=6.0, tone="danger" if delta > 0 else "success")
        layout.footer(page_no, total_pages, title.upper())
    canvas_obj.save()
    return target
