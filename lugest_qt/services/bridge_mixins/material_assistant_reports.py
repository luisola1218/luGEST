from __future__ import annotations

import math
from datetime import date, timedelta
from lugest_infra.pdf.text import (
    clip_text as _pdf_clip_text,
    mix_hex as _pdf_mix_hex,
    wrap_text as _pdf_wrap_text,
)
from typing import Any


class MaterialAssistantReportsBackendMixin:
    """Legacy adapter for material assistant reports; see BACKEND_GUIDE.md."""

    def _material_assistant_planning_week_start(self, rows: list[dict[str, Any]] | None = None) -> date:
        candidates: list[date] = []
        for row in list(rows or []) or []:
            raw = str((row or {}).get("planeamento_dia_iso", "") or "").strip()
            if not raw or raw == "9999-99-99":
                continue
            try:
                candidates.append(date.fromisoformat(raw))
            except Exception:
                continue
        anchor = min(candidates) if candidates else date.today()
        return self._planning_week_start(anchor)

    def _material_assistant_append_planning_page(
        self,
        canvas_obj: Any,
        width: float,
        height: float,
        rows: list[dict[str, Any]] | None = None,
    ) -> None:
        from reportlab.lib import colors

        planning_operation = "Corte Laser"
        week_start = self._material_assistant_planning_week_start(rows)
        week_dates = [week_start + timedelta(days=index) for index in range(6)]
        start_min, end_min, slot = self._planning_grid_metrics()
        total_slots = max(1, int((end_min - start_min) / max(1, slot)))
        margin = 20
        top_margin = 60
        time_w = 66
        footer_box_h = 58
        cols = 6
        usable_w = width - (margin * 2)
        grid_w = usable_w - time_w
        col_w = max(60, grid_w / cols)
        grid_h = height - top_margin - margin - footer_box_h - 10
        row_h = max(18, grid_h / (total_slots + 1))
        dias = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sab"]

        def yinv(y: float) -> float:
            return height - y

        def draw_block_text(cx: float, cy: float, lines: list[Any], box_w: float, box_h: float) -> None:
            raw_specs = []
            for line in list(lines or []):
                if isinstance(line, dict):
                    text = str(line.get("text", "") or "").strip()
                    role = str(line.get("role", "body") or "body").strip()
                else:
                    text = str(line or "").strip()
                    role = "body"
                if text:
                    raw_specs.append({"text": text, "role": role})
            if not raw_specs:
                raw_specs = [{"text": "-", "role": "body"}]
            text_w = max(28.0, box_w - 10)
            for font_size in (9.2, 8.2, 7.2, 6.4, 5.8):
                line_h = font_size + 1.2
                max_lines = max(1, int((box_h - 6) // line_h))
                wrapped: list[dict[str, str]] = []
                for spec in raw_specs:
                    role = str(spec.get("role", "body") or "body")
                    font_name = "Helvetica-Bold" if role in ("title", "time") else "Helvetica"
                    for part in _pdf_wrap_text(spec.get("text", ""), font_name, font_size, text_w, max_lines=2):
                        part_txt = str(part or "").strip()
                        if part_txt:
                            wrapped.append({"text": part_txt, "role": role})
                if len(wrapped) <= max_lines:
                    total_h = len(wrapped) * line_h
                    start_y = cy - (total_h / 2.0) + (line_h / 2.0)
                    canvas_obj.setFillColor(colors.black)
                    for idx, line in enumerate(wrapped):
                        role = str(line.get("role", "body") or "body")
                        canvas_obj.setFont("Helvetica-Bold" if role in ("title", "time") else "Helvetica", font_size)
                        canvas_obj.drawCentredString(cx, yinv(start_y + (idx * line_h)), str(line.get("text", "") or "-"))
                    return

            font_size = 5.8
            line_h = font_size + 1.1
            max_lines = max(1, int((box_h - 6) // line_h))
            wrapped: list[dict[str, str]] = []
            for spec in raw_specs:
                role = str(spec.get("role", "body") or "body")
                font_name = "Helvetica-Bold" if role in ("title", "time") else "Helvetica"
                for part in _pdf_wrap_text(spec.get("text", ""), font_name, font_size, text_w, max_lines=2):
                    part_txt = str(part or "").strip()
                    if part_txt:
                        wrapped.append({"text": part_txt, "role": role})
            wrapped = wrapped[:max_lines]
            if wrapped:
                last = dict(wrapped[-1])
                font_name = "Helvetica-Bold" if last.get("role") in ("title", "time") else "Helvetica"
                text = str(last.get("text", "") or "").strip()
                while text and len(_pdf_wrap_text(f"{text}...", font_name, font_size, text_w, max_lines=1)) != 1:
                    text = text[:-1].rstrip()
                last["text"] = f"{text}..." if text else "..."
                wrapped[-1] = last
            total_h = len(wrapped or [{"text": "-", "role": "body"}]) * line_h
            start_y = cy - (total_h / 2.0) + (line_h / 2.0)
            canvas_obj.setFillColor(colors.black)
            for idx, line in enumerate(wrapped or [{"text": "-", "role": "body"}]):
                role = str(line.get("role", "body") or "body")
                canvas_obj.setFont("Helvetica-Bold" if role in ("title", "time") else "Helvetica", font_size)
                canvas_obj.drawCentredString(cx, yinv(start_y + (idx * line_h)), str(line.get("text", "") or "-"))

        canvas_obj.showPage()
        canvas_obj.setStrokeColor(colors.HexColor("#c7ccd6"))
        canvas_obj.rect(margin, yinv(height - margin), width - (margin * 2), height - (margin * 2), stroke=1, fill=0)

        canvas_obj.setFillColor(colors.HexColor("#0f172a"))
        canvas_obj.setFont("Helvetica-Bold", 15)
        canvas_obj.drawString(margin + 80, yinv(margin + 20), "Planeamento associado")
        canvas_obj.setFont("Helvetica", 9)
        canvas_obj.setFillColor(colors.HexColor("#475569"))
        canvas_obj.drawString(margin + 80, yinv(margin + 34), planning_operation)
        canvas_obj.drawRightString(
            width - margin - 4,
            yinv(margin + 20),
            f"Semana: {week_dates[0].strftime('%d/%m/%Y')} - {week_dates[-1].strftime('%d/%m/%Y')}",
        )
        canvas_obj.line(margin, yinv(55), width - margin, yinv(55))

        canvas_obj.setFont("Helvetica-Bold", 8)
        for col_idx, day in enumerate(week_dates):
            x0 = margin + time_w + (col_idx * col_w)
            canvas_obj.setFillColor(colors.HexColor("#e8eefc"))
            canvas_obj.setStrokeColor(colors.HexColor("#c4d2f3"))
            canvas_obj.rect(x0, yinv(top_margin + row_h), col_w, row_h, stroke=1, fill=1)
            canvas_obj.setFillColor(colors.HexColor("#0f172a"))
            canvas_obj.drawCentredString(x0 + (col_w / 2.0), yinv(top_margin + (row_h / 2.0)), f"{dias[col_idx]} {day.strftime('%d/%m')}")

        canvas_obj.setFillColor(colors.HexColor("#f8fafc"))
        canvas_obj.rect(margin, yinv(top_margin + ((total_slots + 1) * row_h)), time_w, total_slots * row_h, stroke=1, fill=1)
        for row_idx in range(total_slots):
            slot_start = start_min + (row_idx * slot)
            slot_end = slot_start + slot
            y0 = top_margin + ((row_idx + 1) * row_h)
            hhmm = str(self.desktop_main.minutes_to_time(slot_start) or "")
            if hhmm.endswith(":00"):
                canvas_obj.setFont("Helvetica-Bold", 7.5)
                canvas_obj.setFillColor(colors.HexColor("#334155"))
            else:
                canvas_obj.setFont("Helvetica", 6.8)
                canvas_obj.setFillColor(colors.HexColor("#64748b"))
            canvas_obj.drawCentredString(margin + (time_w / 2.0), yinv(y0 + (row_h / 2.0)), hhmm)
            for col_idx in range(cols):
                x0 = margin + time_w + (col_idx * col_w)
                canvas_obj.setFillColor(colors.HexColor("#e5e7eb") if self._planning_interval_blocked(slot_start, slot_end) else colors.white)
                canvas_obj.setStrokeColor(colors.HexColor("#dbe3f0"))
                canvas_obj.rect(x0, yinv(y0 + row_h), col_w, row_h, stroke=1, fill=1)

        clients = {
            str(row.get("codigo", "") or "").strip(): str(row.get("nome", "") or "").strip()
            for row in list(self.ensure_data().get("clientes", []) or [])
            if isinstance(row, dict)
        }
        block_color_fn = getattr(self.plan_actions, "_pdf_block_color_for_item", None)
        planning_items = [
            item
            for item in list(self.ensure_data().get("plano", []) or [])
            if isinstance(item, dict) and self._planning_row_matches_operation(item, planning_operation)
        ]
        for item in planning_items:
            raw_date = str(item.get("data", "") or "").strip()
            raw_start = str(item.get("inicio", "") or "").strip()
            if not raw_date or not raw_start:
                continue
            try:
                block_day = date.fromisoformat(raw_date)
                start_value = self.desktop_main.time_to_minutes(raw_start)
            except Exception:
                continue
            if block_day < week_dates[0] or block_day > week_dates[-1]:
                continue
            duration = self._planning_round_duration(item.get("duracao_min", 0))
            row_start = int((start_value - start_min) // max(1, slot))
            row_span = max(1, int(math.ceil(float(duration) / float(max(1, slot)))))
            if row_start < 0 or row_start >= total_slots:
                continue
            col_idx = (block_day - week_dates[0]).days
            x0 = margin + time_w + (col_idx * col_w)
            y0 = top_margin + ((row_start + 1) * row_h)
            y1 = y0 + (row_span * row_h)
            block_color = str(block_color_fn(item) or "").strip() if callable(block_color_fn) else ""
            if not block_color:
                block_color = str(item.get("color", "") or self._planning_item_color(item.get("encomenda", ""), item.get("material", ""), item.get("espessura", ""))).strip()
            fill_hex = _pdf_mix_hex(block_color or "#c7d2fe", "#ffffff", 0.30)
            edge_hex = _pdf_mix_hex(block_color or "#1f3c88", "#0f172a", 0.20)
            canvas_obj.setFillColor(colors.HexColor(fill_hex))
            canvas_obj.setStrokeColor(colors.HexColor(edge_hex))
            canvas_obj.rect(x0 + 2, yinv(y1 - 2), col_w - 4, (y1 - y0) - 4, stroke=1, fill=1)
            canvas_obj.setFillColor(colors.HexColor(block_color or "#1f3c88"))
            canvas_obj.rect(x0 + 2, yinv(y1 - 2), 6, (y1 - y0) - 4, stroke=0, fill=1)

            enc_num = str(item.get("encomenda", "") or "").strip()
            enc = self.get_encomenda_by_numero(enc_num) or {}
            cliente_txt = clients.get(str(enc.get("cliente", "") or "").strip(), "")
            mat = str(item.get("material", "") or "").strip()
            esp = str(item.get("espessura", "") or "").strip()
            fim_txt = self.desktop_main.minutes_to_time(start_value + duration)
            block_h = (y1 - y0) - 6
            mat_esp = " | ".join(part for part in (mat, f"{esp} mm" if esp else "") if part).strip()
            tempo_txt = f"{duration} min"
            if block_h <= (row_h * 1.2):
                lines = [{"text": enc_num or "-", "role": "title"}, {"text": tempo_txt, "role": "time"}]
                if mat_esp and block_h > (row_h * 0.9):
                    lines.append({"text": mat_esp, "role": "body"})
            elif block_h <= (row_h * 1.9):
                lines = [{"text": enc_num or "-", "role": "title"}]
                if mat_esp:
                    lines.append({"text": mat_esp, "role": "body"})
                lines.append({"text": f"{raw_start} - {fim_txt} | {tempo_txt}", "role": "time"})
            else:
                lines = [{"text": enc_num or "-", "role": "title"}]
                if cliente_txt and block_h >= (row_h * 1.8):
                    lines.append({"text": f"Cliente: {cliente_txt}", "role": "body"})
                if mat_esp:
                    lines.append({"text": mat_esp, "role": "body"})
                lines.append({"text": f"{raw_start} - {fim_txt}", "role": "body"})
                lines.append({"text": tempo_txt, "role": "time"})
            if cliente_txt and block_h >= (row_h * 2.4) and lines and all("Cliente:" not in line for line in lines):
                lines.append({"text": f"Cliente: {cliente_txt}", "role": "body"})
            chapa = str(item.get("chapa", "") or "").strip()
            if chapa and chapa != "-" and block_h >= (row_h * 2.6):
                lines.append({"text": f"Chapa: {chapa}", "role": "body"})
            draw_block_text(x0 + (col_w / 2.0), y0 + ((y1 - y0) / 2.0), lines, col_w - 10, (y1 - y0) - 6)

        box_y = height - margin - footer_box_h
        canvas_obj.setStrokeColor(colors.HexColor("#cbd5e1"))
        canvas_obj.rect(margin, yinv(box_y + footer_box_h), width - (margin * 2), footer_box_h, stroke=1, fill=0)
        canvas_obj.setFillColor(colors.HexColor("#0f172a"))
        canvas_obj.setFont("Helvetica-Bold", 8)
        canvas_obj.drawString(margin + 6, yinv(box_y + 14), "Observacoes:")
        canvas_obj.setFont("Helvetica", 8)
        canvas_obj.line(margin + 74, yinv(box_y + 15), width - margin - 6, yinv(box_y + 15))
        canvas_obj.setFont("Helvetica-Bold", 8)
        canvas_obj.drawString(margin + 6, yinv(box_y + 36), "Data:")
        canvas_obj.drawString(margin + 180, yinv(box_y + 36), "Operador:")

    def _material_assistant_append_suggestions_page(
        self,
        canvas_obj: Any,
        width: float,
        height: float,
        *,
        horizon_days: int = 5,
    ) -> None:
        from reportlab.lib import colors

        rows = [
            dict(row or {})
            for row in list(self.material_assistant_alert_rows(horizon_days=horizon_days) or [])
            if str((row or {}).get("status_key", "") or "").strip().lower() != "ignored"
        ]

        kind_order = {
            "fito_lot": 0,
            "keep_ready": 1,
            "uncativated": 2,
            "shortage": 3,
        }
        rows.sort(
            key=lambda row: (
                int(kind_order.get(str(row.get("kind", "") or "").strip(), 9)),
                0 if str(row.get("status_key", "") or "").strip() == "new" else 1,
                -int(row.get("priority_score", 0) or 0),
                str(row.get("next_action_at", "") or "9999-99-99T99:99"),
                str(row.get("numero", "") or ""),
            )
        )

        margin = 24
        usable_w = width - (margin * 2)
        page_no = 0

        def _kind_label(row: dict[str, Any]) -> str:
            kind = str(row.get("kind", "") or "").strip()
            mapping = {
                "fito_lot": "Troca por urgencia / FIFO",
                "keep_ready": "Nao arrumar / manter pronto",
                "uncativated": "Sem material cativado",
                "shortage": "Sem stock",
            }
            return mapping.get(kind, "Sugestao operacional")

        def _kind_color(row: dict[str, Any]) -> str:
            kind = str(row.get("kind", "") or "").strip()
            mapping = {
                "fito_lot": "#0f3d91",
                "keep_ready": "#b45309",
                "uncativated": "#b45309",
                "shortage": "#b91c1c",
            }
            return mapping.get(kind, "#334155")

        def _draw_page_header() -> float:
            nonlocal page_no
            page_no += 1
            canvas_obj.showPage()
            top_y = height - margin
            header_h = 56
            header_y = top_y - header_h
            canvas_obj.setFillColor(colors.white)
            canvas_obj.setStrokeColor(colors.HexColor("#dbe3f0"))
            canvas_obj.roundRect(margin, header_y, usable_w, header_h, 12, fill=1, stroke=1)
            canvas_obj.setFillColor(colors.HexColor("#0f172a"))
            canvas_obj.setFont("Helvetica-Bold", 17)
            canvas_obj.drawString(margin + 14, header_y + 34, "Sugestoes recomendadas")
            canvas_obj.setFont("Helvetica", 9)
            canvas_obj.setFillColor(colors.HexColor("#475569"))
            canvas_obj.drawString(
                margin + 14,
                header_y + 16,
                "Baseado no planeamento atual e na prioridade FIFO. Serve de apoio ao operador; nao altera nada automaticamente.",
            )
            page_chip_w = 82
            page_chip_h = 24
            page_chip_x = width - margin - page_chip_w
            page_chip_y = header_y + header_h - page_chip_h - 10
            canvas_obj.setFillColor(colors.HexColor("#f8fafc"))
            canvas_obj.setStrokeColor(colors.HexColor("#dbe3f0"))
            canvas_obj.roundRect(page_chip_x, page_chip_y, page_chip_w, page_chip_h, 10, fill=1, stroke=1)
            canvas_obj.setFillColor(colors.HexColor("#0f172a"))
            canvas_obj.setFont("Helvetica-Bold", 9)
            canvas_obj.drawCentredString(page_chip_x + (page_chip_w / 2), page_chip_y + 8, f"Pag. {page_no}")
            info_y = header_y - 8
            info_h = 32
            info_box_w = (usable_w - 16) / 3
            cards = [
                ("Sugestoes", str(len(rows))),
                ("Trocas FIFO", str(len([row for row in rows if str(row.get("kind", "") or "") == "fito_lot"]))),
                ("Horizonte", f"{int(horizon_days or 5)} dias úteis"),
            ]
            for index, (label, value) in enumerate(cards):
                x0 = margin + (index * (info_box_w + 8))
                canvas_obj.setFillColor(colors.HexColor("#f8fafc"))
                canvas_obj.setStrokeColor(colors.HexColor("#dbe3f0"))
                canvas_obj.roundRect(x0, info_y - info_h, info_box_w, info_h, 9, fill=1, stroke=1)
                canvas_obj.setFillColor(colors.HexColor("#64748b"))
                canvas_obj.setFont("Helvetica-Bold", 7.5)
                canvas_obj.drawString(x0 + 10, info_y - 11, label)
                canvas_obj.setFillColor(colors.HexColor("#0f172a"))
                canvas_obj.setFont("Helvetica-Bold", 10)
                canvas_obj.drawRightString(x0 + info_box_w - 10, info_y - 11, str(value or "-"))
            return info_y - info_h - 10

        def _draw_empty_page() -> None:
            y = _draw_page_header()
            canvas_obj.setFillColor(colors.HexColor("#f8fafc"))
            canvas_obj.setStrokeColor(colors.HexColor("#dbe3f0"))
            canvas_obj.roundRect(margin, y - 74, usable_w, 64, 14, fill=1, stroke=1)
            canvas_obj.setFillColor(colors.HexColor("#0f172a"))
            canvas_obj.setFont("Helvetica-Bold", 13)
            canvas_obj.drawString(margin + 16, y - 32, "Sem sugestoes operacionais neste horizonte")
            canvas_obj.setFont("Helvetica", 9)
            canvas_obj.setFillColor(colors.HexColor("#475569"))
            canvas_obj.drawString(
                margin + 16,
                y - 50,
                "Nao existem alertas de troca FIFO, falta de stock ou manutencao de material pronto para os proximos dias.",
            )

        def _draw_row(y: float, row: dict[str, Any]) -> float:
            card_h = 74
            if y < margin + card_h + 8:
                y = _draw_page_header()
            color_hex = _kind_color(row)
            fill_hex = _pdf_mix_hex(color_hex, "#ffffff", 0.90)
            edge_hex = _pdf_mix_hex(color_hex, "#dbeafe", 0.35)
            canvas_obj.setFillColor(colors.HexColor(fill_hex))
            canvas_obj.setStrokeColor(colors.HexColor(edge_hex))
            canvas_obj.roundRect(margin, y - card_h, usable_w, card_h - 2, 12, fill=1, stroke=1)
            canvas_obj.setFillColor(colors.HexColor(color_hex))
            canvas_obj.roundRect(margin, y - card_h, 8, card_h - 2, 8, fill=1, stroke=0)

            header_y = y - 16
            left_x = margin + 16
            right_x = width - margin - 16

            canvas_obj.setFillColor(colors.HexColor("#0f172a"))
            canvas_obj.setFont("Helvetica-Bold", 10)
            canvas_obj.drawString(left_x, header_y, _pdf_clip_text(_kind_label(row), usable_w * 0.42, "Helvetica-Bold", 10))
            canvas_obj.setFont("Helvetica-Bold", 8.5)
            canvas_obj.drawRightString(right_x, header_y, _pdf_clip_text(str(row.get("priority_label", "") or "Media"), usable_w * 0.16, "Helvetica-Bold", 8.5))

            meta_txt = " | ".join(
                part
                for part in (
                    str(row.get("numero", "") or "-").strip() or "-",
                    str(row.get("cliente", "") or "-").strip() or "-",
                    str(row.get("posto_trabalho", "") or "Sem posto").strip() or "Sem posto",
                )
                if part
            )
            canvas_obj.setFont("Helvetica", 8.1)
            canvas_obj.setFillColor(colors.HexColor("#334155"))
            canvas_obj.drawString(left_x, header_y - 13, _pdf_clip_text(meta_txt, usable_w - 32, "Helvetica", 8.1))

            recommendation = str(row.get("recommendation", "") or "").strip()
            if not recommendation:
                recommendation = str(row.get("headline", "") or "").strip() or "Rever sugestao operacional."
            detail_parts = [
                " | ".join(
                    part
                    for part in (
                        str(row.get("material", "") or "-").strip(),
                        f"{str(row.get('espessura', '') or '').strip()} mm".strip() if str(row.get("espessura", "") or "").strip() else "",
                    )
                    if part
                ).strip(" |"),
                "Planeado: " + " ".join(
                    part
                    for part in (
                        str(row.get("planeamento_dia", "") or "").strip(),
                        str(row.get("planeamento_turno", "") or "").strip(),
                        str(row.get("planeamento_hora", "") or "").strip(),
                    )
                    if part and part != "-"
                ).strip(),
                f"Estado: {str(row.get('status_label', '') or 'Nova').strip()}",
            ]
            detail_lines = list(row.get("detail_lines", []) or [])
            if detail_lines:
                detail_parts.append(str(detail_lines[0] or "").strip())

            text_y = y - 44
            canvas_obj.setFillColor(colors.HexColor("#0f172a"))
            canvas_obj.setFont("Helvetica-Bold", 8.6)
            for wrapped in _pdf_wrap_text(recommendation, "Helvetica-Bold", 8.6, usable_w - 32, max_lines=2):
                canvas_obj.drawString(left_x, text_y, wrapped)
                text_y -= 10
            canvas_obj.setFillColor(colors.HexColor("#475569"))
            canvas_obj.setFont("Helvetica", 7.4)
            detail_text = " | ".join(part for part in detail_parts if str(part or "").strip())
            for wrapped in _pdf_wrap_text(detail_text, "Helvetica", 7.4, usable_w - 32, max_lines=2):
                canvas_obj.drawString(left_x, text_y, wrapped)
                text_y -= 8
            return y - card_h - 8

        if not rows:
            _draw_empty_page()
            return

        y = _draw_page_header()
        for row in rows:
            y = _draw_row(y, row)
