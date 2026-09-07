from __future__ import annotations

import re
import tempfile
from datetime import datetime
from lugest_infra.pdf.light_labels import (
    draw_opp_label as _draw_light_opp_label,
    draw_pallet_label as _draw_light_pallet_label,
)
from pathlib import Path
from typing import Any


class OperatorLabelsBackendMixin:
    """Legacy adapter for operator labels; see BACKEND_GUIDE.md."""

    def _operator_pdf_text(self, value: Any) -> str:
        formatter = getattr(self.desktop_main, "pdf_normalize_text", None)
        if callable(formatter):
            try:
                return str(formatter(value) or "")
            except Exception:
                return str(value or "")
        return str(value or "")

    def _operation_pdf_abbrev(self, operation: Any) -> str:
        raw = str(operation or "").strip()
        if not raw:
            return ""
        normalized = str(self.desktop_main.normalize_operacao_nome(raw) or raw).strip()
        key = str(self.desktop_main.norm_text(normalized) or normalized).casefold()
        aliases = {
            "corte laser": "CL",
            "laser": "CL",
            "quinagem": "Q",
            "roscagem": "ROSC.",
            "serralharia": "SR",
            "soldadura": "SOLD.",
            "lacagem": "LAC.",
            "pintura": "PINT.",
            "maquinacao": "MAQ.",
            "maquinacao cnc": "MAQ.",
            "montagem": "MONT.",
            "embalamento": "EMB.",
            "expedicao": "EXP.",
            "furo manual": "FM",
            "departamento de desenho": "DES.",
            "departamento de orcamentacao": "ORC.",
            "outros": "OUT.",
        }
        return aliases.get(key, normalized.upper()[:8])

    def _operations_pdf_abbrev(self, operations: Any) -> str:
        if isinstance(operations, (list, tuple, set)):
            raw_parts = [str(item.get("nome", "") if isinstance(item, dict) else item or "").strip() for item in operations]
        else:
            text = str(operations or "").strip()
            raw_parts = [part.strip() for part in re.split(r"\s*\+\s*|\s*,\s*|\s*\|\s*", text) if part.strip()]
        parts: list[str] = []
        for raw in raw_parts:
            abbrev = self._operation_pdf_abbrev(raw)
            if abbrev and abbrev not in parts:
                parts.append(abbrev)
        return " + ".join(parts) or "-"

    def _operations_pdf_legend(self, operations: Any) -> list[str]:
        if isinstance(operations, (list, tuple, set)):
            raw_parts = [str(item.get("nome", "") if isinstance(item, dict) else item or "").strip() for item in operations]
        else:
            raw_parts = [part.strip() for part in re.split(r"\s*\+\s*|\s*,\s*|\s*\|\s*", str(operations or "")) if part.strip()]
        lines: list[str] = []
        seen: set[str] = set()
        for raw in raw_parts:
            name = str(self.desktop_main.normalize_operacao_nome(raw) or raw).strip()
            if not name:
                continue
            abbrev = self._operation_pdf_abbrev(name)
            key = f"{abbrev}|{name}"
            if key in seen:
                continue
            seen.add(key)
            lines.append(f"{abbrev} - {name}")
        return lines or ["Sem operações registadas"]

    def _operator_label_palette(self) -> dict[str, Any]:
        from reportlab.lib import colors

        primary_hex = self._normalize_pdf_primary_color(self.branding_settings().get("primary_color", "#00A6A6"))
        primary_soft = self._mix_pdf_hex(primary_hex, "#FFFFFF", 0.84)
        primary_soft_2 = self._mix_pdf_hex(primary_hex, "#FFFFFF", 0.94)
        return {
            "primary": colors.HexColor(primary_hex),
            "primary_hex": primary_hex,
            "primary_dark": colors.HexColor("#0B1F33"),
            "primary_soft": colors.HexColor(primary_soft),
            "primary_soft_2": colors.HexColor(primary_soft_2),
            "ink": colors.HexColor("#14212B"),
            "muted": colors.HexColor("#61717F"),
            "line": colors.HexColor("#CAD3DA"),
            "line_strong": colors.HexColor("#AEBCC7"),
            "surface": colors.white,
            "surface_alt": colors.HexColor("#F3F6F8"),
            "success": colors.HexColor("#107569"),
            "danger": colors.HexColor("#555955"),
            "warning": colors.HexColor("#B54708"),
        }

    @staticmethod
    def _normalize_pdf_primary_color(value: Any) -> str:
        text = str(value or "").strip().upper()
        if re.fullmatch(r"#[0-9A-F]{6}", text):
            return text
        return "#00A6A6"

    @classmethod
    def _mix_pdf_hex(cls, base_hex: str, target_hex: str, ratio: float) -> str:
        base = cls._normalize_pdf_primary_color(base_hex).lstrip("#")
        target = cls._normalize_pdf_primary_color(target_hex).lstrip("#")
        amount = max(0.0, min(1.0, float(ratio)))
        values = []
        for index in (0, 2, 4):
            start = int(base[index : index + 2], 16)
            end = int(target[index : index + 2], 16)
            values.append(round(start + ((end - start) * amount)))
        return "#" + "".join(f"{value:02X}" for value in values)

    def _operator_label_tmp_path(self, enc_num: str, variant: str) -> Path:
        safe_enc = "".join(ch if ch.isalnum() else "_" for ch in str(enc_num or "").strip()) or "operador"
        safe_variant = "".join(ch if ch.isalnum() else "_" for ch in str(variant or "").strip()) or "labels"
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        return Path(tempfile.gettempdir()) / f"lugest_{safe_variant}_{safe_enc}_{stamp}.pdf"

    def _operator_operation_from_posto(self, posto: str = "") -> str:
        posto_norm = self.desktop_main.norm_text(self._legacy_workcenter_group_name(posto) or posto or "")
        if "laser" in posto_norm:
            return "Corte Laser"
        if "quin" in posto_norm:
            return "Quinagem"
        if "rosc" in posto_norm:
            return "Roscagem"
        if "sold" in posto_norm:
            return "Soldadura"
        if "embal" in posto_norm:
            return "Embalamento"
        if "mont" in posto_norm:
            return "Montagem"
        if "maquin" in posto_norm:
            return "Maquinacao"
        if "pint" in posto_norm:
            return "Pintura"
        if "furo" in posto_norm:
            return "Furo Manual"
        return ""

    def _operator_posto_for_operation(self, operation: str = "") -> str:
        normalized = self.desktop_main.normalize_operacao_nome(operation or "") or str(operation or "").strip()
        op_norm = self.desktop_main.norm_text(normalized)
        keyword_map = [
            ("laser", "Laser"),
            ("quin", "Quinagem"),
            ("rosc", "Roscagem"),
            ("sold", "Soldadura"),
            ("embal", "Embalamento"),
            ("mont", "Montagem"),
            ("maquin", "Maquinacao"),
            ("pint", "Pintura"),
            ("furo", "Furo Manual"),
            ("exped", "Expedicao"),
        ]
        for token, fallback in keyword_map:
            if token not in op_norm:
                continue
            for posto in self.workcenter_group_options(operation=normalized):
                if token in self.desktop_main.norm_text(posto):
                    return posto
            return fallback
        return normalized or "-"

    def _operator_next_route(self, piece: dict[str, Any], source_posto: str = "Geral", source_operation: str = "") -> dict[str, str]:
        flow = []
        for op in list(self.desktop_main.ensure_peca_operacoes(piece) or []):
            name = self.desktop_main.normalize_operacao_nome(op.get("nome", "")) or str(op.get("nome", "") or "").strip()
            if name:
                flow.append(name)
        pending_ops = list(self.desktop_main.peca_operacoes_pendentes(piece))
        current_operation = self.desktop_main.normalize_operacao_nome(piece.get("operacao_atual", "")) or ""
        chosen_source_operation = self.desktop_main.normalize_operacao_nome(source_operation or self._operator_operation_from_posto(source_posto)) or current_operation
        available_pending_ops: list[str] = []
        for op_name in flow:
            limit = self._piece_operation_limit(piece, op_name)
            total_done = self._piece_operation_total(self._piece_operation_row(piece, op_name), limit)
            if limit > total_done + 1e-9:
                available_pending_ops.append(op_name)
        next_operation = available_pending_ops[0] if available_pending_ops else (pending_ops[0] if pending_ops else "")
        if not current_operation:
            current_operation = next_operation or (flow[-1] if flow else "")
        origin_posto = str(source_posto or "").strip()
        if not origin_posto or self.desktop_main.norm_text(origin_posto) == "geral":
            origin_posto = self._operator_posto_for_operation(chosen_source_operation or current_operation)
        next_posto = "Expedicao" if (not next_operation and flow) else self._operator_posto_for_operation(next_operation)
        return {
            "flow": " -> ".join(flow),
            "source_operation": chosen_source_operation or current_operation or "-",
            "current_operation": current_operation or "-",
            "source_posto": origin_posto or "-",
            "next_operation": next_operation or "Expedicao",
            "next_posto": next_posto or "Expedicao",
        }

    def _ensure_operator_piece_opp(self, piece: dict[str, Any]) -> bool:
        opp = str(piece.get("opp", "") or "").strip()
        if opp:
            return False
        for enc in list(self.ensure_data().get("encomendas", []) or []):
            if any(row is piece for row in list(self.desktop_main.encomenda_pecas(enc) or [])):
                piece["opp"] = self._next_order_opp_codigo(enc)
                if not str(piece.get("of", "") or "").strip():
                    piece["of"] = self._order_of_code(enc, create=True)
                return bool(piece.get("opp"))
        piece["opp"] = str(self.desktop_main.next_opp_numero(self.ensure_data()) or "").strip()
        return bool(piece.get("opp"))

    def _operator_label_row(self, enc: dict[str, Any], piece: dict[str, Any], source_posto: str = "Geral") -> dict[str, Any]:
        cliente_codigo = str(enc.get("cliente", "") or "").strip()
        cliente_obj = {}
        find_cliente_fn = getattr(self.desktop_main, "find_cliente", None)
        if callable(find_cliente_fn) and cliente_codigo:
            try:
                cliente_obj = find_cliente_fn(self.ensure_data(), cliente_codigo) or {}
            except Exception:
                cliente_obj = {}
        route = self._operator_next_route(piece, source_posto=source_posto)
        descricao = str(piece.get("descricao", "") or piece.get("Observacoes", "") or "").strip()
        qtd = self._parse_float(piece.get("quantidade_pedida", 0), 0)
        return {
            "piece_id": str(piece.get("id", "") or "").strip(),
            "opp": str(piece.get("opp", "") or "").strip(),
            "of": str(piece.get("of", "") or "").strip(),
            "encomenda": str(enc.get("numero", "") or "").strip(),
            "cliente": cliente_codigo,
            "cliente_nome": str(cliente_obj.get("nome", "") or "").strip(),
            "cliente_label": f"{cliente_codigo} - {str(cliente_obj.get('nome', '') or '').strip()}".strip(" -"),
            "ref_interna": str(piece.get("ref_interna", "") or "").strip(),
            "ref_externa": str(piece.get("ref_externa", "") or "").strip(),
            "descricao": descricao,
            "material": str(piece.get("material", "") or "").strip(),
            "espessura": str(piece.get("espessura", "") or "").strip(),
            "quantidade": qtd,
            "quantidade_txt": self._fmt(qtd),
            "estado": str(piece.get("estado", "") or "").strip(),
            "operacao_atual": route.get("current_operation", "-"),
            "operacao_origem": route.get("source_operation", "-"),
            "posto_origem": route.get("source_posto", "-"),
            "proxima_operacao": route.get("next_operation", "Expedicao"),
            "proximo_posto": route.get("next_posto", "Expedicao"),
            "fluxo": route.get("flow", ""),
        }

    def _operator_label_rows_for_order(self, enc_num: str, source_posto: str = "Geral") -> tuple[dict[str, Any], list[dict[str, Any]]]:
        enc = self.get_encomenda_by_numero(enc_num)
        if enc is None:
            raise ValueError("Encomenda nao encontrada.")
        changed = False
        rows: list[dict[str, Any]] = []
        for piece in list(self.desktop_main.encomenda_pecas(enc) or []):
            changed = self._ensure_operator_piece_opp(piece) or changed
            rows.append(self._operator_label_row(enc, piece, source_posto=source_posto))
        if changed:
            self._save(force=True)
        rows.sort(
            key=lambda row: (
                str(row.get("proximo_posto", "") or ""),
                str(row.get("ref_interna", "") or ""),
                str(row.get("opp", "") or ""),
            )
        )
        return enc, rows

    def _operator_selected_label_rows(self, enc_num: str, piece_ids: list[str] | tuple[str, ...], source_posto: str = "Geral") -> tuple[dict[str, Any], list[dict[str, Any]]]:
        enc = self.get_encomenda_by_numero(enc_num)
        if enc is None:
            raise ValueError("Encomenda nao encontrada.")
        selected_ids = [str(piece_id or "").strip() for piece_id in list(piece_ids or []) if str(piece_id or "").strip()]
        if not selected_ids:
            raise ValueError("Seleciona pelo menos uma referencia para imprimir etiquetas.")
        selected_set = set(selected_ids)
        changed = False
        rows: list[dict[str, Any]] = []
        for piece in list(self.desktop_main.encomenda_pecas(enc) or []):
            piece_id = str(piece.get("id", "") or "").strip()
            if piece_id not in selected_set:
                continue
            changed = self._ensure_operator_piece_opp(piece) or changed
            rows.append(self._operator_label_row(enc, piece, source_posto=source_posto))
        if not rows:
            raise ValueError("As referencias selecionadas ja nao existem nesta encomenda.")
        order_map = {piece_id: index for index, piece_id in enumerate(selected_ids)}
        rows.sort(key=lambda row: (order_map.get(str(row.get("piece_id", "") or "").strip(), 999999), str(row.get("ref_interna", "") or "")))
        if changed:
            self._save(force=True)
        return enc, rows

    def _draw_operator_logo(self, canvas_obj, logo_path: Path | None, x: float, y: float, width: float, height: float) -> None:
        from reportlab.lib.utils import ImageReader

        if not logo_path or not logo_path.exists():
            return
        canvas_obj.saveState()
        try:
            img_reader = None
            iw = ih = 0
            try:
                from PIL import Image, ImageChops

                img = Image.open(logo_path).convert("RGBA")
                bbox = img.split()[-1].getbbox()
                if not bbox:
                    rgb = img.convert("RGB")
                    diff = ImageChops.difference(rgb, Image.new("RGB", rgb.size, (255, 255, 255)))
                    bbox = diff.getbbox()
                if bbox:
                    img = img.crop(bbox)
                logo_scale = max(1.0, self._branding_logo_scale_factor())
                max_px = (
                    max(96, int(float(width) * 4.0 * logo_scale)),
                    max(72, int(float(height) * 4.0 * logo_scale)),
                )
                if img.width > max_px[0] or img.height > max_px[1]:
                    img.thumbnail(max_px)
                iw, ih = img.size
                img_reader = ImageReader(img)
            except Exception:
                img_reader = ImageReader(str(logo_path))
                iw, ih = img_reader.getSize()
            if not img_reader or iw <= 0 or ih <= 0:
                return
            scale = min(float(width) / float(iw), float(height) / float(ih))
            scale *= self._branding_logo_scale_factor()
            draw_w = max(1.0, float(iw) * scale)
            draw_h = max(1.0, float(ih) * scale)
            draw_x = x + (float(width) - draw_w) / 2.0
            draw_y = y + (float(height) - draw_h) / 2.0
            clip_path = canvas_obj.beginPath()
            clip_path.rect(x, y, width, height)
            canvas_obj.clipPath(clip_path, stroke=0, fill=0)
            canvas_obj.drawImage(img_reader, draw_x, draw_y, width=draw_w, height=draw_h, preserveAspectRatio=True, mask="auto")
        except Exception:
            pass
        finally:
            canvas_obj.restoreState()

    def _draw_operator_logo_plate(
        self,
        canvas_obj,
        palette: dict[str, Any],
        logo_path: Path | None,
        x: float,
        y: float,
        width: float,
        height: float,
        *,
        radius: float = 10,
        padding_x: float = 5,
        padding_y: float = 4,
        line_width: float = 0.9,
    ) -> None:
        canvas_obj.saveState()
        canvas_obj.setFillColor(palette["surface"])
        canvas_obj.setStrokeColor(palette["line"])
        canvas_obj.setLineWidth(line_width)
        canvas_obj.roundRect(x, y, width, height, radius, stroke=1, fill=1)
        canvas_obj.restoreState()
        self._draw_operator_logo(
            canvas_obj,
            logo_path,
            x + (padding_x / 3.2),
            y + (padding_y / 3.2),
            max(10.0, width - ((padding_x / 3.2) * 2)),
            max(8.0, height - ((padding_y / 3.2) * 2)),
        )

    def _draw_operator_unit_label(
        self,
        canvas_obj,
        page_width: float,
        page_height: float,
        row: dict[str, Any],
        palette: dict[str, Any],
        logo_path: Path | None,
        printed_at: str,
    ) -> None:
        _draw_light_opp_label(
            canvas_obj,
            page_width,
            page_height,
            row,
            palette,
            logo_path,
            printed_at,
            draw_logo=self._draw_operator_logo_plate,
            draw_barcode=self._draw_code128_fit,
        )
        return

    def _draw_operator_pallet_page(
        self,
        canvas_obj,
        page_width: float,
        page_height: float,
        rows: list[dict[str, Any]],
        group_rows: list[dict[str, Any]],
        group_name: str,
        palette: dict[str, Any],
        logo_path: Path | None,
        source_posto: str,
        printed_at: str,
        page_number: int,
        total_pages: int,
    ) -> None:
        _draw_light_pallet_label(
            canvas_obj,
            page_width,
            page_height,
            rows,
            group_rows,
            group_name,
            palette,
            logo_path,
            source_posto,
            printed_at,
            page_number,
            total_pages,
            draw_logo=self._draw_operator_logo_plate,
            fmt=self._fmt,
        )
        return

    def operator_label_rows(self, enc_num: str, source_posto: str = "Geral") -> dict[str, Any]:
        enc, rows = self._operator_label_rows_for_order(enc_num, source_posto=source_posto)
        return {
            "encomenda": str(enc.get("numero", "") or "").strip(),
            "cliente": str(enc.get("cliente", "") or "").strip(),
            "source_posto": str(source_posto or "").strip() or "Geral",
            "rows": rows,
        }

    def operator_unit_labels_pdf(
        self,
        enc_num: str,
        piece_ids: list[str] | tuple[str, ...],
        source_posto: str = "Geral",
        output_path: str | Path | None = None,
    ) -> Path:
        from reportlab.lib.units import mm
        from reportlab.pdfgen import canvas as pdf_canvas

        _enc, rows = self._operator_selected_label_rows(enc_num, piece_ids, source_posto=source_posto)
        target = Path(output_path) if output_path else self._operator_label_tmp_path(enc_num, "operator_unit_labels")
        target.parent.mkdir(parents=True, exist_ok=True)
        width, height = (150 * mm, 100 * mm)
        palette = self._operator_label_palette()
        branding = self.branding_settings()
        logo_txt = str(branding.get("logo_path", "") or "").strip()
        logo_path = Path(logo_txt) if logo_txt and Path(logo_txt).exists() else None
        printed_at = str(self.desktop_main.now_iso() or "").replace("T", " ")[:19]
        canvas_obj = pdf_canvas.Canvas(str(target), pagesize=(width, height))
        for index, row in enumerate(rows):
            if index:
                canvas_obj.showPage()
            self._draw_operator_unit_label(canvas_obj, width, height, row, palette, logo_path, printed_at)
        canvas_obj.save()
        return target

    def operator_pallet_labels_pdf(
        self,
        enc_num: str,
        piece_ids: list[str] | tuple[str, ...],
        source_posto: str = "Geral",
        output_path: str | Path | None = None,
    ) -> Path:
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.pdfgen import canvas as pdf_canvas

        _enc, selected_rows = self._operator_selected_label_rows(enc_num, piece_ids, source_posto=source_posto)
        groups: dict[str, list[dict[str, Any]]] = {}
        for row in selected_rows:
            key = str(row.get("proximo_posto", "") or "Expedicao").strip() or "Expedicao"
            groups.setdefault(key, []).append(row)
        ordered_groups = sorted(groups.items(), key=lambda item: item[0].lower())
        rows_per_page = 13
        pages: list[tuple[str, list[dict[str, Any]], list[dict[str, Any]]]] = []
        for group_name, group_rows in ordered_groups:
            group_rows.sort(key=lambda row: (str(row.get("ref_interna", "") or ""), str(row.get("opp", "") or "")))
            for start in range(0, len(group_rows), rows_per_page):
                pages.append((group_name, group_rows[start : start + rows_per_page], group_rows))
        if not pages:
            raise ValueError("Sem referencias para gerar a etiqueta de palete.")
        target = Path(output_path) if output_path else self._operator_label_tmp_path(enc_num, "operator_pallet_labels")
        target.parent.mkdir(parents=True, exist_ok=True)
        page_width, page_height = landscape(A4)
        palette = self._operator_label_palette()
        branding = self.branding_settings()
        logo_txt = str(branding.get("logo_path", "") or "").strip()
        logo_path = Path(logo_txt) if logo_txt and Path(logo_txt).exists() else None
        printed_at = str(self.desktop_main.now_iso() or "").replace("T", " ")[:19]
        canvas_obj = pdf_canvas.Canvas(str(target), pagesize=(page_width, page_height))
        total_pages = len(pages)
        for page_index, (group_name, page_rows, group_rows) in enumerate(pages, start=1):
            if page_index > 1:
                canvas_obj.showPage()
            self._draw_operator_pallet_page(
                canvas_obj,
                page_width,
                page_height,
                page_rows,
                group_rows,
                group_name,
                palette,
                logo_path,
                str(source_posto or "").strip() or "Geral",
                printed_at,
                page_index,
                total_pages,
            )
        canvas_obj.save()
        return target
