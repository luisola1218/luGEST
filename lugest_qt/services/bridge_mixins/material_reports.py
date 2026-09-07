from __future__ import annotations

import os
import re
import tempfile
from lugest_infra.pdf.dossier_reports import render_material_history as _render_dossier_material_history
from lugest_infra.pdf.font_policy import pdf_font_size_limit
from lugest_infra.pdf.light_inventory import render_material_stock_pdf as _render_light_material_stock_pdf
from lugest_infra.pdf.light_labels import draw_material_label as _draw_light_material_label
from pathlib import Path
from types import SimpleNamespace
from typing import Any


class MaterialReportsBackendMixin:
    """Legacy adapter for material reports; see BACKEND_GUIDE.md."""

    def _material_history_entry_row(self, entry: dict[str, Any], record: dict[str, Any] | None = None) -> dict[str, str]:
        details = str(entry.get("detalhes", "") or "").strip()
        material_id = str((record or {}).get("id", "") or "").strip()
        material_name = str((record or {}).get("material", "") or "").strip()
        espessura = str((record or {}).get("espessura", "") or "").strip()
        lote = str((record or {}).get("lote_fornecedor", "") or "").strip()
        local = self._localizacao(record or {}) if record else ""
        dimensao = ""
        quantity = ""
        reserved = ""

        qty_match = re.search(r"\bqtd=([-+]?\d+(?:[.,]\d+)?)", details, flags=re.IGNORECASE)
        if qty_match:
            quantity = qty_match.group(1)
        reserved_match = re.search(r"\breservado=([-+]?\d+(?:[.,]\d+)?)", details, flags=re.IGNORECASE)
        if reserved_match:
            reserved = reserved_match.group(1)
        if not material_name:
            prefix = re.split(r"\bqtd=|\breservado=", details, maxsplit=1, flags=re.IGNORECASE)[0].strip(" |,-")
            if " Lote:" in prefix:
                base_txt, lote_txt = prefix.split(" Lote:", 1)
                if not lote:
                    lote = lote_txt.strip()
                prefix = base_txt.strip()
            dim_match = re.search(r"(\d+(?:[.,]\d+)?x\d+(?:[.,]\d+)?)", prefix)
            if dim_match:
                dimensao = dim_match.group(1)
                prefix = prefix.replace(dimensao, " ").strip()
            esp_match = re.search(r"(\d+(?:[.,]\d+)?)\s*$", prefix)
            if esp_match:
                if not espessura:
                    espessura = esp_match.group(1)
                prefix = prefix[: esp_match.start()].strip(" |-")
            material_name = prefix or material_name
        else:
            comp = str((record or {}).get("comprimento", "") or "").strip()
            larg = str((record or {}).get("largura", "") or "").strip()
            if comp and larg:
                dimensao = f"{comp}x{larg}"

        return {
            "data": str(entry.get("data", "") or "").replace("T", " ")[:19],
            "acao": str(entry.get("acao", "") or "").strip(),
            "operador": str(entry.get("operador", "") or "").strip(),
            "material_id": material_id,
            "material": material_name,
            "espessura": espessura,
            "dimensao": dimensao,
            "lote": lote,
            "local": local,
            "qtd": quantity,
            "reservado": reserved,
            "detalhes": details,
        }

    def material_history_rows(self, material_id: str = "", limit: int = 240) -> list[dict[str, str]]:
        material_id = str(material_id or "").strip()
        record = self.material_by_id(material_id) if material_id else None
        lote = str((record or {}).get("lote_fornecedor", "") or "").strip()
        material_name = str((record or {}).get("material", "") or "").strip()
        rows: list[dict[str, str]] = []
        for entry in list(reversed(self.ensure_data().get("stock_log", [])[-max(limit, 1) * 4 :])):
            detalhes = str(entry.get("detalhes", "") or "").strip()
            if material_id and material_id not in detalhes and (not lote or lote not in detalhes) and (not material_name or material_name not in detalhes):
                continue
            rows.append(self._material_history_entry_row(entry, record))
            if len(rows) >= limit:
                break
        return rows

    def material_stock_formats(self) -> list[str]:
        formats = {
            str(row.get("formato", "") or self.desktop_main.detect_materia_formato(row) or "Chapa").strip()
            for row in list(self.ensure_data().get("materiais", []) or [])
            if isinstance(row, dict)
        }
        return sorted((value for value in formats if value), key=str.casefold)

    def material_open_stock_pdf(self, in_stock_only: bool = False, formats: list[str] | None = None) -> Path:
        target = Path(tempfile.gettempdir()) / "lugest_stock.pdf"
        self.material_render_stock_pdf(target, in_stock_only=in_stock_only, formats=formats)
        os.startfile(str(target))
        return target

    def material_render_stock_pdf(
        self,
        path: str | Path,
        in_stock_only: bool = False,
        formats: list[str] | None = None,
    ) -> Path:
        return _render_light_material_stock_pdf(
            Path(path),
            self.ensure_data(),
            self.branding_settings(),
            in_stock_only=bool(in_stock_only),
            formats=formats,
        )

    def material_open_history_pdf(self) -> Path:
        target = Path(tempfile.gettempdir()) / "lugest_qt_materiais_historico.pdf"
        helper = SimpleNamespace(data=self.ensure_data())
        self.materia_actions.render_stock_log_pdf(helper, str(target))
        os.startfile(str(target))
        return target

    def material_render_history_pdf(self, rows: list[dict[str, Any]], title: str, path: str | Path) -> Path:
        return _render_dossier_material_history(self, Path(path), list(rows or []), str(title or "Historico de materia-prima"))

    def _material_is_retalho(self, record: dict[str, Any]) -> bool:
        checker = getattr(self.materia_actions, "_is_retalho_like", None)
        if callable(checker):
            try:
                return bool(checker(record))
            except Exception:
                pass
        return bool((record or {}).get("is_sobra")) or self._localizacao(record).strip().upper() == "RETALHO"

    def _material_label_lot_text(self, record: dict[str, Any]) -> str:
        lote_interno = str((record or {}).get("lote_interno", "") or "").strip()
        origem_lotes = [str(item or "").strip() for item in list((record or {}).get("origem_lotes_baixa", []) or []) if str(item or "").strip()]
        lote = str((record or {}).get("lote_fornecedor", "") or "").strip()
        origem_lote = str((record or {}).get("origem_lote", "") or "").strip()
        if lote_interno:
            return lote_interno
        if origem_lotes:
            return " + ".join(origem_lotes)
        if lote:
            return lote
        if origem_lote:
            return origem_lote
        return "-"

    def _material_label_dimension_text(self, record: dict[str, Any]) -> str:
        preview = self.material_geometry_preview(record)
        dimension_text = str(preview.get("dimension_label", "") or "").strip()
        metros = float(preview.get("metros", 0) or 0)
        if dimension_text and dimension_text != "-":
            if metros > 0:
                return f"{dimension_text} | {self._fmt(metros)} m"
            return dimension_text
        if metros > 0:
            return f"{self._fmt(metros)} m"
        return "-"

    def _draw_code128_fit(
        self,
        canvas_obj,
        value: Any,
        x: float,
        y: float,
        max_width: float,
        bar_height: float,
        min_bar_width: float = 0.38,
        max_bar_width: float = 1.05,
        align: str = "center",
    ) -> float:
        from reportlab.graphics.barcode import code128

        safe_value = str(value or "-").strip() or "-"
        probe = code128.Code128(safe_value, barHeight=bar_height, barWidth=min_bar_width)
        probe_width = float(getattr(probe, "width", 0.0) or 0.0)
        if probe_width > 0:
            unit_width = probe_width / max(min_bar_width, 0.01)
            target_bar_width = max_width / unit_width if unit_width > 0 else min_bar_width
            target_bar_width = max(min_bar_width, min(max_bar_width, target_bar_width))
            barcode = code128.Code128(safe_value, barHeight=bar_height, barWidth=target_bar_width)
        else:
            barcode = probe
        actual_width = float(getattr(barcode, "width", max_width) or max_width)
        draw_x = x
        if align == "center":
            draw_x = x + max(0.0, (max_width - actual_width) / 2.0)
        elif align == "right":
            draw_x = x + max(0.0, max_width - actual_width)
        barcode.drawOn(canvas_obj, draw_x, y)
        return actual_width

    def _draw_material_stock_label(
        self,
        canvas_obj,
        page_width: float,
        page_height: float,
        record: dict[str, Any],
        palette: dict[str, Any],
        logo_path: Path | None,
        printed_at: str,
    ) -> None:
        stock_id = str(record.get("id", "") or "-").strip() or "-"
        formato = str(record.get("formato") or self.desktop_main.detect_materia_formato(record) or "Chapa").strip() or "Chapa"
        available = max(0.0, self._parse_float(record.get("quantidade", 0), 0) - self._parse_float(record.get("reservado", 0), 0))
        is_retalho = self._material_is_retalho(record)
        _draw_light_material_label(
            canvas_obj,
            page_width,
            page_height,
            record,
            palette,
            logo_path,
            printed_at,
            draw_logo=self._draw_operator_logo_plate,
            draw_barcode=self._draw_code128_fit,
            scan_code=str(record.get("scan_code", "") or self.inventory_scan_code("MAT", stock_id)).strip(),
            dimension_text=self._material_label_dimension_text(record),
            lot_text=self._material_label_lot_text(record),
            location_text=self._localizacao(record) or "-",
            available_text=self._fmt(available),
            format_text=formato,
            kind_text="Retalho" if is_retalho else str(record.get("tipo", "") or "Chapa / Palete"),
            weight_text=f"{self._fmt(record.get('peso_unid', 0))} kg",
        )
        return

    def _draw_material_retalho_label(
        self,
        canvas_obj,
        page_width: float,
        page_height: float,
        record: dict[str, Any],
        palette: dict[str, Any],
        logo_path: Path | None,
        printed_at: str,
    ) -> None:
        stock_id = str(record.get("id", "") or "-").strip() or "-"
        formato = str(record.get("formato") or self.desktop_main.detect_materia_formato(record) or "Chapa").strip() or "Chapa"
        available = max(0.0, self._parse_float(record.get("quantidade", 0), 0) - self._parse_float(record.get("reservado", 0), 0))
        _draw_light_material_label(
            canvas_obj,
            page_width,
            page_height,
            record,
            palette,
            logo_path,
            printed_at,
            draw_logo=self._draw_operator_logo_plate,
            draw_barcode=self._draw_code128_fit,
            scan_code=str(record.get("scan_code", "") or self.inventory_scan_code("MAT", stock_id)).strip(),
            dimension_text=self._material_label_dimension_text(record),
            lot_text=self._material_label_lot_text(record),
            location_text=self._localizacao(record) or "-",
            available_text=self._fmt(available),
            format_text=formato,
            kind_text="Retalho",
            weight_text=f"{self._fmt(record.get('peso_unid', 0))} kg",
        )
        return

    def material_identification_label_pdf(self, material_id: str, output_path: str | Path | None = None) -> Path:
        from reportlab.lib.units import mm
        from reportlab.pdfgen import canvas as pdf_canvas

        record = self.material_by_id(material_id)
        if record is None:
            raise ValueError("Material não encontrado.")
        self.materia_actions._hydrate_retalho_record(self.ensure_data(), record)
        is_retalho = self._material_is_retalho(record)
        target = Path(output_path) if output_path else self._operator_label_tmp_path(material_id, "material_identification")
        target.parent.mkdir(parents=True, exist_ok=True)
        page_size = (110 * mm, 50 * mm)
        palette = self._operator_label_palette()
        branding = self.branding_settings()
        logo_txt = str(branding.get("logo_path", "") or "").strip()
        logo_path = Path(logo_txt) if logo_txt and Path(logo_txt).exists() else None
        printed_at = str(self.desktop_main.now_iso() or "").replace("T", " ")[:19]
        canvas_obj = pdf_canvas.Canvas(str(target), pagesize=page_size)
        with pdf_font_size_limit(14.0):
            if is_retalho:
                self._draw_material_retalho_label(canvas_obj, page_size[0], page_size[1], record, palette, logo_path, printed_at)
            else:
                self._draw_material_stock_label(canvas_obj, page_size[0], page_size[1], record, palette, logo_path, printed_at)
        canvas_obj.save()
        return target
