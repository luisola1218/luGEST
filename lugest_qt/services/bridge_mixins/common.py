from __future__ import annotations

import re
from pathlib import Path
from typing import Any


class CommonBackendMixin:
    """Legacy adapter for common; see BACKEND_GUIDE.md."""

    def _parse_float(self, value: Any, default: float = 0.0) -> float:
        return float(self.desktop_main.parse_float(value, default))

    def _fmt(self, value: Any) -> str:
        return str(self.desktop_main.fmt_num(value))

    def _write_basic_pdf(self, path: str | Path, lines: list[str]) -> Path:
        target = Path(path)
        target.parent.mkdir(parents=True, exist_ok=True)
        safe_lines = [str(line or "") for line in list(lines or [])]
        width = 595
        height = 842
        content_lines = ["BT", "/F1 11 Tf", "50 800 Td", "14 TL"]
        first = True
        for raw in safe_lines[:52]:
            text = raw.encode("latin-1", errors="replace").decode("latin-1")
            text = text.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
            if first:
                content_lines.append(f"({text}) Tj")
                first = False
            else:
                content_lines.append(f"T* ({text}) Tj")
        content_lines.append("ET")
        stream = "\n".join(content_lines).encode("latin-1", errors="replace")
        objects = [
            b"<< /Type /Catalog /Pages 2 0 R >>",
            b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {width} {height}] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>".encode("ascii"),
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            b"<< /Length " + str(len(stream)).encode("ascii") + b" >>\nstream\n" + stream + b"\nendstream",
        ]
        out = bytearray(b"%PDF-1.4\n")
        offsets = [0]
        for index, obj in enumerate(objects, start=1):
            offsets.append(len(out))
            out.extend(f"{index} 0 obj\n".encode("ascii"))
            out.extend(obj)
            out.extend(b"\nendobj\n")
        xref_offset = len(out)
        out.extend(f"xref\n0 {len(objects) + 1}\n".encode("ascii"))
        out.extend(b"0000000000 65535 f \n")
        for offset in offsets[1:]:
            out.extend(f"{offset:010d} 00000 n \n".encode("ascii"))
        out.extend(
            f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref_offset}\n%%EOF\n".encode("ascii")
        )
        target.write_bytes(bytes(out))
        return target

    def _parse_dimension_mm(self, value: Any, default: float = 0.0) -> float:
        if isinstance(value, (int, float)):
            try:
                return float(value)
            except Exception:
                return default
        text = str(value or "").strip().replace(" ", "")
        if not text:
            return default
        if "," in text:
            try:
                return float(text.replace(".", "").replace(",", "."))
            except Exception:
                pass
        if re.fullmatch(r"\d{1,3}(?:\.\d{3})+", text):
            try:
                return float(text.replace(".", ""))
            except Exception:
                pass
        try:
            return float(text)
        except Exception:
            return default

    def _localizacao(self, record: dict[str, Any]) -> str:
        return str(
            record.get("Localizacao")
            or record.get("Localização")
            or ""
            or ""
        ).strip()

    def _fmt_eur(self, value: Any) -> str:
        try:
            number = float(value or 0)
        except Exception:
            number = 0.0
        return f"{number:,.2f} EUR".replace(",", "X").replace(".", ",").replace("X", ".")

    def _planning_norm_esp(self, value: Any) -> str:
        txt = str(value or "").strip().lower().replace("mm", "").replace(",", ".")
        txt = "".join(ch for ch in txt if ch.isdigit() or ch in ".-")
        if not txt:
            return ""
        try:
            number = float(txt)
            if number.is_integer():
                return str(int(number))
            return f"{number:.6f}".rstrip("0").rstrip(".")
        except Exception:
            return txt

    def _next_prefixed_id(self, rows: list[Any], prefix: str, key: str = "id") -> str:
        max_seq = 0
        prefix_txt = str(prefix or "ID").strip().upper()
        for row in list(rows or []):
            if not isinstance(row, dict):
                continue
            raw = str(row.get(key, "") or "").strip().upper()
            if raw.startswith(f"{prefix_txt}-"):
                suffix = raw.split("-", 1)[1]
                if suffix.isdigit():
                    max_seq = max(max_seq, int(suffix))
        return f"{prefix_txt}-{max_seq + 1:04d}"
