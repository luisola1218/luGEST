from lugest_infra.legacy.module_context import configure_module, ensure_module

_CONFIGURED = False

def configure(main_globals):
    configure_module(globals(), main_globals)

def _ensure_configured():
    ensure_module(globals(), "orc_actions")


def _orc_pdf_register_fonts():
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    regular = "Helvetica"
    bold = "Helvetica-Bold"
    candidates = [
        ("Arial", r"C:\Windows\Fonts\arial.ttf"),
        ("Arial-Bold", r"C:\Windows\Fonts\arialbd.ttf"),
        ("SegoeUI", r"C:\Windows\Fonts\segoeui.ttf"),
        ("SegoeUI-Bold", r"C:\Windows\Fonts\segoeuib.ttf"),
    ]
    for name, path in candidates:
        try:
            if path and os.path.exists(path):
                pdfmetrics.registerFont(TTFont(name, path))
        except Exception:
            pass
    try:
        names = set(pdfmetrics.getRegisteredFontNames())
        if "Arial" in names:
            regular = "Arial"
        elif "SegoeUI" in names:
            regular = "SegoeUI"
        if "Arial-Bold" in names:
            bold = "Arial-Bold"
        elif "SegoeUI-Bold" in names:
            bold = "SegoeUI-Bold"
    except Exception:
        pass
    return {"regular": regular, "bold": bold}


def _orc_pdf_hex_to_rgb(value):
    txt = str(value or "").strip().lstrip("#")
    if len(txt) == 3:
        txt = "".join(ch * 2 for ch in txt)
    if len(txt) != 6:
        txt = "1F3C88"
    try:
        return tuple(int(txt[i : i + 2], 16) for i in (0, 2, 4))
    except Exception:
        return (31, 60, 136)


def _orc_pdf_rgb_to_hex(rgb):
    r, g, b = [max(0, min(255, int(v))) for v in rgb]
    return f"#{r:02X}{g:02X}{b:02X}"


def _orc_pdf_mix_hex(base_hex, target_hex, ratio):
    ratio = max(0.0, min(1.0, float(ratio)))
    base = _orc_pdf_hex_to_rgb(base_hex)
    target = _orc_pdf_hex_to_rgb(target_hex)
    out = []
    for base_v, target_v in zip(base, target):
        out.append(round((base_v * (1.0 - ratio)) + (target_v * ratio)))
    return _orc_pdf_rgb_to_hex(tuple(out))


def _orc_pdf_brand_palette():
    from reportlab.lib import colors

    primary_hex = str(get_branding_config().get("primary_color", "") or "#1F3C88").strip() or "#1F3C88"
    return {
        "primary": colors.HexColor(primary_hex),
        "primary_dark": colors.HexColor(_orc_pdf_mix_hex(primary_hex, "#000000", 0.22)),
        "primary_soft": colors.HexColor(_orc_pdf_mix_hex(primary_hex, "#FFFFFF", 0.80)),
        "primary_soft_2": colors.HexColor(_orc_pdf_mix_hex(primary_hex, "#FFFFFF", 0.90)),
        "ink": colors.HexColor(_orc_pdf_mix_hex(primary_hex, "#1A1A1A", 0.72)),
        "muted": colors.HexColor("#667085"),
        "line": colors.HexColor(_orc_pdf_mix_hex(primary_hex, "#D7DEE8", 0.76)),
        "line_strong": colors.HexColor(_orc_pdf_mix_hex(primary_hex, "#708090", 0.36)),
        "surface": colors.white,
        "surface_alt": colors.HexColor("#F7F9FC"),
        "surface_warm": colors.HexColor("#FCFCFD"),
    }


def _orc_pdf_clip_text(value, max_w, font_name, font_size):
    from reportlab.pdfbase import pdfmetrics

    txt = "" if value is None else str(value)
    if pdfmetrics.stringWidth(txt, font_name, font_size) <= max_w:
        return txt
    ellipsis = "..."
    while txt and pdfmetrics.stringWidth(txt + ellipsis, font_name, font_size) > max_w:
        txt = txt[:-1]
    return f"{txt}{ellipsis}" if txt else ""


def _orc_pdf_wrap_text(value, font_name, font_size, max_w, max_lines=None):
    from reportlab.pdfbase import pdfmetrics

    text = str(value or "").replace("\r", " ").replace("\n", " ").strip()
    if not text:
        return []
    words = text.split()
    lines = []
    current = ""
    for word in words:
        candidate = f"{current} {word}".strip() if current else word
        if pdfmetrics.stringWidth(candidate, font_name, font_size) <= max_w:
            current = candidate
            continue
        if current:
            lines.append(current)
        current = word
        if max_lines and len(lines) >= max_lines:
            break
    if current and (not max_lines or len(lines) < max_lines):
        lines.append(current)
    if max_lines and len(lines) > max_lines:
        lines = lines[:max_lines]
    return lines


def _orc_pdf_metric_grid_layout(group_right, banner_top, banner_height, cols=2, rows=2, group_w=230, gap=6):
    cols = max(1, int(cols))
    rows = max(1, int(rows))
    group_w = max(120, int(group_w))
    gap = max(4, int(gap))
    chip_h = max(26, int((float(banner_height) - (gap * (rows + 1))) / rows))
    chip_w = max(48, int((float(group_w) - (gap * (cols + 1))) / cols))
    group_x = float(group_right) - float(group_w)
    row_positions = [float(banner_top) + float(gap) + (idx * (float(chip_h) + float(gap))) for idx in range(rows)]
    col_positions = [group_x + float(gap) + (idx * (float(chip_w) + float(gap))) for idx in range(cols)]
    return {
        "group_x": group_x,
        "chip_w": chip_w,
        "chip_h": chip_h,
        "rows": row_positions,
        "cols": col_positions,
    }

def get_orc_by_numero(self, numero):
    _ensure_configured()
    for o in self.data.get("orcamentos", []):
        if o.get("numero") == numero:
            return o
    return None


def _save_orc_fast(self, orc):
    _ensure_configured()
    if not isinstance(orc, dict):
        return False
    try:
        mysql_upsert_orcamento_com_linhas(self.data, orc)
        try:
            _set_last_save_fingerprint(self.data)
        except Exception:
            pass
        return True
    except Exception:
        return False


def _orc_extract_year(date_txt, numero_txt="", ano_val=None):
    _ensure_configured()
    a = str(ano_val or "").strip()
    if len(a) == 4 and a.isdigit():
        return a
    d = str(date_txt or "").strip()
    if len(d) >= 4 and d[:4].isdigit():
        return d[:4]
    n = str(numero_txt or "").strip()
    if n.startswith("ORC-") and len(n) >= 8 and n[4:8].isdigit():
        return n[4:8]
    return ""


def _mysql_orc_years():
    if not (USE_MYSQL_STORAGE and MYSQL_AVAILABLE):
        return []
    conn = None
    try:
        conn = _mysql_connect()
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT DISTINCT
                    CASE
                        WHEN `ano` IS NOT NULL THEN `ano`
                        WHEN `data` IS NOT NULL THEN YEAR(`data`)
                        WHEN `numero` REGEXP '^ORC-[0-9]{4}-' THEN CAST(SUBSTRING(`numero`, 5, 4) AS UNSIGNED)
                        ELSE NULL
                    END AS ano
                FROM `orcamentos`
                HAVING ano IS NOT NULL
                ORDER BY ano DESC
                """
            )
            rows = cur.fetchall() or []
        out = []
        for r in rows:
            val = r.get("ano") if isinstance(r, dict) else (r[0] if r else None)
            try:
                y = int(val)
            except Exception:
                continue
            if y > 0:
                out.append(str(y))
        return out
    except Exception:
        return []
    finally:
        try:
            if conn is not None:
                conn.close()
        except Exception:
            pass


def refresh_orc_year_options(self, keep_selection=True):
    _ensure_configured()
    current_year = str(datetime.now().year)
    selected = (
        str(self.orc_year_filter.get() or "").strip()
        if keep_selection and hasattr(self, "orc_year_filter")
        else current_year
    )
    years = set()
    years.add(current_year)
    for o in self.data.get("orcamentos", []):
        if not isinstance(o, dict):
            continue
        y = _orc_extract_year(o.get("data", ""), o.get("numero", ""), o.get("ano"))
        if y:
            years.add(y)
    for y in _mysql_orc_years():
        years.add(y)
    year_values = sorted(
        years,
        key=lambda x: int(x) if str(x).isdigit() else 0,
        reverse=True,
    )
    values = year_values + ["Todos"]
    if not values:
        values = [current_year, "Todos"]
    if selected not in values:
        selected = current_year if current_year in values else values[0]
    if hasattr(self, "orc_year_filter"):
        self.orc_year_filter.set(selected)
    try:
        cb = getattr(self, "orc_year_cb", None)
        if cb is not None:
            if hasattr(cb, "configure"):
                try:
                    cb.configure(values=values)
                except Exception:
                    pass
            if hasattr(cb, "__setitem__"):
                try:
                    cb["values"] = values
                except Exception:
                    pass
    except Exception:
        pass
    return values


def _reload_orcamentos_from_mysql(self, year_value=None):
    _ensure_configured()
    if not (USE_MYSQL_STORAGE and MYSQL_AVAILABLE):
        return False
    year_txt = str(
        year_value if year_value is not None else (self.orc_year_filter.get() if hasattr(self, "orc_year_filter") else "")
    ).strip()
    year_num = int(year_txt) if year_txt.isdigit() else None
    conn = None
    try:
        conn = _mysql_connect()
        with conn.cursor() as cur:
            if year_num:
                cur.execute(
                    """
                    SELECT *
                    FROM `orcamentos`
                    WHERE (
                        (`ano` = %s)
                        OR (`ano` IS NULL AND `data` IS NOT NULL AND YEAR(`data`) = %s)
                        OR (`ano` IS NULL AND `data` IS NULL AND `numero` LIKE %s)
                    )
                    ORDER BY `numero` DESC
                    """,
                    (year_num, year_num, f"ORC-{year_num}-%"),
                )
            else:
                cur.execute("SELECT * FROM `orcamentos` ORDER BY `numero` DESC")
            orc_rows = cur.fetchall() or []

            orc_nums = [str(r.get("numero", "") or "") for r in orc_rows if str(r.get("numero", "") or "").strip()]
            linhas_rows = []
            chunk_size = 300
            for i in range(0, len(orc_nums), chunk_size):
                chunk = orc_nums[i : i + chunk_size]
                if not chunk:
                    continue
                placeholders = ",".join(["%s"] * len(chunk))
                cur.execute(
                    f"SELECT * FROM `orcamento_linhas` WHERE `orcamento_numero` IN ({placeholders}) ORDER BY `id`",
                    tuple(chunk),
                )
                linhas_rows.extend(cur.fetchall() or [])

            cur.execute(
                """
                SELECT MAX(CAST(SUBSTRING_INDEX(`numero`, '-', -1) AS UNSIGNED)) AS max_seq
                FROM `orcamentos`
                WHERE `numero` REGEXP '^ORC-[0-9]{4}-[0-9]+$'
                """
            )
            seq_row = cur.fetchone() or {}

        orc_linhas_map = {}
        for l in linhas_rows:
            num = str(l.get("orcamento_numero", "") or "")
            orc_linhas_map.setdefault(num, []).append(
                {
                    "ref_interna": str(l.get("ref_interna", "") or ""),
                    "ref_externa": str(l.get("ref_externa", "") or ""),
                    "descricao": str(l.get("descricao", "") or ""),
                    "material": str(l.get("material", "") or ""),
                    "espessura": str(l.get("espessura", "") or ""),
                    "operacao": str(l.get("operacao", "") or ""),
                    "of": str(l.get("of_codigo", "") or ""),
                    "tempo_peca_min": _to_num(l.get("tempo_peca_min")) or 0.0,
                    "qtd": _to_num(l.get("qtd")) or 0.0,
                    "preco_unit": _to_num(l.get("preco_unit")) or 0.0,
                    "total": _to_num(l.get("total")) or 0.0,
                    "desenho": str(l.get("desenho_path", "") or ""),
                }
            )

        new_orc = []
        for o in orc_rows:
            num = str(o.get("numero", "") or "")
            new_orc.append(
                {
                    "numero": num,
                    "data": _db_to_iso(o.get("data")),
                    "estado": str(o.get("estado", "") or ""),
                    "cliente": _normalize_orc_cliente(str(o.get("cliente_codigo", "") or ""), self.data),
                    "linhas": orc_linhas_map.get(num, []),
                    "iva_perc": _to_num(o.get("iva_perc")) or 0.0,
                    "preco_transporte": _to_num(o.get("preco_transporte")) or 0.0,
                    "subtotal": _to_num(o.get("subtotal")) or 0.0,
                    "total": _to_num(o.get("total")) or 0.0,
                    "numero_encomenda": str(o.get("numero_encomenda", "") or ""),
                    "ano": int(o.get("ano")) if o.get("ano") not in (None, "") else None,
                    "executado_por": str(o.get("executado_por", "") or ""),
                    "nota_transporte": str(o.get("nota_transporte", "") or ""),
                    "notas_pdf": str(o.get("notas_pdf", "") or ""),
                    "nota_cliente": str(o.get("nota_cliente", "") or ""),
                }
            )
        if year_num:
            year_key = str(year_num)
            retained = []
            for old in self.data.get("orcamentos", []):
                if not isinstance(old, dict):
                    continue
                old_year = _orc_extract_year(old.get("data", ""), old.get("numero", ""), old.get("ano"))
                if old_year != year_key:
                    retained.append(old)
            self.data["orcamentos"] = retained + new_orc
        else:
            self.data["orcamentos"] = new_orc

        max_seq = seq_row.get("max_seq") if isinstance(seq_row, dict) else (seq_row[0] if seq_row else None)
        try:
            self.data["orc_seq"] = int(max_seq or 0) + 1
        except Exception:
            self.data["orc_seq"] = max(int(self.data.get("orc_seq", 1)), 1)
        return True
    except Exception:
        return False
    finally:
        try:
            if conn is not None:
                conn.close()
        except Exception:
            pass

def refresh_orc_list(self):
    _ensure_configured()
    self._orc_refreshing = True
    children = self.orc_tbl.get_children()
    if children:
        self.orc_tbl.delete(*children)
    # garantir tags antes de inserir
    self.orc_tbl.tag_configure("even", background="#fbecee")
    self.orc_tbl.tag_configure("odd", background="#fff5f6")
    self.orc_tbl.tag_configure("orc_edicao", background="#fde2e4")
    self.orc_tbl.tag_configure("orc_enviado", background="#fff6cc")
    self.orc_tbl.tag_configure("orc_aprovado", background="#c8e6c9")
    self.orc_tbl.tag_configure("orc_rejeitado", background="#ffcdd2")
    self.orc_tbl.tag_configure("orc_convertido", background="#c8e6c9")
    selected = getattr(self, "selected_orc_numero", None)
    q = (self.orc_search.get().strip().lower() if hasattr(self, "orc_search") else "")
    f = (self.orc_state_filter.get().strip().lower() if hasattr(self, "orc_state_filter") else "ativas")
    y = (self.orc_year_filter.get().strip() if hasattr(self, "orc_year_filter") else "Todos")
    y_norm = y.lower()
    try:
        if hasattr(self, "orc_state_segment") and self.orc_state_segment.winfo_exists():
            self.orc_state_segment.set(self.orc_state_filter.get() or "Ativas")
    except Exception:
        pass
    for idx, o in enumerate(self.data.get("orcamentos", [])):
        if not isinstance(o, dict):
            continue
        total = o.get("total", 0)
        cli = _normalize_orc_cliente(o.get("cliente", {}), self.data)
        o["cliente"] = cli
        if cli.get("nome"):
            cli_nome = f"{cli.get('codigo','')} - {cli.get('nome','')}".strip(" - ")
        else:
            cli_nome = cli.get("codigo", "")
        estado = o.get("estado", "")
        estado_norm = estado.strip().lower()
        row_year = _orc_extract_year(o.get("data", ""), o.get("numero", ""), o.get("ano"))
        if y_norm not in ("todos", "todas", "all") and y and row_year != y:
            continue
        if q:
            try:
                ttxt = f"{float(total):.2f}"
            except Exception:
                ttxt = str(total)
            hay = " ".join([str(o.get("numero", "")), str(cli_nome), str(estado), str(o.get("numero_encomenda", "")), ttxt]).lower()
            if q not in hay:
                continue
        if f and f not in ("todos", "todas", "all"):
            if "ativ" in f:
                if "rejeitado" in estado_norm or "convertido" in estado_norm:
                    continue
            elif "edi" in f and "edi" not in estado_norm:
                continue
            elif "enviado" in f and "enviado" not in estado_norm:
                continue
            elif "aprovado" in f and "aprovado" not in estado_norm:
                continue
            elif "rejeitado" in f and "rejeitado" not in estado_norm:
                continue
            elif "convertido" in f and "convertido" not in estado_norm:
                continue
        if "edi" in estado_norm:
            estado_tag = "orc_edicao"
        elif "enviado" in estado_norm:
            estado_tag = "orc_enviado"
        elif "aprovado" in estado_norm:
            estado_tag = "orc_aprovado"
        elif "rejeitado" in estado_norm:
            estado_tag = "orc_rejeitado"
        elif "convertido" in estado_norm:
            estado_tag = "orc_convertido"
        else:
            estado_tag = "odd" if idx % 2 else "even"
        item_id = self.orc_tbl.insert("", END, values=(
            o.get("numero", ""), cli_nome, estado, o.get("numero_encomenda", ""), f"{total:.2f}"
        ), tags=(estado_tag,))
        if selected and o.get("numero") == selected:
            self.orc_tbl.selection_set(item_id)
    self._orc_refreshing = False
    try:
        if hasattr(self, "nb") and hasattr(self, "tab_menu"):
            if str(self.nb.select()) == str(self.tab_menu):
                self.debounce_call("menu_stats", self.refresh_menu_dashboard_stats, delay_ms=250)
    except Exception:
        pass

def _on_orc_filter_click(self, value=None):
    _ensure_configured()
    try:
        if value is not None:
            self.orc_state_filter.set(value)
    except Exception:
        pass
    self.refresh_orc_list()


def _on_orc_year_change(self, value=None):
    _ensure_configured()
    try:
        if value is not None and hasattr(self, "orc_year_filter"):
            self.orc_year_filter.set(str(value))
    except Exception:
        pass
    # se houver edicao pendente, tenta guardar antes de trocar o dataset do ano
    try:
        if getattr(self, "selected_orc_numero", None):
            self.save_orc_fields(refresh_list=False)
    except Exception:
        pass
    reloaded = _reload_orcamentos_from_mysql(self)
    try:
        self.refresh_orc_year_options(keep_selection=True)
    except Exception:
        pass
    if reloaded:
        self.clear_orc_details()
    self.refresh_orc_list()

def clear_orc_details(self):
    _ensure_configured()
    self.selected_orc_numero = None
    for v in self.orc_cliente_vars.values():
        v.set("")
    if hasattr(self, "orc_nota_transporte"):
        self.orc_nota_transporte.set("")
    self._orc_set_notes_text("")
    children = self.orc_linhas.get_children()
    if children:
        self.orc_linhas.delete(*children)
    self.orc_subtotal.set("0.00")
    self.orc_total.set("0.00")

def on_orc_select(self, _=None):
    _ensure_configured()
    if self._orc_refreshing:
        return
    if self.selected_orc_numero:
        try:
            self.save_orc_fields(refresh_list=False)
        except Exception:
            # não bloquear mudança de seleção por dados temporariamente inválidos
            pass
    sel = self.orc_tbl.selection()
    if not sel:
        return
    current_item = None
    focus_item = self.orc_tbl.focus()
    if focus_item in sel:
        current_item = focus_item
    if not current_item:
        current_item = sel[-1]
    numero = self.orc_tbl.item(current_item, "values")[0]
    self.selected_orc_numero = numero
    orc = self.get_orc_by_numero(numero)
    if not orc:
        return
    cli = _normalize_orc_cliente(orc.get("cliente", {}), self.data)
    orc["cliente"] = cli
    for k in self.orc_cliente_vars:
        self.orc_cliente_vars[k].set(cli.get(k, ""))
    if hasattr(self, "orc_executado"):
        self.orc_executado.set(orc.get("executado_por", ""))
    if hasattr(self, "orc_nota_transporte"):
        self.orc_nota_transporte.set(orc.get("nota_transporte", ""))
    self._orc_set_notes_text(orc.get("notas_pdf", ""))
    try:
        iva_val = float(orc.get("iva_perc", 23.0))
    except Exception:
        iva_val = 23.0
    self.orc_iva.set(iva_val)
    self.refresh_orc_linhas(orc)

def refresh_orc_linhas(self, orc):
    _ensure_configured()
    children = self.orc_linhas.get_children()
    if children:
        self.orc_linhas.delete(*children)
    def fmt_num(val, decimals=2):
        try:
            v = float(val)
        except Exception:
            return str(val) if val is not None else ""
        s = f"{v:.{decimals}f}"
        if "." in s:
            s = s.rstrip("0").rstrip(".")
        return s
    for idx, l in enumerate(orc.get("linhas", [])):
        self.orc_linhas.insert("", END, values=(
            l.get("ref_interna", ""), l.get("ref_externa", ""), l.get("descricao", ""), l.get("material", ""),
            fmt_num(l.get("espessura", "")),
            l.get("operacao", ""),
            fmt_num(l.get("qtd", 0), decimals=2),
            fmt_num(l.get("preco_unit", 0), decimals=2),
            fmt_num(l.get("total", 0), decimals=2),
        ), tags=("odd" if idx % 2 else "even",))
    self.orc_linhas.tag_configure("even", background="#eef7ff")
    self.orc_linhas.tag_configure("odd", background="#fff8f9")
    self.recalc_orc(orc)

def add_orcamento(self):
    _ensure_configured()
    numero = next_orc_numero(self.data)
    orc = {
        "numero": numero,
        "data": now_iso(),
        "estado": "Em edição",
        "cliente": {},
        "linhas": [],
        "iva_perc": 23.0,
        "subtotal": 0.0,
        "total": 0.0,
        "numero_encomenda": "",
        "executado_por": "",
        "nota_transporte": "",
        "notas_pdf": "",
    }
    self.data.setdefault("orcamentos", []).append(orc)
    if not _save_orc_fast(self, orc):
        save_data(self.data)
    if hasattr(self, "orc_year_filter"):
        self.orc_year_filter.set(str(datetime.now().year))
    try:
        self.refresh_orc_year_options(keep_selection=True)
    except Exception:
        pass
    self.selected_orc_numero = numero
    self.refresh_orc_list()
    self.on_orc_select()

def remove_orcamento(self):
    _ensure_configured()
    sel = self.orc_tbl.selection()
    if not sel:
        return
    numero = self.orc_tbl.item(sel[0], "values")[0]
    if not messagebox.askyesno("Confirmar", "Remover orçamento?"):
        return
    self.data["orcamentos"] = [o for o in self.data.get("orcamentos", []) if o.get("numero") != numero]
    removed_mysql = False
    if USE_MYSQL_STORAGE and MYSQL_AVAILABLE:
        conn = None
        try:
            conn = _mysql_connect()
            with conn.cursor() as cur:
                cur.execute("DELETE FROM `orcamento_linhas` WHERE `orcamento_numero`=%s", (numero,))
                cur.execute("DELETE FROM `orcamentos` WHERE `numero`=%s", (numero,))
            conn.commit()
            removed_mysql = True
        except Exception:
            try:
                if conn:
                    conn.rollback()
            except Exception:
                pass
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass
    if not removed_mysql:
        save_data(self.data)
    self.clear_orc_details()
    try:
        self.refresh_orc_year_options(keep_selection=True)
    except Exception:
        pass
    self.refresh_orc_list()

def set_orc_estado(self, estado):
    _ensure_configured()
    numero = self.selected_orc_numero
    if not numero and hasattr(self, "orc_tbl"):
        sel = self.orc_tbl.selection()
        if sel:
            numero = self.orc_tbl.item(sel[0], "values")[0]
            self.selected_orc_numero = numero
    orc = self.get_orc_by_numero(numero) if numero else None
    if not orc:
        return
    orc["estado"] = estado
    if not _save_orc_fast(self, orc):
        save_data(self.data)
    self.refresh_orc_list()
    try:
        self.on_orc_select()
    except Exception:
        pass

def save_orc_fields(self, refresh_list=True):
    _ensure_configured()
    orc = self.get_orc_by_numero(self.selected_orc_numero) if getattr(self, "selected_orc_numero", None) else None
    if not orc:
        return
    cli = {k: v.get().strip() for k, v in self.orc_cliente_vars.items()}
    orc["cliente"] = cli
    try:
        iva_raw = self.orc_iva.get()
    except Exception:
        iva_raw = 0
    try:
        iva = float(str(iva_raw).replace(",", ".") or 0)
    except Exception:
        iva = 0.0
    orc["iva_perc"] = iva
    if hasattr(self, "orc_executado"):
        orc["executado_por"] = self.orc_executado.get().strip()
    if hasattr(self, "orc_nota_transporte"):
        orc["nota_transporte"] = (self.orc_nota_transporte.get() or "").strip()
    orc["notas_pdf"] = self._orc_get_notes_text()
    self.recalc_orc(orc)
    if not _save_orc_fast(self, orc):
        save_data(self.data)
    if refresh_list:
        self.refresh_orc_list()

def fill_orc_from_cliente(self):
    _ensure_configured()
    raw = self.orc_cliente_vars["codigo"].get().strip()
    codigo = raw.split(" - ")[0] if " - " in raw else raw
    if not codigo:
        return
    c = find_cliente(self.data, codigo)
    if not c:
        return
    self.orc_cliente_vars["codigo"].set(c.get("codigo", ""))
    self.orc_cliente_vars["nome"].set(c.get("nome", ""))
    self.orc_cliente_vars["empresa"].set(c.get("nome", ""))
    self.orc_cliente_vars["nif"].set(c.get("nif", ""))
    self.orc_cliente_vars["morada"].set(c.get("morada", ""))
    self.orc_cliente_vars["contacto"].set(c.get("contacto", ""))
    self.orc_cliente_vars["email"].set(c.get("email", ""))
    self.save_orc_fields()

def add_orc_linha(self):
    _ensure_configured()
    self.open_orc_linha()

def edit_orc_linha(self):
    _ensure_configured()
    orc = self.get_orc_by_numero(self.selected_orc_numero) if getattr(self, "selected_orc_numero", None) else None
    if not orc:
        messagebox.showerror("Erro", "Selecione um orçamento.")
        return
    sel = self.orc_linhas.selection()
    if not sel:
        messagebox.showerror("Erro", "Selecione uma linha.")
        return
    idx = self.orc_linhas.index(sel[0])
    self.open_orc_linha(edit_index=idx)

def remove_orc_linha(self):
    _ensure_configured()
    orc = self.get_orc_by_numero(self.selected_orc_numero) if getattr(self, "selected_orc_numero", None) else None
    if not orc:
        messagebox.showerror("Erro", "Selecione um orçamento.")
        return
    sel = self.orc_linhas.selection()
    if not sel:
        messagebox.showerror("Erro", "Selecione uma linha.")
        return
    idx = self.orc_linhas.index(sel[0])
    orc["linhas"].pop(idx)
    self.refresh_orc_linhas(orc)
    if not _save_orc_fast(self, orc):
        save_data(self.data)

def open_orc_linha_desenho(self):
    _ensure_configured()
    numero = getattr(self, "selected_orc_numero", None)
    if not numero and hasattr(self, "orc_tbl"):
        sel_orc = self.orc_tbl.selection()
        if sel_orc:
            numero = self.orc_tbl.item(sel_orc[0], "values")[0]
            self.selected_orc_numero = numero
    orc = self.get_orc_by_numero(numero) if numero else None
    if not orc:
        messagebox.showerror("Erro", "Selecione um orçamento.")
        return
    sel = self.orc_linhas.selection()
    if not sel:
        messagebox.showerror("Erro", "Selecione uma linha.")
        return
    idx = self.orc_linhas.index(sel[0])
    if idx < 0 or idx >= len(orc.get("linhas", [])):
        return
    path = str(orc["linhas"][idx].get("desenho", "") or "").strip()
    if not path:
        messagebox.showinfo("Info", "Esta linha não tem desenho associado.")
        return
    if not os.path.exists(path):
        messagebox.showerror("Erro", f"Ficheiro não encontrado:\n{path}")
        return
    try:
        os.startfile(path)
    except Exception as ex:
        messagebox.showerror("Erro", f"Não foi possível abrir o desenho.\n{ex}")

def recalc_orc(self, orc):
    _ensure_configured()
    subtotal = 0.0
    for l in orc.get("linhas", []):
        qtd = float(l.get("qtd", 0) or 0)
        preco = float(l.get("preco_unit", 0) or 0)
        l["total"] = qtd * preco
        subtotal += l["total"]
    iva = subtotal * (float(orc.get("iva_perc", 0) or 0) / 100.0)
    total = subtotal + iva
    orc["subtotal"] = subtotal
    orc["total"] = total
    self.orc_subtotal.set(f"{subtotal:.2f}")
    self.orc_total.set(f"{total:.2f}")

def _orc_get_notes_text(self):
    _ensure_configured()
    if not hasattr(self, "orc_notas_text"):
        return ""
    try:
        if self.orc_use_custom and CUSTOM_TK_AVAILABLE:
            return (self.orc_notas_text.get("1.0", "end") or "").strip()
        return (self.orc_notas_text.get("1.0", END) or "").strip()
    except Exception:
        return ""

def _orc_set_notes_text(self, text):
    _ensure_configured()
    if not hasattr(self, "orc_notas_text"):
        return
    if isinstance(text, (list, tuple)):
        text = "\n".join([str(x).strip() for x in text if str(x).strip()])
    txt = str(text or "").strip()
    try:
        if self.orc_use_custom and CUSTOM_TK_AVAILABLE:
            self.orc_notas_text.delete("1.0", "end")
            if txt:
                self.orc_notas_text.insert("1.0", txt)
        else:
            self.orc_notas_text.delete("1.0", END)
            if txt:
                self.orc_notas_text.insert("1.0", txt)
    except Exception:
        pass

def _orc_append_pdf_note(self, line):
    _ensure_configured()
    line = (line or "").strip()
    if not line:
        return
    cur = self._orc_get_notes_text()
    lines = [x.strip() for x in cur.splitlines() if x.strip()]
    if line not in lines:
        lines.append(line)
    self._orc_set_notes_text("\n".join(lines))

def _extract_orc_operacoes(self, orc=None):
    _ensure_configured()
    ops = []
    if hasattr(self, "orc_linhas"):
        try:
            for iid in self.orc_linhas.get_children():
                vals = self.orc_linhas.item(iid, "values")
                if len(vals) >= 6:
                    op = str(vals[5] or "").strip()
                    if op:
                        ops.append(op)
        except Exception:
            pass
    if not ops and orc:
        for l in list(orc.get("linhas", [])):
            op = (
                l.get("operacao")
                or l.get("operacoes")
                or l.get("Operacoes")
                or l.get("Operações")
                or ""
            )
            op = str(op).strip()
            if op:
                ops.append(op)
    return ops


def _orc_line_type_value(line):
    _ensure_configured()
    try:
        return normalize_orc_line_type((line or {}).get("tipo_item"))
    except Exception:
        return str((line or {}).get("tipo_item", "") or "").strip() or "peca_fabricada"


def _orc_line_type_title(line):
    _ensure_configured()
    try:
        return orc_line_type_label(_orc_line_type_value(line))
    except Exception:
        kind = _orc_line_type_value(line)
        if kind == "produto_stock":
            return "Produto stock"
        if kind == "servico_montagem":
            return "Montagem"
        return "Peca fabricada"


def _orc_line_ref_display(line):
    _ensure_configured()
    row = dict(line or {})
    kind = _orc_line_type_value(row)
    if kind == "produto_stock":
        return str(row.get("produto_codigo", "") or row.get("ref_externa", "") or "").strip() or "-"
    return str(row.get("ref_externa", "") or row.get("ref_interna", "") or "").strip() or "-"


def _orc_line_material_display(line):
    _ensure_configured()
    row = dict(line or {})
    kind = _orc_line_type_value(row)
    if kind == "produto_stock":
        return str(row.get("produto_codigo", "") or "").strip() or "Stock"
    if kind == "servico_montagem":
        return "Servico"
    return str(row.get("material", "") or "").strip() or "-"


def _orc_line_unit_display(line):
    _ensure_configured()
    row = dict(line or {})
    kind = _orc_line_type_value(row)
    if kind in ("produto_stock", "servico_montagem"):
        return str(row.get("produto_unid", "") or "").strip() or ("SV" if kind == "servico_montagem" else "UN")
    return fmt_num(row.get("espessura", ""))


def _orc_line_operacao_display(line):
    _ensure_configured()
    row = dict(line or {})
    operacao = str(row.get("operacao", "") or "").strip() or "-"
    conjunto = str(row.get("conjunto_codigo", "") or row.get("conjunto_nome", "") or "").strip()
    if conjunto:
        return f"{operacao} | {conjunto}"
    return operacao


def _orc_line_description_display(line):
    _ensure_configured()
    row = dict(line or {})
    desc = str(row.get("descricao", "") or "").strip() or "-"
    kind = _orc_line_type_value(row)
    if kind == "peca_fabricada":
        return desc
    return f"{_orc_line_type_title(row)}: {desc}"

def _build_orc_notes_lines(self, orc):
    _ensure_configured()
    ops_text = " ".join(norm_text(x) for x in self._extract_orc_operacoes(orc))
    rows = list((orc or {}).get("linhas", []) or [])
    notes = []
    seen = set()

    def push(line):
        line = (line or "").strip()
        if not line:
            return
        key = norm_text(line)
        if key in seen:
            return
        seen.add(key)
        notes.append(line)

    push("PROPOSTA RETIFICADA PARA ESPESSURAS DEFINIDAS PELO CLIENTE.")
    transp = (orc.get("nota_transporte", "") or "").strip()
    transp_price = float(orc.get("preco_transporte", 0) or 0)
    if transp:
        t = transp.rstrip(".")
        if "transporte" in norm_text(t):
            push(f"- {t}.")
        else:
            push(f"- Transporte: {t}.")
    if transp_price > 0:
        push(f"- Preco de transporte considerado: {transp_price:.2f} EUR.")

    map_ops = [
        ("laser", "Corte Laser"),
        ("quin", "Quinagem"),
        ("rosc", "Roscagem"),
        ("furo", "Furo Manual"),
        ("sold", "Soldadura"),
        ("mont", "Montagem"),
    ]
    for key, title in map_ops:
        if key in ops_text:
            push(f"- Foi considerado: {title}.")
    if any(_orc_line_type_value(row) == "produto_stock" for row in rows):
        push("- Inclui componentes de stock com baixa prevista no momento de montagem.")
    if any(_orc_line_type_value(row) == "servico_montagem" for row in rows):
        push("- Inclui servico de montagem final e fecho do conjunto.")
    conjuntos = sorted(
        {
            str((row or {}).get("conjunto_nome", "") or (row or {}).get("conjunto_codigo", "") or "").strip()
            for row in rows
            if str((row or {}).get("conjunto_nome", "") or (row or {}).get("conjunto_codigo", "") or "").strip()
        }
    )
    for conjunto in conjuntos[:3]:
        push(f"- Conjunto parametrizado considerado: {conjunto}.")
    extra = (orc.get("notas_pdf", "") or "").strip()
    if extra:
        for ln in extra.splitlines():
            push(ln)
    return notes

def orc_fill_notes_by_ops(self):
    _ensure_configured()
    orc = self.get_orc_by_numero(self.selected_orc_numero) if getattr(self, "selected_orc_numero", None) else None
    ops_text = " ".join(norm_text(x) for x in self._extract_orc_operacoes(orc))
    map_ops = [
        ("laser", "Corte Laser"),
        ("quin", "Quinagem"),
        ("rosc", "Roscagem"),
        ("furo", "Furo Manual"),
        ("sold", "Soldadura"),
    ]
    op_lines = []
    for key, title in map_ops:
        if key in ops_text:
            op_lines.append(f"- Foi considerado: {title}.")
    current = [x.strip() for x in self._orc_get_notes_text().splitlines() if x.strip()]
    keep = [x for x in current if "foi considerado" not in norm_text(x)]
    merged = op_lines + keep
    self._orc_set_notes_text("\n".join(merged))

def open_orc_linha(self, edit_index=None):
    _ensure_configured()
    numero = getattr(self, "selected_orc_numero", None)
    if not numero and hasattr(self, "orc_tbl"):
        sel_orc = self.orc_tbl.selection()
        if sel_orc:
            numero = self.orc_tbl.item(sel_orc[0], "values")[0]
            self.selected_orc_numero = numero
    if not numero and hasattr(self, "orc_tbl"):
        items = self.orc_tbl.get_children()
        if items:
            first = items[0]
            self.orc_tbl.selection_set(first)
            self.orc_tbl.focus(first)
            self.orc_tbl.see(first)
            numero = self.orc_tbl.item(first, "values")[0]
            self.selected_orc_numero = numero
            try:
                self.on_orc_select()
            except Exception:
                pass
    if not numero:
        if messagebox.askyesno("Orçamentos", "Não existe orçamento selecionado. Quer criar um novo agora?"):
            self.add_orcamento()
            numero = getattr(self, "selected_orc_numero", None)
    orc = self.get_orc_by_numero(numero) if numero else None
    if not orc:
        messagebox.showerror("Erro", "Não foi possível abrir a linha sem um orçamento selecionado.")
        return
    refs_db = self.data.get("orc_refs", {})
    use_custom = self.orc_use_custom and CUSTOM_TK_AVAILABLE
    win = ctk.CTkToplevel(self.root) if use_custom else Toplevel(self.root)
    win.title("Linha de Orçamento")
    win.geometry("820x560")
    win.grid_columnconfigure(1, weight=1)
    Lbl = ctk.CTkLabel if use_custom else ttk.Label
    Ent = ctk.CTkEntry if use_custom else ttk.Entry
    Btn = ctk.CTkButton if use_custom else ttk.Button
    Cmb = ctk.CTkComboBox if use_custom else ttk.Combobox
    Chk = ctk.CTkCheckBox if use_custom else ttk.Checkbutton
    vars_ = {
        "ref_int": StringVar(),
        "ref_ext": StringVar(),
        "descricao": StringVar(),
        "material": StringVar(),
        "espessura": StringVar(),
        "operacao": StringVar(),
        "qtd": DoubleVar(value=1),
        "preco": DoubleVar(value=0),
        "desenho": StringVar(),
    }
    if edit_index is not None:
        l = orc["linhas"][edit_index]
        vars_["ref_int"].set(l.get("ref_interna", ""))
        vars_["ref_ext"].set(l.get("ref_externa", ""))
        vars_["descricao"].set(l.get("descricao", ""))
        vars_["material"].set(l.get("material", ""))
        vars_["espessura"].set(l.get("espessura", ""))
        vars_["operacao"].set(" + ".join(parse_operacoes_lista(l.get("operacao", ""))))
        vars_["qtd"].set(l.get("qtd", 0))
        vars_["preco"].set(l.get("preco_unit", 0))
        vars_["desenho"].set(l.get("desenho", ""))
    else:
        cli_code = ""
        cli_val = self.orc_cliente_vars.get("codigo").get() if self.orc_cliente_vars.get("codigo") else ""
        if cli_val:
            cli_code = cli_val.split(" - ")[0]
        if cli_code:
            existing = [l.get("ref_interna", "") for l in orc.get("linhas", [])]
            vars_["ref_int"].set(next_ref_interna_unique(self.data, cli_code, existing))

    def on_ref_pick(_=None):
        ref_ext = vars_["ref_ext"].get().strip()
        if ref_ext in refs_db:
            r = refs_db[ref_ext]
            if r.get("ref_interna"):
                vars_["ref_int"].set(r.get("ref_interna", ""))
            vars_["descricao"].set(r.get("descricao", ""))
            vars_["material"].set(r.get("material", ""))
            vars_["espessura"].set(r.get("espessura", ""))
            vars_["preco"].set(r.get("preco_unit", 0))
            vars_["operacao"].set(" + ".join(parse_operacoes_lista(r.get("operacao", ""))))
            vars_["desenho"].set(r.get("desenho", ""))

    def pick_desenho():
        path = filedialog.askopenfilename(
            title="Selecionar desenho do cliente",
            filetypes=[
                ("Desenhos/PDF", "*.pdf *.dwg *.dxf *.step *.stp *.iges *.igs"),
                ("Imagens", "*.png *.jpg *.jpeg *.bmp"),
                ("Todos", "*.*"),
            ],
        )
        if path:
            vars_["desenho"].set(path)

    def open_desenho():
        path = vars_["desenho"].get().strip()
        if not path:
            messagebox.showinfo("Info", "Sem desenho associado.")
            return
        if not os.path.exists(path):
            messagebox.showerror("Erro", f"Ficheiro não encontrado:\n{path}")
            return
        try:
            os.startfile(path)
        except Exception as ex:
            messagebox.showerror("Erro", f"Não foi possível abrir o desenho.\n{ex}")

    def on_materia_pick():
        mp = self._dialog_escolher_materia_prima_ne()
        if not mp:
            return
        vars_["descricao"].set(mp.get("descricao", ""))
        vars_["material"].set(mp.get("material", ""))
        vars_["espessura"].set(mp.get("espessura", ""))
        vars_["preco"].set(parse_float(mp.get("preco_unid", 0), 0))
        if not vars_["ref_ext"].get().strip():
            vars_["ref_ext"].set(mp.get("id", ""))

    def open_ref_history():
        win2 = ctk.CTkToplevel(self.root) if use_custom else Toplevel(self.root)
        win2.title("Histórico de Referências")
        if use_custom:
            win2.geometry("900x520")
            try:
                win2.configure(fg_color="#f7f8fb")
            except Exception:
                pass
        else:
            win2.geometry("860x500")
        wrap2 = (
            ctk.CTkFrame(
                win2,
                fg_color="#ffffff",
                corner_radius=10,
                border_width=1,
                border_color="#e7cfd3",
            )
            if use_custom
            else ttk.Frame(win2)
        )
        wrap2.pack(fill="both", expand=True, padx=10, pady=10)

        top = ctk.CTkFrame(wrap2, fg_color="transparent") if use_custom else ttk.Frame(wrap2)
        top.pack(fill="x", padx=8, pady=(8, 6))
        (ctk.CTkLabel if use_custom else ttk.Label)(
            top,
            text="Histórico de Referências",
            font=("Segoe UI", 13, "bold") if use_custom else None,
            text_color="#7f1b2c" if use_custom else None,
        ).pack(side="left", padx=(0, 14))
        (ctk.CTkLabel if use_custom else ttk.Label)(top, text="Pesquisar").pack(side="left")
        q = StringVar()
        ent = (ctk.CTkEntry(top, textvariable=q, width=260) if use_custom else ttk.Entry(top, textvariable=q, width=40))
        ent.pack(side="left", padx=6)
        tbl_style = ""
        if use_custom:
            sty = ttk.Style()
            sty.configure(
                "ORCREF.Treeview",
                font=("Segoe UI", 10),
                rowheight=27,
                background="#f8fbff",
                fieldbackground="#f8fbff",
                borderwidth=0,
            )
            sty.configure(
                "ORCREF.Treeview.Heading",
                font=("Segoe UI", 10, "bold"),
                background=THEME_HEADER_BG,
                foreground="white",
                relief="flat",
            )
            sty.map("ORCREF.Treeview.Heading", background=[("active", THEME_HEADER_ACTIVE)])
            sty.map(
                "ORCREF.Treeview",
                background=[("selected", THEME_SELECT_BG)],
                foreground=[("selected", THEME_SELECT_FG)],
            )
            tbl_style = "ORCREF.Treeview"

        tbl_box = ctk.CTkFrame(wrap2, fg_color="#ffffff") if use_custom else ttk.Frame(wrap2)
        tbl_box.pack(fill="both", expand=True, padx=8, pady=(0, 6))
        tbl = ttk.Treeview(
            tbl_box,
            columns=("ref_ext", "ref_int", "desc", "mat", "esp", "preco", "op"),
            show="headings",
            style=tbl_style,
            height=14,
        )
        for col, txt, w in [
            ("ref_ext", "Ref. Externa", 160),
            ("ref_int", "Ref. Interna", 140),
            ("desc", "Descrição", 220),
            ("mat", "Material", 100),
            ("esp", "Espessura", 90),
            ("preco", "Preço", 90),
            ("op", "Operação", 140),
        ]:
            tbl.heading(col, text=txt)
            tbl.column(col, width=w)
        tbl.pack(side="left", fill="both", expand=True, padx=(0, 0), pady=(0, 0))
        if use_custom and CUSTOM_TK_AVAILABLE:
            sb = ctk.CTkScrollbar(tbl_box, orientation="vertical", command=tbl.yview)
        else:
            sb = ttk.Scrollbar(tbl_box, orient="vertical", command=tbl.yview)
        tbl.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        tbl.tag_configure("even", background="#eef6ff")
        tbl.tag_configure("odd", background="#f8fbff")

        def load_rows():
            tbl.delete(*tbl.get_children())
            query = q.get().strip().lower()
            row_i = 0
            for k, r in refs_db.items():
                row = (
                    k, r.get("ref_interna", ""), r.get("descricao", ""), r.get("material", ""), r.get("espessura", ""),
                    r.get("preco_unit", 0), " + ".join(parse_operacoes_lista(r.get("operacao", "")))
                )
                if query and not any(query in str(v).lower() for v in row):
                    continue
                tag = "even" if row_i % 2 == 0 else "odd"
                tbl.insert("", END, values=row, tags=(tag,))
                row_i += 1

        def choose():
            sel = tbl.selection()
            if not sel:
                messagebox.showerror("Erro", "Selecione uma referência")
                return
            vals = tbl.item(sel[0], "values")
            vars_["ref_ext"].set(vals[0])
            vars_["ref_int"].set(vals[1])
            vars_["descricao"].set(vals[2])
            vars_["material"].set(vals[3])
            vars_["espessura"].set(vals[4])
            vars_["preco"].set(float(vals[5] or 0))
            vars_["operacao"].set(" + ".join(parse_operacoes_lista(vals[6])))
            vars_["desenho"].set(refs_db.get(vals[0], {}).get("desenho", ""))
            win2.destroy()

        btns = ctk.CTkFrame(wrap2, fg_color="transparent") if use_custom else ttk.Frame(wrap2)
        btns.pack(fill="x", padx=8, pady=(0, 8))
        (ctk.CTkButton if use_custom else ttk.Button)(btns, text="Selecionar", command=choose, width=130 if use_custom else None).pack(side="left", padx=6)
        (ctk.CTkButton if use_custom else ttk.Button)(btns, text="Cancelar", command=win2.destroy, width=130 if use_custom else None).pack(side="right", padx=6)
        tbl.bind("<Double-Button-1>", lambda _=None: choose())
        ent.bind("<KeyRelease>", lambda _=None: load_rows())
        load_rows()
        ent.focus_set()
        win2.transient(win)
        win2.grab_set()
        try:
            win2.lift()
            win2.focus_force()
        except Exception:
            pass

    Lbl(win, text="Ref. Interna").grid(row=0, column=0, sticky="w", padx=8, pady=6)
    Ent(win, textvariable=vars_["ref_int"], width=(220 if use_custom else 26)).grid(row=0, column=1, padx=8, pady=6, sticky="w")
    btn_ref1 = ctk.CTkFrame(win, fg_color="transparent") if use_custom else ttk.Frame(win)
    btn_ref1.grid(row=0, column=2, columnspan=2, sticky="w", padx=4, pady=6)
    Btn(
        btn_ref1,
        text="Gerar",
        command=lambda: vars_["ref_int"].set(
            next_ref_interna_unique(
                self.data,
                (self.orc_cliente_vars.get("codigo").get().split(" - ")[0] if self.orc_cliente_vars.get("codigo") else ""),
                [l.get("ref_interna", "") for l in orc.get("linhas", [])],
            )
        ),
        width=120 if use_custom else None,
    ).pack(side="left", padx=4)
    Lbl(win, text="Ref. Externa").grid(row=1, column=0, sticky="w", padx=8, pady=6)
    Ent(win, textvariable=vars_["ref_ext"], width=(320 if use_custom else 36)).grid(row=1, column=1, padx=8, pady=6, sticky="w")
    btn_ref = ctk.CTkFrame(win, fg_color="transparent") if use_custom else ttk.Frame(win)
    btn_ref.grid(row=1, column=2, columnspan=2, sticky="w", padx=4, pady=6)
    Btn(btn_ref, text="Histórico", command=open_ref_history, width=120 if use_custom else None).pack(side="left", padx=4)
    Btn(btn_ref, text="Buscar", command=on_ref_pick, width=120 if use_custom else None).pack(side="left", padx=4)
    Btn(btn_ref, text="Matéria-Prima", command=on_materia_pick, width=130 if use_custom else None).pack(side="left", padx=4)
    Lbl(win, text="Desenho Cliente").grid(row=2, column=0, sticky="w", padx=8, pady=6)
    Ent(win, textvariable=vars_["desenho"], width=(500 if use_custom else 48)).grid(row=2, column=1, columnspan=2, padx=8, pady=6, sticky="w")
    draw_btn = ctk.CTkFrame(win, fg_color="transparent") if use_custom else ttk.Frame(win)
    draw_btn.grid(row=2, column=3, sticky="w", padx=4, pady=6)
    Btn(draw_btn, text="Selecionar desenho", command=pick_desenho, width=150 if use_custom else None).pack(side="left", padx=(0, 4))
    Btn(draw_btn, text="Abrir", command=open_desenho, width=90 if use_custom else None).pack(side="left")
    Lbl(win, text="Descrição").grid(row=3, column=0, sticky="w", padx=8, pady=6)
    Ent(win, textvariable=vars_["descricao"], width=(500 if use_custom else 48)).grid(row=3, column=1, columnspan=3, padx=8, pady=6, sticky="w")
    Lbl(win, text="Material").grid(row=4, column=0, sticky="w", padx=8, pady=6)
    mat_opts = list(dict.fromkeys(MATERIAIS_PRESET + self.data.get("materiais_hist", []) + list_unique(self.data, "material")))
    if use_custom:
        Cmb(win, variable=vars_["material"], values=mat_opts, width=210).grid(row=4, column=1, padx=8, pady=6, sticky="w")
    else:
        Cmb(win, textvariable=vars_["material"], values=mat_opts, width=18, state="normal").grid(row=4, column=1, padx=8, pady=6, sticky="w")
    Lbl(win, text="Espessura").grid(row=4, column=2, sticky="w", padx=8, pady=6)
    esp_opts = [str(v).rstrip('0').rstrip('.') if isinstance(v, float) else str(v) for v in ESPESSURAS_PRESET]
    esp_hist = [str(v) for v in self.data.get("espessuras_hist", [])]
    if use_custom:
        Cmb(win, variable=vars_["espessura"], values=list(dict.fromkeys(esp_opts + esp_hist)), width=130).grid(row=4, column=3, padx=8, pady=6, sticky="w")
    else:
        Cmb(win, textvariable=vars_["espessura"], values=list(dict.fromkeys(esp_opts + esp_hist)), width=10, state="normal").grid(row=4, column=3, padx=8, pady=6, sticky="w")
    Lbl(win, text="Operação").grid(row=5, column=0, sticky="w", padx=8, pady=6)
    Ent(win, textvariable=vars_["operacao"], width=(420 if use_custom else 44)).grid(row=5, column=1, columnspan=2, padx=8, pady=6, sticky="w")
    op_btns = ctk.CTkFrame(win, fg_color="transparent") if use_custom else ttk.Frame(win)
    op_btns.grid(row=5, column=3, sticky="w", padx=4, pady=6)
    Btn(
        op_btns,
        text="Selecionar",
        command=lambda: (
            lambda val: vars_["operacao"].set(val)
            if val is not None else None
        )(self.escolher_operacoes_fluxo(vars_["operacao"].get(), parent=win)),
        width=110 if use_custom else None,
    ).pack(side="left", padx=(0, 4))
    Btn(op_btns, text="Padrão", command=lambda: vars_["operacao"].set(OFF_OPERACAO_OBRIGATORIA), width=90 if use_custom else None).pack(side="left")
    Lbl(win, text="Quantidade").grid(row=6, column=0, sticky="w", padx=8, pady=6)
    Ent(win, textvariable=vars_["qtd"], width=(140 if use_custom else None)).grid(row=6, column=1, padx=8, pady=6, sticky="w")
    Lbl(win, text="Preço Unitário (€)").grid(row=7, column=0, sticky="w", padx=8, pady=6)
    Ent(win, textvariable=vars_["preco"], width=(140 if use_custom else None)).grid(row=7, column=1, padx=8, pady=6, sticky="w")
    keep_ref = BooleanVar(value=True)
    Chk(win, text="Guardar referência na base", variable=keep_ref).grid(row=8, column=0, columnspan=2, sticky="w", padx=8, pady=6)

    def save_line():
        ref_int = vars_["ref_int"].get().strip()
        ref_ext = vars_["ref_ext"].get().strip()
        desc = vars_["descricao"].get().strip()
        mat = vars_["material"].get().strip()
        esp = vars_["espessura"].get().strip()
        qtd = parse_float(vars_["qtd"].get(), 0.0)
        preco = parse_float(vars_["preco"].get(), 0.0)
        if qtd <= 0:
            messagebox.showerror("Erro", "Quantidade inválida.")
            return
        if preco < 0:
            messagebox.showerror("Erro", "Preço inválido.")
            return
        operacao = " + ".join(parse_operacoes_lista(vars_["operacao"].get()))
        desenho = vars_["desenho"].get().strip()
        total = qtd * preco
        of_val = orc["linhas"][edit_index].get("of", "") if edit_index is not None else next_of_numero(self.data)
        line = {
            "ref_interna": ref_int,
            "ref_externa": ref_ext,
            "descricao": desc,
            "material": mat,
            "espessura": esp,
            "operacao": operacao,
            "of": of_val,
            "qtd": qtd,
            "preco_unit": preco,
            "total": total,
            "desenho": desenho,
        }
        # impedir referência interna duplicada no mesmo orçamento
        if ref_int:
            for i, existing in enumerate(orc.get("linhas", [])):
                if edit_index is not None and i == edit_index:
                    continue
                if existing.get("ref_interna") == ref_int:
                    cli_val = self.orc_cliente_vars.get("codigo").get() if self.orc_cliente_vars.get("codigo") else ""
                    cli_code = cli_val.split(" - ")[0] if cli_val else ""
                    if cli_code:
                        suggested = next_ref_interna_unique(
                            self.data, cli_code, [l.get("ref_interna", "") for l in orc.get("linhas", [])]
                        )
                        vars_["ref_int"].set(suggested)
                        messagebox.showerror(
                            "Erro",
                            f"Referência interna já existe neste orçamento. Nova sugerida: {suggested}",
                        )
                    else:
                        messagebox.showerror("Erro", "Referência interna já existe neste orçamento.")
                    return
        if edit_index is None:
            orc["linhas"].append(line)
        else:
            orc["linhas"][edit_index] = line
        if keep_ref.get() and ref_ext:
            self.data.setdefault("orc_refs", {})[ref_ext] = {
                "ref_interna": ref_int,
                "ref_externa": ref_ext,
                "descricao": desc,
                "material": mat,
                "espessura": esp,
                "preco_unit": preco,
                "operacao": operacao,
                "desenho": desenho,
            }
            try:
                mysql_upsert_orc_referencia(
                    ref_externa=ref_ext,
                    ref_interna=ref_int,
                    descricao=desc,
                    material=mat,
                    espessura=esp,
                    preco_unit=preco,
                    operacao=operacao,
                    desenho_path=desenho,
                )
            except Exception:
                pass
        direct_saved = False
        try:
            mysql_upsert_orcamento_com_linhas(self.data, orc)
            direct_saved = True
        except Exception:
            direct_saved = False
        self.refresh_orc_linhas(orc)
        if direct_saved:
            try:
                _set_last_save_fingerprint(self.data)
            except Exception:
                pass
        else:
            save_data(self.data)
        win.destroy()

    Btn(win, text="Guardar", command=save_line, width=150 if use_custom else None).grid(row=9, column=0, columnspan=4, pady=12)
    win.grab_set()

def convert_orc_to_encomenda(self):
    _ensure_configured()
    numero = self.selected_orc_numero
    if not numero and hasattr(self, "orc_tbl"):
        sel = self.orc_tbl.selection()
        if sel:
            numero = self.orc_tbl.item(sel[0], "values")[0]
            self.selected_orc_numero = numero
    orc = self.get_orc_by_numero(numero) if numero else None
    if not orc:
        return
    if orc.get("numero_encomenda"):
        messagebox.showerror("Erro", "Orçamento já convertido.")
        return
    estado_norm = (orc.get("estado") or "").strip().lower()
    if "aprovado" not in estado_norm:
        messagebox.showerror("Erro", "Apenas orçamentos aprovados podem ser convertidos.")
        return
    if not orc.get("linhas"):
        messagebox.showerror("Erro", "Sem linhas para converter.")
        return
    nota_cliente = self.prompt_orc_nota_cliente()
    nota_cliente = (nota_cliente or "").strip()
    cli = _normalize_orc_cliente(orc.get("cliente", {}), self.data)
    orc["cliente"] = cli
    codigo = cli.get("codigo", "").strip()
    if codigo and find_cliente(self.data, codigo):
        cliente_code = codigo
    else:
        cliente_code = ""
        for c in self.data.get("clientes", []):
            if cli.get("nif") and c.get("nif") == cli.get("nif"):
                cliente_code = c.get("codigo")
                break
            if cli.get("nome") and c.get("nome") == cli.get("nome"):
                cliente_code = c.get("codigo")
                break
        if not cliente_code:
            cliente_code = next_cliente_codigo(self.data)
            self.data["clientes"].append({
                "codigo": cliente_code,
                "nome": cli.get("nome", ""),
                "nif": cli.get("nif", ""),
                "morada": cli.get("morada", ""),
                "contacto": cli.get("contacto", ""),
                "email": cli.get("email", ""),
                "observacoes": "",
            })
    alert_txt = (
        f"ALERTA: Encomenda gerada por conversao do orcamento {orc.get('numero')}. "
        "Confirmar dados de cliente, materiais, espessuras e prazos."
    )
    obs_txt = f"{alert_txt} | Origem: Orcamento {orc.get('numero')}"
    if nota_cliente:
        obs_txt += f" | Nota cliente: {nota_cliente}"
    enc = {
        "numero": next_encomenda_numero(self.data),
        "cliente": cliente_code,
        "nota_cliente": nota_cliente,
        "data_criacao": now_iso(),
        "data_entrega": "",
        "tempo": 0.0,
        "tempo_estimado": 0.0,
        "cativar": False,
        "observacoes": obs_txt,
        "alerta_conversao": True,
        "estado": "Preparacao",
        "materiais": [],
        "reservas": [],
        "numero_orcamento": orc.get("numero"),
    }
    mats = {}
    peca_idx = 1
    for l in orc.get("linhas", []):
        mat = str(l.get("material", "")).strip()
        esp = str(l.get("espessura", "")).strip()
        if not mat or not esp:
            messagebox.showerror("Erro", "Todas as linhas precisam de Material e Espessura.")
            return
        mats.setdefault(mat, {"material": mat, "estado": "Preparacao", "espessuras": {}})
        mats[mat]["espessuras"].setdefault(esp, {"espessura": esp, "tempo_min": "", "estado": "Preparacao", "pecas": []})
        ref_int = l.get("ref_interna") or next_ref_interna(self.data, cliente_code)
        ops_txt = " + ".join(parse_operacoes_lista(l.get("operacao", "")))
        peca = {
            "id": f"PEC{peca_idx:05d}",
            "ref_interna": ref_int,
            "ref_externa": l.get("ref_externa", ""),
            "material": mat,
            "espessura": esp,
            "quantidade_pedida": float(l.get("qtd", 0) or 0),
            "Operacoes": ops_txt,
            "Observacoes": l.get("descricao", ""),
            "desenho": l.get("desenho", ""),
            "of": next_of_numero(self.data),
            "opp": next_opp_numero(self.data),
            "estado": "Preparacao",
            "produzido_ok": 0.0,
            "produzido_nok": 0.0,
            "inicio_producao": "",
            "fim_producao": "",
        }
        peca["operacoes_fluxo"] = build_operacoes_fluxo(ops_txt)
        peca_idx += 1
        mats[mat]["espessuras"][esp]["pecas"].append(peca)
        update_refs(self.data, peca["ref_interna"], peca["ref_externa"])
    enc["materiais"] = []
    for m in mats.values():
        m["espessuras"] = list(m["espessuras"].values())
        enc["materiais"].append(m)
    self.data["encomendas"].append(enc)
    orc["numero_encomenda"] = enc["numero"]
    if nota_cliente:
        orc["nota_cliente"] = nota_cliente
    orc["estado"] = "Convertido em Encomenda"
    save_data(self.data)
    if hasattr(self, "e_filter"):
        self.e_filter.set("")
    if hasattr(self, "e_filter_entry"):
        self.e_filter_entry.delete(0, END)
    self.selected_encomenda_numero = enc["numero"]
    self.refresh_encomendas()
    try:
        self.refresh()
    except Exception:
        pass
    if hasattr(self, "tab_encomendas"):
        try:
            self.nb.select(self.tab_encomendas)
        except Exception:
            pass
    self.refresh_orc_list()
    messagebox.showinfo("OK", f"Encomenda criada: {enc['numero']}")

def prompt_orc_nota_cliente(self):
    _ensure_configured()
    use_custom = CUSTOM_TK_AVAILABLE and getattr(self, "orc_use_custom", False)
    if not use_custom:
        return simple_input(
            self.root,
            "Nota Cliente",
            "Número da nota de encomenda do cliente (opcional):"
        )
    win = ctk.CTkToplevel(self.root)
    win.title("Nota do Cliente")
    win.geometry("460x180")
    win.resizable(False, False)
    win.transient(self.root)
    win.grab_set()
    out = {"v": None}
    var = StringVar()
    card = ctk.CTkFrame(win, fg_color="#f7f8fb", corner_radius=10)
    card.pack(fill="both", expand=True, padx=10, pady=10)
    ctk.CTkLabel(card, text="Número da nota de encomenda do cliente (opcional)", font=("Segoe UI", 13, "bold"), text_color="#7a0f1a").pack(anchor="w", padx=10, pady=(10, 6))
    ent = ctk.CTkEntry(card, textvariable=var, width=420)
    ent.pack(anchor="w", padx=10, pady=4)
    ent.focus_set()
    row = ctk.CTkFrame(card, fg_color="transparent")
    row.pack(fill="x", padx=10, pady=10)
    ctk.CTkButton(row, text="Confirmar", width=140, fg_color="#a32035", command=lambda: (out.update(v=var.get().strip()), win.destroy())).pack(side="left")
    ctk.CTkButton(row, text="Sem nota", width=120, fg_color="#6b7280", command=lambda: (out.update(v=""), win.destroy())).pack(side="left", padx=8)
    ctk.CTkButton(row, text="Cancelar", width=120, fg_color="#b42318", command=win.destroy).pack(side="right")
    win.wait_window()
    return out.get("v")

def render_orc_pdf(self, path, orc):
    _ensure_configured()
    return _render_orc_pdf_modern(self, path, orc)
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.pdfgen import canvas as pdf_canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.utils import ImageReader
    width, height = landscape(A4)
    c = pdf_canvas.Canvas(path, pagesize=landscape(A4))

    def yinv(y):
        return height - y

    font_regular = "Helvetica"
    font_bold = "Helvetica-Bold"
    try:
        arial = r"C:\Windows\Fonts\arial.ttf"
        arial_b = r"C:\Windows\Fonts\arialbd.ttf"
        if os.path.exists(arial):
            pdfmetrics.registerFont(TTFont("Arial", arial))
            font_regular = "Arial"
        if os.path.exists(arial_b):
            pdfmetrics.registerFont(TTFont("Arial-Bold", arial_b))
            font_bold = "Arial-Bold"
        seg = r"C:\Windows\Fonts\segoeui.ttf"
        seg_b = r"C:\Windows\Fonts\segoeuib.ttf"
        if font_regular == "Helvetica" and os.path.exists(seg):
            pdfmetrics.registerFont(TTFont("SegoeUI", seg))
            font_regular = "SegoeUI"
        if font_bold == "Helvetica-Bold" and os.path.exists(seg_b):
            pdfmetrics.registerFont(TTFont("SegoeUI-Bold", seg_b))
            font_bold = "SegoeUI-Bold"
    except Exception:
        pass

    def set_font(bold, size):
        c.setFont(font_bold if bold else font_regular, size)

    def normalize_text(text):
        return pdf_normalize_text(text)

    def clip_text(text, max_width, font_name=None, font_size=9):
        if text is None:
            return ""
        s = normalize_text(text)
        fname = font_name or font_regular
        if pdfmetrics.stringWidth(s, fname, font_size) <= max_width:
            return s
        ell = "..."
        while s and pdfmetrics.stringWidth(s + ell, fname, font_size) > max_width:
            s = s[:-1]
        return s + ell if s else ""

    def wrap_text(text, max_width, font_name=None, font_size=8):
        text = normalize_text("" if text is None else str(text))
        words = text.split()
        if not words:
            return [""]
        lines = []
        line = words[0]
        fname = font_name or font_regular
        for w in words[1:]:
            test = f"{line} {w}"
            if pdfmetrics.stringWidth(test, fname, font_size) <= max_width:
                line = test
            else:
                lines.append(line)
                line = w
        lines.append(line)
        return lines

    margin = 20
    # Paleta azul para o orçamento
    BLUE_BORDER = (0.16, 0.36, 0.64)
    BLUE_TITLE = (0.10, 0.30, 0.56)
    BLUE_HEADER = (0.15, 0.36, 0.66)
    BLUE_ROW = (0.93, 0.96, 1.00)
    subtotal = float(orc.get("subtotal", 0))
    iva = subtotal * float(orc.get("iva_perc", 0)) / 100.0
    total = subtotal + iva

    avail_w = width - (margin * 2)
    cols_base = [
        ("Linha", 28),
        ("Artigo", 140),
        ("Descricao", 95),
        ("Materia Prima", 80),
        ("Espessura", 48),
        ("Operacoes", 110),
        ("Quant.", 42),
        ("Preco (EUR)", 62),
        ("IVA %", 40),
        ("Total (EUR)", 70),
    ]
    base_sum = sum(w for _, w in cols_base) or 1
    scale = min(1.0, float(avail_w) / float(base_sum))
    cols = []
    for i, (title, w) in enumerate(cols_base):
        if i == len(cols_base) - 1:
            used = sum(x[1] for x in cols)
            cols.append((title, max(34, int(avail_w - used))))
        else:
            cols.append((title, max(24, int(round(w * scale)))))
    table_w = sum(w for _, w in cols)
    table_x = margin

    header_h = 54
    client_h = 60
    gap = 8
    table_header_h = 16
    row_h = 14

    notes_h = 62
    complaints_h = 58
    boxes_h = 70
    footer_h = 44
    bottom_h = notes_h + complaints_h + boxes_h + footer_h + 18
    bottom_h_nonlast = 20

    def draw_header(page_num):
        c.setStrokeColorRGB(*BLUE_BORDER)
        c.rect(margin, yinv(height - margin), width - margin * 2, height - margin * 2, stroke=1, fill=0)
        logo_w = 138
        logo_h = 54
        logo_x = margin + 4
        draw_pdf_logo_box(
            c,
            height,
            logo_x,
            margin + 2,
            box_size=logo_w,
            box_h=logo_h,
            padding=1,
            draw_border=False,
        )
        right_x = width - 210
        title_left = logo_x + logo_w + 10
        title_right = right_x - 10
        box_w = min(340, max(210, title_right - title_left))
        box_x = title_left + max(0.0, ((title_right - title_left) - box_w) / 2.0)
        box_y = margin + 16
        box_h = 22
        c.setStrokeColorRGB(*BLUE_BORDER)
        c.setLineWidth(1.25)
        c.roundRect(box_x, yinv(box_y + box_h), box_w, box_h, 5, stroke=1, fill=0)
        c.setFillColorRGB(*BLUE_TITLE)
        set_font(False, 8.2)
        c.drawCentredString(box_x + (box_w / 2), yinv(box_y - 2), "Orcamento")
        set_font(True, 13)
        c.drawCentredString(
            box_x + (box_w / 2),
            yinv(box_y + 15),
            clip_text(f"N ORC: {orc.get('numero', '')}", box_w - 10, font_name=font_bold, font_size=13),
        )
        data_txt = (orc.get("data", "") or "")[:10]
        right_box_w = 210
        right_box_h = 18
        right_x = width - margin - right_box_w
        y_data = margin + 8
        y_ref = y_data + right_box_h + 4
        c.setLineWidth(1)
        c.setStrokeColorRGB(*BLUE_BORDER)
        c.roundRect(right_x, yinv(y_data + right_box_h), right_box_w, right_box_h, 3, stroke=1, fill=0)
        c.roundRect(right_x, yinv(y_ref + right_box_h), right_box_w, right_box_h, 3, stroke=1, fill=0)
        c.setFillColorRGB(*BLUE_TITLE)
        set_font(False, 9)
        c.drawCentredString(right_x + (right_box_w / 2), yinv(y_data + 12), f"Data: {data_txt}")
        c.drawCentredString(right_x + (right_box_w / 2), yinv(y_ref + 12), f"Ref. Orcamento: {orc.get('numero','')}")
        if page_num > 1:
            c.setFillColorRGB(0, 0, 0)
            c.drawRightString(width - (margin + 6), yinv(margin + 52), f"Pagina {page_num}")
        c.setLineWidth(1)
        c.setFillColorRGB(0, 0, 0)

    def draw_client_box():
        cli = orc.get("cliente", {})
        if not cli.get("nome") and cli.get("codigo"):
            code = cli.get("codigo", "")
            if " - " in code:
                code = code.split(" - ")[0]
            cobj = find_cliente(self.data, code)
            if cobj:
                cli = {**cobj}
        y = margin + header_h + gap
        set_font(False, 9)
        c.rect(margin, yinv(y + client_h), width - (margin * 2), client_h, stroke=1, fill=0)
        c.drawString(margin + 8, yinv(y + 20), clip_text(cli.get("nome", ""), 700, font_size=9))
        c.drawString(margin + 8, yinv(y + 36), f"NIF: {cli.get('nif','')}")
        c.drawString(margin + 8, yinv(y + 52), clip_text(f"Contacto: {cli.get('contacto','')}  Email: {cli.get('email','')}", 700, font_size=9))

    def draw_table_header(y):
        c.setFillColorRGB(*BLUE_HEADER)
        c.rect(table_x, yinv(y + table_header_h), table_w, table_header_h, fill=1, stroke=0)
        c.setFillColorRGB(1, 1, 1)
        set_font(True, 8)
        x = table_x + 4
        for title, w in cols:
            c.drawString(x, yinv(y + 11), title)
            x += w
        c.setFillColorRGB(0, 0, 0)

    def draw_rows(start_y, rows, start_index):
        y = start_y
        set_font(False, 8)
        def fmt_num(val, decimals=2):
            try:
                v = float(val)
            except Exception:
                return str(val) if val is not None else ""
            s = f"{v:.{decimals}f}"
            if "." in s:
                s = s.rstrip("0").rstrip(".")
            return s
        for i, l in enumerate(rows):
            idx = start_index + i
            if idx % 2 == 0:
                c.setFillColorRGB(*BLUE_ROW)
                c.rect(table_x, yinv(y + row_h - 1), table_w, row_h, fill=1, stroke=0)
                c.setFillColorRGB(0, 0, 0)
            x = table_x + 4
            iva_val = l.get("iva", orc.get("iva_perc", 0))
            values = [
                f"{idx+1:03d}",
                clip_text(l.get("ref_externa", ""), cols[1][1] - 6, font_size=8),
                clip_text(l.get("descricao", ""), cols[2][1] - 6, font_size=8),
                clip_text(l.get("material", ""), cols[3][1] - 6, font_size=8),
                clip_text(fmt_num(l.get("espessura", "")), cols[4][1] - 6, font_size=8),
                clip_text(l.get("operacao", ""), cols[5][1] - 6, font_size=8),
                fmt_num(l.get("qtd", 0), decimals=2),
                fmt_num(l.get("preco_unit", 0), decimals=2),
                fmt_num(iva_val, decimals=2),
                fmt_num(l.get("total", 0), decimals=2),
            ]
            for (title, w), val in zip(cols, values):
                if title in ("Preco (EUR)", "IVA %", "Total (EUR)"):
                    c.drawRightString(x + w - 6, yinv(y + 10), val)
                elif title in ("Materia Prima", "Espessura", "Operacoes", "Quant."):
                    c.drawCentredString(x + (w / 2), yinv(y + 10), val)
                else:
                    c.drawString(x, yinv(y + 10), val)
                x += w
            y += row_h
        return y

    def draw_bottom_sections():
        y_top = height - margin - bottom_h
        notes_lines = self._build_orc_notes_lines(orc)

        # Notas
        set_font(True, 9)
        c.rect(margin, yinv(y_top + notes_h), 520, notes_h, stroke=1, fill=0)
        c.drawString(margin + 8, yinv(y_top + 10), "Notas")
        set_font(False, 7)
        ytxt = y_top + 18
        for line in notes_lines:
            if ytxt > y_top + notes_h - 6:
                break
            c.drawString(margin + 8, yinv(ytxt), clip_text(line, 500, font_size=7))
            ytxt += 8

        # Reclamacoes / Devolucoes (caixa fixa, sem sobreposicao)
        y_recl = y_top + notes_h + 4
        c.rect(margin, yinv(y_recl + complaints_h), table_w, complaints_h, stroke=1, fill=0)
        inner_x = margin + 6
        inner_w = table_w - 12
        line_h = 7
        y_line = y_recl + 10
        y_limit = y_recl + complaints_h - 9

        set_font(True, 7.2)
        c.drawString(inner_x, yinv(y_line), "Reclamacoes de cliente:")
        y_line += line_h + 1
        set_font(False, 6.6)
        recl_lines = wrap_text(ORC_RECLAMACOES, inner_w, font_size=6.6)
        for line in recl_lines:
            if y_line > y_limit:
                break
            c.drawString(inner_x, yinv(y_line), clip_text(line, inner_w, font_size=6.6))
            y_line += line_h

        if y_line <= y_limit - (line_h + 2):
            y_line += 1
            set_font(True, 7.2)
            c.drawString(inner_x, yinv(y_line), "Devolucao de materiais:")
            y_line += line_h + 1
            set_font(False, 6.6)
            devol_lines = wrap_text(ORC_DEVOLUCOES, inner_w, font_size=6.6)
            for line in devol_lines:
                if y_line > y_limit:
                    break
                c.drawString(inner_x, yinv(y_line), clip_text(line, inner_w, font_size=6.6))
                y_line += line_h

        # Condicoes / Legenda / Resumo
        y_boxes = y_top + notes_h + complaints_h + 8
        cond_w = 360
        leg_w = 180
        sum_w = 240
        set_font(True, 9)
        c.rect(margin, yinv(y_boxes + boxes_h), cond_w, boxes_h, stroke=1, fill=0)
        c.drawString(margin + 8, yinv(y_boxes + 16), "Condicoes Gerais")
        set_font(False, 8)
        ytxt = y_boxes + 24
        for line in ORC_CONDICOES_GERAIS:
            c.drawString(margin + 8, yinv(ytxt), clip_text(line, cond_w - 16, font_size=8))
            ytxt += 10

        leg_x = margin + cond_w + 8
        set_font(True, 9)
        c.rect(leg_x, yinv(y_boxes + boxes_h), leg_w, boxes_h, stroke=1, fill=0)
        c.drawString(leg_x + 8, yinv(y_boxes + 16), "Legenda")
        set_font(False, 7)
        ytxt = y_boxes + 26
        for line in ORC_LEGENDA_OPERACOES:
            if ytxt > y_boxes + boxes_h - 6:
                break
            c.drawString(leg_x + 8, yinv(ytxt), clip_text(line, leg_w - 16, font_size=7))
            ytxt += 8

        sum_x = width - margin - sum_w
        set_font(True, 9)
        c.rect(sum_x, yinv(y_boxes + boxes_h), sum_w, boxes_h, stroke=1, fill=0)
        c.drawString(sum_x + 8, yinv(y_boxes + 16), "Resumo")
        set_font(False, 9)
        c.drawString(sum_x + 8, yinv(y_boxes + 32), "Subtotal:")
        c.drawRightString(sum_x + sum_w - 8, yinv(y_boxes + 32), f"{subtotal:.2f} EUR")
        c.drawString(sum_x + 8, yinv(y_boxes + 48), f"IVA ({orc.get('iva_perc',0)}%):")
        c.drawRightString(sum_x + sum_w - 8, yinv(y_boxes + 48), f"{iva:.2f} EUR")
        set_font(True, 10)
        c.drawString(sum_x + 8, yinv(y_boxes + 64), "Total:")
        c.drawRightString(sum_x + sum_w - 8, yinv(y_boxes + 64), f"{total:.2f} EUR")

        # Rodap? empresa
        yb = margin + 6
        set_font(False, 7.5)
        for line in get_empresa_rodape_lines():
            c.drawString(margin, yb, line)
            yb += 9
        executado = str(orc.get("executado_por", "") or "").strip() or "Claudio"
        set_font(False, 8)
        c.drawRightString(width - margin - 6, yinv(height - margin - 14), f"Executado por: {executado}")

    def rows_fit_for(first_page, bottom_space):
        y_table_top = margin + header_h + (client_h + gap if first_page else gap)
        y_rows = y_table_top + table_header_h + 2
        max_table_y = height - margin - bottom_space
        return max(0, int((max_table_y - y_rows) // row_h))

    lines = list(orc.get("linhas", []))
    idx = 0
    page_num = 1
    while idx < len(lines) or idx == 0:
        first_page = page_num == 1
        remaining = len(lines) - idx
        rows_fit_last = rows_fit_for(first_page, bottom_h)
        rows_fit_nonlast = rows_fit_for(first_page, bottom_h_nonlast)
        if remaining <= rows_fit_last:
            is_last = True
            rows_fit = remaining
        else:
            is_last = False
            if remaining <= rows_fit_nonlast:
                # reservar espaco para a pagina final com resumo
                rows_fit = max(1, remaining - rows_fit_last)
            else:
                rows_fit = rows_fit_nonlast

        draw_header(page_num)
        if first_page:
            draw_client_box()
        y_table_top = margin + header_h + (client_h + gap if first_page else gap)
        draw_table_header(y_table_top)
        y_rows = y_table_top + table_header_h + 2
        page_rows = lines[idx: idx + rows_fit]
        _ = draw_rows(y_rows, page_rows, idx)
        idx += len(page_rows)

        if is_last:
            draw_bottom_sections()
        else:
            set_font(False, 8)
            c.drawRightString(width - margin, yinv(height - margin - 8), "Continua na proxima pagina...")

        if idx >= len(lines) and len(lines) > 0:
            if is_last:
                break
            # forcar pagina final com resumo
            c.showPage()
            page_num += 1
            continue
        c.showPage()
        page_num += 1
    c.save()


def _render_orc_pdf_modern(self, path, orc):
    _ensure_configured()
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.pdfgen import canvas as pdf_canvas

    fonts = _orc_pdf_register_fonts()
    palette = _orc_pdf_brand_palette()
    width, height = landscape(A4)
    c = pdf_canvas.Canvas(path, pagesize=landscape(A4))
    margin = 22
    content_w = width - (2 * margin)
    header_row_h = 22
    row_h = 18
    first_table_top = 226
    next_table_top = 132
    footer_start = 458
    doc_num = str(orc.get("numero", "") or "ORC").strip() or "ORC"
    c.setTitle(f"Orcamento {doc_num}")
    try:
        c.setPageCompression(0)
    except Exception:
        pass

    subtotal = parse_float(orc.get("subtotal", 0), 0)
    iva_perc = parse_float(orc.get("iva_perc", 0), 0)
    iva = subtotal * (iva_perc / 100.0)
    total = parse_float(orc.get("total", subtotal + iva), subtotal + iva)
    estado = str(orc.get("estado", "") or "Em edicao").strip() or "Em edicao"
    executado = str(orc.get("executado_por", "") or "").strip() or "Claudio"
    nota_transporte = str(orc.get("nota_transporte", "") or "").strip()
    nota_cliente = str(orc.get("nota_cliente", "") or "").strip()
    numero_encomenda = str(orc.get("numero_encomenda", "") or "").strip()
    lines = list(orc.get("linhas", []) or [])
    notes_lines = list(self._build_orc_notes_lines(orc) or [])
    footer_company = list(get_empresa_rodape_lines() or [])

    cli = orc.get("cliente", {}) if isinstance(orc.get("cliente", {}), dict) else {}
    if not cli.get("nome") and cli.get("codigo"):
        code = str(cli.get("codigo", "") or "").strip()
        if " - " in code:
            code = code.split(" - ", 1)[0].strip()
        cobj = find_cliente(self.data, code)
        if cobj:
            cli = {**cobj}

    cols = [
        ("Linha", 34, "center"),
        ("Ref. Ext.", 126, "w"),
        ("Descricao", 188, "w"),
        ("Material", 88, "w"),
        ("Esp.", 34, "center"),
        ("Operacoes", 118, "w"),
        ("Qtd.", 38, "e"),
        ("P.Unit.", 62, "e"),
        ("IVA%", 30, "e"),
        ("Total", 68, "e"),
    ]
    table_w = sum(width for _, width, _ in cols)
    first_capacity = max(1, int((footer_start - (first_table_top + header_row_h + 8)) // row_h))
    next_capacity = max(1, int((footer_start - (next_table_top + header_row_h + 8)) // row_h))

    def yinv(top_y):
        return height - top_y

    def ntxt(value):
        try:
            return pdf_normalize_text(value)
        except Exception:
            return str(value or "")

    def fmt_display_date(value):
        raw = str(value or "").strip().replace("T", " ")
        if not raw:
            return "-"
        try:
            dt = datetime.fromisoformat(raw)
            return dt.strftime("%d/%m/%Y")
        except Exception:
            if len(raw) >= 10 and raw[4] == "-" and raw[7] == "-":
                return f"{raw[8:10]}/{raw[5:7]}/{raw[:4]}"
            return raw[:10]

    def fmt_money(value):
        return f"{parse_float(value, 0):.2f} EUR"

    def fmt_num(value, decimals=2):
        try:
            number = float(value)
        except Exception:
            return str(value) if value is not None else ""
        text = f"{number:.{decimals}f}"
        if "." in text:
            text = text.rstrip("0").rstrip(".")
        return text

    def draw_shell() -> None:
        c.saveState()
        c.setFillColor(palette["primary"])
        c.roundRect(margin, yinv(64), content_w, 5, 2, stroke=0, fill=1)
        c.setFillColor(palette["surface_alt"])
        c.roundRect(margin, yinv(height - (margin * 2)), content_w, height - (margin * 2), 18, stroke=0, fill=1)
        c.setFillColor(colors.white)
        c.roundRect(margin + 1, yinv(height - (margin * 2) - 1), content_w - 2, height - (margin * 2) - 2, 18, stroke=0, fill=1)
        c.setStrokeColor(palette["line"])
        c.setLineWidth(1)
        c.roundRect(margin, yinv(height - (margin * 2)), content_w, height - (margin * 2), 18, stroke=1, fill=0)
        c.restoreState()

    def draw_badge(x, top_y, text, *, fill_color=None, text_color=None, height_box=20, align="left"):
        label = str(text or "").strip() or "-"
        box_w = max(64, min(170, 22 + (len(label) * 6.2)))
        fill = fill_color or palette["primary_soft"]
        txt_color = text_color or palette["primary_dark"]
        c.saveState()
        c.setFillColor(fill)
        c.roundRect(x if align == "left" else x - box_w, yinv(top_y + height_box), box_w, height_box, 10, stroke=0, fill=1)
        c.restoreState()
        c.setFillColor(txt_color)
        c.setFont(fonts["bold"], 8.2)
        draw_x = (x + 10) if align == "left" else (x - box_w + 10)
        c.drawString(draw_x, yinv(top_y + 13.2), ntxt(_orc_pdf_clip_text(label, box_w - 18, fonts["bold"], 8.2)))
        return box_w

    def metric_card(x, top_y, box_w, label, value, *, box_h=34, accent=False, compact=False):
        c.saveState()
        c.setFillColor(colors.white)
        c.setStrokeColor(palette["line_strong"] if accent else palette["line"])
        c.roundRect(x, yinv(top_y + box_h), box_w, box_h, 10, stroke=1, fill=1)
        if accent:
            c.setFillColor(palette["primary_soft"])
            c.roundRect(x + 1, yinv(top_y + box_h) + box_h - 5, box_w - 2, 4, 8, stroke=0, fill=1)
        c.restoreState()
        c.setFillColor(palette["muted"])
        c.setFont(fonts["regular"], 7.0 if compact else 7.3)
        c.drawString(x + 8, yinv(top_y + 10), ntxt(label))
        c.setFillColor(palette["primary_dark"])
        c.setFont(fonts["bold"], 10.2 if compact else 10.6)
        c.drawString(x + 8, yinv(top_y + 23), ntxt(_orc_pdf_clip_text(value, box_w - 16, fonts["bold"], 10.2 if compact else 10.6)))

    def detail_card(x, top_y, box_w, box_h, title, lines_, *, accent=False, font_size=8.4):
        title_fill = palette["primary_soft"] if accent else palette["surface_alt"]
        c.saveState()
        c.setFillColor(colors.white)
        c.setStrokeColor(palette["line_strong"] if accent else palette["line"])
        c.setLineWidth(1.0 if accent else 0.9)
        c.roundRect(x, yinv(top_y + box_h), box_w, box_h, 12, stroke=1, fill=1)
        c.setFillColor(title_fill)
        c.roundRect(x + 1, yinv(top_y + 24), box_w - 2, 23, 11, stroke=0, fill=1)
        c.restoreState()
        c.setFont(fonts["bold"], 8.8)
        c.setFillColor(palette["primary_dark"])
        c.drawString(x + 10, yinv(top_y + 15), ntxt(title))
        yy = top_y + 37
        for idx_line, line in enumerate(lines_):
            line_font = fonts["bold"] if accent and idx_line == 0 else fonts["regular"]
            line_size = font_size + 0.4 if accent and idx_line == 0 else font_size
            wrapped = _orc_pdf_wrap_text(line, line_font, line_size, box_w - 20, max_lines=2)
            for item in wrapped:
                c.setFont(line_font, line_size)
                c.setFillColor(palette["ink"])
                c.drawString(x + 10, yinv(yy), ntxt(item))
                yy += 10.5
                if yy > top_y + box_h - 8:
                    return

    def summary_panel(x, top_y, box_w, box_h):
        c.saveState()
        c.setFillColor(colors.white)
        c.setStrokeColor(palette["line_strong"])
        c.roundRect(x, yinv(top_y + box_h), box_w, box_h, 14, stroke=1, fill=1)
        c.setFillColor(palette["primary_soft"])
        c.roundRect(x + 1, yinv(top_y + box_h) + box_h - 7, box_w - 2, 6, 11, stroke=0, fill=1)
        c.restoreState()
        c.setFillColor(palette["primary_dark"])
        c.setFont(fonts["bold"], 9.6)
        c.drawString(x + 12, yinv(top_y + 15), ntxt(_orc_pdf_clip_text("Resumo financeiro", box_w - 24, fonts["bold"], 9.6)))
        c.setFont(fonts["regular"], 7.4)
        c.drawString(x + 12, yinv(top_y + 28), ntxt(_orc_pdf_clip_text(f"Estado: {estado}", box_w - 24, fonts["regular"], 7.4)))
        c.drawString(x + 12, yinv(top_y + 39), ntxt(f"Linhas: {len(lines)}"))
        c.setStrokeColor(colors.Color(1, 1, 1, alpha=0.22))
        c.line(x + 12, yinv(top_y + 45), x + box_w - 12, yinv(top_y + 45))
        rows = [
            ("Subtotal", fmt_money(subtotal)),
            (f"IVA ({fmt_num(iva_perc)}%)", fmt_money(iva)),
        ]
        yy = top_y + 58
        for label, value in rows:
            c.setFont(fonts["regular"], 8.5)
            c.drawString(x + 12, yinv(yy), ntxt(label))
            c.drawRightString(x + box_w - 12, yinv(yy), ntxt(value))
            yy += 12
        c.setStrokeColor(colors.Color(1, 1, 1, alpha=0.22))
        c.line(x + 12, yinv(top_y + 75), x + box_w - 12, yinv(top_y + 75))
        c.setFont(fonts["bold"], 9.2)
        c.drawString(x + 12, yinv(top_y + 87), ntxt("Total"))
        c.setFont(fonts["bold"], 12.4)
        c.drawRightString(x + box_w - 12, yinv(top_y + 87), ntxt(fmt_money(total)))

    def draw_table_header(top_y):
        c.saveState()
        c.setFillColor(palette["primary_soft"])
        c.setStrokeColor(palette["line"])
        c.setLineWidth(0.8)
        c.roundRect(margin, yinv(top_y + header_row_h), table_w, header_row_h, 8, stroke=1, fill=1)
        c.restoreState()
        c.setFillColor(palette["primary_dark"])
        c.setFont(fonts["bold"], 8.15)
        xx = margin
        for name, width_col, align in cols:
            if align == "e":
                c.drawRightString(xx + width_col - 8, yinv(top_y + 14.2), ntxt(name))
            elif align == "center":
                c.drawCentredString(xx + (width_col / 2.0), yinv(top_y + 14.2), ntxt(name))
            else:
                c.drawString(xx + 8, yinv(top_y + 14.2), ntxt(name))
            xx += width_col
        return top_y + header_row_h + 6

    def draw_row(top_y, idx_line, line, absolute_idx):
        fill = palette["surface_warm"] if idx_line % 2 == 0 else colors.white
        c.saveState()
        c.setFillColor(fill)
        c.setStrokeColor(palette["line"])
        c.setLineWidth(0.45)
        c.roundRect(margin, yinv(top_y + row_h), table_w, row_h, 6, stroke=1, fill=1)
        c.restoreState()
        values = [
            f"{absolute_idx + 1:03d}",
            _orc_line_ref_display(line),
            _orc_line_description_display(line),
            _orc_line_material_display(line),
            fmt_num(_orc_line_unit_display(line), decimals=2),
            _orc_line_operacao_display(line),
            fmt_num(line.get("qtd", 0), decimals=2),
            fmt_num(line.get("preco_unit", 0)),
            fmt_num(line.get("iva", iva_perc), decimals=2),
            fmt_num(line.get("total", 0)),
        ]
        xx = margin
        c.setFont(fonts["regular"], 8.0)
        c.setFillColor(palette["ink"])
        c.setStrokeColor(palette["line"])
        for idx_col, ((_, width_col, align), value) in enumerate(zip(cols, values)):
            if idx_col > 0:
                c.line(xx, yinv(top_y), xx, yinv(top_y + row_h))
            txt = _orc_pdf_clip_text(value, width_col - 16, fonts["regular"], 8.0)
            if idx_col == 1:
                c.setFont(fonts["bold"], 8.0)
            else:
                c.setFont(fonts["regular"], 8.0)
            if align == "e":
                c.drawRightString(xx + width_col - 8, yinv(top_y + 12.1), ntxt(txt))
            elif align == "center":
                c.drawCentredString(xx + (width_col / 2.0), yinv(top_y + 12.1), ntxt(txt))
            else:
                c.drawString(xx + 8, yinv(top_y + 12.1), ntxt(txt))
            xx += width_col

    def draw_header(page_no, total_pages, first_page):
        if first_page:
            draw_shell()
            logo_plate_w = 136
            logo_plate_gap = 12
            hero_x = margin + 10 + logo_plate_w + logo_plate_gap
            hero_w = (content_w - 20) - logo_plate_w - logo_plate_gap
            draw_pdf_logo_plate(c, height, margin + 10, 32, box_w=logo_plate_w, box_h=58, padding=4)
            header_panel = globals().get("draw_pdf_header_panel")
            if callable(header_panel):
                header_panel(c, height, hero_x, 26, hero_w, 88, radius=16, stroke_color="#D5DDE7", accent_color="#EAF0F6", accent_height=5)
            else:
                c.saveState()
                c.setFillColor(colors.white)
                c.setStrokeColor(palette["line_strong"])
                c.setLineWidth(1.0)
                c.roundRect(hero_x, yinv(114), hero_w, 88, 16, stroke=1, fill=1)
                c.restoreState()
            c.setFillColor(palette["muted"])
            c.setFont(fonts["regular"], 8.8)
            c.drawString(hero_x + 18, yinv(48), ntxt("Proposta comercial"))
            hero_center_x = hero_x + (hero_w / 2.0)
            c.setFont(fonts["bold"], 15.5)
            c.setFillColor(palette["primary_dark"])
            c.drawCentredString(hero_center_x, yinv(58), ntxt("Orcamento"))
            c.saveState()
            c.setFillColor(palette["surface_alt"])
            c.setStrokeColor(palette["line_strong"])
            c.roundRect(hero_center_x - 92, yinv(88), 184, 24, 10, stroke=1, fill=1)
            c.restoreState()
            c.setFont(fonts["bold"], 11.8)
            c.setFillColor(palette["primary_dark"])
            c.drawCentredString(hero_center_x, yinv(82), ntxt(_orc_pdf_clip_text(doc_num, 164, fonts["bold"], 11.8)))
            c.setFont(fonts["regular"], 9.1)
            c.setFillColor(palette["muted"])
            c.drawCentredString(hero_center_x, yinv(104), ntxt("Documento preparado para envio ao cliente"))
            draw_badge(margin + content_w - 20, 34, f"Pagina {page_no}/{total_pages}", fill_color=palette["primary_soft"], text_color=palette["primary_dark"], align="right")
            draw_badge(margin + content_w - 20, 58, estado, fill_color=palette["surface_alt"], text_color=palette["primary_dark"], align="right")
            metric_grid = _orc_pdf_metric_grid_layout(width - margin - 8, 126, 64, cols=3, rows=2, group_w=360, gap=8)
            metric_card(metric_grid["cols"][0], metric_grid["rows"][0], metric_grid["chip_w"], "Documento", doc_num, box_h=metric_grid["chip_h"], compact=True)
            metric_card(metric_grid["cols"][1], metric_grid["rows"][0], metric_grid["chip_w"], "Data", fmt_display_date(orc.get("data", "")), box_h=metric_grid["chip_h"], compact=True)
            metric_card(metric_grid["cols"][2], metric_grid["rows"][0], metric_grid["chip_w"], "Executado por", executado, box_h=metric_grid["chip_h"], compact=True)
            metric_card(metric_grid["cols"][0], metric_grid["rows"][1], metric_grid["chip_w"], "Subtotal", fmt_money(subtotal), box_h=metric_grid["chip_h"], compact=True)
            metric_card(metric_grid["cols"][1], metric_grid["rows"][1], metric_grid["chip_w"], "IVA", f"{fmt_num(iva_perc)}%", box_h=metric_grid["chip_h"], compact=True)
            metric_card(metric_grid["cols"][2], metric_grid["rows"][1], metric_grid["chip_w"], "Total", fmt_money(total), box_h=metric_grid["chip_h"], accent=True, compact=True)
            detail_card(
                margin,
                126,
                332,
                86,
                "Cliente",
                [
                    str(cli.get("nome", "") or cli.get("empresa", "") or cli.get("codigo", "") or "-").strip(),
                    f"NIF: {str(cli.get('nif', '') or '-').strip()}",
                    f"Contacto: {str(cli.get('contacto', '') or '-').strip()} | Email: {str(cli.get('email', '') or '-').strip()}",
                    str(cli.get("morada", "") or "-").strip(),
                ],
                accent=True,
            )
            detail_card(
                margin + 344,
                126,
                content_w - 344,
                86,
                "Contexto comercial",
                [
                    f"Nota cliente: {nota_cliente or '-'}",
                    f"Transporte: {nota_transporte or '-'}",
                    f"Encomenda associada: {numero_encomenda or '-'}",
                    f"Operacoes no orcamento: {', '.join(list(self._extract_orc_operacoes(orc) or [])[:4]) or '-'}",
                ],
                font_size=8.2,
            )
            return draw_table_header(first_table_top)

        draw_shell()
        c.saveState()
        c.setFillColor(palette["surface_alt"])
        c.roundRect(margin + 8, yinv(84), content_w - 16, 48, 14, stroke=0, fill=1)
        c.restoreState()
        c.setFillColor(palette["primary_dark"])
        c.setFont(fonts["bold"], 16)
        c.drawString(margin + 18, yinv(48), ntxt(f"Orcamento {doc_num}"))
        c.setFont(fonts["regular"], 8.8)
        c.setFillColor(palette["muted"])
        c.drawString(margin + 18, yinv(63), ntxt(f"Cliente: {str(cli.get('nome', '') or cli.get('codigo', '') or '-').strip()}"))
        metric_grid = _orc_pdf_metric_grid_layout(width - margin - 8, 94, 30, cols=3, rows=1, group_w=324, gap=8)
        metric_card(metric_grid["cols"][0], metric_grid["rows"][0], metric_grid["chip_w"], "Estado", estado, box_h=metric_grid["chip_h"], compact=True)
        metric_card(metric_grid["cols"][1], metric_grid["rows"][0], metric_grid["chip_w"], "Data", fmt_display_date(orc.get("data", "")), box_h=metric_grid["chip_h"], compact=True)
        metric_card(metric_grid["cols"][2], metric_grid["rows"][0], metric_grid["chip_w"], "Pagina", f"{page_no}/{total_pages}", box_h=metric_grid["chip_h"], compact=True)
        return draw_table_header(next_table_top)

    def draw_footer(page_no, total_pages):
        footer_y = footer_start + 4
        footer_gap = 10
        notes_w = 326
        cond_w = 202
        legend_w = 104
        summary_w = max(120, int(table_w - notes_w - cond_w - legend_w - (footer_gap * 3)))
        notes_x = margin
        cond_x = notes_x + notes_w + footer_gap
        legend_x = cond_x + cond_w + footer_gap
        summary_x = legend_x + legend_w + footer_gap
        detail_card(
            notes_x,
            footer_y,
            notes_w,
            92,
            "Notas comerciais",
            notes_lines[:6] or ["Sem notas adicionais para este orcamento."],
            font_size=7.7,
        )
        detail_card(
            cond_x,
            footer_y,
            cond_w,
            92,
            "Condicoes",
            list(ORC_CONDICOES_GERAIS[:3]) or ["Sem condicoes registadas."],
            font_size=7.2,
        )
        detail_card(
            legend_x,
            footer_y,
            legend_w,
            92,
            "Legenda",
            list(ORC_LEGENDA_OPERACOES[:4]) or ["-"],
            font_size=6.9,
        )
        summary_panel(summary_x, footer_y, summary_w, 92)
        c.saveState()
        c.setFillColor(palette["surface_alt"])
        c.roundRect(margin, yinv(574), table_w, 18, 9, stroke=0, fill=1)
        c.restoreState()
        c.setFont(fonts["regular"], 7.0)
        c.setFillColor(palette["ink"])
        left_text = footer_company[0] if footer_company else ""
        right_text = footer_company[2] if len(footer_company) > 2 else (footer_company[1] if len(footer_company) > 1 else "")
        c.drawString(margin + 10, yinv(562), ntxt(_orc_pdf_clip_text(left_text, 320, fonts["regular"], 7.0)))
        c.drawRightString(width - margin - 10, yinv(562), ntxt(_orc_pdf_clip_text(right_text, 260, fonts["regular"], 7.0)))
        c.drawRightString(width - margin - 10, yinv(574), ntxt(f"Executado por {executado} | Pagina {page_no}/{total_pages}"))

    def paginate_items(items):
        if not items:
            return [[]]
        pages = [items[:first_capacity]]
        rem = items[first_capacity:]
        while rem:
            pages.append(rem[:next_capacity])
            rem = rem[next_capacity:]
        return pages

    pages = paginate_items(lines)
    total_pages = len(pages)
    for page_no, page_rows in enumerate(pages, start=1):
        table_y = draw_header(page_no, total_pages, page_no == 1)
        if not page_rows:
            c.setFont(fonts["regular"], 9)
            c.setFillColor(palette["muted"])
            c.drawString(margin + 8, yinv(table_y + 16), ntxt("Sem linhas no orcamento."))
        else:
            y_row = table_y
            absolute_start = sum(len(pages[idx]) for idx in range(page_no - 1))
            for idx_line, line in enumerate(page_rows):
                draw_row(y_row, idx_line, line, absolute_start + idx_line)
                y_row += row_h
        if page_no == total_pages:
            draw_footer(page_no, total_pages)
        if page_no < total_pages:
            c.showPage()

    c.save()


def preview_orcamento(self):
    _ensure_configured()
    if getattr(self, "selected_orc_numero", None):
        self.save_orc_fields(refresh_list=False)
    orc = self.get_orc_by_numero(self.selected_orc_numero) if getattr(self, "selected_orc_numero", None) else None
    if not orc:
        return
    import tempfile
    path_pdf = os.path.join(tempfile.gettempdir(), f"lugest_orcamento_{orc.get('numero','')}.pdf")
    try:
        self.render_orc_pdf(path_pdf, orc)
    except Exception as exc:
        messagebox.showerror("Erro", f"Falha ao gerar PDF: {exc}")
        return
    try:
        os.startfile(path_pdf)
    except Exception:
        try:
            os.startfile(path_pdf, "open")
        except Exception:
            messagebox.showerror("Erro", "Não foi possível abrir o PDF para pré-visualização.")

def save_orc_pdf(self):
    _ensure_configured()
    orc = self.get_orc_by_numero(self.selected_orc_numero) if getattr(self, "selected_orc_numero", None) else None
    if not orc:
        return
    path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
    if not path:
        return
    self.render_orc_pdf(path, orc)
    messagebox.showinfo("OK", "PDF guardado com sucesso.")

def print_orc_pdf(self):
    _ensure_configured()
    orc = self.get_orc_by_numero(self.selected_orc_numero) if getattr(self, "selected_orc_numero", None) else None
    if not orc:
        return
    import tempfile
    path = os.path.join(tempfile.gettempdir(), "lugest_orcamento.pdf")
    self.render_orc_pdf(path, orc)
    try:
        os.startfile(path, "print")
    except Exception:
        try:
            os.startfile(path)
        except Exception:
            messagebox.showerror("Erro", "Não foi possível abrir a impressão.")

def open_orc_pdf_with(self):
    _ensure_configured()
    orc = self.get_orc_by_numero(self.selected_orc_numero) if getattr(self, "selected_orc_numero", None) else None
    if not orc:
        return
    import tempfile
    path = os.path.join(tempfile.gettempdir(), "lugest_orcamento.pdf")
    self.render_orc_pdf(path, orc)
    try:
        subprocess.Popen(["rundll32", "shell32.dll,OpenAs_RunDLL", path])
    except Exception:
        messagebox.showerror("Erro", "Não foi possível abrir o seletor de aplicações.")

def gerir_orcamentistas(self):
    _ensure_configured()
    use_custom = CUSTOM_TK_AVAILABLE and getattr(self, "orc_use_custom", False)
    Win = ctk.CTkToplevel if use_custom else Toplevel
    Frm = ctk.CTkFrame if use_custom else ttk.Frame
    Btn = ctk.CTkButton if use_custom else ttk.Button
    win = Win(self.root)
    win.title("Orçamentistas")
    try:
        if use_custom:
            win.geometry("460x340")
            win.configure(fg_color="#f7f8fb")
        win.transient(self.root)
        win.grab_set()
    except Exception:
        pass
    frame = Frm(win, fg_color="#f7f8fb") if use_custom else Frm(win)
    frame.pack(fill="both", expand=True, padx=10, pady=10)
    list_wrap = Frm(frame, fg_color="#ffffff") if use_custom else Frm(frame)
    list_wrap.pack(side="left", fill="both", expand=True)
    lst = Listbox(list_wrap, height=10, width=34, font=("Segoe UI", 11), bg="#f8fbff", relief="flat")
    lst.pack(side="left", fill="both", expand=True)
    sb = ttk.Scrollbar(list_wrap, orient="vertical", command=lst.yview)
    lst.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    for op in self.data.get("orcamentistas", []):
        lst.insert(END, op)
    btns = Frm(frame, fg_color="#f7f8fb") if use_custom else Frm(frame)
    btns.pack(side="left", padx=8, fill="y")

    def add_op():
        nome = simple_input(self.root, "Orçamentista", "Nome do orçamentista:")
        if not nome:
            return
        if nome in self.data.get("orcamentistas", []):
            messagebox.showerror("Erro", "Orçamentista já existe")
            return
        self.data.setdefault("orcamentistas", []).append(nome)
        lst.insert(END, nome)
        if hasattr(self, "orc_exec_cb"):
            self.orc_exec_cb.configure(values=self.data.get("orcamentistas", []))
        save_data(self.data)

    def remove_op():
        sel = lst.curselection()
        if not sel:
            return
        nome = lst.get(sel[0])
        self.data["orcamentistas"] = [o for o in self.data.get("orcamentistas", []) if o != nome]
        lst.delete(sel[0])
        if hasattr(self, "orc_exec_cb"):
            self.orc_exec_cb.configure(values=self.data.get("orcamentistas", []))
        if hasattr(self, "orc_executado") and self.orc_executado.get() == nome:
            self.orc_executado.set("")
        save_data(self.data)

    Btn(btns, text="Adicionar", command=add_op, width=130 if use_custom else None).pack(fill="x", pady=(0, 6))
    Btn(btns, text="Remover", command=remove_op, width=130 if use_custom else None).pack(fill="x", pady=(0, 6))
    Btn(btns, text="Fechar", command=win.destroy, width=130 if use_custom else None).pack(fill="x")
