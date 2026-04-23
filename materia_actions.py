from module_context import configure_module, ensure_module

_CONFIGURED = False


def configure(main_globals):
    configure_module(globals(), main_globals)


def _ensure_configured():
    ensure_module(globals(), "materia_actions")


def _mat_pdf_register_fonts():
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


def _mat_pdf_hex_to_rgb(value):
    txt = str(value or "").strip().lstrip("#")
    if len(txt) == 3:
        txt = "".join(ch * 2 for ch in txt)
    if len(txt) != 6:
        txt = "1F3C88"
    try:
        return tuple(int(txt[i : i + 2], 16) for i in (0, 2, 4))
    except Exception:
        return (31, 60, 136)


def _mat_pdf_rgb_to_hex(rgb):
    r, g, b = [max(0, min(255, int(v))) for v in tuple(rgb or (31, 60, 136))]
    return f"#{r:02X}{g:02X}{b:02X}"


def _mat_pdf_mix_hex(base_hex, target_hex, ratio):
    ratio = max(0.0, min(1.0, float(ratio)))
    base = _mat_pdf_hex_to_rgb(base_hex)
    target = _mat_pdf_hex_to_rgb(target_hex)
    return _mat_pdf_rgb_to_hex(
        tuple(round((base_v * (1.0 - ratio)) + (target_v * ratio)) for base_v, target_v in zip(base, target))
    )


def _mat_pdf_palette():
    from reportlab.lib import colors

    primary_hex = "#1F3C88"
    try:
        cfg = get_branding_config() if callable(globals().get("get_branding_config")) else {}
        primary_hex = str((cfg or {}).get("primary_color", "") or primary_hex).strip() or primary_hex
    except Exception:
        pass
    return {
        "primary": colors.HexColor(primary_hex),
        "primary_dark": colors.HexColor(_mat_pdf_mix_hex(primary_hex, "#000000", 0.22)),
        "primary_soft": colors.HexColor(_mat_pdf_mix_hex(primary_hex, "#FFFFFF", 0.82)),
        "primary_soft_2": colors.HexColor(_mat_pdf_mix_hex(primary_hex, "#FFFFFF", 0.90)),
        "surface_warm": colors.HexColor("#FCFCFD"),
        "line": colors.HexColor(_mat_pdf_mix_hex(primary_hex, "#D7DEE8", 0.76)),
        "line_strong": colors.HexColor(_mat_pdf_mix_hex(primary_hex, "#708090", 0.34)),
        "muted": colors.HexColor("#667085"),
        "ink": colors.HexColor(_mat_pdf_mix_hex(primary_hex, "#1A1A1A", 0.72)),
        "danger_fill": colors.HexColor("#FEECEC"),
        "danger_text": colors.HexColor("#B42318"),
        "retalho_fill": colors.HexColor("#FFF4E5"),
        "retalho_text": colors.HexColor("#B54708"),
    }


def _mat_pdf_clip_text(value, max_w, font_name, font_size):
    from reportlab.pdfbase import pdfmetrics

    txt = "" if value is None else str(value)
    if pdfmetrics.stringWidth(txt, font_name, font_size) <= max_w:
        return txt
    ellipsis = "..."
    while txt and pdfmetrics.stringWidth(txt + ellipsis, font_name, font_size) > max_w:
        txt = txt[:-1]
    return f"{txt}{ellipsis}" if txt else ""


def _mat_pdf_wrap_text(value, font_name, font_size, max_w, max_lines=None):
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


def _mat_pdf_fit_font_size(text, font_name, max_w, preferred_size, min_size):
    from reportlab.pdfbase import pdfmetrics

    size = float(preferred_size)
    raw = str(text or "")
    max_w = max(12.0, float(max_w))
    while size > float(min_size) and pdfmetrics.stringWidth(raw, font_name, size) > max_w:
        size -= 0.3
    return max(float(min_size), round(size, 2))


def _mat_pdf_metric_grid_layout(group_right, banner_top, banner_height, cols=3, rows=2, group_w=338, gap=6):
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


def _materia_preco_unid_record(m):
    try:
        if "materia_preco_unitario" in globals() and callable(materia_preco_unitario):
            return float(parse_float(materia_preco_unitario(m), 0))
    except Exception:
        pass
    formato = str((m or {}).get("formato", "") or "").strip()
    compra = float(parse_float((m or {}).get("p_compra", 0), 0))
    if formato == "Tubo":
        return float(parse_float((m or {}).get("metros", 0), 0)) * compra
    if formato in ("Chapa", "Perfil", "Cantoneira", "Barra"):
        return float(parse_float((m or {}).get("peso_unid", 0), 0)) * compra
    return compra


def _norm_material_key(value):
    txt = str(value or "").strip()
    try:
        if "norm_text" in globals() and callable(norm_text):
            return str(norm_text(txt) or "").strip()
    except Exception:
        pass
    return txt.lower()


def _num(value, default=0.0):
    try:
        return float(parse_float(value, default))
    except Exception:
        return float(default)


def _is_retalho_like(record):
    if not isinstance(record, dict):
        return False
    if bool(record.get("is_sobra")):
        return True
    local = _norm_material_key(
        record.get("Localização")
        or record.get("Localizacao")
        or record.get("LocalizaÃ§Ã£o")
    )
    return "retalho" in local


_MATERIAL_FAMILY_PROFILES = {
    "": {"label": "Auto", "density": 0.0},
    "steel": {"label": "Aço / Ferro", "density": 7.85},
    "stainless": {"label": "Inox", "density": 7.93},
    "aluminum": {"label": "Alumínio", "density": 2.70},
    "brass": {"label": "Latão", "density": 8.50},
    "copper": {"label": "Cobre", "density": 8.96},
}


def _normalise_material_family_key(value):
    txt = _norm_material_key(value)
    if txt in {"", "auto", "automatico", "automático"}:
        return ""
    if any(tok in txt for tok in ("inox", "stainless", "aisi", "304", "316", "430")):
        return "stainless"
    if any(tok in txt for tok in ("alumin", "aluminum", "aluminium", "aw-")):
        return "aluminum"
    if any(tok in txt for tok in ("latao", "latao", "brass", "cuzn")):
        return "brass"
    if any(tok in txt for tok in ("cobre", "copper", "cu-")):
        return "copper"
    if any(tok in txt for tok in ("aco", "aço", "ferro", "steel", "carbon")):
        return "steel"
    return ""


def _material_family_options(include_auto=True):
    keys = list(_MATERIAL_FAMILY_PROFILES.keys())
    if not include_auto:
        keys = [key for key in keys if key]
    return [
        {
            "key": key,
            "label": str(_MATERIAL_FAMILY_PROFILES.get(key, {}).get("label", "") or ""),
            "density": float(_MATERIAL_FAMILY_PROFILES.get(key, {}).get("density", 0.0) or 0.0),
        }
        for key in keys
    ]


def _resolve_material_family(material="", family=""):
    explicit_key = _normalise_material_family_key(family)
    if explicit_key:
        profile = dict(_MATERIAL_FAMILY_PROFILES.get(explicit_key, _MATERIAL_FAMILY_PROFILES["steel"]))
        return {
            "key": explicit_key,
            "label": str(profile.get("label", "") or ""),
            "density": float(profile.get("density", 7.85) or 7.85),
            "explicit": True,
        }

    txt = _norm_material_key(material)
    guessed_key = "steel"
    if any(tok in txt for tok in ("alumin", "aw-")):
        guessed_key = "aluminum"
    elif "cobre" in txt:
        guessed_key = "copper"
    elif "latao" in txt or "cuzn" in txt:
        guessed_key = "brass"
    elif "aisi" in txt or "inox" in txt or "304" in txt or "316" in txt or "430" in txt:
        guessed_key = "stainless"
    profile = dict(_MATERIAL_FAMILY_PROFILES.get(guessed_key, _MATERIAL_FAMILY_PROFILES["steel"]))
    return {
        "key": guessed_key,
        "label": str(profile.get("label", "") or ""),
        "density": float(profile.get("density", 7.85) or 7.85),
        "explicit": False,
    }


def _guess_density(material, family=""):
    return float(_resolve_material_family(material, family).get("density", 7.85) or 7.85)


def _sheet_weight_kg(material, espessura, comprimento, largura, fallback=0.0, family=""):
    esp = _num(espessura, 0)
    comp = _num(comprimento, 0)
    larg = _num(largura, 0)
    if esp <= 0 or comp <= 0 or larg <= 0:
        return float(fallback)
    dens = _guess_density(material, family)
    return round((dens * esp * larg * comp) / 1000000.0, 3)


def _material_sort_key(record):
    updated = str((record or {}).get("atualizado_em", "") or "").strip()
    qty = _num((record or {}).get("quantidade", 0), 0)
    return (updated, qty)


def _next_material_id(data):
    highest = 0
    for row in list((data or {}).get("materiais", []) or []):
        try:
            highest = max(highest, int(str((row or {}).get("id", "") or "").replace("MAT", "")))
        except Exception:
            continue
    return f"MAT{highest + 1:05d}"


def _find_material_template(data, formato, material, espessura, lote=""):
    materiais = list((data or {}).get("materiais", []) or [])
    mat_key = _norm_material_key(material)
    esp_key = str(espessura or "").strip()
    lote_key = str(lote or "").strip()
    formato_key = str(formato or "").strip() or "Chapa"
    exact_lote = []
    primary = []
    fallback = []
    for row in materiais:
        if not isinstance(row, dict):
            continue
        if _norm_material_key(row.get("material")) != mat_key:
            continue
        if str(row.get("espessura", "") or "").strip() != esp_key:
            continue
        row_formato = str(row.get("formato") or detect_materia_formato(row) or "").strip() or "Chapa"
        if row_formato != formato_key:
            continue
        if lote_key and str(row.get("lote_fornecedor", "") or "").strip() == lote_key and not _is_retalho_like(row):
            exact_lote.append(row)
            continue
        if not _is_retalho_like(row):
            primary.append(row)
        else:
            fallback.append(row)
    for bucket in (exact_lote, primary, fallback):
        if bucket:
            bucket.sort(key=_material_sort_key, reverse=True)
            return bucket[0]
    return None


def _hydrate_retalho_record(data, record, template=None):
    if not isinstance(record, dict):
        return record
    formato = str(record.get("formato") or detect_materia_formato(record) or "Chapa").strip() or "Chapa"
    record["formato"] = formato
    if not _is_retalho_like(record) or formato != "Chapa":
        record["preco_unid"] = _materia_preco_unid_record(record)
        return record

    base = template or _find_material_template(
        data,
        formato,
        record.get("material", ""),
        record.get("espessura", ""),
        record.get("lote_fornecedor", ""),
    )
    comp = _num(record.get("comprimento", 0), 0)
    larg = _num(record.get("largura", 0), 0)
    peso = 0.0
    if base is not None:
        base_comp = _num(base.get("comprimento", 0), 0)
        base_larg = _num(base.get("largura", 0), 0)
        base_area = base_comp * base_larg
        base_peso = _num(base.get("peso_unid", 0), 0)
        if base_area > 0 and base_peso > 0 and comp > 0 and larg > 0:
            peso = round(base_peso * ((comp * larg) / base_area), 3)
        if _num(record.get("p_compra", 0), 0) <= 0:
            record["p_compra"] = round(_num(base.get("p_compra", 0), 0), 6)
        if not str(record.get("lote_fornecedor", "") or "").strip():
            record["lote_fornecedor"] = str(base.get("lote_fornecedor", "") or "").strip()
        if not str(record.get("material_familia", "") or "").strip():
            record["material_familia"] = str(base.get("material_familia", "") or "").strip()
    if peso <= 0:
        family = str(record.get("material_familia", "") or (base or {}).get("material_familia", "") or "").strip()
        peso = _sheet_weight_kg(
            record.get("material", ""),
            record.get("espessura", ""),
            record.get("comprimento", 0),
            record.get("largura", 0),
            fallback=_num(record.get("peso_unid", 0), 0),
            family=family,
        )
    record["peso_unid"] = round(max(0.0, peso), 3)
    record["Localização"] = "RETALHO"
    record["Localizacao"] = "RETALHO"
    record["LocalizaÃ§Ã£o"] = "RETALHO"
    record["is_sobra"] = True
    record["preco_unid"] = round(_materia_preco_unid_record(record), 3)
    return record

def add_material(self):
    _ensure_configured()
    formato = (self.m_formato.get() or "Chapa").strip() or "Chapa"
    mat = self.m_material.get().strip()
    esp = self.m_espessura.get().strip()
    try:
        comp = float(self.m_comprimento.get())
        larg = float(self.m_largura.get())
        qtd = float(self.m_quantidade.get())
        metros = float(self.m_metros.get() or 0)
        peso_unid = float(self.m_peso_unid.get() or 0)
        p_compra = float(self.m_p_compra.get() or 0)
    except Exception:
        messagebox.showerror("Erro", "Comprimento, largura, metros e quantidade obrigatorios")
        return
    if not mat or not esp or qtd <= 0:
        messagebox.showerror("Erro", "Material, espessura e quantidade são obrigatórios")
        return
    if formato == "Chapa" and (comp <= 0 or larg <= 0):
        messagebox.showerror("Erro", "Para chapa, comprimento e largura são obrigatórios")
        return
    if formato == "Tubo" and metros <= 0:
        messagebox.showerror("Erro", "Para tubo, metros por unidade é obrigatório")
        return
    if formato == "Perfil" and peso_unid <= 0:
        messagebox.showerror("Erro", "Para perfil, peso por unidade é obrigatório")
        return
    record = {
        "id": _next_material_id(self.data),
        "formato": formato,
        "material": mat,
        "espessura": esp,
        "comprimento": comp,
        "largura": larg,
        "metros": metros,
        "quantidade": qtd,
        "reservado": 0.0,
        "Localização": self.m_local.get().strip(),
        "lote_fornecedor": self.m_lote.get().strip(),
        "peso_unid": peso_unid,
        "p_compra": p_compra,
        "preco_unid": _materia_preco_unid_record({
            "formato": formato,
            "metros": metros,
            "peso_unid": peso_unid,
            "p_compra": p_compra,
        }),
        "is_sobra": False,
        "atualizado_em": now_iso(),
    }
    record = _hydrate_retalho_record(self.data, record)
    self.data["materiais"].append(record)
    push_unique(self.data.setdefault("materiais_hist", []), mat)
    push_unique(self.data.setdefault("espessuras_hist", []), esp)
    log_stock(self.data, "ADICIONAR", f"{mat} {esp} qtd={qtd}")
    self.sync_all_ne_from_materia()
    save_data(self.data, force=True)
    self.refresh()

def edit_material(self):
    _ensure_configured()
    sel = self.tbl_materia.selection()
    if not sel:
        messagebox.showerror("Erro", "Selecione um material")
        return
    mat_vals = self.tbl_materia.item(sel[0], "values")
    mat_id = mat_vals[-1]
    for m in self.data["materiais"]:
        if m["id"] == mat_id:
            try:
                reservado = float(self.m_reservado.get() or 0)
            except Exception:
                messagebox.showerror("Erro", "Reservado invalido")
                return
            if reservado < 0:
                messagebox.showerror("Erro", "Reservado invalido")
                return
            if reservado > float(self.m_quantidade.get() or 0):
                messagebox.showerror("Erro", "Reservado maior que quantidade")
                return
            m.update({
                "formato": (self.m_formato.get() or m.get("formato") or "Chapa").strip() or "Chapa",
                "material": self.m_material.get().strip(),
                "espessura": self.m_espessura.get().strip(),
                "comprimento": float(self.m_comprimento.get() or 0),
                "largura": float(self.m_largura.get() or 0),
                "metros": float(self.m_metros.get() or 0),
                "quantidade": float(self.m_quantidade.get() or 0),
                "reservado": reservado,
                "Localização": self.m_local.get().strip(),
                "lote_fornecedor": self.m_lote.get().strip(),
                "peso_unid": float(self.m_peso_unid.get() or 0),
                "p_compra": float(self.m_p_compra.get() or 0),
                "atualizado_em": now_iso(),
            })
            _hydrate_retalho_record(self.data, m)
            m["preco_unid"] = _materia_preco_unid_record(m)
            log_stock(
                self.data,
                "EDITAR",
                f"{mat_id} qtd={m.get('quantidade', 0)} reservado={m.get('reservado', 0)}"
            )
            break
    self.sync_all_ne_from_materia()
    save_data(self.data, force=True)
    self.refresh()

def corrigir_stock(self):
    _ensure_configured()
    sel = self.tbl_materia.selection()
    if not sel:
        messagebox.showerror("Erro", "Selecione um material")
        return
    mat_vals = self.tbl_materia.item(sel[0], "values")
    mat_id = mat_vals[-1]
    m = next((item for item in self.data["materiais"] if item["id"] == mat_id), None)
    if not m:
        return
    if CUSTOM_TK_AVAILABLE:
        win = ctk.CTkToplevel(self.root)
        win.configure(fg_color="#f8fafc")
        LabelCls = ctk.CTkLabel
        EntryCls = ctk.CTkEntry
        BtnCls = ctk.CTkButton
    else:
        win = Toplevel(self.root)
        LabelCls = ttk.Label
        EntryCls = ttk.Entry
        BtnCls = ttk.Button
    win.title("Corrigir Stock")
    win.geometry("360x220")
    qtd_var = StringVar(value=str(m.get("quantidade", 0)))
    res_var = StringVar(value=str(m.get("reservado", 0)))
    metros_var = StringVar(value=str(m.get("metros", 0.0)))
    LabelCls(win, text="Quantidade").grid(row=0, column=0, sticky="w", padx=10, pady=6)
    EntryCls(win, textvariable=qtd_var, width=180).grid(row=0, column=1, padx=10, pady=6, sticky="ew")
    LabelCls(win, text="Reservado").grid(row=1, column=0, sticky="w", padx=10, pady=6)
    EntryCls(win, textvariable=res_var, width=180).grid(row=1, column=1, padx=10, pady=6, sticky="ew")
    LabelCls(win, text="Metros (m)").grid(row=2, column=0, sticky="w", padx=10, pady=6)
    EntryCls(win, textvariable=metros_var, width=180).grid(row=2, column=1, padx=10, pady=6, sticky="ew")

    def on_save():
        try:
            qtd = float(qtd_var.get())
            res = float(res_var.get())
            metros = float(metros_var.get() or 0)
        except Exception:
            messagebox.showerror("Erro", "Valores inválidos")
            return
        if qtd < 0 or res < 0 or res > qtd:
            messagebox.showerror("Erro", "Valores inválidos")
            return
        m["quantidade"] = qtd
        m["reservado"] = res
        m["metros"] = metros
        m["preco_unid"] = _materia_preco_unid_record(m)
        m["atualizado_em"] = now_iso()
        log_stock(self.data, "CORRIGIR", f"{m['id']} qtd={qtd} reservado={res}")
        self.sync_all_ne_from_materia()
        save_data(self.data, force=True)
        self.refresh()
        win.destroy()

    BtnCls(win, text="Guardar", command=on_save, width=140).grid(row=3, column=0, columnspan=2, pady=10)
    win.columnconfigure(1, weight=1)

def remove_material(self):
    _ensure_configured()
    sel = self.tbl_materia.selection()
    if not sel:
        messagebox.showerror("Erro", "Selecione um material")
        return
    mat_vals = self.tbl_materia.item(sel[0], "values")
    mat_id = mat_vals[-1]
    removed = None
    kept = []
    for m in self.data["materiais"]:
        if m["id"] == mat_id:
            removed = m
        else:
            kept.append(m)
    self.data["materiais"] = kept
    if removed:
        log_stock(
            self.data,
            "REMOVER",
            f"{removed.get('id')} qtd={removed.get('quantidade', 0)} reservado={removed.get('reservado', 0)}"
        )
    save_data(self.data, force=True)
    self.refresh()

def baixa_material(self):
    _ensure_configured()
    sel = self.tbl_materia.selection()
    if not sel:
        messagebox.showerror("Erro", "Selecione um material")
        return
    mat_vals = self.tbl_materia.item(sel[0], "values")
    mat_id = mat_vals[-1]
    for m in self.data["materiais"]:
        if m["id"] == mat_id:
            use_custom = CUSTOM_TK_AVAILABLE and getattr(self, "materia_use_custom", False)
            Win = ctk.CTkToplevel if use_custom else Toplevel
            Lbl = ctk.CTkLabel if use_custom else ttk.Label
            Ent = ctk.CTkEntry if use_custom else ttk.Entry
            Btn = ctk.CTkButton if use_custom else ttk.Button
            Frm = ctk.CTkFrame if use_custom else ttk.Frame
            win = Win(self.root)
            win.title("Baixa de material")
            try:
                if use_custom:
                    win.geometry("520x320")
                    win.configure(fg_color="#f7f8fb")
                win.transient(self.root)
                win.grab_set()
            except Exception:
                pass

            qtd_var = StringVar()
            sobra_comp_var = StringVar()
            sobra_larg_var = StringVar()
            sobra_qtd_var = StringVar()

            Lbl(win, text="Quantidade a dar baixa").grid(row=0, column=0, sticky="w", padx=8, pady=6)
            Ent(win, textvariable=qtd_var, width=180 if use_custom else None).grid(row=0, column=1, padx=6, pady=4, sticky="w")
            Lbl(win, text="Retalho (comprimento/largura) opcional").grid(row=1, column=0, sticky="w", padx=8, pady=6)
            retalho_row = Frm(win, fg_color="#f7f8fb") if use_custom else Frm(win)
            retalho_row.grid(row=1, column=1, sticky="w", padx=6, pady=4)
            Ent(retalho_row, textvariable=sobra_comp_var, width=90 if use_custom else 8).pack(side="left")
            Ent(retalho_row, textvariable=sobra_larg_var, width=90 if use_custom else 8).pack(side="left", padx=6)
            Lbl(win, text="Retalho (quantidade) opcional").grid(row=2, column=0, sticky="w", padx=8, pady=6)
            Ent(win, textvariable=sobra_qtd_var, width=180 if use_custom else None).grid(row=2, column=1, padx=6, pady=4, sticky="w")
            Lbl(win, text="Retalho (metros) opcional").grid(row=3, column=0, sticky="w", padx=8, pady=6)
            sobra_metros_var = StringVar()
            Ent(win, textvariable=sobra_metros_var, width=180 if use_custom else None).grid(row=3, column=1, padx=6, pady=4, sticky="w")

            def on_save():
                try:
                    qtd = float(qtd_var.get())
                except ValueError:
                    messagebox.showerror("Erro", "Quantidade invalida")
                    return
                if qtd <= 0 or qtd > m["quantidade"]:
                    messagebox.showerror("Erro", "Quantidade invalida")
                    return
                m["quantidade"] -= qtd
                m["atualizado_em"] = now_iso()
                log_stock(self.data, "BAIXA", f"{m['id']} qtd={qtd}")

                sc = sobra_comp_var.get().strip()
                sl = sobra_larg_var.get().strip()
                sq = sobra_qtd_var.get().strip()
                sm = sobra_metros_var.get().strip()
                if sc or sl or sq or sm:
                    try:
                        scf = float(sc)
                        slf = float(sl)
                        sqf = float(sq)
                        smf = float(sm) if sm else 0.0
                    except Exception:
                        messagebox.showerror("Erro", "Valores de sobra invalidos")
                        return
                    retalho = {
                        "id": _next_material_id(self.data),
                        "formato": m.get("formato", detect_materia_formato(m)),
                        "material": m["material"],
                        "espessura": m["espessura"],
                        "comprimento": scf,
                        "largura": slf,
                        "metros": smf,
                        "quantidade": sqf,
                        "reservado": 0.0,
                        "Localização": _get_localizacao(m),
                        "lote_fornecedor": m.get("lote_fornecedor", ""),
                        "peso_unid": 0.0,
                        "p_compra": m.get("p_compra", 0),
                        "preco_unid": 0.0,
                        "is_sobra": True,
                        "atualizado_em": now_iso(),
                    }
                    _hydrate_retalho_record(self.data, retalho, template=m)
                    self.data["materiais"].append(retalho)
                    log_stock(self.data, "RETALHO", f"{m['id']} qtd={sqf}")
                save_data(self.data, force=True)
                self.refresh()
                win.destroy()

            Btn(win, text="Guardar", command=on_save, width=150 if use_custom else None).grid(row=4, column=0, columnspan=2, pady=10)
            return
    self.refresh()

def refresh_materia(self):
    _ensure_configured()
    query = (self.m_filter.get().strip().lower() if hasattr(self, "m_filter") else "")
    selected_id = ""
    try:
        sel = self.tbl_materia.selection()
        if sel:
            vals = self.tbl_materia.item(sel[0], "values")
            if vals:
                selected_id = str(vals[-1] or "")
    except Exception:
        selected_id = ""
    rows = []
    for idx, m in enumerate(self.data["materiais"]):
        m["preco_unid"] = _materia_preco_unid_record(m)
        disponivel = m["quantidade"] - m.get("reservado", 0)
        formato = m.get("formato", detect_materia_formato(m))
        tipo = "Retalho" if m.get("is_sobra") else "Normal"
        values = (
            m.get("lote_fornecedor", ""),
            m["material"],
            fmt_num(m["comprimento"]),
            fmt_num(m["largura"]),
            fmt_num(m["espessura"]),
            fmt_num(m["quantidade"]),
            fmt_num(m.get("reservado", 0)),
            formato,
            fmt_num(m.get("metros", 0.0)),
            fmt_num(m.get("peso_unid", 0)),
            fmt_num(m.get("p_compra", 0)),
            fmt_num(m.get("preco_unid", 0)),
            fmt_num(disponivel),
            f"{formato} / {tipo}",
            _get_localizacao(m),
            m["id"],
        )
        if query and not any(query in str(v).lower() for v in values):
            continue
        if float(m.get("quantidade", 0)) == 1:
            tag = "stock_one"
        elif disponivel <= STOCK_VERMELHO:
            tag = "stock_crit"
        elif disponivel <= STOCK_AMARELO:
            tag = "stock_low"
        else:
            tag = "stock_ok"
        alt_tag = "odd" if idx % 2 else "even"
        rows.append((values, (alt_tag, tag)))

    def _after_fill():
        if not selected_id:
            return
        try:
            for iid in self.tbl_materia.get_children():
                vals = self.tbl_materia.item(iid, "values")
                if vals and str(vals[-1] or "") == selected_id:
                    self.tbl_materia.selection_set(iid)
                    self.tbl_materia.see(iid)
                    break
        except Exception:
            pass

    if hasattr(self, "fill_treeview_in_batches"):
        self.fill_treeview_in_batches(self.tbl_materia, rows, "tbl_materia", chunk_size=160, on_done=_after_fill)
        return
    children = self.tbl_materia.get_children()
    if children:
        self.tbl_materia.delete(*children)
    for values, tags in rows:
        self.tbl_materia.insert("", END, values=values, tags=tags)
    _after_fill()

def show_stock_log(self):
    _ensure_configured()
    prev_dir = os.path.join(BASE_DIR, "previews")
    try:
        os.makedirs(prev_dir, exist_ok=True)
    except Exception:
        prev_dir = tempfile.gettempdir()
    path = os.path.join(prev_dir, "historico_materia_prima.pdf")
    try:
        self.render_stock_log_pdf(path)
        ok = self._open_pdf_default(path)
        if not ok:
            messagebox.showwarning("Aviso", f"PDF gerado, mas não abriu automaticamente.\n{path}")
    except Exception as ex:
        messagebox.showerror("Erro", f"Falha ao gerar PDF do histórico: {ex}")

def on_select_material(self, _):
    _ensure_configured()
    enc = self.get_selected_encomenda()
    if not enc:
        return
    sel = self.tbl_materiais.selection()
    if not sel:
        return
    mat = self.tbl_materiais.item(sel[0], "values")[0]
    self.selected_material = mat
    self.refresh_espessuras(enc, mat)
    self.refresh_pecas(enc, None)
    self.update_cativadas_display(enc)

def on_select_stock_material(self, _):
    _ensure_configured()
    sel = self.tbl_materia.selection()
    if not sel:
        return
    values = self.tbl_materia.item(sel[0], "values")
    self.m_formato.set(values[7] if len(values) > 7 else "Chapa")
    self.m_material.set(values[1] if len(values) > 1 else "")
    self.m_espessura.set(values[4] if len(values) > 4 else "")
    self.m_comprimento.set(values[2] if len(values) > 2 else 0)
    self.m_largura.set(values[3] if len(values) > 3 else 0)
    self.m_metros.set(values[8] if len(values) > 8 else 0)
    self.m_peso_unid.set(values[9] if len(values) > 9 else 0)
    self.m_p_compra.set(values[10] if len(values) > 10 else 0)
    self.m_quantidade.set(values[5] if len(values) > 5 else 0)
    self.m_reservado.set(values[6] if len(values) > 6 else 0)
    self.m_local.set(values[14] if len(values) > 14 else "")
    self.m_lote.set(values[0])
    try:
        mid = values[-1]
        mobj = next((m for m in self.data.get("materiais", []) if m.get("id") == mid), None)
        if mobj:
            self.m_peso_unid.set(parse_float(mobj.get("peso_unid", 0), 0))
            self.m_p_compra.set(parse_float(mobj.get("p_compra", 0), 0))
    except Exception:
        pass

def render_stock_log_pdf(self, path):
    _ensure_configured()
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as pdf_canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib import colors
    from reportlab.lib.utils import ImageReader

    c = pdf_canvas.Canvas(path, pagesize=A4)
    w, h = A4
    m = 20
    rows = list(self.data.get("stock_log", []))

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
    except Exception:
        pass

    def set_font(bold=False, size=9):
        c.setFont(font_bold if bold else font_regular, size)

    def clip_text(txt, max_w, bold=False, size=8):
        s = "" if txt is None else str(txt)
        fname = font_bold if bold else font_regular
        if pdfmetrics.stringWidth(s, fname, size) <= max_w:
            return s
        ell = "..."
        while s and pdfmetrics.stringWidth(s + ell, fname, size) > max_w:
            s = s[:-1]
        return (s + ell) if s else ""

    def draw_logo(x, y_top, max_w=72, max_h=24):
        logo = get_orc_logo_path()
        if not logo or not os.path.exists(logo):
            return
        try:
            img = ImageReader(logo)
            c.drawImage(img, x, h - y_top - max_h, width=max_w, height=max_h, preserveAspectRatio=True, mask="auto")
        except Exception:
            pass

    cols = [
        ("Data", 115, "w"),
        ("Ação", 85, "w"),
        ("Detalhes", 350, "w"),
    ]
    table_w = sum(cw for _, cw, _ in cols)
    header_h = 18
    row_h = 16

    def draw_header(page):
        c.setStrokeColor(colors.HexColor("#c6cfdb"))
        c.rect(m - 4, m - 4, w - (2 * (m - 4)), h - (2 * (m - 4)), stroke=1, fill=0)
        c.setStrokeColor(colors.black)
        draw_logo(m + 2, 24, 74, 24)
        title_left = m + 92
        title_right = w - m - 150
        title_w = min(280, max(210, title_right - title_left))
        title_x = title_left + max(0.0, ((title_right - title_left) - title_w) / 2.0)
        title_y = 20
        title_h = 22
        c.setLineWidth(1.15)
        c.setStrokeColor(colors.HexColor("#8b1e2d"))
        c.setFillColor(colors.white)
        c.roundRect(title_x, h - title_y - title_h, title_w, title_h, 6, stroke=1, fill=1)
        c.setFillColor(colors.HexColor("#8b1e2d"))
        set_font(False, 8)
        c.drawCentredString(title_x + (title_w / 2), h - title_y + 2, "Historico")
        set_font(True, 12)
        c.drawCentredString(title_x + (title_w / 2), h - title_y - 14, "Historico de Materia-Prima")
        c.setFillColor(colors.black)
        set_font(False, 9)
        c.drawRightString(w - m, h - 36, datetime.now().strftime("%d/%m/%Y %H:%M"))
        c.drawRightString(w - m, h - 50, f"Página {page}")

        y = 76
        c.setStrokeColor(colors.HexColor("#8b1e2d"))
        c.setFillColor(colors.white)
        c.rect(m, h - y - header_h, table_w, header_h, stroke=1, fill=1)
        c.setFillColor(colors.HexColor("#8b1e2d"))
        set_font(True, 8.5)
        xx = m
        for name, cw, align in cols:
            if align == "e":
                c.drawRightString(xx + cw - 3, h - y - 12, name)
            elif align == "center":
                c.drawCentredString(xx + cw / 2, h - y - 12, name)
            else:
                c.drawString(xx + 3, h - y - 12, name)
            xx += cw
        c.setFillColor(colors.black)
        c.setStrokeColor(colors.black)
        return y + header_h

    i = 0
    page = 1
    while i < len(rows) or i == 0:
        y_top = draw_header(page)
        max_rows = int((h - (m + 30) - y_top) // row_h)
        if max_rows < 1:
            max_rows = 1
        chunk = rows[i:i + max_rows]
        set_font(False, 8)
        for idx, item in enumerate(chunk):
            y = y_top + idx * row_h
            fill = colors.HexColor("#fff8f9") if ((i + idx) % 2 == 0) else colors.HexColor("#fff0f2")
            c.setFillColor(fill)
            c.rect(m, h - y - row_h, table_w, row_h, stroke=1, fill=1)
            c.setFillColor(colors.black)
            vals = [
                str(item.get("data", ""))[:19].replace("T", " "),
                item.get("acao", ""),
                item.get("detalhes", ""),
            ]
            xx = m
            for (_, cw, align), val in zip(cols, vals):
                txt = clip_text(val, cw - 6, False, 8)
                if align == "e":
                    c.drawRightString(xx + cw - 3, h - y - 11, txt)
                elif align == "center":
                    c.drawCentredString(xx + cw / 2, h - y - 11, txt)
                else:
                    c.drawString(xx + 3, h - y - 11, txt)
                xx += cw
        i += len(chunk)
        if i < len(rows):
            c.showPage()
            page += 1
        else:
            break
    c.save()

def render_stock_a4_pdf(self, path_pdf):
    _ensure_configured()
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.pdfgen import canvas as pdf_canvas

    fonts = _mat_pdf_register_fonts()
    palette = _mat_pdf_palette()
    width, height = landscape(A4)
    c = pdf_canvas.Canvas(path_pdf, pagesize=landscape(A4))
    margin = 22
    content_w = width - (2 * margin)
    row_h = 18
    table_header_h = 20
    table_first_top = 194
    table_next_top = 88
    footer_top = 474

    def ntxt(value):
        try:
            return pdf_normalize_text(value)
        except Exception:
            return str(value or "")

    def yinv(top_y):
        return height - top_y

    def draw_logo(x, top_y, box_w=124, box_h=52):
        drawer = globals().get("draw_pdf_logo_box")
        if callable(drawer):
            try:
                drawer(c, height, x, top_y, box_size=box_w, box_h=box_h, padding=4, draw_border=False)
                return
            except Exception:
                pass

    def draw_logo_plate(x, top_y, box_w=124, box_h=52):
        drawer = globals().get("draw_pdf_logo_plate")
        if callable(drawer):
            try:
                drawer(c, height, x, top_y, box_w=box_w, box_h=box_h, padding=4)
                return
            except Exception:
                pass
        draw_logo(x, top_y, box_w=box_w, box_h=box_h)

    def draw_title_block(left_x, right_x, title, subtitle):
        area_w = max(40.0, float(right_x) - float(left_x))
        center_x = float(left_x) + (area_w / 2.0)
        title_size = _mat_pdf_fit_font_size(title, fonts["bold"], area_w, 22.0, 16.0)
        subtitle_size = _mat_pdf_fit_font_size(subtitle, fonts["regular"], area_w, 10.0, 8.0)
        c.setFillColor(palette["primary_dark"])
        c.setFont(fonts["bold"], title_size)
        c.drawCentredString(center_x, yinv(46), ntxt(_mat_pdf_clip_text(title, area_w, fonts["bold"], title_size)))
        c.setFont(fonts["regular"], subtitle_size)
        c.setFillColor(palette["muted"])
        c.drawCentredString(center_x, yinv(63), ntxt(_mat_pdf_clip_text(subtitle, area_w, fonts["regular"], subtitle_size)))

    def draw_header_panel(x, top_y, box_w, box_h):
        drawer = globals().get("draw_pdf_header_panel")
        if callable(drawer):
            try:
                drawer(c, height, x, top_y, box_w, box_h, radius=12, stroke_color="#D5DDE7", accent_color="#EAF0F6", accent_height=5)
                return
            except Exception:
                pass
        c.saveState()
        c.setFillColor(colors.white)
        c.setStrokeColor(colors.HexColor("#d5dde7"))
        c.setLineWidth(1.0)
        c.roundRect(x, yinv(top_y + box_h), box_w, box_h, 12, stroke=1, fill=1)
        c.restoreState()

    def card(x, top_y, box_w, box_h, title, lines_, accent=False, font_size=8.2):
        title_fill = palette["primary_soft"] if accent else palette["primary_soft_2"]
        c.saveState()
        c.setFillColor(colors.white)
        c.setStrokeColor(palette["line_strong"] if accent else palette["line"])
        c.setLineWidth(0.9)
        c.roundRect(x, yinv(top_y + box_h), box_w, box_h, 8, stroke=1, fill=1)
        c.setFillColor(title_fill)
        c.roundRect(x, yinv(top_y + 20), box_w, 20, 8, stroke=0, fill=1)
        c.restoreState()
        c.setFont(fonts["bold"], 8.6)
        c.setFillColor(palette["primary_dark"])
        c.drawString(x + 8, yinv(top_y + 13), ntxt(title))
        yy = top_y + 34
        for idx_line, line in enumerate(lines_):
            line_font = fonts["bold"] if idx_line == 0 and accent else fonts["regular"]
            line_size = font_size + 0.5 if idx_line == 0 and accent else font_size
            wrapped = _mat_pdf_wrap_text(line, line_font, line_size, box_w - 16, max_lines=2 if idx_line == 0 else 1)
            for item in wrapped:
                c.setFont(line_font, line_size)
                c.setFillColor(palette["ink"])
                c.drawString(x + 8, yinv(yy), ntxt(item))
                yy += 10
            if yy > top_y + box_h - 8:
                break

    def info_chip(x, top_y, box_w, label, value, box_h=34):
        c.saveState()
        c.setFillColor(colors.white)
        c.setStrokeColor(palette["line"])
        c.roundRect(x, yinv(top_y + box_h), box_w, box_h, 8, stroke=1, fill=1)
        c.restoreState()
        c.setFont(fonts["regular"], 7.0)
        c.setFillColor(palette["muted"])
        c.drawString(x + 7, yinv(top_y + 11), ntxt(label))
        value_size = _mat_pdf_fit_font_size(value, fonts["bold"], box_w - 14, 10.2, 8.2)
        c.setFont(fonts["bold"], value_size)
        c.setFillColor(palette["primary_dark"])
        c.drawString(x + 7, yinv(top_y + 24), ntxt(_mat_pdf_clip_text(value, box_w - 14, fonts["bold"], value_size)))

    def dimension_text(record, formato):
        comp = fmt_num(record.get("comprimento", 0))
        larg = fmt_num(record.get("largura", 0))
        metros = fmt_num(record.get("metros", 0))
        if formato == "Chapa":
            if comp != "0" and larg != "0":
                return f"{comp} x {larg} mm"
            return "-"
        if metros != "0":
            return f"{metros} m"
        if comp != "0" or larg != "0":
            return f"{comp} x {larg}"
        return "-"

    cols = [
        ("Formato", 68, "w"),
        ("Material", 150, "w"),
        ("Esp.", 44, "center"),
        ("Dimensao", 132, "w"),
        ("Qtd.", 44, "e"),
        ("Res.", 44, "e"),
        ("Disp.", 46, "e"),
        ("Localizacao", 116, "w"),
        ("Lote", 132, "w"),
    ]
    table_w = sum(cw for _, cw, _ in cols)
    x0 = margin

    source_rows = list(self.data.get("materiais", []) or [])
    source_rows = sorted(
        source_rows,
        key=lambda m: (
            str(m.get("formato") or detect_materia_formato(m) or "Chapa"),
            float(parse_float(m.get("espessura", 0), 0)),
            str(m.get("material", "") or ""),
            str(m.get("Localizacao", m.get("LocalizaÃ§Ã£o", "")) or ""),
        ),
    )

    total_qty = 0.0
    total_reserved = 0.0
    total_available = 0.0
    retalho_count = 0
    format_counts = {}
    rows = []
    last_section = None
    zebra_idx = 0
    for record in source_rows:
        formato = str(record.get("formato") or detect_materia_formato(record) or "Chapa").strip() or "Chapa"
        esp = str(record.get("espessura", "") or "").strip() or "-"
        qty = _num(record.get("quantidade", 0), 0)
        reserved = _num(record.get("reservado", 0), 0)
        available = qty - reserved
        is_retalho = _is_retalho_like(record)
        total_qty += qty
        total_reserved += reserved
        total_available += available
        if is_retalho:
            retalho_count += 1
        format_counts[formato] = int(format_counts.get(formato, 0)) + 1
        section_key = (formato, esp)
        if section_key != last_section:
            rows.append({"_group": f"{formato} | Esp. {esp}"})
            last_section = section_key
        item = dict(record)
        item["_formato"] = formato
        item["_available"] = available
        item["_retalho"] = is_retalho
        item["_material"] = str(record.get("material", "") or "").strip() or "-"
        item["_zebra"] = zebra_idx
        zebra_idx += 1
        rows.append(item)

    top_formats = sorted(format_counts.items(), key=lambda item: (-item[1], item[0]))[:4]

    def draw_table_header(y_top):
        c.saveState()
        c.setFillColor(palette["primary_soft"])
        c.setStrokeColor(palette["line"])
        c.setLineWidth(0.8)
        c.roundRect(x0, yinv(y_top + table_header_h), table_w, table_header_h, 7, stroke=1, fill=1)
        c.restoreState()
        c.setFillColor(palette["primary_dark"])
        c.setFont(fonts["bold"], 8.2)
        xx = x0
        for name, cw, align in cols:
            if align == "e":
                c.drawRightString(xx + cw - 6, yinv(y_top + 13), ntxt(name))
            elif align == "center":
                c.drawCentredString(xx + cw / 2, yinv(y_top + 13), ntxt(name))
            else:
                c.drawString(xx + 6, yinv(y_top + 13), ntxt(name))
            xx += cw

    def draw_header(page_no, total_pages, first_page):
        if first_page:
            logo_w = 124
            logo_gap = 14
            banner_x = margin + logo_w + logo_gap
            banner_w = content_w - logo_w - logo_gap
            metric_grid = _mat_pdf_metric_grid_layout(width - margin, 20, 82, cols=3, rows=2, group_w=338, gap=6)
            draw_header_panel(banner_x, 20, banner_w, 82)
            draw_logo_plate(margin, 30)
            draw_title_block(banner_x + 14, metric_grid["group_x"] - 14, "Stock de Materia-Prima", "Chapas, perfis, tubos e retalhos com leitura consolidada")
            info_chip(metric_grid["cols"][0], metric_grid["rows"][0], metric_grid["chip_w"], "Emitido", datetime.now().strftime("%d/%m/%Y"))
            info_chip(metric_grid["cols"][1], metric_grid["rows"][0], metric_grid["chip_w"], "Pagina", f"{page_no}/{total_pages}")
            info_chip(metric_grid["cols"][2], metric_grid["rows"][0], metric_grid["chip_w"], "Formatos", str(len(format_counts)))
            info_chip(metric_grid["cols"][0], metric_grid["rows"][1], metric_grid["chip_w"], "Registos", str(len(source_rows)))
            info_chip(metric_grid["cols"][1], metric_grid["rows"][1], metric_grid["chip_w"], "Retalhos", str(retalho_count))
            info_chip(metric_grid["cols"][2], metric_grid["rows"][1], metric_grid["chip_w"], "Disponivel", fmt_num(total_available))
            card(
                margin,
                126,
                430,
                62,
                "Cobertura de stock",
                [
                    f"Registos ativos: {len(source_rows)} | Formatos: {len(format_counts)}",
                    f"Qtd total: {fmt_num(total_qty)} | Reservado: {fmt_num(total_reserved)}",
                    f"Disponivel: {fmt_num(total_available)} | Retalhos: {retalho_count}",
                ],
                accent=True,
            )
            format_lines = [f"{name}: {count} registos" for name, count in top_formats] or ["Sem materiais registados."]
            card(
                margin + 442,
                126,
                content_w - 442,
                62,
                "Formatos principais",
                format_lines,
                font_size=8.0,
            )
            return draw_table_header(table_first_top)

        c.saveState()
        c.setFillColor(palette["primary_soft"])
        c.roundRect(margin, yinv(68), content_w, 48, 12, stroke=0, fill=1)
        c.restoreState()
        c.setFillColor(palette["primary_dark"])
        c.setFont(fonts["bold"], 15)
        c.drawString(margin + 12, yinv(40), ntxt("Stock de Materia-Prima"))
        c.setFont(fonts["regular"], 8.6)
        c.drawString(margin + 12, yinv(54), ntxt(f"Disponivel: {fmt_num(total_available)} | Retalhos: {retalho_count}"))
        metric_grid = _mat_pdf_metric_grid_layout(width - margin, 14, 48, cols=2, rows=1, group_w=196, gap=6)
        info_chip(metric_grid["cols"][0], metric_grid["rows"][0], metric_grid["chip_w"], "Emitido", datetime.now().strftime("%d/%m/%Y"), box_h=metric_grid["chip_h"])
        info_chip(metric_grid["cols"][1], metric_grid["rows"][0], metric_grid["chip_w"], "Pagina", f"{page_no}/{total_pages}", box_h=metric_grid["chip_h"])
        return draw_table_header(table_next_top)

    def draw_footer(page_no, total_pages):
        footer_lines = list(get_empresa_rodape_lines() or [])
        card(
            margin,
            486,
            474,
            72,
            "Conferencia de inventario",
            [
                "Data: ____________________",
                "Responsavel: ________________________________",
                "Observacoes: Validar lotes, localizacao e diferencas de stock.",
            ],
            font_size=7.9,
        )
        card(
            margin + 486,
            486,
            content_w - 486,
            72,
            "Resumo de stock",
            [
                f"Qtd total: {fmt_num(total_qty)} | Reservado: {fmt_num(total_reserved)}",
                f"Disponivel: {fmt_num(total_available)} | Retalhos: {retalho_count}",
                f"Pagina: {page_no}/{total_pages}",
            ],
            accent=True,
            font_size=8.0,
        )
        c.setFillColor(palette["muted"])
        c.setFont(fonts["regular"], 7.0)
        if footer_lines:
            c.drawString(margin, yinv(570), ntxt(" | ".join(footer_lines[:2])))
        c.drawRightString(width - margin, yinv(570), ntxt(f"luGEST | {datetime.now().strftime('%Y-%m-%d %H:%M')}"))

    total_pages = 1
    remaining_preview = len(rows)
    while True:
        first = total_pages == 1
        table_top = table_first_top if first else table_next_top
        usable_last = footer_top - (table_top + table_header_h + 6)
        usable_non = (height - margin - 28) - (table_top + table_header_h + 6)
        fit_last = max(1, int(usable_last // row_h))
        fit_non = max(1, int(usable_non // row_h))
        if remaining_preview <= fit_last:
            break
        remaining_preview -= fit_non
        total_pages += 1

    idx = 0
    page_no = 1
    while idx < len(rows) or idx == 0:
        first = page_no == 1
        draw_header(page_no, total_pages, first)
        table_top = table_first_top if first else table_next_top
        usable_last = footer_top - (table_top + table_header_h + 6)
        usable_non = (height - margin - 28) - (table_top + table_header_h + 6)
        fit_last = max(1, int(usable_last // row_h))
        fit_non = max(1, int(usable_non // row_h))
        remaining = len(rows) - idx
        if remaining <= fit_last:
            is_last = True
            count = remaining
        else:
            is_last = False
            count = min(fit_non, remaining)

        y_top = table_top + table_header_h
        for local_i in range(count):
            row = rows[idx + local_i]
            y_row = y_top + (local_i * row_h)
            if row.get("_group") is not None:
                c.saveState()
                c.setFillColor(palette["primary_soft"])
                c.setStrokeColor(palette["line"])
                c.roundRect(x0, yinv(y_row + row_h), table_w, row_h, 5, stroke=1, fill=1)
                c.restoreState()
                c.setFillColor(palette["primary_dark"])
                c.setFont(fonts["bold"], 8.4)
                c.drawString(x0 + 6, yinv(y_row + 11.5), ntxt(str(row.get("_group", ""))))
                continue

            available = _num(row.get("_available", 0), 0)
            if available <= 0:
                fill = palette["danger_fill"]
                text_color = palette["danger_text"]
            elif bool(row.get("_retalho")):
                fill = palette["retalho_fill"]
                text_color = palette["retalho_text"]
            else:
                fill = palette["surface_warm"] if (int(row.get("_zebra", 0)) % 2 == 0) else colors.white
                text_color = palette["ink"]
            c.saveState()
            c.setFillColor(fill)
            c.setStrokeColor(palette["line"])
            c.setLineWidth(0.45)
            c.roundRect(x0, yinv(y_row + row_h), table_w, row_h, 5, stroke=1, fill=1)
            c.restoreState()
            c.setFillColor(text_color)
            c.setFont(fonts["regular"], 8.0)

            vals = [
                row.get("_formato", ""),
                row.get("_material", ""),
                str(row.get("espessura", "") or ""),
                dimension_text(row, str(row.get("_formato", "") or "")),
                fmt_num(row.get("quantidade", 0)),
                fmt_num(row.get("reservado", 0)),
                fmt_num(available),
                row.get("Localizacao", row.get("LocalizaÃ§Ã£o", "")),
                row.get("lote_fornecedor", "") or "-",
            ]
            xx = x0
            for (_name, cw, align), value in zip(cols, vals):
                txt = _mat_pdf_clip_text(ntxt(value), cw - 12, fonts["regular"], 8.0)
                if align == "e":
                    c.drawRightString(xx + cw - 6, yinv(y_row + 11.5), txt)
                elif align == "center":
                    c.drawCentredString(xx + cw / 2, yinv(y_row + 11.5), txt)
                else:
                    c.drawString(xx + 6, yinv(y_row + 11.5), txt)
                xx += cw

        idx += count
        if is_last:
            draw_footer(page_no, total_pages)
            break
        c.setFont(fonts["regular"], 7.4)
        c.setFillColor(palette["muted"])
        c.drawRightString(width - margin, yinv(height - margin - 4), ntxt("Continua na proxima pagina..."))
        c.showPage()
        page_no += 1

    c.save()

def preview_stock_a4(self):
    _ensure_configured()
    path = _temp_pdf_path("lugest_stock")
    try:
        render_stock_a4_pdf(self, path)
    except Exception as exc:
        messagebox.showerror("Erro", f"Falha ao gerar PDF: {exc}")
        return
    try:
        os.startfile(path)
    except Exception:
        messagebox.showerror("Erro", "Nao foi possivel abrir o PDF do stock.")
    return

    def render_pdf(path_pdf):
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas as pdf_canvas
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        from reportlab.lib.utils import ImageReader
        width, height = A4
        c = pdf_canvas.Canvas(path_pdf, pagesize=A4)

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

        def yinv(y):
            return height - y

        def draw_logo(x, y, w, h):
            logo = get_orc_logo_path()
            if not logo or not os.path.exists(logo):
                return
            try:
                img = ImageReader(logo)
                c.drawImage(img, x, yinv(y + h), width=w, height=h, preserveAspectRatio=True, mask="auto")
            except Exception:
                pass

        def fit_text(text, max_w, font_name, font_size):
            s = pdf_normalize_text(text)
            if pdfmetrics.stringWidth(s, font_name, font_size) <= max_w:
                return s
            ell = "..."
            max_w = max(0, max_w - pdfmetrics.stringWidth(ell, font_name, font_size))
            while s and pdfmetrics.stringWidth(s, font_name, font_size) > max_w:
                s = s[:-1]
            return (s + ell) if s else ell

        margin = 20

        def draw_std_header(page_num):
            c.setStrokeColorRGB(0.78, 0.80, 0.84)
            c.rect(margin, yinv(height - margin), width - margin * 2, height - margin * 2, stroke=1, fill=0)
            draw_logo(margin + 4, margin + 6, 90, 40)

            right_info_w = 150
            title_area_left = margin + 100
            title_area_right = width - margin - right_info_w
            title_w = min(260, max(190, title_area_right - title_area_left))
            title_left = title_area_left + max(0.0, ((title_area_right - title_area_left) - title_w) / 2.0)
            title_y = margin + 10
            title_h = 24
            c.setStrokeColorRGB(0.48, 0.06, 0.12)
            c.setLineWidth(1.15)
            c.roundRect(title_left, yinv(title_y + title_h), title_w, title_h, 6, stroke=1, fill=0)
            c.setFillColorRGB(0.55, 0.12, 0.18)
            set_font(True, 13)
            c.drawCentredString(title_left + (title_w / 2), yinv(title_y + 16), "Stock de Chapas")

            c.setLineWidth(1)
            c.setFillColorRGB(0, 0, 0)
            set_font(False, 9)
            c.drawRightString(width - margin - 4, yinv(margin + 18), f"Data: {datetime.now().strftime('%Y-%m-%d')}")
            c.drawRightString(width - margin - 4, yinv(margin + 32), f"Pagina {page_num}")
            c.line(margin, yinv(62), width - margin, yinv(62))

        col_specs = [
            ("Material", 130),
            ("Esp.", 40),
            ("Dimensao", 100),
            ("Qtd.", 40),
            ("Res.", 40),
            ("Disp.", 45),
            ("Localizacao", 85),
            ("Lote", 75),
        ]
        table_w = sum(wc for _, wc in col_specs)
        table_left = max(margin, (width - table_w) / 2.0)
        cols = []
        _x = table_left
        for title, wc in col_specs:
            cols.append((title, _x))
            _x += wc
        x_list = [x for _, x in cols]
        x_ends = x_list[1:] + [table_left + table_w]
        col_w = [x_ends[i] - x_list[i] for i in range(len(x_list))]
        header_h = 18
        header_fs = 9
        row_h = 22
        row_fs = 9
        text_offset = (row_h / 2) + (row_fs / 2 - 1)

        mats_sorted = sorted(self.data.get("materiais", []), key=lambda m: (float(m.get("espessura", 0) or 0), m.get("material","")))
        rows = []
        current_esp = None
        for m in mats_sorted:
            esp_val = str(m.get("espessura", ""))
            if current_esp != esp_val:
                current_esp = esp_val
                rows.append(("section", [f"Espessura: {current_esp}"]))
            disp = float(m.get("quantidade", 0)) - float(m.get("reservado", 0))
            dim = f"{m.get('comprimento','')}x{m.get('largura','')}"
            rows.append((
                "row",
                [
                    m.get("material", ""),
                    str(m.get("espessura", "")),
                    dim,
                    str(m.get("quantidade", 0)),
                    str(m.get("reservado", 0)),
                    str(disp),
                    m.get("Localizacao", m.get("Localização", "")),
                    m.get("lote_fornecedor", ""),
                ],
            ))

        def draw_page_header(page_num):
            draw_std_header(page_num)

            set_font(True, 9)
            y = 72
            header_text_y = y + (header_h / 2) + (header_fs / 2 - 1)
            c.setFillColorRGB(0.90, 0.94, 0.98)
            c.rect(table_left, yinv(y + header_h), table_w, header_h, fill=1, stroke=0)
            c.setFillColorRGB(0, 0, 0)
            for (title, x), w in zip(cols, col_w):
                c.drawCentredString(x + w / 2, yinv(header_text_y), fit_text(title, w - 6, font_bold, header_fs))
            c.line(table_left, yinv(y), table_left + table_w, yinv(y))
            c.line(table_left, yinv(y + header_h), table_left + table_w, yinv(y + header_h))
            return y + header_h + 4, y

        def draw_signature():
            box_h = 40
            box_y = height - margin - box_h
            c.setStrokeColorRGB(0.6, 0.65, 0.7)
            c.rect(margin, yinv(box_y + box_h), width - margin * 2, box_h, stroke=1, fill=0)
            c.setStrokeColorRGB(0, 0, 0)
            set_font(True, 8)
            c.drawString(margin + 6, yinv(box_y + 14), "Data:")
            c.drawString(margin + 180, yinv(box_y + 14), "Assinatura:")

        def draw_footer():
            set_font(False, 7)
            footer_y = 30
            for idx, line in enumerate(get_empresa_rodape_lines()):
                c.drawString(margin, footer_y + idx * 9, line)

        idx = 0
        page_num = 1
        while idx < len(rows) or idx == 0:
            y, table_top = draw_page_header(page_num)
            # available space
            remaining = len(rows) - idx
            max_y_full = height - margin - 12
            max_y_sig = height - margin - 52
            rows_fit_last = max(0, int((max_y_sig - y) // row_h))
            rows_fit_full = max(0, int((max_y_full - y) // row_h))
            if remaining <= rows_fit_last:
                rows_fit = remaining
                last_page = True
            else:
                last_page = False
                if remaining <= rows_fit_full:
                    rows_fit = max(1, remaining - rows_fit_last)
                else:
                    rows_fit = rows_fit_full

            end = min(len(rows), idx + rows_fit)
            set_font(False, 9)
            while idx < end:
                kind, payload = rows[idx]
                if kind == "section":
                    set_font(True, 9)
                    c.drawString(table_left + 4, yinv(y + text_offset), payload[0])
                    c.line(table_left, yinv(y + row_h), table_left + table_w, yinv(y + row_h))
                    y += row_h
                    set_font(False, 9)
                else:
                    for ((title, x), val, w) in zip(cols, payload, col_w):
                        c.drawCentredString(x + w / 2, yinv(y + text_offset), fit_text(val, w - 6, font_regular, row_fs))
                    c.line(table_left, yinv(y + row_h), table_left + table_w, yinv(y + row_h))
                    y += row_h
                idx += 1

            table_bottom = y
            x_lines = [table_left] + [x for _, x in cols[1:]] + [table_left + table_w]
            for x in x_lines:
                c.line(x, yinv(table_top), x, yinv(table_bottom))

            if last_page:
                draw_signature()
                draw_footer()
            if idx >= len(rows):
                if last_page:
                    break
                c.showPage()
                page_num += 1
                continue
            c.showPage()
            page_num += 1
        c.save()

    try:
        render_pdf(path)
    except Exception as exc:
        messagebox.showerror("Erro", f"Falha ao gerar PDF: {exc}")
        return
    try:
        os.startfile(path)
    except Exception:
        messagebox.showerror("Erro", "Nao foi possivel abrir o PDF do stock.")

