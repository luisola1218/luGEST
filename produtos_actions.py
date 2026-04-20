from module_context import configure_module, ensure_module

_CONFIGURED = False

def configure(main_globals):
    configure_module(globals(), main_globals)

def _ensure_configured():
    ensure_module(globals(), "produtos_actions")


def _prod_pdf_register_fonts():
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


def _prod_pdf_hex_to_rgb(value):
    txt = str(value or "").strip().lstrip("#")
    if len(txt) == 3:
        txt = "".join(ch * 2 for ch in txt)
    if len(txt) != 6:
        txt = "1F3C88"
    try:
        return tuple(int(txt[i : i + 2], 16) for i in (0, 2, 4))
    except Exception:
        return (31, 60, 136)


def _prod_pdf_rgb_to_hex(rgb):
    r, g, b = [max(0, min(255, int(v))) for v in tuple(rgb or (31, 60, 136))]
    return f"#{r:02X}{g:02X}{b:02X}"


def _prod_pdf_mix_hex(base_hex, target_hex, ratio):
    ratio = max(0.0, min(1.0, float(ratio)))
    base = _prod_pdf_hex_to_rgb(base_hex)
    target = _prod_pdf_hex_to_rgb(target_hex)
    return _prod_pdf_rgb_to_hex(
        tuple(round((base_v * (1.0 - ratio)) + (target_v * ratio)) for base_v, target_v in zip(base, target))
    )


def _prod_pdf_palette():
    from reportlab.lib import colors

    primary_hex = "#1F3C88"
    try:
        cfg = get_branding_config() if callable(globals().get("get_branding_config")) else {}
        primary_hex = str((cfg or {}).get("primary_color", "") or primary_hex).strip() or primary_hex
    except Exception:
        pass
    return {
        "primary": colors.HexColor(primary_hex),
        "primary_dark": colors.HexColor(_prod_pdf_mix_hex(primary_hex, "#000000", 0.22)),
        "primary_soft": colors.HexColor(_prod_pdf_mix_hex(primary_hex, "#FFFFFF", 0.82)),
        "primary_soft_2": colors.HexColor(_prod_pdf_mix_hex(primary_hex, "#FFFFFF", 0.90)),
        "surface_warm": colors.HexColor("#FCFCFD"),
        "line": colors.HexColor(_prod_pdf_mix_hex(primary_hex, "#D7DEE8", 0.76)),
        "line_strong": colors.HexColor(_prod_pdf_mix_hex(primary_hex, "#708090", 0.34)),
        "muted": colors.HexColor("#667085"),
        "ink": colors.HexColor(_prod_pdf_mix_hex(primary_hex, "#1A1A1A", 0.72)),
        "danger_fill": colors.HexColor("#FEECEC"),
        "danger_text": colors.HexColor("#B42318"),
    }


def _prod_pdf_clip_text(value, max_w, font_name, font_size):
    from reportlab.pdfbase import pdfmetrics

    txt = "" if value is None else str(value)
    if pdfmetrics.stringWidth(txt, font_name, font_size) <= max_w:
        return txt
    ellipsis = "..."
    while txt and pdfmetrics.stringWidth(txt + ellipsis, font_name, font_size) > max_w:
        txt = txt[:-1]
    return f"{txt}{ellipsis}" if txt else ""


def _prod_pdf_wrap_text(value, font_name, font_size, max_w, max_lines=None):
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


def _prod_pdf_fit_font_size(text, font_name, max_w, preferred_size, min_size):
    from reportlab.pdfbase import pdfmetrics

    size = float(preferred_size)
    raw = str(text or "")
    max_w = max(12.0, float(max_w))
    while size > float(min_size) and pdfmetrics.stringWidth(raw, font_name, size) > max_w:
        size -= 0.3
    return max(float(min_size), round(size, 2))


def _prod_pdf_metric_grid_layout(group_right, banner_top, banner_height, cols=3, rows=2, group_w=338, gap=6):
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


def _produto_from_material_record(material):
    if not isinstance(material, dict):
        return {}
    formato = str(material.get("formato") or detect_materia_formato(material) or "Chapa").strip() or "Chapa"
    comp = parse_float(material.get("comprimento", 0), 0)
    larg = parse_float(material.get("largura", 0), 0)
    esp = parse_float(material.get("espessura", 0), 0)
    metros = parse_float(material.get("metros", 0), 0)
    peso = parse_float(material.get("peso_unid", 0), 0)
    if formato == "Tubo":
        desc_dim = f"{fmt_num(metros)}m"
    else:
        desc_dim = f"{fmt_num(comp)}x{fmt_num(larg)}" if (comp > 0 and larg > 0) else "-"
    return {
        "descricao": f"{material.get('material', '')} {fmt_num(esp)}mm | {desc_dim}".strip(),
        "categoria": formato,
        "tipo": str(material.get("material", "") or "").strip() or formato,
        "comp": comp,
        "larg": larg,
        "esp": esp,
        "metros_unid": metros,
        "peso_unid": peso,
        "p_compra": parse_float(material.get("p_compra", 0), 0),
        "unid": "UN",
    }


def _dialog_escolher_materia_para_produto(self):
    _ensure_configured()
    mats = self.data.get("materiais", [])
    if not mats:
        messagebox.showinfo("Info", "Sem registos em Matéria-Prima.")
        return None
    use_custom = CUSTOM_TK_AVAILABLE and os.environ.get("USE_CUSTOM_PROD", "1") != "0"
    Dlg = ctk.CTkToplevel if use_custom else Toplevel
    FrameCls = ctk.CTkFrame if use_custom else ttk.Frame
    LabelCls = ctk.CTkLabel if use_custom else ttk.Label
    EntryCls = ctk.CTkEntry if use_custom else ttk.Entry
    ButtonCls = ctk.CTkButton if use_custom else ttk.Button

    dlg = Dlg(self.root)
    dlg.title("Selecionar da Matéria-Prima em Stock")
    try:
        dlg.geometry("1080x620")
        dlg.transient(self.root)
    except Exception:
        pass
    dlg.grab_set()

    filtro = StringVar()
    top = FrameCls(dlg, fg_color="#f7f8fb") if use_custom else FrameCls(dlg)
    top.pack(fill="x", padx=8, pady=8)
    LabelCls(top, text="Pesquisar").pack(side="left", padx=6)
    ent = EntryCls(top, textvariable=filtro, width=340 if use_custom else 42)
    ent.pack(side="left", padx=6)
    ButtonCls(top, text="Atualizar", command=lambda: refresh()).pack(side="left", padx=6)

    frame = FrameCls(dlg, fg_color="#ffffff") if use_custom else FrameCls(dlg)
    frame.pack(fill="both", expand=True, padx=8, pady=6)
    cols = ("id", "material", "esp", "dim", "metros", "peso", "compra", "stock", "lote", "local")
    tree = ttk.Treeview(frame, columns=cols, show="headings", height=14)
    headers = {
        "id": "ID",
        "material": "Material",
        "esp": "Esp.",
        "dim": "Dimensão",
        "metros": "Metros",
        "peso": "Peso/Unid",
        "compra": "Compra",
        "stock": "Stock",
        "lote": "Lote",
        "local": "Localização",
    }
    widths = {"id": 95, "material": 180, "esp": 70, "dim": 160, "metros": 80, "peso": 90, "compra": 90, "stock": 80, "lote": 120, "local": 160}
    for col in cols:
        tree.heading(col, text=headers[col])
        tree.column(col, width=widths[col], anchor=("w" if col in ("material", "lote", "local") else "center"))
    tree.pack(fill="both", expand=True, side="left")
    sb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")
    tree.tag_configure("even", background="#f8fbff")
    tree.tag_configure("odd", background="#eef4fb")

    mat_map = {m.get("id"): m for m in mats}

    def refresh():
        q = (filtro.get() or "").strip().lower()
        for iid in tree.get_children():
            tree.delete(iid)
        for idx, m in enumerate(mats):
            formato = str(m.get("formato") or detect_materia_formato(m) or "Chapa").strip() or "Chapa"
            comp = parse_float(m.get("comprimento", 0), 0)
            larg = parse_float(m.get("largura", 0), 0)
            metros = parse_float(m.get("metros", 0), 0)
            dim = f"{fmt_num(metros)}m" if formato == "Tubo" else (f"{fmt_num(comp)}x{fmt_num(larg)}" if (comp > 0 and larg > 0) else "-")
            vals = (
                m.get("id", ""),
                m.get("material", ""),
                fmt_num(m.get("espessura", "")),
                dim,
                fmt_num(metros),
                fmt_num(m.get("peso_unid", 0)),
                fmt_num(m.get("p_compra", 0)),
                fmt_num(m.get("quantidade", 0)),
                m.get("lote_fornecedor", ""),
                m.get("Localizacao", m.get("Localização", "")),
            )
            if q and not any(q in str(v).lower() for v in vals):
                continue
            tree.insert("", END, values=vals, tags=("odd" if idx % 2 else "even",))

    chosen = {}

    def on_ok():
        sel = tree.selection()
        if not sel:
            return
        mat_id = tree.item(sel[0], "values")[0]
        material = mat_map.get(mat_id)
        if material:
            chosen.update(_produto_from_material_record(material))
        dlg.destroy()

    bottom = FrameCls(dlg, fg_color="#f7f8fb") if use_custom else FrameCls(dlg)
    bottom.pack(fill="x", padx=8, pady=8)
    ButtonCls(bottom, text="Selecionar", command=on_ok).pack(side="right", padx=4)
    ButtonCls(bottom, text="Cancelar", command=dlg.destroy).pack(side="right", padx=4)

    refresh()
    ent.bind("<KeyRelease>", lambda _e: refresh())
    ent.bind("<Return>", lambda _e: refresh())
    tree.bind("<Double-1>", lambda _e: on_ok())
    tree.bind("<Return>", lambda _e: on_ok())
    dlg.wait_window()
    return chosen if chosen else None

def produto_dict_from_form(self):
    _ensure_configured()
    return {
        "codigo": self.prod_codigo.get().strip(),
        "descricao": self.prod_descricao.get().strip(),
        "categoria": self.prod_categoria.get().strip(),
        "subcat": self.prod_subcat.get().strip(),
        "tipo": self.prod_tipo.get().strip(),
        "dimensoes": self.prod_dimensoes.get().strip(),
        "metros": parse_float(self.prod_metros.get(), 0),
        "peso_unid": parse_float(self.prod_peso_unid.get(), 0),
        "fabricante": self.prod_fab.get().strip(),
        "modelo": self.prod_modelo.get().strip(),
        "unid": self.prod_unid.get().strip() or "UN",
        "qty": parse_float(self.prod_qty.get(), 0),
        "alerta": parse_float(self.prod_alerta.get(), 0),
        "p_compra": parse_float(self.prod_pcompra.get(), 0),
        "pvp1": parse_float(self.prod_pvp1.get(), 0),
        "pvp2": parse_float(self.prod_pvp2.get(), 0),
        "obs": self.prod_obs.get().strip(),
    }

def novo_produto(self):
    _ensure_configured()
    self.prod_sel_codigo = ""
    self.prod_codigo.set(peek_next_produto_numero(self.data))
    self.prod_descricao.set("")
    self.prod_categoria.set("")
    self.prod_subcat.set("")
    self.prod_tipo.set("")
    self.prod_dimensoes.set("")
    self.prod_metros.set("0")
    self.prod_peso_unid.set("0")
    self.prod_fab.set("")
    self.prod_modelo.set("")
    self.prod_unid.set("UN")
    self.prod_qty.set("0")
    self.prod_alerta.set("0")
    self.prod_pcompra.set("0")
    self.prod_pvp1.set("0")
    self.prod_pvp2.set("0")
    self.prod_obs.set("")
    self.prod_filter.set("")
    self.product_apply_category_mode()

def editar_produto_dialog(self):
    _ensure_configured()
    cod = self._get_selected_prod_codigo()
    if not cod:
        messagebox.showerror("Erro", "Selecione um produto para editar")
        return
    prod = next((x for x in self.data.get("produtos", []) if x.get("codigo") == cod), None)
    if not prod:
        messagebox.showerror("Erro", "Produto não encontrado")
        return
    self.dialog_novo_produto(prod)

def dialog_novo_produto(self, prod_init=None):
    _ensure_configured()
    """Janela modal com campos técnicos por categoria."""
    use_custom = CUSTOM_TK_AVAILABLE and os.environ.get("USE_CUSTOM_PROD", "1") != "0"
    try:
        dlg = ctk.CTkToplevel(self.root) if use_custom else Toplevel(self.root)
    except Exception:
        dlg = Toplevel(self.root)
        use_custom = False
    dlg.title("Editar Produto" if prod_init else "Novo Produto")
    dlg.geometry("940x700")
    dlg.minsize(900, 650)
    dlg.grab_set()
    dlg.transient(self.root)

    vars = {
        "codigo": StringVar(value=(prod_init.get("codigo", "") if prod_init else peek_next_produto_numero(self.data))),
        "descricao": StringVar(value=prod_init.get("descricao", "") if prod_init else ""),
        "categoria": StringVar(value=prod_init.get("categoria", "") if prod_init else ""),
        "subcat": StringVar(value=prod_init.get("subcat", "") if prod_init else ""),
        "tipo": StringVar(value=prod_init.get("tipo", "") if prod_init else ""),
        "fabricante": StringVar(value=prod_init.get("fabricante", "") if prod_init else ""),
        "modelo": StringVar(value=prod_init.get("modelo", "") if prod_init else ""),
        "unid": StringVar(value=prod_init.get("unid", "UN") if prod_init else "UN"),
        "qty": StringVar(value=str(prod_init.get("qty", 0)) if prod_init else "0"),
        "alerta": StringVar(value=str(prod_init.get("alerta", 0)) if prod_init else "0"),
        "p_compra": StringVar(value=str(prod_init.get("p_compra", 0)) if prod_init else "0"),
        "pvp1": StringVar(value=str(prod_init.get("pvp1", 0)) if prod_init else "0"),
        "pvp2": StringVar(value=str(prod_init.get("pvp2", 0)) if prod_init else "0"),
        "obs": StringVar(value=prod_init.get("obs", "") if prod_init else ""),
        "comp": StringVar(value=str(prod_init.get("comprimento", 0)) if prod_init else "0"),
        "larg": StringVar(value=str(prod_init.get("largura", 0)) if prod_init else "0"),
        "esp": StringVar(value=str(prod_init.get("espessura", 0)) if prod_init else "0"),
        "metros_unid": StringVar(value=str(prod_init.get("metros_unidade", prod_init.get("metros", 0))) if prod_init else "0"),
        "peso_unid": StringVar(value=str(prod_init.get("peso_unid", 0)) if prod_init else "0"),
        "preco_unid": StringVar(value="0"),
    }
    original_codigo = vars["codigo"].get()

    if use_custom:
        cont = ctk.CTkFrame(dlg, fg_color="#f5f8fd", corner_radius=10)
        cont.pack(fill="both", expand=True, padx=10, pady=10)
        title = ctk.CTkLabel(cont, text=("Editar Produto" if prod_init else "Novo Produto"), font=("Segoe UI", 22, "bold"), text_color="#114a7a")
        title.pack(anchor="w", padx=12, pady=(10, 4))
        frm = ctk.CTkFrame(cont, fg_color="transparent")
        frm.pack(fill="both", expand=True, padx=8, pady=6)
        LabelCls = ctk.CTkLabel
        EntryCls = ctk.CTkEntry
        BtnCls = ctk.CTkButton
    else:
        frm = ttk.Frame(dlg)
        frm.pack(fill="both", expand=True, padx=10, pady=10)
        LabelCls = ttk.Label
        EntryCls = ttk.Entry
        BtnCls = ttk.Button

    def add_row(r, lbl, key, values=None, width=28):
        LabelCls(frm, text=lbl).grid(row=r, column=0, sticky="w", padx=6, pady=4)
        if values:
            if use_custom:
                w = ctk.CTkComboBox(frm, variable=vars[key], values=list(values), width=max(170, int(width * 6)), state="normal")
            else:
                w = ttk.Combobox(frm, textvariable=vars[key], values=values, width=width, state="normal")
            w.grid(row=r, column=1, sticky="we", padx=6, pady=4)
            return w
        if use_custom:
            w = EntryCls(frm, textvariable=vars[key], width=max(170, int(width * 6)))
        else:
            w = EntryCls(frm, textvariable=vars[key], width=width)
        w.grid(row=r, column=1, sticky="we", padx=6, pady=4)
        return w

    cod_entry = add_row(0, "Código", "codigo", width=20)
    if prod_init:
        try:
            cod_entry.configure(state=("disabled" if use_custom else "readonly"))
        except Exception:
            pass
    add_row(1, "Descrição", "descricao", width=52)
    add_row(2, "Categoria", "categoria", values=PROD_CATEGORIAS, width=30)
    add_row(3, "Subcat.", "subcat", values=PROD_SUBCATS, width=30)
    add_row(4, "Tipo", "tipo", values=PROD_TIPOS, width=30)
    add_row(5, "Fabricante", "fabricante", width=30)
    add_row(6, "Modelo", "modelo", width=30)
    add_row(7, "Unid.", "unid", values=PROD_UNIDS, width=12)
    add_row(8, "Quantidade", "qty", width=18)
    add_row(9, "Alerta", "alerta", width=18)
    add_row(10, "Preço Compra", "p_compra", width=18)
    add_row(11, "PVP1", "pvp1", width=18)
    add_row(12, "PVP2", "pvp2", width=18)
    add_row(13, "Observações", "obs", width=52)

    if use_custom:
        tech = ctk.CTkFrame(frm, fg_color="#eef3fb", corner_radius=10, border_width=1, border_color="#e7cfd3")
        ctk.CTkLabel(tech, text="Dados Técnicos", font=("Segoe UI", 14, "bold"), text_color="#1d4b7a").grid(row=0, column=0, columnspan=2, sticky="w", padx=8, pady=(8, 4))
    else:
        tech = ttk.LabelFrame(frm, text="Dados Técnicos")
    tech.grid(row=0, column=2, rowspan=14, sticky="nsew", padx=12, pady=4)

    LblTech = ctk.CTkLabel if use_custom else ttk.Label
    EntTech = ctk.CTkEntry if use_custom else ttk.Entry
    LblTech(tech, text="Comprimento").grid(row=1 if use_custom else 0, column=0, sticky="w", padx=6, pady=3)
    e_comp = EntTech(tech, textvariable=vars["comp"], width=120 if use_custom else 12)
    e_comp.grid(row=1 if use_custom else 0, column=1, padx=6, pady=3)
    LblTech(tech, text="Largura").grid(row=2 if use_custom else 1, column=0, sticky="w", padx=6, pady=3)
    e_larg = EntTech(tech, textvariable=vars["larg"], width=120 if use_custom else 12)
    e_larg.grid(row=2 if use_custom else 1, column=1, padx=6, pady=3)
    LblTech(tech, text="Espessura").grid(row=3 if use_custom else 2, column=0, sticky="w", padx=6, pady=3)
    e_esp = EntTech(tech, textvariable=vars["esp"], width=120 if use_custom else 12)
    e_esp.grid(row=3 if use_custom else 2, column=1, padx=6, pady=3)
    LblTech(tech, text="Metros/Unid").grid(row=4 if use_custom else 3, column=0, sticky="w", padx=6, pady=3)
    e_metros = EntTech(tech, textvariable=vars["metros_unid"], width=120 if use_custom else 12)
    e_metros.grid(row=4 if use_custom else 3, column=1, padx=6, pady=3)
    LblTech(tech, text="Peso/Unid (kg)").grid(row=5 if use_custom else 4, column=0, sticky="w", padx=6, pady=3)
    e_peso = EntTech(tech, textvariable=vars["peso_unid"], width=120 if use_custom else 12)
    e_peso.grid(row=5 if use_custom else 4, column=1, padx=6, pady=3)
    LblTech(tech, text="Preço/Unid (auto)").grid(row=6 if use_custom else 5, column=0, sticky="w", padx=6, pady=3)
    e_preco = EntTech(tech, textvariable=vars["preco_unid"], width=120 if use_custom else 12)
    e_preco.grid(row=6 if use_custom else 5, column=1, padx=6, pady=3)
    try:
        e_preco.configure(state="disabled" if use_custom else "readonly")
    except Exception:
        pass

    if use_custom:
        pick_btn = ctk.CTkButton(
            tech,
            text="Selecionar da Matéria-Prima em Stock",
            width=260,
            fg_color="#f59e0b",
            hover_color="#d97706",
        )
        pick_btn.grid(row=7, column=0, columnspan=2, padx=8, pady=(10, 6), sticky="w")
    else:
        pick_btn = ttk.Button(tech, text="Selecionar da Matéria-Prima em Stock")
        pick_btn.grid(row=6, column=0, columnspan=2, padx=6, pady=(10, 6), sticky="w")

    def _set_state(w, st):
        try:
            w.configure(state=st)
        except Exception:
            pass

    def update_tech_visibility(*_):
        cat = vars["categoria"].get()
        tipo = vars["tipo"].get()
        modo = produto_modo_preco(cat, tipo)
        is_chapa = "chapa" in norm_text(tipo) or "chapa" in norm_text(cat)
        if is_chapa:
            for w in (e_comp, e_larg, e_esp, e_peso):
                _set_state(w, "normal")
            _set_state(e_metros, "disabled")
            vars["unid"].set("UN")
        elif modo == "metros":
            for w in (e_comp, e_larg, e_esp, e_peso):
                _set_state(w, "disabled")
            _set_state(e_metros, "normal")
            vars["unid"].set("UN")
        elif modo == "peso":
            for w in (e_comp, e_larg, e_esp, e_metros):
                _set_state(w, "disabled")
            _set_state(e_peso, "normal")
            vars["unid"].set("UN")
        else:
            for w in (e_comp, e_larg, e_esp, e_peso, e_metros):
                _set_state(w, "normal")

    def update_preco_unid(*_):
        cat = vars["categoria"].get()
        tipo = vars["tipo"].get()
        compra = parse_float(vars["p_compra"].get(), 0)
        modo = produto_modo_preco(cat, tipo)
        if modo == "peso":
            pu = parse_float(vars["peso_unid"].get(), 0) * compra
        elif modo == "metros":
            pu = parse_float(vars["metros_unid"].get(), 0) * compra
        else:
            pu = compra
        vars["preco_unid"].set(fmt_num(pu))

    def apply_material_template():
        chosen = self._dialog_escolher_materia_para_produto()
        if not chosen:
            return
        vars["descricao"].set(str(chosen.get("descricao", "") or ""))
        vars["categoria"].set(str(chosen.get("categoria", "") or ""))
        vars["tipo"].set(str(chosen.get("tipo", "") or ""))
        vars["comp"].set(str(chosen.get("comp", 0) or 0))
        vars["larg"].set(str(chosen.get("larg", 0) or 0))
        vars["esp"].set(str(chosen.get("esp", 0) or 0))
        vars["metros_unid"].set(str(chosen.get("metros_unid", 0) or 0))
        vars["peso_unid"].set(str(chosen.get("peso_unid", 0) or 0))
        vars["p_compra"].set(str(chosen.get("p_compra", 0) or 0))
        vars["unid"].set(str(chosen.get("unid", "UN") or "UN"))
        update_tech_visibility()
        update_preco_unid()

    vars["categoria"].trace_add("write", update_tech_visibility)
    vars["tipo"].trace_add("write", update_tech_visibility)
    for k in ("p_compra", "peso_unid", "metros_unid"):
        vars[k].trace_add("write", update_preco_unid)
    try:
        pick_btn.configure(command=apply_material_template)
    except Exception:
        pass
    if prod_init and not (vars["comp"].get() != "0" or vars["larg"].get() != "0" or vars["esp"].get() != "0"):
        try:
            dim = str(prod_init.get("dimensoes", ""))
            parts = [x.strip() for x in dim.split("x")]
            if len(parts) >= 3:
                vars["comp"].set(parts[0])
                vars["larg"].set(parts[1])
                vars["esp"].set(parts[2])
        except Exception:
            pass
    update_tech_visibility()
    update_preco_unid()

    def on_ok():
        try:
            cat = vars["categoria"].get().strip()
            tipo_txt = vars["tipo"].get().strip()
            comp = parse_float(vars["comp"].get(), 0)
            larg = parse_float(vars["larg"].get(), 0)
            esp = parse_float(vars["esp"].get(), 0)
            metros_unid = parse_float(vars["metros_unid"].get(), 0)
            dimensoes = ""
            if produto_modo_preco(cat, tipo_txt) == "peso" and ("chapa" in norm_text(tipo_txt) or "chapa" in norm_text(cat)):
                dimensoes = f"{fmt_num(comp)}x{fmt_num(larg)}x{fmt_num(esp)}"
            prod = {
                "codigo": vars["codigo"].get().strip(),
                "descricao": vars["descricao"].get().strip(),
                "categoria": cat,
                "subcat": vars["subcat"].get().strip(),
                "tipo": tipo_txt,
                "dimensoes": dimensoes,
                "comprimento": comp,
                "largura": larg,
                "espessura": esp,
                "metros_unidade": metros_unid,
                "metros": metros_unid,
                "peso_unid": parse_float(vars["peso_unid"].get(), 0),
                "fabricante": vars["fabricante"].get().strip(),
                "modelo": vars["modelo"].get().strip(),
                "unid": vars["unid"].get().strip() or "UN",
                "qty": parse_float(vars["qty"].get(), 0),
                "alerta": parse_float(vars["alerta"].get(), 0),
                "p_compra": parse_float(vars["p_compra"].get(), 0),
                "pvp1": parse_float(vars["pvp1"].get(), 0),
                "pvp2": parse_float(vars["pvp2"].get(), 0),
                "obs": vars["obs"].get().strip(),
            }
            if not prod["codigo"]:
                messagebox.showerror("Erro", "Código em falta")
                return
            lst = self.data.setdefault("produtos", [])
            existing = next((x for x in lst if x.get("codigo") == prod["codigo"]), None)
            if existing and prod["codigo"] != original_codigo:
                messagebox.showerror("Erro", "Já existe um produto com este código")
                return
            base = next((x for x in lst if x.get("codigo") == original_codigo), None)
            created_new = not (base or existing)
            old_qty = 0.0
            target = None
            if base:
                old_qty = parse_float(base.get("qty", 0), 0.0)
                base.update(prod)
                target = base
            elif existing:
                old_qty = parse_float(existing.get("qty", 0), 0.0)
                existing.update(prod)
                target = existing
            else:
                lst.append(prod)
                target = prod
            if target is not None:
                target["atualizado_em"] = now_iso()
                new_qty = parse_float(target.get("qty", 0), 0.0)
                operador_mov = str(self.user.get("username", "") or "Sistema")
                codigo_mov = str(target.get("codigo", "") or "")
                descricao_mov = str(target.get("descricao", "") or "")
                if created_new and new_qty > 1e-9:
                    add_produto_mov(
                        self.data,
                        tipo="ENTRADA_INICIAL",
                        operador=operador_mov,
                        codigo=codigo_mov,
                        descricao=descricao_mov,
                        qtd=new_qty,
                        antes=0.0,
                        depois=new_qty,
                        obs="Stock inicial no registo do produto",
                        origem="PRODUTOS",
                        ref_doc=codigo_mov,
                    )
                elif not created_new and abs(new_qty - old_qty) > 1e-9:
                    delta = new_qty - old_qty
                    add_produto_mov(
                        self.data,
                        tipo="AJUSTE_STOCK",
                        operador=operador_mov,
                        codigo=codigo_mov,
                        descricao=descricao_mov,
                        qtd=abs(delta),
                        antes=old_qty,
                        depois=new_qty,
                        obs=f"Ajuste manual no cadastro ({fmt_num(delta)})",
                        origem="PRODUTOS",
                        ref_doc=codigo_mov,
                    )
            ensure_produto_seq(self.data, prod["codigo"])
            self.prod_codigo.set(prod["codigo"])
            self.prod_descricao.set(prod["descricao"])
            self.prod_categoria.set(prod["categoria"])
            self.prod_subcat.set(prod["subcat"])
            self.prod_tipo.set(prod["tipo"])
            self.prod_dimensoes.set(prod.get("dimensoes", ""))
            self.prod_metros.set(str(prod.get("metros", 0)))
            self.prod_peso_unid.set(str(prod.get("peso_unid", 0)))
            self.prod_fab.set(prod["fabricante"])
            self.prod_modelo.set(prod["modelo"])
            self.prod_unid.set(prod["unid"])
            self.prod_qty.set(str(prod["qty"]))
            self.prod_alerta.set(str(prod["alerta"]))
            self.prod_pcompra.set(str(prod["p_compra"]))
            self.prod_pvp1.set(str(prod["pvp1"]))
            self.prod_pvp2.set(str(prod["pvp2"]))
            self.prod_obs.set(prod["obs"])
            save_data(self.data)
            self.refresh_produtos()
            dlg.destroy()
            messagebox.showinfo("OK", "Produto guardado")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao guardar produto: {e}")

    btns = ctk.CTkFrame(frm, fg_color="transparent") if use_custom else ttk.Frame(frm)
    btns.grid(row=14, column=0, columnspan=3, pady=12, sticky="ew")
    if use_custom:
        BtnCls(btns, text="Guardar", command=on_ok, width=140, fg_color="#a32035").pack(side="left", padx=6)
        BtnCls(btns, text="Cancelar", command=dlg.destroy, width=140, fg_color="#6b7280").pack(side="left", padx=6)
    else:
        BtnCls(btns, text="Guardar", command=on_ok).pack(side="left", padx=6)
        BtnCls(btns, text="Cancelar", command=dlg.destroy).pack(side="left", padx=6)

    frm.columnconfigure(1, weight=1)
    frm.columnconfigure(2, weight=1)
    dlg.wait_window()

def refresh_produtos(self):
    _ensure_configured()
    filtro = self.prod_filter.get().lower() if hasattr(self, "prod_filter") else ""
    if self.prod_use_full_custom and hasattr(self, "prod_list"):
        for w in self.prod_list.winfo_children():
            w.destroy()
        self._prod_btns = {}
    else:
        children = self.tbl_produtos.get_children()
        if children:
            self.tbl_produtos.delete(*children)
    selected = (self.prod_sel_codigo or self.prod_codigo.get().strip()).strip()
    for idx, p in enumerate(self.data.get("produtos", [])):
        dim = p.get("dimensoes", "")
        if not dim and (p.get("comprimento") or p.get("largura") or p.get("espessura")):
            dim = f"{fmt_num(p.get('comprimento', 0))}x{fmt_num(p.get('largura', 0))}x{fmt_num(p.get('espessura', 0))}"
        vals = (
            p.get("codigo"),
            p.get("descricao"),
            p.get("categoria"),
            p.get("subcat"),
            p.get("tipo"),
            dim,
            fmt_num(p.get("metros", 0)),
            fmt_num(p.get("peso_unid", 0)),
            p.get("unid"),
            fmt_num(p.get("qty", 0)),
            fmt_num(p.get("p_compra", 0)),
            fmt_num(produto_preco_unitario(p)),
            fmt_num(p.get("pvp1", 0)),
            fmt_num(p.get("pvp2", 0)),
            fmt_num(p.get("alerta", 0)),
        )
        if filtro and not any(filtro in str(v).lower() for v in vals):
            continue
        qty = parse_float(p.get("qty", 0), 0)
        alerta = parse_float(p.get("alerta", 0), 0)
        if qty <= 0 or (alerta > 0 and qty <= alerta):
            tag = "warn"
        else:
            tag = "odd" if idx % 2 else "even"
        if self.prod_use_full_custom and hasattr(self, "prod_list"):
            cod = p.get("codigo", "")
            is_sel = bool(selected and cod == selected)
            if tag == "warn":
                bg = "#ffe6bf"
                txt_color = "#8a5a00"
                hover = "#ffd9a3"
            elif is_sel:
                bg = "#fde2e4"
                txt_color = "#0f172a"
                hover = "#c9e1ff"
            else:
                bg = "#f8fbff" if idx % 2 == 0 else "#edf4ff"
                txt_color = "black"
                hover = "#f8dfe3"
            row = ctk.CTkFrame(self.prod_list, fg_color=bg, corner_radius=7)
            row.pack(fill="x", padx=2, pady=2)
            txt = (
                f"{cod}   |   {p.get('descricao','')}   |   {p.get('categoria','')} / {p.get('tipo','')}   |   "
                f"Qtd: {fmt_num(p.get('qty',0))}   |   Alerta: {fmt_num(p.get('alerta',0))}   |   "
                f"Preco/Unid: {fmt_num(produto_preco_unitario(p))} EUR"
            )
            btn = ctk.CTkButton(
                row,
                text=txt,
                anchor="w",
                fg_color=bg,
                hover_color=hover,
                text_color=txt_color,
                height=34,
                command=lambda c=cod: self._select_produto_custom(c),
            )
            btn.pack(fill="x", padx=4, pady=4)
            self._prod_btns[cod] = btn
        else:
            self.tbl_produtos.insert("", END, values=vals, tags=(tag,))

def _select_produto_custom(self, codigo):
    _ensure_configured()
    p = next((x for x in self.data.get("produtos", []) if x.get("codigo") == codigo), None)
    if not p:
        return
    self._set_prod_form_from_obj(p)
    self.refresh_produtos()

def on_produto_select(self, _event=None):
    _ensure_configured()
    cod = self._get_selected_prod_codigo()
    if not cod:
        return
    p = next((x for x in self.data.get("produtos", []) if x.get("codigo") == cod), None)
    self._set_prod_form_from_obj(p)

def guardar_produto(self):
    _ensure_configured()
    prod = self.produto_dict_from_form()
    if not prod["codigo"]:
        messagebox.showerror("Erro", "Código em falta")
        return
    lst = self.data.setdefault("produtos", [])
    existing = next((x for x in lst if x.get("codigo") == prod["codigo"]), None)
    old_qty = 0.0
    if existing:
        old_qty = parse_float(existing.get("qty", 0), 0.0)
        existing.update(prod)
        target = existing
    else:
        lst.append(prod)
        target = prod
    target["atualizado_em"] = now_iso()
    new_qty = parse_float(target.get("qty", 0), 0.0)
    operador_mov = str(self.user.get("username", "") or "Sistema")
    codigo_mov = str(target.get("codigo", "") or "")
    descricao_mov = str(target.get("descricao", "") or "")
    if not existing and new_qty > 1e-9:
        add_produto_mov(
            self.data,
            tipo="ENTRADA_INICIAL",
            operador=operador_mov,
            codigo=codigo_mov,
            descricao=descricao_mov,
            qtd=new_qty,
            antes=0.0,
            depois=new_qty,
            obs="Stock inicial no registo do produto",
            origem="PRODUTOS",
            ref_doc=codigo_mov,
        )
    elif existing and abs(new_qty - old_qty) > 1e-9:
        delta = new_qty - old_qty
        add_produto_mov(
            self.data,
            tipo="AJUSTE_STOCK",
            operador=operador_mov,
            codigo=codigo_mov,
            descricao=descricao_mov,
            qtd=abs(delta),
            antes=old_qty,
            depois=new_qty,
            obs=f"Ajuste manual no cadastro ({fmt_num(delta)})",
            origem="PRODUTOS",
            ref_doc=codigo_mov,
        )
    ensure_produto_seq(self.data, prod["codigo"])
    save_data(self.data)
    self.refresh_produtos()
    self.prod_sel_codigo = prod["codigo"]
    if not self.prod_use_full_custom:
        # selecionar o recém-guardado
        for item in self.tbl_produtos.get_children():
            if self.tbl_produtos.item(item, "values")[0] == prod["codigo"]:
                self.tbl_produtos.selection_set(item)
                self.tbl_produtos.see(item)
                break
    self.on_produto_select()
    messagebox.showinfo("OK", "Produto guardado")

def remover_produto(self):
    _ensure_configured()
    cod = self._get_selected_prod_codigo()
    if not cod:
        return
    self.data["produtos"] = [p for p in self.data.get("produtos", []) if p.get("codigo") != cod]
    save_data(self.data)
    self.prod_sel_codigo = ""
    self.refresh_produtos()
    self.novo_produto()

def render_produtos_stock_pdf(self, path):
    _ensure_configured()
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.pdfgen import canvas as pdf_canvas

    fonts = _prod_pdf_register_fonts()
    palette = _prod_pdf_palette()
    w, h = landscape(A4)
    c = pdf_canvas.Canvas(path, pagesize=landscape(A4))
    margin = 22
    content_w = w - (2 * margin)
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
        return h - top_y

    def fmt_money(value):
        return f"{parse_float(value, 0):.2f} EUR"

    def draw_logo(x, top_y, box_w=124, box_h=52):
        drawer = globals().get("draw_pdf_logo_box")
        if callable(drawer):
            try:
                drawer(c, h, x, top_y, box_size=box_w, box_h=box_h, padding=4, draw_border=False)
                return
            except Exception:
                pass

    def draw_title_block(left_x, right_x, title, subtitle):
        area_w = max(40.0, float(right_x) - float(left_x))
        center_x = float(left_x) + (area_w / 2.0)
        title_size = _prod_pdf_fit_font_size(title, fonts["bold"], area_w, 22.0, 16.0)
        subtitle_size = _prod_pdf_fit_font_size(subtitle, fonts["regular"], area_w, 10.0, 8.0)
        c.setFillColor(colors.white)
        c.setFont(fonts["bold"], title_size)
        c.drawCentredString(center_x, yinv(46), ntxt(_prod_pdf_clip_text(title, area_w, fonts["bold"], title_size)))
        c.setFont(fonts["regular"], subtitle_size)
        c.drawCentredString(center_x, yinv(63), ntxt(_prod_pdf_clip_text(subtitle, area_w, fonts["regular"], subtitle_size)))

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
            wrapped = _prod_pdf_wrap_text(line, line_font, line_size, box_w - 16, max_lines=2 if idx_line == 0 else 1)
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
        value_size = _prod_pdf_fit_font_size(value, fonts["bold"], box_w - 14, 10.2, 8.2)
        c.setFont(fonts["bold"], value_size)
        c.setFillColor(palette["primary_dark"])
        c.drawString(x + 7, yinv(top_y + 24), ntxt(_prod_pdf_clip_text(value, box_w - 14, fonts["bold"], value_size)))

    cols = [
        ("Codigo", 76, "w"),
        ("Descricao", 224, "w"),
        ("Categoria", 104, "w"),
        ("Subcat.", 90, "w"),
        ("Un.", 34, "center"),
        ("Qtd.", 50, "e"),
        ("Alerta", 52, "e"),
        ("Preco/Un.", 76, "e"),
        ("Valor", 86, "e"),
    ]
    table_w = sum(cw for _, cw, _ in cols)
    x0 = margin

    produtos = list(self.data.get("produtos", []))
    produtos = sorted(produtos, key=lambda p: (str(p.get("categoria", "")), str(p.get("descricao", ""))))
    total_qty = 0.0
    total_val = 0.0
    out_of_stock = 0
    needs_restock = 0
    category_counts = {}
    for p in produtos:
        cat = str(p.get("categoria", "") or "Sem categoria").strip() or "Sem categoria"
        category_counts[cat] = int(category_counts.get(cat, 0)) + 1
        qtd = parse_float(p.get("qty", 0), 0)
        alerta = parse_float(p.get("alerta", 0), 0)
        total_qty += qtd
        total_val += qtd * produto_preco_unitario(p)
        if qtd <= 0:
            out_of_stock += 1
        if qtd <= 0 or (alerta > 0 and qtd <= alerta):
            needs_restock += 1
    top_categories = sorted(category_counts.items(), key=lambda item: (-item[1], item[0]))[:4]

    table_rows = []
    last_cat = None
    zebra_idx = 0
    for p in produtos:
        cat = str(p.get("categoria", "") or "Sem categoria")
        if cat != last_cat:
            table_rows.append({"_group": cat, "_count": category_counts.get(cat, 0)})
            last_cat = cat
        pr = dict(p)
        pr["_zebra"] = zebra_idx
        zebra_idx += 1
        table_rows.append(pr)

    def draw_table_header(y_top):
        c.saveState()
        c.setFillColor(palette["primary"])
        c.roundRect(x0, yinv(y_top + table_header_h), table_w, table_header_h, 7, stroke=0, fill=1)
        c.restoreState()
        c.setFillColor(colors.white)
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
        c.setFillColor(colors.black)

    def draw_header(page_no, total_pages, first_page):
        if first_page:
            metric_grid = _prod_pdf_metric_grid_layout(w - margin, 20, 82, cols=3, rows=2, group_w=338, gap=6)
            c.saveState()
            c.setFillColor(palette["primary"])
            c.roundRect(margin, yinv(102), content_w, 82, 12, stroke=0, fill=1)
            c.restoreState()
            draw_logo(margin + 12, 30)
            draw_title_block(margin + 152, metric_grid["group_x"] - 14, "Stock de Produtos", "Relatorio consolidado de artigos e reposicao")
            info_chip(metric_grid["cols"][0], metric_grid["rows"][0], metric_grid["chip_w"], "Emitido", datetime.now().strftime("%d/%m/%Y"))
            info_chip(metric_grid["cols"][1], metric_grid["rows"][0], metric_grid["chip_w"], "Pagina", f"{page_no}/{total_pages}")
            info_chip(metric_grid["cols"][2], metric_grid["rows"][0], metric_grid["chip_w"], "Categorias", str(len(category_counts)))
            info_chip(metric_grid["cols"][0], metric_grid["rows"][1], metric_grid["chip_w"], "Refs", str(len(produtos)))
            info_chip(metric_grid["cols"][1], metric_grid["rows"][1], metric_grid["chip_w"], "Reposicao", str(needs_restock))
            info_chip(metric_grid["cols"][2], metric_grid["rows"][1], metric_grid["chip_w"], "Valor", fmt_money(total_val))
            card(
                margin,
                126,
                430,
                62,
                "Resumo rapido",
                [
                    f"Produtos ativos: {len(produtos)} | Categorias: {len(category_counts)}",
                    f"Qtd total em stock: {fmt_num(total_qty)} | Sem stock: {out_of_stock}",
                    f"Artigos a repor: {needs_restock} | Valor global: {fmt_money(total_val)}",
                ],
                accent=True,
            )
            category_lines = [f"{name}: {count} refs" for name, count in top_categories] or ["Sem produtos registados."]
            card(
                margin + 442,
                126,
                content_w - 442,
                62,
                "Categorias principais",
                category_lines,
                font_size=8.0,
            )
            return draw_table_header(table_first_top)

        c.saveState()
        c.setFillColor(palette["primary_soft"])
        c.roundRect(margin, yinv(68), content_w, 48, 12, stroke=0, fill=1)
        c.restoreState()
        c.setFillColor(palette["primary_dark"])
        c.setFont(fonts["bold"], 15)
        c.drawString(margin + 12, yinv(40), ntxt("Stock de Produtos"))
        c.setFont(fonts["regular"], 8.6)
        c.drawString(margin + 12, yinv(54), ntxt(f"Reposicao pendente: {needs_restock} | Valor: {fmt_money(total_val)}"))
        metric_grid = _prod_pdf_metric_grid_layout(w - margin, 14, 48, cols=2, rows=1, group_w=196, gap=6)
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
            "Conferencia de stock",
            [
                "Data: ____________________",
                "Responsavel: ________________________________",
                "Observacoes: Validar divergencias antes de fechar a contagem.",
            ],
            font_size=7.9,
        )
        card(
            margin + 486,
            486,
            content_w - 486,
            72,
            "Resumo financeiro",
            [
                f"Valor total: {fmt_money(total_val)}",
                f"Sem stock: {out_of_stock} | Reposicao: {needs_restock}",
                f"Pagina: {page_no}/{total_pages}",
            ],
            accent=True,
            font_size=8.0,
        )
        c.setFillColor(palette["muted"])
        c.setFont(fonts["regular"], 7.0)
        if footer_lines:
            c.drawString(margin, yinv(570), ntxt(" | ".join(footer_lines[:2])))
        c.drawRightString(w - margin, yinv(570), ntxt(f"luGEST | {datetime.now().strftime('%Y-%m-%d %H:%M')}"))

    total_pages = 1
    remaining_preview = len(table_rows)
    while True:
        first = total_pages == 1
        table_top = table_first_top if first else table_next_top
        usable_last = footer_top - (table_top + table_header_h + 6)
        usable_non = (h - margin - 28) - (table_top + table_header_h + 6)
        fit_last = max(1, int(usable_last // row_h))
        fit_non = max(1, int(usable_non // row_h))
        if remaining_preview <= fit_last:
            break
        remaining_preview -= fit_non
        total_pages += 1

    idx = 0
    page = 1
    while idx < len(table_rows) or idx == 0:
        first = page == 1
        draw_header(page, total_pages, first)
        table_top = table_first_top if first else table_next_top

        usable_last = footer_top - (table_top + table_header_h + 6)
        usable_non = (h - margin - 28) - (table_top + table_header_h + 6)
        fit_last = max(1, int(usable_last // row_h))
        fit_non = max(1, int(usable_non // row_h))

        remaining = len(table_rows) - idx
        if remaining <= fit_last:
            is_last = True
            count = remaining
        else:
            is_last = False
            count = min(fit_non, remaining)

        y_top = table_top + table_header_h
        for local_i in range(count):
            p = table_rows[idx + local_i]
            y_row = y_top + (local_i * row_h)
            if p.get("_group") is not None:
                c.saveState()
                c.setFillColor(palette["primary_soft"])
                c.setStrokeColor(palette["line"])
                c.roundRect(x0, yinv(y_row + row_h), table_w, row_h, 5, stroke=1, fill=1)
                c.restoreState()
                c.setFillColor(palette["primary_dark"])
                c.setFont(fonts["bold"], 8.4)
                c.drawString(x0 + 6, yinv(y_row + 11.5), ntxt(f"Categoria: {p.get('_group')} ({p.get('_count', 0)} refs)"))
                c.setFillColor(colors.black)
                continue
            qtd = parse_float(p.get("qty", 0), 0)
            alerta = parse_float(p.get("alerta", 0), 0)
            preco_un = produto_preco_unitario(p)
            val_stock = qtd * preco_un

            if qtd <= 0 or (alerta > 0 and qtd <= alerta):
                fill = palette["danger_fill"]
                text_color = palette["danger_text"]
            else:
                fill = palette["surface_warm"] if (int(p.get("_zebra", 0)) % 2 == 0) else colors.white
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
                p.get("codigo", ""),
                p.get("descricao", ""),
                p.get("categoria", ""),
                p.get("subcat", ""),
                p.get("unid", "UN"),
                fmt_num(qtd),
                fmt_num(alerta),
                fmt_num(preco_un),
                fmt_num(val_stock),
            ]
            xx = x0
            for (_name, cw, align), val in zip(cols, vals):
                txt = _prod_pdf_clip_text(ntxt(val), cw - 12, fonts["regular"], 8.0)
                if align == "e":
                    c.drawRightString(xx + cw - 6, yinv(y_row + 11.5), txt)
                elif align == "center":
                    c.drawCentredString(xx + cw / 2, yinv(y_row + 11.5), txt)
                else:
                    c.drawString(xx + 6, yinv(y_row + 11.5), txt)
                xx += cw

        idx += count
        if is_last:
            draw_footer(page, total_pages)
            break
        c.setFont(fonts["regular"], 7.4)
        c.setFillColor(palette["muted"])
        c.drawRightString(w - margin, yinv(h - margin - 4), ntxt("Continua na proxima pagina..."))
        c.showPage()
        page += 1

    c.save()

def preview_produtos_stock_pdf(self):
    _ensure_configured()
    prev_dir = os.path.join(BASE_DIR, "previews")
    try:
        os.makedirs(prev_dir, exist_ok=True)
    except Exception:
        prev_dir = tempfile.gettempdir()
    path = os.path.join(prev_dir, "lugest_stock_produtos.pdf")
    try:
        self.render_produtos_stock_pdf(path)
        opened = self._open_pdf_default(path)
        if not opened:
            messagebox.showwarning("Aviso", f"PDF de stock gerado mas não abriu automaticamente.\nCaminho:\n{path}")
    except Exception as ex:
        messagebox.showerror("Erro", f"Falha ao gerar PDF de stock: {ex}")

def save_produtos_stock_pdf(self):
    _ensure_configured()
    path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF", "*.pdf")],
        initialfile=f"stock_produtos_{datetime.now().strftime('%Y%m%d')}.pdf",
    )
    if not path:
        return
    try:
        self.render_produtos_stock_pdf(path)
        messagebox.showinfo("OK", "PDF de stock guardado")
    except Exception as ex:
        messagebox.showerror("Erro", f"Falha ao guardar PDF de stock: {ex}")

def print_produtos_stock_pdf(self):
    _ensure_configured()
    path = os.path.join(tempfile.gettempdir(), "lugest_stock_produtos.pdf")
    try:
        self.render_produtos_stock_pdf(path)
        try:
            os.startfile(path, "print")
        except Exception:
            self._open_pdf_default(path)
    except Exception as ex:
        messagebox.showerror("Erro", f"Falha ao imprimir stock: {ex}")

def get_selected_produto_obj(self):
    _ensure_configured()
    cod = self._get_selected_prod_codigo()
    if not cod:
        return None
    return next((p for p in self.data.get("produtos", []) if p.get("codigo") == cod), None)

def produto_dar_baixa_dialog(self):
    _ensure_configured()
    p = self.get_selected_produto_obj()
    if not p:
        messagebox.showerror("Erro", "Selecione um produto para dar baixa")
        return

    use_custom = CUSTOM_TK_AVAILABLE and os.environ.get("USE_CUSTOM_PROD", "1") != "0"
    Dlg = ctk.CTkToplevel if use_custom else Toplevel
    Frm = ctk.CTkFrame if use_custom else ttk.Frame
    Lbl = ctk.CTkLabel if use_custom else ttk.Label
    Ent = ctk.CTkEntry if use_custom else ttk.Entry
    Btn = ctk.CTkButton if use_custom else ttk.Button

    dlg = Dlg(self.root)
    dlg.title("Dar Baixa de Stock")
    dlg.geometry("640x280")
    dlg.grab_set()

    op = StringVar(value=(self.data.get("operadores", [""])[0] if self.data.get("operadores") else ""))
    qtd = StringVar(value="1")
    obs = StringVar()

    top = Frm(dlg, fg_color="#f7f8fb") if use_custom else Frm(dlg)
    top.pack(fill="both", expand=True, padx=10, pady=10)

    Lbl(top, text=f"Produto: {p.get('codigo','')} - {p.get('descricao','')}", font=("Arial", 13, "bold")).grid(row=0, column=0, columnspan=4, sticky="w", pady=(4, 8))
    Lbl(top, text=f"Stock atual: {fmt_num(p.get('qty', 0))} {p.get('unid', 'UN')}").grid(row=1, column=0, columnspan=4, sticky="w", pady=(0, 10))

    Lbl(top, text="Operador").grid(row=2, column=0, sticky="w", padx=4, pady=4)
    cb_op = ttk.Combobox(top, textvariable=op, values=self.data.get("operadores", []), width=28, state="normal")
    cb_op.grid(row=2, column=1, sticky="w", padx=4, pady=4)
    Lbl(top, text="Quantidade").grid(row=2, column=2, sticky="w", padx=4, pady=4)
    Ent(top, textvariable=qtd, width=100 if use_custom else 14).grid(row=2, column=3, sticky="w", padx=4, pady=4)

    Lbl(top, text="Observação").grid(row=3, column=0, sticky="w", padx=4, pady=4)
    Ent(top, textvariable=obs, width=360 if use_custom else 52).grid(row=3, column=1, columnspan=3, sticky="we", padx=4, pady=4)

    def registar():
        operador = op.get().strip()
        if not operador:
            messagebox.showerror("Erro", "Selecione o operador")
            return
        q = parse_float(qtd.get(), 0)
        if q <= 0:
            messagebox.showerror("Erro", "Quantidade deve ser maior que zero")
            return
        before = parse_float(p.get("qty", 0), 0)
        if q > before:
            messagebox.showerror("Erro", "Quantidade superior ao stock disponível")
            return
        after = before - q
        p["qty"] = after
        add_produto_mov(
            self.data,
            tipo="BAIXA",
            operador=operador,
            codigo=str(p.get("codigo", "") or ""),
            descricao=str(p.get("descricao", "") or ""),
            qtd=q,
            antes=before,
            depois=after,
            obs=obs.get().strip(),
            origem="PRODUTOS",
            ref_doc=str(p.get("codigo", "") or ""),
        )
        save_data(self.data)
        self.refresh_produtos()
        self.on_produto_select()
        messagebox.showinfo("OK", "Baixa registada com sucesso")
        dlg.destroy()

    btns = Frm(dlg, fg_color="#f7f8fb") if use_custom else Frm(dlg)
    btns.pack(fill="x", padx=10, pady=8)
    Btn(btns, text="Registar Baixa", command=registar).pack(side="left", padx=4)
    Btn(btns, text="Movimentos Operador", command=lambda: self.produto_mov_operador_dialog(default_operador=op.get().strip())).pack(side="left", padx=4)
    Btn(btns, text="Fechar", command=dlg.destroy).pack(side="right", padx=4)

def _update_produto_preco_from_unit(self, produto_codigo, preco_unit):
    _ensure_configured()
    p = next((x for x in self.data.get("produtos", []) if x.get("codigo") == produto_codigo), None)
    if not p:
        return False
    preco_linha = parse_float(preco_unit, 0)
    cat = p.get("categoria", "")
    tipo = p.get("tipo", "")
    old = parse_float(p.get("p_compra", 0), 0)
    new_val = old
    modo = produto_modo_preco(cat, tipo)
    if modo == "peso":
        peso = parse_float(p.get("peso_unid", 0), 0)
        if peso > 0:
            new_val = round(preco_linha / peso, 6)
    elif modo == "metros":
        metros_un = parse_float(p.get("metros_unidade", p.get("metros", 0)), 0)
        if metros_un > 0:
            new_val = round(preco_linha / metros_un, 6)
    else:
        new_val = preco_linha
    if abs(new_val - old) > 1e-9:
        p["p_compra"] = new_val
        return True
    return False

def _sync_ne_linhas_with_produtos(self, ne):
    _ensure_configured()
    changed = False
    prod_map = {p.get("codigo"): p for p in self.data.get("produtos", [])}
    for l in ne.get("linhas", []):
        if origem_is_materia(l.get("origem", "Produto")):
            continue
        p = prod_map.get(l.get("ref"))
        if not p:
            continue
        novo_preco = round(produto_preco_unitario(p), 6)
        atual_preco = parse_float(l.get("preco", 0), 0)
        if abs(novo_preco - atual_preco) > 1e-9:
            l["preco"] = novo_preco
            q = parse_float(l.get("qtd", 0), 0)
            l["total"] = round(q * novo_preco, 6)
            changed = True
        desc = p.get("descricao", "")
        if desc and l.get("descricao", "") != desc:
            l["descricao"] = desc
            changed = True
        un = p.get("unid", "")
        if un and l.get("unid", "") != un:
            l["unid"] = un
            changed = True
    if changed:
        ne["total"] = sum(parse_float(x.get("total", 0), 0) for x in ne.get("linhas", []))
    return changed

def _dialog_escolher_produto(self, for_ne=False):
    _ensure_configured()
    """Pequena janela para escolher produto da lista."""
    prods_all = self.data.get("produtos", [])
    if for_ne:
        # Em Nota de Encomenda, itens metalúrgicos devem vir da Matéria-Prima.
        prods = [p for p in prods_all if not is_metal_categoria(p.get("categoria", ""))]
    else:
        prods = list(prods_all)
    if not prods:
        msg = "Sem produtos elegíveis."
        if for_ne:
            msg = "Sem produtos não-metalúrgicos. Para chapas/tubos/perfis/vigas use 'Matéria-Prima' na linha."
        messagebox.showinfo("Info", msg)
        return None
    use_custom = CUSTOM_TK_AVAILABLE and os.environ.get("USE_CUSTOM_NE", "1") != "0"
    dlg = ctk.CTkToplevel(self.root) if use_custom else Toplevel(self.root)
    dlg.title("Escolher Produto")
    dlg.geometry("980x560")
    dlg.grab_set()
    filtro = StringVar()
    LabelCls = ctk.CTkLabel if use_custom else ttk.Label
    EntryCls = ctk.CTkEntry if use_custom else ttk.Entry
    ButtonCls = ctk.CTkButton if use_custom else ttk.Button
    FrameCls = ctk.CTkFrame if use_custom else ttk.Frame

    top = FrameCls(dlg, fg_color="#f7f8fb") if use_custom else FrameCls(dlg)
    top.pack(fill="x", padx=8, pady=8)
    LabelCls(top, text="Pesquisar").pack(side="left", padx=6)
    ent = EntryCls(top, textvariable=filtro, width=340 if use_custom else 42)
    ent.pack(side="left", padx=6)
    ButtonCls(top, text="Atualizar", command=lambda: refresh()).pack(side="left", padx=6)

    frame = FrameCls(dlg, fg_color="#ffffff") if use_custom else FrameCls(dlg)
    frame.pack(fill="both", expand=True, padx=8, pady=6)
    cols = ("codigo", "descricao", "categoria", "qtd", "alerta", "unid", "p_compra", "pvp1")
    pick_style = ""
    if use_custom:
        style = ttk.Style()
        style.configure("NEPick.Treeview", font=("Segoe UI", 10), rowheight=26)
        style.configure("NEPick.Treeview.Heading", font=("Segoe UI", 10, "bold"))
        pick_style = "NEPick.Treeview"
    tree = ttk.Treeview(frame, columns=cols, show="headings", height=12, style=pick_style)
    headings = {
        "codigo": "Código",
        "descricao": "Descrição",
        "categoria": "Categoria",
        "qtd": "Stock",
        "alerta": "Alerta",
        "unid": "Un.",
        "p_compra": "Compra (€)",
        "pvp1": "PVP1 (€)",
    }
    for c in cols:
        tree.heading(c, text=headings[c])
        if c == "descricao":
            tree.column(c, width=260, anchor="w")
        elif c == "categoria":
            tree.column(c, width=120, anchor="w")
        elif c in ("qtd", "alerta"):
            tree.column(c, width=80, anchor="e")
        elif c in ("p_compra", "pvp1"):
            tree.column(c, width=95, anchor="e")
        else:
            tree.column(c, width=95, anchor="w")
    tree.pack(fill="both", expand=True, side="left")
    sb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    sbh = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    tree.configure(yscroll=sb.set)
    sb.pack(side="right", fill="y")
    tree.configure(xscroll=sbh.set)
    sbh.pack(side="bottom", fill="x")
    tree.tag_configure("even", background="#f8fbff")
    tree.tag_configure("odd", background="#eef4fb")
    tree.tag_configure("warn", background="#ffe6bf", foreground="#8a5a00")

    def refresh():
        q = filtro.get().lower()
        for i in tree.get_children():
            tree.delete(i)
        for idx, p in enumerate(prods):
            vals = (
                p.get("codigo", ""),
                p.get("descricao", ""),
                p.get("categoria", ""),
                fmt_num(p.get("qty", 0)),
                fmt_num(p.get("alerta", 0)),
                p.get("unid", "UN"),
                fmt_num(p.get("p_compra", 0)),
                fmt_num(p.get("pvp1", 0)),
            )
            if q and not any(q in str(v).lower() for v in vals):
                continue
            qtd = parse_float(p.get("qty", 0), 0)
            alerta = parse_float(p.get("alerta", 0), 0)
            if qtd <= 0 or (alerta > 0 and qtd <= alerta):
                tag = "warn"
            else:
                tag = "odd" if idx % 2 else "even"
            tree.insert("", END, values=vals, tags=(tag,))

    refresh()
    ent.bind("<KeyRelease>", lambda _e: refresh())
    ent.bind("<Return>", lambda _e: refresh())

    chosen = {}

    def on_ok():
        sel = tree.selection()
        if not sel:
            return
        cod = tree.item(sel[0], "values")[0]
        prod = next((x for x in prods if x.get("codigo") == cod), None)
        if prod:
            chosen.update(prod)
        dlg.destroy()

    def on_cancel():
        dlg.destroy()

    bottom = FrameCls(dlg, fg_color="#f7f8fb") if use_custom else FrameCls(dlg)
    bottom.pack(fill="x", padx=8, pady=8)
    ButtonCls(bottom, text="Selecionar", command=on_ok).pack(side="right", padx=4)
    ButtonCls(bottom, text="Cancelar", command=on_cancel).pack(side="right", padx=4)

    tree.bind("<Double-1>", lambda _e: on_ok())
    tree.bind("<Return>", lambda _e: on_ok())
    dlg.wait_window()
    return chosen if chosen else None



