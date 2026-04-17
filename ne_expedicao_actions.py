from module_context import configure_module, ensure_module

_CONFIGURED = False

def configure(main_globals):
    configure_module(globals(), main_globals)

def _ensure_configured():
    ensure_module(globals(), "ne_expedicao_actions")


def _pdf_register_fonts():
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    regular = "Helvetica"
    bold = "Helvetica-Bold"
    candidates = [
        ("Arial", r"C:\Windows\Fonts\arial.ttf"),
        ("Arial-Bold", r"C:\Windows\Fonts\arialbd.ttf"),
    ]
    for name, path in candidates:
        try:
            if path and os.path.exists(path):
                pdfmetrics.registerFont(TTFont(name, path))
        except Exception:
            pass
    try:
        if "Arial" in pdfmetrics.getRegisteredFontNames():
            regular = "Arial"
        if "Arial-Bold" in pdfmetrics.getRegisteredFontNames():
            bold = "Arial-Bold"
    except Exception:
        pass
    return {"regular": regular, "bold": bold}


def _pdf_hex_to_rgb(value):
    txt = str(value or "").strip().lstrip("#")
    if len(txt) == 3:
        txt = "".join(ch * 2 for ch in txt)
    if len(txt) != 6:
        txt = "1F3C88"
    try:
        return tuple(int(txt[i : i + 2], 16) for i in (0, 2, 4))
    except Exception:
        return (31, 60, 136)


def _pdf_rgb_to_hex(rgb):
    r, g, b = [max(0, min(255, int(v))) for v in rgb]
    return f"#{r:02X}{g:02X}{b:02X}"


def _pdf_mix_hex(base_hex, target_hex, ratio):
    ratio = max(0.0, min(1.0, float(ratio)))
    base = _pdf_hex_to_rgb(base_hex)
    target = _pdf_hex_to_rgb(target_hex)
    out = []
    for base_v, target_v in zip(base, target):
        out.append(round((base_v * (1.0 - ratio)) + (target_v * ratio)))
    return _pdf_rgb_to_hex(tuple(out))


def _pdf_brand_palette():
    from reportlab.lib import colors

    primary_hex = str(get_branding_config().get("primary_color", "") or "#1F3C88").strip() or "#1F3C88"
    line_hex = _pdf_mix_hex(primary_hex, "#D7DEE8", 0.76)
    return {
        "primary": colors.HexColor(primary_hex),
        "primary_hex": primary_hex,
        "primary_dark": colors.HexColor(_pdf_mix_hex(primary_hex, "#000000", 0.22)),
        "primary_soft": colors.HexColor(_pdf_mix_hex(primary_hex, "#FFFFFF", 0.80)),
        "primary_soft_2": colors.HexColor(_pdf_mix_hex(primary_hex, "#FFFFFF", 0.90)),
        "primary_mid": colors.HexColor(_pdf_mix_hex(primary_hex, "#FFFFFF", 0.55)),
        "ink": colors.HexColor(_pdf_mix_hex(primary_hex, "#1A1A1A", 0.72)),
        "muted": colors.HexColor("#667085"),
        "line": colors.HexColor(line_hex),
        "line_strong": colors.HexColor(_pdf_mix_hex(primary_hex, "#708090", 0.36)),
        "surface": colors.white,
        "surface_alt": colors.HexColor("#F7F9FC"),
        "surface_warm": colors.HexColor("#FCFCFD"),
        "success": colors.HexColor("#107569"),
        "danger": colors.HexColor("#B42318"),
    }


def _pdf_clip_text(value, max_w, font_name, font_size):
    from reportlab.pdfbase import pdfmetrics

    txt = "" if value is None else str(value)
    if pdfmetrics.stringWidth(txt, font_name, font_size) <= max_w:
        return txt
    ellipsis = "..."
    while txt and pdfmetrics.stringWidth(txt + ellipsis, font_name, font_size) > max_w:
        txt = txt[:-1]
    return f"{txt}{ellipsis}" if txt else ""


def _pdf_wrap_text(value, font_name, font_size, max_w, max_lines=None):
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


def _pdf_metric_grid_layout(group_right, banner_top, banner_height, group_w=230, gap=6):
    group_w = max(120, int(group_w))
    gap = max(4, int(gap))
    chip_h = max(26, int((banner_height - (gap * 3)) / 2))
    chip_w = max(48, int((group_w - (gap * 3)) / 2))
    group_x = float(group_right) - float(group_w)
    row1 = float(banner_top) + float(gap)
    row2 = row1 + float(chip_h) + float(gap)
    col1 = group_x + float(gap)
    col2 = col1 + float(chip_w) + float(gap)
    return {
        "chip_w": chip_w,
        "chip_h": chip_h,
        "row1": row1,
        "row2": row2,
        "col1": col1,
        "col2": col2,
    }


def _pdf_fit_font_size(text, font_name, max_w, preferred_size, min_size):
    from reportlab.pdfbase import pdfmetrics

    size = float(preferred_size)
    text = str(text or "")
    max_w = max(12.0, float(max_w))
    min_size = float(min_size)
    while size > min_size and pdfmetrics.stringWidth(text, font_name, size) > max_w:
        size -= 0.3
    return max(min_size, round(size, 2))


def _pdf_draw_banner_title_block(
    c,
    page_h,
    title,
    subtitle,
    left_x,
    right_x,
    title_y,
    subtitle_y,
    bold_font,
    regular_font,
    title_size=18.0,
    subtitle_size=9.0,
    min_title_size=14.2,
    min_subtitle_size=7.6,
):
    area_w = max(40.0, float(right_x) - float(left_x))
    center_x = float(left_x) + (area_w / 2.0)
    fitted_title_size = _pdf_fit_font_size(title, bold_font, area_w, title_size, min_title_size)
    fitted_subtitle_size = _pdf_fit_font_size(subtitle, regular_font, area_w, subtitle_size, min_subtitle_size)
    title_txt = _pdf_clip_text(title, area_w, bold_font, fitted_title_size)
    subtitle_txt = _pdf_clip_text(subtitle, area_w, regular_font, fitted_subtitle_size)
    c.setFont(bold_font, fitted_title_size)
    c.drawCentredString(center_x, page_h - title_y, pdf_normalize_text(title_txt))
    c.setFont(regular_font, fitted_subtitle_size)
    c.drawCentredString(center_x, page_h - subtitle_y, pdf_normalize_text(subtitle_txt))

def refresh_expedicao(self):
    _ensure_configured()
    if getattr(self, "_refresh_expedicao_busy", False):
        self._refresh_expedicao_pending = True
        return
    self._refresh_expedicao_busy = True
    self._refresh_expedicao_pending = False
    try:
        if not hasattr(self, "exp_tbl_enc"):
            return
        self.refresh_expedicao_pending()
        self.refresh_expedicao_pecas()
        self.refresh_expedicao_draft()
        self.refresh_expedicao_hist()
    finally:
        self._refresh_expedicao_busy = False
        if getattr(self, "_refresh_expedicao_pending", False):
            self._refresh_expedicao_pending = False
            try:
                self.root.after(20, self.refresh_expedicao)
            except Exception:
                pass

def clear_expedicao_selection(self):
    _ensure_configured()
    self.exp_sel_enc_num = None
    self.exp_peca_row_map = {}
    self.exp_draft_linhas = []
    if hasattr(self, "exp_tbl_pecas"):
        self.exp_tbl_pecas.delete(*self.exp_tbl_pecas.get_children())
    if hasattr(self, "exp_tbl_linhas"):
        self.exp_tbl_linhas.delete(*self.exp_tbl_linhas.get_children())

def get_selected_expedicao(self, show_error=True):
    _ensure_configured()
    if not hasattr(self, "exp_tbl_hist"):
        if show_error:
            messagebox.showerror("Erro", "Historico de expedicoes indisponivel.")
        return None
    sel = self.exp_tbl_hist.selection()
    if not sel:
        if show_error:
            messagebox.showerror("Erro", "Selecione uma guia no historico.")
        return None
    num = str(self.exp_tbl_hist.item(sel[0], "values")[0]).strip()
    ex = next((x for x in self.data.get("expedicoes", []) if x.get("numero") == num), None)
    if not ex and show_error:
        messagebox.showerror("Erro", "Guia nao encontrada.")
    return ex

def refresh_expedicao_pending(self):
    _ensure_configured()
    self.exp_tbl_enc.delete(*self.exp_tbl_enc.get_children())
    query = (self.exp_filter.get().strip().lower() if hasattr(self, "exp_filter") else "")
    estado_f = (self.exp_estado_filter.get() or "Todas").strip()
    selected_iid = None
    selected_found = False
    row_i = 0
    for enc in self.data.get("encomendas", []):
        update_estado_expedicao_encomenda(enc)
        disp_total = 0.0
        for p in encomenda_pecas(enc):
            disp_total += peca_qtd_disponivel_expedicao(p)
        if disp_total <= 0:
            continue
        if estado_f != "Todas" and str(enc.get("estado_expedicao", "")) != estado_f:
            continue
        cli_code = str(enc.get("cliente", "") or "")
        cli_obj = find_cliente(self.data, cli_code)
        cli_nome = cli_obj.get("nome", "") if cli_obj else ""
        cli_txt = f"{cli_code} - {cli_nome}".strip(" - ")
        row = (
            enc.get("numero", ""),
            cli_txt,
            enc.get("estado", ""),
            enc.get("estado_expedicao", "Nao expedida"),
            fmt_num(disp_total),
        )
        if query and not any(query in str(v).lower() for v in row):
            continue
        tag = "enc_even" if row_i % 2 == 0 else "enc_odd"
        estado_exp = str(enc.get("estado_expedicao", "Nao expedida"))
        if estado_exp == "Nao expedida":
            tag = "enc_nao"
        elif estado_exp == "Parcialmente expedida":
            tag = "enc_parcial"
        elif estado_exp == "Totalmente expedida":
            tag = "enc_total"
        iid = self.exp_tbl_enc.insert("", END, values=row, tags=(tag,))
        row_i += 1
        if self.exp_sel_enc_num and self.exp_sel_enc_num == enc.get("numero"):
            selected_iid = iid
            selected_found = True
    if not selected_found:
        self.exp_sel_enc_num = None
        self.exp_draft_linhas = []
    if selected_iid:
        try:
            self.exp_tbl_enc.selection_set(selected_iid)
            self.exp_tbl_enc.focus(selected_iid)
            self.exp_tbl_enc.see(selected_iid)
        except Exception:
            pass

def refresh_expedicao_pecas(self):
    _ensure_configured()
    if not hasattr(self, "exp_tbl_pecas"):
        return
    self.exp_tbl_pecas.delete(*self.exp_tbl_pecas.get_children())
    self.exp_peca_row_map = {}
    enc = self.get_exp_selected_encomenda()
    if not enc:
        return
    draft_used = {}
    for l in self.exp_draft_linhas:
        key = str(l.get("_key", ""))
        draft_used[key] = draft_used.get(key, 0.0) + parse_float(l.get("qtd", 0), 0.0)
    for p in encomenda_pecas(enc):
        ensure_peca_operacoes(p)
        base_disp = peca_qtd_disponivel_expedicao(p)
        key = str(p.get("id", "") or f"{p.get('ref_interna','')}|{p.get('ref_externa','')}")
        disp = base_disp - draft_used.get(key, 0.0)
        if disp <= 0:
            continue
        iid = self.exp_tbl_pecas.insert(
            "",
            END,
            values=(
                p.get("id", ""),
                p.get("ref_interna", ""),
                p.get("ref_externa", ""),
                fmt_num(p.get("produzido_ok", 0)),
                fmt_num(p.get("qtd_expedida", 0)),
                fmt_num(disp),
            ),
        )
        self.exp_peca_row_map[iid] = p

def refresh_expedicao_draft(self):
    _ensure_configured()
    if not hasattr(self, "exp_tbl_linhas"):
        return
    self.exp_tbl_linhas.delete(*self.exp_tbl_linhas.get_children())
    for l in self.exp_draft_linhas:
        self.exp_tbl_linhas.insert(
            "",
            END,
            values=(
                l.get("peca_id", ""),
                l.get("ref_interna", ""),
                l.get("ref_externa", ""),
                l.get("descricao", ""),
                fmt_num(l.get("qtd", 0)),
            ),
        )

def refresh_expedicao_hist(self):
    _ensure_configured()
    if not hasattr(self, "exp_tbl_hist"):
        return
    self.exp_tbl_hist.delete(*self.exp_tbl_hist.get_children())
    query = (self.exp_filter.get().strip().lower() if hasattr(self, "exp_filter") else "")
    rows = list(self.data.get("expedicoes", []))
    rows.sort(key=lambda x: str(x.get("data_emissao", "")), reverse=True)
    for idx, ex in enumerate(rows):
        row = (
            ex.get("numero", ""),
            ex.get("tipo", ""),
            ex.get("encomenda", ""),
            ex.get("cliente_nome", ex.get("cliente", "")),
            str(ex.get("data_emissao", "")).replace("T", " ")[:19],
            ex.get("estado", ""),
            len(ex.get("linhas", [])),
        )
        if query and not any(query in str(v).lower() for v in row):
            continue
        tag = "hist_even" if idx % 2 == 0 else "hist_odd"
        if bool(ex.get("anulada")):
            tag = "hist_anulada"
        self.exp_tbl_hist.insert("", END, values=row, tags=(tag,))

def editar_expedicao(self):
    _ensure_configured()
    ex = self.get_selected_expedicao(show_error=True)
    if not ex:
        return
    if ex.get("anulada"):
        messagebox.showerror("Erro", "Nao e possivel editar uma guia anulada.")
        return
    if not messagebox.askyesno(
        "Aviso",
        "Para rastreabilidade, o procedimento recomendado e anular e reemitir.\n\nPretende editar esta guia mesmo assim?",
    ):
        return
    dados = self.prompt_dados_guia(
        f"Editar Guia {ex.get('numero','')}",
        {
            "codigo_at": ex.get("at_validation_code", ex.get("codigo_at", "")),
            "tipo_via": ex.get("tipo_via", "Original"),
            "emitente_nome": ex.get("emitente_nome", ""),
            "emitente_nif": ex.get("emitente_nif", ""),
            "emitente_morada": ex.get("emitente_morada", ""),
            "destinatario": ex.get("destinatario", ""),
            "dest_nif": ex.get("dest_nif", ""),
            "dest_morada": ex.get("dest_morada", ""),
            "local_carga": ex.get("local_carga", ""),
            "local_descarga": ex.get("local_descarga", ""),
            "data_transporte": ex.get("data_transporte", now_iso()),
            "transportador": ex.get("transportador", ""),
            "matricula": ex.get("matricula", ""),
            "observacoes": ex.get("observacoes", ""),
        },
        ex.get("linhas", []),
    )
    if not dados:
        return
    ex["codigo_at"] = dados.get("codigo_at", "")
    ex["at_validation_code"] = dados.get("codigo_at", "")
    seq_num = int(parse_float(ex.get("seq_num", 0), 0) or 0)
    if ex.get("at_validation_code") and seq_num > 0:
        ex["atcud"] = f"{str(ex.get('at_validation_code', '')).strip()}-{seq_num}"
    ex["tipo_via"] = "Original"
    ex["emitente_nome"] = dados.get("emitente_nome", "")
    ex["emitente_nif"] = dados.get("emitente_nif", "")
    ex["emitente_morada"] = dados.get("emitente_morada", "")
    ex["destinatario"] = dados.get("destinatario", "")
    ex["dest_nif"] = dados.get("dest_nif", "")
    ex["dest_morada"] = dados.get("dest_morada", "")
    ex["local_carga"] = dados.get("local_carga", "")
    ex["local_descarga"] = dados.get("local_descarga", "")
    ex["data_transporte"] = dados.get("data_transporte", now_iso())
    ex["transportador"] = dados.get("transportador", "")
    ex["matricula"] = dados.get("matricula", "")
    ex["observacoes"] = dados.get("observacoes", "")
    save_data(self.data)
    self.refresh_expedicao()
    if messagebox.askyesno("OK", "Guia atualizada. Pretende abrir o PDF?"):
        self.preview_expedicao_pdf_by_num(ex.get("numero", ""))

def anular_expedicao(self):
    _ensure_configured()
    if not hasattr(self, "exp_tbl_hist"):
        return
    sel = self.exp_tbl_hist.selection()
    if not sel:
        messagebox.showerror("Erro", "Selecione uma guia no historico.")
        return
    num = str(self.exp_tbl_hist.item(sel[0], "values")[0]).strip()
    ex = next((x for x in self.data.get("expedicoes", []) if x.get("numero") == num), None)
    if not ex:
        messagebox.showerror("Erro", "Guia nao encontrada.")
        return
    if ex.get("anulada"):
        messagebox.showinfo("Info", "A guia ja esta anulada.")
        return
    if not messagebox.askyesno("Confirmar", f"Anular guia {num}?"):
        return
    motivo = simple_input(self.root, "Anular Guia", "Justificacao:")
    if not motivo:
        messagebox.showerror("Erro", "E obrigatorio indicar justificacao.")
        return
    for l in ex.get("linhas", []):
        if l.get("manual"):
            cod = str(l.get("ref_interna", "") or "").strip()
            qtd = parse_float(l.get("qtd", 0), 0.0)
            if cod:
                prod = next((x for x in self.data.get("produtos", []) if x.get("codigo") == cod), None)
                if prod:
                    prod["qty"] = parse_float(prod.get("qty", 0), 0.0) + qtd
                    prod["atualizado_em"] = now_iso()
            continue
        enc = self.get_encomenda_by_numero(l.get("encomenda", ""))
        if not enc:
            continue
        pid = str(l.get("peca_id", "") or "")
        qtd = parse_float(l.get("qtd", 0), 0.0)
        for p in encomenda_pecas(enc):
            if str(p.get("id", "") or "") == pid:
                p["qtd_expedida"] = max(0.0, parse_float(p.get("qtd_expedida", 0), 0.0) - qtd)
                break
        update_estado_expedicao_encomenda(enc)
    ex["anulada"] = True
    ex["estado"] = "Anulada"
    ex["anulada_motivo"] = motivo.strip()
    save_data(self.data)
    self.refresh_expedicao()
    messagebox.showinfo("OK", "Guia anulada.")

def refresh_ne_fornecedores(self):
    _ensure_configured()
    values = []
    for f in self.data.get("fornecedores", []):
        values.append(f"{f.get('id','')} - {f.get('nome','')}".strip(" -"))
    if hasattr(self, "ne_fornecedor_cb"):
        try:
            if CUSTOM_TK_AVAILABLE and isinstance(self.ne_fornecedor_cb, ctk.CTkComboBox):
                self.ne_fornecedor_cb.configure(values=values or [""])
            else:
                self.ne_fornecedor_cb["values"] = values
        except Exception:
            try:
                self.ne_fornecedor_cb["values"] = values
            except Exception:
                pass

def on_ne_fornecedor_change(self, _e=None):
    _ensure_configured()
    txt = self.ne_fornecedor.get().strip()
    fid = ""
    nome = txt
    if " - " in txt:
        fid, nome = txt.split(" - ", 1)
    fobj = None
    if fid:
        fobj = next((f for f in self.data.get("fornecedores", []) if f.get("id") == fid), None)
    if not fobj:
        fobj = next((f for f in self.data.get("fornecedores", []) if f.get("nome", "").strip().lower() == nome.strip().lower()), None)
    if fobj:
        self.ne_fornecedor_id.set(fobj.get("id", ""))
        self.ne_fornecedor.set(f"{fobj.get('id','')} - {fobj.get('nome','')}".strip(" -"))
        self.ne_contacto.set(fobj.get("contacto", ""))

def manage_fornecedores(self):
    _ensure_configured()
    use_custom = CUSTOM_TK_AVAILABLE and os.environ.get("USE_CUSTOM_NE", "1") != "0"
    dlg = ctk.CTkToplevel(self.root) if use_custom else Toplevel(self.root)
    dlg.title("Fornecedores")
    dlg.geometry("1180x680")
    dlg.grab_set()

    top = ctk.CTkFrame(dlg, fg_color="#f7f8fb") if use_custom else ttk.Frame(dlg)
    top.pack(fill="x", padx=8, pady=8)
    (ctk.CTkLabel if use_custom else ttk.Label)(top, text="Pesquisar").pack(side="left", padx=4)
    f_filter = StringVar()
    search_entry = (ctk.CTkEntry if use_custom else ttk.Entry)(top, textvariable=f_filter, width=300 if use_custom else 40)
    search_entry.pack(side="left", padx=4)

    cols = ("id", "nome", "nif", "contacto", "email", "localidade", "pais")
    wrap = ctk.CTkFrame(dlg, fg_color="#ffffff") if use_custom else ttk.Frame(dlg)
    wrap.pack(fill="both", expand=True, padx=8, pady=4)
    tree = ttk.Treeview(wrap, columns=cols, show="headings", height=10, style="NE.Treeview" if use_custom else "")
    heads = {
        "id": "ID",
        "nome": "Nome",
        "nif": "NIF",
        "contacto": "Contacto",
        "email": "Email",
        "localidade": "Localidade",
        "pais": "País",
    }
    for c in cols:
        tree.heading(c, text=heads[c])
    tree.column("id", width=110, anchor="w")
    tree.column("nome", width=220, anchor="w")
    tree.column("nif", width=120, anchor="w")
    tree.column("contacto", width=130, anchor="w")
    tree.column("email", width=220, anchor="w")
    tree.column("localidade", width=170, anchor="w")
    tree.column("pais", width=120, anchor="w")
    vsb = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(wrap, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")
    wrap.rowconfigure(0, weight=1)
    wrap.columnconfigure(0, weight=1)
    tree.tag_configure("even", background="#f8fbff")
    tree.tag_configure("odd", background="#eef4fb")

    frm = ctk.CTkFrame(dlg, fg_color="#f7f8fb") if use_custom else ttk.Frame(dlg)
    frm.pack(fill="x", padx=8, pady=8)
    v_id = StringVar(value=next_fornecedor_numero(self.data))
    v_nome = StringVar()
    v_nif = StringVar()
    v_contacto = StringVar()
    v_email = StringVar()
    v_morada = StringVar()
    v_cp = StringVar()
    v_localidade = StringVar()
    v_pais = StringVar(value="Portugal")
    v_cond_pag = StringVar()
    v_prazo = StringVar()
    v_site = StringVar()
    v_obs = StringVar()
    fields = [
        ("ID", v_id, 0, 0),
        ("Nome", v_nome, 0, 2),
        ("NIF", v_nif, 1, 0),
        ("Contacto", v_contacto, 1, 2),
        ("Email", v_email, 2, 0),
        ("Morada", v_morada, 2, 2),
        ("Cód. Postal", v_cp, 3, 0),
        ("Localidade", v_localidade, 3, 2),
        ("País", v_pais, 4, 0),
        ("Cond. Pagamento", v_cond_pag, 4, 2),
        ("Prazo Entrega (dias)", v_prazo, 5, 0),
        ("Website", v_site, 5, 2),
        ("Obs.", v_obs, 6, 0),
    ]
    for txt, var, r, c in fields:
        (ctk.CTkLabel if use_custom else ttk.Label)(frm, text=txt).grid(row=r, column=c, sticky="w", padx=4, pady=2)
        wide = txt in ("Nome", "Morada", "Email", "Localidade", "Cond. Pagamento", "Website", "Obs.")
        (ctk.CTkEntry if use_custom else ttk.Entry)(
            frm,
            textvariable=var,
            width=280 if (wide and use_custom) else (36 if wide else 20),
        ).grid(row=r, column=c + 1, sticky="we", padx=4, pady=2)

    def refresh():
        q = f_filter.get().lower()
        for i in tree.get_children():
            tree.delete(i)
        for idx, f in enumerate(self.data.get("fornecedores", [])):
            vals = (
                f.get("id", ""),
                f.get("nome", ""),
                f.get("nif", ""),
                f.get("contacto", ""),
                f.get("email", ""),
                f.get("localidade", ""),
                f.get("pais", ""),
            )
            if q and not any(q in str(v).lower() for v in vals):
                continue
            tag = "odd" if idx % 2 else "even"
            tree.insert("", END, values=vals, tags=(tag,))

    def novo():
        v_id.set(next_fornecedor_numero(self.data))
        v_nome.set("")
        v_nif.set("")
        v_contacto.set("")
        v_email.set("")
        v_morada.set("")
        v_cp.set("")
        v_localidade.set("")
        v_pais.set("Portugal")
        v_cond_pag.set("")
        v_prazo.set("")
        v_site.set("")
        v_obs.set("")

    def guardar():
        if not v_nome.get().strip():
            messagebox.showerror("Erro", "Nome do fornecedor e obrigatorio")
            return
        rec = {
            "id": v_id.get().strip() or next_fornecedor_numero(self.data),
            "nome": v_nome.get().strip(),
            "nif": v_nif.get().strip(),
            "contacto": v_contacto.get().strip(),
            "email": v_email.get().strip(),
            "morada": v_morada.get().strip(),
            "codigo_postal": v_cp.get().strip(),
            "localidade": v_localidade.get().strip(),
            "pais": v_pais.get().strip(),
            "cond_pagamento": v_cond_pag.get().strip(),
            "prazo_entrega_dias": int(parse_float(v_prazo.get(), 0)) if str(v_prazo.get()).strip() else 0,
            "website": v_site.get().strip(),
            "obs": v_obs.get().strip(),
        }
        lst = self.data.setdefault("fornecedores", [])
        ex = next((x for x in lst if x.get("id") == rec["id"]), None)
        if ex:
            ex.update(rec)
        else:
            lst.append(rec)
        save_data(self.data)
        refresh()
        self.refresh_ne_fornecedores()
        self.ne_fornecedor.set(f"{rec['id']} - {rec['nome']}")
        self.ne_contacto.set(rec.get("contacto", ""))

    def remover():
        sel = tree.selection()
        if not sel:
            return
        fid = tree.item(sel[0], "values")[0]
        self.data["fornecedores"] = [f for f in self.data.get("fornecedores", []) if f.get("id") != fid]
        save_data(self.data)
        refresh()
        self.refresh_ne_fornecedores()

    def on_sel(_e=None):
        sel = tree.selection()
        if not sel:
            return
        vals = tree.item(sel[0], "values")
        v_id.set(vals[0]); v_nome.set(vals[1]); v_nif.set(vals[2]); v_contacto.set(vals[3]); v_email.set(vals[4])
        fobj = next((f for f in self.data.get("fornecedores", []) if str(f.get("id", "")) == str(vals[0])), {})
        v_morada.set(fobj.get("morada", ""))
        v_cp.set(fobj.get("codigo_postal", ""))
        v_localidade.set(fobj.get("localidade", ""))
        v_pais.set(fobj.get("pais", ""))
        v_cond_pag.set(fobj.get("cond_pagamento", ""))
        v_prazo.set(str(fobj.get("prazo_entrega_dias", "") or ""))
        v_site.set(fobj.get("website", ""))
        v_obs.set(fobj.get("obs", ""))

    tree.bind("<<TreeviewSelect>>", on_sel)
    f_filter.trace_add("write", lambda *_: refresh())

    btns = ctk.CTkFrame(dlg, fg_color="#f7f8fb") if use_custom else ttk.Frame(dlg)
    btns.pack(fill="x", padx=8, pady=8)
    if use_custom:
        ctk.CTkButton(btns, text="Novo", command=novo, width=120, fg_color="#f59e0b", hover_color="#d97706").pack(side="left", padx=4)
        ctk.CTkButton(btns, text="Guardar", command=guardar, width=120).pack(side="left", padx=4)
        ctk.CTkButton(btns, text="Remover", command=remover, width=120).pack(side="left", padx=4)
        ctk.CTkButton(btns, text="Fechar", command=dlg.destroy, width=120).pack(side="right", padx=4)
    else:
        ttk.Button(btns, text="Novo", command=novo).pack(side="left", padx=4)
        ttk.Button(btns, text="Guardar", command=guardar).pack(side="left", padx=4)
        ttk.Button(btns, text="Remover", command=remover).pack(side="left", padx=4)
        ttk.Button(btns, text="Fechar", command=dlg.destroy).pack(side="right", padx=4)

    refresh()

def ne_collect_lines(self):
    _ensure_configured()
    linhas = []
    _ne_meta_cleanup(self)
    for item in self.tbl_ne_linhas.get_children():
        v = self.tbl_ne_linhas.item(item, "values")
        d = self._ne_line_parse_values(v)
        desconto = parse_float(d.get("desconto", _ne_meta_get(self, item, "desconto", 0)), 0)
        iva = parse_float(d.get("iva", _ne_meta_get(self, item, "iva", 23)), 23)
        entregue_txt = str(d.get("entregue_txt", "PENDENTE")).strip().upper()
        line = {
            "ref": d.get("ref", ""),
            "descricao": d.get("descricao", ""),
            "fornecedor_linha": (d.get("fornecedor_linha", "") or "").strip(),
            "origem": d.get("origem", "Produto"),
            "qtd": parse_float(d.get("qtd", 0), 0),
            "unid": d.get("unid", ""),
            "preco": parse_float(d.get("preco", 0), 0),
            "total": parse_float(d.get("total", 0), 0),
            "desconto": desconto,
            "iva": iva,
            "entregue": entregue_txt in ("SIM", "ENTREGUE", "TRUE", "1"),
            "qtd_entregue": parse_float(d.get("qtd", 0), 0) if (entregue_txt in ("SIM", "ENTREGUE", "TRUE", "1")) else 0.0,
        }
        # Para linhas de Materia-Prima, guarda dados tecnicos para criar/atualizar postura no recebimento.
        if origem_is_materia(line.get("origem")):
            m = next((x for x in self.data.get("materiais", []) if x.get("id") == line.get("ref")), None)
            if m:
                line.update(
                    {
                        "material": m.get("material", ""),
                        "espessura": m.get("espessura", ""),
                        "comprimento": parse_float(m.get("comprimento", 0), 0),
                        "largura": parse_float(m.get("largura", 0), 0),
                        "metros": parse_float(m.get("metros", 0), 0),
                        "localizacao": m.get("Localizacao", ""),
                        "lote_fornecedor": m.get("lote_fornecedor", ""),
                        "peso_unid": parse_float(m.get("peso_unid", 0), 0),
                        "p_compra": parse_float(m.get("p_compra", 0), 0),
                        "formato": m.get("formato", detect_materia_formato(m)),
                    }
                )
        linhas.append(line)
    return linhas

def nova_ne(self, create_draft=True):
    _ensure_configured()
    novo_num = next_ne_numero(self.data)
    self.ne_num.set(novo_num)
    if hasattr(self, "ne_filter"):
        self.ne_filter.set("")
    self.ne_fornecedor.set("")
    self.ne_fornecedor_id.set("")
    self.ne_contacto.set("")
    self.ne_entrega.set("")
    self.ne_obs.set("")
    self.ne_local_descarga.set("")
    self.ne_meio_transporte.set("")
    for i in getattr(self, "tbl_ne_linhas", []).get_children():
        self.tbl_ne_linhas.delete(i)
    self.refresh_ne_fornecedores()
    self.refresh_ne_total()
    if create_draft:
        # criar rascunho para aparecer imediatamente na grelha superior
        lst = self.data.setdefault("notas_encomenda", [])
        if not any(n.get("numero") == novo_num for n in lst):
            lst.append(
                {
                    "numero": novo_num,
                    "fornecedor": "",
                    "fornecedor_id": "",
                    "contacto": "",
                    "data_entrega": "",
                    "obs": "",
                    "local_descarga": "",
                    "meio_transporte": "",
                    "linhas": [],
                    "total": 0.0,
                    "estado": "Em edicao",
                    "oculta": False,
                    "_draft": True,
                }
            )
            save_data(self.data)
    self.refresh_ne()
    if create_draft:
        for iid in self.tbl_ne.get_children():
            if self.tbl_ne.item(iid, "values")[0] == novo_num:
                self.tbl_ne.selection_set(iid)
                self.tbl_ne.see(iid)
                break

def remover_ne(self):
    _ensure_configured()
    sel = self.tbl_ne.selection()
    if not sel:
        return
    num = self.tbl_ne.item(sel[0], "values")[0]
    if not messagebox.askyesno("Confirmar", f"Remover nota de encomenda {num}?"):
        return
    self.data["notas_encomenda"] = [n for n in self.data.get("notas_encomenda", []) if n.get("numero") != num]
    save_data(self.data)
    self.refresh_ne()
    # Nao criar automaticamente nova NE ao remover.
    rows = self.tbl_ne.get_children()
    if rows:
        self.tbl_ne.selection_set(rows[0])
        self.tbl_ne.focus(rows[0])
        self.tbl_ne.see(rows[0])
        self.on_ne_select()
    else:
        self.ne_num.set(peek_next_ne_numero(self.data))
        self.ne_fornecedor.set("")
        self.ne_fornecedor_id.set("")
        self.ne_contacto.set("")
        self.ne_entrega.set("")
        self.ne_obs.set("")
        self.ne_local_descarga.set("")
        self.ne_meio_transporte.set("")
        children = self.tbl_ne_linhas.get_children()
        if children:
            self.tbl_ne_linhas.delete(*children)
        self.refresh_ne_total()

def refresh_ne(self):
    _ensure_configured()
    if getattr(self, "_refresh_ne_busy", False):
        self._refresh_ne_pending = True
        return
    self._refresh_ne_busy = True
    self._refresh_ne_pending = False
    try:
        sel_before = self.tbl_ne.selection() if hasattr(self, "tbl_ne") else ()
        selected_num = ""
        if sel_before:
            try:
                selected_num = str(self.tbl_ne.item(sel_before[0], "values")[0])
            except Exception:
                selected_num = ""
        if not selected_num and hasattr(self, "ne_num"):
            selected_num = (self.ne_num.get() or "").strip()
        filtro = self.ne_filter.get().lower() if hasattr(self, "ne_filter") else ""
        estado_filtro = (self.ne_estado_filter.get().strip().lower() if hasattr(self, "ne_estado_filter") else "ativas")
        show_convertidas = bool(self.ne_show_convertidas.get()) if hasattr(self, "ne_show_convertidas") else False
        try:
            if hasattr(self, "ne_estado_segment") and self.ne_estado_segment.winfo_exists():
                desired = self.ne_estado_filter.get() or "Ativas"
                current = ""
                try:
                    current = self.ne_estado_segment.get() or ""
                except Exception:
                    current = ""
                if current != desired:
                    self._suppress_ne_filter_cb = True
                    self.ne_estado_segment.set(desired)
        except Exception:
            pass
        finally:
            self._suppress_ne_filter_cb = False
        self._ne_loaded_num = ""
        children = self.tbl_ne.get_children()
        if children:
            self.tbl_ne.delete(*children)
        try:
            if "normalize_notas_encomenda" in globals():
                normalize_notas_encomenda(self.data)
        except Exception:
            pass
        notas = sorted(self.data.get("notas_encomenda", []), key=lambda x: x.get("numero", ""))
        target_iid = None
        for idx, n in enumerate(notas):
            if n.get("oculta") and not show_convertidas:
                continue
            estado = n.get("estado", "Em edicao")
            estado_norm = norm_text(estado)
            if estado_filtro and estado_filtro not in ("todos", "todas", "all"):
                if "ativ" in estado_filtro:
                    if "entreg" in estado_norm or "convert" in estado_norm:
                        continue
                elif "edi" in estado_filtro and "edi" not in estado_norm:
                    continue
                elif "apro" in estado_filtro and "apro" not in estado_norm:
                    continue
                elif "parcial" in estado_filtro and "parcial" not in estado_norm:
                    continue
                elif "entreg" in estado_filtro and "entreg" not in estado_norm:
                    continue
                elif "convert" in estado_filtro and "convert" not in estado_norm:
                    continue
            vals = (n.get("numero"), n.get("fornecedor"), n.get("data_entrega", ""), estado, fmt_num(n.get("total", 0)))
            if filtro and not any(filtro in str(v).lower() for v in vals):
                continue
            if n.get("_draft"):
                tag = "warn"
            elif (estado or "").strip().lower() == "entregue":
                tag = "entregue"
            elif "parcial" in (estado or "").strip().lower():
                tag = "parcial"
            elif "convertid" in (estado or "").strip().lower():
                tag = "convertida"
            elif "aprovad" in (estado or "").strip().lower():
                tag = "aprovada"
            else:
                tag = "odd" if idx % 2 else "even"
            iid = self.tbl_ne.insert("", END, values=vals, tags=(tag,))
            if selected_num and str(n.get("numero", "")) == selected_num:
                target_iid = iid
        if target_iid:
            try:
                self.tbl_ne.selection_set(target_iid)
                self.tbl_ne.focus(target_iid)
                self.tbl_ne.see(target_iid)
            except Exception:
                pass
    finally:
        self._refresh_ne_busy = False
        if getattr(self, "_refresh_ne_pending", False):
            self._refresh_ne_pending = False
            try:
                self.root.after(20, self.refresh_ne)
            except Exception:
                pass

def on_ne_select(self, _e=None):
    _ensure_configured()
    sel = self.tbl_ne.selection()
    if not sel:
        return
    num = self.tbl_ne.item(sel[0], "values")[0]
    if _e is not None and str(getattr(self, "_ne_loaded_num", "") or "") == str(num):
        return
    ne = next((n for n in self.data.get("notas_encomenda", []) if n.get("numero") == num), None)
    if not ne:
        return
    self._ne_loaded_num = str(num)
    self.ne_num.set(ne.get("numero", ""))
    self.ne_fornecedor.set(ne.get("fornecedor", ""))
    self.ne_fornecedor_id.set(ne.get("fornecedor_id", ""))
    self.ne_contacto.set(ne.get("contacto", ""))
    self.ne_entrega.set(ne.get("data_entrega", ""))
    self.ne_obs.set(ne.get("obs", ""))
    self.ne_local_descarga.set(ne.get("local_descarga", ""))
    self.ne_meio_transporte.set(ne.get("meio_transporte", ""))
    children = self.tbl_ne_linhas.get_children()
    if children:
        self.tbl_ne_linhas.delete(*children)
    self._ne_line_meta = {}
    for l in ne.get("linhas", []):
        qtd_tot = parse_float(l.get("qtd", 0), 0)
        qtd_ent = parse_float(
            l.get("qtd_entregue", l.get("qtd", 0) if l.get("entregue") else 0),
            0,
        )
        if qtd_ent <= 0:
            entregue_txt = "PENDENTE"
        elif qtd_ent < (qtd_tot - 1e-9):
            entregue_txt = f"PARCIAL ({fmt_num(qtd_ent)}/{fmt_num(qtd_tot)})"
        else:
            entregue_txt = "ENTREGUE"
        iid = self.tbl_ne_linhas.insert(
            "",
            END,
            values=(
                l.get("ref", ""),
                l.get("descricao", ""),
                l.get("fornecedor_linha", ne.get("fornecedor", "")),
                l.get("origem", "Produto"),
                fmt_num(l.get("qtd", 0)),
                l.get("unid", ""),
                fmt_num(l.get("preco", 0)),
                fmt_num(l.get("desconto", 0)),
                fmt_num(l.get("iva", 23)),
                fmt_num(l.get("total", 0)),
                entregue_txt,
            ),
        )
        _ne_meta_set(
            self,
            iid,
            desconto=parse_float(l.get("desconto", 0), 0),
            iva=parse_float(l.get("iva", 23), 23),
        )
    self.ne_refresh_line_tags()
    self.refresh_ne_total()

def get_ne_selected(self):
    _ensure_configured()
    sel = self.tbl_ne.selection()
    num = ""
    if sel:
        num = self.tbl_ne.item(sel[0], "values")[0]
    else:
        num = self.ne_num.get().strip() if hasattr(self, "ne_num") else ""
    if not num:
        messagebox.showerror("Erro", "Selecione uma Nota de Encomenda")
        return None
    ne = next((n for n in self.data.get("notas_encomenda", []) if n.get("numero") == num), None)
    if not ne:
        messagebox.showerror("Erro", "Nota de Encomenda nao encontrada")
        return None
    return ne

def aprovar_ne(self):
    _ensure_configured()
    ne = self.get_ne_selected()
    if not ne:
        return
    if not ne.get("linhas"):
        messagebox.showerror("Erro", "A nota nao tem linhas.")
        return
    ne["estado"] = "Aprovada"
    ne["data_aprovacao"] = now_iso()
    ne["_draft"] = False
    save_data(self.data)
    self.refresh_ne()
    for iid in self.tbl_ne.get_children():
        if self.tbl_ne.item(iid, "values")[0] == ne.get("numero"):
            self.tbl_ne.selection_set(iid)
            self.tbl_ne.see(iid)
            break
    self.on_ne_select()
    messagebox.showinfo("OK", f"Nota {ne.get('numero')} aprovada.")

def ne_refresh_line_tags(self):
    _ensure_configured()
    for idx, iid in enumerate(self.tbl_ne_linhas.get_children()):
        vals = self.tbl_ne_linhas.item(iid, "values")
        entregue_txt = str(self._ne_line_parse_values(vals).get("entregue_txt", "PENDENTE")).strip().upper()
        parity = "odd" if idx % 2 else "even"
        if entregue_txt in ("SIM", "ENTREGUE", "TRUE", "1"):
            self.tbl_ne_linhas.item(iid, tags=(f"lin_entregue_{parity}",))
        elif "PARCIAL" in entregue_txt:
            self.tbl_ne_linhas.item(iid, tags=(f"lin_parcial_{parity}",))
        else:
            self.tbl_ne_linhas.item(iid, tags=(parity,))

def refresh_ne_total(self):
    _ensure_configured()
    total = 0.0
    for i in self.tbl_ne_linhas.get_children():
        d = self._ne_line_parse_values(self.tbl_ne_linhas.item(i, "values"))
        total += parse_float(d.get("total", 0), 0)
    self.ne_total_lbl.configure(text=f"Total: {fmt_num(total)} EUR")


def _ne_meta_get(self, iid, key, default=None):
    _ensure_configured()
    meta = getattr(self, "_ne_line_meta", None)
    if not isinstance(meta, dict):
        return default
    row = meta.get(str(iid), {})
    if not isinstance(row, dict):
        return default
    return row.get(key, default)


def _ne_meta_set(self, iid, **kwargs):
    _ensure_configured()
    if not hasattr(self, "_ne_line_meta") or not isinstance(getattr(self, "_ne_line_meta"), dict):
        self._ne_line_meta = {}
    key = str(iid)
    row = self._ne_line_meta.get(key, {})
    if not isinstance(row, dict):
        row = {}
    row.update(kwargs)
    self._ne_line_meta[key] = row


def _ne_meta_cleanup(self):
    _ensure_configured()
    meta = getattr(self, "_ne_line_meta", None)
    if not isinstance(meta, dict):
        self._ne_line_meta = {}
        return
    valid = {str(i) for i in self.tbl_ne_linhas.get_children()}
    dead = [k for k in list(meta.keys()) if k not in valid]
    for k in dead:
        meta.pop(k, None)

def _on_ne_estado_filter_click(self, value=None):
    _ensure_configured()
    if getattr(self, "_suppress_ne_filter_cb", False):
        return
    try:
        if value is not None:
            self.ne_estado_filter.set(value)
    except Exception:
        pass
    self.refresh_ne()

def _ne_line_parse_values(self, vals):
    _ensure_configured()
    v = list(vals or [])
    if len(v) >= 11:
        return {
            "ref": v[0],
            "descricao": v[1],
            "fornecedor_linha": v[2],
            "origem": v[3] or "Produto",
            "qtd": parse_float(v[4], 0),
            "unid": v[5],
            "preco": parse_float(v[6], 0),
            "desconto": parse_float(v[7], 0),
            "iva": parse_float(v[8], 23),
            "total": parse_float(v[9], 0),
            "entregue_txt": str(v[10]),
        }
    if len(v) >= 10:
        return {
            "ref": v[0],
            "descricao": v[1],
            "fornecedor_linha": v[2],
            "origem": v[3] or "Produto",
            "qtd": parse_float(v[4], 0),
            "unid": v[5],
            "preco": parse_float(v[6], 0),
            "desconto": parse_float(v[7], 0),
            "iva": 23.0,
            "total": parse_float(v[8], 0),
            "entregue_txt": str(v[9]),
        }
    if len(v) >= 9:
        return {
            "ref": v[0],
            "descricao": v[1],
            "fornecedor_linha": v[2],
            "origem": v[3] or "Produto",
            "qtd": parse_float(v[4], 0),
            "unid": v[5],
            "preco": parse_float(v[6], 0),
            "desconto": 0.0,
            "iva": 23.0,
            "total": parse_float(v[7], 0),
            "entregue_txt": str(v[8]),
        }
    if len(v) >= 8:
        return {
            "ref": v[0],
            "descricao": v[1],
            "fornecedor_linha": v[2],
            "origem": "Produto",
            "qtd": parse_float(v[3], 0),
            "unid": v[4],
            "preco": parse_float(v[5], 0),
            "desconto": 0.0,
            "iva": 23.0,
            "total": parse_float(v[6], 0),
            "entregue_txt": str(v[7]),
        }
    if len(v) >= 7:
        return {
            "ref": v[0],
            "descricao": v[1],
            "fornecedor_linha": self.ne_fornecedor.get().strip() if hasattr(self, "ne_fornecedor") else "",
            "origem": "Produto",
            "qtd": parse_float(v[2], 0),
            "unid": v[3],
            "preco": parse_float(v[4], 0),
            "desconto": 0.0,
            "iva": 23.0,
            "total": parse_float(v[5], 0),
            "entregue_txt": str(v[6]),
        }
    return {
        "ref": "",
        "descricao": "",
        "fornecedor_linha": "",
        "origem": "Produto",
        "qtd": 0.0,
        "unid": "UN",
        "preco": 0.0,
        "desconto": 0.0,
        "iva": 23.0,
        "total": 0.0,
        "entregue_txt": "PENDENTE",
    }

def ne_add_linha(self):
    _ensure_configured()
    origem = self._dialog_ne_origem_linha()
    if not origem:
        return
    if origem_is_materia(origem):
        mp = self._dialog_escolher_materia_prima_ne()
        if not mp:
            return
        preco = parse_float(mp.get("preco_unid", 0), 0)
        desc = mp.get("descricao", "")
        ref = mp.get("id", "")
        unid = mp.get("unid", "UN")
        forn = self.ne_fornecedor.get().strip()
        origem_txt = "Materia-Prima"
    else:
        prod = self._dialog_escolher_produto(for_ne=True)
        if not prod:
            return
        preco = produto_preco_unitario(prod)
        desc = prod.get("descricao")
        ref = prod.get("codigo")
        unid = prod.get("unid", "UN")
        forn = self.ne_fornecedor.get().strip()
        origem_txt = "Produto"
    iid = self.tbl_ne_linhas.insert(
        "",
        END,
        values=(
            ref,
            desc,
            forn,
            origem_txt,
            1,
            unid,
            fmt_num(preco),
            fmt_num(0),
            fmt_num(23),
            fmt_num(preco),
            "PENDENTE",
        ),
    )
    _ne_meta_set(self, iid, desconto=0.0, iva=23.0)
    self.ne_refresh_line_tags()
    self.refresh_ne_total()

def ne_edit_linha(self, _e=None):
    _ensure_configured()
    sel = self.tbl_ne_linhas.selection()
    if not sel:
        return
    iid = sel[0]
    vals = self.tbl_ne_linhas.item(iid, "values")
    p = self._ne_line_parse_values(vals)
    fornecedor_line = p.get("fornecedor_linha", "") or self.ne_fornecedor.get().strip()
    origem_txt = p.get("origem", "Produto")
    ref = p.get("ref", "")
    desc = p.get("descricao", "")
    unid = p.get("unid", "UN")
    entregue_txt = p.get("entregue_txt", "PENDENTE")
    use_custom = CUSTOM_TK_AVAILABLE and os.environ.get("USE_CUSTOM_NE", "1") != "0"
    DlgCls = ctk.CTkToplevel if use_custom else Toplevel
    dlg = DlgCls(self.root)
    dlg.title("Editar Linha")
    try:
        dlg.transient(self.root)
    except Exception:
        pass
    if use_custom:
        try:
            dlg.geometry("620x520")
            dlg.minsize(620, 520)
            dlg.resizable(False, False)
        except Exception:
            pass
    else:
        try:
            dlg.geometry("600x500")
            dlg.minsize(600, 500)
        except Exception:
            pass
    dlg.grab_set()
    v_qtd = StringVar(value=str(p.get("qtd", 0)))
    v_preco = StringVar(value=str(p.get("preco", 0)))
    v_desc_perc = StringVar(value=str(fmt_num(_ne_meta_get(self, iid, "desconto", 0))))
    v_iva_perc = StringVar(value=str(fmt_num(_ne_meta_get(self, iid, "iva", p.get("iva", 23)))))
    v_forn = StringVar(value=str(fornecedor_line or ""))
    forn_values = []
    for f in self.data.get("fornecedores", []):
        fid = str(f.get("id", "")).strip()
        nome = str(f.get("nome", "")).strip()
        s = f"{fid} - {nome}".strip(" -")
        if s:
            forn_values.append(s)
    if self.ne_fornecedor.get().strip() and self.ne_fornecedor.get().strip() not in forn_values:
        forn_values.insert(0, self.ne_fornecedor.get().strip())

    def on_ok():
        qtd = parse_float(v_qtd.get(), 0)
        preco = parse_float(v_preco.get(), 0)
        desc_perc = max(0.0, min(100.0, parse_float(v_desc_perc.get(), 0)))
        iva_perc = max(0.0, min(100.0, parse_float(v_iva_perc.get(), 23)))
        total_bruto = qtd * preco
        total_sem_iva = total_bruto * (1.0 - (desc_perc / 100.0))
        total = total_sem_iva * (1.0 + (iva_perc / 100.0))
        forn_txt = (v_forn.get() or "").strip()
        if len(vals) >= 11:
            self.tbl_ne_linhas.item(
                iid,
                values=(
                    ref,
                    desc,
                    forn_txt,
                    origem_txt,
                    fmt_num(qtd),
                    unid,
                    fmt_num(preco),
                    fmt_num(desc_perc),
                    fmt_num(iva_perc),
                    fmt_num(total),
                    entregue_txt,
                ),
            )
        elif len(vals) >= 10:
            self.tbl_ne_linhas.item(
                iid,
                values=(ref, desc, forn_txt, origem_txt, fmt_num(qtd), unid, fmt_num(preco), fmt_num(desc_perc), fmt_num(total), entregue_txt),
            )
        elif len(vals) >= 9:
            self.tbl_ne_linhas.item(
                iid,
                values=(ref, desc, forn_txt, origem_txt, fmt_num(qtd), unid, fmt_num(preco), fmt_num(total), entregue_txt),
            )
        elif len(vals) >= 8:
            self.tbl_ne_linhas.item(iid, values=(ref, desc, forn_txt, fmt_num(qtd), unid, fmt_num(preco), fmt_num(total), entregue_txt))
        else:
            self.tbl_ne_linhas.item(iid, values=(ref, desc, fmt_num(qtd), unid, fmt_num(preco), fmt_num(total), entregue_txt))
        if origem_is_materia(origem_txt):
            if self._update_materia_preco_from_unit(ref, preco):
                self.sync_all_ne_from_materia()
                try:
                    self.refresh_materia()
                except Exception:
                    pass
        else:
            if self._update_produto_preco_from_unit(ref, preco):
                save_data(self.data)
                try:
                    self.refresh_produtos()
                except Exception:
                    pass
        self.ne_refresh_line_tags()
        _ne_meta_set(self, iid, desconto=desc_perc, iva=iva_perc)
        self.refresh_ne_total()
        dlg.destroy()

    if use_custom:
        wrap = ctk.CTkFrame(dlg, fg_color="#f7f8fb", corner_radius=12)
        wrap.pack(fill="both", expand=True, padx=10, pady=10)
        try:
            wrap.grid_columnconfigure(0, weight=0)
            wrap.grid_columnconfigure(1, weight=1)
            wrap.grid_columnconfigure(2, weight=1)
        except Exception:
            pass
        ctk.CTkLabel(
            wrap,
            text=f"Item: {ref}",
            font=("Segoe UI", 15, "bold"),
            text_color="#7a0f1a",
        ).grid(row=0, column=0, columnspan=3, padx=12, pady=(12, 6), sticky="w")
        ctk.CTkLabel(wrap, text=f"Origem: {origem_txt}", font=("Segoe UI", 11), text_color="#4a5f77").grid(row=1, column=0, columnspan=3, padx=12, pady=(0, 10), sticky="w")
        ctk.CTkLabel(wrap, text="Fornecedor", font=("Segoe UI", 12, "bold")).grid(row=2, column=0, padx=12, pady=6, sticky="w")
        ctk.CTkComboBox(wrap, variable=v_forn, values=forn_values or [""], width=320).grid(row=2, column=1, columnspan=2, padx=12, pady=6, sticky="w")
        ctk.CTkLabel(wrap, text="Quantidade", font=("Segoe UI", 12, "bold")).grid(row=3, column=0, padx=12, pady=6, sticky="w")
        ctk.CTkEntry(wrap, textvariable=v_qtd, width=180).grid(row=3, column=1, padx=12, pady=6, sticky="w")
        ctk.CTkLabel(wrap, text="Preço (EUR)", font=("Segoe UI", 12, "bold")).grid(row=4, column=0, padx=12, pady=6, sticky="w")
        ctk.CTkEntry(wrap, textvariable=v_preco, width=180).grid(row=4, column=1, padx=12, pady=6, sticky="w")
        desc_box = ctk.CTkFrame(wrap, fg_color="#eef4fb", corner_radius=10, border_width=1, border_color="#cfd9e6")
        desc_box.grid(row=5, column=0, columnspan=3, sticky="we", padx=12, pady=(8, 6))
        ctk.CTkLabel(desc_box, text="Desconto fornecedor (%)", font=("Segoe UI", 12, "bold")).grid(row=0, column=0, padx=10, pady=8, sticky="w")
        ctk.CTkEntry(desc_box, textvariable=v_desc_perc, width=140).grid(row=0, column=1, padx=8, pady=8, sticky="w")
        ctk.CTkLabel(desc_box, text="Aplicado ao total da linha", font=("Segoe UI", 10), text_color="#4a5f77").grid(row=0, column=2, padx=8, pady=8, sticky="w")
        ctk.CTkLabel(desc_box, text="IVA (%)", font=("Segoe UI", 12, "bold")).grid(row=1, column=0, padx=10, pady=8, sticky="w")
        ctk.CTkEntry(desc_box, textvariable=v_iva_perc, width=140).grid(row=1, column=1, padx=8, pady=8, sticky="w")
        ctk.CTkLabel(desc_box, text="Pré-definido: 23", font=("Segoe UI", 10), text_color="#4a5f77").grid(row=1, column=2, padx=8, pady=8, sticky="w")
        totais = ctk.CTkFrame(wrap, fg_color="#ffffff", corner_radius=10, border_width=1, border_color="#d8dee8")
        totais.grid(row=6, column=0, columnspan=3, sticky="we", padx=12, pady=(6, 4))
        total_bruto_var = StringVar(value="0.00")
        total_sem_iva_var = StringVar(value="0.00")
        total_final_var = StringVar(value="0.00")
        ctk.CTkLabel(totais, text="Total bruto:", font=("Segoe UI", 11, "bold")).grid(row=0, column=0, padx=10, pady=6, sticky="w")
        ctk.CTkLabel(totais, textvariable=total_bruto_var, font=("Segoe UI", 11)).grid(row=0, column=1, padx=6, pady=6, sticky="w")
        ctk.CTkLabel(totais, text="Total s/ IVA:", font=("Segoe UI", 11, "bold")).grid(row=0, column=2, padx=12, pady=6, sticky="w")
        ctk.CTkLabel(totais, textvariable=total_sem_iva_var, font=("Segoe UI", 11)).grid(row=0, column=3, padx=6, pady=6, sticky="w")
        ctk.CTkLabel(totais, text="Total final c/ IVA:", font=("Segoe UI", 11, "bold"), text_color="#7a0f1a").grid(row=1, column=0, padx=10, pady=6, sticky="w")
        ctk.CTkLabel(totais, textvariable=total_final_var, font=("Segoe UI", 12, "bold"), text_color="#7a0f1a").grid(row=1, column=1, padx=6, pady=6, sticky="w")
        def _live_totals(*_):
            q = parse_float(v_qtd.get(), 0)
            pr = parse_float(v_preco.get(), 0)
            dp = max(0.0, min(100.0, parse_float(v_desc_perc.get(), 0)))
            ip = max(0.0, min(100.0, parse_float(v_iva_perc.get(), 23)))
            bruto = q * pr
            sem_iva = bruto * (1.0 - dp / 100.0)
            final = sem_iva * (1.0 + ip / 100.0)
            total_bruto_var.set(f"{fmt_num(bruto)} €")
            total_sem_iva_var.set(f"{fmt_num(sem_iva)} €")
            total_final_var.set(f"{fmt_num(final)} €")
        try:
            v_qtd.trace_add("write", _live_totals)
            v_preco.trace_add("write", _live_totals)
            v_desc_perc.trace_add("write", _live_totals)
            v_iva_perc.trace_add("write", _live_totals)
        except Exception:
            pass
        _live_totals()
        btns = ctk.CTkFrame(wrap, fg_color="#f7f8fb")
        btns.grid(row=7, column=0, columnspan=3, sticky="e", padx=12, pady=(10, 10))
        ctk.CTkButton(btns, text="Cancelar", width=120, command=dlg.destroy).pack(side="left", padx=(0, 8))
        ctk.CTkButton(btns, text="Guardar", width=150, command=on_ok, fg_color="#f59e0b", hover_color="#d97706").pack(side="left")
    else:
        ttk.Label(dlg, text=f"Item: {ref} ({origem_txt})").grid(row=0, column=0, columnspan=2, padx=8, pady=4, sticky="w")
        ttk.Label(dlg, text="Fornecedor").grid(row=1, column=0, padx=8, pady=4, sticky="w")
        cb = ttk.Combobox(dlg, textvariable=v_forn, values=forn_values, width=36)
        cb.grid(row=1, column=1, padx=8, pady=4)
        ttk.Label(dlg, text="Quantidade").grid(row=2, column=0, padx=8, pady=4, sticky="w")
        ttk.Entry(dlg, textvariable=v_qtd, width=15).grid(row=2, column=1, padx=8, pady=4)
        ttk.Label(dlg, text="Preço").grid(row=3, column=0, padx=8, pady=4, sticky="w")
        ttk.Entry(dlg, textvariable=v_preco, width=15).grid(row=3, column=1, padx=8, pady=4)
        ttk.Label(dlg, text="Desconto (%)").grid(row=4, column=0, padx=8, pady=4, sticky="w")
        ttk.Entry(dlg, textvariable=v_desc_perc, width=15).grid(row=4, column=1, padx=8, pady=4)
        ttk.Label(dlg, text="IVA (%)").grid(row=5, column=0, padx=8, pady=4, sticky="w")
        ttk.Entry(dlg, textvariable=v_iva_perc, width=15).grid(row=5, column=1, padx=8, pady=4)
        ttk.Button(dlg, text="Guardar", command=on_ok).grid(row=6, column=0, columnspan=2, pady=8)

def ne_del_linha(self):
    _ensure_configured()
    sel = self.tbl_ne_linhas.selection()
    for s in sel:
        self.tbl_ne_linhas.delete(s)
    _ne_meta_cleanup(self)
    self.ne_refresh_line_tags()
    self.refresh_ne_total()

def prompt_dados_guia(self, title, initial, linhas):
    _ensure_configured()
    use_custom = self.exp_use_custom and CUSTOM_TK_AVAILABLE
    Win = ctk.CTkToplevel if use_custom else Toplevel
    Frm = ctk.CTkFrame if use_custom else ttk.Frame
    Lbl = ctk.CTkLabel if use_custom else ttk.Label
    Ent = ctk.CTkEntry if use_custom else ttk.Entry
    Btn = ctk.CTkButton if use_custom else ttk.Button

    win = Win(self.root)
    win.title(title)
    try:
        if use_custom:
            win.geometry("980x760")
            win.configure(fg_color="#f7f8fb")
        else:
            win.geometry("940x720")
        win.transient(self.root)
        win.grab_set()
    except Exception:
        pass

    emit_cfg = get_guia_emitente_info()
    v_dest = StringVar(value=str(initial.get("destinatario", "") or ""))
    v_nif = StringVar(value=str(initial.get("dest_nif", "") or ""))
    v_mor = StringVar(value=str(initial.get("dest_morada", "") or ""))
    v_carga = StringVar(value=str(initial.get("local_carga", "") or emit_cfg.get("local_carga", "")))
    v_desc = StringVar(value=str(initial.get("local_descarga", "") or ""))
    v_data_t = StringVar(value=self._exp_datetime_ui(initial.get("data_transporte", now_iso())))
    v_transp = StringVar(value=str(initial.get("transportador", "") or ""))
    v_matr = StringVar(value=str(initial.get("matricula", "") or ""))
    v_obs = StringVar(value=str(initial.get("observacoes", "") or ""))
    v_codigo_at = StringVar(value=str(initial.get("codigo_at", "") or ""))
    v_emit_nome = StringVar(value=str(initial.get("emitente_nome", "") or emit_cfg.get("nome", "")))
    v_emit_nif = StringVar(value=str(initial.get("emitente_nif", "") or emit_cfg.get("nif", "")))
    v_emit_mor = StringVar(value=str(initial.get("emitente_morada", "") or emit_cfg.get("morada", "")))

    top = Frm(win, fg_color="#ffffff") if use_custom else Frm(win)
    top.pack(fill="x", padx=10, pady=(10, 6))
    Lbl(top, text="Cod. Validacao Serie AT").grid(row=0, column=0, sticky="w", padx=6, pady=4)
    Ent(top, textvariable=v_codigo_at, width=220 if use_custom else 24).grid(row=0, column=1, padx=6, pady=4, sticky="w")
    Lbl(top, text="Vias").grid(row=0, column=2, sticky="w", padx=6, pady=4)
    Lbl(top, text="Original / Duplicado / Triplicado (automatico)").grid(row=0, column=3, sticky="w", padx=6, pady=4)

    Lbl(top, text="Inicio transporte (YYYY-MM-DD HH:MM)").grid(row=1, column=0, sticky="w", padx=6, pady=4)
    Ent(top, textvariable=v_data_t, width=220 if use_custom else 24).grid(row=1, column=1, padx=6, pady=4, sticky="w")

    Lbl(top, text="Emitente (Nome)").grid(row=2, column=0, sticky="w", padx=6, pady=4)
    Ent(top, textvariable=v_emit_nome, width=280 if use_custom else 34).grid(row=2, column=1, padx=6, pady=4, sticky="w")
    Lbl(top, text="Emitente (NIF)").grid(row=2, column=2, sticky="w", padx=6, pady=4)
    Ent(top, textvariable=v_emit_nif, width=140 if use_custom else 16).grid(row=2, column=3, padx=6, pady=4, sticky="w")
    Lbl(top, text="Emitente (Morada)").grid(row=3, column=0, sticky="w", padx=6, pady=4)
    Ent(top, textvariable=v_emit_mor, width=520 if use_custom else 64).grid(row=3, column=1, columnspan=3, padx=6, pady=4, sticky="w")

    Lbl(top, text="Destinatario").grid(row=4, column=0, sticky="w", padx=6, pady=4)
    Ent(top, textvariable=v_dest, width=280 if use_custom else 34).grid(row=4, column=1, padx=6, pady=4, sticky="w")
    Lbl(top, text="NIF").grid(row=4, column=2, sticky="w", padx=6, pady=4)
    Ent(top, textvariable=v_nif, width=140 if use_custom else 16).grid(row=4, column=3, padx=6, pady=4, sticky="w")
    Lbl(top, text="Morada").grid(row=5, column=0, sticky="w", padx=6, pady=4)
    Ent(top, textvariable=v_mor, width=520 if use_custom else 64).grid(row=5, column=1, columnspan=3, padx=6, pady=4, sticky="w")
    Lbl(top, text="Local carga").grid(row=6, column=0, sticky="w", padx=6, pady=4)
    Ent(top, textvariable=v_carga, width=280 if use_custom else 34).grid(row=6, column=1, padx=6, pady=4, sticky="w")
    Lbl(top, text="Local descarga").grid(row=6, column=2, sticky="w", padx=6, pady=4)
    Ent(top, textvariable=v_desc, width=280 if use_custom else 34).grid(row=6, column=3, padx=6, pady=4, sticky="w")
    Lbl(top, text="Transportador").grid(row=7, column=0, sticky="w", padx=6, pady=4)
    Ent(top, textvariable=v_transp, width=280 if use_custom else 34).grid(row=7, column=1, padx=6, pady=4, sticky="w")
    Lbl(top, text="Matricula").grid(row=7, column=2, sticky="w", padx=6, pady=4)
    Ent(top, textvariable=v_matr, width=220 if use_custom else 24).grid(row=7, column=3, padx=6, pady=4, sticky="w")
    Lbl(top, text="Observacoes").grid(row=8, column=0, sticky="w", padx=6, pady=4)
    Ent(top, textvariable=v_obs, width=700 if use_custom else 86).grid(row=8, column=1, columnspan=3, padx=6, pady=4, sticky="w")

    box = Frm(win, fg_color="#ffffff") if use_custom else Frm(win)
    box.pack(fill="both", expand=True, padx=10, pady=(0, 8))
    Lbl(box, text="Linhas da Guia").pack(anchor="w", padx=6, pady=(6, 2))
    tbl = ttk.Treeview(
        box,
        columns=("ref_int", "ref_ext", "desc", "qtd", "unid"),
        show="headings",
        height=10,
        style=("EXP.Treeview" if self.exp_use_custom else ""),
    )
    tbl.heading("ref_int", text="Ref. Interna")
    tbl.heading("ref_ext", text="Ref. Externa")
    tbl.heading("desc", text="Descricao")
    tbl.heading("qtd", text="Qtd")
    tbl.heading("unid", text="Unid")
    tbl.column("ref_int", width=170, anchor="w")
    tbl.column("ref_ext", width=170, anchor="w")
    tbl.column("desc", width=420, anchor="w")
    tbl.column("qtd", width=100, anchor="e")
    tbl.column("unid", width=80, anchor="center")
    tbl.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=(0, 6))
    sb = ttk.Scrollbar(box, orient="vertical", command=tbl.yview)
    tbl.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y", padx=(0, 6), pady=(0, 6))

    for l in linhas:
        tbl.insert(
            "",
            END,
            values=(
                l.get("ref_interna", ""),
                l.get("ref_externa", ""),
                l.get("descricao", ""),
                fmt_num(l.get("qtd", 0)),
                l.get("unid", "UN"),
            ),
        )

    out = {"data": None}

    def on_ok():
        if not v_dest.get().strip():
            messagebox.showerror("Erro", "Indique o destinatario.")
            return
        out["data"] = {
            "codigo_at": v_codigo_at.get().strip(),
            "tipo_via": "Original",
            "emitente_nome": v_emit_nome.get().strip(),
            "emitente_nif": v_emit_nif.get().strip(),
            "emitente_morada": v_emit_mor.get().strip(),
            "destinatario": v_dest.get().strip(),
            "dest_nif": v_nif.get().strip(),
            "dest_morada": v_mor.get().strip(),
            "local_carga": v_carga.get().strip(),
            "local_descarga": v_desc.get().strip(),
            "data_transporte": self._exp_parse_datetime(v_data_t.get().strip(), now_iso()),
            "transportador": v_transp.get().strip(),
            "matricula": v_matr.get().strip(),
            "observacoes": v_obs.get().strip(),
        }
        win.destroy()

    foot = Frm(win, fg_color="transparent") if use_custom else Frm(win)
    foot.pack(fill="x", padx=10, pady=(0, 10))
    Btn(foot, text="Confirmar Emissao", command=on_ok, width=160 if use_custom else None).pack(side="left", padx=4)
    Btn(foot, text="Cancelar", command=win.destroy, width=120 if use_custom else None).pack(side="right", padx=4)
    self.root.wait_window(win)
    return out.get("data")
def render_expedicao_pdf(self, path, ex, include_all_vias=True):
    _ensure_configured()
    return _render_expedicao_pdf_modern(self, path, ex, include_all_vias=include_all_vias)
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as pdf_canvas
    from reportlab.lib import colors
    from reportlab.lib.utils import ImageReader

    c = pdf_canvas.Canvas(path, pagesize=A4)
    w, h = A4
    m = 26
    content_w = w - (2 * m)
    line_h = 16
    max_table_y = 612
    x_art = m + 152
    x_desc = w - m - 94
    x_qtd = w - m - 42
    first_block_y = 126
    blue_main = colors.HexColor("#1F4E8C")
    blue_mid = colors.HexColor("#5D7FA9")
    blue_soft = colors.HexColor("#C8D9EE")
    fill_head = colors.HexColor("#EAF1FB")
    line_frame = blue_mid
    line_soft = blue_soft
    c.setLineJoin(1)
    c.setLineCap(1)

    def yinv(y):
        return h - y

    ex_num = str(ex.get("numero", "") or "-").strip() or "-"
    emit_cfg = get_guia_emitente_info()
    emit_nome = str(ex.get("emitente_nome", "") or emit_cfg.get("nome", "")).strip()
    emit_nif = str(ex.get("emitente_nif", "") or emit_cfg.get("nif", "")).strip()
    emit_morada = str(ex.get("emitente_morada", "") or emit_cfg.get("morada", "")).strip()
    dest_nome = str(ex.get("destinatario", "") or "").strip()
    dest_nif = str(ex.get("dest_nif", "") or "").strip()
    dest_morada = str(ex.get("dest_morada", "") or "").strip()
    local_carga = str(ex.get("local_carga", "") or "").strip()
    local_descarga = str(ex.get("local_descarga", "") or "").strip()
    transportador = str(ex.get("transportador", "") or "").strip()
    matricula = str(ex.get("matricula", "") or "").strip()
    observacoes = str(ex.get("observacoes", "") or "").strip()
    created_by = str(ex.get("created_by", "") or "").strip()
    data_emissao = str(ex.get("data_emissao", "") or "").strip()
    data_transporte = str(ex.get("data_transporte", "") or "").strip()
    codigo_at = str(ex.get("codigo_at", "") or "").strip()
    at_validation_code = str(ex.get("at_validation_code", "") or codigo_at).strip()
    seq_num = int(parse_float(ex.get("seq_num", 0), 0) or 0)
    atcud = str(ex.get("atcud", "") or "").strip()
    if not atcud and at_validation_code and seq_num > 0:
        atcud = f"{at_validation_code}-{seq_num}"
    if not atcud and codigo_at:
        atcud = codigo_at
    if not codigo_at and at_validation_code:
        codigo_at = at_validation_code
    guia_extra_lines = get_guia_extra_info_lines()
    via_options = ("Original", "Duplicado", "Triplicado")
    tipo_via = str(ex.get("tipo_via", "Original") or "Original").strip().title()
    if tipo_via not in via_options:
        tipo_via = "Original"
    linhas = list(ex.get("linhas", []) or [])

    def _ntxt(value):
        try:
            return pdf_normalize_text(value)
        except Exception:
            return str(value or "")

    def _fmt_date(raw):
        txt = str(raw or "").replace("T", " ").strip()
        return txt[:10] if len(txt) >= 10 else (txt or "-")

    def _fmt_datetime(raw):
        txt = str(raw or "").replace("T", " ").strip()
        return txt[:16] if len(txt) >= 16 else (txt or "-")

    def _wrap_text(txt, max_chars=48, max_lines=None):
        base = str(txt or "").replace("\n", " ").strip()
        if not base:
            return []
        words = base.split()
        out = []
        cur = ""
        for wtxt in words:
            candidate = f"{cur} {wtxt}".strip() if cur else wtxt
            if len(candidate) <= max_chars:
                cur = candidate
            else:
                if cur:
                    out.append(cur)
                cur = wtxt
        if cur:
            out.append(cur)
        if max_lines is not None:
            return out[:max_lines]
        return out

    def _first_header_layout():
        left_lines = []
        left_lines.append(emit_nome or "-")
        left_lines.append(f"Contribuinte N.: {emit_nif or '-'}")
        left_lines.extend(_wrap_text(emit_morada, 44, 3))
        for ln in guia_extra_lines[:4]:
            left_lines.extend(_wrap_text(ln, 44, 1))

        right_lines = [dest_nome or "-", *_wrap_text(dest_morada, 42, 4)]
        if dest_nif:
            right_lines.append(f"NIF: {dest_nif}")

        left_box_x = m
        left_box_w = content_w * 0.47
        right_box_w = content_w * 0.43
        right_box_x = w - m - right_box_w
        line_step = 12
        left_box_h = 30 + (min(len(left_lines), 10) * line_step)
        right_box_h = 30 + (min(len(right_lines), 8) * line_step)
        info_h = max(left_box_h, right_box_h)
        y_title = first_block_y + info_h + 18
        y_meta = y_title + 50
        return {
            "left_lines": left_lines,
            "right_lines": right_lines,
            "left_box_x": left_box_x,
            "left_box_w": left_box_w,
            "left_box_h": left_box_h,
            "right_box_x": right_box_x,
            "right_box_w": right_box_w,
            "right_box_h": right_box_h,
            "y_title": y_title,
            "y_meta": y_meta,
        }

    def _calc_first_table_y():
        layout = _first_header_layout()
        return layout["y_meta"] + 48

    def _calc_next_table_y():
        return 80

    first_table_y = _calc_first_table_y()
    next_table_y = _calc_next_table_y()
    first_cap = max(1, int((max_table_y - first_table_y) // line_h))
    next_cap = max(1, int((max_table_y - next_table_y) // line_h))
    total_rows = len(linhas)
    if total_rows <= first_cap:
        total_pages = 1
    else:
        rem = total_rows - first_cap
        total_pages = 1 + ((rem + next_cap - 1) // next_cap)

    def _page_txt(page):
        return f"Pag. {page}/{total_pages}"

    def draw_qr_accud(x, top_y, size=86):
        payload = f"ATCUD:{atcud or '-'}|DOC:{ex_num}|DT:{_fmt_date(data_emissao)}|NIF:{emit_nif or '-'}"
        try:
            from reportlab.graphics.barcode import qr as rl_qr
            from reportlab.graphics.shapes import Drawing
            from reportlab.graphics import renderPDF

            widget = rl_qr.QrCodeWidget(payload)
            b = widget.getBounds()
            bw = max(float(b[2] - b[0]), 1.0)
            bh = max(float(b[3] - b[1]), 1.0)
            d = Drawing(size, size, transform=[size / bw, 0, 0, size / bh, 0, 0])
            d.add(widget)
            y_bottom = yinv(top_y + size)
            renderPDF.draw(d, c, x, y_bottom)
            c.saveState()
            c.setLineWidth(0.6)
            c.setStrokeColor(colors.HexColor("#333333"))
            c.rect(x - 2, y_bottom - 2, size + 4, size + 4, stroke=1, fill=0)
            c.setFont("Helvetica", 6.3)
            c.setFillColor(colors.black)
            c.drawString(x, yinv(top_y - 4), _ntxt("QR ATCUD"))
            c.drawString(x, yinv(top_y + size + 10), _ntxt(f"ATCUD: {atcud or '-'}"))
            c.restoreState()
        except Exception:
            c.saveState()
            c.setFont("Helvetica", 6.5)
            c.setFillColor(colors.black)
            c.drawString(x, yinv(top_y + 8), _ntxt(f"ATCUD: {atcud[:28]}"))
            c.restoreState()

    def draw_logo(x, y, max_w=150, max_h=52):
        lp = get_orc_logo_path()
        if not lp or not os.path.exists(lp):
            return
        try:
            img = ImageReader(lp)
            c.drawImage(img, x, yinv(y + max_h), width=max_w, height=max_h, preserveAspectRatio=True, mask="auto")
        except Exception:
            pass

    def draw_a4_frame():
        c.saveState()
        outer = 10
        c.setStrokeColor(line_frame)
        c.setLineWidth(0.9)
        c.rect(outer, outer, w - (2 * outer), h - (2 * outer), stroke=1, fill=0)
        c.setStrokeColor(line_soft)
        c.setLineWidth(0.45)
        inner_x = m - 8
        inner_y = 18
        c.rect(inner_x, inner_y, w - (2 * inner_x), h - (2 * inner_y), stroke=1, fill=0)
        c.restoreState()

    def draw_table_header(y_top):
        c.saveState()
        c.setFillColor(fill_head)
        c.setStrokeColor(line_frame)
        c.setLineWidth(0.7)
        c.rect(m, yinv(y_top + 18), w - (2 * m), 18, stroke=1, fill=1)
        c.restoreState()

        c.setStrokeColor(line_frame)
        c.setLineWidth(0.95)
        c.line(m, yinv(y_top), w - m, yinv(y_top))
        c.setFont("Helvetica-Bold", 9.8)
        c.setFillColor(blue_main)
        c.drawString(m + 3, yinv(y_top + 13), "Artigo")
        c.drawString(x_art + 3, yinv(y_top + 13), "Descricao")
        c.drawRightString(x_qtd - 4, yinv(y_top + 13), "Qtd.")
        c.drawRightString(w - m - 5, yinv(y_top + 13), "Un.")
        c.setFillColor(colors.black)
        c.setLineWidth(0.7)
        c.line(m, yinv(y_top + 18), w - m, yinv(y_top + 18))
        c.line(x_art, yinv(y_top), x_art, yinv(y_top + 18))
        c.line(x_desc, yinv(y_top), x_desc, yinv(y_top + 18))
        c.line(x_qtd, yinv(y_top), x_qtd, yinv(y_top + 18))
        c.setStrokeColor(colors.black)
        return y_top + 22

    def draw_table_frame(y_top, y_bottom):
        c.setStrokeColor(line_frame)
        c.setLineWidth(0.55)
        c.line(m, yinv(y_top), m, yinv(y_bottom))
        c.line(x_art, yinv(y_top), x_art, yinv(y_bottom))
        c.line(x_desc, yinv(y_top), x_desc, yinv(y_bottom))
        c.line(x_qtd, yinv(y_top), x_qtd, yinv(y_bottom))
        c.line(w - m, yinv(y_top), w - m, yinv(y_bottom))
        c.line(m, yinv(y_bottom), w - m, yinv(y_bottom))
        c.setStrokeColor(colors.black)

    def draw_first_page_header(page_no):
        draw_a4_frame()
        c.setFont("Helvetica", 9.1)
        c.drawRightString(w - m, yinv(24), _ntxt(_page_txt(page_no)))
        draw_logo(m + 2, 14, 172, 102)

        layout = _first_header_layout()
        y_block = first_block_y
        left_x = layout["left_box_x"]
        left_w = layout["left_box_w"]
        left_h = layout["left_box_h"]
        right_x = layout["right_box_x"]
        right_w = layout["right_box_w"]
        right_h = layout["right_box_h"]
        left_lines = layout["left_lines"]
        right_lines = layout["right_lines"]

        c.saveState()
        c.setStrokeColor(line_frame)
        c.setLineWidth(0.95)
        c.roundRect(left_x, yinv(y_block + left_h), left_w, left_h, 6, stroke=1, fill=0)
        c.roundRect(right_x, yinv(y_block + right_h), right_w, right_h, 6, stroke=1, fill=0)
        c.setFillColor(fill_head)
        c.rect(left_x + 1, yinv(y_block + 20), left_w - 2, 18, stroke=0, fill=1)
        c.rect(right_x + 1, yinv(y_block + 20), right_w - 2, 18, stroke=0, fill=1)
        c.restoreState()

        c.setFillColor(blue_main)
        c.setFont("Helvetica-Bold", 9.5)
        c.drawString(left_x + 6, yinv(y_block + 13), "Remetente")
        c.drawString(right_x + 6, yinv(y_block + 13), "Destinatario")
        c.setFillColor(colors.black)

        y_left = y_block + 32
        for idx_ln, ln in enumerate(left_lines[:10]):
            c.setFont("Helvetica-Bold" if idx_ln == 0 else "Helvetica", 10 if idx_ln == 0 else 9.5)
            c.drawString(left_x + 6, yinv(y_left), _ntxt(ln))
            y_left += 12

        y_right = y_block + 32
        for idx_ln, ln in enumerate(right_lines[:8]):
            c.setFont("Helvetica-Bold" if idx_ln == 0 else "Helvetica", 10 if idx_ln == 0 else 9.5)
            c.drawString(right_x + 6, yinv(y_right), _ntxt(ln))
            y_right += 12

        y_title = layout["y_title"]
        c.setFillColor(blue_main)
        c.setFont("Helvetica-Bold", 18)
        c.drawString(m, yinv(y_title), _ntxt(f"Guia de transporte {ex_num}"))
        c.setFont("Helvetica-Bold", 11)
        c.drawRightString(w - m, yinv(y_title + 1), _ntxt(tipo_via))
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(m, yinv(y_title + 18), "Guia de transporte")
        c.drawString(m, yinv(y_title + 32), _ntxt(f"Chave AT: {at_validation_code or '-'}"))
        c.setLineWidth(1.1)
        c.setStrokeColor(blue_main)
        c.line(m, yinv(y_title + 40), w - m, yinv(y_title + 40))
        c.setStrokeColor(colors.black)

        meta_y = layout["y_meta"]
        c.setFont("Helvetica-Bold", 9.8)
        c.drawString(m + 2, yinv(meta_y), "V/N.o Contrib.")
        c.drawString(m + 176, yinv(meta_y), "Requisicao")
        c.drawString(w - m - 94, yinv(meta_y), "Data")
        c.setFont("Helvetica", 9.8)
        requisicao = str(ex.get("encomenda", "") or "").strip() or str(ex.get("observacoes", "") or "").strip()
        c.drawString(m + 2, yinv(meta_y + 14), _ntxt(dest_nif or "-"))
        c.drawString(m + 176, yinv(meta_y + 14), _ntxt((requisicao[:50] if requisicao else "-")))
        c.drawString(w - m - 94, yinv(meta_y + 14), _ntxt(_fmt_date(data_emissao or data_transporte)))
        c.setStrokeColor(line_frame)
        c.setLineWidth(0.8)
        c.line(m, yinv(meta_y + 20), w - m, yinv(meta_y + 20))
        c.setStrokeColor(colors.black)
        return draw_table_header(meta_y + 26)

    def draw_next_page_header(page_no):
        draw_a4_frame()
        c.setFont("Helvetica", 9.1)
        c.drawRightString(w - m, yinv(24), _ntxt(_page_txt(page_no)))
        c.setFillColor(blue_main)
        c.setFont("Helvetica-Bold", 14)
        c.drawString(m, yinv(44), _ntxt(f"Guia de transporte {ex_num}"))
        c.setFillColor(colors.black)
        c.setFont("Helvetica", 9)
        c.drawRightString(w - m, yinv(44), _ntxt(tipo_via))
        c.setFont("Helvetica", 8.6)
        c.drawRightString(w - m, yinv(56), _ntxt(f"ATCUD: {atcud or '-'}"))
        c.setLineWidth(0.9)
        c.setStrokeColor(blue_main)
        c.line(m, yinv(50), w - m, yinv(50))
        c.setStrokeColor(colors.black)
        return draw_table_header(58)

    def draw_final_footer(page_no):
        c.setFont("Helvetica", 8.9)
        c.drawString(m, yinv(655), "Este documento nao serve de fatura")
        c.setFont("Helvetica", 8.2)
        c.drawString(
            m,
            yinv(670),
            _ntxt(f"c/Ho-Processado por Programa Certificado n.o 0030/AT / {ex_num}"),
        )
        c.drawString(
            m,
            yinv(682),
            _ntxt(
                f"Os bens e/ou servicos foram colocados a disposicao na data {_fmt_date(data_transporte or data_emissao)}"
            ),
        )
        c.drawRightString(w - m, yinv(682), _ntxt(f"ATCUD: {atcud or '-'}"))

        c.setLineWidth(0.9)
        c.setStrokeColor(line_frame)
        c.line(m, yinv(694), w - m, yinv(694))
        c.setStrokeColor(colors.black)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(m, yinv(708), "Carga")
        c.drawString(m + (content_w * 0.38), yinv(708), "Descarga")

        c.setFont("Helvetica", 9.2)
        carga_lines = [f"N/ Morada - {_fmt_datetime(data_transporte or data_emissao)}"]
        carga_lines.extend(_wrap_text(local_carga, 38, 5))
        if transportador:
            carga_lines.append(f"Transportador: {transportador}")
        if matricula:
            carga_lines.append(f"Matricula: {matricula}")

        descarga_lines = ["V/ Morada"]
        descarga_lines.extend(_wrap_text(local_descarga, 38, 6))
        if dest_nome:
            descarga_lines.insert(1, dest_nome)

        y_line = 724
        for ln in carga_lines[:8]:
            c.drawString(m, yinv(y_line), _ntxt(ln))
            y_line += 14
        y_line = 724
        right_col_x = m + (content_w * 0.38)
        for ln in descarga_lines[:8]:
            c.drawString(right_col_x, yinv(y_line), _ntxt(ln))
            y_line += 14

        if created_by:
            c.setFont("Helvetica", 7.8)
            c.drawString(m, yinv(822), _ntxt(f"Emitida por: {created_by}"))

        draw_qr_accud(w - m - 112, 716, size=92)
        c.setFont("Helvetica", 9)
        c.drawRightString(w - m, yinv(24), _ntxt(_page_txt(page_no)))

    def draw_row(y_top, linha):
        artigo = str(
            linha.get("produto_codigo")
            or linha.get("ref_externa")
            or linha.get("ref_interna")
            or "-"
        ).strip()
        desc = str(linha.get("descricao", "") or "").strip()
        if not desc:
            desc = f"{str(linha.get('ref_interna', '') or '').strip()} {str(linha.get('ref_externa', '') or '').strip()}".strip() or "-"
        qtd_txt = fmt_num(linha.get("qtd", 0))
        un_txt = str(linha.get("unid", "UN") or "UN").strip()[:6]

        c.setFont("Helvetica", 9.6)
        c.drawString(m + 3, yinv(y_top + 11), _ntxt(artigo[:24]))
        c.drawString(x_art + 3, yinv(y_top + 11), _ntxt(desc[:72]))
        c.drawRightString(x_qtd - 4, yinv(y_top + 11), _ntxt(qtd_txt))
        c.drawRightString(w - m - 5, yinv(y_top + 11), _ntxt(un_txt))
        c.setLineWidth(0.32)
        c.setStrokeColor(line_soft)
        c.line(m, yinv(y_top + 14), w - m, yinv(y_top + 14))
        c.setStrokeColor(colors.black)

    vias_to_render = list(via_options)
    for via_idx, via_nome in enumerate(vias_to_render):
        tipo_via = via_nome
        page_no = 1
        idx = 0
        first_page = True
        while True:
            table_y = draw_first_page_header(page_no) if first_page else draw_next_page_header(page_no)
            first_page = False
            y = table_y

            if not linhas:
                c.setFont("Helvetica-Oblique", 9)
                c.drawString(m + 4, yinv(y + 12), "Sem linhas na guia.")
                y += 18

            while idx < len(linhas):
                if y + line_h > max_table_y:
                    break
                draw_row(y, linhas[idx])
                y += line_h
                idx += 1

            draw_table_frame(table_y - 22, max_table_y)

            if idx < len(linhas):
                c.showPage()
                page_no += 1
                continue

            if observacoes:
                c.setFont("Helvetica", 8.4)
                c.drawString(m, yinv(640), _ntxt(f"Observacoes: {observacoes[:150]}"))
            draw_final_footer(page_no)
            break

        if ex.get("anulada"):
            c.saveState()
            c.setFillColor(colors.HexColor("#b91c1c"))
            c.setFont("Helvetica-Bold", 11)
            c.drawString(m, yinv(h - 44), _ntxt(f"DOCUMENTO ANULADO - Motivo: {ex.get('anulada_motivo', '')}"))
            c.restoreState()

        if via_idx < len(vias_to_render) - 1:
            c.showPage()

    c.save()


def _render_expedicao_pdf_modern(self, path, ex, include_all_vias=True):
    _ensure_configured()
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as pdf_canvas

    fonts = _pdf_register_fonts()
    palette = _pdf_brand_palette()
    c = pdf_canvas.Canvas(path, pagesize=A4)
    w, h = A4
    m = 24
    content_w = w - (2 * m)
    row_h = 18
    header_row_h = 20
    code_w = 112
    qty_w = 58
    un_w = 42
    desc_w = content_w - code_w - qty_w - un_w
    x_code = m
    x_desc = x_code + code_w
    x_qty = x_desc + desc_w
    x_un = x_qty + qty_w
    footer_top = 646
    header_first_top = 334
    header_next_top = 92
    via_options = ("Original", "Duplicado", "Triplicado")

    ex_num = str(ex.get("numero", "") or "-").strip() or "-"
    emit_cfg = get_guia_emitente_info()
    emit_nome = str(ex.get("emitente_nome", "") or emit_cfg.get("nome", "")).strip()
    emit_nif = str(ex.get("emitente_nif", "") or emit_cfg.get("nif", "")).strip()
    emit_morada = str(ex.get("emitente_morada", "") or emit_cfg.get("morada", "")).strip()
    dest_nome = str(ex.get("destinatario", "") or "").strip()
    dest_nif = str(ex.get("dest_nif", "") or "").strip()
    dest_morada = str(ex.get("dest_morada", "") or "").strip()
    local_carga = str(ex.get("local_carga", "") or "").strip()
    local_descarga = str(ex.get("local_descarga", "") or "").strip()
    transportador = str(ex.get("transportador", "") or "").strip()
    matricula = str(ex.get("matricula", "") or "").strip()
    observacoes = str(ex.get("observacoes", "") or "").strip()
    created_by = str(ex.get("created_by", "") or "").strip()
    data_emissao = str(ex.get("data_emissao", "") or "").strip()
    data_transporte = str(ex.get("data_transporte", "") or "").strip()
    codigo_at = str(ex.get("codigo_at", "") or "").strip()
    at_validation_code = str(ex.get("at_validation_code", "") or codigo_at).strip()
    seq_num = int(parse_float(ex.get("seq_num", 0), 0) or 0)
    atcud = str(ex.get("atcud", "") or "").strip()
    if not atcud and at_validation_code and seq_num > 0:
        atcud = f"{at_validation_code}-{seq_num}"
    if not atcud and codigo_at:
        atcud = codigo_at
    if not codigo_at and at_validation_code:
        codigo_at = at_validation_code
    guia_extra_lines = get_guia_extra_info_lines()
    tipo_via = str(ex.get("tipo_via", "Original") or "Original").strip().title()
    if tipo_via not in via_options:
        tipo_via = "Original"
    linhas = list(ex.get("linhas", []) or [])
    vias_to_render = list(via_options) if include_all_vias else [tipo_via]
    first_capacity = max(1, int((footer_top - (header_first_top + header_row_h + 6)) // row_h))
    next_capacity = max(1, int((footer_top - (header_next_top + header_row_h + 6)) // row_h))

    def yinv(top_y):
        return h - top_y

    def ntxt(value):
        try:
            return pdf_normalize_text(value)
        except Exception:
            return str(value or "")

    def fmt_date(value):
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

    def fmt_datetime(value):
        raw = str(value or "").strip().replace("T", " ")
        if not raw:
            return "-"
        try:
            dt = datetime.fromisoformat(raw)
            return dt.strftime("%d/%m/%Y %H:%M")
        except Exception:
            if len(raw) >= 16 and raw[4] == "-" and raw[7] == "-":
                return f"{raw[8:10]}/{raw[5:7]}/{raw[:4]} {raw[11:16]}"
            return raw[:16]

    def card(x, top_y, box_w, box_h, title, lines, title_fill=None, accent=False):
        fill = title_fill or palette["primary_soft_2"]
        radius = 8
        c.saveState()
        c.setStrokeColor(palette["line_strong"] if accent else palette["line"])
        c.setLineWidth(0.9)
        c.roundRect(x, yinv(top_y + box_h), box_w, box_h, radius, stroke=1, fill=0)
        c.setFillColor(fill)
        c.roundRect(x, yinv(top_y + 20), box_w, 20, radius, stroke=0, fill=1)
        c.restoreState()
        c.setFillColor(palette["primary_dark"])
        c.setFont(fonts["bold"], 8.8)
        c.drawString(x + 8, yinv(top_y + 13), ntxt(title))
        yy = top_y + 34
        for idx_line, line in enumerate(lines):
            font_name = fonts["bold"] if idx_line == 0 and accent else fonts["regular"]
            font_size = 9.5 if idx_line == 0 and accent else 8.9
            wrapped = _pdf_wrap_text(line, font_name, font_size, box_w - 16, max_lines=2 if idx_line == 0 else 1)
            for item in wrapped:
                c.setFont(font_name, font_size)
                c.setFillColor(palette["ink"])
                c.drawString(x + 8, yinv(yy), ntxt(item))
                yy += 11
            if yy > top_y + box_h - 10:
                break

    def metric_chip(x, top_y, box_w, label, value, box_h=34):
        c.saveState()
        c.setFillColor(colors.white)
        c.setStrokeColor(palette["line"])
        c.setLineWidth(0.85)
        c.roundRect(x, yinv(top_y + box_h), box_w, box_h, 8, stroke=1, fill=1)
        c.restoreState()
        c.setFont(fonts["regular"], 7.3)
        c.setFillColor(palette["muted"])
        c.drawString(x + 8, yinv(top_y + 11), ntxt(label))
        c.setFont(fonts["bold"], 10)
        c.setFillColor(palette["primary_dark"])
        c.drawString(x + 8, yinv(top_y + 24), ntxt(_pdf_clip_text(value, box_w - 16, fonts["bold"], 10)))

    def draw_qr_block(x, top_y, size=74):
        payload = f"ATCUD:{atcud or '-'}|DOC:{ex_num}|DT:{fmt_date(data_emissao)}|NIF:{emit_nif or '-'}"
        c.saveState()
        c.setFillColor(colors.white)
        c.setStrokeColor(palette["line"])
        c.roundRect(x, yinv(top_y + size + 22), size + 18, size + 22, 8, stroke=1, fill=1)
        c.restoreState()
        try:
            from reportlab.graphics import renderPDF
            from reportlab.graphics.barcode import qr as rl_qr
            from reportlab.graphics.shapes import Drawing

            widget = rl_qr.QrCodeWidget(payload)
            bounds = widget.getBounds()
            bw = max(float(bounds[2] - bounds[0]), 1.0)
            bh = max(float(bounds[3] - bounds[1]), 1.0)
            drawing = Drawing(size, size, transform=[size / bw, 0, 0, size / bh, 0, 0])
            drawing.add(widget)
            renderPDF.draw(drawing, c, x + 9, yinv(top_y + size + 9))
        except Exception:
            c.setStrokeColor(palette["line"])
            c.rect(x + 9, yinv(top_y + size + 9), size, size, stroke=1, fill=0)
        c.setFont(fonts["bold"], 7.5)
        c.setFillColor(palette["muted"])
        c.drawCentredString(x + ((size + 18) / 2.0), yinv(top_y + size + 15), ntxt("QR / ATCUD"))

    def draw_table_header(top_y):
        c.saveState()
        c.setFillColor(palette["primary"])
        c.roundRect(m, yinv(top_y + header_row_h), content_w, header_row_h, 7, stroke=0, fill=1)
        c.restoreState()
        c.setFont(fonts["bold"], 8.6)
        c.setFillColor(colors.white)
        c.drawString(x_code + 8, yinv(top_y + 13), ntxt("Referencia"))
        c.drawString(x_desc + 8, yinv(top_y + 13), ntxt("Descricao"))
        c.drawRightString(x_qty + qty_w - 8, yinv(top_y + 13), ntxt("Qtd."))
        c.drawRightString(x_un + un_w - 8, yinv(top_y + 13), ntxt("Un."))
        return top_y + header_row_h + 6

    def draw_row(top_y, idx_line, linha):
        artigo = str(
            linha.get("produto_codigo")
            or linha.get("ref_externa")
            or linha.get("ref_interna")
            or "-"
        ).strip() or "-"
        desc = str(linha.get("descricao", "") or "").strip()
        if not desc:
            desc = f"{str(linha.get('ref_interna', '') or '').strip()} {str(linha.get('ref_externa', '') or '').strip()}".strip() or "-"
        qtd_txt = fmt_num(linha.get("qtd", 0))
        un_txt = str(linha.get("unid", "UN") or "UN").strip()[:10]
        bg = palette["surface_warm"] if idx_line % 2 == 0 else colors.white
        c.saveState()
        c.setFillColor(bg)
        c.setStrokeColor(palette["line"])
        c.setLineWidth(0.4)
        c.roundRect(m, yinv(top_y + row_h), content_w, row_h, 5, stroke=1, fill=1)
        c.restoreState()
        c.setStrokeColor(palette["line"])
        c.setLineWidth(0.4)
        c.line(x_desc, yinv(top_y), x_desc, yinv(top_y + row_h))
        c.line(x_qty, yinv(top_y), x_qty, yinv(top_y + row_h))
        c.line(x_un, yinv(top_y), x_un, yinv(top_y + row_h))
        c.setFillColor(palette["ink"])
        c.setFont(fonts["regular"], 8.7)
        c.drawString(x_code + 8, yinv(top_y + 11.5), ntxt(_pdf_clip_text(artigo, code_w - 16, fonts["regular"], 8.7)))
        c.drawString(x_desc + 8, yinv(top_y + 11.5), ntxt(_pdf_clip_text(desc, desc_w - 16, fonts["regular"], 8.7)))
        c.drawRightString(x_qty + qty_w - 8, yinv(top_y + 11.5), ntxt(qtd_txt))
        c.drawRightString(x_un + un_w - 8, yinv(top_y + 11.5), ntxt(un_txt))

    def draw_cancel_stamp():
        if not ex.get("anulada"):
            return
        c.saveState()
        c.setStrokeColor(palette["danger"])
        c.setFillColor(palette["danger"])
        c.setLineWidth(1.2)
        c.translate(w / 2.0, h / 2.0)
        c.rotate(23)
        c.setFont(fonts["bold"], 26)
        c.drawCentredString(0, 0, ntxt("DOCUMENTO ANULADO"))
        motivo = str(ex.get("anulada_motivo", "") or "").strip()
        if motivo:
            c.setFont(fonts["regular"], 12)
            c.drawCentredString(0, -18, ntxt(_pdf_clip_text(motivo, 280, fonts["regular"], 12)))
        c.restoreState()

    def draw_first_header(page_no, total_pages, via_nome):
        c.saveState()
        c.setFillColor(palette["primary"])
        c.roundRect(m, yinv(98), content_w, 80, 12, stroke=0, fill=1)
        c.restoreState()
        logo_x = m + 14
        logo_w = 116
        metric_grid = _pdf_metric_grid_layout(w - m, 18, 80, group_w=230, gap=6)
        draw_pdf_logo_box(c, h, logo_x, 28, box_size=logo_w, box_h=50, padding=4, draw_border=False)
        c.setFillColor(colors.white)
        _pdf_draw_banner_title_block(
            c,
            h,
            "Guia de Transporte",
            emit_nome or "Emitente",
            logo_x + logo_w + 8,
            metric_grid["col1"] - 10,
            46,
            64,
            fonts["bold"],
            fonts["regular"],
            title_size=17.0,
            subtitle_size=8.8,
            min_title_size=14.3,
            min_subtitle_size=7.4,
        )
        metric_chip(metric_grid["col1"], metric_grid["row1"], metric_grid["chip_w"], "Numero", ex_num, box_h=metric_grid["chip_h"])
        metric_chip(metric_grid["col2"], metric_grid["row1"], metric_grid["chip_w"], "Pagina", f"{page_no}/{total_pages}", box_h=metric_grid["chip_h"])
        metric_chip(metric_grid["col1"], metric_grid["row2"], metric_grid["chip_w"], "Emissao", fmt_date(data_emissao or data_transporte), box_h=metric_grid["chip_h"])
        metric_chip(metric_grid["col2"], metric_grid["row2"], metric_grid["chip_w"], "Via", via_nome, box_h=metric_grid["chip_h"])

        left_w = (content_w * 0.52) - 6
        right_w = content_w - left_w - 12
        card(
            m,
            124,
            left_w,
            108,
            "Emitente",
            [
                emit_nome or "-",
                f"NIF: {emit_nif or '-'}",
                emit_morada or "-",
                *guia_extra_lines[:2],
            ],
            accent=True,
        )
        card(
            m + left_w + 12,
            124,
            right_w,
            108,
            "Destinatario",
            [
                dest_nome or "-",
                f"NIF: {dest_nif or '-'}" if dest_nif else "NIF: -",
                dest_morada or "-",
            ],
            accent=True,
        )

        card(m, 246, (content_w / 3.0) - 8, 72, "Local de carga", [local_carga or "-", fmt_datetime(data_transporte or data_emissao)])
        card(m + (content_w / 3.0), 246, (content_w / 3.0) - 8, 72, "Local de descarga", [local_descarga or "-", dest_nome or "-"])
        card(
            m + (content_w * (2.0 / 3.0)),
            246,
            (content_w / 3.0),
            72,
            "Transporte",
            [
                f"Transportador: {transportador or '-'}",
                f"Matricula: {matricula or '-'}",
                f"ATCUD: {atcud or '-'}",
            ],
        )
        return draw_table_header(header_first_top)

    def draw_next_header(page_no, total_pages, via_nome):
        c.saveState()
        c.setFillColor(palette["primary_soft"])
        c.roundRect(m, yinv(70), content_w, 52, 12, stroke=0, fill=1)
        c.restoreState()
        c.setFillColor(palette["primary_dark"])
        c.setFont(fonts["bold"], 15)
        c.drawString(m + 12, yinv(42), ntxt(f"Guia de Transporte {ex_num}"))
        c.setFont(fonts["regular"], 8.8)
        c.drawString(m + 12, yinv(56), ntxt(f"Destinatario: {dest_nome or '-'}"))
        metric_chip(w - m - 178, 20, 84, "Pagina", f"{page_no}/{total_pages}")
        metric_chip(w - m - 88, 20, 88, "Via", via_nome)
        return draw_table_header(header_next_top)

    def draw_footer(page_no, total_pages):
        obs_w = content_w - 112
        if observacoes:
            card(m, 660, obs_w, 64, "Observacoes", [observacoes])
        else:
            card(
                m,
                660,
                obs_w,
                64,
                "Observacoes",
                ["Sem observacoes adicionais para esta expedicao."],
            )
        draw_qr_block(m + obs_w + 12, 660, size=70)
        c.setStrokeColor(palette["line"])
        c.setLineWidth(0.8)
        c.line(m, yinv(748), m + 190, yinv(748))
        c.line(m + 220, yinv(748), m + 410, yinv(748))
        c.setFont(fonts["regular"], 8.1)
        c.setFillColor(palette["muted"])
        c.drawString(m, yinv(760), ntxt("Expedido por"))
        c.drawString(m + 220, yinv(760), ntxt("Recebido por"))
        c.setFillColor(palette["ink"])
        c.drawString(m, yinv(775), ntxt(created_by or emit_nome or "-"))
        c.drawString(m + 220, yinv(775), ntxt(dest_nome or "-"))

        c.saveState()
        c.setFillColor(palette["surface_alt"])
        c.roundRect(m, yinv(828), content_w, 52, 10, stroke=0, fill=1)
        c.restoreState()
        c.setFont(fonts["bold"], 8.2)
        c.setFillColor(palette["primary_dark"])
        c.drawString(m + 10, yinv(792), ntxt("Documento de transporte"))
        c.setFont(fonts["regular"], 7.5)
        c.setFillColor(palette["ink"])
        c.drawString(m + 10, yinv(806), ntxt("Este documento nao serve de fatura."))
        c.drawString(
            m + 10,
            yinv(817),
            ntxt(f"Data/hora de inicio do transporte: {fmt_datetime(data_transporte or data_emissao)}"),
        )
        c.drawRightString(w - m - 10, yinv(806), ntxt(f"ATCUD: {atcud or '-'}"))
        c.drawRightString(w - m - 10, yinv(817), ntxt(f"Serie/validacao: {at_validation_code or codigo_at or '-'}"))
        c.setFont(fonts["regular"], 7.2)
        c.setFillColor(palette["muted"])
        c.drawString(m + 10, yinv(826), ntxt(f"c/Ho - Processado por Programa Certificado n.o 0030/AT | Pagina {page_no}/{total_pages}"))

    def paginate_items(items):
        if not items:
            return [[]]
        pages = [items[:first_capacity]]
        rem = items[first_capacity:]
        while rem:
            pages.append(rem[:next_capacity])
            rem = rem[next_capacity:]
        return pages

    for via_idx, via_nome in enumerate(vias_to_render):
        pages = paginate_items(linhas)
        total_pages = len(pages)
        for page_no, page_rows in enumerate(pages, start=1):
            table_y = draw_first_header(page_no, total_pages, via_nome) if page_no == 1 else draw_next_header(page_no, total_pages, via_nome)
            draw_cancel_stamp()
            if not page_rows:
                c.setFont(fonts["regular"], 9.2)
                c.setFillColor(palette["muted"])
                c.drawString(m + 8, yinv(table_y + 16), ntxt("Sem linhas na guia."))
            else:
                y_row = table_y
                for local_idx, linha in enumerate(page_rows):
                    draw_row(y_row, local_idx, linha)
                    y_row += row_h
            if page_no == total_pages:
                draw_footer(page_no, total_pages)
            if page_no < total_pages or via_idx < len(vias_to_render) - 1:
                c.showPage()

    c.save()

def preview_expedicao_pdf_by_num(self, numero):
    _ensure_configured()
    ex = next((x for x in self.data.get("expedicoes", []) if str(x.get("numero", "")) == str(numero)), None)
    if not ex:
        messagebox.showerror("Erro", "Guia nao encontrada.")
        return
    import tempfile
    path = os.path.join(tempfile.gettempdir(), f"lugest_guia_{ex.get('numero','')}.pdf")
    self.render_expedicao_pdf(path, ex)
    try:
        os.startfile(path)
    except Exception:
        messagebox.showerror("Erro", "Nao foi possivel abrir o PDF da guia.")

def save_expedicao_pdf(self):
    _ensure_configured()
    ex = self.get_selected_expedicao(show_error=True)
    if not ex:
        return
    path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF", "*.pdf")],
        initialfile=f"guia_{ex.get('numero','')}.pdf",
    )
    if not path:
        return
    self.render_expedicao_pdf(path, ex)
    messagebox.showinfo("OK", "PDF guardado com sucesso.")

def print_expedicao_pdf(self):
    _ensure_configured()
    ex = self.get_selected_expedicao(show_error=True)
    if not ex:
        return
    import tempfile
    path = os.path.join(tempfile.gettempdir(), f"lugest_guia_{ex.get('numero','')}.pdf")
    self.render_expedicao_pdf(path, ex, include_all_vias=True)
    try:
        os.startfile(path, "print")
    except Exception:
        try:
            os.startfile(path)
        except Exception:
            messagebox.showerror("Erro", "Nao foi possivel imprimir a guia.")

def preview_expedicao_pdf(self):
    _ensure_configured()
    ex = self.get_selected_expedicao(show_error=True)
    if not ex:
        return
    self.preview_expedicao_pdf_by_num(ex.get("numero", ""))

def emitir_expedicao_off(self):
    _ensure_configured()
    enc = self.get_exp_selected_encomenda()
    if not enc:
        messagebox.showerror("Erro", "Selecione uma encomenda.")
        return
    if not self.exp_draft_linhas:
        messagebox.showerror("Erro", "Sem linhas na guia.")
        return
    cli_code = str(enc.get("cliente", "") or "")
    cli = find_cliente(self.data, cli_code) or {}
    cli_nome = str(cli.get("nome", "") or cli_code)
    for l in self.exp_draft_linhas:
        pid = str(l.get("peca_id", "") or "")
        p = None
        for px in encomenda_pecas(enc):
            if str(px.get("id", "") or "") == pid:
                p = px
                break
        if not p:
            messagebox.showerror("Erro", f"Peca nao encontrada para expedicao: {pid}")
            return
        disp = peca_qtd_disponivel_expedicao(p)
        if parse_float(l.get("qtd", 0), 0.0) > disp + 1e-9:
            messagebox.showerror("Erro", f"Quantidade superior ao disponivel na peca {p.get('ref_interna','')}.")
            return

    rodape = get_empresa_rodape_lines()
    local_carga = rodape[1] if len(rodape) > 1 else (rodape[0] if rodape else "")
    emit_cfg = get_guia_emitente_info()
    serie_guess = _exp_default_serie_id("GT", now_iso())
    serie_obj = _find_at_series(self.data, doc_type="GT", serie_id=serie_guess) if hasattr(self, "data") else None
    serie_validation = str((serie_obj or {}).get("validation_code", "") or "").strip()
    guia_data = self.prompt_dados_guia(
        "Confirmar Dados da Guia OFF",
        {
            "codigo_at": serie_validation,
            "tipo_via": "Original",
            "emitente_nome": emit_cfg.get("nome", ""),
            "emitente_nif": emit_cfg.get("nif", ""),
            "emitente_morada": emit_cfg.get("morada", ""),
            "destinatario": cli_nome,
            "dest_nif": str(cli.get("nif", "") or ""),
            "dest_morada": str(cli.get("morada", "") or ""),
            "local_carga": emit_cfg.get("local_carga", "") or local_carga,
            "local_descarga": str(cli.get("morada", "") or ""),
            "data_transporte": now_iso(),
            "transportador": "",
            "matricula": "",
            "observacoes": f"Expedicao da encomenda {enc.get('numero', '')}",
        },
        self.exp_draft_linhas,
    )
    if not guia_data:
        return
    exp_ids, exp_err = next_expedicao_identifiers(
        self.data,
        issue_date=guia_data.get("data_transporte", now_iso()),
        doc_type="GT",
        validation_code_hint=guia_data.get("codigo_at", ""),
    )
    if not exp_ids:
        messagebox.showerror("Serie AT", exp_err or "Nao foi possivel obter serie/ATCUD da guia.")
        return
    ex_num = exp_ids.get("numero", "")
    ex = {
        "numero": ex_num,
        "tipo": "OFF",
        "encomenda": enc.get("numero", ""),
        "cliente": cli_code,
        "cliente_nome": cli_nome,
        "codigo_at": exp_ids.get("validation_code", ""),
        "serie_id": exp_ids.get("serie_id", ""),
        "seq_num": exp_ids.get("seq_num", 0),
        "at_validation_code": exp_ids.get("validation_code", ""),
        "atcud": exp_ids.get("atcud", ""),
        "tipo_via": guia_data.get("tipo_via", "Original"),
        "emitente_nome": guia_data.get("emitente_nome", emit_cfg.get("nome", "")),
        "emitente_nif": guia_data.get("emitente_nif", emit_cfg.get("nif", "")),
        "emitente_morada": guia_data.get("emitente_morada", emit_cfg.get("morada", "")),
        "destinatario": guia_data.get("destinatario", cli_nome),
        "dest_nif": guia_data.get("dest_nif", str(cli.get("nif", "") or "")),
        "dest_morada": guia_data.get("dest_morada", str(cli.get("morada", "") or "")),
        "local_carga": guia_data.get("local_carga", emit_cfg.get("local_carga", "") or local_carga),
        "local_descarga": guia_data.get("local_descarga", str(cli.get("morada", "") or "")),
        "data_emissao": now_iso(),
        "data_transporte": guia_data.get("data_transporte", now_iso()),
        "matricula": guia_data.get("matricula", ""),
        "transportador": guia_data.get("transportador", ""),
        "estado": "Emitida",
        "observacoes": guia_data.get("observacoes", ""),
        "created_by": self.user.get("username", ""),
        "anulada": False,
        "anulada_motivo": "",
        "linhas": [],
    }
    for l in self.exp_draft_linhas:
        pid = str(l.get("peca_id", "") or "")
        p = None
        for px in encomenda_pecas(enc):
            if str(px.get("id", "") or "") == pid:
                p = px
                break
        if not p:
            continue
        qtd = parse_float(l.get("qtd", 0), 0.0)
        p["qtd_expedida"] = parse_float(p.get("qtd_expedida", 0), 0.0) + qtd
        p.setdefault("expedicoes", []).append(ex_num)
        ex["linhas"].append(
            {
                "encomenda": enc.get("numero", ""),
                "peca_id": pid,
                "ref_interna": l.get("ref_interna", ""),
                "ref_externa": l.get("ref_externa", ""),
                "descricao": l.get("descricao", ""),
                "qtd": qtd,
                "unid": l.get("unid", "UN"),
                "peso": parse_float(l.get("peso", 0), 0.0),
                "manual": False,
            }
        )
        atualizar_estado_peca(p)
    self.data.setdefault("expedicoes", []).append(ex)
    update_estado_expedicao_encomenda(enc)
    save_data(self.data)
    self.exp_draft_linhas = []
    self.refresh_expedicao()
    try:
        if hasattr(self, "nb") and hasattr(self, "tab_menu"):
            if str(self.nb.select()) == str(self.tab_menu):
                self.debounce_call("menu_stats", self.refresh_menu_dashboard_stats, delay_ms=250)
    except Exception:
        pass
    if messagebox.askyesno("OK", f"Guia emitida: {ex_num}\n\nPretende abrir o PDF agora?"):
        self.preview_expedicao_pdf_by_num(ex_num)

def criar_guia_manual(self):
    _ensure_configured()
    use_custom = self.exp_use_custom and CUSTOM_TK_AVAILABLE
    Win = ctk.CTkToplevel if use_custom else Toplevel
    Frm = ctk.CTkFrame if use_custom else ttk.Frame
    Lbl = ctk.CTkLabel if use_custom else ttk.Label
    Ent = ctk.CTkEntry if use_custom else ttk.Entry
    Btn = ctk.CTkButton if use_custom else ttk.Button
    Cmb = ctk.CTkComboBox if use_custom else ttk.Combobox

    win = Win(self.root)
    win.title("Criar Guia Manual")
    try:
        if use_custom:
            win.geometry("980x680")
            win.configure(fg_color="#f7f8fb")
        win.transient(self.root)
        win.grab_set()
    except Exception:
        pass

    cli_nome = StringVar()
    cli_nif = StringVar()
    cli_morada = StringVar()
    rodape = get_empresa_rodape_lines()
    emit_cfg = get_guia_emitente_info()
    serie_guess = _exp_default_serie_id("GT", now_iso())
    serie_obj = _find_at_series(self.data, doc_type="GT", serie_id=serie_guess) if hasattr(self, "data") else None
    serie_validation = str((serie_obj or {}).get("validation_code", "") or "").strip()
    local_carga = StringVar(value=str(emit_cfg.get("local_carga", "") or (rodape[1] if len(rodape) > 1 else (rodape[0] if rodape else ""))))
    local_descarga = StringVar()
    transportador = StringVar()
    matricula = StringVar()
    observacoes = StringVar()
    codigo_at = StringVar(value=serie_validation)
    emit_nome_var = StringVar(value=emit_cfg.get("nome", ""))
    prod_code = StringVar()
    desc_var = StringVar()
    qtd_var = StringVar(value="1")
    unid_var = StringVar(value="UN")
    line_items = []

    top = Frm(win, fg_color="#ffffff") if use_custom else Frm(win)
    top.pack(fill="x", padx=10, pady=(10, 6))
    Lbl(top, text="Cod. Validacao Serie AT").grid(row=0, column=0, sticky="w", padx=6, pady=4)
    Ent(top, textvariable=codigo_at, width=220 if use_custom else 24).grid(row=0, column=1, padx=6, pady=4, sticky="w")
    Lbl(top, text="Emitente").grid(row=0, column=2, sticky="w", padx=6, pady=4)
    Ent(top, textvariable=emit_nome_var, state="readonly", width=260 if use_custom else 30).grid(row=0, column=3, padx=6, pady=4, sticky="w")
    Lbl(top, text="Destinatario").grid(row=1, column=0, sticky="w", padx=6, pady=4)
    Ent(top, textvariable=cli_nome, width=260 if use_custom else 30).grid(row=1, column=1, padx=6, pady=4, sticky="w")
    Lbl(top, text="NIF").grid(row=1, column=2, sticky="w", padx=6, pady=4)
    Ent(top, textvariable=cli_nif, width=150 if use_custom else 16).grid(row=1, column=3, padx=6, pady=4, sticky="w")
    Lbl(top, text="Morada").grid(row=2, column=0, sticky="w", padx=6, pady=4)
    Ent(top, textvariable=cli_morada, width=420 if use_custom else 50).grid(row=2, column=1, columnspan=3, padx=6, pady=4, sticky="w")
    Lbl(top, text="Local carga").grid(row=3, column=0, sticky="w", padx=6, pady=4)
    Ent(top, textvariable=local_carga, width=260 if use_custom else 30).grid(row=3, column=1, padx=6, pady=4, sticky="w")
    Lbl(top, text="Local descarga").grid(row=3, column=2, sticky="w", padx=6, pady=4)
    Ent(top, textvariable=local_descarga, width=260 if use_custom else 30).grid(row=3, column=3, padx=6, pady=4, sticky="w")
    Lbl(top, text="Transportador").grid(row=4, column=0, sticky="w", padx=6, pady=4)
    Ent(top, textvariable=transportador, width=260 if use_custom else 30).grid(row=4, column=1, padx=6, pady=4, sticky="w")
    Lbl(top, text="Matricula").grid(row=4, column=2, sticky="w", padx=6, pady=4)
    Ent(top, textvariable=matricula, width=150 if use_custom else 16).grid(row=4, column=3, padx=6, pady=4, sticky="w")
    Lbl(top, text="Observacoes").grid(row=5, column=0, sticky="w", padx=6, pady=4)
    Ent(top, textvariable=observacoes, width=620 if use_custom else 74).grid(row=5, column=1, columnspan=3, padx=6, pady=4, sticky="w")
    Lbl(top, text="Vias").grid(row=6, column=0, sticky="w", padx=6, pady=4)
    Lbl(top, text="Original / Duplicado / Triplicado (automatico)").grid(row=6, column=1, columnspan=3, sticky="w", padx=6, pady=4)

    items = Frm(win, fg_color="#ffffff") if use_custom else Frm(win)
    items.pack(fill="both", expand=True, padx=10, pady=(0, 6))
    hdr = Frm(items, fg_color="transparent") if use_custom else Frm(items)
    hdr.pack(fill="x", padx=6, pady=(6, 2))
    Lbl(hdr, text="Produto stock").pack(side="left", padx=(0, 6))
    prod_values = [f"{p.get('codigo','')} - {p.get('descricao','')}" for p in self.data.get("produtos", [])]
    if use_custom:
        cmb = Cmb(hdr, variable=prod_code, values=prod_values, width=260)
    else:
        cmb = Cmb(hdr, textvariable=prod_code, values=prod_values, width=36)
    cmb.pack(side="left", padx=4)
    Lbl(hdr, text="Descricao").pack(side="left", padx=(10, 4))
    Ent(hdr, textvariable=desc_var, width=260 if use_custom else 30).pack(side="left", padx=4)
    Lbl(hdr, text="Qtd").pack(side="left", padx=(10, 4))
    Ent(hdr, textvariable=qtd_var, width=80 if use_custom else 8).pack(side="left", padx=4)
    Lbl(hdr, text="Unid").pack(side="left", padx=(10, 4))
    Ent(hdr, textvariable=unid_var, width=70 if use_custom else 8).pack(side="left", padx=4)

    tbl = ttk.Treeview(
        items,
        columns=("codigo", "descricao", "qtd", "unid"),
        show="headings",
        height=10,
        style=("EXP.Treeview" if self.exp_use_custom else ""),
    )
    tbl.heading("codigo", text="Codigo")
    tbl.heading("descricao", text="Descricao")
    tbl.heading("qtd", text="Qtd")
    tbl.heading("unid", text="Unid")
    tbl.column("codigo", width=140, anchor="center")
    tbl.column("descricao", width=520, anchor="w")
    tbl.column("qtd", width=100, anchor="e")
    tbl.column("unid", width=90, anchor="center")
    tbl.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=(0, 6))
    sb = ttk.Scrollbar(items, orient="vertical", command=tbl.yview)
    tbl.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y", padx=(0, 6), pady=(0, 6))

    def on_prod_pick(_=None):
        txt = str(prod_code.get() or "").strip()
        cod = txt.split(" - ", 1)[0].strip() if " - " in txt else txt.strip()
        prod = next((x for x in self.data.get("produtos", []) if x.get("codigo") == cod), None)
        if prod and not desc_var.get().strip():
            desc_var.set(prod.get("descricao", ""))
    try:
        cmb.bind("<<ComboboxSelected>>", on_prod_pick)
    except Exception:
        pass

    def add_line():
        txt = str(prod_code.get() or "").strip()
        cod = txt.split(" - ", 1)[0].strip() if " - " in txt else txt.strip()
        desc = desc_var.get().strip()
        if not desc and not cod:
            messagebox.showerror("Erro", "Selecione um artigo de stock ou introduza descricao.")
            return
        qtd = parse_float(qtd_var.get(), 0.0)
        if qtd <= 0:
            messagebox.showerror("Erro", "Quantidade invalida.")
            return
        unid = unid_var.get().strip() or "UN"
        if cod:
            prod = next((x for x in self.data.get("produtos", []) if x.get("codigo") == cod), None)
            if not prod:
                messagebox.showerror("Erro", "Produto nao encontrado.")
                return
            if not desc:
                desc = prod.get("descricao", "")
        line_items.append({"produto_codigo": cod, "descricao": desc, "qtd": qtd, "unid": unid})
        tbl.insert("", END, values=(cod, desc, fmt_num(qtd), unid))
        desc_var.set("")
        qtd_var.set("1")

    def remove_line():
        sel = tbl.selection()
        if not sel:
            return
        idx = tbl.index(sel[0])
        if idx < 0 or idx >= len(line_items):
            return
        line_items.pop(idx)
        tbl.delete(sel[0])

    bar = Frm(win, fg_color="transparent") if use_custom else Frm(win)
    bar.pack(fill="x", padx=10, pady=(0, 8))
    Btn(bar, text="Adicionar linha", command=add_line, width=130 if use_custom else None).pack(side="left", padx=4)
    Btn(bar, text="Remover linha", command=remove_line, width=130 if use_custom else None).pack(side="left", padx=4)

    def emit():
        if not cli_nome.get().strip():
            messagebox.showerror("Erro", "Indique o destinatario.")
            return
        if not line_items:
            messagebox.showerror("Erro", "Sem linhas na guia.")
            return
        for ln in line_items:
            cod = str(ln.get("produto_codigo", "") or "").strip()
            if not cod:
                continue
            prod = next((x for x in self.data.get("produtos", []) if x.get("codigo") == cod), None)
            if not prod:
                messagebox.showerror("Erro", f"Produto nao encontrado: {cod}")
                return
            if parse_float(ln.get("qtd", 0), 0.0) > parse_float(prod.get("qty", 0), 0.0) + 1e-9:
                messagebox.showerror("Erro", f"Stock insuficiente para {cod}.")
                return

        exp_ids, exp_err = next_expedicao_identifiers(
            self.data,
            issue_date=now_iso(),
            doc_type="GT",
            validation_code_hint=codigo_at.get().strip(),
        )
        if not exp_ids:
            messagebox.showerror("Serie AT", exp_err or "Nao foi possivel obter serie/ATCUD da guia.")
            return
        ex_num = exp_ids.get("numero", "")
        ex = {
            "numero": ex_num,
            "tipo": "Manual",
            "encomenda": "",
            "cliente": "",
            "cliente_nome": cli_nome.get().strip(),
            "codigo_at": exp_ids.get("validation_code", ""),
            "serie_id": exp_ids.get("serie_id", ""),
            "seq_num": exp_ids.get("seq_num", 0),
            "at_validation_code": exp_ids.get("validation_code", ""),
            "atcud": exp_ids.get("atcud", ""),
            "tipo_via": "Original",
            "emitente_nome": emit_cfg.get("nome", ""),
            "emitente_nif": emit_cfg.get("nif", ""),
            "emitente_morada": emit_cfg.get("morada", ""),
            "destinatario": cli_nome.get().strip(),
            "dest_nif": cli_nif.get().strip(),
            "dest_morada": cli_morada.get().strip(),
            "local_carga": local_carga.get().strip(),
            "local_descarga": local_descarga.get().strip(),
            "data_emissao": now_iso(),
            "data_transporte": now_iso(),
            "matricula": matricula.get().strip(),
            "transportador": transportador.get().strip(),
            "estado": "Emitida",
            "observacoes": observacoes.get().strip(),
            "created_by": self.user.get("username", ""),
            "anulada": False,
            "anulada_motivo": "",
            "linhas": [],
        }
        for ln in line_items:
            cod = str(ln.get("produto_codigo", "") or "").strip()
            qtd = parse_float(ln.get("qtd", 0), 0.0)
            if cod:
                prod = next((x for x in self.data.get("produtos", []) if x.get("codigo") == cod), None)
                if prod:
                    prod["qty"] = max(0.0, parse_float(prod.get("qty", 0), 0.0) - qtd)
                    prod["atualizado_em"] = now_iso()
            ex["linhas"].append(
                {
                    "encomenda": "",
                    "peca_id": "",
                    "ref_interna": cod,
                    "ref_externa": "",
                    "descricao": ln.get("descricao", ""),
                    "qtd": qtd,
                    "unid": ln.get("unid", "UN"),
                    "peso": 0.0,
                    "manual": True,
                }
            )
        self.data.setdefault("expedicoes", []).append(ex)
        save_data(self.data)
        self.refresh_expedicao()
        try:
            self.refresh_produtos()
        except Exception:
            pass
        if messagebox.askyesno("OK", f"Guia manual emitida: {ex_num}\n\nPretende abrir o PDF agora?"):
            self.preview_expedicao_pdf_by_num(ex_num)
        win.destroy()

    foot = Frm(win, fg_color="transparent") if use_custom else Frm(win)
    foot.pack(fill="x", padx=10, pady=(0, 10))
    Btn(foot, text="Emitir Guia Manual", command=emit, width=170 if use_custom else None).pack(side="left", padx=4)
    Btn(foot, text="Cancelar", command=win.destroy, width=130 if use_custom else None).pack(side="right", padx=4)

def gerar_nes_por_fornecedor(self):
    _ensure_configured()
    ne = self.get_ne_selected()
    if not ne:
        return
    # sincroniza o que esta no ecra antes de repartir
    if self.ne_num.get().strip() == ne.get("numero"):
        ne["fornecedor"] = self.ne_fornecedor.get().strip()
        ne["fornecedor_id"] = self.ne_fornecedor_id.get().strip()
        ne["contacto"] = self.ne_contacto.get().strip()
        ne["data_entrega"] = self.ne_entrega.get().strip()
        ne["obs"] = self.ne_obs.get().strip()
        ne["linhas"] = self.ne_collect_lines()
        ne["total"] = sum(parse_float(l.get("total", 0), 0) for l in ne.get("linhas", []))

    if not ne.get("linhas"):
        messagebox.showerror("Erro", "A nota nao tem linhas.")
        return

    fornecedores = self.data.get("fornecedores", [])

    def resolve_forn(txt):
        t = (txt or "").strip()
        if not t:
            return "", "", ""
        fid = ""
        nome = t
        if " - " in t:
            fid, nome = [x.strip() for x in t.split(" - ", 1)]
        fobj = None
        if fid:
            fobj = next((f for f in fornecedores if str(f.get("id", "")).strip() == fid), None)
        if not fobj and nome:
            fobj = next((f for f in fornecedores if str(f.get("nome", "")).strip().lower() == nome.lower()), None)
        if fobj:
            fid = str(fobj.get("id", "")).strip()
            nome = str(fobj.get("nome", "")).strip()
            contacto = str(fobj.get("contacto", "")).strip()
            return fid, f"{fid} - {nome}".strip(" -"), contacto
        return fid, t, ""

    grupos = {}
    faltam = []
    for l in ne.get("linhas", []):
        alvo = (l.get("fornecedor_linha") or ne.get("fornecedor") or "").strip()
        fid, forn_txt, contacto = resolve_forn(alvo)
        if not forn_txt:
            faltam.append(l.get("ref", ""))
            continue
        k = fid or forn_txt
        if k not in grupos:
            grupos[k] = {"fornecedor_id": fid, "fornecedor": forn_txt, "contacto": contacto, "linhas": []}
        novo_l = dict(l)
        novo_l["fornecedor_linha"] = forn_txt
        novo_l["entregue"] = False
        novo_l["_stock_in"] = False
        grupos[k]["linhas"].append(novo_l)

    if faltam:
        messagebox.showerror(
            "Erro",
            "Existem linhas sem fornecedor adjudicado.\n"
            "Edite as linhas e defina 'Fornecedor (linha)' para todas antes de gerar.\n\n"
            f"Produtos: {', '.join(sorted(set([x for x in faltam if x])))}",
        )
        return
    if len(grupos) <= 1:
        messagebox.showinfo("Info", "So existe um fornecedor adjudicado. Nao ha divisao a fazer.")
        return
    if not messagebox.askyesno(
        "Confirmar",
        f"Gerar {len(grupos)} Notas de Encomenda separadas por fornecedor\n"
        f"a partir da nota {ne.get('numero')}?",
    ):
        return

    notas = self.data.setdefault("notas_encomenda", [])
    criadas = []
    for g in grupos.values():
        novo_num = next_ne_numero(self.data)
        linhas = g.get("linhas", [])
        novo = {
            "numero": novo_num,
            "fornecedor": g.get("fornecedor", ""),
            "fornecedor_id": g.get("fornecedor_id", ""),
            "contacto": g.get("contacto", ""),
            "data_entrega": ne.get("data_entrega", ""),
            "obs": (f"Gerada de {ne.get('numero')} | " + (ne.get("obs", "") or "")).strip(" |"),
            "local_descarga": ne.get("local_descarga", ""),
            "meio_transporte": ne.get("meio_transporte", ""),
            "linhas": linhas,
            "estado": "Aprovada",
            "_draft": False,
            "origem_cotacao": ne.get("numero", ""),
            "total": sum(parse_float(x.get("total", 0), 0) for x in linhas),
        }
        notas.append(novo)
        criadas.append(novo_num)

    ne["estado"] = "Convertida"
    ne["ne_geradas"] = criadas
    ne["oculta"] = True
    ne["_draft"] = False
    save_data(self.data)
    self.refresh_ne()
    if criadas:
        for iid in self.tbl_ne.get_children():
            if self.tbl_ne.item(iid, "values")[0] == criadas[0]:
                self.tbl_ne.selection_set(iid)
                self.tbl_ne.focus(iid)
                self.tbl_ne.see(iid)
                break
        self.on_ne_select()
    messagebox.showinfo("OK", "Notas geradas: " + ", ".join(criadas))

def _next_materia_id(self):
    _ensure_configured()
    max_n = 0
    for m in self.data.get("materiais", []):
        mid = str(m.get("id", ""))
        if mid.startswith("MAT") and mid[3:].isdigit():
            max_n = max(max_n, int(mid[3:]))
    return f"MAT{max_n + 1:05d}"

def _receber_linha_materia_ne(self, l, ne_num, qtd_recebida=None):
    _ensure_configured()
    qtd = parse_float(qtd_recebida if qtd_recebida is not None else l.get("qtd", 0), 0)
    if qtd <= 0:
        return True
    mats = self.data.setdefault("materiais", [])
    ref = (l.get("ref") or "").strip()
    m = next((x for x in mats if x.get("id") == ref), None)
    if not m:
        material = (l.get("material") or "").strip()
        if not material:
            return False
        comp = parse_float(l.get("comprimento", 0), 0)
        larg = parse_float(l.get("largura", 0), 0)
        metros = parse_float(l.get("metros", 0), 0)
        formato = (l.get("formato") or "").strip() or ("Tubo" if metros > 0 and (comp <= 0 or larg <= 0) else "Chapa")
        m = {
            "id": self._next_materia_id(),
            "formato": formato,
            "material": material,
            "espessura": str(l.get("espessura", "")).strip(),
            "comprimento": comp,
            "largura": larg,
            "metros": metros,
            "quantidade": 0.0,
            "reservado": 0.0,
            "Localizacao": (l.get("localizacao") or "").strip(),
            "lote_fornecedor": (l.get("lote_fornecedor") or "").strip(),
            "peso_unid": parse_float(l.get("peso_unid", 0), 0),
            "p_compra": parse_float(l.get("p_compra", 0), 0),
            "is_sobra": False,
            "atualizado_em": now_iso(),
        }
        mats.append(m)
        ref = m["id"]
        l["ref"] = ref
    else:
        if not m.get("formato"):
            m["formato"] = (l.get("formato") or "").strip() or detect_materia_formato(m)
    m["quantidade"] = parse_float(m.get("quantidade", 0), 0) + qtd
    preco_unit = parse_float(l.get("preco", 0), 0)
    if preco_unit > 0:
        formato = m.get("formato", detect_materia_formato(m))
        peso = parse_float(m.get("peso_unid", 0), 0)
        metros = parse_float(m.get("metros", 0), 0)
        if formato == "Tubo" and metros > 0:
            m["p_compra"] = round(preco_unit / metros, 6)
        elif formato in ("Chapa", "Perfil") and peso > 0:
            m["p_compra"] = round(preco_unit / peso, 6)
        elif parse_float(m.get("p_compra", 0), 0) <= 0:
            m["p_compra"] = preco_unit
    m["atualizado_em"] = now_iso()
    push_unique(self.data.setdefault("materiais_hist", []), m.get("material", ""))
    push_unique(self.data.setdefault("espessuras_hist", []), m.get("espessura", ""))
    log_stock(self.data, "ENTRADA_NE", f"{m.get('id','')} qtd+={qtd} via {ne_num}")
    return True

def _dialog_associar_fatura_ne(self, ne):
    _ensure_configured()
    use_custom = CUSTOM_TK_AVAILABLE and os.environ.get("USE_CUSTOM_NE", "1") != "0"
    Win = ctk.CTkToplevel if use_custom else Toplevel
    Frm = ctk.CTkFrame if use_custom else ttk.Frame
    Lbl = ctk.CTkLabel if use_custom else ttk.Label
    Ent = ctk.CTkEntry if use_custom else ttk.Entry
    Btn = ctk.CTkButton if use_custom else ttk.Button
    Chk = ctk.CTkCheckBox if use_custom else ttk.Checkbutton

    win = Win(self.root)
    win.title("Associar Documento")
    try:
        win.geometry("900x360")
        win.resizable(False, False)
    except Exception:
        pass

    out = {"ok": False}
    box = Frm(win, fg_color="#f7f8fb") if use_custom else Frm(win)
    box.pack(fill="both", expand=True, padx=10, pady=10)

    title = Lbl(
        box,
        text=f"Associar documento apos entrega - {ne.get('numero','')}",
        font=("Segoe UI", 16, "bold") if use_custom else None,
    )
    title.grid(row=0, column=0, columnspan=7, sticky="w", padx=8, pady=(8, 12))

    titulo_var = StringVar(value="")
    guia_var = StringVar(value=str(ne.get("guia_ultima", "") or ""))
    fatura_var = StringVar(value=str(ne.get("fatura_ultima", "") or ""))
    caminho_var = StringVar(value=str(ne.get("fatura_caminho_ultima", "") or ""))
    data_doc_var = StringVar(value=str(ne.get("data_doc_ultima", "") or now_iso()[:10]))
    data_ent_var = StringVar(value=str(ne.get("data_ultima_entrega", "") or now_iso()[:10]))
    obs_var = StringVar(value="")
    aplica_linhas = BooleanVar(value=True)

    Lbl(box, text="Titulo").grid(row=1, column=0, sticky="w", padx=8, pady=5)
    Ent(box, textvariable=titulo_var, width=260 if use_custom else 30).grid(row=1, column=1, columnspan=3, sticky="we", padx=8, pady=5)
    Lbl(box, text="Data Documento").grid(row=1, column=4, sticky="w", padx=8, pady=5)
    Ent(box, textvariable=data_doc_var, width=130 if use_custom else 14).grid(row=1, column=5, sticky="w", padx=8, pady=5)
    Btn(
        box,
        text="Calendario",
        width=95 if use_custom else None,
        command=lambda: self.pick_date(data_doc_var, parent=win),
    ).grid(row=1, column=6, sticky="w", padx=8, pady=5)

    Lbl(box, text="Guia Transporte").grid(row=2, column=0, sticky="w", padx=8, pady=5)
    Ent(box, textvariable=guia_var, width=170 if use_custom else 22).grid(row=2, column=1, sticky="we", padx=8, pady=5)
    Lbl(box, text="Fatura").grid(row=2, column=2, sticky="w", padx=8, pady=5)
    Ent(box, textvariable=fatura_var, width=200 if use_custom else 24).grid(row=2, column=3, sticky="we", padx=8, pady=5)
    Lbl(box, text="Data Registo").grid(row=2, column=4, sticky="w", padx=8, pady=5)
    Ent(box, textvariable=data_ent_var, width=130 if use_custom else 14).grid(row=2, column=5, sticky="w", padx=8, pady=5)
    Btn(
        box,
        text="Calendario",
        width=95 if use_custom else None,
        command=lambda: self.pick_date(data_ent_var, parent=win),
    ).grid(row=2, column=6, sticky="w", padx=8, pady=5)

    def pick_path():
        path = filedialog.askopenfilename(
            title="Selecionar documento",
            filetypes=[
                ("Documentos", "*.pdf *.png *.jpg *.jpeg *.bmp *.xlsx *.xls *.doc *.docx"),
                ("Todos", "*.*"),
            ],
        )
        if path:
            caminho_var.set(path)

    Lbl(box, text="Caminho").grid(row=3, column=0, sticky="w", padx=8, pady=5)
    Ent(box, textvariable=caminho_var, width=420 if use_custom else 52).grid(row=3, column=1, columnspan=4, sticky="we", padx=8, pady=5)
    Btn(box, text="Ficheiro", width=95 if use_custom else None, command=pick_path).grid(row=3, column=5, sticky="w", padx=8, pady=5)
    Lbl(box, text="Observacoes").grid(row=4, column=0, sticky="w", padx=8, pady=5)
    Ent(box, textvariable=obs_var, width=520 if use_custom else 70).grid(row=4, column=1, columnspan=6, sticky="we", padx=8, pady=5)
    Chk(
        box,
        text="Aplicar aos registos de linhas ja entregues",
        variable=aplica_linhas,
        onvalue=True,
        offvalue=False,
    ).grid(row=5, column=0, columnspan=7, sticky="w", padx=8, pady=(8, 10))

    def on_ok():
        titulo = (titulo_var.get() or "").strip()
        guia = (guia_var.get() or "").strip()
        fatura = (fatura_var.get() or "").strip()
        caminho = (caminho_var.get() or "").strip()
        obs = (obs_var.get() or "").strip()
        if not titulo and not guia and not fatura and not caminho and not obs:
            messagebox.showerror("Erro", "Preencha pelo menos Titulo, Guia, Fatura, Caminho ou Observacoes.")
            return
        if guia and fatura:
            tipo = "GUIA_FATURA"
        elif fatura:
            tipo = "FATURA"
        elif guia:
            tipo = "GUIA"
        else:
            tipo = "DOCUMENTO"
        out.update(
            {
                "ok": True,
                "tipo": tipo,
                "titulo": titulo,
                "guia": guia,
                "fatura": fatura,
                "caminho": caminho,
                "data_doc": (data_doc_var.get() or "").strip(),
                "data_entrega": (data_ent_var.get() or "").strip(),
                "obs": obs,
                "aplica_linhas": bool(aplica_linhas.get()),
            }
        )
        win.destroy()

    for col_idx in range(7):
        try:
            box.grid_columnconfigure(col_idx, weight=1 if col_idx in (1, 3, 4) else 0)
        except Exception:
            pass

    btm = Frm(box, fg_color="#f7f8fb") if use_custom else Frm(box)
    btm.grid(row=6, column=0, columnspan=7, sticky="ew", padx=8, pady=(10, 4))
    Btn(btm, text="Guardar", command=on_ok, width=120 if use_custom else None).pack(side="left", padx=4)
    Btn(btm, text="Cancelar", command=win.destroy, width=120 if use_custom else None).pack(side="right", padx=4)

    win.grab_set()
    self.root.wait_window(win)
    return out if out.get("ok") else None

def associar_fatura_ne(self):
    _ensure_configured()
    ne = self.get_ne_selected()
    if not ne:
        return
    entrega_info = self._dialog_associar_fatura_ne(ne)
    if not entrega_info:
        return

    if entrega_info.get("aplica_linhas"):
        for l in ne.get("linhas", []):
            qtd_tot = parse_float(l.get("qtd", 0), 0)
            qtd_ent = parse_float(l.get("qtd_entregue", l.get("qtd", 0) if l.get("entregue") else 0), 0)
            if qtd_ent <= 0 and not l.get("entregue"):
                continue
            l["guia_entrega"] = entrega_info.get("guia", "")
            l["fatura_entrega"] = entrega_info.get("fatura", "")
            l["data_doc_entrega"] = entrega_info.get("data_doc", "")
            l["data_entrega_real"] = entrega_info.get("data_entrega", "")
            l.setdefault("entregas_linha", []).append(
                {
                    "data_registo": now_iso(),
                    "data_entrega": entrega_info.get("data_entrega", ""),
                    "data_documento": entrega_info.get("data_doc", ""),
                    "guia": entrega_info.get("guia", ""),
                    "fatura": entrega_info.get("fatura", ""),
                    "obs": entrega_info.get("obs", ""),
                    "qtd": min(qtd_ent, qtd_tot),
                    "tipo": entrega_info.get("tipo", "DOCUMENTO"),
                }
            )

    ne["guia_ultima"] = entrega_info.get("guia", "")
    ne["fatura_ultima"] = entrega_info.get("fatura", "")
    ne["data_doc_ultima"] = entrega_info.get("data_doc", "")
    ne["data_ultima_entrega"] = entrega_info.get("data_entrega", "") or ne.get("data_ultima_entrega", "")
    if entrega_info.get("caminho") and entrega_info.get("fatura"):
        ne["fatura_caminho_ultima"] = entrega_info.get("caminho", "")
    ne.setdefault("documentos", []).append(
        {
            "data_registo": now_iso(),
            "tipo": entrega_info.get("tipo", "DOCUMENTO"),
            "titulo": entrega_info.get("titulo", ""),
            "caminho": entrega_info.get("caminho", ""),
            "guia": entrega_info.get("guia", ""),
            "fatura": entrega_info.get("fatura", ""),
            "data_entrega": entrega_info.get("data_entrega", ""),
            "data_documento": entrega_info.get("data_doc", ""),
            "obs": entrega_info.get("obs", ""),
        }
    )
    ne.setdefault("entregas", []).append(
        {
            "data_registo": now_iso(),
            "data_entrega": entrega_info.get("data_entrega", ""),
            "guia": entrega_info.get("guia", ""),
            "fatura": entrega_info.get("fatura", ""),
            "data_documento": entrega_info.get("data_doc", ""),
            "obs": entrega_info.get("obs", ""),
            "linhas": [],
            "quantidade_linhas": 0,
            "quantidade_total": 0,
            "tipo": entrega_info.get("tipo", "DOCUMENTO"),
            "titulo": entrega_info.get("titulo", ""),
            "caminho": entrega_info.get("caminho", ""),
        }
    )
    save_data(self.data)
    self.refresh_ne()
    for iid in self.tbl_ne.get_children():
        if self.tbl_ne.item(iid, "values")[0] == ne.get("numero"):
            self.tbl_ne.selection_set(iid)
            self.tbl_ne.see(iid)
            break
    self.on_ne_select()
    messagebox.showinfo("OK", "Documento associado com sucesso.")

def show_ne_documentos(self):
    _ensure_configured()
    ne = self.get_ne_selected()
    if not ne:
        return
    docs = []
    for doc in ne.get("documentos", []):
        if not isinstance(doc, dict):
            continue
        docs.append(
            {
                "tipo": str(doc.get("tipo", "") or ""),
                "titulo": str(doc.get("titulo", "") or ""),
                "guia": str(doc.get("guia", "") or ""),
                "fatura": str(doc.get("fatura", "") or ""),
                "data_documento": str(doc.get("data_documento", "") or ""),
                "data_entrega": str(doc.get("data_entrega", "") or ""),
                "data_registo": str(doc.get("data_registo", "") or ""),
                "caminho": str(doc.get("caminho", "") or ""),
                "obs": str(doc.get("obs", "") or ""),
            }
        )
    if not docs:
        for ent in ne.get("entregas", []):
            if not isinstance(ent, dict):
                continue
            docs.append(
                {
                    "tipo": str(ent.get("tipo", "") or ""),
                    "titulo": str(ent.get("titulo", "") or ""),
                    "guia": str(ent.get("guia", "") or ""),
                    "fatura": str(ent.get("fatura", "") or ""),
                    "data_documento": str(ent.get("data_documento", "") or ""),
                    "data_entrega": str(ent.get("data_entrega", "") or ""),
                    "data_registo": str(ent.get("data_registo", "") or ""),
                    "caminho": str(ent.get("caminho", "") or ""),
                    "obs": str(ent.get("obs", "") or ""),
                }
            )
    if not docs:
        messagebox.showinfo("Info", "Esta nota ainda nao tem documentos associados.")
        return

    use_custom = CUSTOM_TK_AVAILABLE and os.environ.get("USE_CUSTOM_NE", "1") != "0"
    Win = ctk.CTkToplevel if use_custom else Toplevel
    Frm = ctk.CTkFrame if use_custom else ttk.Frame
    Btn = ctk.CTkButton if use_custom else ttk.Button
    win = Win(self.root)
    win.title(f"Documentos {ne.get('numero','')}")
    try:
        win.geometry("1080x460")
    except Exception:
        pass
    frame = Frm(win, fg_color="#f7f8fb") if use_custom else Frm(win)
    frame.pack(fill="both", expand=True, padx=10, pady=10)
    cols = ("tipo", "titulo", "guia", "fatura", "data_doc", "data_ent", "registo", "caminho")
    tree = ttk.Treeview(frame, columns=cols, show="headings", height=14, style="NE.Treeview" if use_custom else "")
    headers = {
        "tipo": "Tipo",
        "titulo": "Titulo",
        "guia": "Guia",
        "fatura": "Fatura",
        "data_doc": "Data Doc",
        "data_ent": "Data Entrega",
        "registo": "Registo",
        "caminho": "Caminho",
    }
    widths = {"tipo": 110, "titulo": 220, "guia": 110, "fatura": 110, "data_doc": 95, "data_ent": 95, "registo": 145, "caminho": 260}
    for col in cols:
        tree.heading(col, text=headers[col])
        tree.column(col, width=widths[col], anchor=("w" if col in ("titulo", "caminho") else "center"))
    for doc in docs:
        tree.insert(
            "",
            END,
            values=(
                doc.get("tipo", "") or "-",
                doc.get("titulo", "") or "-",
                doc.get("guia", "") or "-",
                doc.get("fatura", "") or "-",
                doc.get("data_documento", "") or "-",
                doc.get("data_entrega", "") or "-",
                str(doc.get("data_registo", "") or "").replace("T", " ")[:19] or "-",
                doc.get("caminho", "") or "-",
            ),
        )
    tree.pack(side="left", fill="both", expand=True, padx=(0, 6))
    sb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=sb.set)
    sb.pack(side="right", fill="y")

    def open_selected():
        sel = tree.selection()
        if not sel:
            messagebox.showerror("Erro", "Selecione um documento.")
            return
        idx = tree.index(sel[0])
        if idx < 0 or idx >= len(docs):
            return
        path = str(docs[idx].get("caminho", "") or "").strip()
        if not path:
            messagebox.showinfo("Info", "O registo selecionado nao tem caminho associado.")
            return
        if not os.path.isabs(path):
            path = os.path.abspath(path)
        if not os.path.exists(path):
            messagebox.showerror("Erro", f"Ficheiro nao encontrado:\n{path}")
            return
        try:
            os.startfile(path)
        except Exception as exc:
            messagebox.showerror("Erro", f"Nao foi possivel abrir o ficheiro:\n{exc}")

    bar = Frm(win, fg_color="#f7f8fb") if use_custom else Frm(win)
    bar.pack(fill="x", padx=10, pady=(0, 10))
    Btn(bar, text="Abrir ficheiro", command=open_selected, width=140 if use_custom else None).pack(side="left", padx=4)
    Btn(bar, text="Fechar", command=win.destroy, width=120 if use_custom else None).pack(side="right", padx=4)
    win.grab_set()
    self.root.wait_window(win)

def _dialog_confirmar_entrega_ne(self, ne):
    _ensure_configured()
    use_custom = CUSTOM_TK_AVAILABLE and os.environ.get("USE_CUSTOM_NE", "1") != "0"
    Win = ctk.CTkToplevel if use_custom else Toplevel
    Frm = ctk.CTkFrame if use_custom else ttk.Frame
    Lbl = ctk.CTkLabel if use_custom else ttk.Label
    Ent = ctk.CTkEntry if use_custom else ttk.Entry
    Btn = ctk.CTkButton if use_custom else ttk.Button
    Chk = ctk.CTkCheckBox if use_custom else ttk.Checkbutton

    win = Win(self.root)
    win.title("Confirmar Entrega")
    try:
        win.geometry("980x700")
        win.minsize(860, 600)
    except Exception:
        pass

    out = {"ok": False}

    top = Frm(win, fg_color="#f7f8fb") if use_custom else Frm(win)
    top.pack(fill="x", padx=10, pady=(10, 6))
    Lbl(top, text=f"Nota: {ne.get('numero','')}", font=("Segoe UI", 14, "bold") if use_custom else None).pack(side="left", padx=4)
    Lbl(top, text=f"Fornecedor: {ne.get('fornecedor','')}").pack(side="left", padx=12)

    meta = Frm(win, fg_color="#f7f8fb") if use_custom else Frm(win)
    meta.pack(fill="x", padx=10, pady=6)
    guia_var = StringVar(value=str(ne.get("guia_ultima", "") or ""))
    fatura_var = StringVar(value=str(ne.get("fatura_ultima", "") or ""))
    data_doc_var = StringVar(value=str(ne.get("data_doc_ultima", "") or now_iso()[:10]))
    data_ent_var = StringVar(value=now_iso()[:10])
    obs_var = StringVar(value="")

    Lbl(meta, text="Guia Transporte").grid(row=0, column=0, padx=4, pady=4, sticky="w")
    Ent(meta, textvariable=guia_var, width=170 if use_custom else 24).grid(row=0, column=1, padx=4, pady=4, sticky="w")
    Lbl(meta, text="Fatura").grid(row=0, column=2, padx=4, pady=4, sticky="w")
    Ent(meta, textvariable=fatura_var, width=170 if use_custom else 20).grid(row=0, column=3, padx=4, pady=4, sticky="w")
    Lbl(meta, text="Data Documento").grid(row=0, column=4, padx=4, pady=4, sticky="w")
    Ent(meta, textvariable=data_doc_var, width=130 if use_custom else 14).grid(row=0, column=5, padx=4, pady=4, sticky="w")
    Btn(meta, text="Calendario", width=90 if use_custom else None, command=lambda: self.pick_date(data_doc_var, parent=win)).grid(row=0, column=6, padx=4, pady=4)
    Lbl(meta, text="Data Entrega").grid(row=1, column=0, padx=4, pady=4, sticky="w")
    Ent(meta, textvariable=data_ent_var, width=130 if use_custom else 14).grid(row=1, column=1, padx=4, pady=4, sticky="w")
    Btn(meta, text="Calendario", width=90 if use_custom else None, command=lambda: self.pick_date(data_ent_var, parent=win)).grid(row=1, column=2, padx=4, pady=4, sticky="w")
    Lbl(meta, text="Obs").grid(row=1, column=3, padx=4, pady=4, sticky="w")
    Ent(meta, textvariable=obs_var, width=360 if use_custom else 46).grid(row=1, column=4, columnspan=3, padx=4, pady=4, sticky="we")
    try:
        meta.grid_columnconfigure(6, weight=1)
    except Exception:
        pass

    wrap = Frm(win, fg_color="#ffffff") if use_custom else Frm(win)
    wrap.pack(fill="both", expand=True, padx=10, pady=8)
    canvas = Canvas(wrap, borderwidth=0, highlightthickness=0, background="#ffffff")
    vsb = ttk.Scrollbar(wrap, orient="vertical", command=canvas.yview)
    inner = ttk.Frame(canvas)
    inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=inner, anchor="nw")
    canvas.configure(yscrollcommand=vsb.set)
    canvas.pack(side="left", fill="both", expand=True)
    vsb.pack(side="right", fill="y")

    Lbl(inner, text="Selecionar posturas a entregar", font=("Segoe UI", 12, "bold") if use_custom else None).grid(row=0, column=0, columnspan=5, sticky="w", padx=6, pady=(6, 4))
    Lbl(inner, text="Linha").grid(row=1, column=0, sticky="w", padx=8, pady=(2, 4))
    Lbl(inner, text="Qtd Total").grid(row=1, column=1, sticky="e", padx=8, pady=(2, 4))
    Lbl(inner, text="Qtd em Falta").grid(row=1, column=2, sticky="e", padx=8, pady=(2, 4))
    Lbl(inner, text="Rececionar").grid(row=1, column=3, sticky="e", padx=8, pady=(2, 4))
    line_rows = []
    row = 2
    for idx, l in enumerate(ne.get("linhas", [])):
        qtd_total = parse_float(l.get("qtd", 0), 0)
        qtd_entregue = parse_float(
            l.get("qtd_entregue", l.get("qtd", 0) if l.get("entregue") else 0),
            0,
        )
        qtd_falta = max(0.0, qtd_total - qtd_entregue)
        entregue = bool(l.get("_stock_in")) or qtd_falta <= 1e-9
        txt = f"{l.get('ref','')} | {l.get('descricao','')}"
        if entregue:
            txt += " [ENTREGUE]"
        var = BooleanVar(value=(not entregue))
        qtd_var = StringVar(value=(fmt_num(qtd_falta) if qtd_falta > 0 else "0"))
        chk = Chk(inner, text=txt, variable=var)
        chk.grid(row=row, column=0, sticky="w", padx=8, pady=3)
        Lbl(inner, text=fmt_num(qtd_total)).grid(row=row, column=1, sticky="e", padx=8, pady=3)
        Lbl(inner, text=fmt_num(qtd_falta)).grid(row=row, column=2, sticky="e", padx=8, pady=3)
        Ent(inner, textvariable=qtd_var, width=100 if use_custom else 12).grid(row=row, column=3, sticky="e", padx=8, pady=3)
        if entregue:
            try:
                chk.configure(state="disabled")
            except Exception:
                pass
        line_rows.append(
            {
                "idx": idx,
                "var": var,
                "qtd_var": qtd_var,
                "qtd_falta": qtd_falta,
                "entregue": entregue,
                "ref": l.get("ref", ""),
            }
        )
        row += 1

    def on_confirm():
        itens = []
        for it in line_rows:
            if it.get("entregue") or not bool(it["var"].get()):
                continue
            qtd = parse_float(it["qtd_var"].get(), 0)
            if qtd <= 0:
                continue
            if qtd > parse_float(it.get("qtd_falta", 0), 0) + 1e-9:
                messagebox.showerror("Erro", f"Quantidade acima do pendente na linha {it.get('ref','')}.")
                return
            itens.append({"idx": it["idx"], "qtd": qtd})
        if not itens:
            messagebox.showerror("Erro", "Selecione pelo menos uma postura para entregar.")
            return
        out.update(
            {
                "ok": True,
                "indices": [x["idx"] for x in itens],
                "itens": itens,
                "guia": (guia_var.get() or "").strip(),
                "fatura": (fatura_var.get() or "").strip(),
                "data_doc": (data_doc_var.get() or "").strip(),
                "data_entrega": (data_ent_var.get() or "").strip(),
                "obs": (obs_var.get() or "").strip(),
            }
        )
        win.destroy()

    def on_cancel():
        win.destroy()

    btns = Frm(win, fg_color="#f7f8fb") if use_custom else Frm(win)
    btns.pack(fill="x", padx=10, pady=(0, 10))
    Btn(btns, text="Confirmar Entrega", command=on_confirm, width=170 if use_custom else None).pack(side="left", padx=4)
    Btn(btns, text="Cancelar", command=on_cancel, width=120 if use_custom else None).pack(side="right", padx=4)

    win.grab_set()
    self.root.wait_window(win)
    return out if out.get("ok") else None

def confirmar_entrega_ne(self):
    _ensure_configured()
    ne = self.get_ne_selected()
    if not ne:
        return
    if not ne.get("linhas"):
        messagebox.showerror("Erro", "A nota nao tem linhas.")
        return
    if "aprovad" not in str(ne.get("estado", "")).lower():
        if not messagebox.askyesno("Confirmar", "A nota ainda nao esta aprovada. Continuar assim mesmo?"):
            return
    entrega_data = self._dialog_confirmar_entrega_ne(ne)
    if not entrega_data:
        return
    selected_items = {}
    for it in entrega_data.get("itens", []) or []:
        try:
            idx = int(it.get("idx"))
        except Exception:
            continue
        selected_items[idx] = parse_float(it.get("qtd", 0), 0)
    if not selected_items:
        for i in entrega_data.get("indices", []) or []:
            try:
                idx = int(i)
                linha = ne.get("linhas", [])[idx]
            except Exception:
                continue
            selected_items[idx] = parse_float(linha.get("qtd", 0), 0)
    moved = 0
    moved_qtd = 0.0
    missing = []
    delivered_refs = []
    for idx, l in enumerate(ne.get("linhas", [])):
        if idx not in selected_items:
            continue
        qtd_total = parse_float(l.get("qtd", 0), 0)
        qtd_entregue = parse_float(
            l.get("qtd_entregue", l.get("qtd", 0) if l.get("entregue") else 0),
            0,
        )
        qtd_pendente = max(0.0, qtd_total - qtd_entregue)
        if qtd_pendente <= 1e-9:
            l["qtd_entregue"] = qtd_total
            l["entregue"] = True
            l["_stock_in"] = True
            continue
        qtd = min(parse_float(selected_items.get(idx, 0), 0), qtd_pendente)
        if qtd <= 0:
            continue
        origem = l.get("origem", "Produto")
        ref = l.get("ref", "")
        if origem_is_materia(origem):
            ok = self._receber_linha_materia_ne(l, ne.get("numero", ""), qtd_recebida=qtd)
            if not ok:
                missing.append(ref or l.get("descricao", ""))
                continue
            qtd_new = min(qtd_total, qtd_entregue + qtd)
            l["qtd_entregue"] = qtd_new
            l["entregue"] = qtd_new >= (qtd_total - 1e-9)
            l["_stock_in"] = bool(l.get("entregue"))
            l["guia_entrega"] = entrega_data.get("guia", "")
            l["fatura_entrega"] = entrega_data.get("fatura", "")
            l["data_doc_entrega"] = entrega_data.get("data_doc", "")
            l["data_entrega_real"] = entrega_data.get("data_entrega", "")
            l["obs_entrega"] = entrega_data.get("obs", "")
            moved += 1
            moved_qtd += qtd
            delivered_refs.append(f"{ref or l.get('descricao', '')} ({fmt_num(qtd)})")
            l.setdefault("entregas_linha", []).append(
                {
                    "data_registo": now_iso(),
                    "data_entrega": entrega_data.get("data_entrega", ""),
                    "data_documento": entrega_data.get("data_doc", ""),
                    "guia": entrega_data.get("guia", ""),
                    "fatura": entrega_data.get("fatura", ""),
                    "obs": entrega_data.get("obs", ""),
                    "qtd": qtd,
                }
            )
            continue
        p = next((x for x in self.data.get("produtos", []) if x.get("codigo") == ref), None)
        if not p:
            missing.append(ref)
            continue
        p["qty"] = parse_float(p.get("qty", 0), 0) + qtd
        p["atualizado_em"] = now_iso()
        qtd_new = min(qtd_total, qtd_entregue + qtd)
        l["qtd_entregue"] = qtd_new
        l["entregue"] = qtd_new >= (qtd_total - 1e-9)
        l["_stock_in"] = bool(l.get("entregue"))
        l["guia_entrega"] = entrega_data.get("guia", "")
        l["fatura_entrega"] = entrega_data.get("fatura", "")
        l["data_doc_entrega"] = entrega_data.get("data_doc", "")
        l["data_entrega_real"] = entrega_data.get("data_entrega", "")
        l["obs_entrega"] = entrega_data.get("obs", "")
        moved += 1
        moved_qtd += qtd
        delivered_refs.append(f"{ref or l.get('descricao', '')} ({fmt_num(qtd)})")
        l.setdefault("entregas_linha", []).append(
            {
                "data_registo": now_iso(),
                "data_entrega": entrega_data.get("data_entrega", ""),
                "data_documento": entrega_data.get("data_doc", ""),
                "guia": entrega_data.get("guia", ""),
                "fatura": entrega_data.get("fatura", ""),
                "obs": entrega_data.get("obs", ""),
                "qtd": qtd,
            }
        )
        self.data.setdefault("produtos_mov", []).append(
            {
                "data": now_iso(),
                "produto": ref,
                "descricao": l.get("descricao", ""),
                "operador": "Entrada NE",
                "tipo": "Entrada",
                "quantidade": qtd,
                "obs": ne.get("numero", ""),
            }
        )
    if missing:
        messagebox.showwarning("Aviso", "Itens nao encontrados para entrada:\n" + "\n".join(sorted(set(missing))))
    if moved == 0 and all(l.get("_stock_in") for l in ne.get("linhas", [])):
        messagebox.showinfo("Info", "Esta nota ja estava toda entregue.")
    all_done = all(
        parse_float(
            l.get("qtd_entregue", l.get("qtd", 0) if l.get("entregue") else 0),
            0,
        )
        >= (parse_float(l.get("qtd", 0), 0) - 1e-9)
        for l in ne.get("linhas", [])
    )
    any_done = any(
        parse_float(
            l.get("qtd_entregue", l.get("qtd", 0) if l.get("entregue") else 0),
            0,
        )
        > 0
        for l in ne.get("linhas", [])
    )
    if all_done:
        ne["estado"] = "Entregue"
        ne["data_entregue"] = entrega_data.get("data_entrega", "") or now_iso()
    elif any_done:
        ne["estado"] = "Parcialmente Entregue"
    else:
        ne["estado"] = "Aprovada"
    ne["data_ultima_entrega"] = entrega_data.get("data_entrega", "") or now_iso()
    ne["guia_ultima"] = entrega_data.get("guia", "")
    ne["fatura_ultima"] = entrega_data.get("fatura", "")
    ne["data_doc_ultima"] = entrega_data.get("data_doc", "")
    ne.setdefault("entregas", []).append(
        {
            "data_registo": now_iso(),
            "data_entrega": entrega_data.get("data_entrega", ""),
            "guia": entrega_data.get("guia", ""),
            "fatura": entrega_data.get("fatura", ""),
            "data_documento": entrega_data.get("data_doc", ""),
            "obs": entrega_data.get("obs", ""),
            "linhas": delivered_refs,
            "quantidade_linhas": moved,
            "quantidade_total": moved_qtd,
        }
    )
    ne["_draft"] = False
    ne["total"] = sum(parse_float(x.get("total", 0), 0) for x in ne.get("linhas", []))
    save_data(self.data)
    try:
        self.refresh_produtos()
    except Exception:
        pass
    self.refresh_ne()
    for iid in self.tbl_ne.get_children():
        if self.tbl_ne.item(iid, "values")[0] == ne.get("numero"):
            self.tbl_ne.selection_set(iid)
            self.tbl_ne.see(iid)
            break
    self.on_ne_select()
    if moved > 0:
        messagebox.showinfo("OK", f"Entrega confirmada. {moved} linhas lancadas em stock.")

def ne_blink_schedule(self):
    _ensure_configured()
    try:
        if hasattr(self, "nb") and hasattr(self, "tab_ne"):
            if str(self.nb.select()) != str(self.tab_ne):
                self.root.after(1400, self.ne_blink_schedule)
                return
        if not hasattr(self, "tbl_ne") or not self.tbl_ne.winfo_exists():
            return
        self.ne_blink_on = not getattr(self, "ne_blink_on", False)
        if self.ne_blink_on:
            self.tbl_ne.tag_configure("aprovada", background="#ffe08a", foreground="#7c5a00")
        else:
            self.tbl_ne.tag_configure("aprovada", background="#fff4cc", foreground="#8a6d00")
    except Exception:
        return
    self.root.after(900, self.ne_blink_schedule)

def render_ne_pdf(self, path, ne):
    _ensure_configured()
    return _render_ne_pdf_modern(self, path, ne)
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as pdf_canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.utils import ImageReader
    from reportlab.lib import colors

    c = pdf_canvas.Canvas(path, pagesize=A4)
    w, h = A4
    m = 20
    c.setTitle(f"Nota de Encomenda {ne.get('numero','')}")

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

    cols = [
        ("Designacao", 249, "w"),
        ("Qtd", 34, "e"),
        ("Un.", 24, "center"),
        ("Preco EUR", 62, "e"),
        ("Desc %", 36, "e"),
        ("IVA %", 34, "e"),
        ("Total EUR", 64, "e"),
        ("Entrega", 50, "center"),
    ]
    table_w = sum(cw for _, cw, _ in cols)
    row_h = 16
    header_row_h = 18
    lines = list(ne.get("linhas", []))

    subtotal = 0.0
    iva = 0.0
    total = 0.0
    for l in lines:
        qtd_l = parse_float(l.get("qtd", 0), 0)
        preco_l = parse_float(l.get("preco", 0), 0)
        desc_l = max(0.0, min(100.0, parse_float(l.get("desconto", 0), 0)))
        iva_l = max(0.0, min(100.0, parse_float(l.get("iva", 23), 23)))
        base_l = (qtd_l * preco_l) * (1.0 - (desc_l / 100.0))
        iva_amt = base_l * (iva_l / 100.0)
        total_l = parse_float(l.get("total", 0), 0)
        if total_l <= 0:
            total_l = base_l + iva_amt
        subtotal += base_l
        iva += iva_amt
        total += total_l

    # resolve fornecedor completo para cabecalho (nome + morada + contacto + nif + email)
    forn_obj = None
    forn_id = str(ne.get("fornecedor_id", "") or "").strip()
    forn_txt = str(ne.get("fornecedor", "") or "").strip()
    fornecedores = self.data.get("fornecedores", []) if isinstance(self.data, dict) else []
    if forn_id:
        forn_obj = next((f for f in fornecedores if str(f.get("id", "")).strip() == forn_id), None)
    if not forn_obj and " - " in forn_txt:
        pid, pnome = [x.strip() for x in forn_txt.split(" - ", 1)]
        forn_obj = next((f for f in fornecedores if str(f.get("id", "")).strip() == pid), None)
        if not forn_obj:
            forn_obj = next((f for f in fornecedores if str(f.get("nome", "")).strip().lower() == pnome.lower()), None)
    if not forn_obj and forn_txt:
        forn_obj = next((f for f in fornecedores if str(f.get("nome", "")).strip().lower() == forn_txt.lower()), None)
    if not forn_obj:
        forn_obj = {}

    forn_nome = (forn_obj.get("nome") or "").strip() or forn_txt
    forn_morada = (forn_obj.get("morada") or "").strip()
    forn_contacto = (forn_obj.get("contacto") or "").strip() or str(ne.get("contacto", "") or "").strip()
    forn_nif = (forn_obj.get("nif") or "").strip()
    forn_email = (forn_obj.get("email") or "").strip()

    def draw_page_header(page_num, full=True):
        c.setStrokeColor(colors.HexColor("#b8c7dd"))
        c.rect(m - 4, m - 4, w - (2 * (m - 4)), h - (2 * (m - 4)), stroke=1, fill=0)
        c.setStrokeColor(colors.black)

        logo_w = 170
        logo_h = 92
        logo_x = m - 6
        logo_top = 2
        draw_pdf_logo_box(
            c,
            h,
            logo_x,
            logo_top,
            box_size=logo_w,
            box_h=logo_h,
            padding=1,
            draw_border=False,
        )
        c.setFillColor(colors.HexColor("#0f2f66"))
        set_font(True, 14)
        rx = w - 210
        # Mantem apenas o bloco da direita com N. ENC / DATA para evitar repeticao.
        c.setFillColor(colors.black)

        # Dados da empresa no topo, ao lado do logo (evita concentrar tudo no rodape).
        emp_lines = get_empresa_rodape_lines()
        if emp_lines:
            info_x = logo_x + logo_w + 10
            info_y = h - 42
            c.setFillColor(colors.HexColor("#1d3557"))
            set_font(False, 7.6)
            for ln in emp_lines[:4]:
                c.drawString(info_x, info_y, clip_text(ln, 250, False, 7.6))
                info_y -= 9
            c.setFillColor(colors.black)

        c.setStrokeColor(colors.HexColor("#7f95ba"))
        c.setFillColor(colors.HexColor("#f4f8ff"))
        c.roundRect(rx, h - 48, 190, 14, 2, stroke=1, fill=1)
        c.roundRect(rx, h - 65, 190, 14, 2, stroke=1, fill=1)
        c.setFillColor(colors.black)
        set_font(True, 8.5)
        c.drawString(rx + 5, h - 44, f"N. ENC: {ne.get('numero','')}")
        c.drawString(rx + 5, h - 61, f"DATA: {datetime.now().strftime('%d/%m/%Y')}")
        set_font(False, 8)
        c.drawRightString(w - m, h - 76, f"Pag. {page_num}")

        if full:
            box_y = h - 144
            box_h = 72
            split_x = m + 360
            c.setStrokeColor(colors.HexColor("#7f95ba"))
            c.setFillColor(colors.HexColor("#f8fbff"))
            c.roundRect(m, box_y, w - (2 * m), box_h, 3, stroke=1, fill=1)
            c.line(split_x, box_y, split_x, box_y + box_h)
            c.setFillColor(colors.black)
            set_font(True, 9)
            c.drawString(m + 7, box_y + 55, "Fornecedor")
            set_font(False, 9)
            c.drawString(m + 7, box_y + 42, clip_text(forn_nome, split_x - m - 16, False, 9))
            if forn_morada:
                c.drawString(m + 7, box_y + 29, clip_text(f"Morada: {forn_morada}", split_x - m - 16, False, 8.5))
            c.drawString(m + 7, box_y + 16, clip_text(f"Contacto: {forn_contacto}", split_x - m - 16, False, 8.5))

            set_font(True, 8.5)
            c.drawString(split_x + 8, box_y + 55, "Dados Complementares")
            set_font(False, 8.5)
            c.drawString(split_x + 8, box_y + 42, clip_text(f"NIF: {forn_nif}", w - split_x - 16, False, 8.5))
            c.drawString(split_x + 8, box_y + 29, clip_text(f"Email: {forn_email}", w - split_x - 16, False, 8.5))
            c.drawString(split_x + 8, box_y + 16, clip_text(f"Entrega prevista: {ne.get('data_entrega','')}", w - split_x - 16, False, 8.5))

    def draw_table_header(y_top):
        c.setFillColor(colors.HexColor("#143f85"))
        c.rect(m, h - y_top - header_row_h, table_w, header_row_h, stroke=1, fill=1)
        c.setFillColor(colors.white)
        set_font(True, 8.5)
        xx = m
        for name, cw, align in cols:
            if align == "e":
                c.drawRightString(xx + cw - 3, h - y_top - 12, name)
            elif align == "center":
                c.drawCentredString(xx + cw / 2, h - y_top - 12, name)
            else:
                c.drawString(xx + 3, h - y_top - 12, name)
            xx += cw
        c.setFillColor(colors.black)

    def draw_bottom():
        by = 84
        docs_h = 56
        docs_y = by + 104
        c.setStrokeColor(colors.HexColor("#7f95ba"))
        c.roundRect(m, docs_y, table_w, docs_h, 3, stroke=1, fill=0)
        c.setFillColor(colors.HexColor("#f8fbff"))
        c.roundRect(m, by, 275, 90, 3, stroke=1, fill=1)
        c.roundRect(m + 283, by, table_w - 283, 90, 3, stroke=1, fill=1)
        c.setFillColor(colors.black)
        c.setStrokeColor(colors.black)

        set_font(True, 8.5)
        c.drawString(m + 6, docs_y + docs_h + 3, "Documentos de entrega")
        # grelha docs
        header_y = docs_y + docs_h - 16
        c.line(m, header_y, m + table_w, header_y)
        col_data = m + 78
        col_guia = m + 240
        col_fat = m + 410
        c.line(col_data, docs_y, col_data, docs_y + docs_h)
        c.line(col_guia, docs_y, col_guia, docs_y + docs_h)
        c.line(col_fat, docs_y, col_fat, docs_y + docs_h)
        set_font(True, 7.6)
        c.drawCentredString((m + col_data) / 2, docs_y + docs_h - 12, "Data")
        c.drawCentredString((col_data + col_guia) / 2, docs_y + docs_h - 12, "Guia")
        c.drawCentredString((col_guia + col_fat) / 2, docs_y + docs_h - 12, "Fatura")
        c.drawCentredString((col_fat + (m + table_w)) / 2, docs_y + docs_h - 12, "Obs.")

        set_font(False, 7.2)
        entregas = ne.get("entregas", []) or []
        if entregas:
            docs = list(reversed(entregas[-4:]))
            y_docs = docs_y + docs_h - 26
            for ent in docs:
                data_txt = clip_text(ent.get("data_entrega", ""), col_data - m - 6, False, 7.2)
                guia_txt = clip_text(ent.get("guia", ""), col_guia - col_data - 6, False, 7.2)
                fat_txt = clip_text(ent.get("fatura", ""), col_fat - col_guia - 6, False, 7.2)
                obs_txt = clip_text(ent.get("obs", ""), (m + table_w) - col_fat - 6, False, 7.2)
                c.drawString(m + 3, y_docs, data_txt)
                c.drawString(col_data + 3, y_docs, guia_txt)
                c.drawString(col_guia + 3, y_docs, fat_txt)
                c.drawString(col_fat + 3, y_docs, obs_txt)
                y_docs -= 9
        else:
            c.drawString(m + 6, docs_y + 12, "Sem entregas registadas")

        set_font(True, 9)
        c.drawString(m + 6, by + 74, "Condicoes")
        set_font(False, 8)
        c.drawString(m + 6, by + 60, "Pagamento: 60 dias")
        c.drawString(m + 6, by + 47, f"Observacoes: {clip_text(ne.get('obs', ''), 255, False, 8)}")
        c.drawString(m + 6, by + 34, f"Local de descarga: {clip_text(ne.get('local_descarga', ''), 255, False, 8)}")
        c.drawString(m + 6, by + 21, f"Meio de transporte: {clip_text(ne.get('meio_transporte', ''), 255, False, 8)}")

        set_font(True, 9)
        c.drawString(m + 289, by + 74, "Resumo")
        set_font(False, 9)
        c.drawString(m + 289, by + 58, "Mercadorias (s/ IVA):")
        c.drawRightString(m + table_w - 8, by + 58, f"{fmt_num(subtotal)} EUR")
        c.drawString(m + 289, by + 44, "IVA:")
        c.drawRightString(m + table_w - 8, by + 44, f"{fmt_num(iva)} EUR")
        set_font(True, 11)
        c.drawString(m + 289, by + 26, "TOTAL:")
        c.drawRightString(m + table_w - 8, by + 26, f"{fmt_num(total)} EUR")

        c.setStrokeColor(colors.HexColor("#8096bb"))
        sig_y = by - 16
        left_sig_x1 = m + 28
        left_sig_x2 = m + 248
        right_sig_x1 = m + 302
        right_sig_x2 = m + table_w - 20
        c.line(left_sig_x1, sig_y, left_sig_x2, sig_y)
        c.line(right_sig_x1, sig_y, right_sig_x2, sig_y)
        c.setFillColor(colors.HexColor("#445a7e"))
        set_font(False, 7.3)
        c.drawCentredString((left_sig_x1 + left_sig_x2) / 2.0, sig_y - 11, "Assinatura Responsavel")
        c.drawCentredString((right_sig_x1 + right_sig_x2) / 2.0, sig_y - 11, "Assinatura Aprovacao")
        c.setFillColor(colors.black)
        c.drawRightString(w - m, 22, "Software luGEST - Processado por computador")

    def rows_fit(first_page, reserve_bottom):
        table_top = 174 if first_page else 70
        usable = h - m - reserve_bottom - (table_top + header_row_h)
        return max(0, int(usable // row_h))

    idx = 0
    page = 1
    while idx < len(lines) or idx == 0:
        first = page == 1
        fit_last = rows_fit(first, 220)
        fit_non = max(1, rows_fit(first, 36))
        remaining = len(lines) - idx
        if remaining <= fit_last:
            is_last = True
            count = remaining
        else:
            is_last = False
            count = fit_non

        draw_page_header(page, full=first)
        table_top = 174 if first else 70
        draw_table_header(table_top)

        set_font(False, 8)
        y_top = table_top + header_row_h
        for local_i in range(count):
            l = lines[idx + local_i]
            y_row = y_top + local_i * row_h
            fill = colors.HexColor("#f9fbff") if ((idx + local_i) % 2 == 0) else colors.HexColor("#edf3ff")
            c.setFillColor(fill)
            c.rect(m, h - y_row - row_h, table_w, row_h, stroke=1, fill=1)
            c.setFillColor(colors.black)

            desc_l = parse_float(l.get("desconto", 0), 0)
            iva_l = parse_float(l.get("iva", 23), 23)
            qtd_ent_l = parse_float(l.get("qtd_entregue", l.get("qtd", 0) if l.get("entregue") else 0), 0)
            qtd_tot_l = parse_float(l.get("qtd", 0), 0)
            if qtd_ent_l <= 0:
                ent_txt = "PENDENTE"
            elif qtd_ent_l < (qtd_tot_l - 1e-9):
                ent_txt = "PARCIAL"
            else:
                ent_txt = "ENTREGUE"
            vals = [
                l.get("descricao", ""),
                fmt_num(l.get("qtd", 0)),
                l.get("unid", "UN"),
                fmt_num(l.get("preco", 0)),
                fmt_num(desc_l),
                fmt_num(iva_l),
                fmt_num(l.get("total", 0)),
                ent_txt,
            ]
            xx = m
            for (_, cw, align), val in zip(cols, vals):
                txt = clip_text(val, cw - 6, False, 8)
                if align == "e":
                    c.drawRightString(xx + cw - 3, h - y_row - 11, txt)
                elif align == "center":
                    c.drawCentredString(xx + cw / 2, h - y_row - 11, txt)
                else:
                    c.drawString(xx + 3, h - y_row - 11, txt)
                xx += cw

        idx += count
        if is_last:
            draw_bottom()
            break
        c.showPage()
        page += 1

    c.save()


def _render_ne_pdf_modern(self, path, ne):
    _ensure_configured()
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as pdf_canvas

    fonts = _pdf_register_fonts()
    palette = _pdf_brand_palette()
    c = pdf_canvas.Canvas(path, pagesize=A4)
    w, h = A4
    m = 24
    content_w = w - (2 * m)
    header_row_h = 20
    row_h = 18
    table_first_top = 252
    table_next_top = 92
    footer_top = 622
    doc_num = str(ne.get("numero", "") or "NE").strip() or "NE"
    c.setTitle(f"Nota de Encomenda {doc_num}")

    cols = [
        ("Ref.", 68, "w"),
        ("Designacao", 171, "w"),
        ("Qtd.", 38, "e"),
        ("Un.", 30, "center"),
        ("Preco", 58, "e"),
        ("Desc%", 36, "e"),
        ("IVA%", 34, "e"),
        ("Total", 64, "e"),
        ("Estado", 48, "center"),
    ]
    table_w = sum(width for _, width, _ in cols)
    first_capacity = max(1, int((footer_top - (table_first_top + header_row_h + 6)) // row_h))
    next_capacity = max(1, int((footer_top - (table_next_top + header_row_h + 6)) // row_h))

    subtotal = 0.0
    iva = 0.0
    total = 0.0
    for line in list(ne.get("linhas", []) or []):
        qtd_l = parse_float(line.get("qtd", 0), 0)
        preco_l = parse_float(line.get("preco", 0), 0)
        desc_l = max(0.0, min(100.0, parse_float(line.get("desconto", 0), 0)))
        iva_l = max(0.0, min(100.0, parse_float(line.get("iva", 23), 23)))
        base_l = (qtd_l * preco_l) * (1.0 - (desc_l / 100.0))
        iva_amt = base_l * (iva_l / 100.0)
        total_l = parse_float(line.get("total", 0), 0)
        if total_l <= 0:
            total_l = base_l + iva_amt
        subtotal += base_l
        iva += iva_amt
        total += total_l

    forn_obj = None
    forn_id = str(ne.get("fornecedor_id", "") or "").strip()
    forn_txt = str(ne.get("fornecedor", "") or "").strip()
    fornecedores = self.data.get("fornecedores", []) if isinstance(self.data, dict) else []
    if forn_id:
        forn_obj = next((row for row in fornecedores if str(row.get("id", "")).strip() == forn_id), None)
    if not forn_obj and " - " in forn_txt:
        pid, pnome = [chunk.strip() for chunk in forn_txt.split(" - ", 1)]
        forn_obj = next((row for row in fornecedores if str(row.get("id", "")).strip() == pid), None)
        if not forn_obj:
            forn_obj = next(
                (row for row in fornecedores if str(row.get("nome", "")).strip().lower() == pnome.lower()),
                None,
            )
    if not forn_obj and forn_txt:
        forn_obj = next(
            (row for row in fornecedores if str(row.get("nome", "")).strip().lower() == forn_txt.lower()),
            None,
        )
    if not forn_obj:
        forn_obj = {}

    forn_nome = str(forn_obj.get("nome", "") or forn_txt).strip() or "-"
    forn_morada = str(forn_obj.get("morada", "") or "").strip()
    forn_contacto = str(forn_obj.get("contacto", "") or ne.get("contacto", "") or "").strip()
    forn_nif = str(forn_obj.get("nif", "") or "").strip()
    forn_email = str(forn_obj.get("email", "") or "").strip()
    pagamento = str(forn_obj.get("cond_pagamento", "") or "60 dias").strip()
    entrega_prevista = str(ne.get("data_entrega", "") or "").strip()
    estado_doc = str(ne.get("estado", "") or "Em edicao").strip() or "Em edicao"
    docs = list(reversed(list(ne.get("entregas", []) or [])[-4:]))
    lines = list(ne.get("linhas", []) or [])
    footer_company = list(get_empresa_rodape_lines() or [])

    def yinv(top_y):
        return h - top_y

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

    def delivery_status(line):
        qtd_ent = parse_float(line.get("qtd_entregue", line.get("qtd", 0) if line.get("entregue") else 0), 0)
        qtd_tot = parse_float(line.get("qtd", 0), 0)
        if qtd_ent <= 1e-9:
            return "Pendente"
        if qtd_ent < (qtd_tot - 1e-9):
            return "Parcial"
        return "Entregue"

    def card(x, top_y, box_w, box_h, title, lines_, accent=False):
        title_fill = palette["primary_soft"] if accent else palette["primary_soft_2"]
        c.saveState()
        c.setFillColor(colors.white)
        c.setStrokeColor(palette["line_strong"] if accent else palette["line"])
        c.setLineWidth(0.9)
        c.roundRect(x, yinv(top_y + box_h), box_w, box_h, 8, stroke=1, fill=1)
        c.setFillColor(title_fill)
        c.roundRect(x, yinv(top_y + 20), box_w, 20, 8, stroke=0, fill=1)
        c.restoreState()
        c.setFont(fonts["bold"], 8.7)
        c.setFillColor(palette["primary_dark"])
        c.drawString(x + 8, yinv(top_y + 13), ntxt(title))
        yy = top_y + 34
        for idx_line, line in enumerate(lines_):
            font_name = fonts["bold"] if idx_line == 0 and accent else fonts["regular"]
            font_size = 9.6 if idx_line == 0 and accent else 8.6
            wrapped = _pdf_wrap_text(line, font_name, font_size, box_w - 16, max_lines=2 if idx_line == 0 else 1)
            for item in wrapped:
                c.setFont(font_name, font_size)
                c.setFillColor(palette["ink"])
                c.drawString(x + 8, yinv(yy), ntxt(item))
                yy += 11
            if yy > top_y + box_h - 10:
                break

    def info_chip(x, top_y, box_w, label, value, box_h=34):
        c.saveState()
        c.setFillColor(colors.white)
        c.setStrokeColor(palette["line"])
        c.roundRect(x, yinv(top_y + box_h), box_w, box_h, 8, stroke=1, fill=1)
        c.restoreState()
        c.setFont(fonts["regular"], 7.2)
        c.setFillColor(palette["muted"])
        c.drawString(x + 7, yinv(top_y + 11), ntxt(label))
        c.setFont(fonts["bold"], 9.7)
        c.setFillColor(palette["primary_dark"])
        c.drawString(x + 7, yinv(top_y + 23.5), ntxt(_pdf_clip_text(value, box_w - 14, fonts["bold"], 9.7)))

    def draw_table_header(top_y):
        c.saveState()
        c.setFillColor(palette["primary"])
        c.roundRect(m, yinv(top_y + header_row_h), table_w, header_row_h, 7, stroke=0, fill=1)
        c.restoreState()
        c.setFillColor(colors.white)
        c.setFont(fonts["bold"], 8.15)
        xx = m
        for name, width, align in cols:
            if align == "e":
                c.drawRightString(xx + width - 7, yinv(top_y + 13), ntxt(name))
            elif align == "center":
                c.drawCentredString(xx + (width / 2.0), yinv(top_y + 13), ntxt(name))
            else:
                c.drawString(xx + 7, yinv(top_y + 13), ntxt(name))
            xx += width
        return top_y + header_row_h + 6

    def draw_row(top_y, idx_line, line):
        fill = palette["surface_warm"] if idx_line % 2 == 0 else colors.white
        c.saveState()
        c.setFillColor(fill)
        c.setStrokeColor(palette["line"])
        c.setLineWidth(0.45)
        c.roundRect(m, yinv(top_y + row_h), table_w, row_h, 5, stroke=1, fill=1)
        c.restoreState()

        values = [
            str(line.get("ref", "") or "").strip() or "-",
            str(line.get("descricao", "") or "").strip() or "-",
            fmt_num(line.get("qtd", 0)),
            str(line.get("unid", "UN") or "UN").strip(),
            fmt_num(line.get("preco", 0)),
            fmt_num(line.get("desconto", 0)),
            fmt_num(line.get("iva", 23)),
            fmt_num(line.get("total", 0)),
            delivery_status(line),
        ]
        xx = m
        c.setFont(fonts["regular"], 8.1)
        c.setFillColor(palette["ink"])
        c.setStrokeColor(palette["line"])
        for idx_col, ((_, width, align), value) in enumerate(zip(cols, values)):
            if idx_col > 0:
                c.line(xx, yinv(top_y), xx, yinv(top_y + row_h))
            txt = _pdf_clip_text(value, width - 14, fonts["regular"], 8.1)
            if align == "e":
                c.drawRightString(xx + width - 7, yinv(top_y + 11.5), ntxt(txt))
            elif align == "center":
                if idx_col == len(cols) - 1:
                    if value == "Entregue":
                        c.setFillColor(palette["success"])
                    elif value == "Parcial":
                        c.setFillColor(palette["primary_dark"])
                    else:
                        c.setFillColor(palette["danger"])
                    c.setFont(fonts["bold"], 7.8)
                c.drawCentredString(xx + (width / 2.0), yinv(top_y + 11.5), ntxt(txt))
                c.setFillColor(palette["ink"])
                c.setFont(fonts["regular"], 8.1)
            else:
                c.drawString(xx + 7, yinv(top_y + 11.5), ntxt(txt))
            xx += width

    def draw_docs_card(top_y, box_h=66):
        c.saveState()
        c.setFillColor(colors.white)
        c.setStrokeColor(palette["line"])
        c.roundRect(m, yinv(top_y + box_h), table_w, box_h, 8, stroke=1, fill=1)
        c.setFillColor(palette["primary_soft_2"])
        c.roundRect(m, yinv(top_y + 20), table_w, 20, 8, stroke=0, fill=1)
        c.restoreState()
        c.setFont(fonts["bold"], 8.7)
        c.setFillColor(palette["primary_dark"])
        c.drawString(m + 8, yinv(top_y + 13), ntxt("Documentos de entrega"))
        headers = [("Data", 70), ("Guia", 112), ("Fatura", 112), ("Observacoes", table_w - 294)]
        xx = m + 8
        header_y = top_y + 30
        c.setFont(fonts["bold"], 7.5)
        c.setFillColor(palette["muted"])
        for name, width in headers:
            c.drawString(xx, yinv(header_y), ntxt(name))
            xx += width
        line_y = top_y + 34
        c.setStrokeColor(palette["line"])
        c.line(m + 8, yinv(line_y), m + table_w - 8, yinv(line_y))
        if not docs:
            c.setFont(fonts["regular"], 8)
            c.setFillColor(palette["ink"])
            c.drawString(m + 8, yinv(top_y + 50), ntxt("Sem entregas registadas."))
            return
        yy = top_y + 48
        c.setFont(fonts["regular"], 7.6)
        for entry in docs[:3]:
            xx = m + 8
            values = [
                fmt_display_date(entry.get("data_entrega", "")),
                str(entry.get("guia", "") or "-").strip(),
                str(entry.get("fatura", "") or "-").strip(),
                str(entry.get("obs", "") or "").strip() or "-",
            ]
            for (name, width), value in zip(headers, values):
                c.drawString(xx, yinv(yy), ntxt(_pdf_clip_text(value, width - 8, fonts["regular"], 7.6)))
                xx += width
            yy += 10

    def draw_header(page_no, total_pages, first_page):
        if first_page:
            c.saveState()
            c.setFillColor(palette["primary"])
            c.roundRect(m, yinv(102), content_w, 82, 12, stroke=0, fill=1)
            c.restoreState()
            logo_x = m + 12
            logo_w = 118
            metric_grid = _pdf_metric_grid_layout(w - m, 20, 82, group_w=230, gap=6)
            draw_pdf_logo_box(c, h, logo_x, 30, box_size=logo_w, box_h=52, padding=4, draw_border=False)
            c.setFillColor(colors.white)
            _pdf_draw_banner_title_block(
                c,
                h,
                "Nota de Encomenda",
                "Documento de adjudicacao ao fornecedor",
                logo_x + logo_w + 8,
                metric_grid["col1"] - 10,
                47,
                64,
                fonts["bold"],
                fonts["regular"],
                title_size=17.0,
                subtitle_size=8.6,
                min_title_size=14.0,
                min_subtitle_size=7.2,
            )
            info_chip(metric_grid["col1"], metric_grid["row1"], metric_grid["chip_w"], "Documento", doc_num, box_h=metric_grid["chip_h"])
            info_chip(metric_grid["col2"], metric_grid["row1"], metric_grid["chip_w"], "Data", datetime.now().strftime("%d/%m/%Y"), box_h=metric_grid["chip_h"])
            info_chip(metric_grid["col1"], metric_grid["row2"], metric_grid["chip_w"], "Estado", estado_doc, box_h=metric_grid["chip_h"])
            info_chip(metric_grid["col2"], metric_grid["row2"], metric_grid["chip_w"], "Pagina", f"{page_no}/{total_pages}", box_h=metric_grid["chip_h"])
            card(
                m,
                126,
                288,
                100,
                "Fornecedor",
                [
                    forn_nome,
                    f"NIF: {forn_nif or '-'}",
                    forn_morada or "-",
                    f"Contacto: {forn_contacto or '-'}",
                    f"Email: {forn_email or '-'}",
                ],
                accent=True,
            )
            card(
                m + 300,
                126,
                content_w - 300,
                100,
                "Dados da encomenda",
                [
                    f"Entrega prevista: {fmt_display_date(entrega_prevista)}",
                    f"Local de descarga: {str(ne.get('local_descarga', '') or '-').strip()}",
                    f"Meio de transporte: {str(ne.get('meio_transporte', '') or '-').strip()}",
                    f"Ultima guia/fatura: {str(ne.get('guia_ultima', '') or '-').strip()} / {str(ne.get('fatura_ultima', '') or '-').strip()}",
                ],
            )
            return draw_table_header(table_first_top)

        c.saveState()
        c.setFillColor(palette["primary_soft"])
        c.roundRect(m, yinv(70), content_w, 50, 12, stroke=0, fill=1)
        c.restoreState()
        c.setFillColor(palette["primary_dark"])
        c.setFont(fonts["bold"], 15)
        c.drawString(m + 12, yinv(42), ntxt(f"Nota de Encomenda {doc_num}"))
        c.setFont(fonts["regular"], 8.6)
        c.drawString(m + 12, yinv(56), ntxt(f"Fornecedor: {forn_nome}"))
        info_chip(w - m - 182, 20, 84, "Estado", estado_doc)
        info_chip(w - m - 92, 20, 92, "Pagina", f"{page_no}/{total_pages}")
        return draw_table_header(table_next_top)

    def draw_footer(page_no, total_pages):
        draw_docs_card(630, box_h=66)
        card(
            m,
            708,
            332,
            72,
            "Condicoes",
            [
                f"Pagamento: {pagamento}",
                f"Observacoes: {str(ne.get('obs', '') or '-').strip()}",
                f"Local de descarga: {str(ne.get('local_descarga', '') or '-').strip()}",
                f"Meio de transporte: {str(ne.get('meio_transporte', '') or '-').strip()}",
            ],
        )
        card(
            m + 344,
            708,
            table_w - 344,
            72,
            "Resumo",
            [
                f"Subtotal: {fmt_money(subtotal)}",
                f"IVA: {fmt_money(iva)}",
                f"Total: {fmt_money(total)}",
                f"Linhas: {len(lines)}",
            ],
            accent=True,
        )
        c.saveState()
        c.setFillColor(palette["surface_alt"])
        c.roundRect(m, yinv(826), table_w, 28, 10, stroke=0, fill=1)
        c.restoreState()
        c.setFont(fonts["regular"], 7.3)
        c.setFillColor(palette["ink"])
        left_text = footer_company[0] if footer_company else ""
        mid_text = footer_company[1] if len(footer_company) > 1 else ""
        c.drawString(m + 10, yinv(806), ntxt(left_text))
        c.drawString(m + 10, yinv(816), ntxt(mid_text))
        c.drawRightString(w - m - 10, yinv(806), ntxt(f"Software luGEST | Pagina {page_no}/{total_pages}"))
        c.drawRightString(w - m - 10, yinv(816), ntxt(f"Documento: {doc_num}"))

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
            c.drawString(m + 8, yinv(table_y + 16), ntxt("Sem linhas na nota de encomenda."))
        else:
            y_row = table_y
            for idx_line, line in enumerate(page_rows):
                draw_row(y_row, idx_line, line)
                y_row += row_h
        if page_no == total_pages:
            draw_footer(page_no, total_pages)
        if page_no < total_pages:
            c.showPage()

    c.save()


def preview_ne(self):
    _ensure_configured()
    ne = self.get_ne_selected()
    if not ne:
        return
    prev_dir = os.path.join(BASE_DIR, "previews")
    try:
        os.makedirs(prev_dir, exist_ok=True)
    except Exception:
        prev_dir = tempfile.gettempdir()
    tmp = os.path.join(prev_dir, f"lugest_ne_{ne.get('numero','')}.pdf")
    try:
        self.render_ne_pdf(tmp, ne)
        opened = self._open_pdf_default(tmp)
        if not opened:
            messagebox.showwarning("Aviso", f"PDF gerado mas nao abriu automaticamente.\nCaminho:\n{tmp}")
    except Exception as ex:
        messagebox.showerror("Erro", f"Falha ao pre-visualizar PDF: {ex}")

def save_ne_pdf(self):
    _ensure_configured()
    ne = self.get_ne_selected()
    if not ne:
        return
    path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")], initialfile=f"{ne.get('numero','NE')}.pdf")
    if not path:
        return
    try:
        self.render_ne_pdf(path, ne)
        messagebox.showinfo("OK", "PDF guardado")
    except Exception as ex:
        messagebox.showerror("Erro", f"Falha ao guardar PDF: {ex}")

def open_ne_pdf_with(self):
    _ensure_configured()
    ne = self.get_ne_selected()
    if not ne:
        return
    tmp = os.path.join(tempfile.gettempdir(), f"lugest_ne_{ne.get('numero','')}.pdf")
    try:
        self.render_ne_pdf(tmp, ne)
        subprocess.Popen(["rundll32.exe", "shell32.dll,OpenAs_RunDLL", tmp])
    except Exception as ex:
        messagebox.showerror("Erro", f"Falha ao abrir com: {ex}")

def render_ne_cotacao_pdf(self, path, ne):
    _ensure_configured()
    return _render_ne_cotacao_pdf_modern(self, path, ne)
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as pdf_canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.utils import ImageReader
    from reportlab.lib import colors

    c = pdf_canvas.Canvas(path, pagesize=A4)
    w, h = A4
    m = 20
    c.setTitle(f"Pedido de Cotacao {ne.get('numero','')}")

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

    cols = [
        ("Codigo", 70, "w"),
        ("Descricao", 220, "w"),
        ("Qtd", 40, "e"),
        ("Un.", 32, "center"),
        ("Entrega", 70, "center"),
        ("Preco Cot. (EUR)", 55, "e"),
        ("Total Cot. (EUR)", 58, "e"),
    ]
    table_w = sum(cw for _, cw, _ in cols)
    row_h = 17
    header_row_h = 18
    lines = list(ne.get("linhas", []))

    def draw_page_header(page_num, full=True):
        c.setStrokeColor(colors.HexColor("#c6cfdb"))
        c.rect(m - 4, m - 4, w - (2 * (m - 4)), h - (2 * (m - 4)), stroke=1, fill=0)
        c.setStrokeColor(colors.black)

        logo_w = 170
        logo_h = 92
        logo_x = m - 6
        logo_top = 2
        draw_pdf_logo_box(
            c,
            h,
            logo_x,
            logo_top,
            box_size=logo_w,
            box_h=logo_h,
            padding=1,
            draw_border=False,
        )
        c.setFillColor(colors.HexColor("#8b1e2d"))
        set_font(True, 16)
        title_left = logo_x + logo_w + 10
        title_right = w - m - 120
        title_x = (title_left + title_right) / 2.0 if title_right > title_left else title_left
        title_y = h - 62
        c.drawCentredString(title_x, title_y, "Pedido de Cotacao")
        c.setFillColor(colors.black)
        set_font(True, 10)
        c.drawRightString(w - m, h - 36, f"N. {ne.get('numero','')}")
        set_font(False, 9)
        c.drawRightString(w - m, h - 50, f"Data: {datetime.now().strftime('%d/%m/%Y')}")
        if page_num > 1:
            c.drawRightString(w - m, h - 64, f"Pagina {page_num}")

        if full:
            box_y = h - 134
            c.rect(m, box_y, w - (2 * m), 62, stroke=1, fill=0)
            set_font(True, 9)
            c.drawString(m + 7, box_y + 45, "Fornecedor")
            set_font(False, 9)
            c.drawString(m + 7, box_y + 31, clip_text(ne.get("fornecedor", ""), 320, False, 9))
            c.drawString(m + 7, box_y + 17, f"Contacto: {ne.get('contacto','')}")
            c.drawRightString(w - m - 8, box_y + 17, f"Entrega prevista: {ne.get('data_entrega','')}")

            set_font(False, 8)
            c.drawString(m, box_y - 12, "Documento para pedido de precos ao fornecedor (sem valores internos).")

    def draw_table_header(y_top):
        c.setFillColor(colors.HexColor("#8b1e2d"))
        c.rect(m, h - y_top - header_row_h, table_w, header_row_h, stroke=1, fill=1)
        c.setFillColor(colors.white)
        set_font(True, 8.5)
        xx = m
        for name, cw, align in cols:
            if align == "e":
                c.drawRightString(xx + cw - 3, h - y_top - 12, name)
            elif align == "center":
                c.drawCentredString(xx + cw / 2, h - y_top - 12, name)
            else:
                c.drawString(xx + 3, h - y_top - 12, name)
            xx += cw
        c.setFillColor(colors.black)

    def draw_bottom():
        by = 78
        c.setStrokeColor(colors.HexColor("#c6cfdb"))
        c.rect(m, by, table_w, 70, stroke=1, fill=0)
        c.setStrokeColor(colors.black)
        set_font(True, 9)
        c.drawString(m + 6, by + 52, "Observacoes para cotacao")
        set_font(False, 8.5)
        c.drawString(m + 6, by + 38, "- Prazo de entrega:")
        c.drawString(m + 6, by + 24, "- Condicoes de pagamento:")
        c.drawString(m + 6, by + 10, "- Observacoes:")
        c.drawRightString(w - m, 34, f"Total de linhas: {len(lines)}")

    def rows_fit(first_page, reserve_bottom):
        table_top = 170 if first_page else 70
        usable = h - m - reserve_bottom - (table_top + header_row_h)
        return max(0, int(usable // row_h))

    idx = 0
    page = 1
    while idx < len(lines) or idx == 0:
        first = page == 1
        fit_last = rows_fit(first, 154)
        fit_non = max(1, rows_fit(first, 36))
        remaining = len(lines) - idx
        if remaining <= fit_last:
            is_last = True
            count = remaining
        else:
            is_last = False
            count = fit_non

        draw_page_header(page, full=first)
        table_top = 170 if first else 70
        draw_table_header(table_top)

        y_top = table_top + header_row_h
        set_font(False, 8)
        for local_i in range(count):
            l = lines[idx + local_i]
            y_row = y_top + local_i * row_h
            fill = colors.HexColor("#fff8f9") if ((idx + local_i) % 2 == 0) else colors.HexColor("#fff0f2")
            c.setFillColor(fill)
            c.rect(m, h - y_row - row_h, table_w, row_h, stroke=1, fill=1)
            c.setFillColor(colors.black)

            vals = [
                l.get("ref", ""),
                l.get("descricao", ""),
                fmt_num(l.get("qtd", 0)),
                l.get("unid", "UN"),
                ne.get("data_entrega", ""),
                "",
                "",
            ]
            xx = m
            for (_, cw, align), val in zip(cols, vals):
                txt = clip_text(val, cw - 6, False, 8)
                if align == "e":
                    c.drawRightString(xx + cw - 3, h - y_row - 11, txt)
                elif align == "center":
                    c.drawCentredString(xx + cw / 2, h - y_row - 11, txt)
                else:
                    c.drawString(xx + 3, h - y_row - 11, txt)
                xx += cw

        idx += count
        if is_last:
            draw_bottom()
            break
        c.showPage()
        page += 1

    c.save()


def _render_ne_cotacao_pdf_modern(self, path, ne):
    _ensure_configured()
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as pdf_canvas

    fonts = _pdf_register_fonts()
    palette = _pdf_brand_palette()
    c = pdf_canvas.Canvas(path, pagesize=A4)
    w, h = A4
    m = 23
    content_w = w - (2 * m)
    row_h = 18
    header_row_h = 20
    footer_top = 648
    table_first_top = 246
    table_next_top = 92
    rodape_lines = list(get_empresa_rodape_lines() or [])
    linhas = list(ne.get("linhas", []) or [])
    doc_num = str(ne.get("numero", "") or "NE").strip() or "NE"
    fornecedor = str(ne.get("fornecedor", "") or "").strip()
    contacto = str(ne.get("contacto", "") or "").strip()
    entrega = str(ne.get("data_entrega", "") or "").strip()
    observacoes = str(ne.get("obs", "") or "").strip()
    data_doc = datetime.now().strftime("%d/%m/%Y")

    cols = [
        ("Referencia", 78, "w"),
        ("Descricao", 195, "w"),
        ("Qtd.", 42, "e"),
        ("Un.", 34, "center"),
        ("Entrega", 68, "center"),
        ("Preco Unit.", 62, "e"),
        ("Total", 70, "e"),
    ]
    table_w = sum(width for _, width, _ in cols)
    first_capacity = max(1, int((footer_top - (table_first_top + header_row_h + 6)) // row_h))
    next_capacity = max(1, int((footer_top - (table_next_top + header_row_h + 6)) // row_h))

    def yinv(top_y):
        return h - top_y

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

    def card(x, top_y, box_w, box_h, title, lines):
        c.saveState()
        c.setFillColor(colors.white)
        c.setStrokeColor(palette["line"])
        c.setLineWidth(0.9)
        c.roundRect(x, yinv(top_y + box_h), box_w, box_h, 8, stroke=1, fill=1)
        c.setFillColor(palette["primary_soft_2"])
        c.roundRect(x, yinv(top_y + 20), box_w, 20, 8, stroke=0, fill=1)
        c.restoreState()
        c.setFont(fonts["bold"], 8.7)
        c.setFillColor(palette["primary_dark"])
        c.drawString(x + 8, yinv(top_y + 13), ntxt(title))
        yy = top_y + 34
        for idx_line, line in enumerate(lines):
            font_name = fonts["bold"] if idx_line == 0 else fonts["regular"]
            font_size = 9.6 if idx_line == 0 else 8.7
            wrapped = _pdf_wrap_text(line, font_name, font_size, box_w - 16, max_lines=2 if idx_line == 0 else 2)
            for item in wrapped:
                c.setFont(font_name, font_size)
                c.setFillColor(palette["ink"])
                c.drawString(x + 8, yinv(yy), ntxt(item))
                yy += 11
            if yy > top_y + box_h - 10:
                break

    def info_chip(x, top_y, box_w, label, value, box_h=34):
        c.saveState()
        c.setFillColor(colors.white)
        c.setStrokeColor(palette["line"])
        c.roundRect(x, yinv(top_y + box_h), box_w, box_h, 8, stroke=1, fill=1)
        c.restoreState()
        c.setFont(fonts["regular"], 7.2)
        c.setFillColor(palette["muted"])
        c.drawString(x + 7, yinv(top_y + 11), ntxt(label))
        c.setFont(fonts["bold"], 9.8)
        c.setFillColor(palette["primary_dark"])
        c.drawString(x + 7, yinv(top_y + 23.5), ntxt(_pdf_clip_text(value, box_w - 14, fonts["bold"], 9.8)))

    def draw_table_header(top_y):
        c.saveState()
        c.setFillColor(palette["primary"])
        c.roundRect(m, yinv(top_y + header_row_h), table_w, header_row_h, 7, stroke=0, fill=1)
        c.restoreState()
        c.setFillColor(colors.white)
        c.setFont(fonts["bold"], 8.2)
        xx = m
        for name, width, align in cols:
            if align == "e":
                c.drawRightString(xx + width - 7, yinv(top_y + 13), ntxt(name))
            elif align == "center":
                c.drawCentredString(xx + (width / 2.0), yinv(top_y + 13), ntxt(name))
            else:
                c.drawString(xx + 7, yinv(top_y + 13), ntxt(name))
            xx += width
        return top_y + header_row_h + 6

    def draw_row(top_y, idx_line, line):
        fill = palette["surface_warm"] if idx_line % 2 == 0 else colors.white
        c.saveState()
        c.setFillColor(fill)
        c.setStrokeColor(palette["line"])
        c.setLineWidth(0.45)
        c.roundRect(m, yinv(top_y + row_h), table_w, row_h, 5, stroke=1, fill=1)
        c.restoreState()
        values = [
            str(line.get("ref", "") or "").strip(),
            str(line.get("descricao", "") or "").strip() or "-",
            fmt_num(line.get("qtd", 0)),
            str(line.get("unid", "UN") or "UN").strip(),
            fmt_display_date(entrega),
            "",
            "",
        ]
        xx = m
        c.setFont(fonts["regular"], 8.3)
        c.setFillColor(palette["ink"])
        c.setStrokeColor(palette["line"])
        for idx_col, ((_, width, align), value) in enumerate(zip(cols, values)):
            if idx_col > 0:
                c.line(xx, yinv(top_y), xx, yinv(top_y + row_h))
            txt = _pdf_clip_text(value, width - 14, fonts["regular"], 8.3)
            if align == "e":
                c.drawRightString(xx + width - 7, yinv(top_y + 11.5), ntxt(txt))
            elif align == "center":
                c.drawCentredString(xx + (width / 2.0), yinv(top_y + 11.5), ntxt(txt))
            else:
                c.drawString(xx + 7, yinv(top_y + 11.5), ntxt(txt))
            xx += width

    def draw_header(page_no, total_pages, first_page):
        if first_page:
            c.saveState()
            c.setFillColor(palette["primary"])
            c.roundRect(m, yinv(102), content_w, 82, 12, stroke=0, fill=1)
            c.restoreState()
            logo_x = m + 12
            logo_w = 118
            metric_grid = _pdf_metric_grid_layout(w - m, 20, 82, group_w=230, gap=6)
            draw_pdf_logo_box(c, h, logo_x, 30, box_size=logo_w, box_h=52, padding=4, draw_border=False)
            c.setFillColor(colors.white)
            _pdf_draw_banner_title_block(
                c,
                h,
                "Pedido de Cotacao",
                "Documento para consulta comercial ao fornecedor",
                logo_x + logo_w + 8,
                metric_grid["col1"] - 10,
                47,
                64,
                fonts["bold"],
                fonts["regular"],
                title_size=16.8,
                subtitle_size=8.6,
                min_title_size=13.8,
                min_subtitle_size=7.2,
            )
            info_chip(metric_grid["col1"], metric_grid["row1"], metric_grid["chip_w"], "Documento", doc_num, box_h=metric_grid["chip_h"])
            info_chip(metric_grid["col2"], metric_grid["row1"], metric_grid["chip_w"], "Data", data_doc, box_h=metric_grid["chip_h"])
            info_chip(metric_grid["col1"], metric_grid["row2"], metric_grid["chip_w"], "Pagina", f"{page_no}/{total_pages}", box_h=metric_grid["chip_h"])
            info_chip(metric_grid["col2"], metric_grid["row2"], metric_grid["chip_w"], "Entrega", fmt_display_date(entrega), box_h=metric_grid["chip_h"])
            card(
                m,
                126,
                268,
                92,
                "Fornecedor",
                [
                    fornecedor or "-",
                    f"Contacto: {contacto or '-'}",
                    f"Entrega prevista: {fmt_display_date(entrega)}",
                ],
            )
            card(
                m + 282,
                126,
                content_w - 282,
                92,
                "Resposta pretendida",
                [
                    "Preencher preco unitario, total e prazo de entrega.",
                    "Indicar validade da proposta e eventuais condicoes comerciais.",
                    "Responder para os contactos indicados no rodape.",
                ],
            )
            c.setFont(fonts["regular"], 8.3)
            c.setFillColor(palette["muted"])
            c.drawString(m, yinv(234), ntxt("Os campos de preco ficam em branco para preenchimento do fornecedor."))
            return draw_table_header(table_first_top)

        c.saveState()
        c.setFillColor(palette["primary_soft"])
        c.roundRect(m, yinv(70), content_w, 50, 12, stroke=0, fill=1)
        c.restoreState()
        c.setFillColor(palette["primary_dark"])
        c.setFont(fonts["bold"], 15)
        c.drawString(m + 12, yinv(42), ntxt(f"Pedido de Cotacao {doc_num}"))
        c.setFont(fonts["regular"], 8.6)
        c.drawString(m + 12, yinv(56), ntxt(f"Fornecedor: {fornecedor or '-'}"))
        info_chip(w - m - 182, 24, 84, "Pagina", f"{page_no}/{total_pages}")
        info_chip(w - m - 92, 24, 92, "Entrega", fmt_display_date(entrega))
        return draw_table_header(table_next_top)

    def draw_footer(page_no, total_pages):
        response_h = 74
        c.saveState()
        c.setFillColor(colors.white)
        c.setStrokeColor(palette["line"])
        c.roundRect(m, yinv(662 + response_h), table_w, response_h, 8, stroke=1, fill=1)
        c.restoreState()
        c.setFont(fonts["bold"], 8.7)
        c.setFillColor(palette["primary_dark"])
        c.drawString(m + 8, yinv(675), ntxt("Dados de resposta do fornecedor"))
        c.setFont(fonts["regular"], 8.2)
        c.setFillColor(palette["ink"])
        c.drawString(m + 8, yinv(692), ntxt("Prazo de entrega:"))
        c.drawString(m + 8, yinv(708), ntxt("Validade da proposta:"))
        c.drawString(m + 8, yinv(724), ntxt("Observacoes/condicoes:"))
        c.setStrokeColor(palette["line"])
        c.line(m + 98, yinv(694), m + 235, yinv(694))
        c.line(m + 116, yinv(710), m + 235, yinv(710))
        c.line(m + 132, yinv(726), m + table_w - 10, yinv(726))
        c.line(m + 8, yinv(738), m + table_w - 10, yinv(738))

        c.saveState()
        c.setFillColor(palette["surface_alt"])
        c.roundRect(m, yinv(828), table_w, 54, 10, stroke=0, fill=1)
        c.restoreState()
        c.setFont(fonts["bold"], 8.2)
        c.setFillColor(palette["primary_dark"])
        c.drawString(m + 10, yinv(792), ntxt("Resposta / envio"))
        c.setFont(fonts["regular"], 7.5)
        c.setFillColor(palette["ink"])
        footer_left = rodape_lines[0] if rodape_lines else ""
        footer_mid = rodape_lines[1] if len(rodape_lines) > 1 else ""
        footer_right = rodape_lines[2] if len(rodape_lines) > 2 else ""
        c.drawString(m + 10, yinv(805), ntxt(footer_left))
        c.drawString(m + 10, yinv(816), ntxt(footer_mid))
        c.drawString(m + 10, yinv(827), ntxt(footer_right))
        c.drawRightString(w - m - 10, yinv(805), ntxt(f"Total de linhas: {len(linhas)}"))
        c.drawRightString(w - m - 10, yinv(816), ntxt(f"Documento: {doc_num}"))
        c.drawRightString(w - m - 10, yinv(827), ntxt(f"Pagina {page_no}/{total_pages}"))

    def paginate_items(items):
        if not items:
            return [[]]
        pages = [items[:first_capacity]]
        rem = items[first_capacity:]
        while rem:
            pages.append(rem[:next_capacity])
            rem = rem[next_capacity:]
        return pages

    pages = paginate_items(linhas)
    total_pages = len(pages)
    for page_no, page_rows in enumerate(pages, start=1):
        first_page = page_no == 1
        table_y = draw_header(page_no, total_pages, first_page)
        if not page_rows:
            c.setFont(fonts["regular"], 9)
            c.setFillColor(palette["muted"])
            c.drawString(m + 8, yinv(table_y + 16), ntxt("Sem linhas para cotacao."))
        else:
            y_row = table_y
            for idx_line, line in enumerate(page_rows):
                draw_row(y_row, idx_line, line)
                y_row += row_h
        if page_no == total_pages:
            draw_footer(page_no, total_pages)
        if page_no < total_pages:
            c.showPage()

    c.save()

def preview_ne_cotacao(self):
    _ensure_configured()
    ne = self.get_ne_selected()
    if not ne:
        return
    prev_dir = os.path.join(BASE_DIR, "previews")
    try:
        os.makedirs(prev_dir, exist_ok=True)
    except Exception:
        prev_dir = tempfile.gettempdir()
    tmp = os.path.join(prev_dir, f"lugest_ne_cotacao_{ne.get('numero','')}.pdf")
    try:
        self.render_ne_cotacao_pdf(tmp, ne)
        opened = self._open_pdf_default(tmp)
        if not opened:
            messagebox.showwarning("Aviso", f"PDF de cotacao gerado mas nao abriu automaticamente.\nCaminho:\n{tmp}")
    except Exception as ex:
        messagebox.showerror("Erro", f"Falha ao gerar PDF de cotacao: {ex}")

def save_ne_cotacao_pdf(self):
    _ensure_configured()
    ne = self.get_ne_selected()
    if not ne:
        return
    path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF", "*.pdf")],
        initialfile=f"{ne.get('numero','NE')}_cotacao.pdf",
    )
    if not path:
        return
    try:
        self.render_ne_cotacao_pdf(path, ne)
        messagebox.showinfo("OK", "PDF de cotacao guardado")
    except Exception as ex:
        messagebox.showerror("Erro", f"Falha ao guardar PDF de cotacao: {ex}")

def open_ne_cotacao_pdf_with(self):
    _ensure_configured()
    ne = self.get_ne_selected()
    if not ne:
        return
    tmp = os.path.join(tempfile.gettempdir(), f"lugest_ne_cotacao_{ne.get('numero','')}.pdf")
    try:
        self.render_ne_cotacao_pdf(tmp, ne)
        subprocess.Popen(["rundll32.exe", "shell32.dll,OpenAs_RunDLL", tmp])
    except Exception as ex:
        messagebox.showerror("Erro", f"Falha ao abrir cotacao com: {ex}")

def add_exp_linha(self):
    _ensure_configured()
    enc = self.get_exp_selected_encomenda()
    if not enc:
        messagebox.showerror("Erro", "Selecione uma encomenda para expedir.")
        return
    sel = self.exp_tbl_pecas.selection()
    if not sel:
        messagebox.showerror("Erro", "Selecione uma peca disponivel.")
        return
    p = self.exp_peca_row_map.get(sel[0])
    if not p:
        messagebox.showerror("Erro", "Peca nao encontrada.")
        return
    try:
        qtd = parse_float(self.exp_qtd_var.get(), 0.0)
    except Exception:
        qtd = 0.0
    if qtd <= 0:
        messagebox.showerror("Erro", "Quantidade invalida.")
        return
    key = str(p.get("id", "") or f"{p.get('ref_interna','')}|{p.get('ref_externa','')}")
    used = sum(parse_float(x.get("qtd", 0), 0.0) for x in self.exp_draft_linhas if str(x.get("_key", "")) == key)
    disp = peca_qtd_disponivel_expedicao(p) - used
    if qtd > disp + 1e-9:
        messagebox.showerror("Erro", f"Quantidade superior ao disponivel ({fmt_num(disp)}).")
        return
    merged = False
    for l in self.exp_draft_linhas:
        if str(l.get("_key", "")) == key:
            l["qtd"] = parse_float(l.get("qtd", 0), 0.0) + qtd
            merged = True
            break
    if not merged:
        pid = p.get("id", "") or key
        self.exp_draft_linhas.append(
            {
                "_key": key,
                "encomenda": enc.get("numero", ""),
                "peca_id": pid,
                "ref_interna": p.get("ref_interna", ""),
                "ref_externa": p.get("ref_externa", ""),
                "descricao": p.get("Observacoes", p.get("Observações", "")) or p.get("ref_externa", "") or p.get("ref_interna", ""),
                "qtd": qtd,
                "unid": "UN",
                "peso": 0.0,
                "manual": False,
            }
        )
    self.refresh_expedicao_pecas()
    self.refresh_expedicao_draft()

def _sync_ne_linhas_with_materia(self, ne):
    _ensure_configured()
    changed = False
    mat_map = {m.get("id"): m for m in self.data.get("materiais", [])}
    for l in ne.get("linhas", []):
        if not origem_is_materia(l.get("origem", "")):
            continue
        if l.get("_stock_in") or l.get("entregue"):
            continue
        m = mat_map.get(l.get("ref"))
        if not m:
            continue
        new_preco = round(materia_preco_unitario(m), 6)
        cur_preco = parse_float(l.get("preco", 0), 0)
        if abs(new_preco - cur_preco) > 1e-9:
            l["preco"] = new_preco
            q = parse_float(l.get("qtd", 0), 0)
            desconto = max(0.0, min(100.0, parse_float(l.get("desconto", 0), 0)))
            iva = max(0.0, min(100.0, parse_float(l.get("iva", 23), 23)))
            l["total"] = round(((q * new_preco) * (1.0 - (desconto / 100.0))) * (1.0 + (iva / 100.0)), 6)
            changed = True
        formato = m.get("formato", detect_materia_formato(m))
        comp = parse_float(m.get("comprimento", 0), 0)
        larg = parse_float(m.get("largura", 0), 0)
        metros = parse_float(m.get("metros", 0), 0)
        if formato == "Tubo":
            dim_txt = f"{fmt_num(metros)}m"
        else:
            dim_txt = f"{fmt_num(comp)}x{fmt_num(larg)}" if (comp > 0 and larg > 0) else "-"
        new_desc = f"{m.get('material','')} {fmt_num(m.get('espessura',''))}mm | {dim_txt}"
        if (l.get("descricao", "") or "") != new_desc:
            l["descricao"] = new_desc
            changed = True
        if l.get("unid", "") != "UN":
            l["unid"] = "UN"
            changed = True
        # manter dados tecnicos sincronizados para recebimento
        new_meta = {
            "material": m.get("material", ""),
            "espessura": m.get("espessura", ""),
            "comprimento": parse_float(m.get("comprimento", 0), 0),
            "largura": parse_float(m.get("largura", 0), 0),
            "metros": parse_float(m.get("metros", 0), 0),
            "localizacao": m.get("Localizacao", ""),
            "lote_fornecedor": m.get("lote_fornecedor", ""),
            "peso_unid": parse_float(m.get("peso_unid", 0), 0),
            "p_compra": parse_float(m.get("p_compra", 0), 0),
            "formato": formato,
        }
        for k, v in new_meta.items():
            if l.get(k) != v:
                l[k] = v
                changed = True
    if changed:
        ne["total"] = sum(parse_float(x.get("total", 0), 0) for x in ne.get("linhas", []))
    return changed

def _dialog_ne_origem_linha(self):
    _ensure_configured()
    use_custom = CUSTOM_TK_AVAILABLE and os.environ.get("USE_CUSTOM_NE", "1") != "0"
    if not use_custom:
        ans = messagebox.askyesnocancel(
            "Origem da Linha",
            "Adicionar linha de Materia-Prima?\n\nSim = Materia-Prima\nNao = Produto",
        )
        if ans is None:
            return None
        return "Materia-Prima" if ans else "Produto"

    dlg = ctk.CTkToplevel(self.root)
    dlg.title("Origem da Linha")
    dlg.geometry("380x190")
    dlg.resizable(False, False)
    dlg.grab_set()
    chosen = {"v": None}
    box = ctk.CTkFrame(dlg, fg_color="#f7f8fb", corner_radius=12)
    box.pack(fill="both", expand=True, padx=10, pady=10)
    ctk.CTkLabel(
        box,
        text="Adicionar linha de:",
        font=("Segoe UI", 16, "bold"),
        text_color="#7a0f1a",
    ).pack(anchor="w", padx=12, pady=(10, 8))
    row = ctk.CTkFrame(box, fg_color="#f7f8fb")
    row.pack(fill="x", padx=10, pady=8)
    ctk.CTkButton(
        row,
        text="Materia-Prima",
        width=150,
        height=36,
        fg_color="#b42318",
        command=lambda: (chosen.update(v="Materia-Prima"), dlg.destroy()),
    ).pack(side="left", padx=4)
    ctk.CTkButton(
        row,
        text="Produto",
        width=120,
        height=36,
        fg_color="#c43a42",
        command=lambda: (chosen.update(v="Produto"), dlg.destroy()),
    ).pack(side="left", padx=4)
    ctk.CTkButton(
        box,
        text="Cancelar",
        width=100,
        command=dlg.destroy,
    ).pack(anchor="e", padx=12, pady=(8, 10))
    self.root.wait_window(dlg)
    return chosen["v"]

def _dialog_escolher_materia_prima_ne(self):
    _ensure_configured()
    mats = self.data.get("materiais", [])
    if not mats:
        messagebox.showinfo("Info", "Sem registos em Matéria-Prima.")
        return None
    use_custom = CUSTOM_TK_AVAILABLE and os.environ.get("USE_CUSTOM_NE", "1") != "0"
    DlgCls = ctk.CTkToplevel if use_custom else Toplevel
    FrameCls = ctk.CTkFrame if use_custom else ttk.Frame
    LabelCls = ctk.CTkLabel if use_custom else ttk.Label
    EntryCls = ctk.CTkEntry if use_custom else ttk.Entry
    ButtonCls = ctk.CTkButton if use_custom else ttk.Button

    dlg = DlgCls(self.root)
    dlg.title("Escolher Matéria-Prima")
    try:
        dlg.geometry("1100x620")
        dlg.transient(self.root)
    except Exception:
        pass
    dlg.grab_set()
    filtro = StringVar()

    top = FrameCls(dlg, fg_color="#f7f8fb") if use_custom else FrameCls(dlg)
    top.pack(fill="x", padx=8, pady=8)
    LabelCls(top, text="Pesquisar").pack(side="left", padx=6)
    ent = EntryCls(top, textvariable=filtro, width=320 if use_custom else 45)
    ent.pack(side="left", padx=6)

    cols = ("id", "material", "espessura", "dimensao", "metros", "disp", "lote", "local")
    frame = FrameCls(dlg, fg_color="#ffffff") if use_custom else FrameCls(dlg)
    frame.pack(fill="both", expand=True, padx=8, pady=6)
    tree = ttk.Treeview(frame, columns=cols, show="headings", height=14, style="NEL.Treeview" if use_custom else "")
    heads = {
        "id": "ID",
        "material": "Material",
        "espessura": "Esp.",
        "dimensao": "Dimensão",
        "metros": "Metros",
        "disp": "Disponivel",
        "lote": "Lote",
        "local": "Localização",
    }
    for c in cols:
        tree.heading(c, text=heads[c])
    tree.column("id", width=95, anchor="w")
    tree.column("material", width=170, anchor="w")
    tree.column("espessura", width=75, anchor="center")
    tree.column("dimensao", width=150, anchor="center")
    tree.column("metros", width=80, anchor="e")
    tree.column("disp", width=90, anchor="e")
    tree.column("lote", width=130, anchor="w")
    tree.column("local", width=180, anchor="w")
    tree.pack(fill="both", expand=True, side="left")
    vs = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    hs = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
    vs.pack(side="right", fill="y")
    hs.pack(side="bottom", fill="x")
    tree.tag_configure("even", background="#f8fbff")
    tree.tag_configure("odd", background="#eef4fb")

    mat_map = {m.get("id"): m for m in mats}

    def refresh():
        q = (filtro.get() or "").strip().lower()
        for i in tree.get_children():
            tree.delete(i)
        for idx, m in enumerate(mats):
            disp = parse_float(m.get("quantidade", 0), 0) - parse_float(m.get("reservado", 0), 0)
            dim = f"{fmt_num(m.get('comprimento', 0))}x{fmt_num(m.get('largura', 0))}"
            vals = (
                m.get("id", ""),
                m.get("material", ""),
                fmt_num(m.get("espessura", "")),
                dim,
                fmt_num(m.get("metros", 0)),
                fmt_num(disp),
                m.get("lote_fornecedor", ""),
                m.get("Localizacao", ""),
            )
            if q and not any(q in str(v).lower() for v in vals):
                continue
            tree.insert("", END, values=vals, tags=("odd" if idx % 2 else "even",))

    picked = {}

    def on_ok():
        sel = tree.selection()
        if not sel:
            return
        mid = tree.item(sel[0], "values")[0]
        m = mat_map.get(mid)
        if not m:
            return
        formato = m.get("formato", detect_materia_formato(m))
        comp = parse_float(m.get("comprimento", 0), 0)
        larg = parse_float(m.get("largura", 0), 0)
        metros = parse_float(m.get("metros", 0), 0)
        preco_unid = materia_preco_unitario(m)
        un = "UN"
        if formato == "Tubo":
            dim_txt = f"{fmt_num(metros)}m"
        else:
            dim_txt = f"{fmt_num(comp)}x{fmt_num(larg)}" if (comp > 0 and larg > 0) else "-"
        desc = f"{m.get('material','')} {fmt_num(m.get('espessura',''))}mm | {dim_txt}"
        picked.update(
            {
                "id": m.get("id", ""),
                "descricao": desc,
                "preco_unid": preco_unid,
                "unid": un,
                "material": m.get("material", ""),
                "espessura": fmt_num(m.get("espessura", "")),
                "formato": formato,
            }
        )
        dlg.destroy()

    bottom = FrameCls(dlg, fg_color="#f7f8fb") if use_custom else FrameCls(dlg)
    bottom.pack(fill="x", padx=8, pady=8)
    if use_custom:
        ButtonCls(bottom, text="Selecionar", command=on_ok, width=140, fg_color="#f59e0b", hover_color="#d97706").pack(side="right", padx=4)
        ButtonCls(bottom, text="Cancelar", command=dlg.destroy, width=120).pack(side="right", padx=4)
    else:
        ButtonCls(bottom, text="Selecionar", command=on_ok).pack(side="right", padx=4)
        ButtonCls(bottom, text="Cancelar", command=dlg.destroy).pack(side="right", padx=4)
    ent.bind("<KeyRelease>", lambda _e: refresh())
    ent.bind("<Return>", lambda _e: refresh())
    tree.bind("<Double-1>", lambda _e: on_ok())
    tree.bind("<Return>", lambda _e: on_ok())
    refresh()
    dlg.wait_window()
    return picked if picked else None

def guardar_ne(self):
    _ensure_configured()
    lst = self.data.setdefault("notas_encomenda", [])
    existing = next((x for x in lst if x.get("numero") == self.ne_num.get()), None)
    old_lines = existing.get("linhas", []) if existing else []
    ne = {
        "numero": self.ne_num.get(),
        "fornecedor": self.ne_fornecedor.get().strip(),
        "fornecedor_id": self.ne_fornecedor_id.get().strip(),
        "contacto": self.ne_contacto.get().strip(),
        "data_entrega": self.ne_entrega.get().strip(),
        "obs": self.ne_obs.get().strip(),
        "local_descarga": self.ne_local_descarga.get().strip(),
        "meio_transporte": self.ne_meio_transporte.get().strip(),
        "linhas": self.ne_collect_lines(),
        "estado": (existing.get("estado", "Em edicao") if existing else "Em edicao"),
        "oculta": (existing.get("oculta", False) if existing else False),
        "_draft": False,
    }
    for i, l in enumerate(ne["linhas"]):
        if i < len(old_lines):
            old = old_lines[i]
            qtd_tot = parse_float(l.get("qtd", 0), 0)
            qtd_old = parse_float(
                old.get(
                    "qtd_entregue",
                    old.get("qtd", 0) if old.get("entregue") else 0,
                ),
                0,
            )
            qtd_old = max(0.0, min(qtd_tot, qtd_old))
            l["qtd_entregue"] = qtd_old
            if old.get("entregue") or (qtd_tot > 0 and qtd_old >= (qtd_tot - 1e-9)):
                l["entregue"] = True
            if old.get("_stock_in") and l.get("entregue"):
                l["_stock_in"] = True
            for k in (
                "guia_entrega",
                "fatura_entrega",
                "data_doc_entrega",
                "data_entrega_real",
                "obs_entrega",
                "entregas_linha",
            ):
                if old.get(k):
                    l[k] = old.get(k)
        else:
            qtd_tot = parse_float(l.get("qtd", 0), 0)
            if l.get("entregue"):
                l["qtd_entregue"] = qtd_tot
    # calc total
    ne["total"] = sum(l.get("total", 0) for l in ne["linhas"])
    # ponte: se preco de linha foi alterado manualmente, atualiza preco base do produto
    prod_map = {p.get("codigo"): p for p in self.data.get("produtos", [])}
    mat_changed = False
    for l in ne["linhas"]:
        if origem_is_materia(l.get("origem", "")):
            if self._update_materia_preco_from_unit(l.get("ref", ""), l.get("preco", 0)):
                mat_changed = True
            continue
        p = prod_map.get(l.get("ref"))
        if not p:
            continue
        preco_linha = parse_float(l.get("preco", 0), 0)
        cat = p.get("categoria", "")
        tipo = p.get("tipo", "")
        modo = produto_modo_preco(cat, tipo)
        if modo == "peso":
            peso = parse_float(p.get("peso_unid", 0), 0)
            if peso > 0:
                p["p_compra"] = round(preco_linha / peso, 6)
        elif modo == "metros":
            metros_un = parse_float(p.get("metros_unidade", p.get("metros", 0)), 0)
            if metros_un > 0:
                p["p_compra"] = round(preco_linha / metros_un, 6)
        else:
            p["p_compra"] = preco_linha
    if mat_changed:
        self._sync_ne_linhas_with_materia(ne)
    existing = next((x for x in lst if x.get("numero") == ne["numero"]), None)
    if existing:
        existing.update(ne)
    else:
        lst.append(ne)
    save_data(self.data)
    self.refresh_ne()
    try:
        self.on_ne_select()
    except Exception:
        pass
    messagebox.showinfo("OK", "Nota de Encomenda guardada")

def _update_materia_preco_from_unit(self, materia_id, preco_unit):
    _ensure_configured()
    m = next((x for x in self.data.get("materiais", []) if x.get("id") == materia_id), None)
    if not m:
        return False
    preco_linha = parse_float(preco_unit, 0)
    old = parse_float(m.get("p_compra", 0), 0)
    new_val = old
    formato = m.get("formato", detect_materia_formato(m))
    if formato == "Tubo":
        metros = parse_float(m.get("metros", 0), 0)
        if metros > 0:
            new_val = round(preco_linha / metros, 6)
    elif formato in ("Chapa", "Perfil"):
        peso = parse_float(m.get("peso_unid", 0), 0)
        if peso > 0:
            new_val = round(preco_linha / peso, 6)
    else:
        new_val = preco_linha
    if abs(new_val - old) > 1e-9:
        m["p_compra"] = new_val
        m["atualizado_em"] = now_iso()
        return True
    return False

def _open_pdf_default(self, path):
    _ensure_configured()
    path = os.path.abspath(path)
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    if os.path.getsize(path) == 0:
        raise RuntimeError("PDF gerado vazio")
    try:
        os.startfile(path)
        return True
    except Exception:
        pass
    try:
        os.startfile(path, "open")
        return True
    except Exception:
        pass
    try:
        subprocess.Popen(["rundll32.exe", "shell32.dll,OpenAs_RunDLL", path])
        return True
    except Exception:
        pass
    try:
        return bool(webbrowser.open_new(f"file:///{path.replace(os.sep, '/')}"))
    except Exception:
        return False





