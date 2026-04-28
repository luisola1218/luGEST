from lugest_infra.legacy.module_context import configure_module, ensure_module

_CONFIGURED = False

def configure(main_globals):
    configure_module(globals(), main_globals)

def _ensure_configured():
    ensure_module(globals(), "app_misc_actions")

def _repair_widget_mojibake(widget):
    _ensure_configured()
    repair_fn = globals().get("_repair_mojibake_text")
    if not callable(repair_fn):
        return

    stack = [widget]
    while stack:
        w = stack.pop()
        try:
            txt = w.cget("text")
            if isinstance(txt, str) and txt:
                fixed = repair_fn(txt)
                if fixed != txt:
                    w.configure(text=fixed)
        except Exception:
            pass

        try:
            vals = w.cget("values")
            if vals:
                if isinstance(vals, str):
                    try:
                        vals_seq = tuple(w.tk.splitlist(vals))
                    except Exception:
                        vals_seq = (vals,)
                else:
                    vals_seq = tuple(vals)
                fixed_vals = tuple(repair_fn(v) if isinstance(v, str) else v for v in vals_seq)
                if vals_seq != fixed_vals:
                    w.configure(values=fixed_vals)
        except Exception:
            pass

        try:
            if isinstance(w, ttk.Treeview):
                cols = list(w.cget("columns") or [])
                for col in cols:
                    try:
                        info = w.heading(col)
                        label = info.get("text")
                        if isinstance(label, str) and label:
                            fixed_label = repair_fn(label)
                            if fixed_label != label:
                                w.heading(col, text=fixed_label)
                    except Exception:
                        pass
        except Exception:
            pass

        try:
            children = w.winfo_children()
        except Exception:
            children = []
        for child in children:
            stack.append(child)

def _repair_visible_ui_texts(self):
    _ensure_configured()
    try:
        current = self.nb.select() if hasattr(self, "nb") else None
    except Exception:
        current = None
    if current:
        try:
            _repair_widget_mojibake(self.nametowidget(current))
            return
        except Exception:
            pass
    try:
        _repair_widget_mojibake(self.root)
    except Exception:
        pass

def escolher_operacoes_fluxo(self, current_text="", parent=None):
    _ensure_configured()
    use_custom = CUSTOM_TK_AVAILABLE and (
        getattr(self, "encomendas_use_custom", False)
        or getattr(self, "orc_use_custom", False)
        or getattr(self, "op_use_custom", False)
    )
    Win = ctk.CTkToplevel if use_custom else Toplevel
    Frm = ctk.CTkFrame if use_custom else ttk.Frame
    Lbl = ctk.CTkLabel if use_custom else ttk.Label
    Btn = ctk.CTkButton if use_custom else ttk.Button
    Chk = ctk.CTkCheckBox if use_custom else ttk.Checkbutton

    parent_win = parent if parent is not None else self.root
    win = Win(parent_win)
    win.title("Operações da OFF")
    if use_custom:
        try:
            win.geometry("430x440")
            win.configure(fg_color="#f7f8fb")
        except Exception:
            pass
    else:
        win.geometry("380x430")
    try:
        win.transient(parent_win)
        win.grab_set()
    except Exception:
        pass

    Lbl(win, text="Seleciona as operações da peça", font=("Segoe UI", 12, "bold") if use_custom else None).pack(anchor="w", padx=10, pady=(10, 4))
    Lbl(win, text=f"'{OFF_OPERACAO_OBRIGATORIA}' é obrigatória.", text_color="#5b6f84" if use_custom else None).pack(anchor="w", padx=10, pady=(0, 6))

    box = ctk.CTkScrollableFrame(win, fg_color="#ffffff") if use_custom else Frm(win)
    box.pack(fill="both", expand=True, padx=10, pady=6)

    selected = set(parse_operacoes_lista(current_text))
    vars_map = {}
    for nome in OFF_OPERACOES_DISPONIVEIS:
        v = BooleanVar(value=(nome in selected or nome == OFF_OPERACAO_OBRIGATORIA))
        vars_map[nome] = v
        kwargs = {}
        if nome == OFF_OPERACAO_OBRIGATORIA:
            kwargs["state"] = "disabled"
        Chk(box, text=nome, variable=v, **kwargs).pack(anchor="w", padx=8, pady=5)

    out = {"value": None}

    def on_ok():
        ops = [nome for nome in OFF_OPERACOES_DISPONIVEIS if bool(vars_map[nome].get())]
        if OFF_OPERACAO_OBRIGATORIA not in ops:
            ops.append(OFF_OPERACAO_OBRIGATORIA)
        out["value"] = " + ".join(ops)
        win.destroy()

    def on_cancel():
        win.destroy()

    btns = Frm(win, fg_color="transparent") if use_custom else Frm(win)
    btns.pack(fill="x", padx=10, pady=(4, 10))
    Btn(btns, text="Confirmar", command=on_ok, width=120 if use_custom else None).pack(side="left", padx=4)
    Btn(btns, text="Cancelar", command=on_cancel, width=120 if use_custom else None).pack(side="right", padx=4)
    win.wait_window()
    return out["value"]

def escolher_operacoes_concluir(self, peca, parent=None):
    _ensure_configured()
    pendentes = peca_operacoes_pendentes(peca)
    if not pendentes:
        return []
    use_custom = CUSTOM_TK_AVAILABLE and (
        getattr(self, "encomendas_use_custom", False)
        or getattr(self, "orc_use_custom", False)
        or getattr(self, "op_use_custom", False)
    )
    Win = ctk.CTkToplevel if use_custom else Toplevel
    Frm = ctk.CTkFrame if use_custom else ttk.Frame
    Lbl = ctk.CTkLabel if use_custom else ttk.Label
    Btn = ctk.CTkButton if use_custom else ttk.Button
    Chk = ctk.CTkCheckBox if use_custom else ttk.Checkbutton

    parent_win = parent if parent is not None else self.root
    win = Win(parent_win)
    win.title("Concluir Operações")
    if use_custom:
        try:
            win.geometry("430x440")
            win.configure(fg_color="#f7f8fb")
        except Exception:
            pass
    else:
        win.geometry("380x420")
    try:
        win.transient(parent_win)
        win.grab_set()
    except Exception:
        pass

    concluidas = peca_operacoes_concluidas(peca)
    Lbl(
        win,
        text="Seleciona as operações concluídas nesta baixa (inclui Embalamento)",
        font=("Segoe UI", 12, "bold") if use_custom else None,
    ).pack(anchor="w", padx=10, pady=(10, 4))
    if concluidas:
        Lbl(
            win,
            text=f"Já concluídas: {', '.join(concluidas)}",
            text_color="#5b6f84" if use_custom else None,
        ).pack(anchor="w", padx=10, pady=(0, 6))

    box = ctk.CTkScrollableFrame(win, fg_color="#ffffff") if use_custom else Frm(win)
    box.pack(fill="both", expand=True, padx=10, pady=6)
    vars_map = {}
    pre_select_single = len(pendentes) == 1
    for nome in pendentes:
        # Não auto-confirma embalamento para evitar conclusão acidental.
        v = BooleanVar(value=pre_select_single)
        vars_map[nome] = v
        Chk(box, text=nome, variable=v).pack(anchor="w", padx=8, pady=5)

    out = {"value": None}

    def on_ok():
        out["value"] = [nome for nome in pendentes if bool(vars_map[nome].get())]
        win.destroy()

    def on_cancel():
        out["value"] = None
        win.destroy()

    btns = Frm(win, fg_color="transparent") if use_custom else Frm(win)
    btns.pack(fill="x", padx=10, pady=(4, 10))
    Btn(btns, text="Confirmar", command=on_ok, width=120 if use_custom else None).pack(side="left", padx=4)
    Btn(btns, text="Cancelar", command=on_cancel, width=120 if use_custom else None).pack(side="right", padx=4)
    win.wait_window()
    return out.get("value")

def apply_permissions(self):
    _ensure_configured()
    role = str(self.user.get("role", "") or "").strip()
    role_norm = norm_text(role)
    if "orcament" in role_norm:
        role = "Orcamentista"
    if role == "Producao":
        self.nb.tab(self.tab_clientes, state="hidden")
        self.nb.tab(self.tab_qualidade, state="hidden")
        self.nb.tab(self.tab_plano, state="hidden")
        self.nb.tab(self.tab_operador, state="hidden")
        self.nb.tab(self.tab_orc, state="hidden")
        for t in [self.tab_expedicao, self.tab_produtos, self.tab_ne]:
            self.nb.tab(t, state="hidden")
    elif role == "Qualidade":
        self.nb.tab(self.tab_materia, state="hidden")
        self.nb.tab(self.tab_encomendas, state="hidden")
        self.nb.tab(self.tab_plano, state="hidden")
        self.nb.tab(self.tab_operador, state="hidden")
        self.nb.tab(self.tab_orc, state="hidden")
        for t in [self.tab_expedicao, self.tab_produtos, self.tab_ne]:
            self.nb.tab(t, state="hidden")
    elif role == "Planeamento":
        self.nb.tab(self.tab_clientes, state="hidden")
        self.nb.tab(self.tab_materia, state="hidden")
        self.nb.tab(self.tab_encomendas, state="hidden")
        self.nb.tab(self.tab_qualidade, state="hidden")
        self.nb.tab(self.tab_operador, state="hidden")
        self.nb.tab(self.tab_orc, state="hidden")
        for t in [self.tab_expedicao, self.tab_produtos, self.tab_ne]:
            self.nb.tab(t, state="hidden")
    elif role == "Operador":
        # esconder tudo exceto Operador e Plano
        for tab in [
            self.tab_menu,
            self.tab_clientes,
            self.tab_materia,
            self.tab_encomendas,
            self.tab_qualidade,
            self.tab_export,
            self.tab_expedicao,
            self.tab_orc,
            self.tab_ordens,
            self.tab_produtos,
            self.tab_ne,
        ]:
            self.nb.tab(tab, state="hidden")
        self.nb.tab(self.tab_operador, state="normal")
        self.nb.tab(self.tab_plano, state="normal")
        self.show_operador_menu()
        self.nb.select(self.tab_operador)
    elif role == "Orçamentista":
        for tab in [
            self.tab_clientes,
            self.tab_materia,
            self.tab_encomendas,
            self.tab_plano,
            self.tab_qualidade,
            self.tab_export,
            self.tab_expedicao,
            self.tab_operador,
            self.tab_ordens,
            self.tab_produtos,
            self.tab_ne,
        ]:
            self.nb.tab(tab, state="hidden")
    self.apply_menu_only_mode()

def sort_treeview(self, tree, col, reverse):
    _ensure_configured()
    data = [(tree.set(k, col), k) for k in tree.get_children("")]
    def sort_key(item):
        value = item[0]
        try:
            return float(value)
        except ValueError:
            return value.lower()
    data.sort(key=sort_key, reverse=reverse)
    for index, (_, k) in enumerate(data):
        tree.move(k, "", index)
    tree.heading(col, command=lambda: self.sort_treeview(tree, col, not reverse))

def refresh(self, full=False, persist=False):
    _ensure_configured()
    # Coalesce de refresh para evitar cascatas (save -> refresh -> refresh...)
    if getattr(self, "_refresh_busy", False):
        self._refresh_pending = True
        self._refresh_pending_full = bool(getattr(self, "_refresh_pending_full", False) or full)
        self._refresh_pending_persist = bool(getattr(self, "_refresh_pending_persist", False) or persist)
        return

    self._refresh_busy = True
    try:
        # por defeito atualiza so o menu visivel (mais rapido em bases grandes)
        if persist:
            save_data(self.data)

        def do_full_refresh():
            self.refresh_clientes()
            self.refresh_materia()
            self.refresh_encomendas()
            self.refresh_plano()
            self.refresh_qualidade()
            if hasattr(self, "refresh_ordens"):
                self.refresh_ordens()
            if hasattr(self, "refresh_operador"):
                self.refresh_operador()
            if hasattr(self, "refresh_expedicao"):
                self.refresh_expedicao()
            if hasattr(self, "refresh_produtos"):
                self.refresh_produtos()
            if hasattr(self, "refresh_orc_list"):
                self.refresh_orc_list()
            if hasattr(self, "refresh_ne"):
                self.refresh_ne()

        try:
            cur = self.nb.select()
        except Exception:
            cur = None

        if full:
            do_full_refresh()
        else:
            if not cur:
                do_full_refresh()
            elif hasattr(self, "tab_menu") and cur == str(self.tab_menu):
                # No menu principal nao faz refresh completo para evitar bloqueio no arranque.
                if hasattr(self, "refresh_menu_dashboard_stats"):
                    self.refresh_menu_dashboard_stats()
            elif cur == str(self.tab_clientes):
                self.refresh_clientes()
            elif cur == str(self.tab_materia):
                self.refresh_materia()
            elif cur == str(self.tab_encomendas):
                self.refresh_encomendas()
            elif cur == str(self.tab_plano):
                self.refresh_plano()
            elif cur == str(self.tab_qualidade):
                self.refresh_qualidade()
            elif hasattr(self, "tab_orc") and cur == str(self.tab_orc):
                if hasattr(self, "refresh_orc_list"):
                    self.refresh_orc_list()
            elif hasattr(self, "tab_ordens") and cur == str(self.tab_ordens):
                if hasattr(self, "refresh_ordens"):
                    self.refresh_ordens()
            elif hasattr(self, "tab_export") and cur == str(self.tab_export):
                if hasattr(self, "refresh_export_data"):
                    self.refresh_export_data()
            elif hasattr(self, "tab_expedicao") and cur == str(self.tab_expedicao):
                if hasattr(self, "refresh_expedicao"):
                    self.refresh_expedicao()
            elif hasattr(self, "tab_produtos") and cur == str(self.tab_produtos):
                if hasattr(self, "refresh_produtos"):
                    self.refresh_produtos()
            elif hasattr(self, "tab_ne") and cur == str(self.tab_ne):
                if hasattr(self, "refresh_ne"):
                    self.refresh_ne()
            elif hasattr(self, "tab_operador") and cur == str(self.tab_operador):
                if hasattr(self, "refresh_operador"):
                    self.refresh_operador()
            else:
                do_full_refresh()

        # Fora do menu principal, evitar recalculo extra de dashboard em cada refresh.
        if hasattr(self, "refresh_menu_dashboard_stats") and hasattr(self, "tab_menu"):
            try:
                if self.nb.select() == str(self.tab_menu):
                    self.refresh_menu_dashboard_stats()
            except Exception:
                pass
        try:
            _repair_visible_ui_texts(self)
        except Exception:
            pass
    finally:
        self._refresh_busy = False

    if getattr(self, "_refresh_pending", False):
        pfull = bool(getattr(self, "_refresh_pending_full", False))
        ppersist = bool(getattr(self, "_refresh_pending_persist", False))
        self._refresh_pending = False
        self._refresh_pending_full = False
        self._refresh_pending_persist = False
        try:
            self.root.after(1, lambda: self.refresh(full=pfull, persist=ppersist))
        except Exception:
            self.refresh(full=pfull, persist=ppersist)

def refresh_materiais(self, enc):
    _ensure_configured()
    children = self.tbl_materiais.get_children()
    if children:
        self.tbl_materiais.delete(*children)
    if not enc:
        return
    selected = self.selected_material
    for idx, m in enumerate(enc.get("materiais", [])):
        tag = "mat_even" if idx % 2 == 0 else "mat_odd"
        item_id = self.tbl_materiais.insert("", END, values=(m.get("material", ""), m.get("estado", "Preparacao")), tags=(tag,))
        if selected and m.get("material") == selected:
            self.tbl_materiais.selection_set(item_id)
    self.tbl_materiais.tag_configure("mat_even", background="#fff0f2")
    self.tbl_materiais.tag_configure("mat_odd", background="#fff5f6")

def libertar_reserva(self):
    _ensure_configured()
    enc = self.get_selected_encomenda()
    if not enc:
        return
    if not enc["reservas"]:
        messagebox.showinfo("Info", "Sem reservas")
        return
    for r in enc["reservas"]:
        log_stock(self.data, "LIBERTAR", f"{r.get('material_id','')} qtd={r.get('quantidade',0)} encomenda={enc['numero']}")
    aplicar_reserva_em_stock(self.data, enc["reservas"], -1)
    enc["reservas"] = []
    save_data(self.data)
    self.refresh()

def confirmar_baixa_stock(self, mat, esp, total):
    _ensure_configured()
    def _norm_mat(v):
        return str(v or "").strip().lower()

    def _norm_esp(v):
        txt = str(v or "").strip().lower().replace("mm", "").replace(",", ".")
        txt = "".join(ch for ch in txt if ch.isdigit() or ch in ".-")
        if not txt:
            return ""
        try:
            num = float(txt)
            if num.is_integer():
                return str(int(num))
            return f"{num:.6f}".rstrip("0").rstrip(".")
        except Exception:
            return txt

    def _loc_txt(m):
        return m.get("Localização", m.get("Localizacao", m.get("Localiza?o", "")))

    candidatos = []
    for m in self.data.get("materiais", []):
        if _norm_mat(m.get("material")) != _norm_mat(mat):
            continue
        if _norm_esp(m.get("espessura")) != _norm_esp(esp):
            continue
        disp = parse_float(m.get("quantidade"), 0.0) - parse_float(m.get("reservado"), 0.0)
        if disp <= 0:
            continue
        candidatos.append((m, disp))
    if not candidatos:
        messagebox.showinfo("Info", f"Sem stock disponivel para {mat} esp. {esp}.")
        return None

    use_custom = CUSTOM_TK_AVAILABLE and getattr(self, "op_use_custom", False)
    Win = ctk.CTkToplevel if use_custom else Toplevel
    Frm = ctk.CTkFrame if use_custom else ttk.Frame
    Lbl = ctk.CTkLabel if use_custom else ttk.Label
    Ent = ctk.CTkEntry if use_custom else ttk.Entry
    Btn = ctk.CTkButton if use_custom else ttk.Button

    if use_custom:
        win = ctk.CTkToplevel(self.root)
        win.title("Dar baixa de stock")
        win.transient(self.root)
        win.lift()
        win.grab_set()
        try:
            win.geometry("900x500")
            win.configure(fg_color="#f7f8fb")
        except Exception:
            pass
        top = ctk.CTkFrame(win, fg_color="#f7f8fb")
        top.pack(fill="x", padx=10, pady=(10, 6))
        ctk.CTkLabel(
            top,
            text=f"Material: {mat} | Espessura: {esp} | Produzido: {fmt_num(total)}",
            font=("Segoe UI", 13, "bold"),
        ).pack(anchor="w")
        ctk.CTkLabel(
            top,
            text="Defina a baixa numa unica chapa (padrao custom)",
            text_color="#6b7280",
            font=("Segoe UI", 11),
        ).pack(anchor="w", pady=(2, 0))

        list_frame = ctk.CTkScrollableFrame(win, width=860, height=320, fg_color="#ffffff")
        list_frame.pack(fill="both", expand=True, padx=10, pady=6)
        headers = ["Dimensao", "Disponivel", "Local", "Lote", "Dar baixa"]
        for i, h in enumerate(headers):
            ctk.CTkLabel(list_frame, text=h, font=("Segoe UI", 12, "bold")).grid(row=0, column=i, sticky="w", padx=8, pady=(4, 6))
        rows = []
        r = 1
        for m, disp in candidatos:
            dimensao = f"{m.get('comprimento','')}x{m.get('largura','')}"
            ctk.CTkLabel(list_frame, text=dimensao).grid(row=r, column=0, sticky="w", padx=8, pady=2)
            ctk.CTkLabel(list_frame, text=fmt_num(disp)).grid(row=r, column=1, sticky="w", padx=8, pady=2)
            ctk.CTkLabel(list_frame, text=_loc_txt(m)).grid(row=r, column=2, sticky="w", padx=8, pady=2)
            ctk.CTkLabel(list_frame, text=m.get("lote_fornecedor", "")).grid(row=r, column=3, sticky="w", padx=8, pady=2)
            v = StringVar()
            ctk.CTkEntry(list_frame, textvariable=v, width=110).grid(row=r, column=4, sticky="w", padx=8, pady=2)
            rows.append((m, disp, v))
            r += 1
        ok_var = {"ok": False, "qtd": None, "material_id": None, "lote": None}

        def on_ok_custom():
            picks = []
            for m, disp, v in rows:
                txt = (v.get() or "").strip()
                if not txt:
                    continue
                try:
                    qtd = float(txt)
                except Exception:
                    messagebox.showerror("Erro", "Quantidade invalida")
                    return
                if qtd <= 0:
                    messagebox.showerror("Erro", "Quantidade invalida")
                    return
                if qtd > disp:
                    messagebox.showerror("Erro", f"Quantidade maior que disponivel para {m.get('id','')}")
                    return
                picks.append((m, qtd))
            if not picks:
                messagebox.showerror("Erro", "Defina a quantidade de baixa")
                return
            if len(picks) > 1:
                messagebox.showerror("Erro", "Defina baixa em apenas uma chapa por vez")
                return
            m_sel, qtd_sel = picks[0]
            ok_var["ok"] = True
            ok_var["qtd"] = qtd_sel
            ok_var["material_id"] = m_sel.get("id")
            ok_var["lote"] = m_sel.get("lote_fornecedor", "")
            win.destroy()

        btns = ctk.CTkFrame(win, fg_color="#f7f8fb")
        btns.pack(fill="x", padx=10, pady=(2, 10))
        ctk.CTkButton(btns, text="Dar baixa", command=on_ok_custom, width=140).pack(side="left", padx=6)
        ctk.CTkButton(btns, text="Cancelar", command=win.destroy, width=140).pack(side="right", padx=6)
        self.root.wait_window(win)
        return (ok_var["qtd"], ok_var["material_id"], ok_var["lote"]) if ok_var["ok"] else None

    win = Win(self.root)
    win.title("Dar baixa de stock")
    win.transient(self.root)
    win.lift()
    win.grab_set()
    if use_custom:
        try:
            win.geometry("820x420")
        except Exception:
            pass
    top = Frm(win, fg_color="#f7f8fb") if use_custom else Frm(win)
    top.pack(fill="x", padx=10, pady=(10, 6))
    Lbl(top, text=f"Material: {mat} | Espessura: {esp} | Produzido: {fmt_num(total)}", font=("Segoe UI", 13, "bold") if use_custom else None).pack(anchor="w")

    tbl_wrap = Frm(win, fg_color="#ffffff") if use_custom else Frm(win)
    tbl_wrap.pack(fill="both", expand=True, padx=10, pady=6)
    tbl = ttk.Treeview(tbl_wrap, columns=("lote", "material", "dim", "disp", "local"), show="headings", height=8)
    tbl.heading("lote", text="Lote")
    tbl.heading("material", text="Material")
    tbl.heading("dim", text="Dimensao")
    tbl.heading("disp", text="Disponivel")
    tbl.heading("local", text="Localizacao")
    tbl.column("lote", width=120)
    tbl.column("material", width=120)
    tbl.column("dim", width=140)
    tbl.column("disp", width=90)
    tbl.column("local", width=140)
    tbl.pack(side="left", fill="both", expand=True)
    vsb = ttk.Scrollbar(tbl_wrap, orient="vertical", command=tbl.yview)
    tbl.configure(yscrollcommand=vsb.set)
    vsb.pack(side="right", fill="y")
    row_map = {}
    for m, disp in candidatos:
        dim = f"{m.get('comprimento','')}x{m.get('largura','')}"
        iid = m.get("id") or f"row{len(row_map)+1}"
        row_map[iid] = (m, disp)
        tbl.insert("", END, iid=iid, values=(m.get("lote_fornecedor", ""), m.get("material", ""), dim, fmt_num(disp), _loc_txt(m)))
    qtd_var = StringVar(value="1")
    form = Frm(win, fg_color="#f7f8fb") if use_custom else Frm(win)
    form.pack(fill="x", padx=10, pady=(0, 6))
    Lbl(form, text="Chapas a dar baixa").pack(side="left")
    Ent(form, textvariable=qtd_var, width=100 if use_custom else 10).pack(side="left", padx=6)
    ok_var = {"ok": False, "qtd": None, "material_id": None, "lote": None}
    def on_ok():
        try:
            v = float(qtd_var.get())
        except Exception:
            messagebox.showerror("Erro", "Quantidade invalida")
            return
        if v <= 0:
            messagebox.showerror("Erro", "Quantidade invalida")
            return
        sel = tbl.selection()
        if not sel:
            messagebox.showerror("Erro", "Selecione uma chapa")
            return
        sel_id = sel[0]
        mat_sel, disp_sel = row_map.get(sel_id, ({}, 0.0))
        if v > disp_sel:
            messagebox.showerror("Erro", "Quantidade maior que disponivel")
            return
        ok_var["ok"] = True
        ok_var["qtd"] = v
        ok_var["material_id"] = mat_sel.get("id")
        ok_var["lote"] = mat_sel.get("lote_fornecedor", "")
        win.destroy()
    def on_cancel():
        win.destroy()
    btns = Frm(win, fg_color="#f7f8fb") if use_custom else Frm(win)
    btns.pack(fill="x", padx=10, pady=(2, 10))
    Btn(btns, text="Dar baixa", command=on_ok, width=120 if use_custom else None).pack(side="left", padx=6)
    Btn(btns, text="Cancelar", command=on_cancel, width=120 if use_custom else None).pack(side="right", padx=6)
    self.root.wait_window(win)
    return (ok_var["qtd"], ok_var["material_id"], ok_var["lote"]) if ok_var["ok"] else None

def get_plano_grid_metrics(self):
    _ensure_configured()
    try:
        start_min = time_to_minutes(self.p_inicio.get())
        end_min = time_to_minutes(self.p_fim.get())
    except Exception:
        start_min = 480
        end_min = 1080
    slot = 30
    rows = (end_min - start_min) // slot
    cols = 6
    w = max(self.plano.winfo_width(), 800)
    h = max(self.plano.winfo_height(), 500)
    col_w = w // (cols + 1)
    row_h = max(20, h // (rows + 1))
    return start_min, end_min, slot, rows, cols, col_w, row_h

def on_plano_drag_start(self, _):
    _ensure_configured()
    sel = self.tbl_plano_enc.selection()
    if not sel:
        self.plano_drag_item = None
        return
    vals = self.tbl_plano_enc.item(sel[0], "values")
    self.plano_drag_item = {
        "encomenda": vals[0],
        "material": vals[1],
        "espessura": vals[2],
        "tempo": vals[3],
    }
    self.plano_drag_preview = None

def on_plano_drag_motion(self, event):
    _ensure_configured()
    if not hasattr(self, "plano_drag_item") or not self.plano_drag_item:
        return
    start_min, end_min, slot, rows, cols, col_w, row_h = self.get_plano_grid_metrics()
    col = (event.x // col_w) - 1
    row = (event.y // row_h) - 1
    if col < 0 or col >= cols or row < 0 or row >= rows:
        return
    x0 = (col + 1) * col_w
    y0 = (row + 1) * row_h
    if self.plano_drag_preview:
        self.plano.delete(self.plano_drag_preview)
    self.plano_drag_preview = self.plano.create_rectangle(x0 + 2, y0 + 2, x0 + col_w - 2, y0 + row_h - 2, outline="#ba2d3d", dash=(3, 3))

def get_chapa_reservada(self, numero, material=None, espessura=None):
    _ensure_configured()
    def _norm_mat(v):
        return str(v or "").strip().lower()
    def _norm_esp(v):
        txt = str(v or "").strip().lower().replace("mm", "").replace(",", ".")
        txt = "".join(ch for ch in txt if ch.isdigit() or ch in ".-")
        if not txt:
            return ""
        try:
            num = float(txt)
            if num.is_integer():
                return str(int(num))
            return f"{num:.6f}".rstrip("0").rstrip(".")
        except Exception:
            return txt
    enc = self.get_encomenda_by_numero(numero)
    if not enc or not enc.get("reservas"):
        return "-"
    for r in enc["reservas"]:
        if material and _norm_mat(r.get("material")) != _norm_mat(material):
            continue
        if espessura and _norm_esp(r.get("espessura")) != _norm_esp(espessura):
            continue
        return f"{r.get('material','')} {r.get('espessura','')} ({r.get('quantidade',0)})"
    return "-"

def _op_select_esp_combo(self):
    _ensure_configured()
    val = (self.op_esp_var.get() if hasattr(self, "op_esp_var") else "").strip()
    if not val:
        return
    if "|" not in val:
        return
    mat, esp_part = val.split("|", 1)
    mat = mat.strip()
    esp = esp_part.replace("mm", "").strip()
    if mat and esp:
        self._op_select_esp_custom(mat, esp)

def _op_select_esp_custom(self, mat, esp):
    _ensure_configured()
    self.op_material = mat
    self.op_espessura = esp
    try:
        self.op_sel_pecas_ids = set()
    except Exception:
        pass
    try:
        if hasattr(self, "op_esp_var"):
            self.op_esp_var.set(self._op_esp_label(mat, esp))
    except Exception:
        pass
    try:
        self.on_operador_select_espessura(None)
    except Exception:
        pass

def op_blink_schedule(self):
    _ensure_configured()
    try:
        if hasattr(self, "nb") and hasattr(self, "tab_operador"):
            if str(self.nb.select()) != str(self.tab_operador):
                self.root.after(1400, self.op_blink_schedule)
                return
        if not hasattr(self, "op_tbl_enc") or not self.op_tbl_enc.winfo_exists():
            return
        self.op_blink_on = not getattr(self, "op_blink_on", False)
        bg = "#ffbf80" if self.op_blink_on else "#ffd8a8"
        self.op_tbl_enc.tag_configure("op_em_curso", background=bg, foreground="#7c5a00")
        if hasattr(self, "op_tbl_esp"):
            self.op_tbl_esp.tag_configure("op_em_curso", background=bg, foreground="#7c5a00")
        if hasattr(self, "op_tbl_pecas"):
            self.op_tbl_pecas.tag_configure("op_em_curso", background=bg, foreground="#7c5a00")
    except Exception:
        return
    self.root.after(900, self.op_blink_schedule)

def preview_opp_selected(self):
    _ensure_configured()
    opp = self._ordens_get_selected_opp()
    if not opp:
        messagebox.showerror("Erro", "Selecione uma Ordem de Fabrico")
        return
    enc = None
    p = None
    for e in self.data.get("encomendas", []):
        for px in encomenda_pecas(e):
            if px.get("opp") == opp:
                enc = e
                p = px
                break
        if p:
            break
    if not p:
        messagebox.showerror("Erro", "Ordem de Fabrico nao encontrada")
        return
    self._preview_piece_pdf(enc, p)

def on_exp_hist_click_open_pdf(self, event=None):
    _ensure_configured()
    if not hasattr(self, "exp_tbl_hist") or event is None:
        return
    try:
        region = self.exp_tbl_hist.identify("region", event.x, event.y)
        col = self.exp_tbl_hist.identify_column(event.x)
        row = self.exp_tbl_hist.identify_row(event.y)
    except Exception:
        return
    if region not in ("cell", "tree") or not row:
        return
    # Abrir por clique apenas quando carregar no numero da guia (coluna 1).
    if col != "#1":
        return
    self.exp_tbl_hist.selection_set(row)
    self.exp_tbl_hist.focus(row)
    self.preview_expedicao_pdf()

def _exp_parse_datetime(self, value, default_iso=""):
    _ensure_configured()
    txt = str(value or "").strip()
    if not txt:
        return default_iso or now_iso()
    for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
        try:
            dt = datetime.strptime(txt, fmt)
            return dt.strftime("%Y-%m-%dT%H:%M:%S")
        except Exception:
            pass
    try:
        return datetime.fromisoformat(txt.replace("Z", "+00:00")).strftime("%Y-%m-%dT%H:%M:%S")
    except Exception:
        return default_iso or now_iso()

def edit_empresa_info(self):
    _ensure_configured()
    use_custom = CUSTOM_TK_AVAILABLE
    WinCls = ctk.CTkToplevel if use_custom else Toplevel
    FrameCls = ctk.CTkFrame if use_custom else ttk.Frame
    LabelCls = ctk.CTkLabel if use_custom else ttk.Label
    BtnCls = ctk.CTkButton if use_custom else Button

    win = WinCls(self.root)
    win.title("Informacao da Empresa")
    win.geometry("940x760")
    try:
        win.minsize(860, 700)
    except Exception:
        pass
    try:
        sw = int(win.winfo_screenwidth())
        sh = int(win.winfo_screenheight())
        ww = min(980, max(860, sw - 80))
        wh = min(820, max(700, sh - 120))
        x = max(10, (sw - ww) // 2)
        y = max(10, (sh - wh) // 2 - 18)
        win.geometry(f"{ww}x{wh}+{x}+{y}")
    except Exception:
        pass
    try:
        win.transient(self.root)
        win.grab_set()
    except Exception:
        pass

    top = FrameCls(win, fg_color="#f7f8fb") if use_custom else FrameCls(win)
    top.pack(fill="x", padx=10, pady=(10, 6))
    if use_custom:
        LabelCls(
            top,
            text="Rodape dos PDFs",
            font=("Segoe UI", 13, "bold"),
            text_color="#7f1b2c",
        ).pack(anchor="w", padx=4, pady=(2, 0))
    else:
        LabelCls(
            top,
            text="Rodape dos PDFs",
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w", padx=4, pady=(2, 0))
    LabelCls(
        top,
        text="Uma linha por linha. Esta informacao e usada em Orcamento, Nota de Encomenda e outros PDFs.",
    ).pack(anchor="w", padx=4, pady=(2, 4))

    body = FrameCls(win, fg_color="#ffffff") if use_custom else FrameCls(win)
    body.pack(fill="x", expand=False, padx=10, pady=(0, 8))

    txt = Text(body, wrap="word", font=("Segoe UI", 11), height=10)
    txt.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
    ysb = ttk.Scrollbar(body, orient="vertical", command=txt.yview)
    ysb.pack(side="right", fill="y", padx=(0, 6), pady=6)
    txt.configure(yscrollcommand=ysb.set)
    txt.insert("1.0", "\n".join(get_empresa_rodape_lines()))

    emit_cfg = get_guia_emitente_info()
    emit_box = FrameCls(win, fg_color="#f7f8fb") if use_custom else FrameCls(win)
    emit_box.pack(fill="x", padx=10, pady=(0, 8))
    LabelCls(
        emit_box,
        text="Emitente da Guia de Transporte",
        font=("Segoe UI", 11, "bold") if use_custom else None,
        text_color="#7f1b2c" if use_custom else None,
    ).grid(row=0, column=0, columnspan=4, sticky="w", padx=4, pady=(4, 2))
    emit_nome = StringVar(value=emit_cfg.get("nome", ""))
    emit_nif = StringVar(value=emit_cfg.get("nif", ""))
    emit_morada = StringVar(value=emit_cfg.get("morada", ""))
    emit_carga = StringVar(value=emit_cfg.get("local_carga", ""))
    EntryCls = ctk.CTkEntry if use_custom else ttk.Entry
    LabelCls(emit_box, text="Nome").grid(row=1, column=0, sticky="w", padx=4, pady=3)
    EntryCls(emit_box, textvariable=emit_nome, width=260 if use_custom else 30).grid(row=1, column=1, sticky="w", padx=4, pady=3)
    LabelCls(emit_box, text="NIF").grid(row=1, column=2, sticky="w", padx=4, pady=3)
    EntryCls(emit_box, textvariable=emit_nif, width=140 if use_custom else 16).grid(row=1, column=3, sticky="w", padx=4, pady=3)
    LabelCls(emit_box, text="Morada").grid(row=2, column=0, sticky="w", padx=4, pady=3)
    EntryCls(emit_box, textvariable=emit_morada, width=420 if use_custom else 50).grid(row=2, column=1, columnspan=3, sticky="w", padx=4, pady=3)
    LabelCls(emit_box, text="Local carga padrao").grid(row=3, column=0, sticky="w", padx=4, pady=3)
    EntryCls(emit_box, textvariable=emit_carga, width=420 if use_custom else 50).grid(row=3, column=1, columnspan=3, sticky="w", padx=4, pady=3)

    extra_box = FrameCls(win, fg_color="#f7f8fb") if use_custom else FrameCls(win)
    extra_box.pack(fill="both", expand=True, padx=10, pady=(0, 8))
    LabelCls(
        extra_box,
        text="Bloco inferior da Guia de Transporte",
        font=("Segoe UI", 11, "bold") if use_custom else None,
        text_color="#7f1b2c" if use_custom else None,
    ).pack(anchor="w", padx=4, pady=(4, 2))
    LabelCls(
        extra_box,
        text="Uma linha por linha (ex.: IBAN, publicidade, contatos comerciais).",
    ).pack(anchor="w", padx=4, pady=(0, 4))
    extra_wrap = FrameCls(extra_box, fg_color="#ffffff") if use_custom else FrameCls(extra_box)
    extra_wrap.pack(fill="both", expand=True, padx=4, pady=(0, 4))
    txt_extra = Text(extra_wrap, wrap="word", font=("Segoe UI", 10), height=4)
    txt_extra.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
    ysb2 = ttk.Scrollbar(extra_wrap, orient="vertical", command=txt_extra.yview)
    ysb2.pack(side="right", fill="y", padx=(0, 6), pady=6)
    txt_extra.configure(yscrollcommand=ysb2.set)
    txt_extra.insert("1.0", "\n".join(get_guia_extra_info_lines()))

    def restore_defaults():
        txt.delete("1.0", END)
        txt.insert("1.0", "\n".join(ORC_EMPRESA_INFO_RODAPE))
        emit_nome.set(ORC_EMPRESA_INFO_RODAPE[0] if ORC_EMPRESA_INFO_RODAPE else "")
        emit_nif.set("")
        emit_morada.set(ORC_EMPRESA_INFO_RODAPE[1] if len(ORC_EMPRESA_INFO_RODAPE) > 1 else "")
        emit_carga.set(ORC_EMPRESA_INFO_RODAPE[1] if len(ORC_EMPRESA_INFO_RODAPE) > 1 else "")
        txt_extra.delete("1.0", END)

    def save_info():
        raw = txt.get("1.0", END)
        lines = [ln.strip() for ln in raw.replace("\r", "").split("\n") if ln.strip()]
        if not lines:
            messagebox.showerror("Erro", "Introduza pelo menos uma linha de informacao.")
            return
        raw_extra = txt_extra.get("1.0", END)
        extra_lines = [ln.strip() for ln in raw_extra.replace("\r", "").split("\n") if ln.strip()]

        cfg = {}
        try:
            loaded = get_branding_config()
            if isinstance(loaded, dict):
                cfg = dict(loaded)
        except Exception:
            cfg = {}
        cfg["empresa_info_rodape"] = lines
        cfg["guia_emitente"] = {
            "nome": emit_nome.get().strip(),
            "nif": emit_nif.get().strip(),
            "morada": emit_morada.get().strip(),
            "local_carga": emit_carga.get().strip(),
        }
        cfg["guia_info_extra"] = extra_lines

        ok, ex = _persist_branding_config_mysql(cfg)
        if not ok:
            messagebox.showerror("Erro", f"Nao foi possivel guardar a informacao na base de dados.\n{ex}")
            return

        _invalidate_branding_cache()
        messagebox.showinfo("OK", "Informacao atualizada. Os novos PDFs ja usarao este texto.")
        win.destroy()

    # barra de acoes visivel acima da area de texto (evita ficar fora do ecra em escalas altas)
    quick = FrameCls(win, fg_color="#f7f8fb") if use_custom else FrameCls(win)
    quick.pack(fill="x", padx=10, pady=(0, 6), before=body)
    BtnCls(quick, text="Padrao", command=restore_defaults, width=100 if use_custom else None).pack(side="left", padx=4)
    BtnCls(quick, text="Cancelar", command=win.destroy, width=110 if use_custom else None).pack(side="right", padx=4)
    BtnCls(quick, text="Guardar", command=save_info, width=110 if use_custom else None).pack(side="right", padx=4)

def manage_user_accounts(self):
    _ensure_configured()
    role_now = str((getattr(self, "user", {}) or {}).get("role", "") or "").strip().lower()
    if role_now != "admin":
        messagebox.showerror("Permissao", "Apenas o utilizador Admin pode gerir contas.")
        return

    if hasattr(self, "users_manage_win") and self.users_manage_win and self.users_manage_win.winfo_exists():
        try:
            self.users_manage_win.deiconify()
            self.users_manage_win.lift()
            self.users_manage_win.focus_force()
        except Exception:
            pass
        return

    use_custom = CUSTOM_TK_AVAILABLE and os.environ.get("USE_CUSTOM_EXPORT", "1") != "0"
    WinCls = ctk.CTkToplevel if use_custom else Toplevel
    FrameCls = ctk.CTkFrame if use_custom else ttk.Frame
    LabelCls = ctk.CTkLabel if use_custom else ttk.Label
    EntryCls = ctk.CTkEntry if use_custom else ttk.Entry
    ComboCls = ctk.CTkComboBox if use_custom else ttk.Combobox
    BtnCls = ctk.CTkButton if use_custom else ttk.Button

    win = WinCls(self.root)
    self.users_manage_win = win
    win.title("Utilizadores")
    try:
        win.minsize(900, 620)
        try:
            win.resizable(True, True)
        except Exception:
            pass
        sw = int(win.winfo_screenwidth())
        sh = int(win.winfo_screenheight())
        # Abre grande por defeito para evitar cortes em escalas altas.
        ww = max(900, min(1220, sw - 40))
        wh = max(620, min(820, sh - 90))
        x = max(10, (sw - ww) // 2)
        y = max(10, (sh - wh) // 2 - 20)
        win.geometry(f"{ww}x{wh}+{x}+{y}")
        # Em ecrãs pequenos/escala alta, maximiza para garantir visibilidade total.
        if sh <= 760:
            try:
                win.state("zoomed")
            except Exception:
                pass
        win.transient(self.root)
        win.grab_set()
    except Exception:
        pass

    def _on_close():
        try:
            win.destroy()
        finally:
            self.users_manage_win = None

    try:
        win.protocol("WM_DELETE_WINDOW", _on_close)
    except Exception:
        pass

    users = self.data.setdefault("users", [])

    search_var = StringVar(value="")
    username_var = StringVar(value="")
    password_var = StringVar(value="")
    role_var = StringVar(value="Operador")
    selected_username = StringVar(value="")

    outer = FrameCls(win, fg_color="#f7f8fb") if use_custom else FrameCls(win)
    outer.pack(fill="both", expand=True, padx=10, pady=10)

    if use_custom:
        LabelCls(outer, text="Gestão de Utilizadores", font=("Segoe UI", 16, "bold"), text_color="#7a0f1a").pack(anchor="w", padx=4, pady=(2, 8))
    else:
        LabelCls(outer, text="Gestão de Utilizadores", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=4, pady=(2, 8))

    top = FrameCls(outer, fg_color="#f7f8fb") if use_custom else FrameCls(outer)
    top.pack(fill="x", pady=(0, 8))
    LabelCls(top, text="Pesquisar").pack(side="left", padx=(4, 8))
    EntryCls(top, textvariable=search_var, width=340 if use_custom else None).pack(side="left", fill="x", expand=True, padx=(0, 8))

    mid = FrameCls(outer, fg_color="transparent") if use_custom else FrameCls(outer)
    mid.pack(fill="both", expand=True, pady=(0, 8))
    if use_custom:
        style = ttk.Style()
        style.configure(
            "USR.Treeview",
            font=("Segoe UI", 11),
            rowheight=30,
            background="#f8fbff",
            fieldbackground="#f8fbff",
            borderwidth=0,
        )
        style.configure(
            "USR.Treeview.Heading",
            font=("Segoe UI", 11, "bold"),
            background=THEME_HEADER_BG,
            foreground="white",
            relief="flat",
        )
        style.map("USR.Treeview.Heading", background=[("active", THEME_HEADER_ACTIVE)])
        style.map(
            "USR.Treeview",
            background=[("selected", THEME_SELECT_BG)],
            foreground=[("selected", THEME_SELECT_FG)],
        )
        tbl_style = "USR.Treeview"
    else:
        tbl_style = ""

    tbl = ttk.Treeview(mid, columns=("username", "role"), show="headings", height=14, style=tbl_style)
    tbl.heading("username", text="Utilizador")
    tbl.heading("role", text="Perfil")
    tbl.column("username", width=420, anchor="w")
    tbl.column("role", width=180, anchor="center")
    tbl.pack(side="left", fill="both", expand=True)
    sb = ttk.Scrollbar(mid, orient="vertical", command=tbl.yview)
    sb.pack(side="right", fill="y")
    tbl.configure(yscrollcommand=sb.set)
    if use_custom:
        tbl.tag_configure("odd", background="#f8fbff")
        tbl.tag_configure("even", background="#edf3ff")

    def _fit_user_columns(_event=None):
        try:
            total = int(tbl.winfo_width())
        except Exception:
            total = 0
        if total <= 80:
            return
        # Reserva para scrollbar/bordas
        usable = max(120, total - 24)
        user_w = int(usable * 0.70)
        role_w = max(120, usable - user_w)
        try:
            tbl.column("username", width=user_w, minwidth=180, stretch=True)
            tbl.column("role", width=role_w, minwidth=120, stretch=True, anchor="center")
        except Exception:
            pass

    try:
        tbl.bind("<Configure>", _fit_user_columns)
        win.bind("<Configure>", _fit_user_columns, add="+")
    except Exception:
        pass

    sep1 = FrameCls(outer, fg_color="#d9e3f2", height=1) if use_custom else FrameCls(outer)
    if use_custom:
        sep1.pack(fill="x", pady=(0, 8))

    form = FrameCls(outer, fg_color="#f7f8fb") if use_custom else FrameCls(outer)
    form.pack(fill="x", pady=(0, 8))
    for col in (1, 3):
        try:
            form.grid_columnconfigure(col, weight=1)
        except Exception:
            pass

    LabelCls(form, text="Utilizador").grid(row=0, column=0, sticky="w", padx=6, pady=4)
    EntryCls(form, textvariable=username_var, width=260 if use_custom else None).grid(row=0, column=1, sticky="we", padx=6, pady=4)
    LabelCls(form, text="Password").grid(row=0, column=2, sticky="w", padx=6, pady=4)
    EntryCls(form, textvariable=password_var, show="*", width=260 if use_custom else None).grid(row=0, column=3, sticky="we", padx=6, pady=4)
    LabelCls(form, text="Perfil").grid(row=1, column=0, sticky="w", padx=6, pady=4)
    if use_custom:
        role_cb = ComboCls(form, values=["Operador", "Admin"], variable=role_var, width=180)
    else:
        role_cb = ComboCls(form, textvariable=role_var, values=["Operador", "Admin"], state="readonly", width=18)
    role_cb.grid(row=1, column=1, sticky="w", padx=6, pady=4)
    if use_custom:
        LabelCls(
            form,
            text="Deixe a password em branco para manter a atual quando estiver a editar.",
            text_color="#5b6f84",
        ).grid(row=2, column=0, columnspan=4, sticky="w", padx=6, pady=(0, 2))
    else:
        LabelCls(
            form,
            text="Deixe a password em branco para manter a atual quando estiver a editar.",
        ).grid(row=2, column=0, columnspan=4, sticky="w", padx=6, pady=(0, 2))

    sep2 = FrameCls(outer, fg_color="#d9e3f2", height=1) if use_custom else FrameCls(outer)
    if use_custom:
        sep2.pack(fill="x", pady=(2, 8))

    btns = FrameCls(outer, fg_color="transparent") if use_custom else FrameCls(outer)
    btns.pack(fill="x", pady=(2, 2))

    def _refresh_table():
        flt = str(search_var.get() or "").strip().lower()
        for iid in tbl.get_children():
            tbl.delete(iid)
        row_i = 0
        for u in users:
            un = str(u.get("username", "") or "").strip()
            rl = str(u.get("role", "") or "").strip()
            if not un:
                continue
            hay = f"{un} {rl}".lower()
            if flt and flt not in hay:
                continue
            tag = "even" if (row_i % 2 == 0) else "odd"
            tbl.insert("", "end", values=(un, rl), tags=(tag,))
            row_i += 1

    def _select_username(un):
        selected_username.set(un)
        found = False
        for u in users:
            if str(u.get("username", "") or "").strip().lower() == un.lower():
                username_var.set(str(u.get("username", "") or "").strip())
                password_var.set("")
                r = str(u.get("role", "") or "").strip()
                role_var.set(r if r in ("Operador", "Admin") else "Operador")
                found = True
                break
        if not found:
            username_var.set("")
            password_var.set("")
            role_var.set("Operador")

    def _on_select(_e=None):
        sel = tbl.selection()
        if not sel:
            return
        vals = tbl.item(sel[0], "values")
        if vals:
            _select_username(str(vals[0]))

    def _new_user():
        selected_username.set("")
        username_var.set("")
        password_var.set("")
        role_var.set("Operador")
        try:
            tbl.selection_remove(tbl.selection())
        except Exception:
            pass

    def _save_user():
        un = str(username_var.get() or "").strip()
        pw = str(password_var.get() or "")
        rl = str(role_var.get() or "").strip()
        if not un:
            messagebox.showerror("Erro", "Introduza o utilizador.")
            return
        if rl not in ("Operador", "Admin"):
            messagebox.showerror("Erro", "Perfil invalido.")
            return

        old_un = str(selected_username.get() or "").strip()
        dup = next((u for u in users if str(u.get("username", "") or "").strip().lower() == un.lower()), None)
        if dup and (not old_un or str(dup.get("username", "") or "").strip().lower() != old_un.lower()):
            messagebox.showerror("Erro", "Ja existe um utilizador com esse nome.")
            return

        target = None
        if old_un:
            target = next((u for u in users if str(u.get("username", "") or "").strip().lower() == old_un.lower()), None)
        if target is None and not pw.strip():
            messagebox.showerror("Erro", "Introduza a password para o novo utilizador.")
            return
        normalize_pw = globals().get("normalize_password_for_storage")
        validate_pw = globals().get("validate_local_password")
        stored_password = ""
        if pw.strip():
            try:
                if callable(validate_pw):
                    validate_pw(un, pw)
                stored_password = normalize_pw(un, pw, True) if callable(normalize_pw) else pw.strip()
            except Exception as exc:
                messagebox.showerror("Erro", str(exc))
                return
        if target is None:
            target = {"username": un, "password": stored_password, "role": rl}
            users.append(target)
        elif not stored_password:
            stored_password = str(target.get("password", "") or "").strip()
        target["username"] = un
        target["password"] = stored_password
        target["role"] = rl

        if rl == "Operador":
            ops = self.data.setdefault("operadores", [])
            if un not in ops:
                ops.append(un)
            try:
                if getattr(self, "op_user_cb", None) is not None:
                    self.op_user_cb.configure(values=ops)
            except Exception:
                pass

        selected_username.set(un)
        save_data(self.data)
        _refresh_table()
        messagebox.showinfo("OK", "Utilizador guardado.")

    def _remove_user():
        sel = tbl.selection()
        if not sel:
            messagebox.showerror("Erro", "Selecione um utilizador.")
            return
        vals = tbl.item(sel[0], "values")
        un = str(vals[0]) if vals else ""
        if not un:
            return
        logged_user = str((getattr(self, "user", {}) or {}).get("username", "") or "").strip()
        if logged_user and logged_user.lower() == un.lower():
            messagebox.showerror("Erro", "Nao pode remover o utilizador atualmente autenticado.")
            return
        if not messagebox.askyesno("Confirmar", f"Remover utilizador '{un}'?"):
            return
        self.data["users"] = [u for u in users if str(u.get("username", "") or "").strip().lower() != un.lower()]
        save_data(self.data)
        selected_username.set("")
        _refresh_table()
        _new_user()

    try:
        tbl.bind("<<TreeviewSelect>>", _on_select)
    except Exception:
        pass
    try:
        search_var.trace_add("write", lambda *_: _refresh_table())
    except Exception:
        pass

    if use_custom:
        BtnCls(btns, text="Novo", command=_new_user, width=130, height=38, corner_radius=10).pack(side="left", padx=4)
        BtnCls(btns, text="Guardar", command=_save_user, width=130, height=38, corner_radius=10, fg_color="#f59e0b", hover_color="#d97706", text_color="#ffffff").pack(side="left", padx=4)
        BtnCls(btns, text="Remover", command=_remove_user, width=130, height=38, corner_radius=10).pack(side="left", padx=4)
        BtnCls(btns, text="Fechar", command=_on_close, width=130, height=38, corner_radius=10).pack(side="right", padx=4)
    else:
        BtnCls(btns, text="Novo", command=_new_user, width=14).pack(side="left", padx=6, ipady=4)
        BtnCls(btns, text="Guardar", command=_save_user, width=14, bg="#f59e0b", fg="white", activebackground="#d97706", activeforeground="white").pack(side="left", padx=6, ipady=4)
        BtnCls(btns, text="Remover", command=_remove_user, width=14).pack(side="left", padx=6, ipady=4)
        BtnCls(btns, text="Fechar", command=_on_close, width=14).pack(side="right", padx=6, ipady=4)

    _refresh_table()
    _fit_user_columns()

def _extract_hex_set(value):
    out = set()
    if isinstance(value, (list, tuple)):
        for v in value:
            out.update(_extract_hex_set(v))
        return out
    txt = str(value or "").strip().lower()
    if not txt:
        return out
    for m in re.findall(r"#[0-9a-f]{6}", txt):
        out.add(m)
    return out


def _retheme_widget_tree(widget, primary_aliases, hover_aliases, select_aliases):
    opts = (
        ("fg_color", CTK_PRIMARY_RED, primary_aliases),
        ("hover_color", CTK_PRIMARY_RED_HOVER, hover_aliases),
        ("selected_color", CTK_PRIMARY_RED, primary_aliases),
        ("selected_hover_color", CTK_PRIMARY_RED_HOVER, hover_aliases),
        ("button_color", CTK_PRIMARY_RED, primary_aliases),
        ("button_hover_color", CTK_PRIMARY_RED_HOVER, hover_aliases),
        ("unselected_hover_color", THEME_SELECT_BG, select_aliases),
    )
    for opt_name, target, aliases in opts:
        try:
            cur = widget.cget(opt_name)
        except Exception:
            continue
        try:
            cur_hex = _extract_hex_set(cur)
            if cur_hex and cur_hex.intersection(aliases):
                widget.configure(**{opt_name: target})
        except Exception:
            pass
    for child in widget.winfo_children():
        _retheme_widget_tree(child, primary_aliases, hover_aliases, select_aliases)


def apply_primary_color_runtime(self, old_primary="", old_hover=""):
    _ensure_configured()
    if not CUSTOM_TK_AVAILABLE:
        return
    p_alias = {
        _normalize_hex_color(DEFAULT_PRIMARY_RED, DEFAULT_PRIMARY_RED),
        _normalize_hex_color(old_primary, DEFAULT_PRIMARY_RED),
        _normalize_hex_color(CTK_PRIMARY_RED, DEFAULT_PRIMARY_RED),
        _normalize_hex_color(THEME_HEADER_BG, DEFAULT_THEME_HEADER_BG),
        _normalize_hex_color(DEFAULT_THEME_HEADER_BG, DEFAULT_THEME_HEADER_BG),
        "#a32035",
        "#c62828",
    }
    h_alias = {
        _normalize_hex_color(DEFAULT_PRIMARY_RED_HOVER, DEFAULT_PRIMARY_RED_HOVER),
        _normalize_hex_color(old_hover, DEFAULT_PRIMARY_RED_HOVER),
        _normalize_hex_color(CTK_PRIMARY_RED_HOVER, DEFAULT_PRIMARY_RED_HOVER),
        _normalize_hex_color(THEME_HEADER_ACTIVE, DEFAULT_THEME_HEADER_ACTIVE),
        _normalize_hex_color(DEFAULT_THEME_HEADER_ACTIVE, DEFAULT_THEME_HEADER_ACTIVE),
        "#8f1d14",
    }
    s_alias = {
        _normalize_hex_color(DEFAULT_THEME_SELECT_BG, DEFAULT_THEME_SELECT_BG),
        _normalize_hex_color(THEME_SELECT_BG, DEFAULT_THEME_SELECT_BG),
        _normalize_hex_color(_mix_hex_color(old_primary or DEFAULT_PRIMARY_RED, "#ffffff", 0.82), DEFAULT_THEME_SELECT_BG),
        "#fde2e4",
    }
    p_alias = {x for x in p_alias if x}
    h_alias = {x for x in h_alias if x}
    s_alias = {x for x in s_alias if x}
    try:
        _retheme_widget_tree(self.root, p_alias, h_alias, s_alias)
    except Exception:
        pass
    try:
        style = ttk.Style()
        headings = [
            "Clientes.Treeview.Heading",
            "Stock.Treeview.Heading",
            "Encomendas.Treeview.Heading",
            "EncomendasSub.Treeview.Heading",
            "Plan.Treeview.Heading",
            "Qual.Treeview.Heading",
            "Orc.Treeview.Heading",
            "OrcLin.Treeview.Heading",
            "Operador.Treeview.Heading",
            "EXP.Treeview.Heading",
            "PROD.Treeview.Heading",
            "NE.Treeview.Heading",
            "NEL.Treeview.Heading",
        ]
        for hs in headings:
            style.configure(hs, background=THEME_HEADER_BG)
            style.map(hs, background=[("active", THEME_HEADER_ACTIVE)])
        tree_styles = [
            "Clientes.Treeview",
            "Stock.Treeview",
            "Encomendas.Treeview",
            "EncomendasSub.Treeview",
            "Plan.Treeview",
            "Qual.Treeview",
            "Orc.Treeview",
            "OrcLin.Treeview",
            "Operador.Treeview",
            "EXP.Treeview",
            "PROD.Treeview",
            "NE.Treeview",
            "NEL.Treeview",
            "ORCREF.Treeview",
        ]
        for ts in tree_styles:
            style.map(
                ts,
                background=[("selected", THEME_SELECT_BG)],
                foreground=[("selected", THEME_SELECT_FG)],
            )
    except Exception:
        pass


def _persist_primary_color(color_hex):
    cfg = {}
    try:
        loaded = get_branding_config()
        if isinstance(loaded, dict):
            cfg = dict(loaded)
    except Exception:
        cfg = {}
    cfg["primary_color"] = _normalize_hex_color(color_hex, DEFAULT_PRIMARY_RED)
    ok, ex = _persist_branding_config_mysql(cfg)
    if not ok:
        raise RuntimeError(f"Falha ao guardar cor na base de dados: {ex}")
    _invalidate_branding_cache()


def _persist_branding_config_mysql(cfg):
    conn = None
    try:
        conn = _mysql_connect()
        with conn.cursor() as cur:
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS app_config (
                    ckey VARCHAR(80) PRIMARY KEY,
                    cvalue LONGTEXT NULL,
                    updated_at DATETIME NULL
                )
                """
            )
            payload = json.dumps(cfg, ensure_ascii=False)
            ts = _to_mysql_datetime(now_iso())
            cur.execute(
                """
                INSERT INTO app_config (ckey, cvalue, updated_at)
                VALUES (%s, %s, %s)
                ON DUPLICATE KEY UPDATE
                    cvalue=VALUES(cvalue),
                    updated_at=VALUES(updated_at)
                """,
                ("branding_config", payload, ts),
            )
        conn.commit()
        return True, None
    except Exception as ex:
        try:
            if conn:
                conn.rollback()
        except Exception:
            pass
        return False, ex
    finally:
        try:
            if conn:
                conn.close()
        except Exception:
            pass


def _invalidate_branding_cache():
    """Limpa o cache de branding no módulo atual e no módulo principal."""
    global _BRANDING_CACHE
    _BRANDING_CACHE = None
    try:
        import sys as _sys
        for _mod_name in ("main", "__main__"):
            _mod = _sys.modules.get(_mod_name)
            if _mod is not None and hasattr(_mod, "_BRANDING_CACHE"):
                setattr(_mod, "_BRANDING_CACHE", None)
    except Exception:
        pass


def _apply_primary_color(self, color_hex):
    _ensure_configured()
    chosen = _normalize_hex_color(color_hex, DEFAULT_PRIMARY_RED) or DEFAULT_PRIMARY_RED
    old_primary = _normalize_hex_color(CTK_PRIMARY_RED, DEFAULT_PRIMARY_RED)
    old_hover = _normalize_hex_color(CTK_PRIMARY_RED_HOVER, DEFAULT_PRIMARY_RED_HOVER)
    apply_primary_theme_color(chosen)
    try:
        self.CTK_PRIMARY_RED = CTK_PRIMARY_RED
        self.CTK_PRIMARY_RED_HOVER = CTK_PRIMARY_RED_HOVER
        self.THEME_HEADER_BG = THEME_HEADER_BG
        self.THEME_HEADER_ACTIVE = THEME_HEADER_ACTIVE
        self.THEME_SELECT_BG = THEME_SELECT_BG
        self.THEME_SELECT_FG = THEME_SELECT_FG
    except Exception:
        pass
    _set_ctk_button_defaults_red()
    self.refresh_module_contexts()
    apply_primary_color_runtime(self, old_primary, old_hover)
    self.setup_styles()


def choose_primary_color(self):
    _ensure_configured()
    initial = _normalize_hex_color(CTK_PRIMARY_RED, DEFAULT_PRIMARY_RED) or DEFAULT_PRIMARY_RED
    try:
        _, picked = colorchooser.askcolor(color=initial, title="Cor principal do programa")
    except Exception:
        picked = None
    if not picked:
        return
    color_hex = _normalize_hex_color(picked, initial)
    if not color_hex:
        return
    try:
        _persist_primary_color(color_hex)
    except Exception as ex:
        messagebox.showerror("Erro", f"Nao foi possivel guardar a cor.\n{ex}")
        return
    _apply_primary_color(self, color_hex)
    messagebox.showinfo("OK", "Cor principal atualizada.")


def reset_primary_color(self):
    _ensure_configured()
    try:
        _persist_primary_color(DEFAULT_PRIMARY_RED)
    except Exception as ex:
        messagebox.showerror("Erro", f"Nao foi possivel repor a cor padrao.\n{ex}")
        return
    _apply_primary_color(self, DEFAULT_PRIMARY_RED)
    messagebox.showinfo("OK", "Cor padrao reposta.")


def export_csv(self, tipo):
    _ensure_configured()
    path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
    if not path:
        return
    if tipo == "clientes":
        rows = self.data["clientes"]
    elif tipo == "materiais":
        rows = self.data["materiais"]
    elif tipo == "encomendas":
        rows = self.data["encomendas"]
    elif tipo == "plano":
        rows = self.data.get("plano", [])
    elif tipo == "qualidade":
        rows = self.data["qualidade"]
    else:
        return
    if not rows:
        messagebox.showinfo("Info", "Sem dados para exportar")
        return
    keys = sorted(rows[0].keys())
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=keys)
        w.writeheader()
        w.writerows(rows)
    messagebox.showinfo("OK", "Exportado")


def open_sheet_calculator(self):
    _ensure_configured()
    use_custom = CUSTOM_TK_AVAILABLE and os.environ.get("USE_CUSTOM_EXPORT", "1") != "0"
    WinCls = ctk.CTkToplevel if use_custom else Toplevel
    FrameCls = ctk.CTkFrame if use_custom else ttk.Frame
    LabelCls = ctk.CTkLabel if use_custom else ttk.Label
    EntryCls = ctk.CTkEntry if use_custom else ttk.Entry
    BtnCls = ctk.CTkButton if use_custom else ttk.Button
    ComboCls = ctk.CTkComboBox if use_custom else ttk.Combobox

    materiais = [
        ("Aço carbono", 7.85),
        ("Aço galvanizado", 7.85),
        ("Aço inox 304", 7.93),
        ("Aço inox 316", 7.99),
        ("Alumínio", 2.70),
        ("Cobre", 8.96),
        ("Latão", 8.50),
    ]
    dens_map = {nome: dens for nome, dens in materiais}

    win = WinCls(self.root)
    win.title("Calculadora de Chapa")
    try:
        win.geometry("760x520")
        win.minsize(700, 480)
    except Exception:
        pass
    try:
        win.transient(self.root)
        win.grab_set()
    except Exception:
        pass

    top = FrameCls(win, fg_color="#ffffff") if use_custom else FrameCls(win)
    top.pack(fill="both", expand=True, padx=12, pady=12)

    LabelCls(
        top,
        text="Calculadora de Chapa",
        font=("Segoe UI", 16, "bold") if use_custom else ("Segoe UI", 12, "bold"),
        text_color="#7a0f1a" if use_custom else None,
    ).grid(row=0, column=0, columnspan=4, sticky="w", padx=6, pady=(4, 8))

    mat_var = StringVar(value=materiais[0][0])
    dens_var = StringVar(value=f"{materiais[0][1]:.3f}")
    esp_var = StringVar(value="1.5")
    lar_var = StringVar(value="1000")
    comp_var = StringVar(value="2000")
    qtd_var = StringVar(value="1")

    area_var = StringVar(value="0.000")
    kgm2_var = StringVar(value="0.000")
    peso_unit_var = StringVar(value="0.000")
    peso_total_var = StringVar(value="0.000")
    preco_kg_var = StringVar(value="0.00")
    custo_total_var = StringVar(value="0.00")

    def parse_float_local(v, default=0.0):
        txt = str(v or "").strip().replace(",", ".")
        try:
            return float(txt)
        except Exception:
            return default

    LabelCls(top, text="Material").grid(row=1, column=0, sticky="w", padx=6, pady=4)
    if use_custom:
        mat_cb = ComboCls(top, variable=mat_var, values=[m[0] for m in materiais], width=220)
    else:
        mat_cb = ComboCls(top, textvariable=mat_var, values=[m[0] for m in materiais], width=26, state="readonly")
    mat_cb.grid(row=1, column=1, sticky="w", padx=6, pady=4)

    LabelCls(top, text="Densidade (g/cm³)").grid(row=1, column=2, sticky="w", padx=6, pady=4)
    EntryCls(top, textvariable=dens_var, width=140 if use_custom else 16).grid(row=1, column=3, sticky="w", padx=6, pady=4)

    LabelCls(top, text="Espessura (mm)").grid(row=2, column=0, sticky="w", padx=6, pady=4)
    EntryCls(top, textvariable=esp_var, width=140 if use_custom else 16).grid(row=2, column=1, sticky="w", padx=6, pady=4)
    LabelCls(top, text="Largura (mm)").grid(row=2, column=2, sticky="w", padx=6, pady=4)
    EntryCls(top, textvariable=lar_var, width=140 if use_custom else 16).grid(row=2, column=3, sticky="w", padx=6, pady=4)

    LabelCls(top, text="Comprimento (mm)").grid(row=3, column=0, sticky="w", padx=6, pady=4)
    EntryCls(top, textvariable=comp_var, width=140 if use_custom else 16).grid(row=3, column=1, sticky="w", padx=6, pady=4)
    LabelCls(top, text="Quantidade (chapas)").grid(row=3, column=2, sticky="w", padx=6, pady=4)
    EntryCls(top, textvariable=qtd_var, width=140 if use_custom else 16).grid(row=3, column=3, sticky="w", padx=6, pady=4)

    LabelCls(top, text="Preço por kg (€) [opcional]").grid(row=4, column=0, sticky="w", padx=6, pady=4)
    EntryCls(top, textvariable=preco_kg_var, width=140 if use_custom else 16).grid(row=4, column=1, sticky="w", padx=6, pady=4)

    out = FrameCls(top, fg_color="#f7f8fb") if use_custom else FrameCls(top)
    out.grid(row=5, column=0, columnspan=4, sticky="ew", padx=6, pady=(12, 8))

    LabelCls(out, text="Área por chapa (m²)").grid(row=0, column=0, sticky="w", padx=6, pady=4)
    LabelCls(out, textvariable=area_var, font=("Segoe UI", 12, "bold") if use_custom else None).grid(row=0, column=1, sticky="w", padx=6, pady=4)
    LabelCls(out, text="Peso por m² (kg/m²)").grid(row=0, column=2, sticky="w", padx=6, pady=4)
    LabelCls(out, textvariable=kgm2_var, font=("Segoe UI", 12, "bold") if use_custom else None).grid(row=0, column=3, sticky="w", padx=6, pady=4)

    LabelCls(out, text="Peso por chapa (kg)").grid(row=1, column=0, sticky="w", padx=6, pady=4)
    LabelCls(out, textvariable=peso_unit_var, font=("Segoe UI", 13, "bold") if use_custom else ("Segoe UI", 10, "bold")).grid(row=1, column=1, sticky="w", padx=6, pady=4)
    LabelCls(out, text="Peso total (kg)").grid(row=1, column=2, sticky="w", padx=6, pady=4)
    LabelCls(out, textvariable=peso_total_var, font=("Segoe UI", 13, "bold") if use_custom else ("Segoe UI", 10, "bold")).grid(row=1, column=3, sticky="w", padx=6, pady=4)

    LabelCls(out, text="Custo total (€)").grid(row=2, column=0, sticky="w", padx=6, pady=4)
    LabelCls(out, textvariable=custo_total_var, font=("Segoe UI", 13, "bold") if use_custom else ("Segoe UI", 10, "bold")).grid(row=2, column=1, sticky="w", padx=6, pady=4)

    ajuda = (
        "Fórmula: Peso (kg) = Densidade(g/cm³) × Esp(mm) × Larg(mm) × Comp(mm) / 1.000.000"
    )
    LabelCls(top, text=ajuda, text_color="#4a5f77" if use_custom else None).grid(
        row=6, column=0, columnspan=4, sticky="w", padx=6, pady=(4, 8)
    )

    def recalc(*_):
        dens = parse_float_local(dens_var.get(), 0.0)
        esp = parse_float_local(esp_var.get(), 0.0)
        larg = parse_float_local(lar_var.get(), 0.0)
        comp = parse_float_local(comp_var.get(), 0.0)
        qtd = parse_float_local(qtd_var.get(), 0.0)
        preco_kg = parse_float_local(preco_kg_var.get(), 0.0)
        if dens <= 0 or esp <= 0 or larg <= 0 or comp <= 0 or qtd < 0:
            area_var.set("0.000")
            kgm2_var.set("0.000")
            peso_unit_var.set("0.000")
            peso_total_var.set("0.000")
            custo_total_var.set("0.00")
            return
        area_m2 = (larg * comp) / 1000000.0
        kg_m2 = dens * esp
        peso_unit = (dens * esp * larg * comp) / 1000000.0
        peso_total = peso_unit * qtd
        custo_total = peso_total * preco_kg
        area_var.set(f"{area_m2:.3f}")
        kgm2_var.set(f"{kg_m2:.3f}")
        peso_unit_var.set(f"{peso_unit:.3f}")
        peso_total_var.set(f"{peso_total:.3f}")
        custo_total_var.set(f"{custo_total:.2f}")

    def on_material_change(_=None):
        dens = dens_map.get((mat_var.get() or "").strip())
        if dens is not None:
            dens_var.set(f"{dens:.3f}")
        recalc()

    def limpar():
        mat_var.set(materiais[0][0])
        dens_var.set(f"{materiais[0][1]:.3f}")
        esp_var.set("1.5")
        lar_var.set("1000")
        comp_var.set("2000")
        qtd_var.set("1")
        preco_kg_var.set("0.00")
        recalc()

    try:
        if use_custom:
            mat_cb.configure(command=lambda _v=None: on_material_change())
        else:
            mat_cb.bind("<<ComboboxSelected>>", on_material_change)
    except Exception:
        pass

    for var in (dens_var, esp_var, lar_var, comp_var, qtd_var, preco_kg_var):
        try:
            var.trace_add("write", recalc)
        except Exception:
            pass

    bbar = FrameCls(top, fg_color="transparent") if use_custom else FrameCls(top)
    bbar.grid(row=7, column=0, columnspan=4, sticky="w", padx=6, pady=(2, 8))
    if use_custom:
        BtnCls(bbar, text="Calcular", command=recalc, width=120, fg_color=CTK_PRIMARY_RED, hover_color=CTK_PRIMARY_RED_HOVER).pack(side="left", padx=4)
        BtnCls(bbar, text="Limpar", command=limpar, width=120).pack(side="left", padx=4)
        BtnCls(bbar, text="Fechar", command=win.destroy, width=120).pack(side="left", padx=4)
    else:
        BtnCls(bbar, text="Calcular", command=recalc).pack(side="left", padx=4)
        BtnCls(bbar, text="Limpar", command=limpar).pack(side="left", padx=4)
        BtnCls(bbar, text="Fechar", command=win.destroy).pack(side="left", padx=4)

    recalc()

def product_apply_category_mode(self):
    _ensure_configured()
    cat = self.prod_categoria.get()
    tipo = self.prod_tipo.get()
    modo = produto_modo_preco(cat, tipo)
    is_chapa = "chapa" in norm_text(tipo) or "chapa" in norm_text(cat)
    if is_chapa:
        if not self.prod_dimensoes.get().strip():
            self.prod_dimensoes.set("2000x1000")
        self.prod_metros.set("0")
    elif modo == "metros":
        self.prod_dimensoes.set("")
        if self.prod_metros.get() in ("", "0"):
            self.prod_metros.set("6")
    elif modo == "peso":
        self.prod_dimensoes.set("")
        self.prod_metros.set("0")

def _set_prod_form_from_obj(self, p):
    _ensure_configured()
    if not p:
        return
    self.prod_sel_codigo = p.get("codigo", "")
    self.prod_codigo.set(p.get("codigo", ""))
    self.prod_descricao.set(p.get("descricao", ""))
    self.prod_categoria.set(p.get("categoria", ""))
    self.prod_subcat.set(p.get("subcat", ""))
    self.prod_tipo.set(p.get("tipo", ""))
    self.prod_dimensoes.set(p.get("dimensoes", ""))
    self.prod_metros.set(str(p.get("metros", 0)))
    self.prod_peso_unid.set(str(p.get("peso_unid", 0)))
    self.prod_fab.set(p.get("fabricante", ""))
    self.prod_modelo.set(p.get("modelo", ""))
    self.prod_unid.set(p.get("unid", "UN"))
    self.prod_qty.set(str(p.get("qty", 0)))
    self.prod_alerta.set(str(p.get("alerta", 0)))
    self.prod_pcompra.set(str(p.get("p_compra", 0)))
    self.prod_pvp1.set(str(p.get("pvp1", 0)))
    self.prod_pvp2.set(str(p.get("pvp2", 0)))
    self.prod_obs.set(p.get("obs", ""))

def sync_all_ne_from_materia(self):
    _ensure_configured()
    changed_any = False
    for ne in self.data.get("notas_encomenda", []):
        if self._sync_ne_linhas_with_materia(ne):
            changed_any = True
    if not changed_any:
        return
    save_data(self.data)
    try:
        self.refresh_ne()
        self.on_ne_select()
    except Exception:
        pass


def edit_tempos_operacao_planeada(self):
    _ensure_configured()
    use_custom = CUSTOM_TK_AVAILABLE and (
        getattr(self, "export_use_custom", False)
        or getattr(self, "op_use_custom", False)
        or getattr(self, "op_use_full_custom", False)
    )
    Win = ctk.CTkToplevel if use_custom else Toplevel
    FrameCls = ctk.CTkFrame if use_custom else ttk.Frame
    LabelCls = ctk.CTkLabel if use_custom else ttk.Label
    EntryCls = ctk.CTkEntry if use_custom else ttk.Entry
    BtnCls = ctk.CTkButton if use_custom else ttk.Button

    win = Win(self.root)
    win.title("Tempos Planeados por Operação")
    try:
        win.transient(self.root)
        win.grab_set()
    except Exception:
        pass
    try:
        if use_custom:
            win.geometry("560x420")
            win.resizable(False, False)
            win.configure(fg_color="#f7f8fb")
        else:
            win.geometry("520x390")
    except Exception:
        pass

    if not isinstance(self.data.get("tempos_operacao_planeada_min", None), dict):
        self.data["tempos_operacao_planeada_min"] = {}
    raw_map = dict(self.data.get("tempos_operacao_planeada_min", {}) or {})
    ops = [
        ("laser", "Laser"),
        ("quinagem", "Quinagem"),
        ("soldadura", "Soldadura"),
        ("roscagem", "Roscagem"),
        ("embalamento", "Embalamento"),
    ]

    top = FrameCls(win, fg_color="transparent") if use_custom else FrameCls(win)
    top.pack(fill="both", expand=True, padx=12, pady=10)
    LabelCls(
        top,
        text="Tempo planeado por operação (min por peça)",
        font=("Segoe UI", 14, "bold") if use_custom else ("Segoe UI", 11, "bold"),
        text_color="#7a0f1a" if use_custom else None,
    ).pack(anchor="w", pady=(0, 4))
    LabelCls(
        top,
        text="Nota: atualmente o desvio principal é avaliado no Laser (tempo da espessura). "
             "Os restantes postos ficam preparados para planeamento por operação.",
        wraplength=520,
        text_color="#4b5563" if use_custom else None,
    ).pack(anchor="w", pady=(0, 8))

    grid = FrameCls(top, fg_color="#ffffff" if use_custom else None, corner_radius=10 if use_custom else None, border_width=1 if use_custom else 0, border_color="#dce3ee" if use_custom else None) if use_custom else FrameCls(top)
    grid.pack(fill="x", pady=(0, 10))
    vars_map = {}
    for i, (op_key, op_label) in enumerate(ops):
        row = FrameCls(grid, fg_color="transparent") if use_custom else FrameCls(grid)
        row.pack(fill="x", padx=10, pady=5)
        LabelCls(row, text=op_label, width=180 if use_custom else None, anchor="w").pack(side="left")
        v = StringVar(value=str(raw_map.get(op_key, 0) or 0))
        vars_map[op_key] = v
        EntryCls(row, textvariable=v, width=120 if use_custom else 16).pack(side="left", padx=(8, 6))
        LabelCls(row, text="min/peça", text_color="#64748b" if use_custom else None).pack(side="left")

    def on_save():
        out = {}
        for op_key, _op_label in ops:
            txt = str(vars_map[op_key].get() or "0").strip().replace(",", ".")
            try:
                val = max(0.0, float(txt))
            except Exception:
                messagebox.showerror("Erro", f"Valor inválido em {op_key}.")
                return
            out[op_key] = val
        self.data["tempos_operacao_planeada_min"] = out
        save_data(self.data)
        messagebox.showinfo("OK", "Tempos planeados guardados.")
        win.destroy()

    bbar = FrameCls(top, fg_color="transparent") if use_custom else FrameCls(top)
    bbar.pack(fill="x", pady=(4, 2))
    if use_custom:
        BtnCls(
            bbar,
            text="Guardar",
            command=on_save,
            width=130,
            fg_color="#f59e0b",
            hover_color="#d97706",
            text_color="#ffffff",
        ).pack(side="left", padx=4)
        BtnCls(bbar, text="Fechar", command=win.destroy, width=120).pack(side="right", padx=4)
    else:
        BtnCls(bbar, text="Guardar", command=on_save).pack(side="left", padx=4)
        BtnCls(bbar, text="Fechar", command=win.destroy).pack(side="right", padx=4)

