from lugest_infra.legacy.module_context import configure_module, ensure_module

_CONFIGURED = False


def configure(main_globals):
    configure_module(globals(), main_globals)


def _ensure_configured():
    ensure_module(globals(), "ui_build_blocks")

def build_materia(self):
    _ensure_configured()
    self.materia_use_custom = CUSTOM_TK_AVAILABLE and os.environ.get("USE_CUSTOM_STOCK", "1") != "0"
    parent_m = self.tab_materia
    if self.materia_use_custom:
        parent_m = ctk.CTkFrame(self.tab_materia, fg_color="#ffffff")
        parent_m.pack(fill="both", expand=True)

    # filtros e campos
    top_frame = ctk.CTkFrame(
        parent_m,
        fg_color="white",
        corner_radius=10,
        border_width=1,
        border_color="#e7cfd3",
    ) if self.materia_use_custom else ttk.Frame(parent_m)
    top_frame.pack(fill="x", padx=10, pady=10)

    filter_row = ctk.CTkFrame(
        parent_m,
        fg_color="white",
        corner_radius=10,
        border_width=1,
        border_color="#e7cfd3",
    ) if self.materia_use_custom else ttk.Frame(parent_m)
    filter_row.pack(fill="x", padx=10, pady=(0, 6))
    (ctk.CTkLabel if self.materia_use_custom else ttk.Label)(filter_row, text="Pesquisa").pack(side="left")
    EntryCls = ctk.CTkEntry if self.materia_use_custom else ttk.Entry
    if self.materia_use_custom and CUSTOM_TK_AVAILABLE:
        def BtnCls(parent, **kwargs):
            kwargs.setdefault("height", 34)
            kwargs.setdefault("corner_radius", 10)
            kwargs.setdefault("fg_color", CTK_PRIMARY_RED)
            kwargs.setdefault("hover_color", CTK_PRIMARY_RED_HOVER)
            kwargs.setdefault("text_color", "#ffffff")
            kwargs.setdefault("border_width", 0)
            return ctk.CTkButton(parent, **kwargs)
        orange_btn = {"fg_color": "#f59e0b", "hover_color": "#d97706"}
    else:
        BtnCls = ttk.Button
        orange_btn = {}
    self.m_filter = StringVar()
    self.m_filter_entry = EntryCls(filter_row, textvariable=self.m_filter, width=260 if self.materia_use_custom else 40)
    self.m_filter_entry.pack(side="left", padx=6, pady=2)
    self.m_filter_entry.bind("<KeyRelease>", lambda _: self.debounce_call("materia_filter", self.refresh_materia, delay_ms=170))

    frm = top_frame
    self.m_formato = StringVar(value="Chapa")
    self.m_material = StringVar()
    self.m_espessura = StringVar()
    self.m_comprimento = DoubleVar()
    self.m_largura = DoubleVar()
    self.m_metros = DoubleVar()
    self.m_quantidade = DoubleVar()
    self.m_local = StringVar()
    self.m_lote = StringVar()
    self.m_reservado = DoubleVar()
    self.m_peso_unid = DoubleVar()
    self.m_p_compra = DoubleVar()

    labels = [
        ("formato", "Formato", self.m_formato, "formato"),
        ("material", "Material", self.m_material, "material"),
        ("espessura", "Espessura", self.m_espessura, "espessura"),
        ("comprimento", "Comprimento", self.m_comprimento, "num"),
        ("largura", "Largura", self.m_largura, "num"),
        ("metros", "Metros", self.m_metros, "num"),
        ("quantidade", "Quantidade", self.m_quantidade, "num"),
        ("local", "Localização", self.m_local, "local"),
        ("lote", "Lote fornecedor", self.m_lote, "texto"),
        ("reservado", "Reservado", self.m_reservado, "num"),
        ("peso_unid", "Peso/Un. (kg)", self.m_peso_unid, "num"),
        ("p_compra", "Preço Compra (€/kg|€/m)", self.m_p_compra, "num"),
    ]
    self.m_widgets = {}
    def _mw(v):
        try:
            n = int(v)
        except Exception:
            n = 220
        if self.materia_use_custom:
            return n if n >= 120 else n * 8
        return n

    for i, (key, lbl, var, tipo) in enumerate(labels):
        (ctk.CTkLabel if self.materia_use_custom else ttk.Label)(frm, text=lbl).grid(row=i // 2, column=(i % 2) * 2, sticky="w", padx=4, pady=2)
        if tipo == "formato":
            if self.materia_use_custom:
                entry = ctk.CTkComboBox(frm, variable=var, values=MATERIA_FORMATOS, width=_mw(27), state="readonly")
            else:
                entry = ttk.Combobox(frm, textvariable=var, values=MATERIA_FORMATOS, width=27, state="readonly")
        elif tipo == "material":
            materiais = list(dict.fromkeys(MATERIAIS_PRESET + self.data.get("materiais_hist", []) + list_unique(self.data, "material")))
            if self.materia_use_custom:
                entry = ctk.CTkComboBox(frm, variable=var, values=materiais or [""], width=_mw(27), state="normal")
            else:
                entry = ttk.Combobox(frm, textvariable=var, values=materiais, width=27)
        elif tipo == "espessura":
            esp_opts = [str(v).rstrip('0').rstrip('.') if isinstance(v, float) else str(v) for v in ESPESSURAS_PRESET]
            esp_hist = [str(v) for v in self.data.get("espessuras_hist", [])] + list_unique(self.data, "espessura")
            esp = list(dict.fromkeys(esp_opts + esp_hist))
            if self.materia_use_custom:
                entry = ctk.CTkComboBox(frm, variable=var, values=esp or [""], width=_mw(27), state="normal")
            else:
                entry = ttk.Combobox(frm, textvariable=var, values=esp, width=27)
        elif tipo == "local":
            if self.materia_use_custom:
                entry = ctk.CTkComboBox(frm, variable=var, values=LOCALIZACOES_PRESET or [""], width=_mw(27), state="normal")
            else:
                entry = ttk.Combobox(frm, textvariable=var, values=LOCALIZACOES_PRESET, width=27)
        else:
            entry = EntryCls(frm, textvariable=var, width=_mw(220) if self.materia_use_custom else 30)
        entry.grid(row=i // 2, column=(i % 2) * 2 + 1, sticky="w", padx=4, pady=2)
        self.m_widgets[key] = entry

    def _set_widget_state(w, state):
        if not w:
            return
        try:
            w.configure(state=state)
        except Exception:
            pass

    def apply_m_formato(*_):
        fmt = (self.m_formato.get() or "Chapa").strip()
        if fmt == "Chapa":
            _set_widget_state(self.m_widgets.get("comprimento"), "normal")
            _set_widget_state(self.m_widgets.get("largura"), "normal")
            _set_widget_state(self.m_widgets.get("espessura"), "normal")
            _set_widget_state(self.m_widgets.get("peso_unid"), "normal")
            _set_widget_state(self.m_widgets.get("metros"), "normal")
        elif fmt == "Tubo":
            _set_widget_state(self.m_widgets.get("comprimento"), "normal")
            _set_widget_state(self.m_widgets.get("largura"), "normal")
            _set_widget_state(self.m_widgets.get("espessura"), "normal")
            _set_widget_state(self.m_widgets.get("metros"), "normal")
            _set_widget_state(self.m_widgets.get("peso_unid"), "normal")
        else:  # Perfil
            _set_widget_state(self.m_widgets.get("comprimento"), "normal")
            _set_widget_state(self.m_widgets.get("largura"), "normal")
            _set_widget_state(self.m_widgets.get("espessura"), "normal")
            _set_widget_state(self.m_widgets.get("peso_unid"), "normal")
            _set_widget_state(self.m_widgets.get("metros"), "normal")

    self.m_formato.trace_add("write", apply_m_formato)
    apply_m_formato()

    btns = ctk.CTkFrame(
        parent_m,
        fg_color="white",
        corner_radius=10,
        border_width=1,
        border_color="#e7cfd3",
    ) if self.materia_use_custom else ttk.Frame(parent_m)
    btns.pack(fill="x", padx=10)
    BtnCls(btns, text="Adicionar", command=self.add_material, **orange_btn).pack(side="left", padx=4, pady=2)
    BtnCls(btns, text="Editar", command=self.edit_material).pack(side="left", padx=4, pady=2)
    BtnCls(btns, text="Corrigir Stock", command=self.corrigir_stock).pack(side="left", padx=4, pady=2)
    BtnCls(btns, text="Remover", command=self.remove_material).pack(side="left", padx=4, pady=2)
    BtnCls(btns, text="Dar baixa", command=self.baixa_material).pack(side="left", padx=4, pady=2)
    BtnCls(btns, text="Atualizar", command=self.refresh_materia, **orange_btn).pack(side="left", padx=4, pady=2)
    BtnCls(btns, text="Histórico", command=self.show_stock_log).pack(side="left", padx=4, pady=2)
    BtnCls(btns, text="Exportar CSV", command=lambda: self.export_csv("materiais")).pack(side="left", padx=4, pady=2)
    BtnCls(btns, text="Pre-visualizar A4", command=self.preview_stock_a4).pack(side="left", padx=4, pady=2)

    tbl_style = ""
    if self.materia_use_custom:
        style = ttk.Style()
        style.configure(
            "Stock.Treeview",
            font=("Segoe UI", 10),
            rowheight=26,
            background="#f8fbff",
            fieldbackground="#f8fbff",
        )
        style.configure(
            "Stock.Treeview.Heading",
            font=("Segoe UI", 10, "bold"),
            background=THEME_HEADER_BG,
            foreground="white",
            relief="flat",
        )
        style.map("Stock.Treeview.Heading", background=[("active", THEME_HEADER_ACTIVE)])
        style.map(
            "Stock.Treeview",
            background=[("selected", THEME_SELECT_BG)],
            foreground=[("selected", THEME_SELECT_FG)],
        )
        tbl_style = "Stock.Treeview"

    tbl_wrap = ctk.CTkFrame(
        parent_m,
        fg_color="white",
        corner_radius=10,
        border_width=1,
        border_color="#e7cfd3",
    ) if self.materia_use_custom else ttk.Frame(parent_m)
    tbl_wrap.pack(fill="both", expand=True, padx=10, pady=10)

    self.tbl_materia = ttk.Treeview(
        tbl_wrap,
        columns=("lote", "material", "comprimento", "largura", "espessura", "quantidade", "reservado", "formato", "metros", "peso_unid", "p_compra", "preco_unid", "disponivel", "tipo", "local", "id"),
        show="headings",
        style=tbl_style,
    )
    self.tbl_materia.heading("lote", text="Chapa", command=lambda c="lote": self.sort_treeview(self.tbl_materia, c, False))
    self.tbl_materia.heading("material", text="Material", command=lambda c="material": self.sort_treeview(self.tbl_materia, c, False))
    self.tbl_materia.heading("comprimento", text="Comprimento", command=lambda c="comprimento": self.sort_treeview(self.tbl_materia, c, False))
    self.tbl_materia.heading("largura", text="Largura", command=lambda c="largura": self.sort_treeview(self.tbl_materia, c, False))
    self.tbl_materia.heading("espessura", text="Espessura", command=lambda c="espessura": self.sort_treeview(self.tbl_materia, c, False))
    self.tbl_materia.heading("quantidade", text="Quantidade", command=lambda c="quantidade": self.sort_treeview(self.tbl_materia, c, False))
    self.tbl_materia.heading("reservado", text="Reserva", command=lambda c="reservado": self.sort_treeview(self.tbl_materia, c, False))
    self.tbl_materia.heading("formato", text="Formato", command=lambda c="formato": self.sort_treeview(self.tbl_materia, c, False))
    self.tbl_materia.heading("metros", text="Metros (m)", command=lambda c="metros": self.sort_treeview(self.tbl_materia, c, False))
    self.tbl_materia.heading("peso_unid", text="Peso/Un. (kg)", command=lambda c="peso_unid": self.sort_treeview(self.tbl_materia, c, False))
    self.tbl_materia.heading("p_compra", text="Compra (€/kg|€/m)", command=lambda c="p_compra": self.sort_treeview(self.tbl_materia, c, False))
    self.tbl_materia.heading("preco_unid", text="Preço/Unid (€)", command=lambda c="preco_unid": self.sort_treeview(self.tbl_materia, c, False))
    self.tbl_materia.heading("disponivel", text="Disponível", command=lambda c="disponivel": self.sort_treeview(self.tbl_materia, c, False))
    self.tbl_materia.heading("tipo", text="Tipo", command=lambda c="tipo": self.sort_treeview(self.tbl_materia, c, False))
    self.tbl_materia.heading("local", text="Localização", command=lambda c="local": self.sort_treeview(self.tbl_materia, c, False))
    self.tbl_materia.heading("id", text="ID", command=lambda c="id": self.sort_treeview(self.tbl_materia, c, False))
    col_widths = {
        "lote": 125,
        "material": 290,
        "comprimento": 110,
        "largura": 110,
        "espessura": 74,
        "quantidade": 86,
        "reservado": 82,
        "formato": 90,
        "metros": 92,
        "peso_unid": 110,
        "p_compra": 125,
        "preco_unid": 115,
        "disponivel": 100,
        "tipo": 90,
        "local": 120,
        "id": 90,
    }
    for col in self.tbl_materia["columns"]:
        self.tbl_materia.column(col, width=col_widths.get(col, 100))
    self.tbl_materia.column("material", anchor="w")
    self.tbl_materia.tag_configure("even", background="#eef2f8")
    self.tbl_materia.tag_configure("odd", background="#e6ecf5")
    self.tbl_materia.tag_configure("stock_ok", foreground="#000000")
    self.tbl_materia.tag_configure("stock_low", foreground="#000000")
    self.tbl_materia.tag_configure("stock_crit", foreground="#000000")
    self.tbl_materia.tag_configure("stock_one", foreground="#b91c1c")
    if self.materia_use_custom and CUSTOM_TK_AVAILABLE:
        vsb = ctk.CTkScrollbar(tbl_wrap, orientation="vertical", command=self.tbl_materia.yview)
        hsb = ctk.CTkScrollbar(tbl_wrap, orientation="horizontal", command=self.tbl_materia.xview)
    else:
        vsb = ttk.Scrollbar(tbl_wrap, orient="vertical", command=self.tbl_materia.yview)
        hsb = ttk.Scrollbar(tbl_wrap, orient="horizontal", command=self.tbl_materia.xview)
    self.tbl_materia.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    self.tbl_materia.pack(side="left", fill="both", expand=True)
    vsb.pack(side="right", fill="y")
    hsb.pack(side="bottom", fill="x")
    self.tbl_materia.bind("<<TreeviewSelect>>", self.on_select_stock_material)
    self.refresh_materia()

def build_encomendas(self):
    _ensure_configured()
    self.encomendas_use_custom = CUSTOM_TK_AVAILABLE and os.environ.get("USE_CUSTOM_ENC", "1") != "0"
    parent_enc = self.tab_encomendas
    if self.encomendas_use_custom:
        parent_enc = ctk.CTkFrame(self.tab_encomendas, fg_color="#ffffff")

    top = ctk.CTkFrame(
        parent_enc,
        fg_color="white",
        corner_radius=10,
        border_width=1,
        border_color="#e7cfd3",
    ) if self.encomendas_use_custom else ttk.Frame(parent_enc)
    top.pack(fill="x", padx=10, pady=10)

    filter_row = ctk.CTkFrame(
        parent_enc,
        fg_color="white",
        corner_radius=10,
        border_width=1,
        border_color="#e7cfd3",
    ) if self.encomendas_use_custom else ttk.Frame(parent_enc)
    filter_row.pack(fill="x", padx=10, pady=(0, 6))
    (ctk.CTkLabel if self.encomendas_use_custom else ttk.Label)(filter_row, text="Pesquisa").pack(side="left")
    EntryCls = ctk.CTkEntry if self.encomendas_use_custom else ttk.Entry
    self.e_filter = StringVar()
    self.e_filter_entry = EntryCls(filter_row, textvariable=self.e_filter, width=260 if self.encomendas_use_custom else 40)
    self.e_filter_entry.pack(side="left", padx=6, pady=2)
    self.e_filter_entry.bind("<KeyRelease>", lambda _: self.debounce_call("encomendas_filter", self.refresh_encomendas, delay_ms=170))
    self.e_estado_filter = StringVar(value="Ativas")
    self.e_year_filter = StringVar(value=str(datetime.now().year))
    self.e_cliente_filter = StringVar(value="Todos")
    if self.encomendas_use_custom:
        ctk.CTkLabel(filter_row, text="Estado").pack(side="left", padx=(12, 6))
        self.e_estado_segment = ctk.CTkSegmentedButton(
            filter_row,
            values=["Ativas", "Todas", "Preparação", "Em produção", "Concluída"],
            command=lambda v=None: (
                None
                if getattr(self, "_suppress_encomendas_filter_cb", False)
                else (self.e_estado_filter.set(v or "Ativas"), self.refresh_encomendas())
            ),
            width=420,
        )
        self.e_estado_segment.configure(
            fg_color="#f9e8ea",
            selected_color=CTK_PRIMARY_RED,
            selected_hover_color=CTK_PRIMARY_RED_HOVER,
            unselected_color="#f3f7fc",
            unselected_hover_color=THEME_SELECT_BG,
        )
        self.e_estado_segment.pack(side="left", padx=4, pady=2)
        self.e_estado_segment.set("Ativas")
    else:
        ttk.Label(filter_row, text="Estado").pack(side="left", padx=(12, 6))
        self.e_estado_cb = ttk.Combobox(
            filter_row,
            textvariable=self.e_estado_filter,
            values=["Ativas", "Todas", "Preparação", "Em produção", "Concluída"],
            width=16,
            state="readonly",
        )
        self.e_estado_cb.pack(side="left", padx=4, pady=2)
        self.e_estado_cb.bind("<<ComboboxSelected>>", lambda _e=None: self.refresh_encomendas())
    (ctk.CTkLabel if self.encomendas_use_custom else ttk.Label)(filter_row, text="Ano").pack(side="left", padx=(12, 6))
    if self.encomendas_use_custom:
        self.e_year_cb = ctk.CTkComboBox(
            filter_row,
            variable=self.e_year_filter,
            values=[str(datetime.now().year), "Todos"],
            width=120,
            command=self._on_encomendas_year_change,
        )
        self.e_year_cb.pack(side="left", padx=4, pady=2)
    else:
        self.e_year_cb = ttk.Combobox(
            filter_row,
            textvariable=self.e_year_filter,
            values=[str(datetime.now().year), "Todos"],
            width=10,
            state="readonly",
        )
        self.e_year_cb.pack(side="left", padx=4, pady=2)
        self.e_year_cb.bind("<<ComboboxSelected>>", lambda _e=None: self._on_encomendas_year_change(self.e_year_filter.get()))

    (ctk.CTkLabel if self.encomendas_use_custom else ttk.Label)(filter_row, text="Cliente").pack(side="left", padx=(12, 6))
    clientes_filter_vals = ["Todos"] + (self.get_clientes_display() or [])
    if self.encomendas_use_custom:
        self.e_cliente_filter_cb = ctk.CTkComboBox(
            filter_row,
            variable=self.e_cliente_filter,
            values=clientes_filter_vals,
            width=300,
            command=lambda _v=None: (
                None if getattr(self, "_suppress_encomendas_filter_cb", False) else self.refresh_encomendas()
            ),
        )
        self.e_cliente_filter_cb.pack(side="left", padx=4, pady=2)
    else:
        self.e_cliente_filter_cb = ttk.Combobox(
            filter_row,
            textvariable=self.e_cliente_filter,
            values=clientes_filter_vals,
            width=34,
            state="readonly",
        )
        self.e_cliente_filter_cb.pack(side="left", padx=4, pady=2)
        self.e_cliente_filter_cb.bind("<<ComboboxSelected>>", lambda _e=None: self.refresh_encomendas())

    self.e_cliente = StringVar()
    self.e_nota_cliente = StringVar()
    self.e_data_entrega = StringVar()
    self.e_tempo = DoubleVar()
    self.e_cativar = BooleanVar()
    self.e_obs = StringVar()

    LabelCls = ctk.CTkLabel if self.encomendas_use_custom else ttk.Label
    CombCls = ctk.CTkComboBox if self.encomendas_use_custom else ttk.Combobox
    if self.encomendas_use_custom and CUSTOM_TK_AVAILABLE:
        def BtnCls(parent, **kwargs):
            kwargs.setdefault("height", 34)
            kwargs.setdefault("corner_radius", 10)
            kwargs.setdefault("fg_color", CTK_PRIMARY_RED)
            kwargs.setdefault("hover_color", CTK_PRIMARY_RED_HOVER)
            kwargs.setdefault("text_color", "#ffffff")
            kwargs.setdefault("border_width", 0)
            return ctk.CTkButton(parent, **kwargs)
        orange_btn = {"fg_color": "#f59e0b", "hover_color": "#d97706"}
    else:
        BtnCls = ttk.Button
        orange_btn = {}

    LabelCls(top, text="Cliente").grid(row=0, column=0, sticky="w")
    if self.encomendas_use_custom:
        self.e_cliente_cb = CombCls(top, variable=self.e_cliente, values=self.get_clientes_codes() or [""], width=220, state="readonly")
    else:
        self.e_cliente_cb = CombCls(top, textvariable=self.e_cliente, values=self.get_clientes_codes(), width=26)
    self.e_cliente_cb.grid(row=0, column=1, padx=4, pady=2)

    LabelCls(top, text="Tempo estimado (h)").grid(row=0, column=5, sticky="w")
    EntryCls(top, textvariable=self.e_tempo, width=12).grid(row=0, column=6, padx=4, pady=2)

    BtnCls(top, text="Cativar MP", command=self.cativar_stock_selecao).grid(row=1, column=0, padx=(0, 4), pady=2, sticky="w")
    BtnCls(top, text="Descativar MP", command=self.descativar_stock_selecao).grid(row=1, column=1, padx=(4, 6), pady=2, sticky="w")
    LabelCls(top, text="Chapa cativada").grid(row=1, column=4, sticky="w")
    self.e_chapa = StringVar()
    EntryCls(top, textvariable=self.e_chapa, state="readonly", width=(220 if self.encomendas_use_custom else 30)).grid(row=1, column=5, padx=4, pady=2)
    LabelCls(top, text="Observacoes").grid(row=1, column=2, sticky="w")
    EntryCls(top, textvariable=self.e_obs, width=(220 if self.encomendas_use_custom else 22)).grid(row=1, column=3, sticky="w", padx=4, pady=2)

    LabelCls(top, text="Chapas cativadas (todas)").grid(row=0, column=7, sticky="w")
    cativ_wrap = ctk.CTkFrame(top, fg_color="transparent") if self.encomendas_use_custom else ttk.Frame(top)
    cativ_wrap.grid(row=1, column=7, rowspan=2, padx=6, pady=2, sticky="w")
    if self.encomendas_use_custom:
        self.e_chapas_txt = ctk.CTkTextbox(
            cativ_wrap,
            width=520,
            height=120,
            font=("Segoe UI", 11),
            fg_color="#f8fbff",
            border_width=1,
            border_color="#d8b9bf",
        )
        self.e_chapas_txt.pack(side="left", fill="both", expand=True)
    else:
        self.e_chapas_txt = Text(cativ_wrap, height=4, width=80)
        sb_chapas = ttk.Scrollbar(cativ_wrap, orient="vertical", command=self.e_chapas_txt.yview)
        self.e_chapas_txt.configure(yscrollcommand=sb_chapas.set)
        self.e_chapas_txt.pack(side="left")
        sb_chapas.pack(side="right", fill="y")
    self.e_chapas_txt.configure(state="disabled")
    try:
        top.grid_columnconfigure(3, weight=1)
        top.grid_columnconfigure(7, weight=1)
    except Exception:
        pass

    meta = ctk.CTkFrame(
        parent_enc,
        fg_color="white",
        corner_radius=10,
        border_width=1,
        border_color="#e7cfd3",
    ) if self.encomendas_use_custom else ttk.Frame(parent_enc)
    meta.pack(fill="x", padx=10, pady=(0, 6), before=filter_row)

    LabelCls(meta, text="Data entrega").grid(row=0, column=0, sticky="w", padx=(10, 6), pady=6)
    self.e_data_entrega_entry = EntryCls(meta, textvariable=self.e_data_entrega, width=(180 if self.encomendas_use_custom else 22))
    self.e_data_entrega_entry.grid(row=0, column=1, padx=4, pady=6, sticky="w")
    BtnCls(meta, text="Calendário", command=lambda: self.pick_date(self.e_data_entrega)).grid(row=0, column=2, padx=6, pady=6, sticky="w")

    LabelCls(meta, text="Nota cliente").grid(row=0, column=3, sticky="w", padx=(18, 6), pady=6)
    self.e_nota_cliente_entry = EntryCls(meta, textvariable=self.e_nota_cliente, width=(320 if self.encomendas_use_custom else 36))
    self.e_nota_cliente_entry.grid(row=0, column=4, padx=4, pady=6, sticky="w")
    self.e_nota_cliente_entry.bind("<FocusOut>", self.save_nota_cliente_encomenda)
    self.e_nota_cliente_entry.bind("<Return>", self.save_nota_cliente_encomenda)
    if self.encomendas_use_custom:
        self.e_nota_status_lbl = ctk.CTkLabel(meta, text="", width=150, anchor="w")
    else:
        self.e_nota_status_lbl = ttk.Label(meta, text="")
    self.e_nota_status_lbl.grid(row=0, column=5, padx=(6, 10), pady=6, sticky="w")
    self._update_nota_cliente_visual()

    self.e_data_entrega_entry.bind("<FocusOut>", self.save_data_entrega_encomenda)
    self.e_data_entrega_entry.bind("<Return>", self.save_data_entrega_encomenda)
    self.e_data_entrega.trace_add("write", lambda *_: self.debounce_call("encomendas_data_entrega", self.save_data_entrega_encomenda, delay_ms=320))

    try:
        meta.grid_columnconfigure(4, weight=1)
    except Exception:
        pass

    btns = ctk.CTkFrame(
        parent_enc,
        fg_color="white",
        corner_radius=10,
        border_width=1,
        border_color="#e7cfd3",
    ) if self.encomendas_use_custom else ttk.Frame(parent_enc)
    btns.pack(fill="x", padx=10)
    BtnCls(btns, text="Criar encomenda", command=self.add_encomenda, **orange_btn).pack(side="left", padx=4, pady=2)
    BtnCls(btns, text="Editar", command=self.edit_encomenda).pack(side="left", padx=4, pady=2)
    BtnCls(btns, text="Atualizar", command=self.refresh).pack(side="left", padx=4, pady=2)
    BtnCls(btns, text="Reabrir encomenda", command=self.reabrir_encomenda).pack(side="left", padx=4, pady=2)
    BtnCls(btns, text="Remover encomenda", command=self.remove_encomenda).pack(side="left", padx=4, pady=2)
    BtnCls(btns, text="Libertar reserva", command=self.libertar_reserva).pack(side="left", padx=4, pady=2)
    BtnCls(btns, text="Exportar CSV", command=lambda: self.export_csv("encomendas")).pack(side="left", padx=4, pady=2)
    BtnCls(btns, text="Pre-visualizar Encomenda", command=self.preview_encomenda).pack(side="left", padx=4, pady=2)

    resumo = ctk.CTkFrame(
        parent_enc,
        fg_color="white",
        corner_radius=10,
        border_width=1,
        border_color="#e7cfd3",
    ) if self.encomendas_use_custom else ttk.LabelFrame(parent_enc, text="Informações da encomenda")
    resumo.pack(fill="x", padx=10, pady=(6, 2))
    self.enc_info_numero = StringVar(value="-")
    self.enc_info_cliente = StringVar(value="-")
    self.enc_info_entrega = StringVar(value="-")
    self.enc_info_estado = StringVar(value="-")
    self.enc_info_nota = StringVar(value="-")
    LabelInfo = ctk.CTkLabel if self.encomendas_use_custom else ttk.Label
    LabelInfo(resumo, text="Número").grid(row=0, column=0, sticky="w", padx=8, pady=4)
    LabelInfo(resumo, textvariable=self.enc_info_numero).grid(row=0, column=1, sticky="w", padx=8, pady=4)
    LabelInfo(resumo, text="Cliente").grid(row=0, column=2, sticky="w", padx=8, pady=4)
    LabelInfo(resumo, textvariable=self.enc_info_cliente).grid(row=0, column=3, sticky="w", padx=8, pady=4)
    LabelInfo(resumo, text="Entrega").grid(row=1, column=0, sticky="w", padx=8, pady=4)
    LabelInfo(resumo, textvariable=self.enc_info_entrega).grid(row=1, column=1, sticky="w", padx=8, pady=4)
    LabelInfo(resumo, text="Estado").grid(row=1, column=2, sticky="w", padx=8, pady=4)
    LabelInfo(resumo, textvariable=self.enc_info_estado).grid(row=1, column=3, sticky="w", padx=8, pady=4)
    LabelInfo(resumo, text="Nota").grid(row=2, column=0, sticky="w", padx=8, pady=4)
    LabelInfo(resumo, textvariable=self.enc_info_nota).grid(row=2, column=1, columnspan=3, sticky="w", padx=8, pady=4)
    LabelInfo(resumo, text="Cativar MP").grid(row=0, column=4, sticky="w", padx=(18, 6), pady=4)
    if self.encomendas_use_custom and CUSTOM_TK_AVAILABLE:
        self.enc_cativar_switch = ctk.CTkSwitch(
            resumo,
            text="",
            variable=self.e_cativar,
            command=self.on_cativar_toggle,
            onvalue=True,
            offvalue=False,
            width=64,
        )
        self.enc_cativar_switch.grid(row=0, column=5, sticky="w", padx=4, pady=4)
    else:
        self.enc_cativar_switch = ttk.Checkbutton(
            resumo,
            variable=self.e_cativar,
            command=self.on_cativar_toggle,
        )
        self.enc_cativar_switch.grid(row=0, column=5, sticky="w", padx=4, pady=4)

    tbl_style = ""
    sub_tbl_style = ""
    if self.encomendas_use_custom:
        style = ttk.Style()
        style.configure(
            "Encomendas.Treeview",
            font=("Segoe UI", 11),
            rowheight=31,
            background="#f8fbff",
            fieldbackground="#f8fbff",
            borderwidth=0,
            bordercolor="#d7deea",
            lightcolor="#d7deea",
            darkcolor="#d7deea",
            relief="solid",
        )
        style.configure(
            "Encomendas.Treeview.Heading",
            font=("Segoe UI", 11, "bold"),
            background=THEME_HEADER_BG,
            foreground="white",
            relief="flat",
        )
        style.map("Encomendas.Treeview.Heading", background=[("active", THEME_HEADER_ACTIVE)])
        style.map(
            "Encomendas.Treeview",
            background=[("selected", THEME_SELECT_BG)],
            foreground=[("selected", THEME_SELECT_FG)],
        )
        style.configure(
            "EncomendasSub.Treeview",
            font=("Segoe UI", 11),
            rowheight=30,
            background="#f8fbff",
            fieldbackground="#f8fbff",
            borderwidth=0,
            bordercolor="#d7deea",
            lightcolor="#d7deea",
            darkcolor="#d7deea",
            relief="solid",
        )
        style.configure(
            "EncomendasSub.Treeview.Heading",
            font=("Segoe UI", 11, "bold"),
            background=THEME_HEADER_BG,
            foreground="white",
            relief="flat",
        )
        style.map("EncomendasSub.Treeview.Heading", background=[("active", THEME_HEADER_ACTIVE)])
        style.map(
            "EncomendasSub.Treeview",
            background=[("selected", THEME_SELECT_BG)],
            foreground=[("selected", THEME_SELECT_FG)],
        )
        tbl_style = "Encomendas.Treeview"
        sub_tbl_style = "EncomendasSub.Treeview"

    tbl_wrap = ctk.CTkFrame(
        parent_enc,
        fg_color="white",
        corner_radius=10,
        border_width=1,
        border_color="#e7cfd3",
    ) if self.encomendas_use_custom else ttk.Frame(parent_enc)
    tbl_wrap.pack(fill="x", padx=10, pady=(8, 6))
    try:
        tbl_wrap.configure(height=300)
        tbl_wrap.pack_propagate(False)
    except Exception:
        pass
    self.tbl_encomendas = ttk.Treeview(tbl_wrap, columns=("numero", "nota_cliente", "cliente", "data_criacao", "data_entrega", "tempo", "estado", "cativar"), show="headings", style=tbl_style)
    self.tbl_encomendas.heading("numero", text="Número", command=lambda c="numero": self.sort_treeview(self.tbl_encomendas, c, False))
    self.tbl_encomendas.heading("nota_cliente", text="Nota Cliente", command=lambda c="nota_cliente": self.sort_treeview(self.tbl_encomendas, c, False))
    self.tbl_encomendas.heading("cliente", text="Cliente", command=lambda c="cliente": self.sort_treeview(self.tbl_encomendas, c, False))
    self.tbl_encomendas.heading("data_criacao", text="Data Criação", command=lambda c="data_criacao": self.sort_treeview(self.tbl_encomendas, c, False))
    self.tbl_encomendas.heading("data_entrega", text="Data Entrega", command=lambda c="data_entrega": self.sort_treeview(self.tbl_encomendas, c, False))
    self.tbl_encomendas.heading("tempo", text="Tempo", command=lambda c="tempo": self.sort_treeview(self.tbl_encomendas, c, False))
    self.tbl_encomendas.heading("estado", text="Estado", command=lambda c="estado": self.sort_treeview(self.tbl_encomendas, c, False))
    self.tbl_encomendas.heading("cativar", text="Cativar", command=lambda c="cativar": self.sort_treeview(self.tbl_encomendas, c, False))
    self.tbl_encomendas.column("numero", width=110)
    self.tbl_encomendas.column("nota_cliente", width=140)
    self.tbl_encomendas.column("cliente", width=200)
    self.tbl_encomendas.column("data_criacao", width=120)
    self.tbl_encomendas.column("data_entrega", width=120)
    self.tbl_encomendas.column("tempo", width=80)
    self.tbl_encomendas.column("estado", width=90)
    self.tbl_encomendas.column("cativar", width=70)
    self.tbl_encomendas.tag_configure("even", background="#f7fbff")
    self.tbl_encomendas.tag_configure("odd", background="#f1f7fe")
    self.tbl_encomendas.tag_configure("estado_preparacao_even", background="#fbecee", foreground="#7a0f1a")
    self.tbl_encomendas.tag_configure("estado_preparacao_odd", background="#f8e5e8", foreground="#7a0f1a")
    self.tbl_encomendas.tag_configure("estado_producao_even", background="#ffe7c1", foreground="#8a4b00")
    self.tbl_encomendas.tag_configure("estado_producao_odd", background="#ffe1b2", foreground="#8a4b00")
    self.tbl_encomendas.tag_configure("estado_concluida_even", background="#d9f2df", foreground="#0f5132")
    self.tbl_encomendas.tag_configure("estado_concluida_odd", background="#cfe9d5", foreground="#0f5132")
    self.tbl_encomendas.pack(side="left", fill="both", expand=True)
    if self.encomendas_use_custom and CUSTOM_TK_AVAILABLE:
        sb_en = ctk.CTkScrollbar(tbl_wrap, orientation="vertical", command=self.tbl_encomendas.yview)
        sb_enx = ctk.CTkScrollbar(tbl_wrap, orientation="horizontal", command=self.tbl_encomendas.xview)
    else:
        sb_en = ttk.Scrollbar(tbl_wrap, orient="vertical", command=self.tbl_encomendas.yview)
        sb_enx = ttk.Scrollbar(tbl_wrap, orient="horizontal", command=self.tbl_encomendas.xview)
    self.tbl_encomendas.configure(yscrollcommand=sb_en.set, xscrollcommand=sb_enx.set)
    sb_en.pack(side="right", fill="y")
    sb_enx.pack(side="bottom", fill="x")
    self.tbl_encomendas.bind("<<TreeviewSelect>>", self.on_select_encomenda)
    self.tbl_encomendas.bind("<Double-1>", lambda _e=None: self.edit_encomenda())
    self.bind_clear_on_empty(self.tbl_encomendas, self.clear_encomenda_selection)

    mid = ctk.CTkFrame(parent_enc, fg_color="transparent") if self.encomendas_use_custom else ttk.Frame(parent_enc)
    mid.pack(fill="x", padx=10, pady=(0, 6))
    try:
        mid.configure(height=268)
        mid.pack_propagate(False)
    except Exception:
        pass
    left = ctk.CTkFrame(
        mid,
        fg_color="white",
        corner_radius=10,
        border_width=1,
        border_color="#e7cfd3",
    ) if self.encomendas_use_custom else ttk.Frame(mid)
    left.pack(side="left", fill="both", expand=True)
    right = ctk.CTkFrame(
        mid,
        fg_color="white",
        corner_radius=10,
        border_width=1,
        border_color="#e7cfd3",
    ) if self.encomendas_use_custom else ttk.Frame(mid)
    right.pack(side="left", fill="both", expand=True, padx=(10, 0))

    (ctk.CTkLabel if self.encomendas_use_custom else ttk.Label)(left, text="Materiais").pack(anchor="w", padx=6, pady=(6, 2))
    self.tbl_materiais = ttk.Treeview(left, columns=("material", "estado"), show="headings", height=3, style=sub_tbl_style)
    self.tbl_materiais.heading("material", text="Material")
    self.tbl_materiais.heading("estado", text="Estado")
    self.tbl_materiais.column("material", width=130)
    self.tbl_materiais.column("estado", width=90)
    self.tbl_materiais.pack(side="left", fill="both", expand=True)
    if self.encomendas_use_custom and CUSTOM_TK_AVAILABLE:
        sb_mat = ctk.CTkScrollbar(left, orientation="vertical", command=self.tbl_materiais.yview)
        sb_matx = ctk.CTkScrollbar(left, orientation="horizontal", command=self.tbl_materiais.xview)
    else:
        sb_mat = ttk.Scrollbar(left, orient="vertical", command=self.tbl_materiais.yview)
        sb_matx = ttk.Scrollbar(left, orient="horizontal", command=self.tbl_materiais.xview)
    self.tbl_materiais.configure(yscrollcommand=sb_mat.set, xscrollcommand=sb_matx.set)
    sb_mat.pack(side="right", fill="y")
    sb_matx.pack(side="bottom", fill="x")
    self.tbl_materiais.bind("<<TreeviewSelect>>", self.on_select_material)
    self.bind_clear_on_empty(self.tbl_materiais, self.clear_material_selection)
    mat_btns = ctk.CTkFrame(left, fg_color="transparent") if self.encomendas_use_custom else ttk.Frame(left)
    mat_btns.pack(fill="x", pady=(4, 0))
    BtnCls(mat_btns, text="Adicionar material", command=self.add_material_encomenda).pack(fill="x", padx=4, pady=(0, 4))
    BtnCls(mat_btns, text="Remover material", command=self.remove_material_encomenda).pack(fill="x", padx=4)

    (ctk.CTkLabel if self.encomendas_use_custom else ttk.Label)(right, text="Espessuras").pack(anchor="w", padx=6, pady=(6, 2))
    self.tbl_espessuras = ttk.Treeview(right, columns=("espessura", "tempo", "estado"), show="headings", height=3, style=sub_tbl_style)
    self.tbl_espessuras.heading("espessura", text="Espessura")
    self.tbl_espessuras.heading("tempo", text="Tempo")
    self.tbl_espessuras.heading("estado", text="Estado")
    self.tbl_espessuras.column("espessura", width=75)
    self.tbl_espessuras.column("tempo", width=90)
    self.tbl_espessuras.column("estado", width=90)
    self.tbl_espessuras.pack(side="left", fill="both", expand=True)
    if self.encomendas_use_custom and CUSTOM_TK_AVAILABLE:
        sb_esp = ctk.CTkScrollbar(right, orientation="vertical", command=self.tbl_espessuras.yview)
        sb_espx = ctk.CTkScrollbar(right, orientation="horizontal", command=self.tbl_espessuras.xview)
    else:
        sb_esp = ttk.Scrollbar(right, orient="vertical", command=self.tbl_espessuras.yview)
        sb_espx = ttk.Scrollbar(right, orient="horizontal", command=self.tbl_espessuras.xview)
    self.tbl_espessuras.configure(yscrollcommand=sb_esp.set, xscrollcommand=sb_espx.set)
    sb_esp.pack(side="right", fill="y")
    sb_espx.pack(side="bottom", fill="x")
    self.tbl_espessuras.bind("<<TreeviewSelect>>", self.on_select_espessura)
    self.bind_clear_on_empty(self.tbl_espessuras, self.clear_espessura_selection)
    esp_btns = ctk.CTkFrame(right, fg_color="transparent") if self.encomendas_use_custom else ttk.Frame(right)
    esp_btns.pack(fill="x", pady=(2, 0))
    BtnCls(esp_btns, text="Adicionar espessura", command=self.add_espessura, height=30 if self.encomendas_use_custom else None).pack(fill="x", padx=4, pady=(0, 2))
    BtnCls(esp_btns, text="Remover espessura", command=self.remove_espessura, height=30 if self.encomendas_use_custom else None).pack(fill="x", padx=4, pady=(0, 2))
    BtnCls(esp_btns, text="Adicionar tempo", command=self.edit_tempo_espessura, height=30 if self.encomendas_use_custom else None).pack(fill="x", padx=4, pady=(0, 2))

    peca_btns = ctk.CTkFrame(right, fg_color="transparent") if self.encomendas_use_custom else ttk.Frame(right)
    peca_btns.pack(fill="x", pady=(2, 0))
    BtnCls(peca_btns, text="Adicionar peca", command=self.add_peca, height=30 if self.encomendas_use_custom else None).pack(fill="x", padx=4, pady=(0, 2))
    BtnCls(peca_btns, text="Remover Peça", command=self.remove_peca, height=30 if self.encomendas_use_custom else None).pack(fill="x", padx=4, pady=(0, 2))
    BtnCls(peca_btns, text="Ver desenho", command=self.open_selected_peca_desenho, height=30 if self.encomendas_use_custom else None).pack(fill="x", padx=4, pady=(0, 2))

    pecas_wrap = ctk.CTkFrame(
        parent_enc,
        fg_color="white",
        corner_radius=10,
        border_width=1,
        border_color="#e7cfd3",
    ) if self.encomendas_use_custom else ttk.Frame(parent_enc)
    pecas_wrap.pack(fill="both", expand=True, padx=6, pady=(0, 8))
    self.tbl_pecas = ttk.Treeview(pecas_wrap, columns=("ref_int", "ref_ext", "material", "espessura", "Operações", "qtd_pedida", "estado"), show="headings", style=sub_tbl_style)
    for col in self.tbl_pecas["columns"]:
        self.tbl_pecas.heading(col, text=col, command=lambda c=col: self.sort_treeview(self.tbl_pecas, c, False))
    self.tbl_pecas.heading("ref_int", text="Referência Interna")
    self.tbl_pecas.heading("ref_ext", text="Referência Externa")
    self.tbl_pecas.column("ref_int", width=260, stretch=True)
    self.tbl_pecas.column("ref_ext", width=300, stretch=True)
    self.tbl_pecas.column("material", width=110, stretch=True)
    self.tbl_pecas.column("espessura", width=85, stretch=True)
    self.tbl_pecas.column("Operações", width=220, stretch=True)
    self.tbl_pecas.column("qtd_pedida", width=90, stretch=True)
    self.tbl_pecas.column("estado", width=105, stretch=True)
    self.tbl_pecas.tag_configure("even", background="#f7fbff")
    self.tbl_pecas.tag_configure("odd", background="#f1f7fe")
    self.tbl_pecas.tag_configure("estado_preparacao_even", background="#fbecee", foreground="#7a0f1a")
    self.tbl_pecas.tag_configure("estado_preparacao_odd", background="#f8e5e8", foreground="#7a0f1a")
    self.tbl_pecas.tag_configure("estado_producao_even", background="#ffe7c1", foreground="#8a4b00")
    self.tbl_pecas.tag_configure("estado_producao_odd", background="#ffe1b2", foreground="#8a4b00")
    self.tbl_pecas.tag_configure("estado_concluida_even", background="#d9f2df", foreground="#0f5132")
    self.tbl_pecas.tag_configure("estado_concluida_odd", background="#cfe9d5", foreground="#0f5132")
    self.tbl_pecas.tag_configure("esp_1", background="#e3f2fd")
    self.tbl_pecas.tag_configure("esp_2", background="#fff3e0")
    self.tbl_pecas.tag_configure("esp_3", background="#e8f5e9")
    self.tbl_pecas.tag_configure("esp_4", background="#f3e5f5")
    self.tbl_pecas.pack(side="left", fill="both", expand=True)
    self.bind_clear_on_empty(self.tbl_pecas, self.clear_peca_selection)
    if self.encomendas_use_custom and CUSTOM_TK_AVAILABLE:
        sb_pv = ctk.CTkScrollbar(pecas_wrap, orientation="vertical", command=self.tbl_pecas.yview)
        sb_pvx = ctk.CTkScrollbar(pecas_wrap, orientation="horizontal", command=self.tbl_pecas.xview)
    else:
        sb_pv = ttk.Scrollbar(pecas_wrap, orient="vertical", command=self.tbl_pecas.yview)
        sb_pvx = ttk.Scrollbar(pecas_wrap, orient="horizontal", command=self.tbl_pecas.xview)
    self.tbl_pecas.configure(yscrollcommand=sb_pv.set, xscrollcommand=sb_pvx.set)
    sb_pv.pack(side="right", fill="y")
    sb_pvx.pack(side="bottom", fill="x")

    self.refresh_encomendas_year_options(keep_selection=False)
    self._on_encomendas_year_change(self.e_year_filter.get())

    # Interface principal simplificada: todo o detalhe fica no editor custom (Criar/Editar).
    self.encomendas_advanced_frames = [top, meta, mid, pecas_wrap]
    for _f in self.encomendas_advanced_frames:
        try:
            _f.pack_forget()
        except Exception:
            pass
    if self.encomendas_use_custom:
        # Mostrar o container apenas no fim evita o "flash" do layout antigo.
        parent_enc.pack(fill="both", expand=True)

def build_plano(self):
    _ensure_configured()
    self.plano_use_custom = CUSTOM_TK_AVAILABLE and os.environ.get("USE_CUSTOM_PLANO", "1") != "0"
    container = self.tab_plano
    if self.plano_use_custom:
        container = ctk.CTkFrame(self.tab_plano, fg_color="#ffffff")
        container.pack(fill="both", expand=True)
        try:
            s = ttk.Style()
            s.theme_use("clam")
            s.configure(
                "Plan.Treeview",
                background="#fff8f9",
                fieldbackground="#fff8f9",
                rowheight=24,
                borderwidth=0,
            )
            s.configure(
                "Plan.Treeview.Heading",
                background=THEME_HEADER_BG,
                foreground="white",
                font=("Segoe UI", 9, "bold"),
                relief="flat",
            )
            s.map("Plan.Treeview.Heading", background=[("active", THEME_HEADER_ACTIVE)])
            s.map(
                "Plan.Treeview",
                background=[("selected", THEME_SELECT_BG)],
                foreground=[("selected", THEME_SELECT_FG)],
            )
        except Exception:
            pass

    top = (
        ctk.CTkFrame(
            container,
            fg_color="#ffffff",
            corner_radius=10,
            border_width=1,
            border_color="#e7cfd3",
        )
        if self.plano_use_custom
        else ttk.Frame(container)
    )
    top.pack(fill="x", padx=10, pady=10)

    self.p_inicio = StringVar(value="08:00")
    self.p_fim = StringVar(value="18:00")
    (ctk.CTkLabel if self.plano_use_custom else ttk.Label)(top, text="Horário").pack(side="left")
    (ctk.CTkEntry if self.plano_use_custom else ttk.Entry)(top, textvariable=self.p_inicio, width=70 if self.plano_use_custom else 6).pack(side="left", padx=4)
    (ctk.CTkLabel if self.plano_use_custom else ttk.Label)(top, text="até").pack(side="left")
    (ctk.CTkEntry if self.plano_use_custom else ttk.Entry)(top, textvariable=self.p_fim, width=70 if self.plano_use_custom else 6).pack(side="left", padx=4)

    self.p_week_start = week_start(datetime.now().date())
    if self.plano_use_custom:
        ctk.CTkButton(
            top,
            text="Semana -",
            command=self.prev_week,
            width=96,
            height=34,
            corner_radius=10,
            fg_color=CTK_PRIMARY_RED,
            hover_color=CTK_PRIMARY_RED_HOVER,
            text_color="#ffffff",
            border_width=0,
        ).pack(side="left", padx=10, pady=2)
    else:
        ttk.Button(top, text="Semana -", command=self.prev_week).pack(side="left", padx=10)
    self.p_week_lbl = (ctk.CTkLabel if self.plano_use_custom else ttk.Label)(top, text=self.p_week_start.strftime("%d/%m/%Y"))
    self.p_week_lbl.pack(side="left")
    if self.plano_use_custom:
        ctk.CTkButton(
            top,
            text="Semana +",
            command=self.next_week,
            width=96,
            height=34,
            corner_radius=10,
            fg_color=CTK_PRIMARY_RED,
            hover_color=CTK_PRIMARY_RED_HOVER,
            text_color="#ffffff",
            border_width=0,
        ).pack(side="left", padx=6, pady=2)
    else:
        ttk.Button(top, text="Semana +", command=self.next_week).pack(side="left", padx=6)

    btn_cls = ctk.CTkButton if self.plano_use_custom else ttk.Button
    if self.plano_use_custom:
        for txt, cmd, w in (
            ("Atualizar", self.refresh_plano, 104),
            ("Exportar CSV", lambda: self.export_csv("plano"), 118),
            ("Pre-visualizar A4", self.preview_plano_a4, 136),
            ("Auto planear", self.auto_planear, 118),
            ("Desplanear", self.desplanear_tudo, 118),
        ):
            btn_cls(
                top,
                text=txt,
                command=cmd,
                width=w,
                height=34,
                corner_radius=10,
                fg_color=CTK_PRIMARY_RED,
                hover_color=CTK_PRIMARY_RED_HOVER,
                text_color="#ffffff",
                border_width=0,
            ).pack(side="left", padx=(10 if txt == "Atualizar" else 6), pady=2)
    else:
        btn_cls(top, text="Atualizar", command=self.refresh_plano).pack(side="left", padx=10)
        btn_cls(top, text="Exportar CSV", command=lambda: self.export_csv("plano")).pack(side="left", padx=6)
        btn_cls(top, text="Pre-visualizar A4", command=self.preview_plano_a4).pack(side="left", padx=6)
        btn_cls(top, text="Auto planear", command=self.auto_planear).pack(side="left", padx=6)
        btn_cls(top, text="Desplanear", command=self.desplanear_tudo).pack(side="left", padx=6)

    body = ctk.CTkFrame(container, fg_color="#ffffff") if self.plano_use_custom else ttk.Frame(container)
    body.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    left = (
        ctk.CTkFrame(
            body,
            fg_color="#ffffff",
            corner_radius=10,
            border_width=1,
            border_color="#e7cfd3",
        )
        if self.plano_use_custom
        else ttk.Frame(body)
    )
    left.pack(side="left", fill="y")
    (ctk.CTkLabel if self.plano_use_custom else ttk.Label)(left, text="Planeamento", font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=4, pady=(6, 2))
    filt_row = ctk.CTkFrame(left, fg_color="#ffffff") if self.plano_use_custom else ttk.Frame(left)
    filt_row.pack(fill="x", pady=(4, 6))
    (ctk.CTkLabel if self.plano_use_custom else ttk.Label)(filt_row, text="Filtro").pack(side="left")
    self.p_filter = StringVar()
    (ctk.CTkEntry if self.plano_use_custom else ttk.Entry)(filt_row, textvariable=self.p_filter, width=180 if self.plano_use_custom else 18).pack(side="left", padx=6)
    if self.plano_use_custom:
        filt_row.winfo_children()[-1].bind("<KeyRelease>", lambda _=None: self.debounce_call("plano_filter", self.refresh_plano, delay_ms=170))
    else:
        self.p_filter_entry = filt_row.winfo_children()[-1]
        self.p_filter_entry.bind("<KeyRelease>", lambda _=None: self.debounce_call("plano_filter", self.refresh_plano, delay_ms=170))
    self.p_status_filter = StringVar(value="Pendentes")
    if self.plano_use_custom:
        self.p_status_segment = ctk.CTkSegmentedButton(
            filt_row,
            values=["Pendentes", "Concluídas", "Todas"],
            command=lambda v=None: (self.p_status_filter.set(v or "Pendentes"), self.refresh_plano()),
            width=320,
        )
        self.p_status_segment.configure(
            fg_color="#f9e8ea",
            selected_color=CTK_PRIMARY_RED,
            selected_hover_color=CTK_PRIMARY_RED_HOVER,
            unselected_color="#fff5f6",
            unselected_hover_color=THEME_SELECT_BG,
        )
        self.p_status_segment.pack(side="left", padx=8)
        self.p_status_segment.set("Pendentes")
    else:
        self.p_status_cb = ttk.Combobox(
            filt_row,
            textvariable=self.p_status_filter,
            values=["Pendentes", "Concluídas", "Todas"],
            state="readonly",
            width=14,
        )
        self.p_status_cb.pack(side="left", padx=8)
        self.p_status_cb.bind("<<ComboboxSelected>>", lambda _e=None: self.refresh_plano())

    self.tbl_plano_enc = ttk.Treeview(left, columns=("numero", "material", "espessura", "tempo"), show="headings", height=32, style="Plan.Treeview" if self.plano_use_custom else "")
    self.tbl_plano_enc.heading("numero", text="Encomenda")
    self.tbl_plano_enc.heading("material", text="Material")
    self.tbl_plano_enc.heading("espessura", text="Espessura")
    self.tbl_plano_enc.heading("tempo", text="Tempo (min)")
    self.tbl_plano_enc.column("numero", width=120)
    self.tbl_plano_enc.column("material", width=140)
    self.tbl_plano_enc.column("espessura", width=80)
    self.tbl_plano_enc.column("tempo", width=90)
    self.tbl_plano_enc.pack(fill="y", padx=(0, 10))
    self.tbl_plano_enc.bind("<ButtonPress-1>", self.on_plano_drag_start)
    self.tbl_plano_enc.tag_configure("pl_even", background="#fff5f6")
    self.tbl_plano_enc.tag_configure("pl_odd", background="#fbecee")
    self.tbl_plano_enc.tag_configure("pl_warn", background="#fff6cc")
    if self.plano_use_custom:
        ctk.CTkScrollbar(
            left,
            command=self.tbl_plano_enc.yview,
            fg_color="#f9e8ea",
            button_color=CTK_PRIMARY_RED,
            button_hover_color=CTK_PRIMARY_RED_HOVER,
        ).pack(side="right", fill="y")
        self.tbl_plano_enc.configure(yscrollcommand=lambda *args: None)  # basic, visual only

    self.plano_canvas = (
        ctk.CTkFrame(
            body,
            fg_color="#ffffff",
            corner_radius=10,
            border_width=1,
            border_color="#e7cfd3",
        )
        if self.plano_use_custom
        else ttk.Frame(body)
    )
    self.plano_canvas.pack(side="left", fill="both", expand=True)

    self.plano_grid = ttk.Treeview(self.plano_canvas, columns=(), show="tree")
    self.plano_grid.pack_forget()

    self.plano = Canvas(self.plano_canvas, bg="white", highlightthickness=0)
    self.plano.pack(fill="both", expand=True, padx=4, pady=4)
    self.plano.bind("<Button-1>", self.on_plano_click)
    self.plano.bind("<B1-Motion>", self.on_plano_drag_motion)
    self.plano.bind("<ButtonRelease-1>", self.on_plano_drop)

    self.refresh_plano()

def build_qualidade(self):
    _ensure_configured()
    self.qual_use_custom = CUSTOM_TK_AVAILABLE and os.environ.get("USE_CUSTOM_QUALIDADE", "1") != "0"
    FrameCls = ctk.CTkFrame if self.qual_use_custom else ttk.Frame
    LabelCls = ctk.CTkLabel if self.qual_use_custom else ttk.Label
    EntryCls = ctk.CTkEntry if self.qual_use_custom else ttk.Entry
    BtnCls = ctk.CTkButton if self.qual_use_custom else ttk.Button

    parent_q = self.tab_qualidade
    if self.qual_use_custom:
        parent_q = ctk.CTkFrame(self.tab_qualidade, fg_color="#fff8f9")
        parent_q.pack(fill="both", expand=True)

    top = (
        FrameCls(
            parent_q,
            fg_color="white",
            corner_radius=10,
            border_width=1,
            border_color="#e7cfd3",
        )
        if self.qual_use_custom
        else FrameCls(parent_q)
    )
    top.pack(fill="x", padx=10, pady=(10, 6))
    if self.qual_use_custom:
        LabelCls(
            top,
            text="Controlo de Qualidade",
            font=("Segoe UI", 13, "bold"),
            text_color="#7f1b2c",
        ).pack(side="left", padx=(6, 10))
    LabelCls(top, text="Pesquisa").pack(side="left")
    self.q_filter = StringVar()
    self.q_filter_entry = EntryCls(top, textvariable=self.q_filter, width=260 if self.qual_use_custom else 40)
    self.q_filter_entry.pack(side="left", padx=6, pady=2)
    self.q_filter_entry.bind("<KeyRelease>", lambda _: self.debounce_call("qualidade_filter", self.refresh_qualidade, delay_ms=170))
    try:
        self.q_filter_entry.bind("<Return>", lambda _e=None: self.refresh_qualidade())
    except Exception:
        pass
    if self.qual_use_custom:
        BtnCls(top, text="Atualizar", command=self.refresh_qualidade, width=100).pack(side="left", padx=6, pady=2)
        BtnCls(top, text="Exportar CSV", command=lambda: self.export_csv("qualidade"), width=120).pack(side="left", padx=6, pady=2)
    else:
        BtnCls(top, text="Atualizar", command=self.refresh_qualidade).pack(side="left", padx=6, pady=2)
        BtnCls(top, text="Exportar CSV", command=lambda: self.export_csv("qualidade")).pack(side="left", padx=6, pady=2)

    qual_style = ""
    if self.qual_use_custom:
        style = ttk.Style()
        style.configure(
            "Qual.Treeview",
            font=("Segoe UI", 10),
            rowheight=26,
            background="#f8fbff",
            fieldbackground="#f8fbff",
            borderwidth=0,
        )
        style.configure(
            "Qual.Treeview.Heading",
            font=("Segoe UI", 10, "bold"),
            background=THEME_HEADER_BG,
            foreground="white",
            relief="flat",
        )
        style.map("Qual.Treeview.Heading", background=[("active", THEME_HEADER_ACTIVE)])
        style.map(
            "Qual.Treeview",
            background=[("selected", THEME_SELECT_BG)],
            foreground=[("selected", THEME_SELECT_FG)],
        )
        qual_style = "Qual.Treeview"

    tbl_wrap = (
        ctk.CTkFrame(
            parent_q,
            fg_color="white",
            corner_radius=10,
            border_width=1,
            border_color="#e7cfd3",
        )
        if self.qual_use_custom
        else ttk.Frame(parent_q)
    )
    tbl_wrap.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    self.tbl_qualidade = ttk.Treeview(tbl_wrap, columns=("encomenda", "peca", "ok", "nok", "motivo", "data"), show="headings", style=qual_style)
    self.tbl_qualidade.heading("encomenda", text="Encomenda", command=lambda c="encomenda": self.sort_treeview(self.tbl_qualidade, c, False))
    self.tbl_qualidade.heading("peca", text="Peça", command=lambda c="peca": self.sort_treeview(self.tbl_qualidade, c, False))
    self.tbl_qualidade.heading("ok", text="OK", command=lambda c="ok": self.sort_treeview(self.tbl_qualidade, c, False))
    self.tbl_qualidade.heading("nok", text="NOK", command=lambda c="nok": self.sort_treeview(self.tbl_qualidade, c, False))
    self.tbl_qualidade.heading("motivo", text="Motivo", command=lambda c="motivo": self.sort_treeview(self.tbl_qualidade, c, False))
    self.tbl_qualidade.heading("data", text="Data", command=lambda c="data": self.sort_treeview(self.tbl_qualidade, c, False))
    self.tbl_qualidade.column("encomenda", width=140)
    self.tbl_qualidade.column("peca", width=180)
    self.tbl_qualidade.column("ok", width=70)
    self.tbl_qualidade.column("nok", width=70)
    self.tbl_qualidade.column("motivo", width=220)
    self.tbl_qualidade.column("data", width=160)
    if self.qual_use_custom:
        self.tbl_qualidade.tag_configure("even", background="#eaf3ff")
        self.tbl_qualidade.tag_configure("odd", background="#fff8f9")
    else:
        self.tbl_qualidade.tag_configure("even", background="#fbecee")
        self.tbl_qualidade.tag_configure("odd", background="#fff5f6")
    vsb = ttk.Scrollbar(tbl_wrap, orient="vertical", command=self.tbl_qualidade.yview)
    hsb = ttk.Scrollbar(tbl_wrap, orient="horizontal", command=self.tbl_qualidade.xview)
    self.tbl_qualidade.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    self.tbl_qualidade.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")
    tbl_wrap.rowconfigure(0, weight=1)
    tbl_wrap.columnconfigure(0, weight=1)
    self.refresh_qualidade()

def build_orc(self):
    _ensure_configured()
    self.orc_use_custom = CUSTOM_TK_AVAILABLE and os.environ.get("USE_CUSTOM_ORC", "1") != "0"
    if self.orc_use_custom and CUSTOM_TK_AVAILABLE:
        def BtnCls(parent, **kwargs):
            kwargs.setdefault("width", 122)
            kwargs.setdefault("height", 34)
            kwargs.setdefault("corner_radius", 10)
            kwargs.setdefault("fg_color", CTK_PRIMARY_RED)
            kwargs.setdefault("hover_color", CTK_PRIMARY_RED_HOVER)
            kwargs.setdefault("text_color", "#ffffff")
            kwargs.setdefault("border_width", 0)
            return ctk.CTkButton(parent, **kwargs)
        orange_btn = {"fg_color": "#f59e0b", "hover_color": "#d97706"}
    else:
        BtnCls = ttk.Button
        orange_btn = {}
    FrameCls = ctk.CTkFrame if self.orc_use_custom else ttk.Frame
    LabelCls = ctk.CTkLabel if self.orc_use_custom else ttk.Label
    EntryCls = ctk.CTkEntry if self.orc_use_custom else ttk.Entry

    top = (
        FrameCls(
            self.tab_orc,
            fg_color="#ffffff",
            corner_radius=10,
            border_width=1,
            border_color="#e7cfd3",
        )
        if self.orc_use_custom
        else FrameCls(self.tab_orc)
    )
    top.pack(fill="x", padx=10, pady=(10, 4))
    if self.orc_use_custom:
        self.orc_logo_img = None
        try:
            logo_path = get_orc_logo_path()
            if logo_path and os.path.exists(logo_path):
                from PIL import Image, ImageTk
                img = Image.open(logo_path).convert("RGBA")
                rw = 90
                rh = max(26, int(img.height * (rw / max(1, img.width))))
                img = img.resize((rw, rh))
                self.orc_logo_img = ImageTk.PhotoImage(img)
        except Exception:
            self.orc_logo_img = None
        if self.orc_logo_img:
            ctk.CTkLabel(top, image=self.orc_logo_img, text="").pack(side="left", padx=(6, 8), pady=2)
    BtnCls(top, text="Novo Orçamento", command=self.add_orcamento, **orange_btn).pack(side="left", padx=4, pady=2)
    BtnCls(top, text="Remover", command=self.remove_orcamento).pack(side="left", padx=4, pady=2)
    BtnCls(top, text="Em Edição", command=lambda: self.set_orc_estado("Em edição")).pack(side="left", padx=4, pady=2)
    BtnCls(top, text="Enviado", command=lambda: self.set_orc_estado("Enviado")).pack(side="left", padx=4, pady=2)
    BtnCls(top, text="Aprovado", command=lambda: self.set_orc_estado("Aprovado")).pack(side="left", padx=4, pady=2)
    BtnCls(top, text="Rejeitado", command=lambda: self.set_orc_estado("Rejeitado")).pack(side="left", padx=4, pady=2)

    top2 = (
        FrameCls(
            self.tab_orc,
            fg_color="#ffffff",
            corner_radius=10,
            border_width=1,
            border_color="#e7cfd3",
        )
        if self.orc_use_custom
        else FrameCls(self.tab_orc)
    )
    top2.pack(fill="x", padx=10, pady=(0, 8))
    BtnCls(top2, text="Converter em Encomenda", command=self.convert_orc_to_encomenda, **orange_btn).pack(side="left", padx=4, pady=2)
    BtnCls(top2, text="Pré-visualizar", command=self.preview_orcamento).pack(side="left", padx=4, pady=2)
    BtnCls(top2, text="Guardar PDF", command=self.save_orc_pdf).pack(side="left", padx=4, pady=2)
    BtnCls(top2, text="Imprimir PDF", command=self.print_orc_pdf).pack(side="left", padx=4, pady=2)
    BtnCls(top2, text="Abrir com...", command=self.open_orc_pdf_with).pack(side="left", padx=4, pady=2)

    self.orc_search = StringVar(value="")
    self.orc_state_filter = StringVar(value="Ativas")
    self.orc_year_filter = StringVar(value=str(datetime.now().year))
    top3 = (
        FrameCls(
            self.tab_orc,
            fg_color="#ffffff",
            corner_radius=10,
            border_width=1,
            border_color="#e7cfd3",
        )
        if self.orc_use_custom
        else FrameCls(self.tab_orc)
    )
    top3.pack(fill="x", padx=10, pady=(0, 8))
    LabelCls(top3, text="Filtro").pack(side="left", padx=(0, 6), pady=2)
    self.orc_filter_entry = EntryCls(top3, textvariable=self.orc_search, width=260 if self.orc_use_custom else 36)
    self.orc_filter_entry.pack(side="left", padx=(0, 8), pady=2)
    try:
        self.orc_filter_entry.bind("<Return>", lambda _e=None: self.refresh_orc_list())
    except Exception:
        pass
    if self.orc_use_custom:
        self.orc_state_segment = ctk.CTkSegmentedButton(
            top3,
            values=["Ativas", "Todos", "Em edição", "Enviado", "Aprovado", "Rejeitado", "Convertido"],
            command=self._on_orc_filter_click,
            width=520,
        )
        self.orc_state_segment.pack(side="left", padx=(0, 8), pady=2)
        self.orc_state_segment.set("Ativas")
    else:
        self.orc_state_cb = ttk.Combobox(
            top3,
            textvariable=self.orc_state_filter,
            values=["Ativas", "Todos", "Em edição", "Enviado", "Aprovado", "Rejeitado", "Convertido"],
            width=20,
            state="readonly",
        )
        self.orc_state_cb.pack(side="left", padx=(0, 8), pady=2)
        self.orc_state_cb.bind("<<ComboboxSelected>>", lambda _e=None: self.refresh_orc_list())
    LabelCls(top3, text="Ano").pack(side="left", padx=(6, 4), pady=2)
    if self.orc_use_custom:
        self.orc_year_cb = ctk.CTkComboBox(
            top3,
            variable=self.orc_year_filter,
            values=[str(datetime.now().year), "Todos"],
            width=120,
            command=self._on_orc_year_change,
        )
        self.orc_year_cb.pack(side="left", padx=(0, 8), pady=2)
    else:
        self.orc_year_cb = ttk.Combobox(
            top3,
            textvariable=self.orc_year_filter,
            values=[str(datetime.now().year), "Todos"],
            width=10,
            state="readonly",
        )
        self.orc_year_cb.pack(side="left", padx=(0, 8), pady=2)
        self.orc_year_cb.bind("<<ComboboxSelected>>", lambda _e=None: self._on_orc_year_change(self.orc_year_filter.get()))
    BtnCls(top3, text="Aplicar", command=self.refresh_orc_list).pack(side="left", padx=4, pady=2)

    main = FrameCls(self.tab_orc, fg_color="#ffffff") if self.orc_use_custom else ttk.Frame(self.tab_orc)
    main.pack(fill="both", expand=True, padx=10, pady=6)
    left = FrameCls(main, fg_color="white", corner_radius=10) if self.orc_use_custom else ttk.Frame(main)
    right = FrameCls(main, fg_color="white", corner_radius=10) if self.orc_use_custom else ttk.Frame(main)
    left.pack(side="left", fill="y", padx=(0, 8), pady=2)
    right.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=2)

    orc_style = ""
    if self.orc_use_custom:
        style = ttk.Style()
        style.configure(
            "Orc.Treeview",
            font=("Segoe UI", 10),
            rowheight=26,
            background="#f8fafc",
            fieldbackground="#f8fafc",
            borderwidth=0,
        )
        style.configure(
            "Orc.Treeview.Heading",
            font=("Segoe UI", 10, "bold"),
            padding=(6, 5),
            background=THEME_HEADER_BG,
            foreground="white",
            relief="flat",
        )
        style.map("Orc.Treeview.Heading", background=[("active", THEME_HEADER_ACTIVE)])
        style.map(
            "Orc.Treeview",
            background=[("selected", THEME_SELECT_BG)],
            foreground=[("selected", THEME_SELECT_FG)],
        )
        style.configure(
            "OrcLin.Treeview",
            font=("Segoe UI", 10),
            rowheight=25,
            background="#f8fbff",
            fieldbackground="#f8fbff",
            borderwidth=0,
        )
        style.configure(
            "OrcLin.Treeview.Heading",
            font=("Segoe UI", 10, "bold"),
            background=THEME_HEADER_BG,
            foreground="white",
            relief="flat",
        )
        style.map("OrcLin.Treeview.Heading", background=[("active", THEME_HEADER_ACTIVE)])
        style.map(
            "OrcLin.Treeview",
            background=[("selected", THEME_SELECT_BG)],
            foreground=[("selected", THEME_SELECT_FG)],
        )
        orc_style = "Orc.Treeview"

    if self.orc_use_custom:
        ctk.CTkLabel(
            left,
            text="Orçamentos",
            font=("Segoe UI", 12, "bold"),
            text_color="#7a0f1a",
        ).pack(anchor="w", padx=8, pady=(8, 2))
    self.orc_tbl = ttk.Treeview(
        left,
        columns=("numero", "cliente", "estado", "encomenda", "total"),
        show="headings",
        height=11,
        style=orc_style,
        selectmode="browse",
    )
    self.orc_tbl.heading("numero", text="Número")
    self.orc_tbl.heading("cliente", text="Cliente")
    self.orc_tbl.heading("estado", text="Estado")
    self.orc_tbl.heading("encomenda", text="Encomenda")
    self.orc_tbl.heading("total", text="Total")
    self.orc_tbl.column("numero", width=140)
    self.orc_tbl.column("cliente", width=180)
    self.orc_tbl.column("estado", width=110)
    self.orc_tbl.column("encomenda", width=130)
    self.orc_tbl.column("total", width=90)
    self.orc_tbl.pack(side="left", fill="y")
    if self.orc_use_custom and CUSTOM_TK_AVAILABLE:
        sc = ctk.CTkScrollbar(left, orientation="vertical", command=self.orc_tbl.yview)
    else:
        sc = ttk.Scrollbar(left, orient="vertical", command=self.orc_tbl.yview)
    self.orc_tbl.configure(yscrollcommand=sc.set)
    sc.pack(side="left", fill="y")
    self.orc_tbl.bind("<<TreeviewSelect>>", self.on_orc_select)
    self.bind_clear_on_empty(self.orc_tbl, self.clear_orc_details)

    if self.orc_use_custom:
        cliente = ctk.CTkFrame(
            right,
            fg_color="#ffffff",
            corner_radius=10,
            border_width=1,
            border_color="#e7cfd3",
        )
        ctk.CTkLabel(
            cliente,
            text="Cliente (selecionar)",
            font=("Segoe UI", 12, "bold"),
            text_color="#7a0f1a",
        ).grid(row=0, column=0, columnspan=4, sticky="w", padx=10, pady=(8, 4))
        base_row = 1
    else:
        cliente = ttk.LabelFrame(right, text="Cliente (selecionar)")
        base_row = 0
    cliente.pack(fill="x")
    self.orc_cliente_vars = {
        "codigo": StringVar(),
        "nome": StringVar(),
        "empresa": StringVar(),
        "nif": StringVar(),
        "morada": StringVar(),
        "contacto": StringVar(),
        "email": StringVar(),
    }
    self.orc_executado = StringVar()
    if self.orc_use_custom:
        ctk.CTkLabel(cliente, text="Cliente").grid(row=base_row + 0, column=0, sticky="w", padx=8, pady=4)
        self.orc_cliente_cb = ctk.CTkComboBox(
            cliente,
            variable=self.orc_cliente_vars["codigo"],
            values=self.get_clientes_display() or [""],
            width=260,
            state="readonly",
        )
        self.orc_cliente_cb.grid(row=base_row + 0, column=1, padx=8, pady=2, sticky="w")
        BtnCls(cliente, text="Selecionar", command=self.fill_orc_from_cliente, width=120, **orange_btn).grid(row=base_row + 0, column=2, padx=8, pady=2)

        ctk.CTkLabel(cliente, text="Orçamentista").grid(row=base_row + 1, column=0, sticky="w", padx=8, pady=4)
        self.orc_exec_cb = ctk.CTkComboBox(
            cliente,
            variable=self.orc_executado,
            values=self.data.get("orcamentistas", []) or [""],
            width=260,
            state="normal",
        )
        self.orc_exec_cb.grid(row=base_row + 1, column=1, padx=8, pady=2, sticky="w")
        BtnCls(cliente, text="Gerir", command=self.gerir_orcamentistas, width=120).grid(row=base_row + 1, column=2, padx=8, pady=2)

        ctk.CTkLabel(cliente, text="Nome").grid(row=base_row + 2, column=0, sticky="w", padx=8, pady=2)
        ctk.CTkEntry(cliente, textvariable=self.orc_cliente_vars["nome"], width=270).grid(row=base_row + 2, column=1, padx=8, pady=2, sticky="w")
        ctk.CTkLabel(cliente, text="Empresa").grid(row=base_row + 2, column=2, sticky="w", padx=8, pady=2)
        ctk.CTkEntry(cliente, textvariable=self.orc_cliente_vars["empresa"], width=260).grid(row=base_row + 2, column=3, padx=8, pady=2, sticky="w")

        ctk.CTkLabel(cliente, text="NIF").grid(row=base_row + 3, column=0, sticky="w", padx=8, pady=2)
        ctk.CTkEntry(cliente, textvariable=self.orc_cliente_vars["nif"], width=170).grid(row=base_row + 3, column=1, padx=8, pady=2, sticky="w")
        ctk.CTkLabel(cliente, text="Morada").grid(row=base_row + 3, column=2, sticky="w", padx=8, pady=2)
        ctk.CTkEntry(cliente, textvariable=self.orc_cliente_vars["morada"], width=260).grid(row=base_row + 3, column=3, padx=8, pady=2, sticky="w")

        ctk.CTkLabel(cliente, text="Contacto").grid(row=base_row + 4, column=0, sticky="w", padx=8, pady=2)
        ctk.CTkEntry(cliente, textvariable=self.orc_cliente_vars["contacto"], width=170).grid(row=base_row + 4, column=1, padx=8, pady=2, sticky="w")
        ctk.CTkLabel(cliente, text="Email").grid(row=base_row + 4, column=2, sticky="w", padx=8, pady=2)
        ctk.CTkEntry(cliente, textvariable=self.orc_cliente_vars["email"], width=260).grid(row=base_row + 4, column=3, padx=8, pady=2, sticky="w")
    else:
        ttk.Label(cliente, text="Cliente").grid(row=base_row + 0, column=0, sticky="w", padx=6, pady=4)
        self.orc_cliente_cb = ttk.Combobox(cliente, textvariable=self.orc_cliente_vars["codigo"], values=self.get_clientes_display(), width=22, state="readonly")
        self.orc_cliente_cb.grid(row=base_row + 0, column=1, padx=6, pady=4, sticky="w")
        BtnCls(cliente, text="Selecionar", command=self.fill_orc_from_cliente, **orange_btn).grid(row=base_row + 0, column=2, padx=6, pady=4)
        ttk.Label(cliente, text="Orçamentista").grid(row=base_row + 1, column=0, sticky="w", padx=6, pady=4)
        self.orc_exec_cb = ttk.Combobox(cliente, textvariable=self.orc_executado, values=self.data.get("orcamentistas", []), width=22)
        self.orc_exec_cb.grid(row=base_row + 1, column=1, padx=6, pady=4, sticky="w")
        BtnCls(cliente, text="Gerir", command=self.gerir_orcamentistas).grid(row=base_row + 1, column=2, padx=6, pady=4)
        ttk.Label(cliente, text="Nome").grid(row=base_row + 2, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(cliente, textvariable=self.orc_cliente_vars["nome"], width=30, state="readonly").grid(row=base_row + 2, column=1, padx=6, pady=4, sticky="w")
        ttk.Label(cliente, text="Empresa").grid(row=base_row + 2, column=2, sticky="w", padx=6, pady=4)
        ttk.Entry(cliente, textvariable=self.orc_cliente_vars["empresa"], width=28, state="readonly").grid(row=base_row + 2, column=3, padx=6, pady=4, sticky="w")
        ttk.Label(cliente, text="NIF").grid(row=base_row + 3, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(cliente, textvariable=self.orc_cliente_vars["nif"], width=18, state="readonly").grid(row=base_row + 3, column=1, padx=6, pady=4, sticky="w")
        ttk.Label(cliente, text="Morada").grid(row=base_row + 3, column=2, sticky="w", padx=6, pady=4)
        ttk.Entry(cliente, textvariable=self.orc_cliente_vars["morada"], width=28, state="readonly").grid(row=base_row + 3, column=3, padx=6, pady=4, sticky="w")
        ttk.Label(cliente, text="Contacto").grid(row=base_row + 4, column=0, sticky="w", padx=6, pady=4)
        ttk.Entry(cliente, textvariable=self.orc_cliente_vars["contacto"], width=18, state="readonly").grid(row=base_row + 4, column=1, padx=6, pady=4, sticky="w")
        ttk.Label(cliente, text="Email").grid(row=base_row + 4, column=2, sticky="w", padx=6, pady=4)
        ttk.Entry(cliente, textvariable=self.orc_cliente_vars["email"], width=28, state="readonly").grid(row=base_row + 4, column=3, padx=6, pady=4, sticky="w")

    if self.orc_use_custom:
        linhas_box = ctk.CTkFrame(
            right,
            fg_color="#ffffff",
            corner_radius=10,
            border_width=1,
            border_color="#e7cfd3",
        )
        ctk.CTkLabel(
            linhas_box,
            text="Referências",
            font=("Segoe UI", 12, "bold"),
            text_color="#7a0f1a",
        ).pack(anchor="w", padx=8, pady=(8, 2))
    else:
        linhas_box = ttk.LabelFrame(right, text="Referências")
    # manter altura controlada para garantir visibilidade da caixa de notas
    linhas_box.pack(fill="x", pady=(4, 0))
    btns = ctk.CTkFrame(linhas_box, fg_color="#ffffff") if self.orc_use_custom else ttk.Frame(linhas_box)
    btns.pack(fill="x", padx=6, pady=2)
    BtnCls(btns, text="Adicionar Linha", command=self.add_orc_linha).pack(side="left", padx=4, pady=2)
    BtnCls(btns, text="Editar", command=self.edit_orc_linha).pack(side="left", padx=4, pady=2)
    BtnCls(btns, text="Remover", command=self.remove_orc_linha).pack(side="left", padx=4, pady=2)
    BtnCls(btns, text="Ver Desenho", command=self.open_orc_linha_desenho).pack(side="left", padx=4, pady=2)
    BtnCls(btns, text="Guardar Dados", command=self.save_orc_fields, **orange_btn).pack(side="left", padx=8, pady=2)

    linhas_tbl_wrap = ctk.CTkFrame(linhas_box, fg_color="#ffffff") if self.orc_use_custom else ttk.Frame(linhas_box)
    linhas_tbl_wrap.pack(fill="x", padx=6, pady=(0, 2))
    self.orc_linhas = ttk.Treeview(
        linhas_tbl_wrap,
        columns=("ref_int", "ref_ext", "descricao", "material", "espessura", "operacao", "qtd", "preco", "total"),
        show="headings",
        height=3,
        style=("OrcLin.Treeview" if self.orc_use_custom else ""),
    )
    for col, txt, w, anchor in [
        ("ref_int", "Ref. Interna", 120, "w"),
        ("ref_ext", "Ref. Externa", 220, "w"),
        ("descricao", "Descrição", 140, "w"),
        ("material", "Material", 90, "center"),
        ("espessura", "Espessura", 70, "center"),
        ("operacao", "Operação", 110, "center"),
        ("qtd", "Qtd", 60, "center"),
        ("preco", "Preço Unit.", 80, "e"),
        ("total", "Total", 80, "e"),
    ]:
        self.orc_linhas.heading(col, text=txt)
        self.orc_linhas.column(col, width=w, anchor=anchor)
    self.orc_linhas.grid(row=0, column=0, sticky="ew")
    if self.orc_use_custom and CUSTOM_TK_AVAILABLE:
        sc2 = ctk.CTkScrollbar(linhas_tbl_wrap, orientation="vertical", command=self.orc_linhas.yview)
    else:
        sc2 = ttk.Scrollbar(linhas_tbl_wrap, orient="vertical", command=self.orc_linhas.yview)
    self.orc_linhas.configure(yscrollcommand=sc2.set)
    sc2.grid(row=0, column=1, sticky="ns")
    linhas_tbl_wrap.grid_columnconfigure(0, weight=1)

    totals = ctk.CTkFrame(
        right,
        fg_color="white",
        corner_radius=10,
        border_width=1,
        border_color="#e7cfd3",
    ) if self.orc_use_custom else ttk.Frame(right)
    totals.pack(fill="x", pady=(2, 0))
    self.orc_iva = DoubleVar(value=23.0)
    self.orc_subtotal = StringVar(value="0.00")
    self.orc_total = StringVar(value="0.00")
    if self.orc_use_custom:
        ctk.CTkLabel(totals, text="IVA (%)", font=("Segoe UI", 12, "bold"), text_color="#6f1624").pack(side="left", padx=(10, 6), pady=6)
        ctk.CTkEntry(totals, textvariable=self.orc_iva, width=90, font=("Segoe UI", 12)).pack(side="left", pady=6)
        ctk.CTkLabel(totals, text="Subtotal", font=("Segoe UI", 12, "bold"), text_color="#6f1624").pack(side="left", padx=(20, 6), pady=6)
        ctk.CTkLabel(totals, textvariable=self.orc_subtotal, font=("Segoe UI", 12)).pack(side="left", pady=6)
        ctk.CTkLabel(totals, text="Total", font=("Segoe UI", 13, "bold"), text_color="#7a0f1a").pack(side="left", padx=(20, 6), pady=6)
        ctk.CTkLabel(totals, textvariable=self.orc_total, font=("Segoe UI", 13, "bold"), text_color="#7a0f1a").pack(side="left", pady=6)
    else:
        ttk.Label(totals, text="IVA (%)").pack(side="left", padx=6)
        ttk.Entry(totals, textvariable=self.orc_iva, width=6).pack(side="left")
        ttk.Label(totals, text="Subtotal").pack(side="left", padx=(14, 6))
        ttk.Label(totals, textvariable=self.orc_subtotal).pack(side="left")
        ttk.Label(totals, text="Total").pack(side="left", padx=(14, 6))
        ttk.Label(totals, textvariable=self.orc_total).pack(side="left")

    self.orc_nota_transporte = StringVar(value="")
    if self.orc_use_custom and CUSTOM_TK_AVAILABLE:
        notas_box = ctk.CTkFrame(
            right,
            fg_color="#ffffff",
            corner_radius=10,
            border_width=1,
            border_color="#dfc4c9",
        )
        notas_box.pack(fill="x", pady=(3, 0))
        ctk.CTkLabel(
            notas_box,
            text="Notas do orçamento (PDF)",
            font=("Segoe UI", 13, "bold"),
            text_color="#7a0f1a",
        ).grid(row=0, column=0, columnspan=4, sticky="w", padx=10, pady=(6, 3))
        ctk.CTkLabel(
            notas_box,
            text="Transporte",
            font=("Segoe UI", 11, "bold"),
            text_color="#6f1624",
        ).grid(row=1, column=0, sticky="w", padx=10, pady=4)
        self.orc_transporte_cb = ctk.CTkComboBox(
            notas_box,
            variable=self.orc_nota_transporte,
            values=["", "Transporte a Cargo do Cliente", "Transporte a Nosso Cargo"],
            width=280,
            font=("Segoe UI", 11),
            fg_color="#f8fbff",
            border_color="#d8b9bf",
            button_color=CTK_PRIMARY_RED_HOVER,
            dropdown_fg_color="#ffffff",
        )
        self.orc_transporte_cb.grid(row=1, column=1, sticky="w", padx=6, pady=4)
        ctk.CTkButton(
            notas_box,
            text="Notas por Operações",
            command=self.orc_fill_notes_by_ops,
            width=170,
            fg_color=CTK_PRIMARY_RED_HOVER,
            hover_color="#8f1d14",
        ).grid(row=1, column=2, padx=6, pady=4, sticky="w")

        ops_btns = ctk.CTkFrame(notas_box, fg_color="transparent")
        ops_btns.grid(row=2, column=0, columnspan=4, sticky="w", padx=10, pady=(0, 4))
        for txt in ("Corte Laser", "Quinagem", "Roscagem", "Furo Manual", "Soldadura"):
            ctk.CTkButton(
                ops_btns,
                text=txt,
                command=lambda t=txt: self._orc_append_pdf_note(f"- Foi considerado: {t}."),
                width=122,
                height=30,
                fg_color=CTK_PRIMARY_RED,
                hover_color=CTK_PRIMARY_RED_HOVER,
            ).pack(side="left", padx=(0, 6), pady=2)

        self.orc_notas_text = ctk.CTkTextbox(
            notas_box,
            height=54,
            font=("Segoe UI", 11),
            fg_color="#f8fbff",
            border_color="#d8b9bf",
            border_width=1,
        )
        self.orc_notas_text.grid(row=3, column=0, columnspan=4, padx=10, pady=(3, 8), sticky="nsew")
        notas_box.grid_columnconfigure(1, weight=1)
        notas_box.grid_columnconfigure(2, weight=1)
        notas_box.grid_columnconfigure(3, weight=1)
    else:
        notas_box = ttk.LabelFrame(right, text="Notas do orçamento (PDF)")
        notas_box.pack(fill="x", pady=(4, 0))
        ttk.Label(notas_box, text="Transporte").grid(row=0, column=0, sticky="w", padx=6, pady=4)
        self.orc_transporte_cb = ttk.Combobox(
            notas_box,
            textvariable=self.orc_nota_transporte,
            values=("", "Transporte a Cargo do Cliente", "Transporte a Nosso Cargo"),
            width=34,
            state="readonly",
        )
        self.orc_transporte_cb.grid(row=0, column=1, padx=6, pady=4, sticky="w")
        ttk.Button(notas_box, text="Notas por Operações", command=self.orc_fill_notes_by_ops).grid(row=0, column=2, padx=6, pady=4)

        ops_btns = ttk.Frame(notas_box)
        ops_btns.grid(row=1, column=0, columnspan=3, sticky="w", padx=6, pady=(0, 4))
        for txt in ("Corte Laser", "Quinagem", "Roscagem", "Furo Manual", "Soldadura"):
            ttk.Button(ops_btns, text=txt, command=lambda t=txt: self._orc_append_pdf_note(f"- Foi considerado: {t}.")).pack(side="left", padx=(0, 4))

        self.orc_notas_text = Text(notas_box, width=110, height=4)
        self.orc_notas_text.grid(row=2, column=0, columnspan=3, padx=6, pady=(2, 6), sticky="we")

    self.refresh_orc_year_options(keep_selection=False)
    self._on_orc_year_change(self.orc_year_filter.get())

def build_operador(self):
    _ensure_configured()
    style = ttk.Style()
    style.configure("Operador.TFrame", background="#fff8f9")
    style.configure("Operador.TLabelframe", background="#fff8f9")
    style.configure("Operador.TLabelframe.Label", background="#fff8f9")
    style.configure("Operador.TLabel", background="#fff8f9")
    style.configure("Operador.Treeview", background="white", fieldbackground="white")

    self.tab_operador.configure(style="Operador.TFrame")

    use_custom_op = CUSTOM_TK_AVAILABLE and os.environ.get("USE_CUSTOM_OP", "1") != "0"
    self.op_use_custom = use_custom_op
    # no modo custom, o nível 1 do Operador passa para lista simplificada
    self.op_use_full_custom = bool(use_custom_op)
    if use_custom_op:
        try:
            style.theme_use("clam")
            style.configure("Operador.Treeview", background="#f8fbff", fieldbackground="#f8fbff", rowheight=24, borderwidth=0)
            style.configure(
                "Operador.Treeview.Heading",
                background="#1f4e78",
                foreground="white",
                font=("Segoe UI", 9, "bold"),
                relief="flat",
            )
            style.map("Operador.Treeview.Heading", background=[("active", THEME_HEADER_ACTIVE)])
            style.map(
                "Operador.Treeview",
                background=[("selected", "#2f80ed")],
                foreground=[("selected", "#ffffff")],
            )
        except Exception:
            pass

    container = self.tab_operador
    if use_custom_op:
        container = ctk.CTkFrame(self.tab_operador, fg_color="#fff8f9")
        container.pack(fill="both", expand=True)

    def make_label(parent, text, **kwargs):
        if use_custom_op:
            return ctk.CTkLabel(parent, text=text, **kwargs)
        return ttk.Label(parent, text=text, style="Operador.TLabel", **kwargs)

    def make_frame(parent, **kwargs):
        if use_custom_op:
            return ctk.CTkFrame(parent, fg_color="transparent", **kwargs)
        return ttk.Frame(parent, **kwargs)

    if use_custom_op:
        header = ctk.CTkFrame(container, corner_radius=10, fg_color="white", border_width=1, border_color="#e7cfd3")
        header.pack(fill="x", padx=10, pady=(10, 6))
        ctk.CTkLabel(header, text="Operador", font=("Segoe UI", 12, "bold")).grid(row=0, column=0, sticky="w", padx=6, pady=(6, 2))
    else:
        header = ttk.LabelFrame(container, text="Operador", style="Operador.TLabelframe")
        header.pack(fill="x", padx=10, pady=(10, 6))

    self.op_user = StringVar()
    self.op_posto = StringVar()
    self.op_status_filter = StringVar(value="Ativas")
    self.op_year_filter = StringVar(value=str(datetime.now().year))
    self.op_quick_filter = StringVar()
    self.op_user_cb = None
    self.op_manage_btn = None
    op_role = str(self.user.get("role", "") or "") == "Operador"
    op_login_name = str(self.user.get("username", "") or "").strip()
    try:
        operator_accounts = [
            str(u.get("username", "") or "").strip()
            for u in list(self.data.get("users", []) or [])
            if str(u.get("role", "") or "").strip().lower() == "operador" and str(u.get("username", "") or "").strip()
        ]
    except Exception:
        operator_accounts = []
    try:
        existing_ops = [str(o).strip() for o in list(self.data.get("operadores", []) or []) if str(o).strip()]
        merged_ops = list(dict.fromkeys(existing_ops + operator_accounts))
        if merged_ops != existing_ops:
            self.data["operadores"] = merged_ops
            save_data(self.data)
    except Exception:
        pass
    if use_custom_op:
        top_line = ctk.CTkFrame(header, fg_color="transparent")
        top_line.grid(row=1, column=0, columnspan=12, sticky="ew", padx=6, pady=(2, 4))
        top_line.grid_columnconfigure(6 if not op_role else 5, weight=1)

        if op_role:
            ctk.CTkLabel(top_line, text="Sessao").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=3)
            user_badge = ctk.CTkFrame(
                top_line,
                fg_color="#f8fbff",
                border_width=1,
                border_color="#d7e3f5",
                corner_radius=8,
            )
            user_badge.grid(row=0, column=1, sticky="w", padx=(0, 10), pady=3)
            ctk.CTkLabel(
                user_badge,
                text=(op_login_name or "Operador"),
                text_color="#7a0f1a",
                font=("Segoe UI", 12, "bold"),
            ).pack(side="left", padx=10, pady=4)
            ctk.CTkButton(top_line, text="Atualizar", command=self.refresh, width=120).grid(row=0, column=2, padx=0, pady=3, sticky="w")
            ctk.CTkLabel(top_line, text="Posto").grid(row=0, column=3, sticky="e", padx=(12, 6), pady=3)
            self.op_posto_cb = ctk.CTkComboBox(
                top_line,
                variable=self.op_posto,
                values=list(self.data.get("postos_trabalho", []) or ["Geral"]),
                width=170,
            )
            self.op_posto_cb.grid(row=0, column=4, sticky="w", padx=(0, 10), pady=3)
            try:
                self.op_posto_cb.configure(state="disabled")
            except Exception:
                pass
        else:
            ctk.CTkLabel(top_line, text="Nome").grid(row=0, column=0, sticky="w", padx=(0, 6), pady=3)
            self.op_user_cb = ctk.CTkComboBox(top_line, variable=self.op_user, values=self.data.get("operadores", []), width=220)
            self.op_user_cb.grid(row=0, column=1, sticky="w", padx=(0, 10), pady=3)
            self.op_manage_btn = ctk.CTkButton(top_line, text="Gerir Operadores", command=self.gerir_operadores, width=150)
            self.op_manage_btn.grid(row=0, column=2, padx=(0, 10), pady=3)
            ctk.CTkLabel(top_line, text="Posto").grid(row=0, column=3, sticky="e", padx=(12, 6), pady=3)
            self.op_posto_cb = ctk.CTkComboBox(
                top_line,
                variable=self.op_posto,
                values=list(self.data.get("postos_trabalho", []) or ["Geral"]),
                width=170,
            )
            self.op_posto_cb.grid(row=0, column=4, sticky="w", padx=(0, 10), pady=3)
            ctk.CTkButton(top_line, text="Atualizar", command=self.refresh, width=120).grid(row=0, column=5, padx=0, pady=3, sticky="w")

        ctk.CTkLabel(top_line, text="Filtro de Estado").grid(row=1, column=0, sticky="w", padx=(0, 6), pady=3)
        self.op_status_segment = ctk.CTkSegmentedButton(
            top_line,
            values=["Ativas", "Todas", "Preparação", "Em produção", "Concluída"],
            command=self._on_op_status_click,
            width=520,
        )
        self.op_status_segment.grid(row=1, column=1, columnspan=(2 if op_role else 3), sticky="w", padx=(0, 10), pady=3)
        self.op_status_segment.set(self.op_status_filter.get())
        year_col = 5 if op_role else 6
        ctk.CTkLabel(top_line, text="Ano").grid(row=1, column=year_col, sticky="e", padx=(6, 4), pady=3)
        self.op_year_cb = ctk.CTkComboBox(
            top_line,
            variable=self.op_year_filter,
            values=[str(datetime.now().year), "Todos"],
            width=110,
            command=self._on_operador_year_change,
        )
        self.op_year_cb.grid(row=1, column=year_col + 1, sticky="w", padx=(0, 8), pady=3)
        ctk.CTkLabel(top_line, text="Filtro").grid(row=2, column=0, sticky="w", padx=(0, 6), pady=3)
        self.op_filter_entry = ctk.CTkEntry(
            top_line,
            textvariable=self.op_quick_filter,
            width=360,
            placeholder_text="Numero / Cliente / Estado",
        )
        self.op_filter_entry.grid(row=2, column=1, columnspan=(4 if op_role else 5), sticky="w", padx=(0, 10), pady=3)
        try:
            self.op_filter_entry.bind("<KeyRelease>", lambda _e=None: self.debounce_call("operador_filter", self.refresh_operador, delay_ms=140))
        except Exception:
            pass
        ctk.CTkLabel(
            header,
            text="Clique numa encomenda para abrir o detalhe de produção.",
            text_color="#4b5563",
            font=("Segoe UI", 11),
        ).grid(row=3, column=0, columnspan=12, sticky="w", padx=8, pady=(0, 8))
    else:
        ttk.Label(header, text=("Sessão" if op_role else "Nome"), style="Operador.TLabel").grid(row=0, column=0, sticky="w", padx=6, pady=6)
        self.op_user_cb = ttk.Combobox(header, textvariable=self.op_user, values=self.data.get("operadores", []), width=28)
        self.op_user_cb.grid(row=0, column=1, sticky="w", padx=6, pady=6)
        if not op_role:
            self.op_manage_btn = Button(header, text="Gerir Operadores", command=self.gerir_operadores, width=18, **self.btn_cfg)
            self.op_manage_btn.grid(row=0, column=2, padx=6, pady=6)
        ttk.Label(header, text="Posto", style="Operador.TLabel").grid(row=0, column=3, sticky="e", padx=(12, 4), pady=6)
        self.op_posto_cb = ttk.Combobox(
            header,
            textvariable=self.op_posto,
            values=list(self.data.get("postos_trabalho", []) or ["Geral"]),
            width=18,
            state=("disabled" if op_role else "readonly"),
        )
        self.op_posto_cb.grid(row=0, column=4, padx=4, pady=6, sticky="w")
        Button(header, text="Pré-visualizar Produção", command=self.preview_operador_pdf, width=22, **self.btn_cfg).grid(row=0, column=5, padx=6, pady=6)
        Button(header, text="Pré-visualizar Peça", command=self.preview_peca_pdf, width=20, **self.btn_cfg).grid(row=0, column=6, padx=6, pady=6)
        Button(header, text="Iniciar Produção", command=self.operador_iniciar_espessura, width=18, **self.btn_cfg).grid(row=0, column=7, padx=6, pady=6)
        Button(header, text="Fim Peça", command=self.operador_fim_peca, width=12, **self.btn_cfg).grid(row=0, column=8, padx=6, pady=6)
        Button(header, text="Reabrir Peça", command=self.operador_reabrir_peca_total, width=14, **self.btn_cfg).grid(row=0, column=9, padx=6, pady=6)
        Button(header, text="Histórico Rejeitadas", command=self.show_rejeitadas_hist, width=18, **self.btn_cfg).grid(row=0, column=10, padx=6, pady=6)
        Button(header, text="Atualizar", command=self.refresh, width=14, **self.btn_cfg).grid(row=0, column=11, padx=6, pady=6)
        Button(header, text="Imprimir Ordem de Fabrico", command=self.print_of_selected_encomenda, width=24, **self.btn_cfg).grid(row=0, column=12, padx=6, pady=6)
        ttk.Label(header, text="Estado", style="Operador.TLabel").grid(row=0, column=13, sticky="e", padx=(12, 4), pady=6)
        self.op_status_cb = ttk.Combobox(header, textvariable=self.op_status_filter, values=["Ativas", "Todas", "Preparação", "Em produção", "Concluída"], width=16, state="readonly")
        self.op_status_cb.grid(row=0, column=14, padx=4, pady=6, sticky="w")
        self.op_status_cb.bind("<<ComboboxSelected>>", lambda _e=None: self._on_op_status_click(self.op_status_filter.get()))
        ttk.Label(header, text="Ano", style="Operador.TLabel").grid(row=0, column=15, sticky="e", padx=(12, 4), pady=6)
        self.op_year_cb = ttk.Combobox(
            header,
            textvariable=self.op_year_filter,
            values=[str(datetime.now().year), "Todos"],
            width=10,
            state="readonly",
        )
        self.op_year_cb.grid(row=0, column=16, padx=4, pady=6, sticky="w")
        self.op_year_cb.bind("<<ComboboxSelected>>", lambda _e=None: self._on_operador_year_change(self.op_year_filter.get()))
        ttk.Label(header, text="Filtro", style="Operador.TLabel").grid(row=1, column=0, sticky="w", padx=6, pady=(0, 6))
        self.op_filter_entry = ttk.Entry(header, textvariable=self.op_quick_filter, width=36)
        self.op_filter_entry.grid(row=1, column=1, columnspan=3, sticky="w", padx=6, pady=(0, 6))
        self.op_filter_entry.bind("<KeyRelease>", lambda _e=None: self.debounce_call("operador_filter", self.refresh_operador, delay_ms=140))

    make_label(container, "Encomendas", font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=10, pady=(4, 0))
    enc_frame = make_frame(container)
    enc_frame.pack(fill="x", padx=10, pady=(0, 4))
    if self.op_use_full_custom:
        self.op_enc_frame = enc_frame
        top_band = ctk.CTkFrame(enc_frame, fg_color=THEME_HEADER_BG, corner_radius=8)
        top_band.pack(fill="x", pady=(0, 4))
        ctk.CTkLabel(
            top_band,
            text="NÚMERO | CLIENTE | ESTADO",
            text_color="white",
            font=("Segoe UI", 11, "bold"),
            anchor="w",
        ).pack(fill="x", padx=10, pady=7)
        self.op_enc_list = ctk.CTkScrollableFrame(enc_frame, fg_color="white", corner_radius=10, border_width=1, border_color="#e7cfd3", height=560)
        self.op_enc_list.pack(fill="x", expand=True)
    else:
        self.op_tbl_enc = ttk.Treeview(enc_frame, columns=("numero", "cliente", "estado"), show="headings", height=8, style="Operador.Treeview")
        self.op_tbl_enc.heading("numero", text="Número")
        self.op_tbl_enc.heading("cliente", text="Cliente")
        self.op_tbl_enc.heading("estado", text="Estado")
        self.op_tbl_enc.column("numero", width=120)
        self.op_tbl_enc.column("cliente", width=220)
        self.op_tbl_enc.column("estado", width=140)
        self.op_tbl_enc.pack(side="left", fill="x", expand=True)
        if use_custom_op:
            enc_scroll = ctk.CTkScrollbar(enc_frame, orientation="vertical", command=self.op_tbl_enc.yview)
        else:
            enc_scroll = ttk.Scrollbar(enc_frame, orient="vertical", command=self.op_tbl_enc.yview)
        enc_scroll.pack(side="right", fill="y")
        self.op_tbl_enc.configure(yscrollcommand=enc_scroll.set)
        self.op_tbl_enc.bind("<<TreeviewSelect>>", self.on_operador_select)
        self.bind_clear_on_empty(self.op_tbl_enc, self.clear_operador_selection)

    if not self.op_use_full_custom:
        esp_frame = make_frame(container)
        esp_frame.pack(fill="x", padx=10, pady=(6, 6))
        self.op_tbl_esp = ttk.Treeview(
            esp_frame,
            columns=("material", "espessura", "planeado", "produzido", "estado"),
            show="headings",
            height=8,
            style="Operador.Treeview",
        )
        self.op_tbl_esp.heading("material", text="Material")
        self.op_tbl_esp.heading("espessura", text="Espessura")
        self.op_tbl_esp.heading("planeado", text="Planeado")
        self.op_tbl_esp.heading("produzido", text="Produzido")
        self.op_tbl_esp.heading("estado", text="Estado")
        self.op_tbl_esp.column("material", width=140)
        self.op_tbl_esp.column("espessura", width=100)
        self.op_tbl_esp.column("planeado", width=90)
        self.op_tbl_esp.column("produzido", width=90)
        self.op_tbl_esp.column("estado", width=120)
        self.op_tbl_esp.pack(side="left", fill="x", expand=True)
        if use_custom_op:
            esp_scroll = ctk.CTkScrollbar(esp_frame, orientation="vertical", command=self.op_tbl_esp.yview)
        else:
            esp_scroll = ttk.Scrollbar(esp_frame, orient="vertical", command=self.op_tbl_esp.yview)
        esp_scroll.pack(side="right", fill="y")
        self.op_tbl_esp.configure(yscrollcommand=esp_scroll.set)
        self.op_tbl_esp.bind("<<TreeviewSelect>>", self.on_operador_select_espessura)
        self.bind_clear_on_empty(self.op_tbl_esp, self.clear_operador_esp)
        self.op_tbl_esp.tag_configure("op_em_curso", background="#ffd8a8")
        self.op_tbl_esp.tag_configure("op_concluida", background="#b7e4c7")
        self.op_tbl_esp.tag_configure("op_even", background="#fff5f6")
        self.op_tbl_esp.tag_configure("op_odd", background="#fbecee")

        pecas_frame = make_frame(container)
        pecas_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.op_tbl_pecas = ttk.Treeview(
            pecas_frame,
            columns=("ref_int", "ref_ext", "ops", "em_curso", "qtd", "produzido", "estado"),
            show="headings",
            height=12,
            style="Operador.Treeview",
        )
        self.op_tbl_pecas.heading("ref_int", text="Ref. Interna")
        self.op_tbl_pecas.heading("ref_ext", text="Ref. Externa")
        self.op_tbl_pecas.heading("ops", text="Operações Pendentes")
        self.op_tbl_pecas.heading("em_curso", text="Em curso / Operador")
        self.op_tbl_pecas.heading("qtd", text="Planeado")
        self.op_tbl_pecas.heading("produzido", text="Produzido")
        self.op_tbl_pecas.heading("estado", text="Estado")
        self.op_tbl_pecas.column("ref_int", width=160)
        self.op_tbl_pecas.column("ref_ext", width=200)
        self.op_tbl_pecas.column("ops", width=260)
        self.op_tbl_pecas.column("em_curso", width=260)
        self.op_tbl_pecas.column("qtd", width=100)
        self.op_tbl_pecas.column("produzido", width=100)
        self.op_tbl_pecas.column("estado", width=120)
        self.op_tbl_pecas.pack(side="left", fill="both", expand=True)
        pecas_scroll_y = ttk.Scrollbar(pecas_frame, orient="vertical", command=self.op_tbl_pecas.yview)
        pecas_scroll_y.pack(side="right", fill="y")
        pecas_scroll_x = ttk.Scrollbar(pecas_frame, orient="horizontal", command=self.op_tbl_pecas.xview)
        pecas_scroll_x.pack(side="bottom", fill="x")
        self.op_tbl_pecas.configure(yscrollcommand=pecas_scroll_y.set, xscrollcommand=pecas_scroll_x.set)
        self.bind_clear_on_empty(self.op_tbl_pecas, lambda: None)

    # botões movidos para o topo junto ao operador

    self.op_sel_enc_num = None
    self.op_sel_esp = None
    self.op_sel_peca = None

    op_login_name = str(self.user.get("username", "") or "").strip()
    if str(self.user.get("role", "") or "") == "Operador":
        if op_login_name:
            existing_ops = [str(o).strip() for o in self.data.get("operadores", []) if str(o).strip()]
            existing_ops_norm = {name.casefold() for name in existing_ops}
            if op_login_name.casefold() not in existing_ops_norm:
                existing_ops.append(op_login_name)
                self.data["operadores"] = existing_ops
                try:
                    save_data(self.data)
                except Exception:
                    pass
            if self.op_user_cb is not None:
                try:
                    self.op_user_cb.configure(values=self.data.get("operadores", []))
                except Exception:
                    pass
            self.op_user.set(op_login_name)
        if self.op_user_cb is not None:
            try:
                self.op_user_cb.configure(state="disabled")
            except Exception:
                try:
                    self.op_user_cb.configure(state="readonly")
                except Exception:
                    pass
        try:
            if hasattr(self, "op_manage_btn") and self.op_manage_btn is not None:
                self.op_manage_btn.configure(state="disabled")
        except Exception:
            pass
        if getattr(self, "op_posto_cb", None) is not None:
            try:
                self.op_posto_cb.configure(state="disabled")
            except Exception:
                pass

    # Posto de trabalho por operador.
    try:
        if not isinstance(self.data.get("postos_trabalho", None), list) or not self.data.get("postos_trabalho"):
            self.data["postos_trabalho"] = ["Geral", "Laser 1", "Quinagem 1", "Soldadura 1", "Embalamento 1"]
        if not isinstance(self.data.get("tempos_operacao_planeada_min", None), dict):
            self.data["tempos_operacao_planeada_min"] = {
                "laser": 0,
                "quinagem": 0,
                "soldadura": 0,
                "roscagem": 0,
                "embalamento": 0,
            }
        if not isinstance(self.data.get("operador_posto_map", None), dict):
            self.data["operador_posto_map"] = {}
        current_user = str(self.op_user.get() or "").strip()
        if op_role and op_login_name:
            current_user = op_login_name
        default_posto = str(self.data.get("operador_posto_map", {}).get(current_user, "") or "")
        if not default_posto:
            default_posto = str((self.data.get("postos_trabalho", []) or ["Geral"])[0])
        self.op_posto.set(default_posto)
        if hasattr(self, "op_posto_cb") and self.op_posto_cb is not None:
            try:
                self.op_posto_cb.configure(values=list(self.data.get("postos_trabalho", []) or ["Geral"]))
            except Exception:
                pass
        if hasattr(self, "_on_operador_user_change"):
            try:
                self.op_user.trace_add("write", lambda *_: self._on_operador_user_change())
            except Exception:
                pass
        if hasattr(self, "_on_operador_posto_change"):
            try:
                self.op_posto.trace_add("write", lambda *_: self._on_operador_posto_change())
            except Exception:
                pass
    except Exception:
        pass

    self.refresh_operador_year_options(keep_selection=False)
    self._on_operador_year_change(self.op_year_filter.get())
    self.op_blink_on = False
    if not getattr(self, "_op_blink_started", False):
        self._op_blink_started = True
        self.root.after(650, self.op_blink_schedule)

def build_ordens(self):
    _ensure_configured()
    self.ordens_use_custom = CUSTOM_TK_AVAILABLE and os.environ.get("USE_CUSTOM_OF", "1") != "0"
    FrameCls = ctk.CTkFrame if self.ordens_use_custom else ttk.Frame
    BtnCls = ctk.CTkButton if self.ordens_use_custom else ttk.Button
    EntryCls = ctk.CTkEntry if self.ordens_use_custom else ttk.Entry

    self.ordens_sel_opp = ""
    self._ordens_btns = {}

    top = FrameCls(self.tab_ordens, fg_color="#f7f8fb") if self.ordens_use_custom else FrameCls(self.tab_ordens)
    top.pack(fill="x", padx=10, pady=10)
    (ctk.CTkLabel if self.ordens_use_custom else ttk.Label)(
        top,
        text="Pesquisar Ordem de Fabrico",
        font=("Segoe UI", 13, "bold") if self.ordens_use_custom else None,
        text_color="#7f1b2c" if self.ordens_use_custom else None,
    ).pack(side="left")
    self.opp_search = StringVar()
    self.ordens_status_filter = StringVar(value="Ativas")
    self.ordens_year_filter = StringVar(value=str(datetime.now().year))
    entry = EntryCls(top, textvariable=self.opp_search, width=260 if self.ordens_use_custom else 30)
    entry.pack(side="left", padx=6, pady=2)
    try:
        entry.bind("<Return>", lambda _e=None: self.refresh_ordens())
    except Exception:
        pass
    if self.ordens_use_custom:
        ctk.CTkLabel(top, text="Estado", text_color="#7f1b2c").pack(side="left", padx=(10, 4))
        self.ordens_status_segment = ctk.CTkSegmentedButton(
            top,
            values=["Ativas", "Todas", "Preparação", "Em produção", "Concluída"],
            command=lambda v=None: (
                None
                if getattr(self, "_suppress_ordens_filter_cb", False)
                else (self.ordens_status_filter.set(v or "Ativas"), self.refresh_ordens())
            ),
            width=520,
        )
        self.ordens_status_segment.pack(side="left", padx=4, pady=2)
        self.ordens_status_segment.set("Ativas")
    else:
        ttk.Label(top, text="Estado").pack(side="left", padx=(10, 4))
        self.ordens_status_cb = ttk.Combobox(
            top,
            textvariable=self.ordens_status_filter,
            values=["Ativas", "Todas", "Preparação", "Em produção", "Concluída"],
            width=16,
            state="readonly",
        )
        self.ordens_status_cb.pack(side="left", padx=4, pady=2)
        self.ordens_status_cb.bind("<<ComboboxSelected>>", lambda _e=None: self.refresh_ordens())
    (ctk.CTkLabel if self.ordens_use_custom else ttk.Label)(top, text="Ano").pack(side="left", padx=(10, 4))
    if self.ordens_use_custom:
        self.ordens_year_cb = ctk.CTkComboBox(
            top,
            variable=self.ordens_year_filter,
            values=[str(datetime.now().year), "Todos"],
            width=120,
            command=self._on_ordens_year_change,
        )
        self.ordens_year_cb.pack(side="left", padx=4, pady=2)
    else:
        self.ordens_year_cb = ttk.Combobox(
            top,
            textvariable=self.ordens_year_filter,
            values=[str(datetime.now().year), "Todos"],
            width=10,
            state="readonly",
        )
        self.ordens_year_cb.pack(side="left", padx=4, pady=2)
        self.ordens_year_cb.bind("<<ComboboxSelected>>", lambda _e=None: self._on_ordens_year_change(self.ordens_year_filter.get()))
    BtnCls(top, text="Atualizar", command=self.refresh_ordens).pack(side="left", padx=6, pady=2)
    BtnCls(top, text="Pré-visualizar Peça", command=self.preview_opp_selected).pack(side="left", padx=6, pady=2)

    if self.ordens_use_custom:
        self.tbl_ordens = None
        wrap = ctk.CTkFrame(self.tab_ordens, fg_color="#f7f8fb", corner_radius=10)
        wrap.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        head = ctk.CTkFrame(wrap, fg_color=THEME_HEADER_BG, corner_radius=8)
        head.pack(fill="x", padx=8, pady=(8, 4))
        ctk.CTkLabel(
            head,
            text="OF | ENCOMENDA | REF. INTERNA | REF. EXTERNA | MATERIAL | ESPESSURA | QTD | ESTADO",
            font=("Segoe UI", 11, "bold"),
            text_color="white",
            anchor="w",
        ).pack(fill="x", padx=10, pady=7)
        self.ordens_list = ctk.CTkScrollableFrame(wrap, fg_color="#f7f8fb", corner_radius=8)
        self.ordens_list.pack(fill="both", expand=True, padx=8, pady=(0, 8))
    else:
        ord_style = ""
        tbl_wrap = ttk.Frame(self.tab_ordens)
        tbl_wrap.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.tbl_ordens = ttk.Treeview(
            tbl_wrap,
            columns=("opp", "encomenda", "ref_int", "ref_ext", "material", "espessura", "qtd", "estado"),
            show="headings",
            height=18,
            style=ord_style,
        )
        headings = {
            "opp": "Ordem de Fabrico",
            "encomenda": "Encomenda",
            "ref_int": "Ref. Interna",
            "ref_ext": "Ref. Externa",
            "material": "Material",
            "espessura": "Espessura",
            "qtd": "Quantidade",
            "estado": "Estado",
        }
        for key, text in headings.items():
            self.tbl_ordens.heading(key, text=text)
            self.tbl_ordens.column(key, width=140)
        self.tbl_ordens.column("opp", width=160)
        self.tbl_ordens.column("ref_ext", width=200)
        self.tbl_ordens.column("qtd", width=90)
        vsb = ttk.Scrollbar(tbl_wrap, orient="vertical", command=self.tbl_ordens.yview)
        hsb = ttk.Scrollbar(tbl_wrap, orient="horizontal", command=self.tbl_ordens.xview)
        self.tbl_ordens.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tbl_ordens.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tbl_ordens.bind("<Double-Button-1>", lambda _=None: self.preview_opp_selected())
        self.tbl_ordens.tag_configure("ord_even", background="#fff5f6")
        self.tbl_ordens.tag_configure("ord_odd", background="#fbecee")
        self.tbl_ordens.tag_configure("ord_em_producao", background="#ffd8a8")
        self.tbl_ordens.tag_configure("ord_concluida", background="#b7e4c7")

    self.refresh_ordens_year_options(keep_selection=False)
    self._on_ordens_year_change(self.ordens_year_filter.get())

def build_export(self):
    _ensure_configured()
    use_custom_export = CUSTOM_TK_AVAILABLE and os.environ.get("USE_CUSTOM_EXPORT", "1") != "0"
    FrameCls = ctk.CTkFrame if use_custom_export else ttk.Frame
    LabelCls = ctk.CTkLabel if use_custom_export else ttk.Label
    if use_custom_export and CUSTOM_TK_AVAILABLE:
        def BtnCls(parent, **kwargs):
            kwargs.setdefault("height", 34)
            kwargs.setdefault("corner_radius", 10)
            kwargs.setdefault("fg_color", CTK_PRIMARY_RED)
            kwargs.setdefault("hover_color", CTK_PRIMARY_RED_HOVER)
            kwargs.setdefault("text_color", "#ffffff")
            kwargs.setdefault("border_width", 0)
            return ctk.CTkButton(parent, **kwargs)
    else:
        BtnCls = ttk.Button

    parent = self.tab_export
    if use_custom_export:
        parent = ctk.CTkFrame(self.tab_export, fg_color="#ffffff")
        parent.pack(fill="both", expand=True)

    card = (
        FrameCls(
            parent,
            fg_color="#ffffff",
            corner_radius=10,
            border_width=1,
            border_color="#e7cfd3",
        )
        if use_custom_export
        else FrameCls(parent)
    )
    card.pack(fill="x", padx=10, pady=10)

    if use_custom_export:
        LabelCls(
            card,
            text="Exportações CSV",
            font=("Segoe UI", 14, "bold"),
            text_color="#7a0f1a",
        ).pack(pady=(10, 4))
        LabelCls(
            card,
            text="Exportar dados da aplicação e editar informação de rodapé dos PDFs.",
            text_color="#4a5f77",
        ).pack(pady=(0, 8))
    else:
        LabelCls(card, text="Exportações CSV").pack(pady=(10, 4))

    btns = FrameCls(card, fg_color="transparent") if use_custom_export else FrameCls(card)
    btns.pack(pady=(2, 10))
    BtnCls(btns, text="Clientes", command=lambda: self.export_csv("clientes"), width=120 if use_custom_export else None).pack(side="left", padx=6, pady=2)
    BtnCls(btns, text="Stock", command=lambda: self.export_csv("materiais"), width=120 if use_custom_export else None).pack(side="left", padx=6, pady=2)
    BtnCls(btns, text="Encomendas", command=lambda: self.export_csv("encomendas"), width=120 if use_custom_export else None).pack(side="left", padx=6, pady=2)
    BtnCls(btns, text="Plano Produção", command=lambda: self.export_csv("plano"), width=140 if use_custom_export else None).pack(side="left", padx=6, pady=2)
    BtnCls(btns, text="Qualidade", command=lambda: self.export_csv("qualidade"), width=120 if use_custom_export else None).pack(side="left", padx=6, pady=2)
    BtnCls(btns, text="Calc. Chapa", command=self.open_sheet_calculator, width=120 if use_custom_export else None).pack(side="left", padx=6, pady=2)
    BtnCls(btns, text="Tempos Posto", command=self.edit_tempos_operacao_planeada, width=130 if use_custom_export else None).pack(side="left", padx=6, pady=2)
    if str((getattr(self, "user", {}) or {}).get("role", "") or "").strip() == "Admin":
        BtnCls(btns, text="Utilizadores", command=self.manage_user_accounts, width=130 if use_custom_export else None).pack(side="left", padx=6, pady=2)
    BtnCls(btns, text="Inf.", command=self.edit_empresa_info, width=90 if use_custom_export else None).pack(side="left", padx=6, pady=2)
    BtnCls(btns, text="Cor tema", command=self.choose_primary_color, width=120 if use_custom_export else None).pack(side="left", padx=6, pady=2)
    BtnCls(btns, text="Cor padrao", command=self.reset_primary_color, width=120 if use_custom_export else None).pack(side="left", padx=6, pady=2)

def build_expedicao(self):
    _ensure_configured()
    self.exp_use_custom = CUSTOM_TK_AVAILABLE and os.environ.get("USE_CUSTOM_EXPORT", "1") != "0"
    FrameCls = ctk.CTkFrame if self.exp_use_custom else ttk.Frame
    LabelCls = ctk.CTkLabel if self.exp_use_custom else ttk.Label
    EntryCls = ctk.CTkEntry if self.exp_use_custom else ttk.Entry
    CombCls = ctk.CTkComboBox if self.exp_use_custom else ttk.Combobox
    if self.exp_use_custom and CUSTOM_TK_AVAILABLE:
        def BtnCls(parent, **kwargs):
            kwargs.setdefault("height", 34)
            kwargs.setdefault("corner_radius", 10)
            kwargs.setdefault("fg_color", CTK_PRIMARY_RED)
            kwargs.setdefault("hover_color", CTK_PRIMARY_RED_HOVER)
            kwargs.setdefault("text_color", "#ffffff")
            kwargs.setdefault("border_width", 0)
            return ctk.CTkButton(parent, **kwargs)
        orange_btn = {"fg_color": "#f59e0b", "hover_color": "#d97706"}
    else:
        BtnCls = ttk.Button
        orange_btn = {}
    exp_tbl_style = ""
    if self.exp_use_custom:
        style = ttk.Style()
        style.configure(
            "EXP.Treeview",
            font=("Segoe UI", 10),
            rowheight=27,
            background="#f8fbff",
            fieldbackground="#f8fbff",
            borderwidth=0,
        )
        style.configure(
            "EXP.Treeview.Heading",
            font=("Segoe UI", 10, "bold"),
            background=THEME_HEADER_BG,
            foreground="white",
            relief="flat",
        )
        style.map("EXP.Treeview.Heading", background=[("active", THEME_HEADER_ACTIVE)])
        style.map(
            "EXP.Treeview",
            background=[("selected", THEME_SELECT_BG)],
            foreground=[("selected", THEME_SELECT_FG)],
        )
        exp_tbl_style = "EXP.Treeview"

    parent = self.tab_expedicao
    if self.exp_use_custom:
        parent = ctk.CTkFrame(self.tab_expedicao, fg_color="#ffffff")
        parent.pack(fill="both", expand=True)

    top = (
        FrameCls(
            parent,
            fg_color="#ffffff",
            corner_radius=10,
            border_width=1,
            border_color="#e7cfd3",
        )
        if self.exp_use_custom
        else FrameCls(parent)
    )
    top.pack(fill="x", padx=10, pady=10)
    LabelCls(
        top,
        text="Expedição",
        font=("Segoe UI", 14, "bold") if self.exp_use_custom else None,
        text_color="#7a0f1a" if self.exp_use_custom else None,
    ).pack(side="left", padx=6, pady=4)
    BtnCls(top, text="Emitir Guia OFF", command=self.emitir_expedicao_off, width=150 if self.exp_use_custom else None, **orange_btn).pack(side="left", padx=6, pady=4)
    BtnCls(top, text="Criar Guia Manual", command=self.criar_guia_manual, width=160 if self.exp_use_custom else None).pack(side="left", padx=6, pady=4)
    BtnCls(top, text="Editar Guia", command=self.editar_expedicao, width=130 if self.exp_use_custom else None).pack(side="left", padx=6, pady=4)
    BtnCls(top, text="Anular Guia", command=self.anular_expedicao, width=130 if self.exp_use_custom else None).pack(side="left", padx=6, pady=4)
    BtnCls(top, text="Pré-visualizar PDF", command=self.preview_expedicao_pdf, width=160 if self.exp_use_custom else None).pack(side="left", padx=6, pady=4)
    BtnCls(top, text="Guardar PDF", command=self.save_expedicao_pdf, width=130 if self.exp_use_custom else None).pack(side="left", padx=6, pady=4)
    BtnCls(top, text="Imprimir PDF", command=self.print_expedicao_pdf, width=130 if self.exp_use_custom else None).pack(side="left", padx=6, pady=4)
    BtnCls(top, text="Atualizar", command=self.refresh_expedicao, width=120 if self.exp_use_custom else None, **orange_btn).pack(side="right", padx=6, pady=4)

    filt = (
        FrameCls(
            parent,
            fg_color="#ffffff",
            corner_radius=10,
            border_width=1,
            border_color="#e7cfd3",
        )
        if self.exp_use_custom
        else FrameCls(parent)
    )
    filt.pack(fill="x", padx=10, pady=(0, 6))
    self.exp_filter = StringVar()
    self.exp_estado_filter = StringVar(value="Todas")
    LabelCls(filt, text="Pesquisa").pack(side="left", padx=(6, 4), pady=4)
    self.exp_filter_entry = EntryCls(filt, textvariable=self.exp_filter, width=250 if self.exp_use_custom else 34)
    self.exp_filter_entry.pack(side="left", padx=4, pady=4)
    self.exp_filter_entry.bind("<KeyRelease>", lambda _=None: self.debounce_call("expedicao_filter", self.refresh_expedicao, delay_ms=170))
    LabelCls(filt, text="Estado").pack(side="left", padx=(12, 4), pady=4)
    if self.exp_use_custom:
        self.exp_estado_cb = CombCls(
            filt,
            variable=self.exp_estado_filter,
            values=["Todas", "Não expedida", "Parcialmente expedida", "Totalmente expedida"],
            width=220,
            command=lambda _v=None: self.refresh_expedicao(),
        )
    else:
        self.exp_estado_cb = CombCls(
            filt,
            textvariable=self.exp_estado_filter,
            values=["Todas", "Não expedida", "Parcialmente expedida", "Totalmente expedida"],
            width=24,
            state="readonly",
        )
        self.exp_estado_cb.bind("<<ComboboxSelected>>", lambda _=None: self.refresh_expedicao())
    self.exp_estado_cb.pack(side="left", padx=4, pady=4)
    if self.exp_use_custom:
        self.exp_estado_cb.set("Todas")

    split = FrameCls(parent, fg_color="transparent") if self.exp_use_custom else FrameCls(parent)
    split.pack(fill="both", expand=True, padx=10, pady=(0, 6))
    left = (
        FrameCls(split, fg_color="#ffffff", corner_radius=10, border_width=1, border_color="#e7cfd3")
        if self.exp_use_custom
        else FrameCls(split)
    )
    left.pack(side="left", fill="both", expand=True)
    right = (
        FrameCls(split, fg_color="#ffffff", corner_radius=10, border_width=1, border_color="#e7cfd3")
        if self.exp_use_custom
        else FrameCls(split)
    )
    right.pack(side="left", fill="both", expand=True, padx=(10, 0))

    LabelCls(left, text="Expedições Pendentes (OFF concluídas)").pack(anchor="w", padx=6, pady=(6, 2))
    self.exp_tbl_enc = ttk.Treeview(
        left,
        columns=("numero", "cliente", "estado_prod", "estado_exp", "disponivel"),
        show="headings",
        height=9,
        style=exp_tbl_style,
    )
    self.exp_tbl_enc.heading("numero", text="Encomenda")
    self.exp_tbl_enc.heading("cliente", text="Cliente")
    self.exp_tbl_enc.heading("estado_prod", text="Produção")
    self.exp_tbl_enc.heading("estado_exp", text="Expedição")
    self.exp_tbl_enc.heading("disponivel", text="Qtd Disponível")
    self.exp_tbl_enc.column("numero", width=130, anchor="center")
    self.exp_tbl_enc.column("cliente", width=190, anchor="w")
    self.exp_tbl_enc.column("estado_prod", width=110, anchor="center")
    self.exp_tbl_enc.column("estado_exp", width=130, anchor="center")
    self.exp_tbl_enc.column("disponivel", width=110, anchor="e")
    self.exp_tbl_enc.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
    if self.exp_use_custom and CUSTOM_TK_AVAILABLE:
        sb_enc = ctk.CTkScrollbar(left, orientation="vertical", command=self.exp_tbl_enc.yview)
    else:
        sb_enc = ttk.Scrollbar(left, orient="vertical", command=self.exp_tbl_enc.yview)
    self.exp_tbl_enc.configure(yscrollcommand=sb_enc.set)
    sb_enc.pack(side="right", fill="y", padx=(0, 6), pady=6)
    self.exp_tbl_enc.bind("<<TreeviewSelect>>", self.on_exp_select_encomenda)

    LabelCls(right, text="Artigos Disponíveis para Expedir").pack(anchor="w", padx=6, pady=(6, 2))
    self.exp_tbl_pecas = ttk.Treeview(
        right,
        columns=("peca_id", "ref_int", "ref_ext", "qtd_ok", "qtd_exp", "disp"),
        show="headings",
        height=9,
        style=exp_tbl_style,
    )
    self.exp_tbl_pecas.heading("peca_id", text="Peça")
    self.exp_tbl_pecas.heading("ref_int", text="Ref. Interna")
    self.exp_tbl_pecas.heading("ref_ext", text="Ref. Externa")
    self.exp_tbl_pecas.heading("qtd_ok", text="Produzido")
    self.exp_tbl_pecas.heading("qtd_exp", text="Expedido")
    self.exp_tbl_pecas.heading("disp", text="Disponível")
    self.exp_tbl_pecas.column("peca_id", width=100, anchor="center")
    self.exp_tbl_pecas.column("ref_int", width=160, anchor="w")
    self.exp_tbl_pecas.column("ref_ext", width=180, anchor="w")
    self.exp_tbl_pecas.column("qtd_ok", width=90, anchor="e")
    self.exp_tbl_pecas.column("qtd_exp", width=90, anchor="e")
    self.exp_tbl_pecas.column("disp", width=90, anchor="e")
    self.exp_tbl_pecas.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
    if self.exp_use_custom and CUSTOM_TK_AVAILABLE:
        sb_peca = ctk.CTkScrollbar(right, orientation="vertical", command=self.exp_tbl_pecas.yview)
    else:
        sb_peca = ttk.Scrollbar(right, orient="vertical", command=self.exp_tbl_pecas.yview)
    self.exp_tbl_pecas.configure(yscrollcommand=sb_peca.set)
    sb_peca.pack(side="right", fill="y", padx=(0, 6), pady=6)

    draft = (
        FrameCls(
            parent,
            fg_color="#ffffff",
            corner_radius=10,
            border_width=1,
            border_color="#e7cfd3",
        )
        if self.exp_use_custom
        else FrameCls(parent)
    )
    draft.pack(fill="both", expand=True, padx=10, pady=(0, 6))
    top_d = FrameCls(draft, fg_color="transparent") if self.exp_use_custom else FrameCls(draft)
    top_d.pack(fill="x", padx=6, pady=(6, 2))
    LabelCls(top_d, text="Guia em preparação").pack(side="left", padx=(0, 10))
    LabelCls(top_d, text="Qtd").pack(side="left")
    self.exp_qtd_var = StringVar(value="1")
    EntryCls(top_d, textvariable=self.exp_qtd_var, width=80 if self.exp_use_custom else 10).pack(side="left", padx=4)
    BtnCls(top_d, text="Adicionar linha", command=self.add_exp_linha, width=130 if self.exp_use_custom else None, **orange_btn).pack(side="left", padx=4)
    BtnCls(top_d, text="Remover linha", command=self.remove_exp_linha, width=130 if self.exp_use_custom else None).pack(side="left", padx=4)

    self.exp_tbl_linhas = ttk.Treeview(
        draft,
        columns=("peca_id", "ref_int", "ref_ext", "descricao", "qtd"),
        show="headings",
        height=6,
        style=exp_tbl_style,
    )
    self.exp_tbl_linhas.heading("peca_id", text="Peça")
    self.exp_tbl_linhas.heading("ref_int", text="Ref. Interna")
    self.exp_tbl_linhas.heading("ref_ext", text="Ref. Externa")
    self.exp_tbl_linhas.heading("descricao", text="Descrição")
    self.exp_tbl_linhas.heading("qtd", text="Qtd")
    self.exp_tbl_linhas.column("peca_id", width=90, anchor="center")
    self.exp_tbl_linhas.column("ref_int", width=180, anchor="w")
    self.exp_tbl_linhas.column("ref_ext", width=180, anchor="w")
    self.exp_tbl_linhas.column("descricao", width=360, anchor="w")
    self.exp_tbl_linhas.column("qtd", width=100, anchor="e")
    self.exp_tbl_linhas.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=(0, 6))
    if self.exp_use_custom and CUSTOM_TK_AVAILABLE:
        sb_d = ctk.CTkScrollbar(draft, orientation="vertical", command=self.exp_tbl_linhas.yview)
    else:
        sb_d = ttk.Scrollbar(draft, orient="vertical", command=self.exp_tbl_linhas.yview)
    self.exp_tbl_linhas.configure(yscrollcommand=sb_d.set)
    sb_d.pack(side="right", fill="y", padx=(0, 6), pady=(0, 6))

    hist = (
        FrameCls(
            parent,
            fg_color="#ffffff",
            corner_radius=10,
            border_width=1,
            border_color="#e7cfd3",
        )
        if self.exp_use_custom
        else FrameCls(parent)
    )
    hist.pack(fill="both", expand=True, padx=10, pady=(0, 10))
    LabelCls(hist, text="Histórico de Expedições").pack(anchor="w", padx=6, pady=(6, 2))
    self.exp_tbl_hist = ttk.Treeview(
        hist,
        columns=("numero", "tipo", "encomenda", "cliente", "data", "estado", "linhas"),
        show="headings",
        height=7,
        style=exp_tbl_style,
    )
    self.exp_tbl_hist.heading("numero", text="Número")
    self.exp_tbl_hist.heading("tipo", text="Tipo")
    self.exp_tbl_hist.heading("encomenda", text="Encomenda")
    self.exp_tbl_hist.heading("cliente", text="Cliente")
    self.exp_tbl_hist.heading("data", text="Data")
    self.exp_tbl_hist.heading("estado", text="Estado")
    self.exp_tbl_hist.heading("linhas", text="Linhas")
    self.exp_tbl_hist.column("numero", width=120, anchor="center")
    self.exp_tbl_hist.column("tipo", width=90, anchor="center")
    self.exp_tbl_hist.column("encomenda", width=120, anchor="center")
    self.exp_tbl_hist.column("cliente", width=180, anchor="w")
    self.exp_tbl_hist.column("data", width=140, anchor="center")
    self.exp_tbl_hist.column("estado", width=120, anchor="center")
    self.exp_tbl_hist.column("linhas", width=80, anchor="e")
    self.exp_tbl_hist.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=(0, 6))
    if self.exp_use_custom and CUSTOM_TK_AVAILABLE:
        sb_h = ctk.CTkScrollbar(hist, orientation="vertical", command=self.exp_tbl_hist.yview)
    else:
        sb_h = ttk.Scrollbar(hist, orient="vertical", command=self.exp_tbl_hist.yview)
    self.exp_tbl_hist.configure(yscrollcommand=sb_h.set)
    sb_h.pack(side="right", fill="y", padx=(0, 6), pady=(0, 6))
    self.exp_tbl_hist.bind("<Double-Button-1>", lambda _=None: self.preview_expedicao_pdf())
    self.exp_tbl_hist.bind("<ButtonRelease-1>", self.on_exp_hist_click_open_pdf)

    self.exp_sel_enc_num = None
    self.exp_peca_row_map = {}
    self.exp_draft_linhas = []
    self.exp_tbl_enc.tag_configure("enc_even", background="#eef6ff")
    self.exp_tbl_enc.tag_configure("enc_odd", background="#f8fbff")
    self.exp_tbl_enc.tag_configure("enc_nao", background="#fde68a")
    self.exp_tbl_enc.tag_configure("enc_parcial", background="#bfdbfe")
    self.exp_tbl_enc.tag_configure("enc_total", background="#c8e6c9")
    self.exp_tbl_hist.tag_configure("hist_even", background="#eef6ff")
    self.exp_tbl_hist.tag_configure("hist_odd", background="#f8fbff")
    self.exp_tbl_hist.tag_configure("hist_anulada", background="#ffdada", foreground="#7f1d1d")
    self.refresh_expedicao()

def build_produtos(self):
    _ensure_configured()
    use_custom = CUSTOM_TK_AVAILABLE and os.environ.get("USE_CUSTOM_PROD", "1") != "0"
    # manter em tabela para leitura por colunas
    self.prod_use_full_custom = False
    self.prod_sel_codigo = ""
    FrameCls = ctk.CTkFrame if use_custom else ttk.Frame
    if use_custom and CUSTOM_TK_AVAILABLE:
        def BtnCls(parent, **kwargs):
            kwargs.setdefault("height", 34)
            kwargs.setdefault("corner_radius", 10)
            kwargs.setdefault("fg_color", CTK_PRIMARY_RED)
            kwargs.setdefault("hover_color", CTK_PRIMARY_RED_HOVER)
            kwargs.setdefault("text_color", "#ffffff")
            kwargs.setdefault("border_width", 0)
            return ctk.CTkButton(parent, **kwargs)
        orange_btn = {"fg_color": "#f59e0b", "hover_color": "#d97706"}
    else:
        BtnCls = ttk.Button
        orange_btn = {}
    LabelCls = ctk.CTkLabel if use_custom else ttk.Label
    EntryCls = ctk.CTkEntry if use_custom else ttk.Entry
    bg = "#ffffff" if use_custom else None
    parent = self.tab_produtos
    if use_custom:
        parent = ctk.CTkFrame(self.tab_produtos, fg_color=bg)
        parent.pack(fill="both", expand=True)

    # barra topo
    top = (
        FrameCls(parent, fg_color="white", corner_radius=10, border_width=0)
        if use_custom
        else FrameCls(parent)
    )
    top.pack(fill="x", padx=8, pady=8)
    LabelCls(top, text="Pesquisa").pack(side="left", padx=4)
    self.prod_filter = StringVar()
    prod_filter_entry = EntryCls(top, textvariable=self.prod_filter, width=200)
    prod_filter_entry.pack(side="left", padx=4)
    try:
        prod_filter_entry.bind("<Return>", lambda _e: self.refresh_produtos())
    except Exception:
        pass
    BtnCls(top, text="Novo", command=self.dialog_novo_produto, width=90, **orange_btn).pack(side="left", padx=4, pady=2)
    BtnCls(top, text="Guardar", command=self.guardar_produto, width=90).pack(side="left", padx=4, pady=2)
    BtnCls(top, text="Editar", command=self.editar_produto_dialog, width=90).pack(side="left", padx=4, pady=2)
    BtnCls(top, text="Remover", command=self.remover_produto, width=90).pack(side="left", padx=4, pady=2)
    BtnCls(top, text="Baixa", command=self.produto_dar_baixa_dialog, width=90).pack(side="left", padx=4, pady=2)
    BtnCls(top, text="Movimentos Operador", command=self.produto_mov_operador_dialog, width=165).pack(side="left", padx=4, pady=2)
    BtnCls(top, text="Pre-visualizar PDF", command=self.preview_produtos_stock_pdf, width=170).pack(side="left", padx=4, pady=2)
    BtnCls(top, text="Atualizar", command=self.refresh_produtos, width=90).pack(side="right", padx=4, pady=2)

    # listagem produtos
    if self.prod_use_full_custom:
        self.tbl_produtos = None
        tbl_wrap = ctk.CTkFrame(parent, fg_color=bg, corner_radius=10)
        tbl_wrap.pack(fill="both", expand=True, padx=8, pady=4)
        head = ctk.CTkFrame(tbl_wrap, fg_color=THEME_HEADER_BG, corner_radius=8)
        head.pack(fill="x", padx=8, pady=(8, 4))
        ctk.CTkLabel(
            head,
            text="CÓDIGO | DESCRIÇÃO | CATEGORIA | TIPO | QTD | ALERTA | PREÇO/UNID",
            font=("Segoe UI", 11, "bold"),
            text_color="white",
            anchor="w",
        ).pack(fill="x", padx=10, pady=7)
        self.prod_list = ctk.CTkScrollableFrame(tbl_wrap, fg_color="#f7f8fb", corner_radius=8)
        self.prod_list.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        self._prod_btns = {}
    else:
        cols = ("codigo", "descricao", "categoria", "subcat", "tipo", "dimensoes", "metros", "peso_unid", "unid", "qty", "p_compra", "preco_unid", "pvp1", "pvp2", "alerta")
        tbl_wrap = ctk.CTkFrame(
            parent,
            fg_color="white",
            corner_radius=10,
            border_width=1,
            border_color="#e7cfd3",
        ) if use_custom else FrameCls(parent)
        tbl_wrap.pack(fill="both", expand=True, padx=8, pady=4)
        prod_style = ""
        if use_custom:
            style = ttk.Style()
            style.configure(
                "PROD.Treeview",
                font=("Segoe UI", 10),
                rowheight=27,
                background="#f8fbff",
                fieldbackground="#f8fbff",
                borderwidth=0,
            )
            style.configure(
                "PROD.Treeview.Heading",
                font=("Segoe UI", 10, "bold"),
                background=THEME_HEADER_BG,
                foreground="white",
                relief="flat",
            )
            style.map("PROD.Treeview.Heading", background=[("active", THEME_HEADER_ACTIVE)])
            style.map(
                "PROD.Treeview",
                background=[("selected", THEME_SELECT_BG)],
                foreground=[("selected", THEME_SELECT_FG)],
            )
            prod_style = "PROD.Treeview"
        self.tbl_produtos = ttk.Treeview(tbl_wrap, columns=cols, show="headings", height=12, style=prod_style)
        headings = {
            "codigo": "Código",
            "descricao": "Descrição",
            "categoria": "Categoria",
            "subcat": "Subcategoria",
            "tipo": "Tipo",
            "dimensoes": "Dimensões",
            "metros": "Metros",
            "peso_unid": "Peso/Un.",
            "unid": "Un.",
            "qty": "Qtd",
            "p_compra": "Compra (â‚¬)",
            "preco_unid": "Preço/Unid (€)",
            "pvp1": "PVP1 (â‚¬)",
            "pvp2": "PVP2 (â‚¬)",
            "alerta": "Alerta",
        }
        for c in cols:
            self.tbl_produtos.heading(c, text=headings[c])
            anchor = "center" if c in ("unid", "qty", "alerta", "metros", "peso_unid") else ("e" if c in ("p_compra", "preco_unid", "pvp1", "pvp2") else "w")
            width = 70 if c in ("unid", "qty", "alerta") else 90
            if c == "descricao":
                width = 220
            if c == "dimensoes":
                width = 130
            self.tbl_produtos.column(c, width=width, anchor=anchor)
        # vista principal simplificada (mais legível)
        self.tbl_produtos["displaycolumns"] = (
            "codigo",
            "descricao",
            "categoria",
            "tipo",
            "qty",
            "alerta",
            "preco_unid",
        )
        self.tbl_produtos.heading("preco_unid", text="Preço (€/Un)")
        self.tbl_produtos.column("codigo", width=110, anchor="center")
        self.tbl_produtos.column("descricao", width=360, anchor="w")
        self.tbl_produtos.column("categoria", width=170, anchor="w")
        self.tbl_produtos.column("tipo", width=180, anchor="w")
        self.tbl_produtos.column("qty", width=90, anchor="center")
        self.tbl_produtos.column("alerta", width=90, anchor="center")
        self.tbl_produtos.column("preco_unid", width=120, anchor="e")
        if use_custom and CUSTOM_TK_AVAILABLE:
            vsb = ctk.CTkScrollbar(tbl_wrap, orientation="vertical", command=self.tbl_produtos.yview)
            hsb = ctk.CTkScrollbar(tbl_wrap, orientation="horizontal", command=self.tbl_produtos.xview)
        else:
            vsb = ttk.Scrollbar(tbl_wrap, orient="vertical", command=self.tbl_produtos.yview)
            hsb = ttk.Scrollbar(tbl_wrap, orient="horizontal", command=self.tbl_produtos.xview)
        self.tbl_produtos.configure(yscroll=vsb.set, xscroll=hsb.set)
        self.tbl_produtos.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tbl_wrap.rowconfigure(0, weight=1)
        tbl_wrap.columnconfigure(0, weight=1)
        self.tbl_produtos.bind("<<TreeviewSelect>>", self.on_produto_select)
        self.tbl_produtos.tag_configure("even", background="#fff5f6")
        self.tbl_produtos.tag_configure("odd", background="#fbecee")
        self.tbl_produtos.tag_configure("warn", background="#ffe6bf", foreground="#8a5a00")

    # formulário
    form = (
        ctk.CTkFrame(
            parent,
            fg_color="white",
            corner_radius=10,
            border_width=1,
            border_color="#e7cfd3",
        )
        if use_custom
        else FrameCls(parent)
    )
    form.pack(fill="x", padx=10, pady=6)

    def _prod_w(v):
        try:
            n = int(v)
        except Exception:
            n = 120
        if use_custom:
            return n if n >= 100 else n * 8
        return n

    def row(label, var, width=200, values=None, parent_row=None):
        host = parent_row if parent_row is not None else form
        if use_custom:
            card = ctk.CTkFrame(
                host,
                fg_color="#ffffff",
                corner_radius=8,
                border_width=1,
                border_color="#dfc4c9",
            )
            card.pack(side="left", padx=4, pady=4, fill="y")
            ctk.CTkLabel(card, text=label, font=("Segoe UI", 11, "bold"), text_color="#6f1624").pack(anchor="w", padx=8, pady=(4, 2))
            if values:
                cb = ctk.CTkComboBox(
                    card,
                    variable=var,
                    values=[str(v) for v in values],
                    width=_prod_w(width),
                    state="readonly",
                )
                cb.pack(padx=8, pady=(0, 6))
                return cb
            e = EntryCls(card, textvariable=var, width=_prod_w(width))
            e.pack(padx=8, pady=(0, 6))
            return e
        else:
            LabelCls(host, text=label).pack(side="left", padx=4, pady=2)
            if values:
                cb = ttk.Combobox(host, textvariable=var, values=values, width=width, state="normal")
                cb.pack(side="left", padx=4, pady=2)
                return cb
            e = EntryCls(host, textvariable=var, width=_prod_w(width))
            e.pack(side="left", padx=4, pady=2)
            return e

    self.prod_codigo = StringVar()
    self.prod_descricao = StringVar()
    self.prod_categoria = StringVar()
    self.prod_subcat = StringVar()
    self.prod_tipo = StringVar()
    self.prod_fab = StringVar()
    self.prod_modelo = StringVar()
    self.prod_unid = StringVar(value="UN")
    self.prod_qty = StringVar(value="0")
    self.prod_alerta = StringVar(value="0")
    self.prod_pcompra = StringVar(value="0")
    self.prod_pvp1 = StringVar(value="0")
    self.prod_pvp2 = StringVar(value="0")
    self.prod_dimensoes = StringVar()
    self.prod_metros = StringVar(value="0")
    self.prod_peso_unid = StringVar(value="0")
    self.prod_obs = StringVar()

    if use_custom:
        r1 = ctk.CTkFrame(form, fg_color="transparent")
        r1.pack(fill="x", padx=6, pady=(4, 2))
        r2 = ctk.CTkFrame(form, fg_color="transparent")
        r2.pack(fill="x", padx=6, pady=(2, 6))

        row("Código", self.prod_codigo, width=90, parent_row=r1)
        row("Descrição", self.prod_descricao, width=240, parent_row=r1)
        row("Categoria", self.prod_categoria, width=180, values=PROD_CATEGORIAS, parent_row=r1)
        row("Subcat.", self.prod_subcat, width=150, values=PROD_SUBCATS, parent_row=r1)
        row("Tipo", self.prod_tipo, width=150, values=PROD_TIPOS, parent_row=r1)
        row("Dimensões", self.prod_dimensoes, width=160, parent_row=r1)
        row("Metros", self.prod_metros, width=90, parent_row=r1)
        row("Peso/Un", self.prod_peso_unid, width=90, parent_row=r1)

        row("Fabricante", self.prod_fab, width=170, parent_row=r2)
        row("Modelo", self.prod_modelo, width=170, parent_row=r2)
        row("Unid.", self.prod_unid, width=90, values=PROD_UNIDS, parent_row=r2)
        row("Qtd", self.prod_qty, width=90, parent_row=r2)
        row("Alerta", self.prod_alerta, width=90, parent_row=r2)
        row("Compra â‚¬", self.prod_pcompra, width=110, parent_row=r2)
        row("PVP1 â‚¬", self.prod_pvp1, width=110, parent_row=r2)
        row("PVP2 â‚¬", self.prod_pvp2, width=110, parent_row=r2)
        row("Obs", self.prod_obs, width=260, parent_row=r2)
    else:
        row("Código", self.prod_codigo, width=80)
        row("Descrição", self.prod_descricao, width=220)
        row("Categoria", self.prod_categoria, width=18, values=PROD_CATEGORIAS)
        row("Subcat.", self.prod_subcat, width=16, values=PROD_SUBCATS)
        row("Tipo", self.prod_tipo, width=16, values=PROD_TIPOS)
        row("Dimensões", self.prod_dimensoes, width=120)
        row("Metros", self.prod_metros, width=60)
        row("Peso/Un", self.prod_peso_unid, width=60)
        row("Fabricante", self.prod_fab, width=120)
        row("Modelo", self.prod_modelo, width=120)
        row("Unid.", self.prod_unid, width=6, values=PROD_UNIDS)
        row("Qtd", self.prod_qty, width=60)
        row("Alerta", self.prod_alerta, width=60)
        row("Compra â‚¬", self.prod_pcompra, width=80)
        row("PVP1 â‚¬", self.prod_pvp1, width=80)
        row("PVP2 â‚¬", self.prod_pvp2, width=80)
        row("Obs", self.prod_obs, width=180)
    self.prod_categoria.trace_add("write", lambda *_: self.product_apply_category_mode())

    self.novo_produto()
    self.refresh_produtos()

def build_ne(self):
    _ensure_configured()
    use_custom = CUSTOM_TK_AVAILABLE and os.environ.get("USE_CUSTOM_NE", "1") != "0"
    FrameCls = ctk.CTkFrame if use_custom else ttk.Frame
    BtnCls = ctk.CTkButton if use_custom else ttk.Button
    LabelCls = ctk.CTkLabel if use_custom else ttk.Label
    EntryCls = ctk.CTkEntry if use_custom else ttk.Entry
    bg = "#fff8f9" if use_custom else None
    parent_ne = self.tab_ne
    if use_custom:
        parent_ne = ctk.CTkFrame(self.tab_ne, fg_color=bg)
        parent_ne.pack(fill="both", expand=True)
    w_num = 20 if not use_custom else 220
    w_forn = 42 if not use_custom else 420
    w_contacto = 34 if not use_custom else 300
    w_entrega = 20 if not use_custom else 210
    top = ctk.CTkFrame(
        parent_ne,
        fg_color="white",
        corner_radius=10,
        border_width=1,
        border_color="#e7cfd3",
    ) if use_custom else FrameCls(parent_ne)
    top.pack(fill="x", padx=8, pady=8)
    if use_custom:
        self.ne_logo_img = None
        try:
            logo_path = get_orc_logo_path()
            if logo_path and os.path.exists(logo_path):
                from PIL import Image, ImageTk
                img = Image.open(logo_path).convert("RGBA")
                rw = 88
                rh = max(24, int(img.height * (rw / max(1, img.width))))
                img = img.resize((rw, rh))
                self.ne_logo_img = ImageTk.PhotoImage(img)
        except Exception:
            self.ne_logo_img = None
        if self.ne_logo_img:
            ctk.CTkLabel(top, image=self.ne_logo_img, text="").pack(side="left", padx=(6, 8), pady=2)
    self.ne_filter = StringVar()
    self.ne_show_convertidas = BooleanVar(value=False)
    self.ne_estado_filter = StringVar(value="Ativas")
    LabelCls(top, text="Pesquisa").pack(side="left", padx=4)
    EntryCls(top, textvariable=self.ne_filter, width=170 if use_custom else None).pack(side="left", padx=4)
    if use_custom:
        ctk.CTkSwitch(
            top,
            text="Ver convertidas",
            variable=self.ne_show_convertidas,
            onvalue=True,
            offvalue=False,
            command=self.refresh_ne,
            width=130,
        ).pack(side="left", padx=(6, 8))
    else:
        ttk.Checkbutton(top, text="Ver convertidas", variable=self.ne_show_convertidas, command=self.refresh_ne).pack(side="left", padx=(6, 8))
    if use_custom:
        def add_top_btn(txt, cmd, w=88):
            BtnCls(
                top,
                text=txt,
                command=cmd,
                width=w,
                height=28,
                corner_radius=8,
                font=("Segoe UI", 11, "bold"),
            ).pack(side="left", padx=3)
        add_top_btn("Nova NE", self.nova_ne, 88)
        add_top_btn("Aprovar NE", self.aprovar_ne, 104)
        add_top_btn("Entregar NE", self.confirmar_entrega_ne, 106)
        add_top_btn("Apagar NE", self.remover_ne, 102)
        add_top_btn("Pedido Cotacao", self.preview_ne_cotacao, 124)
        add_top_btn("Pre-visualizar NE", self.preview_ne, 132)
        add_top_btn("Gerar NEs", self.gerar_nes_por_fornecedor, 96)
    else:
        BtnCls(top, text="Nova NE", command=self.nova_ne).pack(side="left", padx=4)
        BtnCls(top, text="Aprovar NE", command=self.aprovar_ne).pack(side="left", padx=4)
        BtnCls(top, text="Entregar NE", command=self.confirmar_entrega_ne).pack(side="left", padx=4)
        BtnCls(top, text="Apagar NE", command=self.remover_ne).pack(side="left", padx=4)
        BtnCls(top, text="Pedido Cotacao", command=self.preview_ne_cotacao).pack(side="left", padx=4)
        BtnCls(top, text="Pre-visualizar NE", command=self.preview_ne).pack(side="left", padx=4)
        BtnCls(top, text="Gerar NEs por Fornecedor", command=self.gerar_nes_por_fornecedor).pack(side="left", padx=4)

    top_filter = ctk.CTkFrame(
        parent_ne,
        fg_color="white",
        corner_radius=10,
        border_width=1,
        border_color="#e7cfd3",
    ) if use_custom else FrameCls(parent_ne)
    top_filter.pack(fill="x", padx=8, pady=(0, 4))
    LabelCls(top_filter, text="Estado").pack(side="left", padx=(4, 6))
    if use_custom:
        self.ne_estado_segment = ctk.CTkSegmentedButton(
            top_filter,
            values=["Ativas", "Todos", "Em edição", "Aprovada", "Parcial", "Entregue", "Convertida"],
            command=self._on_ne_estado_filter_click,
            width=520,
        )
        self.ne_estado_segment.configure(
            fg_color="#f9e8ea",
            selected_color="#dbe8ff",
            selected_hover_color="#cddfff",
            unselected_color="#f4f6f9",
            unselected_hover_color="#e7ebf0",
            text_color="#111111",
        )
        self.ne_estado_segment.pack(side="left", padx=4, pady=2)
        self.ne_estado_segment.set("Ativas")
    else:
        self.ne_estado_cb = ttk.Combobox(
            top_filter,
            textvariable=self.ne_estado_filter,
            values=["Ativas", "Todos", "Em edição", "Aprovada", "Parcial", "Entregue", "Convertida"],
            width=18,
            state="readonly",
        )
        self.ne_estado_cb.pack(side="left", padx=4, pady=2)
        self.ne_estado_cb.bind("<<ComboboxSelected>>", lambda _e=None: self.refresh_ne())

    cols = ("numero", "fornecedor", "data_entrega", "estado", "total")
    ne_tbl_wrap = ctk.CTkFrame(
        parent_ne,
        fg_color="white",
        corner_radius=10,
        border_width=1,
        border_color="#e7cfd3",
    ) if use_custom else FrameCls(parent_ne)
    ne_tbl_wrap.pack(fill="x", padx=8, pady=6)
    ne_style = ""
    ne_lin_style = ""
    if use_custom:
        style = ttk.Style()
        style.configure(
            "NE.Treeview",
            font=("Segoe UI", 10),
            rowheight=29,
            background="#f8fbff",
            fieldbackground="#f8fbff",
            borderwidth=0,
            bordercolor="#d7deea",
            lightcolor="#d7deea",
            darkcolor="#d7deea",
            relief="solid",
        )
        style.configure(
            "NE.Treeview.Heading",
            font=("Segoe UI", 10, "bold"),
            background=THEME_HEADER_BG,
            foreground="white",
            relief="flat",
        )
        style.map("NE.Treeview.Heading", background=[("active", THEME_HEADER_ACTIVE)])
        style.map(
            "NE.Treeview",
            background=[("selected", THEME_SELECT_BG)],
            foreground=[("selected", THEME_SELECT_FG)],
        )
        style.configure(
            "NEL.Treeview",
            font=("Segoe UI", 10),
            rowheight=27,
            background="#f8fbff",
            fieldbackground="#f8fbff",
            borderwidth=0,
            bordercolor="#d7deea",
            lightcolor="#d7deea",
            darkcolor="#d7deea",
            relief="solid",
        )
        style.configure(
            "NEL.Treeview.Heading",
            font=("Segoe UI", 10, "bold"),
            background=THEME_HEADER_BG,
            foreground="white",
            relief="flat",
        )
        style.map("NEL.Treeview.Heading", background=[("active", THEME_HEADER_ACTIVE)])
        style.map(
            "NEL.Treeview",
            background=[("selected", THEME_SELECT_BG)],
            foreground=[("selected", THEME_SELECT_FG)],
        )
        ne_style = "NE.Treeview"
        ne_lin_style = "NEL.Treeview"
    self.tbl_ne = ttk.Treeview(ne_tbl_wrap, columns=cols, show="headings", height=8, style=ne_style)
    self.tbl_ne.heading("numero", text="Número")
    self.tbl_ne.heading("fornecedor", text="Fornecedor")
    self.tbl_ne.heading("data_entrega", text="Entrega")
    self.tbl_ne.heading("estado", text="Estado")
    self.tbl_ne.heading("total", text="Total (€)")
    self.tbl_ne.column("numero", width=120, anchor="w")
    self.tbl_ne.column("fornecedor", width=240, anchor="w")
    self.tbl_ne.column("data_entrega", width=110, anchor="center")
    self.tbl_ne.column("estado", width=110, anchor="center")
    self.tbl_ne.column("total", width=100, anchor="e")
    ne_vsb = ttk.Scrollbar(ne_tbl_wrap, orient="vertical", command=self.tbl_ne.yview)
    ne_hsb = ttk.Scrollbar(ne_tbl_wrap, orient="horizontal", command=self.tbl_ne.xview)
    self.tbl_ne.configure(yscroll=ne_vsb.set, xscroll=ne_hsb.set)
    self.tbl_ne.grid(row=0, column=0, sticky="nsew")
    ne_vsb.grid(row=0, column=1, sticky="ns")
    ne_hsb.grid(row=1, column=0, sticky="ew")
    ne_tbl_wrap.columnconfigure(0, weight=1)
    self.tbl_ne.bind("<<TreeviewSelect>>", self.on_ne_select)
    self.tbl_ne.tag_configure("even", background="#edf4ff")
    self.tbl_ne.tag_configure("odd", background="#e5eefb")
    self.tbl_ne.tag_configure("warn_even", background="#fff3d4", foreground="#8a6d00")
    self.tbl_ne.tag_configure("warn_odd", background="#ffefc7", foreground="#8a6d00")
    self.tbl_ne.tag_configure("aprovada_even", background="#fff6db", foreground="#8a6d00")
    self.tbl_ne.tag_configure("aprovada_odd", background="#fff2cd", foreground="#8a6d00")
    self.tbl_ne.tag_configure("parcial_even", background="#ffe7c1", foreground="#8a4b00")
    self.tbl_ne.tag_configure("parcial_odd", background="#ffe1b2", foreground="#8a4b00")
    self.tbl_ne.tag_configure("entregue_even", background="#d9f2df", foreground="#0f5132")
    self.tbl_ne.tag_configure("entregue_odd", background="#cfe9d5", foreground="#0f5132")
    self.tbl_ne.tag_configure("convertida_even", background="#eceff3", foreground="#5c6773")
    self.tbl_ne.tag_configure("convertida_odd", background="#e5e9ee", foreground="#5c6773")

    form = ctk.CTkFrame(
        parent_ne,
        fg_color="white",
        corner_radius=10,
        border_width=1,
        border_color="#e7cfd3",
    ) if use_custom else FrameCls(parent_ne)
    form.pack(fill="x", padx=10, pady=6)
    if use_custom:
        ctk.CTkLabel(
            form,
            text="Detalhe da Nota Selecionada",
            font=("Segoe UI", 12, "bold"),
            text_color="#7a0f1a",
        ).grid(row=0, column=0, columnspan=10, sticky="w", padx=6, pady=(6, 2))
        row_base = 1
    else:
        row_base = 0
    self.ne_num = StringVar()
    self.ne_fornecedor = StringVar()
    self.ne_fornecedor_id = StringVar()
    self.ne_contacto = StringVar()
    self.ne_entrega = StringVar()
    self.ne_obs = StringVar()
    self.ne_local_descarga = StringVar()
    self.ne_meio_transporte = StringVar()
    LabelCls(form, text="Número").grid(row=row_base + 0, column=0, sticky="w", padx=4, pady=2)
    self.ne_num_entry = EntryCls(form, textvariable=self.ne_num, width=w_num)
    try:
        self.ne_num_entry.configure(state="readonly")
    except Exception:
        pass
    self.ne_num_entry.grid(row=row_base + 0, column=1, padx=4, pady=2, sticky="w")
    LabelCls(form, text="Fornecedor").grid(row=row_base + 0, column=2, sticky="w", padx=4, pady=2)
    if use_custom:
        self.ne_fornecedor_cb = ctk.CTkComboBox(form, variable=self.ne_fornecedor, values=[""], width=w_forn, state="normal", command=lambda _v=None: self.on_ne_fornecedor_change())
        self.ne_fornecedor_cb.grid(row=row_base + 0, column=3, padx=4, pady=2, sticky="we")
    else:
        self.ne_fornecedor_cb = ttk.Combobox(form, textvariable=self.ne_fornecedor, width=w_forn, state="normal")
        self.ne_fornecedor_cb.grid(row=row_base + 0, column=3, padx=4, pady=2)
        self.ne_fornecedor_cb.bind("<<ComboboxSelected>>", self.on_ne_fornecedor_change)
    LabelCls(form, text="Contacto").grid(row=row_base + 0, column=4, sticky="w", padx=4, pady=2)
    EntryCls(form, textvariable=self.ne_contacto, width=w_contacto).grid(row=row_base + 0, column=5, padx=4, pady=2)
    LabelCls(form, text="Data Entrega").grid(row=row_base + 0, column=6, sticky="w", padx=4, pady=2)
    EntryCls(form, textvariable=self.ne_entrega, width=w_entrega).grid(row=row_base + 0, column=7, padx=4, pady=2)
    BtnCls(form, text="Calendario", command=lambda: self.pick_date(self.ne_entrega, parent=form.winfo_toplevel())).grid(row=row_base + 0, column=8, padx=4, pady=2)
    BtnCls(form, text="Gerir Fornecedores", command=self.manage_fornecedores).grid(row=row_base + 0, column=9, padx=4, pady=2)
    BtnCls(form, text="Atualizar", command=self.refresh_ne).grid(row=row_base + 0, column=10, padx=4, pady=2)
    if use_custom:
        ctk.CTkButton(
            form,
            text="Guardar",
            command=self.guardar_ne,
            width=100,
            height=30,
            corner_radius=8,
            fg_color="#f59e0b",
            hover_color="#d97706",
            text_color="#ffffff",
        ).grid(row=row_base + 0, column=11, padx=4, pady=2)
    else:
        BtnCls(form, text="Guardar", command=self.guardar_ne).grid(row=row_base + 0, column=11, padx=4, pady=2)
    LabelCls(form, text="Observações").grid(row=row_base + 1, column=0, sticky="w", padx=4, pady=2)
    EntryCls(form, textvariable=self.ne_obs, width=58 if use_custom else 56).grid(row=row_base + 1, column=1, columnspan=3, padx=4, pady=2, sticky="we")
    LabelCls(form, text="Local de descarga").grid(row=row_base + 1, column=4, sticky="w", padx=4, pady=2)
    if use_custom:
        self.ne_local_cb = ctk.CTkComboBox(
            form,
            variable=self.ne_local_descarga,
            values=["Nossas Instalações", "Vossas Instalações"],
            width=320,
            state="normal",
        )
        self.ne_local_cb.grid(row=row_base + 1, column=5, padx=4, pady=2, sticky="we")
    else:
        self.ne_local_cb = ttk.Combobox(
            form,
            textvariable=self.ne_local_descarga,
            values=("Nossas Instalações", "Vossas Instalações"),
            width=30,
            state="normal",
        )
        self.ne_local_cb.grid(row=row_base + 1, column=5, padx=4, pady=2, sticky="we")
    LabelCls(form, text="Meio de transporte").grid(row=row_base + 1, column=6, sticky="w", padx=4, pady=2)
    if use_custom:
        self.ne_transporte_cb = ctk.CTkComboBox(
            form,
            variable=self.ne_meio_transporte,
            values=["Nosso Transporte", "Vosso transporte"],
            width=250,
            state="normal",
        )
        self.ne_transporte_cb.grid(row=row_base + 1, column=7, padx=4, pady=2, sticky="we")
    else:
        self.ne_transporte_cb = ttk.Combobox(
            form,
            textvariable=self.ne_meio_transporte,
            values=("Nosso Transporte", "Vosso transporte"),
            width=24,
            state="normal",
        )
        self.ne_transporte_cb.grid(row=row_base + 1, column=7, padx=4, pady=2, sticky="we")
    form.grid_columnconfigure(3, weight=1)
    form.grid_columnconfigure(5, weight=1)

    self.ne_lin_frame = ctk.CTkFrame(
        parent_ne,
        fg_color="white",
        corner_radius=10,
        border_width=1,
        border_color="#e7cfd3",
    ) if use_custom else FrameCls(parent_ne)
    self.ne_lin_frame.pack(fill="both", expand=True, padx=8, pady=8)
    if use_custom:
        ctk.CTkLabel(
            self.ne_lin_frame,
            text="Linhas da Nota Selecionada",
            font=("Segoe UI", 12, "bold"),
            text_color="#7a0f1a",
        ).pack(anchor="w", padx=8, pady=(8, 4))
    lin_cols = ("ref", "descricao", "fornecedor", "origem", "qtd", "unid", "preco", "desconto", "iva", "total", "entregue")
    self.tbl_ne_linhas = ttk.Treeview(self.ne_lin_frame, columns=lin_cols, show="headings", height=10, style=ne_lin_style)
    headings_lin = {
        "ref": "Produto",
        "descricao": "Descrição",
        "fornecedor": "Fornecedor (linha)",
        "origem": "Origem",
        "qtd": "Qtd",
        "unid": "Un.",
        "preco": "Preço (€)",
        "desconto": "Desc. (%)",
        "iva": "IVA (%)",
        "total": "Total (€)",
        "entregue": "Entregue",
    }
    for c in lin_cols:
        anchor = "w" if c in ("ref", "descricao", "fornecedor", "origem") else ("center" if c in ("qtd", "unid", "entregue", "desconto", "iva") else "e")
        width = 120 if c in ("ref", "descricao") else 180 if c == "fornecedor" else 110 if c == "origem" else 70 if c in ("qtd", "unid", "entregue") else 85 if c in ("desconto", "iva") else 90
        self.tbl_ne_linhas.heading(c, text=headings_lin[c])
        self.tbl_ne_linhas.column(c, width=width, anchor=anchor)
    self.tbl_ne_linhas.column("descricao", width=220, anchor="w")
    self.tbl_ne_linhas.column("fornecedor", width=240, anchor="w")
    self.tbl_ne_linhas.column("origem", width=120, anchor="center")
    ne_lin_vsb = ttk.Scrollbar(self.ne_lin_frame, orient="vertical", command=self.tbl_ne_linhas.yview)
    ne_lin_hsb = ttk.Scrollbar(self.ne_lin_frame, orient="horizontal", command=self.tbl_ne_linhas.xview)
    self.tbl_ne_linhas.configure(yscroll=ne_lin_vsb.set, xscroll=ne_lin_hsb.set)
    self.tbl_ne_linhas.pack(fill="both", expand=True, side="left")
    ne_lin_vsb.pack(side="right", fill="y")
    ne_lin_hsb.pack(side="bottom", fill="x")
    self.tbl_ne_linhas.bind("<<TreeviewSelect>>", self.on_ne_lin_select)
    self.tbl_ne_linhas.bind("<Double-1>", self.ne_edit_linha)
    self.tbl_ne_linhas.tag_configure("even", background="#eef6ff")
    self.tbl_ne_linhas.tag_configure("odd", background="#e6f0fc")
    self.tbl_ne_linhas.tag_configure("lin_parcial_even", background="#fff5d8", foreground="#8a6d00")
    self.tbl_ne_linhas.tag_configure("lin_parcial_odd", background="#fff1ca", foreground="#8a6d00")
    self.tbl_ne_linhas.tag_configure("lin_entregue_even", background="#d9f2df", foreground="#0f5132")
    self.tbl_ne_linhas.tag_configure("lin_entregue_odd", background="#cfe9d5", foreground="#0f5132")
    btn_col = FrameCls(self.ne_lin_frame, fg_color="transparent") if use_custom else FrameCls(self.ne_lin_frame)
    btn_col.pack(side="right", fill="y", padx=6)
    BtnCls(btn_col, text="+ Linha", command=self.ne_add_linha).pack(side="top", padx=6, pady=4)
    BtnCls(btn_col, text="Remover Linha", command=self.ne_del_linha).pack(side="top", padx=6, pady=4)
    BtnCls(btn_col, text="Associar Fatura", command=self.associar_fatura_ne).pack(side="top", padx=6, pady=4)
    BtnCls(btn_col, text="Documentos", command=self.show_ne_documentos).pack(side="top", padx=6, pady=4)

    total_bar = ctk.CTkFrame(
        parent_ne,
        fg_color="white",
        corner_radius=10,
        border_width=1,
        border_color="#e7cfd3",
    ) if use_custom else FrameCls(parent_ne)
    total_bar.pack(fill="x", padx=10, pady=(0, 6))
    self.ne_total_lbl = LabelCls(total_bar, text="Total: 0 €", font=("Arial", 12, "bold"))
    self.ne_total_lbl.pack(anchor="e", padx=10, pady=6)

    # No arranque, abrir o formulario em branco sem criar registo na BD.
    self.nova_ne(create_draft=False)
    self.refresh_ne()
    self.ne_blink_on = False
    self.ne_blink_schedule()

