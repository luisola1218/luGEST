import os
from tkinter import END, StringVar
from tkinter import ttk

try:
    import customtkinter as ctk  # type: ignore
except Exception:
    ctk = None  # type: ignore


def build_clientes(app, *, custom_tk_available, cond_pagamento_opcoes):
    app.clientes_use_custom = custom_tk_available and ctk is not None and os.environ.get("USE_CUSTOM_CLIENTES", "1") != "0"
    app.clientes_use_full_custom = False
    app._clientes_btns = {}
    app.selected_cliente_codigo = ""
    primary = getattr(app, "CTK_PRIMARY_RED", "#ba2d3d")
    primary_hover = getattr(app, "CTK_PRIMARY_RED_HOVER", "#a32035")
    theme_header_bg = getattr(app, "THEME_HEADER_BG", "#9b2233")
    theme_header_active = getattr(app, "THEME_HEADER_ACTIVE", "#992233")
    theme_select_bg = getattr(app, "THEME_SELECT_BG", "#fde2e4")
    theme_select_fg = getattr(app, "THEME_SELECT_FG", "#7a0f1a")
    parent_c = app.tab_clientes
    if app.clientes_use_custom:
        parent_c = ctk.CTkFrame(app.tab_clientes, fg_color="#ffffff")
        parent_c.pack(fill="both", expand=True)

    top_frame = ctk.CTkFrame(
        parent_c,
        fg_color="white",
        corner_radius=10,
        border_width=1,
        border_color="#e7cfd3",
    ) if app.clientes_use_custom else ttk.Frame(parent_c)
    top_frame.pack(fill="x", padx=10, pady=10)

    filter_row = ctk.CTkFrame(
        parent_c,
        fg_color="white",
        corner_radius=10,
        border_width=1,
        border_color="#e7cfd3",
    ) if app.clientes_use_custom else ttk.Frame(parent_c)
    filter_row.pack(fill="x", padx=10, pady=(0, 6))
    (ctk.CTkLabel if app.clientes_use_custom else ttk.Label)(filter_row, text="Pesquisa").pack(side="left")
    EntryCls = ctk.CTkEntry if app.clientes_use_custom else ttk.Entry
    if app.clientes_use_custom and ctk is not None:
        def BtnCls(parent, **kwargs):
            kwargs.setdefault("height", 34)
            kwargs.setdefault("corner_radius", 10)
            kwargs.setdefault("fg_color", primary)
            kwargs.setdefault("hover_color", primary_hover)
            kwargs.setdefault("text_color", "#ffffff")
            kwargs.setdefault("border_width", 0)
            return ctk.CTkButton(parent, **kwargs)
        orange_btn = {"fg_color": "#f59e0b", "hover_color": "#d97706"}
    else:
        BtnCls = ttk.Button
        orange_btn = {}
    app.c_filter = StringVar()
    app.c_filter_entry = EntryCls(filter_row, textvariable=app.c_filter, width=260 if app.clientes_use_custom else 40)
    app.c_filter_entry.pack(side="left", padx=6, pady=2)
    app.c_filter_entry.bind("<KeyRelease>", lambda _: app.refresh_clientes())

    frm = top_frame
    app.c_nome = StringVar()
    app.c_codigo = StringVar()
    app.c_nif = StringVar()
    app.c_morada = StringVar()
    app.c_contacto = StringVar()
    app.c_email = StringVar()
    app.c_obs = StringVar()
    app.c_prazo = StringVar()
    app.c_pagamento = StringVar()
    app.c_obs_tec = StringVar()

    for i, (lbl, var) in enumerate([
        ("Nome", app.c_nome),
        ("Código", app.c_codigo),
        ("NIF", app.c_nif),
        ("Morada", app.c_morada),
        ("Contacto", app.c_contacto),
        ("Email", app.c_email),
        ("Observações", app.c_obs),
        ("Prazo entrega (dias)", app.c_prazo),
        ("Condições de Pagamento", app.c_pagamento),
        ("Obs. Técnicas", app.c_obs_tec),
    ]):
        (ctk.CTkLabel if app.clientes_use_custom else ttk.Label)(frm, text=lbl).grid(row=i // 2, column=(i % 2) * 2, sticky="w", padx=4, pady=2)
        if lbl == "Código":
            entry = EntryCls(frm, textvariable=var, width=260 if app.clientes_use_custom else 40, state="readonly")
        elif lbl == "Condições de Pagamento":
            if app.clientes_use_custom:
                entry = ctk.CTkComboBox(frm, variable=var, values=cond_pagamento_opcoes, state="readonly", width=260)
            else:
                entry = ttk.Combobox(frm, textvariable=var, values=cond_pagamento_opcoes, state="readonly", width=37)
        else:
            entry = EntryCls(frm, textvariable=var, width=260 if app.clientes_use_custom else 40)
        entry.grid(row=i // 2, column=(i % 2) * 2 + 1, sticky="w", padx=4, pady=2)

    btns = ctk.CTkFrame(
        parent_c,
        fg_color="white",
        corner_radius=10,
        border_width=1,
        border_color="#e7cfd3",
    ) if app.clientes_use_custom else ttk.Frame(parent_c)
    btns.pack(fill="x", padx=10, pady=4)
    BtnCls(btns, text="Adicionar", command=app.add_cliente, **orange_btn).pack(side="left", padx=4, pady=2)
    BtnCls(btns, text="Atualizar", command=app.edit_cliente).pack(side="left", padx=4, pady=2)
    BtnCls(btns, text="Remover", command=app.remove_cliente).pack(side="left", padx=4, pady=2)
    BtnCls(btns, text="Exportar CSV", command=lambda: app.export_csv("clientes")).pack(side="left", padx=4, pady=2)

    tbl_style = ""
    if app.clientes_use_custom:
        style = ttk.Style()
        style.configure(
            "Clientes.Treeview",
            font=("Segoe UI", 10),
            rowheight=27,
            background="#f8fbff",
            fieldbackground="#f8fbff",
            borderwidth=0,
        )
        style.configure(
            "Clientes.Treeview.Heading",
            font=("Segoe UI", 10, "bold"),
            background=theme_header_bg,
            foreground="white",
            relief="flat",
        )
        style.map("Clientes.Treeview.Heading", background=[("active", theme_header_active)])
        style.map(
            "Clientes.Treeview",
            background=[("selected", theme_select_bg)],
            foreground=[("selected", theme_select_fg)],
        )
        tbl_style = "Clientes.Treeview"
    if app.clientes_use_full_custom:
        app.tbl_clientes = None
        tbl_wrap = ctk.CTkFrame(parent_c, fg_color="#fff8f9", corner_radius=10)
        tbl_wrap.pack(fill="both", expand=True, padx=10, pady=10)
        head = ctk.CTkFrame(tbl_wrap, fg_color="#9b2233", corner_radius=8)
        head.pack(fill="x", padx=8, pady=(8, 4))
        ctk.CTkLabel(
            head,
            text="CÓDIGO | NOME | NIF | CONTACTO | EMAIL | CONDIÇÕES",
            font=("Segoe UI", 11, "bold"),
            text_color="white",
            anchor="w",
        ).pack(fill="x", padx=10, pady=7)
        app.clientes_list = ctk.CTkScrollableFrame(tbl_wrap, fg_color="#f7f8fb", corner_radius=8)
        app.clientes_list.pack(fill="both", expand=True, padx=8, pady=(0, 8))
    else:
        tbl_wrap = ctk.CTkFrame(
            parent_c,
            fg_color="white",
            corner_radius=10,
            border_width=1,
            border_color="#e7cfd3",
        ) if app.clientes_use_custom else ttk.Frame(parent_c)
        tbl_wrap.pack(fill="both", expand=True, padx=10, pady=10)

        app.tbl_clientes = ttk.Treeview(
            tbl_wrap,
            columns=("codigo", "nome", "nif", "morada", "contacto", "email", "obs", "prazo", "pagamento", "obs_tec"),
            show="headings",
            style=tbl_style,
        )
        app.tbl_clientes.heading("codigo", text="Código", command=lambda c="codigo": app.sort_treeview(app.tbl_clientes, c, False))
        app.tbl_clientes.heading("nome", text="Nome", command=lambda c="nome": app.sort_treeview(app.tbl_clientes, c, False))
        app.tbl_clientes.heading("nif", text="NIF", command=lambda c="nif": app.sort_treeview(app.tbl_clientes, c, False))
        app.tbl_clientes.heading("morada", text="Morada", command=lambda c="morada": app.sort_treeview(app.tbl_clientes, c, False))
        app.tbl_clientes.heading("contacto", text="Contacto", command=lambda c="contacto": app.sort_treeview(app.tbl_clientes, c, False))
        app.tbl_clientes.heading("email", text="Email", command=lambda c="email": app.sort_treeview(app.tbl_clientes, c, False))
        app.tbl_clientes.heading("obs", text="Observações", command=lambda c="obs": app.sort_treeview(app.tbl_clientes, c, False))
        app.tbl_clientes.heading("prazo", text="Prazo Entrega", command=lambda c="prazo": app.sort_treeview(app.tbl_clientes, c, False))
        app.tbl_clientes.heading("pagamento", text="Condições de Pagamento", command=lambda c="pagamento": app.sort_treeview(app.tbl_clientes, c, False))
        app.tbl_clientes.heading("obs_tec", text="Obs. Técnicas", command=lambda c="obs_tec": app.sort_treeview(app.tbl_clientes, c, False))
        app.tbl_clientes.column("codigo", width=100, anchor="center")
        app.tbl_clientes.column("nome", width=220, anchor="w")
        app.tbl_clientes.column("nif", width=115, anchor="center")
        app.tbl_clientes.column("morada", width=220, anchor="w")
        app.tbl_clientes.column("contacto", width=120, anchor="center")
        app.tbl_clientes.column("email", width=180, anchor="w")
        app.tbl_clientes.column("obs", width=160, anchor="w")
        app.tbl_clientes.column("prazo", width=110, anchor="center")
        app.tbl_clientes.column("pagamento", width=190, anchor="w")
        app.tbl_clientes.column("obs_tec", width=180, anchor="w")
        app.tbl_clientes["displaycolumns"] = (
            "codigo",
            "nome",
            "nif",
            "contacto",
            "email",
            "pagamento",
            "prazo",
            "morada",
            "obs",
            "obs_tec",
        )
        app.tbl_clientes.tag_configure("even", background="#eef4fb")
        app.tbl_clientes.tag_configure("odd", background="#f8fbff")
        if app.clientes_use_custom and ctk is not None:
            vsb = ctk.CTkScrollbar(tbl_wrap, orientation="vertical", command=app.tbl_clientes.yview)
            hsb = ctk.CTkScrollbar(tbl_wrap, orientation="horizontal", command=app.tbl_clientes.xview)
        else:
            vsb = ttk.Scrollbar(tbl_wrap, orient="vertical", command=app.tbl_clientes.yview)
            hsb = ttk.Scrollbar(tbl_wrap, orient="horizontal", command=app.tbl_clientes.xview)
        app.tbl_clientes.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        app.tbl_clientes.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        app.tbl_clientes.bind("<<TreeviewSelect>>", app.on_select_cliente)
    app.refresh_clientes()
