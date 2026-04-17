from module_context import configure_module, ensure_module

_CONFIGURED = False


def configure(main_globals):
    configure_module(globals(), main_globals)


def _ensure_configured():
    ensure_module(globals(), "clientes_actions")

def on_select_cliente(self, _):
    _ensure_configured()
    sel = self.tbl_clientes.selection()
    if not sel:
        return
    values = self.tbl_clientes.item(sel[0], "values")
    self._set_cliente_form_values(values)

def _set_cliente_form_values(self, values):
    _ensure_configured()
    self.c_codigo.set(values[0] if len(values) > 0 else "")
    self.c_nome.set(values[1] if len(values) > 1 else "")
    self.c_nif.set(values[2] if len(values) > 2 else "")
    self.c_morada.set(values[3] if len(values) > 3 else "")
    self.c_contacto.set(values[4] if len(values) > 4 else "")
    self.c_email.set(values[5] if len(values) > 5 else "")
    self.c_obs.set(values[6] if len(values) > 6 else "")
    self.c_prazo.set(values[7] if len(values) > 7 else "")
    self.c_pagamento.set(values[8] if len(values) > 8 else "")
    self.c_obs_tec.set(values[9] if len(values) > 9 else "")
    self.selected_cliente_codigo = self.c_codigo.get().strip()

def _on_select_cliente_custom(self, codigo):
    _ensure_configured()
    cli = find_cliente(self.data, codigo)
    if not cli:
        return
    values = (
        cli.get("codigo", ""), cli.get("nome", ""), cli.get("nif", ""), cli.get("morada", ""),
        cli.get("contacto", ""), cli.get("email", ""), cli.get("obs", ""),
        cli.get("prazo_entrega", ""), cli.get("cond_pagamento", ""), cli.get("obs_tecnicas", "")
    )
    self._set_cliente_form_values(values)
    self.refresh_clientes()

def add_cliente(self):
    _ensure_configured()
    codigo = next_cliente_codigo(self.data)
    if not codigo:
        messagebox.showerror("Erro", "Código obrigatório")
        return
    if find_cliente(self.data, codigo):
        messagebox.showerror("Erro", "Código já existe")
        return
    self.c_codigo.set(codigo)
    self.data["clientes"].append({
        "codigo": codigo,
        "nome": self.c_nome.get().strip(),
        "nif": self.c_nif.get().strip(),
        "morada": self.c_morada.get().strip(),
        "contacto": self.c_contacto.get().strip(),
        "email": self.c_email.get().strip(),
        "obs": self.c_obs.get().strip(),
        "prazo_entrega": self.c_prazo.get().strip(),
        "cond_pagamento": self.c_pagamento.get().strip(),
        "obs_tecnicas": self.c_obs_tec.get().strip(),
        "criado_em": now_iso(),
    })
    save_data(self.data)
    self.refresh()

def edit_cliente(self):
    _ensure_configured()
    codigo = self.c_codigo.get().strip()
    c = find_cliente(self.data, codigo)
    if not c:
        messagebox.showerror("Erro", "Cliente nao encontrado")
        return
    c.update({
        "nome": self.c_nome.get().strip(),
        "nif": self.c_nif.get().strip(),
        "morada": self.c_morada.get().strip(),
        "contacto": self.c_contacto.get().strip(),
        "email": self.c_email.get().strip(),
        "obs": self.c_obs.get().strip(),
        "prazo_entrega": self.c_prazo.get().strip(),
        "cond_pagamento": self.c_pagamento.get().strip(),
        "obs_tecnicas": self.c_obs_tec.get().strip(),
    })
    save_data(self.data)
    self.refresh()

def remove_cliente(self):
    _ensure_configured()
    codigo = self.c_codigo.get().strip()
    c = find_cliente(self.data, codigo)
    if not c:
        messagebox.showerror("Erro", "Cliente nao encontrado")
        return
    self.data["clientes"].remove(c)
    save_data(self.data)
    self.refresh()

def refresh_clientes(self):
    _ensure_configured()
    query = (self.c_filter.get().strip().lower() if hasattr(self, "c_filter") else "")
    if self.clientes_use_full_custom and hasattr(self, "clientes_list"):
        for w in self.clientes_list.winfo_children():
            w.destroy()
        selected = (self.c_codigo.get().strip() or self.selected_cliente_codigo).strip()
        self._clientes_btns = {}
        row_i = 0
        for c in self.data["clientes"]:
            values = (
                c.get("codigo", ""), c.get("nome", ""), c.get("nif", ""), c.get("morada", ""),
                c.get("contacto", ""), c.get("email", ""), c.get("obs", ""),
                c.get("prazo_entrega", ""), c.get("cond_pagamento", ""), c.get("obs_tecnicas", "")
            )
            if query and not any(query in str(v).lower() for v in values):
                continue
            codigo = c.get("codigo", "")
            is_sel = bool(selected and codigo == selected)
            base_bg = THEME_SELECT_BG if is_sel else ("#f8fbff" if row_i % 2 == 0 else "#eef4fb")
            row = ctk.CTkFrame(self.clientes_list, fg_color=base_bg, corner_radius=7)
            row.pack(fill="x", padx=2, pady=2)
            txt = f"{codigo}   |   {c.get('nome','')}   |   {c.get('nif','')}   |   {c.get('contacto','')}   |   {c.get('email','')}   |   {c.get('cond_pagamento','')}"
            btn = ctk.CTkButton(
                row,
                text=txt,
                anchor="w",
                fg_color=base_bg,
                hover_color=THEME_SELECT_BG,
                text_color="black",
                height=34,
                command=lambda cod=codigo: self._on_select_cliente_custom(cod),
            )
            btn.pack(fill="x", padx=4, pady=4)
            self._clientes_btns[codigo] = btn
            row_i += 1
    else:
        children = self.tbl_clientes.get_children()
        if children:
            self.tbl_clientes.delete(*children)
        for idx, c in enumerate(self.data["clientes"]):
            values = (
                c.get("codigo", ""), c.get("nome", ""), c.get("nif", ""), c.get("morada", ""),
                c.get("contacto", ""), c.get("email", ""), c.get("obs", ""),
                c.get("prazo_entrega", ""), c.get("cond_pagamento", ""), c.get("obs_tecnicas", "")
            )
            if query and not any(query in str(v).lower() for v in values):
                continue
            tag = "odd" if idx % 2 else "even"
            self.tbl_clientes.insert("", END, values=values, tags=(tag,))
    if hasattr(self, "orc_cliente_cb"):
        vals = self.get_clientes_display()
        try:
            if CUSTOM_TK_AVAILABLE and isinstance(self.orc_cliente_cb, ctk.CTkComboBox):
                self.orc_cliente_cb.configure(values=vals or [""])
            else:
                self.orc_cliente_cb["values"] = vals
        except Exception:
            try:
                self.orc_cliente_cb["values"] = vals
            except Exception:
                pass

def get_clientes_display(self):
    _ensure_configured()
    out = []
    for c in self.data["clientes"]:
        codigo = c.get("codigo", "")
        nome = c.get("nome", "")
        if nome:
            out.append(f"{codigo} - {nome}")
        else:
            out.append(codigo)
    return out

