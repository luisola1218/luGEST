import math

from module_context import configure_module, ensure_module

_CONFIGURED = False

def configure(main_globals):
    configure_module(globals(), main_globals)

def _ensure_configured():
    ensure_module(globals(), "plan_actions")

PLANO_PDF_CORES = [
    "#ffb74d",
    "#4fc3f7",
    "#81c784",
    "#f48fb1",
    "#ba68c8",
    "#ffd54f",
    "#4db6ac",
    "#90caf9",
    "#ff8a65",
    "#a5d6a7",
    "#ce93d8",
    "#80cbc4",
]


def _find_esp_obj_for_item(enc, item):
    if not enc:
        return None
    mat_item = str(item.get("material", "") or "").strip()
    esp_item = fmt_num(item.get("espessura", ""))
    for m in list(enc.get("materiais", []) or []):
        if str(m.get("material", "") or "").strip() != mat_item:
            continue
        for esp_obj in list(m.get("espessuras", []) or []):
            if fmt_num(esp_obj.get("espessura", "")) == esp_item:
                return esp_obj
    return None


def _laser_done_for_item(self, item):
    enc_num = str(item.get("encomenda", "") or "").strip()
    if not enc_num:
        return False
    enc = self.get_encomenda_by_numero(enc_num)
    esp_obj = _find_esp_obj_for_item(enc, item)
    if not esp_obj:
        return False
    if bool(esp_obj.get("laser_concluido")):
        return True
    pecas = list(esp_obj.get("pecas", []) or [])
    if not pecas:
        return False
    has_laser = False
    all_done = True
    for p in pecas:
        fluxo = list(ensure_peca_operacoes(p) or [])
        p_has_laser = False
        p_laser_done = False
        for op in fluxo:
            nome = normalize_operacao_nome(op.get("nome", ""))
            if "laser" in norm_text(nome):
                p_has_laser = True
                if "concl" in norm_text(op.get("estado", "")):
                    p_laser_done = True
                break
        if p_has_laser:
            has_laser = True
            if not p_laser_done:
                all_done = False
    if has_laser and all_done:
        esp_obj["laser_concluido"] = True
        esp_obj["laser_concluido_em"] = now_iso()
        return True
    return False


def _pdf_block_color_for_item(item):
    # Cor por bloco (encomenda/material/espessura/data/hora) para diferenciar melhor quadrados.
    key = "|".join(
        [
            str(item.get("encomenda", "") or "").strip().upper(),
            str(item.get("material", "") or "").strip().upper(),
            str(item.get("espessura", "") or "").strip().upper(),
            str(item.get("data", "") or "").strip().upper(),
            str(item.get("inicio", "") or "").strip().upper(),
        ]
    )
    if not key.strip("|"):
        key = str(item or "")
    idx = sum(ord(ch) for ch in key) % len(PLANO_PDF_CORES)
    return PLANO_PDF_CORES[idx]

def on_shortcut_new(self, _):
    _ensure_configured()
    tab = self.nb.select()
    if tab == str(self.tab_clientes):
        self.add_cliente()
    elif tab == str(self.tab_materia):
        self.add_material()
    elif tab == str(self.tab_encomendas):
        self.add_encomenda()

def pick_date(self, target_var, parent=None):
    _ensure_configured()
    use_custom = CUSTOM_TK_AVAILABLE and (
        getattr(self, "encomendas_use_custom", False)
        or getattr(self, "plano_use_custom", False)
        or getattr(self, "orc_use_custom", False)
        or getattr(self, "ne_use_custom", False)
    )
    parent_win = parent
    if parent_win is None:
        try:
            fw = self.root.focus_get()
            parent_win = fw.winfo_toplevel() if fw is not None else None
        except Exception:
            parent_win = None
    if parent_win is None:
        parent_win = self.root
    win = ctk.CTkToplevel(parent_win) if use_custom else Toplevel(parent_win)
    win.title("Calendário")
    win.geometry("360x390")
    try:
        win.transient(parent_win)
        win.grab_set()
    except Exception:
        pass

    today = datetime.now().date()
    selected_day = IntVar(value=0)
    selected_text = StringVar(value="")
    year_var = StringVar(value=str(today.year))
    month_var = StringVar(value=f"{today.month:02d}")
    cur_val = (target_var.get() or "").strip() if hasattr(target_var, "get") else ""
    try:
        cur_dt = datetime.strptime(cur_val, "%Y-%m-%d").date() if cur_val else None
    except Exception:
        cur_dt = None
    if cur_dt:
        year_var.set(str(cur_dt.year))
        month_var.set(f"{cur_dt.month:02d}")
        selected_day.set(cur_dt.day)
        selected_text.set(cur_dt.strftime("%Y-%m-%d"))

    def get_ym():
        try:
            y = int(year_var.get())
        except Exception:
            y = today.year
        try:
            m = int(month_var.get())
        except Exception:
            m = today.month
        m = max(1, min(12, m))
        return y, m

    header = ctk.CTkFrame(win, fg_color="transparent") if use_custom else ttk.Frame(win)
    header.pack(fill="x", padx=10, pady=6)
    if use_custom:
        ctk.CTkButton(header, text="◀", width=40, command=lambda: change_month(-1)).pack(side="left")
        month_cb = ctk.CTkComboBox(
            header,
            variable=month_var,
            values=[f"{i:02d}" for i in range(1, 13)],
            width=78,
            command=lambda _v: draw_calendar(),
        )
        month_cb.pack(side="left", padx=6)
        year_cb = ctk.CTkComboBox(
            header,
            variable=year_var,
            values=[str(y) for y in range(today.year - 2, today.year + 6)],
            width=96,
            command=lambda _v: draw_calendar(),
        )
        year_cb.pack(side="left")
        ctk.CTkButton(header, text="▶", width=40, command=lambda: change_month(1)).pack(side="left", padx=6)
    else:
        ttk.Button(header, text="<", command=lambda: change_month(-1)).pack(side="left")
        month_cb = ttk.Combobox(header, values=[f"{i:02d}" for i in range(1, 13)], textvariable=month_var, width=3, state="readonly")
        month_cb.pack(side="left", padx=6)
        year_cb = ttk.Combobox(header, values=[str(y) for y in range(today.year - 2, today.year + 6)], textvariable=year_var, width=5, state="readonly")
        year_cb.pack(side="left")
        ttk.Button(header, text=">", command=lambda: change_month(1)).pack(side="left", padx=6)

    grid = ctk.CTkFrame(win, fg_color="transparent") if use_custom else ttk.Frame(win)
    grid.pack(fill="both", expand=True, padx=10, pady=(0, 6))

    days_lbl = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sab", "Dom"]
    for i, d in enumerate(days_lbl):
        if use_custom:
            ctk.CTkLabel(grid, text=d, font=("Segoe UI", 12, "bold")).grid(row=0, column=i, padx=2, pady=2)
        else:
            ttk.Label(grid, text=d).grid(row=0, column=i, padx=2, pady=2)

    def change_month(delta):
        y, m = get_ym()
        m += delta
        if m < 1:
            m = 12
            y -= 1
        elif m > 12:
            m = 1
            y += 1
        month_var.set(f"{m:02d}")
        year_var.set(str(y))
        draw_calendar()

    def pick_day(day):
        if day <= 0:
            return
        y, m = get_ym()
        dt = datetime(y, m, day).date()
        selected_day.set(day)
        selected_text.set(dt.strftime("%Y-%m-%d"))

    def apply_date():
        d = int(selected_day.get() or 0)
        if d <= 0:
            messagebox.showerror("Erro", "Selecione um dia no calendário.")
            return
        y, m = get_ym()
        try:
            dt = datetime(y, m, d).date()
        except Exception:
            messagebox.showerror("Erro", "Data inválida.")
            return
        target_var.set(dt.strftime("%Y-%m-%d"))
        win.destroy()

    def set_today():
        year_var.set(str(today.year))
        month_var.set(f"{today.month:02d}")
        selected_day.set(today.day)
        selected_text.set(today.strftime("%Y-%m-%d"))
        draw_calendar()

    def draw_calendar():
        for w in grid.winfo_children():
            if int(w.grid_info().get("row", 0)) > 0:
                w.destroy()
        y, m = get_ym()
        cal = calendar.monthcalendar(y, m)
        sel = int(selected_day.get() or 0)
        for r, week in enumerate(cal, start=1):
            for c, day in enumerate(week):
                if day == 0:
                    if use_custom:
                        ctk.CTkLabel(grid, text="", width=36).grid(row=r, column=c, padx=2, pady=2)
                    else:
                        ttk.Label(grid, text="").grid(row=r, column=c, padx=2, pady=2)
                else:
                    is_weekend = c in (5, 6)
                    is_selected = (day == sel)
                    if use_custom:
                        if is_selected:
                            fg = "#ba2d3d"
                            hover = "#a32035"
                            txt_color = "#ffffff"
                        else:
                            fg = "#f9d9d9" if is_weekend else "#fbecee"
                            hover = "#f2c5c5" if is_weekend else "#f8dfe3"
                            txt_color = "#1f2937"
                        ctk.CTkButton(
                            grid,
                            text=str(day),
                            width=36,
                            height=30,
                            corner_radius=7,
                            fg_color=fg,
                            hover_color=hover,
                            text_color=txt_color,
                            command=lambda d=day: pick_day(d),
                        ).grid(row=r, column=c, padx=2, pady=2)
                    else:
                        bg = "#ba2d3d" if is_selected else ("#ffe5e5" if is_weekend else "#fff5f6")
                        Button(grid, text=str(day), width=3, command=lambda d=day: pick_day(d), bg=bg).grid(row=r, column=c, padx=2, pady=2)

    if not use_custom:
        month_cb.bind("<<ComboboxSelected>>", lambda _: draw_calendar())
        year_cb.bind("<<ComboboxSelected>>", lambda _: draw_calendar())
    draw_calendar()

    status = ctk.CTkFrame(win, fg_color="transparent") if use_custom else ttk.Frame(win)
    status.pack(fill="x", padx=10, pady=(2, 0))
    if use_custom:
        ctk.CTkLabel(status, text="Selecionada:").pack(side="left")
        ctk.CTkLabel(status, textvariable=selected_text, text_color="#7a0f1a").pack(side="left", padx=6)
    else:
        ttk.Label(status, text="Selecionada:").pack(side="left")
        ttk.Label(status, textvariable=selected_text).pack(side="left", padx=6)

    btns = ctk.CTkFrame(win, fg_color="transparent") if use_custom else ttk.Frame(win)
    btns.pack(pady=6, fill="x", padx=10)
    if use_custom:
        ctk.CTkButton(btns, text="Hoje", command=set_today, width=110).pack(side="left", padx=6)
        ctk.CTkButton(btns, text="Confirmar", command=apply_date, width=110, fg_color="#ba2d3d", hover_color="#a32035").pack(side="left", padx=6)
        ctk.CTkButton(btns, text="Fechar", command=win.destroy, width=110, fg_color="#6b7280", hover_color="#4b5563").pack(side="left", padx=6)
    else:
        ttk.Button(btns, text="Hoje", command=set_today).pack(side="left", padx=6)
        ttk.Button(btns, text="Confirmar", command=apply_date).pack(side="left", padx=6)
        ttk.Button(btns, text="Fechar", command=win.destroy).pack(side="left", padx=6)

def refresh_plano(self):
    _ensure_configured()
    self._sync_plano_hist()
    children = self.tbl_plano_enc.get_children()
    if children:
        self.tbl_plano_enc.delete(*children)
    filtro = ""
    if hasattr(self, "p_filter") and self.p_filter.get():
        filtro = self.p_filter.get().strip().lower()
    estado_filtro = norm_text(self.p_status_filter.get() if hasattr(self, "p_status_filter") else "pendentes")
    if hasattr(self, "p_status_segment"):
        try:
            self.p_status_segment.set(self.p_status_filter.get() or "Pendentes")
        except Exception:
            pass
    planeado = {(p.get("encomenda"), p.get("material"), p.get("espessura")) for p in self.data.get("plano", [])}
    row_i = 0
    for e in self.data["encomendas"]:
        enc_estado_norm = norm_text(e.get("estado", ""))
        # Em modo pendentes, nunca mostrar encomendas já concluídas/canceladas.
        if estado_filtro.startswith("pend") and ("concl" in enc_estado_norm or "cancel" in enc_estado_norm):
            continue
        for m in e.get("materiais", []):
            for esp_obj in m.get("espessuras", []):
                esp = esp_obj.get("espessura", "")
                laser_done = _laser_done_for_item(
                    self,
                    {"encomenda": e.get("numero", ""), "material": m.get("material", ""), "espessura": esp},
                )
                if estado_filtro.startswith("pend") and laser_done:
                    continue
                if estado_filtro.startswith("concl") and (not laser_done):
                    continue
                # Em lista pendente/todas, esconder linhas já alocadas em blocos ativos.
                if estado_filtro != "concluidas" and (e["numero"], f"{m.get('material','')}", esp) in planeado:
                    continue
                if filtro:
                    hay = " ".join([
                        str(e.get("numero", "")),
                        str(m.get("material", "")),
                        str(esp),
                        str(esp_obj.get("tempo_min", "")),
                        str(e.get("estado", "")),
                    ]).lower()
                    if filtro not in hay:
                        continue
                tempo = esp_obj.get("tempo_min", "")
                tag = "pl_even" if row_i % 2 == 0 else "pl_odd"
                if not tempo:
                    tag = "pl_warn"
                self.tbl_plano_enc.insert("", END, values=(e["numero"], m.get("material",""), esp, tempo), tags=(tag,))
                row_i += 1

    # CTkLabel não implementa .config; usar .configure para ambos
    self.p_week_lbl.configure(text=self.p_week_start.strftime("%d/%m/%Y"))
    start_min, end_min, slot, rows, cols, col_w, row_h = self.get_plano_grid_metrics()

    self.plano.delete("all")
    dias = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sab"]
    dates = [self.p_week_start + timedelta(days=i) for i in range(cols)]
    dia_cores = ["#fde2e4", "#fbecee", "#fff0f2", "#ffe6e9", "#f9e8ea", "#fff5f6"]
    for c in range(cols):
        x0 = (c + 1) * col_w
        self.plano.create_rectangle(x0, 0, x0 + col_w, row_h, fill=dia_cores[c % len(dia_cores)], outline="#e7cfd3")
        self.plano.create_text(x0 + col_w // 2, row_h // 2, text=f"{dias[c]} {dates[c].strftime('%d/%m')}")

    for r in range(rows):
        y0 = (r + 1) * row_h
        slot_start = start_min + r * slot
        slot_end = slot_start + slot
        t = minutes_to_time(slot_start)
        self.plano.create_text(col_w // 2, y0 + row_h // 2, text=t)
        for c in range(cols):
            x0 = (c + 1) * col_w
            fill = "#f5f5f5" if self.plano_intervalo_bloqueado(slot_start, slot_end) else "white"
            self.plano.create_rectangle(x0, y0, x0 + col_w, y0 + row_h, outline="#e7cfd3", fill=fill)

    def _hex_to_rgb01(hx):
        try:
            hh = str(hx or "").strip().lstrip("#")
            if len(hh) != 6:
                return (0.82, 0.90, 0.98)
            return (int(hh[0:2], 16) / 255.0, int(hh[2:4], 16) / 255.0, int(hh[4:6], 16) / 255.0)
        except Exception:
            return (0.82, 0.90, 0.98)

    def _rgb01_to_hex(rgb):
        r, g, b = rgb
        return f"#{int(max(0, min(1, r))*255):02x}{int(max(0, min(1, g))*255):02x}{int(max(0, min(1, b))*255):02x}"

    def _blend_white(rgb, alpha=0.30):
        r, g, b = rgb
        return (r + (1 - r) * alpha, g + (1 - g) * alpha, b + (1 - b) * alpha)

    def _darken(rgb, factor=0.65):
        r, g, b = rgb
        return (max(0, r * factor), max(0, g * factor), max(0, b * factor))

    for item in self.data.get("plano", []):
        try:
            d = datetime.strptime(item["data"], "%Y-%m-%d").date()
        except Exception:
            continue
        if d < dates[0] or d > dates[-1]:
            continue
        col = (d - dates[0]).days
        start = time_to_minutes(item["inicio"])
        dur = int(item["duracao_min"])
        r_start = (start - start_min) // slot
        r_span = max(1, int(math.ceil(float(dur) / float(slot))))
        if r_start < 0 or r_start >= rows:
            continue
        x0 = (col + 1) * col_w
        y0 = (r_start + 1) * row_h
        y1 = y0 + r_span * row_h
        base_rgb = _hex_to_rgb01(_pdf_block_color_for_item(item))
        fill_hex = _rgb01_to_hex(_blend_white(base_rgb, 0.30))
        edge_hex = _rgb01_to_hex(_darken(base_rgb, 0.65))
        rect = self.plano.create_rectangle(x0 + 2, y0 + 2, x0 + col_w - 2, y1 - 2, fill=fill_hex, outline=edge_hex, width=2)
        self.plano.create_rectangle(x0 + 2, y0 + 2, x0 + 8, y1 - 2, fill=_rgb01_to_hex(base_rgb), outline="")

        enc_num = str(item.get("encomenda", "") or "").strip()
        mat = str(item.get("material", "") or "").strip()
        esp = str(item.get("espessura", "") or "").strip()
        mat_esp = " | ".join(x for x in [mat, f"{esp} mm" if esp else ""] if x).strip()
        box_h = max(1, r_span * row_h)
        if box_h <= (row_h * 1.2):
            texto = f"{enc_num} | {mat_esp}" if mat_esp else enc_num
            font_size = 8
        else:
            linhas = [enc_num]
            if mat:
                linhas.append(mat)
            if esp:
                linhas.append(f"{esp} mm")
            texto = "\n".join([ln for ln in linhas if ln])
            font_size = 9 if box_h >= (row_h * 2) else 8
        self.plano.create_text(
            x0 + col_w // 2,
            y0 + (box_h / 2),
            text=texto,
            width=col_w - 14,
            justify="center",
            font=("Segoe UI", font_size),
        )
        self.plano.tag_bind(rect, "<Button-1>", lambda e, it=item: self.on_plano_block(it))

def prev_week(self):
    _ensure_configured()
    self.p_week_start = self.p_week_start - timedelta(days=7)
    self.refresh_plano()

def next_week(self):
    _ensure_configured()
    self.p_week_start = self.p_week_start + timedelta(days=7)
    self.refresh_plano()

def on_plano_click(self, event):
    _ensure_configured()
    numero = None
    espessura = None
    tempo_def = None
    if hasattr(self, "plano_drag_item") and self.plano_drag_item:
        numero = self.plano_drag_item.get("encomenda")
        material = self.plano_drag_item.get("material")
        espessura = self.plano_drag_item.get("espessura")
        tempo_def = self.plano_drag_item.get("tempo")
    else:
        selection = self.tbl_plano_enc.selection()
        if selection:
            vals = self.tbl_plano_enc.item(selection[0], "values")
            numero = vals[0]
            material = vals[1]
            espessura = vals[2]
            tempo_def = vals[3]

    try:
        start_min = time_to_minutes(self.p_inicio.get())
        end_min = time_to_minutes(self.p_fim.get())
    except Exception:
        return
    slot = 30
    rows = (end_min - start_min) // slot
    cols = 6
    w = max(self.plano.winfo_width(), 800)
    h = max(self.plano.winfo_height(), 500)
    col_w = w // (cols + 1)
    row_h = max(20, h // (rows + 1))

    col = (event.x // col_w) - 1
    row = (event.y // row_h) - 1
    if col < 0 or col >= cols or row < 0 or row >= rows:
        return

    inicio = minutes_to_time(start_min + row * slot)
    data = (self.p_week_start + timedelta(days=col)).strftime("%Y-%m-%d")

    if hasattr(self, "plano_selected") and self.plano_selected:
        item = self.plano_selected
        try:
            dur_sel = int(item.get("duracao_min", 0))
        except Exception:
            dur_sel = 0
        s_try = start_min + row * slot
        if dur_sel > 0 and self.plano_intervalo_bloqueado(s_try, s_try + dur_sel):
            messagebox.showerror("Erro", "Horário de almoço (12:30-14:00) bloqueado para planeamento.")
            return
        item["data"] = data
        item["inicio"] = inicio
        item["chapa"] = self.get_chapa_reservada(item["encomenda"])
        self.plano_selected = None
        save_data(self.data)
        self.refresh_plano()
        return

    if not numero:
        return
    if self.is_plano_duplicado(numero, material, espessura):
        messagebox.showerror("Erro", "Esta encomenda ja esta planeada")
        return

    if not tempo_def:
        messagebox.showerror("Erro", "Encomenda sem tempo associado")
        return
    default_dur = str(tempo_def)
    prompt = "Duracao em minutos (multiplo de 30)"
    if default_dur:
        prompt += f" [default {default_dur}]"
    dur = simple_input(self.root, "Duracao", prompt)
    if not dur and default_dur:
        dur = default_dur
    if not dur:
        return
    try:
        dur = int(dur)
    except ValueError:
        messagebox.showerror("Erro", "Duracao invalida")
        return
    if dur % 30 != 0:
        messagebox.showerror("Erro", "Duracao deve ser multiplo de 30")
        return
    s_try = start_min + row * slot
    if self.plano_intervalo_bloqueado(s_try, s_try + dur):
        messagebox.showerror("Erro", "Horário de almoço (12:30-14:00) bloqueado para planeamento.")
        return

    self.data.setdefault("plano", []).append({
        "id": f"PL{int(datetime.now().timestamp())}{len(self.data.get('plano', []))}",
        "encomenda": numero,
        "material": material,
        "espessura": espessura,
        "data": data,
        "inicio": inicio,
        "duracao_min": dur,
        "color": PLANO_CORES[len(self.data.get("plano", [])) % len(PLANO_CORES)],
        "chapa": self.get_chapa_reservada(numero),
    })
    save_data(self.data)
    self.refresh_plano()

def on_plano_drop(self, event):
    _ensure_configured()
    if not hasattr(self, "plano_drag_item") or not self.plano_drag_item:
        return
    start_min, end_min, slot, rows, cols, col_w, row_h = self.get_plano_grid_metrics()
    col = (event.x // col_w) - 1
    row = (event.y // row_h) - 1
    if col < 0 or col >= cols or row < 0 or row >= rows:
        return
    inicio = minutes_to_time(start_min + row * slot)
    data = (self.p_week_start + timedelta(days=col)).strftime("%Y-%m-%d")

    numero = self.plano_drag_item.get("encomenda")
    material = self.plano_drag_item.get("material")
    espessura = self.plano_drag_item.get("espessura")
    tempo_def = self.plano_drag_item.get("tempo")
    if self.is_plano_duplicado(numero, material, espessura):
        messagebox.showerror("Erro", "Esta encomenda ja esta planeada")
        self.plano_drag_item = None
        if self.plano_drag_preview:
            self.plano.delete(self.plano_drag_preview)
            self.plano_drag_preview = None
        return
    if not tempo_def:
        messagebox.showerror("Erro", "Encomenda sem tempo associado")
        self.plano_drag_item = None
        if self.plano_drag_preview:
            self.plano.delete(self.plano_drag_preview)
            self.plano_drag_preview = None
        return
    default_dur = str(tempo_def)
    prompt = "Duracao em minutos (multiplo de 30)"
    if default_dur:
        prompt += f" [default {default_dur}]"
    dur = simple_input(self.root, "Duracao", prompt)
    if not dur and default_dur:
        dur = default_dur
    if not dur:
        return
    try:
        dur = int(dur)
    except ValueError:
        messagebox.showerror("Erro", "Duracao invalida")
        return
    if dur % 30 != 0:
        messagebox.showerror("Erro", "Duracao deve ser multiplo de 30")
        return
    s_try = start_min + row * slot
    if self.plano_intervalo_bloqueado(s_try, s_try + dur):
        messagebox.showerror("Erro", "Horário de almoço (12:30-14:00) bloqueado para planeamento.")
        self.plano_drag_item = None
        if self.plano_drag_preview:
            self.plano.delete(self.plano_drag_preview)
            self.plano_drag_preview = None
        return

    self.data.setdefault("plano", []).append({
        "id": f"PL{int(datetime.now().timestamp())}{len(self.data.get('plano', []))}",
        "encomenda": numero,
        "material": material,
        "espessura": espessura,
        "data": data,
        "inicio": inicio,
        "duracao_min": dur,
        "color": PLANO_CORES[len(self.data.get("plano", [])) % len(PLANO_CORES)],
        "chapa": self.get_chapa_reservada(numero),
    })
    save_data(self.data)
    self.plano_drag_item = None
    if self.plano_drag_preview:
        self.plano.delete(self.plano_drag_preview)
        self.plano_drag_preview = None
    self.refresh_plano()

def auto_planear(self):
    _ensure_configured()
    start_min, end_min, slot, rows, cols, col_w, row_h = self.get_plano_grid_metrics()
    dates = [self.p_week_start + timedelta(days=i) for i in range(cols)]

    # Build occupied intervals per day
    occupied = {d.strftime("%Y-%m-%d"): [] for d in dates}
    for item in self.data.get("plano", []):
        d = item.get("data")
        if d not in occupied:
            continue
        try:
            s = time_to_minutes(item["inicio"])
            e = s + int(item["duracao_min"])
        except Exception:
            continue
        occupied[d].append((s, e))

    def is_free(d, s, e):
        if self.plano_intervalo_bloqueado(s, e):
            return False
        for a, b in occupied.get(d, []):
            if not (e <= a or s >= b):
                return False
        return True

    def find_slot(dur):
        for d in dates:
            ds = d.strftime("%Y-%m-%d")
            s = start_min
            while s + dur <= end_min:
                e = s + dur
                if is_free(ds, s, e):
                    return ds, s
                s += slot
        return None, None

    items = list(self.tbl_plano_enc.get_children())
    sel = self.tbl_plano_enc.selection()
    if sel:
        start_idx = items.index(sel[0])
        items = items[start_idx:]

    items = self.select_plano_order(items)
    if items is None:
        return

    cur_day_idx = 0
    cur_min = start_min

    def advance_day():
        nonlocal cur_day_idx, cur_min
        cur_day_idx += 1
        cur_min = start_min

    def find_next_slot(dur):
        nonlocal cur_day_idx, cur_min
        while cur_day_idx < len(dates):
            ds = dates[cur_day_idx].strftime("%Y-%m-%d")
            s = cur_min
            while s + dur <= end_min:
                if is_free(ds, s, s + dur):
                    cur_min = s + dur
                    return ds, s
                s += slot
            advance_day()
        return None, None

    for item_id in items:
        vals = self.tbl_plano_enc.item(item_id, "values")
        numero, material, espessura, tempo_def = vals[0], vals[1], vals[2], vals[3]
        if not tempo_def:
            messagebox.showerror("Erro", f"Encomenda {numero} sem tempo associado")
            return
        try:
            dur = int(float(tempo_def))
        except Exception:
            messagebox.showerror("Erro", f"Tempo invalido em {numero}")
            return
        ds, s = find_next_slot(dur)
        if not ds:
            messagebox.showerror("Erro", "Sem espaco no plano para auto planeamento")
            return
        self.data.setdefault("plano", []).append({
            "id": f"PL{int(datetime.now().timestamp())}{len(self.data.get('plano', []))}",
            "encomenda": numero,
            "material": material,
            "espessura": espessura,
            "data": ds,
            "inicio": minutes_to_time(s),
            "duracao_min": dur,
            "color": PLANO_CORES[len(self.data.get("plano", [])) % len(PLANO_CORES)],
            "chapa": self.get_chapa_reservada(numero),
        })
        occupied[ds].append((s, s + dur))

    save_data(self.data)
    self.refresh_plano()

def select_plano_order(self, items):
    _ensure_configured()
    use_custom = getattr(self, "plano_use_custom", False) and CUSTOM_TK_AVAILABLE
    win = ctk.CTkToplevel(self.root) if use_custom else Toplevel(self.root)
    win.title("Ordem de planeamento")
    win.geometry("620x430")
    if use_custom:
        ctk.CTkLabel(win, text="Clique para adicionar pela ordem desejada:", font=("Segoe UI", 13, "bold")).pack(anchor="w", padx=12, pady=(12, 6))
    else:
        ttk.Label(win, text="Clique para adicionar pela ordem desejada:").pack(anchor="w", padx=10, pady=(10, 4))
    list_frame = ctk.CTkFrame(win, fg_color="transparent") if use_custom else ttk.Frame(win)
    list_frame.pack(fill="both", expand=True, padx=10)

    left_wrap = ctk.CTkFrame(list_frame, fg_color="#f7fafc") if use_custom else ttk.Frame(list_frame)
    right_wrap = ctk.CTkFrame(list_frame, fg_color="#f7fafc") if use_custom else ttk.Frame(list_frame)
    left_wrap.pack(side="left", fill="both", expand=True)
    right_wrap.pack(side="left", fill="both", expand=True, padx=(10, 0))
    if use_custom:
        ctk.CTkLabel(left_wrap, text="Pendentes", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=8, pady=(6, 2))
        ctk.CTkLabel(right_wrap, text="Ordem selecionada", font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=8, pady=(6, 2))
    left = Listbox(left_wrap, height=12, activestyle="none")
    right = Listbox(right_wrap, height=12, activestyle="none")
    left.pack(side="left", fill="both", expand=True, padx=6, pady=(0, 6))
    right.pack(side="left", fill="both", expand=True, padx=6, pady=(0, 6))
    sb_l = ttk.Scrollbar(left_wrap, orient="vertical", command=left.yview)
    sb_r = ttk.Scrollbar(right_wrap, orient="vertical", command=right.yview)
    left.configure(yscrollcommand=sb_l.set)
    right.configure(yscrollcommand=sb_r.set)
    sb_l.pack(side="right", fill="y", pady=(0, 6))
    sb_r.pack(side="right", fill="y", pady=(0, 6))

    id_map = {}
    for item_id in items:
        vals = self.tbl_plano_enc.item(item_id, "values")
        label = f"{vals[0]} | {vals[1]} | {vals[2]} | {vals[3]}"
        left.insert(END, label)
        id_map[label] = item_id

    def add_item(_=None):
        sel = left.curselection()
        if not sel:
            return
        label = left.get(sel[0])
        right.insert(END, label)
        left.delete(sel[0])

    def remove_item():
        sel = right.curselection()
        if not sel:
            return
        label = right.get(sel[0])
        right.delete(sel[0])
        left.insert(END, label)

    left.bind("<Double-Button-1>", add_item)

    btns = ctk.CTkFrame(win, fg_color="transparent") if use_custom else ttk.Frame(win)
    btns.pack(fill="x", padx=10, pady=8)
    Btn = ctk.CTkButton if use_custom else ttk.Button
    Btn(btns, text="Adicionar", command=add_item, width=120 if use_custom else None).pack(side="left", padx=6)
    Btn(btns, text="Remover", command=remove_item, width=120 if use_custom else None).pack(side="left", padx=6)

    result = {"items": None}

    def confirm():
        ordered = []
        for i in range(right.size()):
            label = right.get(i)
            ordered.append(id_map[label])
        result["items"] = ordered
        win.destroy()

    Btn(btns, text="Confirmar", command=confirm, width=120 if use_custom else None).pack(side="right", padx=6)
    if use_custom:
        Btn(btns, text="Cancelar", command=win.destroy, width=120, fg_color="#6b7280", hover_color="#4b5563").pack(side="right", padx=6)
    else:
        Btn(btns, text="Cancelar", command=win.destroy).pack(side="right", padx=6)

    win.grab_set()
    self.root.wait_window(win)
    return result["items"]

def desplanear_tudo(self):
    _ensure_configured()
    if messagebox.askyesno("Desplanear", "Limpar todo o planeamento?"):
        self.data["plano"] = []
        save_data(self.data)
        self.refresh_plano()

def on_plano_block(self, item):
    _ensure_configured()
    enc = self.get_encomenda_by_numero(item["encomenda"])
    if not enc:
        messagebox.showinfo("Info", "Encomenda nao encontrada")
        return
    chapa = self.get_chapa_reservada(enc["numero"])
    materiais = {}
    for p in encomenda_pecas(enc):
        key = f"{p.get('material','')} {p.get('espessura','')}"
        materiais[key] = materiais.get(key, 0) + p.get("quantidade_pedida", 0)
    materiais_str = "\\n".join([f"{k}: {v}" for k, v in materiais.items()]) or "-"
    info = (
        f"Encomenda: {enc['numero']}\\n"
        f"Cliente: {enc['cliente']}\\n"
        f"Data: {item['data']} {item['inicio']}\\n"
        f"Duracao: {item['duracao_min']} min\\n"
        f"Espessura: {item.get('espessura','-')}\\n"
        f"Chapa cativada: {chapa}\\n"
        f"Materiais:\\n{materiais_str}"
    )
    self.plano_selected = item
    if messagebox.askyesno("Detalhes", info + "\\n\\nRemover bloco?"):
        item_id = item.get("id")
        if item_id:
            self.data["plano"] = [p for p in self.data.get("plano", []) if p.get("id") != item_id]
        else:
            self.data["plano"] = [p for p in self.data.get("plano", []) if p is not item]
        save_data(self.data)
        self.refresh_plano()

def _sync_plano_hist(self):
    _ensure_configured()
    plano = list(self.data.get("plano", []))
    if not plano:
        return
    active = []
    moved = False
    hist = self.data.setdefault("plano_hist", [])
    for item in plano:
        enc_num = str(item.get("encomenda", "")).strip()
        enc = self.get_encomenda_by_numero(enc_num) if enc_num else None
        laser_done = _laser_done_for_item(self, item)
        if not laser_done:
            active.append(item)
            continue
        moved = True
        if not any(str(h.get("id", "")) == str(item.get("id", "")) and str(h.get("encomenda", "")) == enc_num for h in hist):
            hrow = dict(item)
            hrow["movido_em"] = now_iso()
            hrow["estado_final"] = "Laser Concluido"
            hrow["tipo_planeamento"] = "Laser"
            hrow["tempo_planeado_min"] = parse_float(item.get("duracao_min", item.get("dur", 0)), 0)
            hrow["tempo_real_min"] = parse_float(
                enc.get("tempo_producao_min", enc.get("tempo_espessuras_min", enc.get("tempo_pecas_min", 0))) if enc else 0,
                0,
            )
            hist.append(hrow)
    if moved:
        self.data["plano"] = active
        save_data(self.data)

def preview_plano_a4(self):
    _ensure_configured()
    path = os.path.join(tempfile.gettempdir(), "lugest_plano.pdf")

    def render_pdf(out_path):
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.pdfgen import canvas as pdf_canvas
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        from reportlab.lib.utils import ImageReader
        width, height = landscape(A4)
        c = pdf_canvas.Canvas(out_path, pagesize=landscape(A4))

        font_regular = "Helvetica"
        font_bold = "Helvetica-Bold"
        try:
            seg = r"C:\Windows\Fonts\segoeui.ttf"
            seg_b = r"C:\Windows\Fonts\segoeuib.ttf"
            if os.path.exists(seg):
                pdfmetrics.registerFont(TTFont("SegoeUI", seg))
                font_regular = "SegoeUI"
            if os.path.exists(seg_b):
                pdfmetrics.registerFont(TTFont("SegoeUI-Bold", seg_b))
                font_bold = "SegoeUI-Bold"
            if font_regular == "Helvetica":
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

        margin = 20
        c.setStrokeColorRGB(0.78, 0.80, 0.84)
        c.rect(margin, yinv(height - margin), width - margin * 2, height - margin * 2, stroke=1, fill=0)
        draw_logo(margin + 4, margin + 6, 70, 32)

        start_min, end_min, slot, rows, cols, col_w, row_h = self.get_plano_grid_metrics()
        top_margin = 60
        time_w = 66
        footer_box_h = 60
        grid_w = width - (margin * 2) - time_w
        col_w = max(60, grid_w // cols)
        grid_h = height - top_margin - margin - footer_box_h - 10
        row_h = max(18, grid_h // (rows + 1))
        dias = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sab"]
        dates = [self.p_week_start + timedelta(days=i) for i in range(cols)]

        set_font(True, 14)
        c.drawString(margin + 82, yinv(margin + 20), "Plano de Produção")
        set_font(False, 9)
        c.drawRightString(width - margin - 4, yinv(margin + 20), f"Semana: {dates[0].strftime('%d/%m/%Y')} - {dates[-1].strftime('%d/%m/%Y')}")
        c.line(margin, yinv(55), width - margin, yinv(55))

        set_font(True, 8)
        for cidx in range(cols):
            x0 = margin + time_w + (cidx * col_w)
            c.setFillColorRGB(0.90, 0.94, 1.0)
            c.rect(x0, yinv(top_margin + row_h), col_w, row_h, stroke=1, fill=1)
            c.setFillColorRGB(0, 0, 0)
            c.drawCentredString(x0 + col_w / 2, yinv(top_margin + row_h / 2), f"{dias[cidx]} {dates[cidx].strftime('%d/%m')}")

        c.setFillColorRGB(0.96, 0.97, 0.99)
        c.rect(margin, yinv(top_margin + (rows + 1) * row_h), time_w, rows * row_h, stroke=1, fill=1)
        set_font(False, 7)
        for r in range(rows):
            y0 = top_margin + (r + 1) * row_h
            slot_start = start_min + r * slot
            slot_end = slot_start + slot
            t = minutes_to_time(slot_start)
            hhmm = str(t)
            if hhmm.endswith(":00"):
                set_font(True, 7.5)
                c.setFillColorRGB(0.18, 0.22, 0.28)
            else:
                set_font(False, 6.8)
                c.setFillColorRGB(0.45, 0.50, 0.58)
            c.drawCentredString(margin + time_w / 2, yinv(y0 + row_h / 2), hhmm)
            for cidx in range(cols):
                x0 = margin + time_w + (cidx * col_w)
                if self.plano_intervalo_bloqueado(slot_start, slot_end):
                    c.setFillColorRGB(0.86, 0.86, 0.86)
                else:
                    c.setFillColorRGB(1, 1, 1)
                c.rect(x0, yinv(y0 + row_h), col_w, row_h, stroke=1, fill=1)
        c.setStrokeColorRGB(0.72, 0.76, 0.82)
        for cidx in range(cols + 1):
            xv = margin + time_w + (cidx * col_w)
            c.line(xv, yinv(top_margin + row_h), xv, yinv(top_margin + (rows + 1) * row_h))
        c.setFillColorRGB(0, 0, 0)
        c.setStrokeColorRGB(0, 0, 0)

        set_font(False, 7)
        def _wrap_line_for_width(text, size, max_w, font_name=None):
            txt = str(text or "").strip()
            if not txt:
                return []
            parts = txt.split()
            if not parts:
                return [txt]
            out = []
            cur = parts[0]
            font_used = font_name or font_regular
            for part in parts[1:]:
                candidate = f"{cur} {part}"
                if pdfmetrics.stringWidth(candidate, font_used, size) <= max_w:
                    cur = candidate
                else:
                    out.append(cur)
                    cur = part
            out.append(cur)
            return out

        def _fit_block_lines(lines, box_w, box_h):
            raw_specs = []
            for line in list(lines or []):
                if isinstance(line, dict):
                    text = str(line.get("text", "") or "").strip()
                    role = str(line.get("role", "body") or "body").strip()
                else:
                    text = str(line or "").strip()
                    role = "body"
                if text:
                    raw_specs.append({"text": text, "role": role})
            if not raw_specs:
                raw_specs = [{"text": "-", "role": "body"}]
            for size in (9.8, 8.8, 7.8, 7.0, 6.2, 5.8):
                wrapped = []
                line_h = size + 1.25
                max_lines = max(1, int((box_h - 6) // line_h))
                for spec in raw_specs:
                    role = str(spec.get("role", "body") or "body")
                    font_used = font_bold if role in ("title", "time") else font_regular
                    for part in _wrap_line_for_width(spec.get("text", ""), size, box_w, font_name=font_used):
                        wrapped.append({"text": part, "role": role})
                if len(wrapped) <= max_lines:
                    return wrapped, size
            size = 5.8
            line_h = size + 1.15
            max_lines = max(1, int((box_h - 6) // line_h))
            wrapped = []
            for spec in raw_specs:
                role = str(spec.get("role", "body") or "body")
                font_used = font_bold if role in ("title", "time") else font_regular
                for part in _wrap_line_for_width(spec.get("text", ""), size, box_w, font_name=font_used):
                    wrapped.append({"text": part, "role": role})
            wrapped = wrapped[:max_lines]
            if wrapped:
                last = dict(wrapped[-1])
                font_used = font_bold if last.get("role") in ("title", "time") else font_regular
                text = str(last.get("text", "") or "").strip()
                while text and pdfmetrics.stringWidth(f"{text}...", font_used, size) > box_w:
                    text = text[:-1].rstrip()
                last["text"] = f"{text}..." if text else "..."
                wrapped[-1] = last
            return wrapped, size

        def draw_block_text(cx, cy, lines, box_w, box_h):
            wrapped, size = _fit_block_lines(lines, box_w, box_h)
            line_h = size + 1.25
            total_h = len(wrapped) * line_h
            start_y = cy - (total_h / 2) + (line_h / 2)
            for i, item_line in enumerate(wrapped):
                role = str(item_line.get("role", "body") or "body")
                text = str(item_line.get("text", "") or "").strip() or "-"
                c.setFont(font_bold if role in ("title", "time") else font_regular, size)
                c.drawCentredString(cx, yinv(start_y + i * line_h), text)

        enc_map = {str(e.get("numero", "") or ""): e for e in self.data.get("encomendas", [])}

        def _hex_to_rgb01(hx):
            try:
                hh = str(hx or "").strip().lstrip("#")
                if len(hh) != 6:
                    return (0.82, 0.90, 0.98)
                return (int(hh[0:2], 16) / 255.0, int(hh[2:4], 16) / 255.0, int(hh[4:6], 16) / 255.0)
            except Exception:
                return (0.82, 0.90, 0.98)

        def _blend_with_white(rgb, alpha=0.30):
            r, g, b = rgb
            return (r + (1 - r) * alpha, g + (1 - g) * alpha, b + (1 - b) * alpha)

        def _darken(rgb, factor=0.65):
            r, g, b = rgb
            return (max(0, r * factor), max(0, g * factor), max(0, b * factor))

        for item in self.data.get("plano", []):
            try:
                d = datetime.strptime(item["data"], "%Y-%m-%d").date()
            except Exception:
                continue
            if d < dates[0] or d > dates[-1]:
                continue
            col = (d - dates[0]).days
            start = time_to_minutes(item["inicio"])
            dur = int(item["duracao_min"])
            r_start = (start - start_min) // slot
            r_span = max(1, int(math.ceil(float(dur) / float(slot))))
            if r_start < 0 or r_start >= rows:
                continue
            x0 = margin + time_w + (col * col_w)
            y0 = top_margin + (r_start + 1) * row_h
            y1 = y0 + r_span * row_h
            base_rgb = _hex_to_rgb01(_pdf_block_color_for_item(item))
            fill_rgb = _blend_with_white(base_rgb, 0.30)
            edge_rgb = _darken(base_rgb, 0.65)
            c.setFillColorRGB(*fill_rgb)
            c.setStrokeColorRGB(*edge_rgb)
            c.rect(x0 + 2, yinv(y1 - 2), col_w - 4, (y1 - y0) - 4, stroke=1, fill=1)
            c.setFillColorRGB(*base_rgb)
            c.rect(x0 + 2, yinv(y1 - 2), 6, (y1 - y0) - 4, stroke=0, fill=1)
            c.setFillColorRGB(0, 0, 0)
            bloco_h = (y1 - y0) - 6
            enc_num = str(item.get("encomenda", "") or "").strip()
            enc_obj = enc_map.get(enc_num, {}) if enc_num else {}
            cliente = str(enc_obj.get("cliente", "") or "").strip()
            tempo_txt = f"{dur} min"
            linhas = [{"text": enc_num or "-", "role": "title"}]
            if cliente:
                linhas.append({"text": f"Cliente: {cliente}", "role": "body"})
            mat = str(item.get("material", "") or "").strip()
            esp = str(item.get("espessura", "") or "").strip()
            mat_esp = " | ".join(x for x in [mat, f"{esp}mm" if esp else ""] if x).strip()
            if mat_esp:
                linhas.append({"text": mat_esp, "role": "body"})
            fim = minutes_to_time(start + dur)
            linhas.append({"text": f"{item.get('inicio','')} - {fim}", "role": "body"})
            linhas.append({"text": tempo_txt, "role": "time"})
            chapa = str(item.get("chapa", "") or "").strip()
            if chapa and chapa != "-" and bloco_h >= (row_h * 2):
                linhas.append({"text": f"Chapa: {chapa}", "role": "body"})
            if bloco_h <= (row_h * 1.2):
                compact = [{"text": enc_num or "-", "role": "title"}]
                compact.append({"text": tempo_txt, "role": "time"})
                if mat_esp and bloco_h > (row_h * 0.9):
                    compact.append({"text": mat_esp, "role": "body"})
                linhas = compact
            elif bloco_h <= (row_h * 1.9):
                compact = [{"text": enc_num or "-", "role": "title"}]
                if mat_esp:
                    compact.append({"text": mat_esp, "role": "body"})
                compact.append({"text": f"{item.get('inicio','')} - {fim} | {tempo_txt}", "role": "time"})
                linhas = compact
            draw_block_text(x0 + col_w / 2, y0 + (y1 - y0) / 2, linhas, col_w - 10, (y1 - y0) - 6)

        c.setStrokeColorRGB(0.6, 0.65, 0.7)

        # Caixa de Observações / Data / Operador para assinatura
        box_h = footer_box_h
        box_y = height - margin - box_h
        c.setStrokeColorRGB(0.6, 0.65, 0.7)
        c.rect(margin, yinv(box_y + box_h), width - margin * 2, box_h, stroke=1, fill=0)
        c.setStrokeColorRGB(0, 0, 0)
        set_font(True, 8)
        c.drawString(margin + 6, yinv(box_y + 14), "Observacoes:")
        set_font(False, 8)
        c.line(margin + 6, yinv(box_y + 34), width - margin - 6, yinv(box_y + 34))
        set_font(True, 8)
        c.drawString(margin + 6, yinv(box_y + 52), "Data:")
        c.drawString(margin + 180, yinv(box_y + 52), "Operador:")
        c.save()

    try:
        render_pdf(path)
    except Exception as exc:
        messagebox.showerror("Erro", f"Falha ao gerar PDF: {exc}")
        return
    try:
        os.startfile(path)
    except Exception:
        try:
            os.startfile(path, "open")
        except Exception:
            messagebox.showerror("Erro", "Não foi possível abrir o PDF do plano.")

