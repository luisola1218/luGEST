from module_context import configure_module, ensure_module

_CONFIGURED = False

def configure(main_globals):
    configure_module(globals(), main_globals)

def _ensure_configured():
    ensure_module(globals(), "encomendas_actions")

def _is_orc_based_encomenda(enc):
    return bool(str((enc or {}).get("numero_orcamento", "") or "").strip())

def _norm_material(value):
    return str(value or "").strip().lower()

def _norm_espessura(value):
    txt = str(value or "").strip().lower().replace("mm", "").replace(",", ".")
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

def _match_material(a, b):
    ma = _norm_material(a)
    mb = _norm_material(b)
    if not ma or not mb:
        return False
    return ma == mb or ma in mb or mb in ma


def _enc_extract_year(data_criacao="", data_entrega="", numero_txt="", ano_val=None):
    _ensure_configured()
    a = str(ano_val or "").strip()
    if len(a) == 4 and a.isdigit():
        return a
    for raw in (data_criacao, data_entrega):
        txt = str(raw or "").strip()
        if len(txt) >= 4 and txt[:4].isdigit():
            return txt[:4]
    num = str(numero_txt or "").strip()
    if len(num) >= 4 and num[:4].isdigit():
        return num[:4]
    return ""


def _mysql_encomendas_years():
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
                        WHEN `data_criacao` IS NOT NULL THEN YEAR(`data_criacao`)
                        WHEN `data_entrega` IS NOT NULL THEN YEAR(`data_entrega`)
                        ELSE NULL
                    END AS ano
                FROM `encomendas`
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


def _mysql_encomendas_years_cached(self, ttl_sec=120):
    _ensure_configured()
    try:
        now_ts = time.time()
    except Exception:
        now_ts = 0
    cache_ts = float(getattr(self, "_enc_years_cache_ts", 0) or 0)
    cache_vals = getattr(self, "_enc_years_cache_vals", None)
    if isinstance(cache_vals, list) and (now_ts - cache_ts) <= max(5, int(ttl_sec)):
        return list(cache_vals)
    vals = _mysql_encomendas_years()
    try:
        self._enc_years_cache_vals = list(vals)
        self._enc_years_cache_ts = now_ts
    except Exception:
        pass
    return vals


def refresh_encomendas_year_options(self, keep_selection=True):
    _ensure_configured()
    current_year = str(datetime.now().year)
    selected = (
        str(self.e_year_filter.get() or "").strip()
        if keep_selection and hasattr(self, "e_year_filter")
        else current_year
    )
    years = {current_year}
    for e in self.data.get("encomendas", []):
        if not isinstance(e, dict):
            continue
        y = _enc_extract_year(e.get("data_criacao", ""), e.get("data_entrega", ""), e.get("numero", ""), e.get("ano"))
        if y:
            years.add(y)
    for y in _mysql_encomendas_years_cached(self):
        years.add(y)
    year_values = sorted(
        years,
        key=lambda x: int(x) if str(x).isdigit() else 0,
        reverse=True,
    )
    values = year_values + ["Todos"]
    if selected not in values:
        selected = current_year if current_year in values else values[0]
    if hasattr(self, "e_year_filter"):
        self.e_year_filter.set(selected)
    try:
        cb = getattr(self, "e_year_cb", None)
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


def _reload_encomendas_from_mysql(self, year_value=None):
    _ensure_configured()
    if not (USE_MYSQL_STORAGE and MYSQL_AVAILABLE):
        return False
    year_txt = str(
        year_value if year_value is not None else (self.e_year_filter.get() if hasattr(self, "e_year_filter") else "")
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
                    FROM `encomendas`
                    WHERE (
                        (`ano` = %s)
                        OR (`ano` IS NULL AND `data_criacao` IS NOT NULL AND YEAR(`data_criacao`) = %s)
                        OR (`ano` IS NULL AND `data_criacao` IS NULL AND `data_entrega` IS NOT NULL AND YEAR(`data_entrega`) = %s)
                    )
                    ORDER BY `numero`
                    """,
                    (year_num, year_num, year_num),
                )
            else:
                cur.execute("SELECT * FROM `encomendas` ORDER BY `numero`")
            enc_rows = cur.fetchall() or []

            enc_nums = [str(r.get("numero", "") or "") for r in enc_rows if str(r.get("numero", "") or "").strip()]

            def _fetch_in_chunks(table_name, col_name):
                rows = []
                chunk_size = 300
                for i in range(0, len(enc_nums), chunk_size):
                    chunk = enc_nums[i : i + chunk_size]
                    if not chunk:
                        continue
                    placeholders = ",".join(["%s"] * len(chunk))
                    cur.execute(
                        f"SELECT * FROM `{table_name}` WHERE `{col_name}` IN ({placeholders}) ORDER BY `id`",
                        tuple(chunk),
                    )
                    rows.extend(cur.fetchall() or [])
                return rows

            peca_rows = _fetch_in_chunks("pecas", "encomenda_numero")
            esp_rows = _fetch_in_chunks("encomenda_espessuras", "encomenda_numero")
            reservas_rows = _fetch_in_chunks("encomenda_reservas", "encomenda_numero")

        peca_map = {}
        for p in peca_rows:
            enc_num = str(p.get("encomenda_numero", "") or "")
            fluxo = []
            hist_rows = []
            try:
                raw_fluxo = p.get("operacoes_fluxo_json")
                if raw_fluxo:
                    parsed = json.loads(raw_fluxo)
                    if isinstance(parsed, list):
                        fluxo = parsed
            except Exception:
                fluxo = []
            try:
                raw_hist = p.get("hist_json")
                if raw_hist:
                    parsed_h = json.loads(raw_hist)
                    if isinstance(parsed_h, list):
                        hist_rows = parsed_h
            except Exception:
                hist_rows = []
            peca = {
                "id": str(p.get("id", "") or ""),
                "ref_interna": str(p.get("ref_interna", "") or ""),
                "ref_externa": str(p.get("ref_externa", "") or ""),
                "material": str(p.get("material", "") or ""),
                "espessura": str(p.get("espessura", "") or ""),
                "quantidade_pedida": _to_num(p.get("quantidade_pedida")) or 0.0,
                "Operacoes": str(p.get("operacoes", "") or ""),
                "Observacoes": str(p.get("observacoes", "") or ""),
                "of": str(p.get("of_codigo", "") or ""),
                "opp": str(p.get("opp_codigo", "") or ""),
                "estado": str(p.get("estado", "") or ""),
                "produzido_ok": _to_num(p.get("produzido_ok")) or 0.0,
                "produzido_nok": _to_num(p.get("produzido_nok")) or 0.0,
                "inicio_producao": _db_to_iso(p.get("inicio_producao")),
                "fim_producao": _db_to_iso(p.get("fim_producao")),
                "produzido_qualidade": 0.0,
                "tempo_producao_min": _to_num(p.get("tempo_producao_min")) or 0.0,
                "hist": hist_rows,
                "lote_baixa": str(p.get("lote_baixa", "") or ""),
                "desenho": str(p.get("desenho_path", "") or ""),
                "operacoes_fluxo": fluxo,
                "qtd_expedida": _to_num(p.get("qtd_expedida")) or 0.0,
                "expedicoes": [],
            }
            peca_map.setdefault(enc_num, []).append(peca)

        esp_meta_map = {}
        for r in esp_rows:
            enc_num = str(r.get("encomenda_numero", "") or "")
            mat = str(r.get("material", "") or "").strip()
            esp = str(r.get("espessura", "") or "").strip()
            if not enc_num or not mat or not esp:
                continue
            esp_meta_map.setdefault(enc_num, {})[(mat, esp)] = {
                "tempo_min": _to_num(r.get("tempo_min")),
                "estado": str(r.get("estado", "") or "Preparacao"),
                "inicio_producao": _db_to_iso(r.get("inicio_producao")),
                "fim_producao": _db_to_iso(r.get("fim_producao")),
                "tempo_producao_min": _to_num(r.get("tempo_producao_min")) or 0.0,
                "lote_baixa": str(r.get("lote_baixa", "") or ""),
            }

        reservas_map = {}
        for r in reservas_rows:
            enc_num = str(r.get("encomenda_numero", "") or "").strip()
            mat = str(r.get("material", "") or "").strip()
            esp = str(r.get("espessura", "") or "").strip()
            qtd = _to_num(r.get("quantidade")) or 0.0
            if not enc_num or not mat or not esp or qtd <= 0:
                continue
            reservas_map.setdefault(enc_num, []).append(
                {
                    "material_id": str(r.get("material_id", "") or ""),
                    "material": mat,
                    "espessura": esp,
                    "quantidade": qtd,
                }
            )

        new_rows = []
        for e in enc_rows:
            num = str(e.get("numero", "") or "")
            mats = {}
            for p in peca_map.get(num, []):
                mat = str(p.get("material", "") or "")
                esp = str(p.get("espessura", "") or "")
                mats.setdefault(mat, {"material": mat, "estado": "Preparacao", "esp_map": {}})
                mats[mat]["esp_map"].setdefault(
                    esp,
                    {
                        "espessura": esp,
                        "tempo_min": "",
                        "estado": "Preparacao",
                        "pecas": [],
                        "inicio_producao": "",
                        "fim_producao": "",
                        "tempo_producao_min": 0.0,
                        "lote_baixa": "",
                    },
                )
                mats[mat]["esp_map"][esp]["pecas"].append(p)
            for (mat, esp), meta in esp_meta_map.get(num, {}).items():
                mats.setdefault(mat, {"material": mat, "estado": "Preparacao", "esp_map": {}})
                mats[mat]["esp_map"].setdefault(
                    esp,
                    {
                        "espessura": esp,
                        "tempo_min": "",
                        "estado": "Preparacao",
                        "pecas": [],
                        "inicio_producao": "",
                        "fim_producao": "",
                        "tempo_producao_min": 0.0,
                        "lote_baixa": "",
                    },
                )
                slot = mats[mat]["esp_map"][esp]
                if meta.get("tempo_min") is not None:
                    slot["tempo_min"] = meta.get("tempo_min")
                slot["estado"] = meta.get("estado") or slot.get("estado", "Preparacao")
                slot["inicio_producao"] = meta.get("inicio_producao", slot.get("inicio_producao", ""))
                slot["fim_producao"] = meta.get("fim_producao", slot.get("fim_producao", ""))
                slot["tempo_producao_min"] = meta.get("tempo_producao_min", slot.get("tempo_producao_min", 0.0))
                slot["lote_baixa"] = meta.get("lote_baixa", slot.get("lote_baixa", ""))
            materiais = []
            for m in mats.values():
                m["espessuras"] = list(m.pop("esp_map").values())
                materiais.append(m)
            new_rows.append(
                {
                    "numero": num,
                    "ano": int(e.get("ano")) if e.get("ano") not in (None, "") else None,
                    "cliente": str(e.get("cliente_codigo", "") or ""),
                    "nota_cliente": str(e.get("nota_cliente", "") or ""),
                    "data_criacao": _db_to_iso(e.get("data_criacao")),
                    "data_entrega": _db_to_iso(e.get("data_entrega"))[:10],
                    "tempo": 0.0,
                    "tempo_estimado": _to_num(e.get("tempo_estimado")) or 0.0,
                    "cativar": bool(e.get("cativar")) or bool(reservas_map.get(num)),
                    "observacoes": str(e.get("observacoes", "") or ""),
                    "ObservaÃ§Ãµes": str(e.get("observacoes", "") or ""),
                    "estado": str(e.get("estado", "") or "Preparacao"),
                    "materiais": materiais,
                    "reservas": reservas_map.get(num, []),
                    "numero_orcamento": str(e.get("numero_orcamento", "") or ""),
                    "Observacoes": "",
                    "inicio_producao": "",
                    "tempo_pecas_min": 0.0,
                    "tempo_espessuras_min": 0.0,
                    "fim_producao": "",
                    "tempo_producao_min": 0.0,
                    "inicio_encomenda": "",
                    "fim_encomenda": "",
                    "estado_operador": "",
                    "obs_inicio": "",
                    "obs_interrupcao": "",
                    "tempo_por_espessura": {},
                    "espessuras": [],
                }
            )

        if year_num:
            year_key = str(year_num)
            retained = []
            for enc in self.data.get("encomendas", []):
                if not isinstance(enc, dict):
                    continue
                y = _enc_extract_year(enc.get("data_criacao", ""), enc.get("data_entrega", ""), enc.get("numero", ""), enc.get("ano"))
                if y != year_key:
                    retained.append(enc)
            self.data["encomendas"] = retained + new_rows
        else:
            self.data["encomendas"] = new_rows
        return True
    except Exception:
        return False
    finally:
        try:
            if conn is not None:
                conn.close()
        except Exception:
            pass


def _on_encomendas_year_change(self, value=None):
    _ensure_configured()
    try:
        if value is not None and hasattr(self, "e_year_filter"):
            self.e_year_filter.set(str(value))
    except Exception:
        pass
    year_key = str(self.e_year_filter.get() if hasattr(self, "e_year_filter") else value or "").strip()
    if not year_key:
        year_key = str(value or "")
    if not year_key:
        year_key = str(datetime.now().year)
    cache_key = year_key.lower()
    try:
        now_ts = time.time()
    except Exception:
        now_ts = 0
    last_key = str(getattr(self, "_enc_mysql_reload_key", "") or "")
    last_ts = float(getattr(self, "_enc_mysql_reload_ts", 0) or 0)
    ttl_sec = 45
    need_reload = (cache_key != last_key) or ((now_ts - last_ts) > ttl_sec)
    if need_reload:
        _reload_encomendas_from_mysql(self, year_value=year_key)
        try:
            self._enc_mysql_reload_key = cache_key
            self._enc_mysql_reload_ts = now_ts
        except Exception:
            pass
    try:
        self.refresh_encomendas_year_options(keep_selection=True)
    except Exception:
        pass
    self.clear_encomenda_selection()
    self.refresh_encomendas()


def _open_encomenda_header_dialog(self, enc=None):
    _ensure_configured()
    from tkinter import simpledialog
    if not (CUSTOM_TK_AVAILABLE and getattr(self, "encomendas_use_custom", False)):
        messagebox.showerror("Erro", "Este editor requer interface CustomTkinter ativa.")
        return

    is_edit = bool(enc)
    working = copy.deepcopy(enc) if is_edit else {
        "id": "",
        "numero": "",
        "cliente": "",
        "posto_trabalho": "",
        "nota_cliente": "",
        "data_criacao": now_iso(),
        "data_entrega": "",
        "tempo_estimado": 0.0,
        "cativar": False,
        "ObservaÃ§Ãµes": "",
        "Observacoes": "",
        "estado": "Preparacao",
        "materiais": [],
        "reservas": [],
        "espessuras": [],
        "numero_orcamento": "",
    }
    working.setdefault("materiais", [])

    def Btn(parent, **kwargs):
        kwargs.setdefault("height", 34)
        kwargs.setdefault("corner_radius", 10)
        kwargs.setdefault("fg_color", CTK_PRIMARY_RED)
        kwargs.setdefault("hover_color", CTK_PRIMARY_RED_HOVER)
        kwargs.setdefault("text_color", "#ffffff")
        kwargs.setdefault("border_width", 0)
        return ctk.CTkButton(parent, **kwargs)

    win = ctk.CTkToplevel(self.root)
    win.title("Editar encomenda" if is_edit else "Criar encomenda")
    try:
        sw = int(win.winfo_screenwidth() or 1366)
        sh = int(win.winfo_screenheight() or 768)
        ww = max(1240, min(1520, sw - 60))
        wh = max(780, min(940, sh - 90))
        pos_x = max(0, (sw - ww) // 2)
        pos_y = max(0, (sh - wh) // 2)
        win.geometry(f"{ww}x{wh}+{pos_x}+{pos_y}")
        win.minsize(1240, 780)
    except Exception:
        win.geometry("1400x860")
    win.configure(fg_color="#f5f7fb")
    try:
        win.transient(self.root)
        win.grab_set()
        win.lift()
        try:
            win.focus_force()
        except Exception:
            pass
    except Exception:
        pass
    self._enc_editor_win = win

    hero = ctk.CTkFrame(win, fg_color="#eaf2ff", corner_radius=12, border_width=1, border_color="#c9d8ef")
    hero.pack(fill="x", padx=12, pady=(12, 6))
    ctk.CTkLabel(
        hero,
        text=("Editar Encomenda" if is_edit else "Criar Encomenda"),
        font=("Segoe UI", 19, "bold"),
        text_color="#0b1f4d",
    ).pack(side="left", padx=12, pady=8)
    ctk.CTkLabel(
        hero,
        text=(f"Numero: {working.get('numero','(novo)')}" if is_edit else "Novo registo ligado ao MySQL"),
        font=("Segoe UI", 12, "bold"),
        text_color="#334155",
    ).pack(side="right", padx=12, pady=8)

    cli_var = StringVar(value=str(working.get("cliente", "") or ""))
    posto_options = [str(v).strip() for v in list(self.data.get("postos_trabalho", []) or ["Geral"]) if str(v).strip()]
    if not posto_options:
        posto_options = ["Geral"]
    posto_current = str(working.get("posto_trabalho", "") or working.get("posto", "") or working.get("maquina", "") or "").strip()
    if posto_current and posto_current not in posto_options:
        posto_options.append(posto_current)
    posto_var = StringVar(value=posto_current or posto_options[0])
    nota_var = StringVar(value=str(working.get("nota_cliente", "") or ""))
    entrega_var = StringVar(value=str(working.get("data_entrega", "") or ""))
    tempo_var = StringVar(value=str(working.get("tempo_estimado", 0) or 0))
    obs_var = StringVar(
        value=str(
            working.get("ObservaÃ§Ãµes", "")
            or working.get("Observacoes", "")
            or ""
        )
    )
    cativar_var = BooleanVar(value=bool(working.get("cativar", False)))

    top = ctk.CTkFrame(win, fg_color="white", corner_radius=10, border_width=1, border_color="#d8dee8")
    top.pack(fill="x", padx=12, pady=(12, 8))

    ctk.CTkLabel(top, text="Cliente").grid(row=0, column=0, padx=8, pady=6, sticky="w")
    ctk.CTkComboBox(top, variable=cli_var, values=self.get_clientes_codes() or [""], width=230).grid(row=0, column=1, padx=8, pady=6, sticky="w")

    ctk.CTkLabel(top, text="Data entrega").grid(row=0, column=2, padx=8, pady=6, sticky="w")
    ctk.CTkEntry(top, textvariable=entrega_var, width=130).grid(row=0, column=3, padx=8, pady=6, sticky="w")
    Btn(top, text="Calendário", width=110, command=lambda: self.pick_date(entrega_var)).grid(row=0, column=4, padx=6, pady=6, sticky="w")

    ctk.CTkLabel(top, text="Tempo (h)").grid(row=0, column=5, padx=8, pady=6, sticky="w")
    ctk.CTkEntry(top, textvariable=tempo_var, width=90).grid(row=0, column=6, padx=8, pady=6, sticky="w")
    ctk.CTkLabel(top, text="Posto").grid(row=0, column=7, padx=8, pady=6, sticky="w")
    ctk.CTkComboBox(top, variable=posto_var, values=posto_options, width=160).grid(row=0, column=8, padx=8, pady=6, sticky="w")

    ctk.CTkLabel(top, text="Nota cliente").grid(row=1, column=0, padx=8, pady=6, sticky="w")
    ctk.CTkEntry(top, textvariable=nota_var, width=420).grid(row=1, column=1, columnspan=3, padx=8, pady=6, sticky="we")

    ctk.CTkLabel(top, text="Observações").grid(row=1, column=4, padx=8, pady=6, sticky="w")
    ctk.CTkEntry(top, textvariable=obs_var, width=320).grid(row=1, column=5, columnspan=2, padx=8, pady=6, sticky="we")

    ctk.CTkCheckBox(top, text="Cativar MP", variable=cativar_var).grid(row=1, column=7, columnspan=2, padx=10, pady=6, sticky="w")

    cativ_info = ctk.CTkFrame(win, fg_color="white", corner_radius=10, border_width=1, border_color="#d8dee8")
    cativ_info.pack(fill="x", padx=12, pady=(0, 8))
    ctk.CTkLabel(cativ_info, text="Cativações (material reservado)", font=("Segoe UI", 13, "bold")).pack(anchor="w", padx=10, pady=(8, 4))
    cativ_txt = ctk.CTkTextbox(cativ_info, height=72, wrap="word", fg_color="#f8fbff", border_width=1, border_color="#d8dee8")
    cativ_txt.pack(fill="x", padx=10, pady=(0, 8))

    body = ctk.CTkFrame(win, fg_color="transparent")
    body.pack(fill="both", expand=True, padx=12, pady=(0, 8))

    # Treeviews customizadas para evitar visual ttk clássico na janela de edição.
    tv_style = ttk.Style(win)
    try:
        tv_style.theme_use("clam")
    except Exception:
        pass
    tv_style.configure(
        "EncEditor.Treeview",
        font=("Segoe UI", 10),
        rowheight=28,
        background="#f8fbff",
        fieldbackground="#f8fbff",
        borderwidth=0,
        relief="flat",
    )
    tv_style.configure(
        "EncEditor.Treeview.Heading",
        font=("Segoe UI", 10, "bold"),
        background=THEME_HEADER_BG,
        foreground="#ffffff",
        relief="flat",
    )
    tv_style.map("EncEditor.Treeview.Heading", background=[("active", THEME_HEADER_ACTIVE)])
    tv_style.map(
        "EncEditor.Treeview",
        background=[("selected", THEME_SELECT_BG)],
        foreground=[("selected", THEME_SELECT_FG)],
    )

    sec_m = ctk.CTkFrame(body, fg_color="white", corner_radius=10, border_width=1, border_color="#d8dee8")
    sec_e = ctk.CTkFrame(body, fg_color="white", corner_radius=10, border_width=1, border_color="#d8dee8")
    sec_p = ctk.CTkFrame(body, fg_color="white", corner_radius=10, border_width=1, border_color="#d8dee8")
    sec_m.pack(side="left", fill="both", expand=True)
    sec_e.pack(side="left", fill="both", expand=True, padx=8)
    sec_p.pack(side="left", fill="both", expand=True)

    ctk.CTkLabel(sec_m, text="Materiais", font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=8, pady=(8, 4))
    ctk.CTkLabel(sec_e, text="Espessuras", font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=8, pady=(8, 4))
    ctk.CTkLabel(sec_p, text="Peças", font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=8, pady=(8, 4))

    tbl_m = ttk.Treeview(sec_m, columns=("material", "estado"), show="headings", height=16, style="EncEditor.Treeview")
    tbl_m.heading("material", text="Material")
    tbl_m.heading("estado", text="Estado")
    tbl_m.column("material", width=230, anchor="w")
    tbl_m.column("estado", width=100, anchor="center")
    tbl_m.pack(fill="both", expand=True, padx=8, pady=4)
    sb_m = ctk.CTkScrollbar(sec_m, orientation="vertical", command=tbl_m.yview)
    tbl_m.configure(yscrollcommand=sb_m.set)
    sb_m.pack(side="right", fill="y", padx=(0, 6), pady=(0, 6))

    tbl_e = ttk.Treeview(sec_e, columns=("esp", "tempo", "estado"), show="headings", height=16, style="EncEditor.Treeview")
    tbl_e.heading("esp", text="Espessura")
    tbl_e.heading("tempo", text="Tempo")
    tbl_e.heading("estado", text="Estado")
    tbl_e.column("esp", width=100, anchor="center")
    tbl_e.column("tempo", width=90, anchor="center")
    tbl_e.column("estado", width=110, anchor="center")
    tbl_e.pack(fill="both", expand=True, padx=8, pady=4)
    sb_e = ctk.CTkScrollbar(sec_e, orientation="vertical", command=tbl_e.yview)
    tbl_e.configure(yscrollcommand=sb_e.set)
    sb_e.pack(side="right", fill="y", padx=(0, 6), pady=(0, 6))

    tbl_p = ttk.Treeview(sec_p, columns=("refi", "refe", "qtd", "ops", "estado"), show="headings", height=16, style="EncEditor.Treeview")
    tbl_p.heading("refi", text="Ref. Interna")
    tbl_p.heading("refe", text="Ref. Externa")
    tbl_p.heading("qtd", text="Qtd")
    tbl_p.heading("ops", text="Operações")
    tbl_p.heading("estado", text="Estado")
    tbl_p.column("refi", width=130, anchor="w")
    tbl_p.column("refe", width=150, anchor="w")
    tbl_p.column("qtd", width=70, anchor="e")
    tbl_p.column("ops", width=170, anchor="w")
    tbl_p.column("estado", width=95, anchor="center")
    tbl_p.pack(fill="both", expand=True, padx=8, pady=4)
    sb_p = ctk.CTkScrollbar(sec_p, orientation="vertical", command=tbl_p.yview)
    tbl_p.configure(yscrollcommand=sb_p.set)
    sb_p.pack(side="right", fill="y", padx=(0, 6), pady=(0, 6))

    sel = {"m": None, "e": None}
    target = {"enc": enc if is_edit else None}
    sync_state = {"busy": False}

    def _select_main_encomenda(numero):
        try:
            for item in self.tbl_encomendas.get_children():
                vals = self.tbl_encomendas.item(item, "values")
                if vals and str(vals[0]) == str(numero):
                    self.tbl_encomendas.selection_set(item)
                    self.tbl_encomendas.see(item)
                    break
        except Exception:
            pass

    def _set_main_context(tgt, mat=None, esp=None):
        if not tgt:
            return
        self.selected_encomenda_numero = str(tgt.get("numero", "") or "")
        self.selected_material = mat
        self.selected_espessura = esp
        _select_main_encomenda(self.selected_encomenda_numero)
        try:
            self.refresh_materiais(tgt)
            self.refresh_espessuras(tgt, self.selected_material)
            self.refresh_pecas(tgt, self.selected_espessura)
        except Exception:
            pass

    def _render_editor_reservas():
        tgt = target.get("enc") or {}
        linhas = []
        for r in (tgt.get("reservas", []) or []):
            mat = str(r.get("material", "") or "")
            esp = str(r.get("espessura", "") or "")
            qtd = parse_float(r.get("quantidade", 0), 0.0)
            lote = "-"
            dim = "-"
            mid = r.get("material_id")
            if mid:
                for mrow in self.data.get("materiais", []) or []:
                    if str(mrow.get("id")) == str(mid):
                        lote = str(mrow.get("lote_fornecedor", "") or "-")
                        dim = f"{mrow.get('comprimento','')}x{mrow.get('largura','')}"
                        break
            linhas.append(f"{mat} {esp} | {dim} | Lote: {lote} | Qtd: {qtd:g}")
        if not linhas:
            linhas = ["Sem chapas cativadas"]
        try:
            cativ_txt.configure(state="normal")
            cativ_txt.delete("1.0", "end")
            cativ_txt.insert("1.0", "\n".join(linhas))
            cativ_txt.configure(state="disabled")
        except Exception:
            pass

    def _sync_from_target(prefer_material=None, prefer_esp=None, prefer_ref=None):
        tgt = target.get("enc")
        if not tgt:
            return
        sync_state["busy"] = True
        try:
            mat_sel = str(prefer_material or (sel.get("m") or {}).get("material", "") or "")
            esp_sel = str(prefer_esp or (sel.get("e") or {}).get("espessura", "") or "")
            working["materiais"] = copy.deepcopy(tgt.get("materiais", []))
            _refresh_m()
            if mat_sel:
                for iid in tbl_m.get_children():
                    vals = tbl_m.item(iid, "values")
                    if vals and str(vals[0]) == mat_sel:
                        tbl_m.selection_set(iid)
                        tbl_m.see(iid)
                        sel["m"] = _find_m()
                        break
            if sel.get("m"):
                _refresh_e()
            # Se não vier espessura explícita, tenta descobrir pela peça (ref interna) para preservar contexto.
            if not esp_sel and prefer_ref and sel.get("m"):
                ref_txt = str(prefer_ref or "").strip()
                if ref_txt:
                    for ee in sel["m"].get("espessuras", []):
                        for pp in ee.get("pecas", []):
                            if str(pp.get("ref_interna", "") or "").strip() == ref_txt:
                                esp_sel = str(ee.get("espessura", "") or "").strip()
                                break
                        if esp_sel:
                            break

            selected_esp = False
            if esp_sel:
                esp_norm = _norm_espessura(esp_sel)
                for iid in tbl_e.get_children():
                    vals = tbl_e.item(iid, "values")
                    if vals and _norm_espessura(vals[0]) == esp_norm:
                        tbl_e.selection_set(iid)
                        tbl_e.see(iid)
                        sel["e"] = _find_e()
                        _refresh_p()
                        selected_esp = True
                        break

            # fallback: se houver só uma espessura, mantém-na selecionada automaticamente.
            if not selected_esp and sel.get("m") and len(sel["m"].get("espessuras", [])) == 1:
                only_esp = str(sel["m"]["espessuras"][0].get("espessura", "") or "")
                for iid in tbl_e.get_children():
                    vals = tbl_e.item(iid, "values")
                    if vals and _norm_espessura(vals[0]) == _norm_espessura(only_esp):
                        tbl_e.selection_set(iid)
                        tbl_e.see(iid)
                        sel["e"] = _find_e()
                        _refresh_p()
                        break
            # fallback final: mantém sempre uma espessura ativa para não esconder referências
            if not sel.get("e"):
                all_esp = tbl_e.get_children()
                if all_esp:
                    iid0 = all_esp[0]
                    tbl_e.selection_set(iid0)
                    tbl_e.see(iid0)
                    sel["e"] = _find_e()
                    _refresh_p()
            if prefer_ref:
                ref_txt = str(prefer_ref or "").strip()
                if ref_txt:
                    for iid in tbl_p.get_children():
                        vals = tbl_p.item(iid, "values")
                        if vals and str(vals[0] or "").strip() == ref_txt:
                            tbl_p.selection_set(iid)
                            tbl_p.see(iid)
                            break
            _render_editor_reservas()
        finally:
            sync_state["busy"] = False

    def _ensure_target():
        if target.get("enc"):
            return target.get("enc")
        cliente = (cli_var.get() or "").strip()
        if not cliente:
            messagebox.showerror("Erro", "Cliente obrigatório.")
            return None
        try:
            tempo_init = float((tempo_var.get() or "0").replace(",", "."))
        except Exception:
            tempo_init = 0.0
        numero = next_encomenda_numero(self.data)
        tgt = {
            "id": f"ENC{len(self.data['encomendas'])+1:05d}",
            "numero": numero,
            "cliente": cliente,
            "nota_cliente": (nota_var.get() or "").strip(),
            "data_criacao": now_iso(),
            "data_entrega": (entrega_var.get() or "").strip(),
            "tempo_estimado": tempo_init,
            "tempo": tempo_init,
            "cativar": bool(cativar_var.get()),
            "ObservaÃ§Ãµes": (obs_var.get() or "").strip(),
            "Observacoes": (obs_var.get() or "").strip(),
            "estado": "Preparacao",
            "materiais": [],
            "reservas": [],
            "espessuras": [],
            "numero_orcamento": "",
        }
        self.data["encomendas"].append(tgt)
        target["enc"] = tgt
        try:
            self.refresh_encomendas_year_options(keep_selection=True)
        except Exception:
            pass
        _set_main_context(tgt)
        save_data(self.data)
        return tgt

    def _find_m():
        s = tbl_m.selection()
        if not s:
            return None
        mat = str(tbl_m.item(s[0], "values")[0] or "")
        for m in working.get("materiais", []):
            if str(m.get("material", "")) == mat:
                return m
        return None

    def _find_e():
        m = _find_m()
        s = tbl_e.selection()
        if not m or not s:
            return None
        esp = str(tbl_e.item(s[0], "values")[0] or "")
        for e in m.get("espessuras", []):
            if str(e.get("espessura", "")) == esp:
                return e
        return None

    def _refresh_m():
        for i in tbl_m.get_children():
            tbl_m.delete(i)
        for m in working.get("materiais", []):
            tbl_m.insert("", END, values=(m.get("material", ""), m.get("estado", "Preparacao")))
        _refresh_e()

    def _refresh_e():
        for i in tbl_e.get_children():
            tbl_e.delete(i)
        m = sel.get("m")
        if not m:
            _refresh_p()
            return
        for e in m.get("espessuras", []):
            tbl_e.insert("", END, values=(e.get("espessura", ""), e.get("tempo_min", ""), e.get("estado", "Preparacao")))
        _refresh_p()

    def _refresh_p():
        for i in tbl_p.get_children():
            tbl_p.delete(i)
        e = sel.get("e")
        if not e:
            return
        for p in e.get("pecas", []):
            ops_txt = p.get("OperaÃ§Ãµes", "") or p.get("Operacoes", "")
            tbl_p.insert(
                "",
                END,
                values=(
                    p.get("ref_interna", ""),
                    p.get("ref_externa", ""),
                    p.get("quantidade_pedida", 0),
                    ops_txt,
                    p.get("estado", "Preparacao"),
                ),
            )

    def _on_m(_=None):
        if sync_state.get("busy"):
            return
        sel["m"] = _find_m()
        sel["e"] = None
        _refresh_e()

    def _on_e(_=None):
        if sync_state.get("busy"):
            return
        sel["e"] = _find_e()
        _refresh_p()

    tbl_m.bind("<<TreeviewSelect>>", _on_m)
    tbl_e.bind("<<TreeviewSelect>>", _on_e)

    btn_m = ctk.CTkFrame(sec_m, fg_color="transparent")
    btn_m.pack(fill="x", padx=8, pady=(2, 8))
    btn_e = ctk.CTkFrame(sec_e, fg_color="transparent")
    btn_e.pack(fill="x", padx=8, pady=(2, 8))
    btn_p = ctk.CTkFrame(sec_p, fg_color="transparent")
    btn_p.pack(fill="x", padx=8, pady=(2, 8))

    def _add_m():
        tgt = _ensure_target()
        if not tgt:
            return
        _set_main_context(tgt, None, None)
        dlg = self.add_material_encomenda()
        try:
            if dlg is not None and dlg.winfo_exists():
                win.wait_window(dlg)
        except Exception:
            pass
        _sync_from_target()

    def _rm_m():
        tgt = _ensure_target()
        if not tgt:
            return
        m = _find_m()
        if not m:
            return
        _set_main_context(tgt, str(m.get("material", "") or ""), None)
        self.remove_material_encomenda()
        sel["m"] = None; sel["e"] = None
        _sync_from_target()

    def _add_e():
        tgt = _ensure_target()
        if not tgt:
            return
        m = _find_m()
        if not m:
            messagebox.showerror("Erro", "Selecione um material.")
            return
        mat_name = str(m.get("material", "") or "")
        while True:
            before = 0
            mcur = next((x for x in tgt.get("materiais", []) if str(x.get("material", "")) == mat_name), None)
            if mcur:
                before = len(mcur.get("espessuras", []))
            _set_main_context(tgt, mat_name, None)
            dlg = self.add_espessura(enc_override=tgt, material_override=mat_name)
            try:
                if dlg is not None and dlg.winfo_exists():
                    win.wait_window(dlg)
            except Exception:
                pass
            _sync_from_target(prefer_material=mat_name)
            mcur = next((x for x in tgt.get("materiais", []) if str(x.get("material", "")) == mat_name), None)
            after = len(mcur.get("espessuras", [])) if mcur else 0
            if after <= before:
                break
            last_esp = ""
            if mcur and mcur.get("espessuras"):
                last_esp = str(mcur.get("espessuras", [])[-1].get("espessura", "") or "")
            _sync_from_target(prefer_material=mat_name, prefer_esp=last_esp)
            more = messagebox.askyesno("Espessuras", "Quer adicionar mais espessuras?")
            if not more:
                break

    def _rm_e():
        tgt = _ensure_target()
        if not tgt:
            return
        m = _find_m(); e = _find_e()
        if not m or not e:
            return
        mat_name = str(m.get("material", "") or "")
        _set_main_context(tgt, mat_name, str(e.get("espessura", "") or ""))
        self.remove_espessura()
        sel["e"] = None
        _sync_from_target(prefer_material=mat_name)

    def _tempo_e():
        tgt = _ensure_target()
        if not tgt:
            return
        m = _find_m()
        e = _find_e()
        if not m or not e:
            messagebox.showerror("Erro", "Selecione material e espessura.")
            return
        mat_name = str(m.get("material", "") or "")
        esp_name = str(e.get("espessura", "") or "")
        _set_main_context(tgt, mat_name, esp_name)
        self.edit_tempo_espessura()
        _sync_from_target(prefer_material=mat_name, prefer_esp=esp_name)

    def _edit_tempo_e():
        tgt = _ensure_target()
        if not tgt:
            return
        m = _find_m()
        e = _find_e()
        if not m or not e:
            messagebox.showerror("Erro", "Selecione material e espessura.")
            return
        mat_name = str(m.get("material", "") or "")
        esp_name = str(e.get("espessura", "") or "")
        esp_obj = next(
            (
                x
                for x in (m.get("espessuras", []) or [])
                if str(x.get("espessura", "") or "") == esp_name
            ),
            None,
        )
        atual = ""
        if esp_obj is not None:
            atual = str(esp_obj.get("tempo_min", "") or "").strip()

        use_custom_local = self.encomendas_use_custom and CUSTOM_TK_AVAILABLE
        parent_local = getattr(self, "_enc_editor_win", None) or win
        w2 = ctk.CTkToplevel(parent_local) if use_custom_local else Toplevel(parent_local)
        w2.title("Editar tempo")
        try:
            w2.geometry("380x170" if use_custom_local else "320x150")
            w2.transient(parent_local)
            w2.grab_set()
            w2.lift()
            w2.focus_force()
        except Exception:
            pass
        Lbl2 = ctk.CTkLabel if use_custom_local else ttk.Label
        Ent2 = ctk.CTkEntry if use_custom_local else ttk.Entry
        Btn2 = ctk.CTkButton if use_custom_local else ttk.Button
        tempo_v = StringVar(value=atual)
        Lbl2(w2, text=f"{mat_name} | {esp_name} mm").grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 6), sticky="w")
        Lbl2(w2, text="Tempo (min)").grid(row=1, column=0, padx=10, pady=6, sticky="w")
        Ent2(w2, textvariable=tempo_v, width=140 if use_custom_local else 14).grid(row=1, column=1, padx=10, pady=6, sticky="w")

        def _save_tempo():
            val = (tempo_v.get() or "").strip()
            if val:
                try:
                    int(val)
                except Exception:
                    messagebox.showerror("Erro", "Tempo inválido (minutos inteiros).")
                    return
            _set_main_context(tgt, mat_name, esp_name)
            for mm in tgt.get("materiais", []):
                if str(mm.get("material", "") or "") != mat_name:
                    continue
                for ee in mm.get("espessuras", []):
                    if str(ee.get("espessura", "") or "") == esp_name:
                        ee["tempo_min"] = val
                        break
            save_data(self.data)
            _sync_from_target(prefer_material=mat_name, prefer_esp=esp_name)
            w2.destroy()

        if use_custom_local:
            Btn2(
                w2,
                text="Guardar",
                command=_save_tempo,
                width=180,
                height=40,
                fg_color="#f59e0b",
                hover_color="#d97706",
            ).grid(row=2, column=0, columnspan=2, pady=(12, 10))
        else:
            Btn2(w2, text="Guardar", command=_save_tempo, width=18).grid(row=2, column=0, columnspan=2, pady=10)

    def _add_p():
        tgt = _ensure_target()
        if not tgt:
            return
        m = _find_m(); e = _find_e()
        if not m or not e:
            messagebox.showerror("Erro", "Selecione material e espessura.")
            return
        esp_sel = str(e.get("espessura", "") or "")
        mat_name = str(m.get("material", "") or "")
        _set_main_context(tgt, mat_name, esp_sel)
        dlg = self.add_peca(esp_override=esp_sel)
        try:
            if dlg is not None and dlg.winfo_exists():
                win.wait_window(dlg)
        except Exception:
            pass
        _sync_from_target(prefer_material=mat_name, prefer_esp=esp_sel)

    def _rm_p():
        tgt = _ensure_target()
        if not tgt:
            return
        m = _find_m(); e = _find_e(); s = tbl_p.selection()
        if not m or not e or not s:
            return
        mat_name = str(m.get("material", "") or "")
        esp_name = str(e.get("espessura", "") or "")
        _set_main_context(tgt, mat_name, esp_name)
        refi = str(tbl_p.item(s[0], "values")[0] or "")
        try:
            for iid in self.tbl_pecas.get_children():
                vals = self.tbl_pecas.item(iid, "values")
                if vals and str(vals[0]) == refi:
                    self.tbl_pecas.selection_set(iid)
                    self.tbl_pecas.see(iid)
                    break
        except Exception:
            pass
        self.remove_peca()
        _sync_from_target(prefer_material=mat_name, prefer_esp=esp_name)

    def _edit_p():
        tgt = _ensure_target()
        if not tgt:
            return
        m = _find_m()
        e = _find_e()
        s = tbl_p.selection()
        if not m or not e or not s:
            messagebox.showerror("Erro", "Selecione uma peça.")
            return
        mat_name = str(m.get("material", "") or "")
        esp_name = str(e.get("espessura", "") or "")
        vals = tbl_p.item(s[0], "values")
        refi = str(vals[0] if vals else "").strip()
        if not refi:
            messagebox.showerror("Erro", "Peça sem referência interna.")
            return

        peca_obj = None
        esp_obj = None
        for mm in tgt.get("materiais", []):
            if str(mm.get("material", "") or "") != mat_name:
                continue
            for ee in mm.get("espessuras", []):
                if str(ee.get("espessura", "") or "") != esp_name:
                    continue
                for pp in ee.get("pecas", []):
                    if str(pp.get("ref_interna", "") or "") == refi:
                        peca_obj = pp
                        esp_obj = ee
                        break
                if peca_obj:
                    break
            if peca_obj:
                break
        if not peca_obj:
            messagebox.showerror("Erro", "Peça não encontrada.")
            return

        use_custom_local = self.encomendas_use_custom and CUSTOM_TK_AVAILABLE
        parent_local = getattr(self, "_enc_editor_win", None) or win
        w2 = ctk.CTkToplevel(parent_local) if use_custom_local else Toplevel(parent_local)
        w2.title("Editar peça")
        try:
            w2.geometry("430x220" if use_custom_local else "360x200")
            w2.transient(parent_local)
            w2.grab_set()
            w2.lift()
            w2.focus_force()
        except Exception:
            pass
        Lbl2 = ctk.CTkLabel if use_custom_local else ttk.Label
        Ent2 = ctk.CTkEntry if use_custom_local else ttk.Entry
        Btn2 = ctk.CTkButton if use_custom_local else ttk.Button
        qtd_v = StringVar(value=str(peca_obj.get("quantidade_pedida", 0) or 0))
        Lbl2(w2, text=f"Ref. Interna: {refi}").grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 6), sticky="w")
        Lbl2(w2, text="Quantidade").grid(row=1, column=0, padx=10, pady=6, sticky="w")
        Ent2(w2, textvariable=qtd_v, width=160 if use_custom_local else 14).grid(row=1, column=1, padx=10, pady=6, sticky="w")

        def _save_peca():
            txt = (qtd_v.get() or "").strip().replace(",", ".")
            try:
                q = float(txt)
            except Exception:
                messagebox.showerror("Erro", "Quantidade inválida.")
                return
            if q <= 0:
                messagebox.showerror("Erro", "Quantidade deve ser maior que zero.")
                return
            peca_obj["quantidade_pedida"] = q
            # se já estava concluída e nova qtd for maior, regressa a produção
            total_prod = parse_float(peca_obj.get("produzido_ok", 0), 0.0) + parse_float(peca_obj.get("produzido_nok", 0), 0.0)
            if total_prod < q and "concl" in norm_text(peca_obj.get("estado", "")):
                peca_obj["estado"] = "Em producao"
                peca_obj["fim_producao"] = ""
            try:
                atualizar_estado_peca(peca_obj)
            except Exception:
                pass
            try:
                if esp_obj is not None:
                    if all(norm_text(px.get("estado", "")).find("concl") >= 0 for px in esp_obj.get("pecas", [])):
                        esp_obj["estado"] = "Concluida"
                    elif any((parse_float(px.get("produzido_ok", 0), 0.0) + parse_float(px.get("produzido_nok", 0), 0.0)) > 0 for px in esp_obj.get("pecas", [])):
                        esp_obj["estado"] = "Em producao"
                    else:
                        esp_obj["estado"] = "Preparacao"
            except Exception:
                pass
            save_data(self.data)
            _sync_from_target(prefer_material=mat_name, prefer_esp=esp_name, prefer_ref=refi)
            w2.destroy()

        if use_custom_local:
            Btn2(
                w2,
                text="Guardar",
                command=_save_peca,
                width=180,
                height=40,
                fg_color="#f59e0b",
                hover_color="#d97706",
            ).grid(row=2, column=0, columnspan=2, pady=(12, 10))
        else:
            Btn2(w2, text="Guardar", command=_save_peca, width=18).grid(row=2, column=0, columnspan=2, pady=10)

    def _open_p_desenho():
        tgt = _ensure_target()
        if not tgt:
            return
        s = tbl_p.selection()
        if not s:
            messagebox.showerror("Erro", "Selecione uma peça.")
            return
        vals = tbl_p.item(s[0], "values")
        refi = str(vals[0] if vals else "").strip()
        refe = str(vals[1] if vals and len(vals) > 1 else "").strip()
        open_peca_desenho_by_refs(self, enc=tgt, ref_interna=refi, ref_externa=refe)

    Btn(btn_m, text="Adicionar", width=118, command=_add_m).pack(side="left", padx=3)
    Btn(btn_m, text="Remover", width=118, command=_rm_m).pack(side="left", padx=3)
    Btn(btn_e, text="Adicionar", width=118, command=_add_e).pack(side="left", padx=3)
    Btn(btn_e, text="Remover", width=118, command=_rm_e).pack(side="left", padx=3)
    Btn(btn_e, text="Adicionar tempo", width=140, command=_tempo_e).pack(side="left", padx=3)
    Btn(btn_e, text="Editar tempo", width=130, command=_edit_tempo_e).pack(side="left", padx=3)
    Btn(btn_p, text="Adicionar", width=118, command=_add_p).pack(side="left", padx=3)
    Btn(btn_p, text="Editar", width=118, command=_edit_p).pack(side="left", padx=3)
    Btn(btn_p, text="Remover", width=118, command=_rm_p).pack(side="left", padx=3)
    Btn(btn_p, text="Ver desenho técnico", width=164, command=_open_p_desenho).pack(side="left", padx=3)
    Btn(btn_p, text="Resumo encomenda", width=164, command=lambda: open_encomenda_info_by_numero(self, (target.get("enc") or {}).get("numero", ""))).pack(side="left", padx=3)
    try:
        tbl_p.bind("<Double-1>", lambda _e=None: _open_p_desenho())
    except Exception:
        pass

    top_actions = ctk.CTkFrame(top, fg_color="transparent")
    top_actions.grid(row=2, column=0, columnspan=8, padx=(8, 8), pady=(0, 8), sticky="w")

    foot = ctk.CTkFrame(win, fg_color="white", corner_radius=10, border_width=1, border_color="#d8dee8")
    foot.pack(side="bottom", fill="x", padx=12, pady=(0, 12))

    def _editor_cativar():
        tgt = _ensure_target()
        if not tgt:
            return
        m = _find_m()
        e = _find_e()
        if not m and tbl_m.get_children():
            iid0 = tbl_m.get_children()[0]
            tbl_m.selection_set(iid0)
            sel["m"] = _find_m()
            _refresh_e()
            m = _find_m()
        if m and not e and tbl_e.get_children():
            iid0 = tbl_e.get_children()[0]
            tbl_e.selection_set(iid0)
            sel["e"] = _find_e()
            _refresh_p()
            e = _find_e()
        if not m or not e:
            messagebox.showerror("Erro", "Selecione material e espessura para cativar.")
            return
        mat_name = str(m.get("material", "") or "")
        esp_name = str(e.get("espessura", "") or "")
        _set_main_context(tgt, mat_name, esp_name)
        try:
            win.grab_release()
        except Exception:
            pass
        try:
            opened = self.cativar_stock(
                enc_override=tgt,
                mat_override=mat_name,
                esp_override=esp_name,
                parent_window=win,
            )
            if opened and hasattr(opened, "winfo_exists") and opened.winfo_exists():
                win.wait_window(opened)
        finally:
            try:
                if win.winfo_exists():
                    win.grab_set()
                    win.lift()
                    win.focus_force()
            except Exception:
                pass
        cativar_var.set(bool(tgt.get("reservas")))
        _sync_from_target(prefer_material=mat_name, prefer_esp=esp_name)

    def _editor_descativar():
        tgt = _ensure_target()
        if not tgt:
            return
        m = _find_m()
        e = _find_e()
        if not m or not e:
            messagebox.showerror("Erro", "Selecione material e espessura para descativar.")
            return
        mat_name = str(m.get("material", "") or "")
        esp_name = str(e.get("espessura", "") or "")
        _set_main_context(tgt, mat_name, esp_name)
        ok = self.descativar_stock_selecao(enc_override=tgt, mat_override=mat_name, esp_override=esp_name)
        if ok:
            cativar_var.set(bool(tgt.get("reservas")))
            _sync_from_target(prefer_material=mat_name, prefer_esp=esp_name)

    def _save():
        cliente = (cli_var.get() or "").strip()
        if not cliente:
            messagebox.showerror("Erro", "Cliente obrigatório.")
            return
        try:
            tempo = float((tempo_var.get() or "0").replace(",", "."))
        except Exception:
            messagebox.showerror("Erro", "Tempo estimado inválido.")
            return

        alvo = target.get("enc") if target.get("enc") else (enc if is_edit else None)
        if alvo is None:
            alvo = _ensure_target()
            if alvo is None:
                return

        alvo["cliente"] = cliente
        alvo["posto_trabalho"] = (posto_var.get() or "").strip()
        alvo["nota_cliente"] = (nota_var.get() or "").strip()
        alvo["data_entrega"] = (entrega_var.get() or "").strip()
        alvo["tempo_estimado"] = tempo
        alvo["tempo"] = tempo
        alvo["cativar"] = bool(cativar_var.get())
        obs_txt = (obs_var.get() or "").strip()
        alvo["ObservaÃ§Ãµes"] = obs_txt
        alvo["Observacoes"] = obs_txt
        alvo["estado"] = str(alvo.get("estado", "Preparacao") or "Preparacao")
        alvo["materiais"] = working.get("materiais", [])
        alvo.setdefault("reservas", [])
        alvo.setdefault("espessuras", [])
        alvo.setdefault("numero_orcamento", "")

        try:
            if not bool(alvo.get("cativar")):
                if alvo.get("reservas"):
                    aplicar_reserva_em_stock(self.data, alvo.get("reservas", []), -1)
                alvo["reservas"] = []
                alvo["cativar"] = False
            else:
                # cativação ativa só quando existem reservas reais (manual por chapa)
                alvo["cativar"] = bool(alvo.get("reservas"))
                cativar_var.set(alvo["cativar"])
            update_estado_encomenda_por_espessuras(alvo)
        except Exception:
            pass

        for p in encomenda_pecas(alvo):
            try:
                update_refs(self.data, p.get("ref_interna", ""), p.get("ref_externa", ""))
            except Exception:
                pass

        save_data(self.data)
        try:
            self.refresh_encomendas_year_options(keep_selection=True)
        except Exception:
            pass
        try:
            self.refresh_encomendas()
        except Exception:
            pass
        self.selected_encomenda_numero = alvo.get("numero")
        self.refresh()
        try:
            self._enc_editor_win = None
        except Exception:
            pass
        win.destroy()

    Btn(top_actions, text="Guardar", width=156, height=38, font=("Segoe UI", 13, "bold"), command=_save, fg_color="#f59e0b", hover_color="#d97706").pack(side="left", padx=4)
    Btn(top_actions, text="Cativar MP", width=126, height=34, font=("Segoe UI", 12, "bold"), command=_editor_cativar).pack(side="left", padx=4)
    Btn(top_actions, text="Descativar MP", width=133, height=34, font=("Segoe UI", 12, "bold"), command=_editor_descativar).pack(side="left", padx=4)
    def _close_editor():
        try:
            self._enc_editor_win = None
        except Exception:
            pass
        win.destroy()
    Btn(top_actions, text="Fechar", width=156, height=38, font=("Segoe UI", 13, "bold"), command=_close_editor).pack(side="left", padx=4)

    Btn(foot, text="Guardar", width=196, height=38, font=("Segoe UI", 13, "bold"), command=_save, fg_color="#f59e0b", hover_color="#d97706").pack(side="left", padx=8, pady=10)
    Btn(foot, text="Cancelar", width=168, height=38, font=("Segoe UI", 13, "bold"), command=_close_editor).pack(side="left", padx=8, pady=10)

    _refresh_m()
    _render_editor_reservas()


def add_encomenda(self):
    _ensure_configured()
    _open_encomenda_header_dialog(self, None)

def edit_encomenda(self):
    _ensure_configured()
    enc = self.get_selected_encomenda()
    if not enc:
        return
    _open_encomenda_header_dialog(self, enc)

def remove_encomenda(self):
    _ensure_configured()
    enc = self.get_selected_encomenda()
    if not enc:
        return
    if messagebox.askyesno("Confirmar", "Remover esta encomenda?"):
        if enc.get("reservas"):
            aplicar_reserva_em_stock(self.data, enc["reservas"], -1)
        self.data["encomendas"].remove(enc)
        save_data(self.data)
        try:
            self.refresh_encomendas_year_options(keep_selection=True)
        except Exception:
            pass
        self.refresh()

def on_select_encomenda(self, _):
    _ensure_configured()
    sel = self.tbl_encomendas.selection()
    if not sel:
        return
    numero = self.tbl_encomendas.item(sel[0], "values")[0]
    enc = self.get_encomenda_by_numero(numero)
    self.selected_encomenda_numero = numero
    self.selected_material = None
    self.selected_espessura = None
    if enc:
        try:
            if hasattr(self, "enc_info_numero"): self.enc_info_numero.set(str(enc.get("numero", "") or "-"))
            if hasattr(self, "enc_info_cliente"): self.enc_info_cliente.set(str(enc.get("cliente", "") or "-"))
            if hasattr(self, "enc_info_entrega"): self.enc_info_entrega.set(str(enc.get("data_entrega", "") or "-"))
            if hasattr(self, "enc_info_estado"): self.enc_info_estado.set(str(enc.get("estado", "") or "-"))
            if hasattr(self, "enc_info_nota"): self.enc_info_nota.set(str(enc.get("nota_cliente", "") or "-"))
        except Exception:
            pass
        self.e_cliente.set(enc.get("cliente", ""))
        self.e_nota_cliente.set(enc.get("nota_cliente", ""))
        self._update_nota_cliente_visual()
        self.e_data_entrega.set(enc.get("data_entrega", ""))
        self.e_tempo.set(enc.get("tempo_estimado", 0))
        self.e_cativar.set(bool(enc.get("cativar", False)))
        if hasattr(self, "update_cativar_button"):
            self.update_cativar_button()
        self.e_obs.set(enc.get("ObservaÃ§Ãµes", ""))
        self.e_chapa.set(self.get_chapa_reservada(enc["numero"]))
        self.update_cativadas_display(enc)
    self.refresh_materiais(enc)
    self.refresh_espessuras(enc, self.selected_material)
    self.refresh_pecas(enc, None)

def clear_encomenda_selection(self):
    _ensure_configured()
    self.selected_encomenda_numero = None
    self.selected_material = None
    self.selected_espessura = None
    self.e_cliente.set("")
    self.e_nota_cliente.set("")
    self._update_nota_cliente_visual()
    self.e_data_entrega.set("")
    self.e_tempo.set(0)
    self.e_cativar.set(False)
    if hasattr(self, "update_cativar_button"):
        self.update_cativar_button()
    self.e_obs.set("")
    self.e_chapa.set("")
    try:
        if hasattr(self, "enc_info_numero"): self.enc_info_numero.set("-")
        if hasattr(self, "enc_info_cliente"): self.enc_info_cliente.set("-")
        if hasattr(self, "enc_info_entrega"): self.enc_info_entrega.set("-")
        if hasattr(self, "enc_info_estado"): self.enc_info_estado.set("-")
        if hasattr(self, "enc_info_nota"): self.enc_info_nota.set("-")
    except Exception:
        pass
    self.update_cativadas_display(None)
    for t in [self.tbl_materiais, self.tbl_espessuras, self.tbl_pecas]:
        for i in t.get_children():
            t.delete(i)

def clear_espessura_selection(self):
    _ensure_configured()
    self.selected_espessura = None
    for i in self.tbl_pecas.get_children():
        self.tbl_pecas.delete(i)

def clear_peca_selection(self):
    _ensure_configured()
    pass

def save_nota_cliente_encomenda(self, _e=None):
    _ensure_configured()
    numero = (self.selected_encomenda_numero or "").strip()
    if not numero:
        self._update_nota_cliente_visual()
        return
    enc = self.get_encomenda_by_numero(numero)
    if not enc:
        self._update_nota_cliente_visual()
        return
    novo = (self.e_nota_cliente.get() or "").strip()
    atual = (enc.get("nota_cliente", "") or "").strip()
    if novo == atual:
        self._update_nota_cliente_visual()
        return
    enc["nota_cliente"] = novo
    save_data(self.data)
    self.refresh_encomendas()
    self._update_nota_cliente_visual()

def save_data_entrega_encomenda(self, _e=None):
    _ensure_configured()
    numero = (self.selected_encomenda_numero or "").strip()
    if not numero:
        return
    enc = self.get_encomenda_by_numero(numero)
    if not enc:
        return
    novo = (self.e_data_entrega.get() or "").strip()
    if novo and len(novo) < 10:
        return
    atual = (enc.get("data_entrega", "") or "").strip()
    if novo == atual:
        return
    enc["data_entrega"] = novo
    save_data(self.data)
    try:
        sel = self.tbl_encomendas.selection()
        if sel:
            iid = sel[0]
            vals = list(self.tbl_encomendas.item(iid, "values"))
            if len(vals) >= 5 and str(vals[0]) == numero:
                vals[4] = novo
                self.tbl_encomendas.item(iid, values=tuple(vals))
                return
    except Exception:
        pass
    self.refresh_encomendas()

def save_cativar_encomenda(self, _e=None):
    _ensure_configured()
    numero = (self.selected_encomenda_numero or "").strip()
    if not numero:
        return False
    enc = self.get_encomenda_by_numero(numero)
    if not enc:
        return False
    novo = bool(self.e_cativar.get())
    atual = bool(enc.get("cativar", False))
    if novo == atual:
        if hasattr(self, "update_cativar_button"):
            self.update_cativar_button()
        return True
    enc["cativar"] = novo
    try:
        if novo:
            # Cativação ativa sem recálculo automático.
            # As reservas são geridas manualmente por chapa/lote no fluxo Cativar MP.
            enc["cativar"] = bool(enc.get("reservas"))
            try:
                self.e_cativar.set(enc["cativar"])
            except Exception:
                pass
        else:
            # Ao desativar, liberta reservas existentes do stock.
            if enc.get("reservas"):
                aplicar_reserva_em_stock(self.data, enc.get("reservas", []), -1)
            enc["reservas"] = []
    except Exception:
        pass
    save_data(self.data)
    self.refresh_encomendas()
    try:
        self.e_chapa.set(self.get_chapa_reservada(enc.get("numero", "")))
        self.update_cativadas_display(enc)
    except Exception:
        pass
    if hasattr(self, "update_cativar_button"):
        self.update_cativar_button()
    return True

def get_encomenda_by_numero(self, numero):
    _ensure_configured()
    for e in self.data["encomendas"]:
        if e["numero"] == numero:
            return e
    return None

def on_select_espessura(self, _):
    _ensure_configured()
    enc = self.get_selected_encomenda()
    if not enc:
        return
    sel = self.tbl_espessuras.selection()
    if not sel:
        return
    esp = self.tbl_espessuras.item(sel[0], "values")[0]
    self.selected_espessura = esp
    if self.selected_material:
        self.e_chapa.set(self.get_chapa_reservada(enc["numero"], self.selected_material, esp))
    self.refresh_pecas(enc, esp)
    self.update_cativadas_display(enc)

def refresh_espessuras(self, enc, material):
    _ensure_configured()
    children = self.tbl_espessuras.get_children()
    if children:
        self.tbl_espessuras.delete(*children)
    if not enc:
        return
    selected = self.selected_espessura
    row = 0
    for m in enc.get("materiais", []):
        if material and m.get("material") != material:
            continue
        for e in m.get("espessuras", []):
            tag = "esp_even" if row % 2 == 0 else "esp_odd"
            item_id = self.tbl_espessuras.insert("", END, values=(
                e.get("espessura", ""), e.get("tempo_min", ""), e.get("estado", "Preparacao")
            ), tags=(tag,))
            if selected and e.get("espessura") == selected:
                self.tbl_espessuras.selection_set(item_id)
            row += 1
    self.tbl_espessuras.tag_configure("esp_even", background="#eef7e6")
    self.tbl_espessuras.tag_configure("esp_odd", background="#f7fbf2")

def get_selected_encomenda(self):
    _ensure_configured()
    sel = self.tbl_encomendas.selection()
    if not sel:
        numero = str(getattr(self, "selected_encomenda_numero", "") or "").strip()
        if numero:
            return self.get_encomenda_by_numero(numero)
        messagebox.showerror("Erro", "Selecione uma encomenda")
        return None
    numero = self.tbl_encomendas.item(sel[0], "values")[0]
    return self.get_encomenda_by_numero(numero)

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
    self.selected_espessura = None
    self.e_chapa.set(self.get_chapa_reservada(enc["numero"], mat))
    self.refresh_espessuras(enc, mat)
    self.refresh_pecas(enc, None)

def refresh_encomendas(self):
    _ensure_configured()
    if not hasattr(self, "tbl_encomendas") or self.tbl_encomendas is None:
        return
    if getattr(self, "_refresh_encomendas_busy", False):
        self._refresh_encomendas_pending = True
        return
    self._refresh_encomendas_busy = True
    self._refresh_encomendas_pending = False
    try:
        query = (self.e_filter.get().strip().lower() if hasattr(self, "e_filter") else "")
        estado_filtro = (self.e_estado_filter.get().strip().lower() if hasattr(self, "e_estado_filter") else "ativas")
        ano_filtro = (self.e_year_filter.get().strip() if hasattr(self, "e_year_filter") else "Todos")
        cliente_filtro_raw = (self.e_cliente_filter.get().strip() if hasattr(self, "e_cliente_filter") else "Todos")
        cliente_filtro_norm = norm_text(cliente_filtro_raw)
        cliente_filtro_codigo = ""
        if cliente_filtro_norm not in ("", "todos", "todas", "all"):
            try:
                if "_extract_cliente_codigo" in globals():
                    cliente_filtro_codigo = _extract_cliente_codigo(cliente_filtro_raw, self.data) or ""
            except Exception:
                cliente_filtro_codigo = ""
            if not cliente_filtro_codigo:
                cliente_filtro_codigo = cliente_filtro_raw.split(" - ", 1)[0].strip()
        ano_norm = ano_filtro.lower()
        try:
            if hasattr(self, "e_estado_segment") and self.e_estado_segment.winfo_exists():
                desired = self.e_estado_filter.get() or "Ativas"
                current = ""
                try:
                    current = self.e_estado_segment.get() or ""
                except Exception:
                    current = ""
                if current != desired:
                    self._suppress_encomendas_filter_cb = True
                    self.e_estado_segment.set(desired)
        except Exception:
            pass
        finally:
            self._suppress_encomendas_filter_cb = False
        selected = self.selected_encomenda_numero
        clientes_nome = {}
        try:
            clientes_nome = {
                str(c.get("codigo", "") or "").strip(): str(c.get("nome", "") or "")
                for c in self.data.get("clientes", [])
                if isinstance(c, dict)
            }
        except Exception:
            clientes_nome = {}
        rows = []
        for idx, e in enumerate(self.data["encomendas"]):
            enc_year = _enc_extract_year(e.get("data_criacao", ""), e.get("data_entrega", ""), e.get("numero", ""), e.get("ano"))
            if ano_norm not in ("todos", "todas", "all") and ano_filtro and enc_year != ano_filtro:
                continue
            if cliente_filtro_codigo and str(e.get("cliente", "")).strip() != cliente_filtro_codigo:
                continue
            cliente_nome = clientes_nome.get(str(e.get("cliente", "") or "").strip(), "")
            cliente_display = f"{e['cliente']} - {cliente_nome}" if cliente_nome else e["cliente"]
            values = (
                e["numero"], e.get("nota_cliente", ""), cliente_display, e["data_criacao"], e["data_entrega"],
                e["tempo_estimado"], e["estado"], "SIM" if e["cativar"] else "NAO"
            )
            if query and not any(query in str(v).lower() for v in values):
                continue
            est_norm = norm_text(e.get("estado", ""))
            if estado_filtro and estado_filtro not in ("todas", "todos", "all"):
                if "ativ" in estado_filtro:
                    if "concl" in est_norm:
                        continue
                elif "prepar" in estado_filtro and "prepar" not in est_norm:
                    continue
                elif "produ" in estado_filtro and "produ" not in est_norm:
                    continue
                elif "concl" in estado_filtro and "concl" not in est_norm:
                    continue
            alt_tag = "odd" if idx % 2 else "even"
            est_norm_row = norm_text(e.get("estado", ""))
            if "concl" in est_norm_row:
                estado_tag = f"estado_concluida_{alt_tag}"
            elif "produ" in est_norm_row:
                estado_tag = f"estado_producao_{alt_tag}"
            else:
                estado_tag = f"estado_preparacao_{alt_tag}"
            rows.append((values, (alt_tag, estado_tag)))
        try:
            vals = self.get_clientes_codes()
            if self.encomendas_use_custom and CUSTOM_TK_AVAILABLE and isinstance(self.e_cliente_cb, ctk.CTkComboBox):
                self.e_cliente_cb.configure(values=vals or [""])
            else:
                self.e_cliente_cb["values"] = vals
        except Exception:
            pass
        try:
            filtros_cli = ["Todos"] + (self.get_clientes_display() or [])
            self._suppress_encomendas_filter_cb = True
            if self.encomendas_use_custom and CUSTOM_TK_AVAILABLE and isinstance(self.e_cliente_filter_cb, ctk.CTkComboBox):
                self.e_cliente_filter_cb.configure(values=filtros_cli)
                if not (self.e_cliente_filter.get() or "").strip():
                    self.e_cliente_filter.set("Todos")
            else:
                self.e_cliente_filter_cb["values"] = filtros_cli
                if not (self.e_cliente_filter.get() or "").strip():
                    self.e_cliente_filter.set("Todos")
        except Exception:
            pass
        finally:
            self._suppress_encomendas_filter_cb = False
        def _after_fill():
            try:
                if selected:
                    for iid in self.tbl_encomendas.get_children():
                        vals = self.tbl_encomendas.item(iid, "values")
                        if vals and str(vals[0] or "") == str(selected):
                            self.tbl_encomendas.selection_set(iid)
                            self.tbl_encomendas.see(iid)
                            break
            except Exception:
                pass
            try:
                enc = self.get_encomenda_by_numero(selected) if selected else None
                self.update_cativadas_display(enc)
            except Exception:
                pass

        if hasattr(self, "fill_treeview_in_batches"):
            self.fill_treeview_in_batches(self.tbl_encomendas, rows, "tbl_encomendas", chunk_size=120, on_done=_after_fill)
        else:
            children = self.tbl_encomendas.get_children()
            if children:
                self.tbl_encomendas.delete(*children)
            for values, tags in rows:
                self.tbl_encomendas.insert("", END, values=values, tags=tags)
            _after_fill()
    finally:
        self._refresh_encomendas_busy = False
        if getattr(self, "_refresh_encomendas_pending", False):
            self._refresh_encomendas_pending = False
            try:
                self.root.after(15, self.refresh_encomendas)
            except Exception:
                pass

def refresh_pecas(self, enc, espessura):
    _ensure_configured()
    children = self.tbl_pecas.get_children()
    if children:
        self.tbl_pecas.delete(*children)
    if not enc:
        return
    esp_map = {}
    esp_colors = ["esp_1", "esp_2", "esp_3", "esp_4"]
    esp_keys = sorted({p.get("espessura", "") for p in encomenda_pecas(enc)})
    for i, e in enumerate(esp_keys):
        esp_map[e] = esp_colors[i % len(esp_colors)]
    pecas = encomenda_pecas(enc)
    if espessura:
        pecas = [p for p in pecas if p.get("espessura") == espessura]
    for idx, p in enumerate(pecas):
        alt_tag = "odd" if idx % 2 else "even"
        est_norm_row = norm_text(p.get("estado", ""))
        if "concl" in est_norm_row:
            estado_tag = f"estado_concluida_{alt_tag}"
        elif "produ" in est_norm_row:
            estado_tag = f"estado_producao_{alt_tag}"
        else:
            estado_tag = f"estado_preparacao_{alt_tag}"
        esp_tag = esp_map.get(p.get("espessura", ""), alt_tag)
        self.tbl_pecas.insert("", END, values=(
            p["ref_interna"], p["ref_externa"], p["material"], p["espessura"],
            p.get("OperaÃ§Ãµes", ""), p["quantidade_pedida"], p["estado"]
        ), tags=(esp_tag, estado_tag))

def reselect_encomenda_material_espessura(self):
    _ensure_configured()
    if self.selected_encomenda_numero:
        for item in self.tbl_encomendas.get_children():
            if self.tbl_encomendas.item(item, "values")[0] == self.selected_encomenda_numero:
                self.tbl_encomendas.selection_set(item)
                self.tbl_encomendas.see(item)
                break
    if self.selected_material:
        for item in self.tbl_materiais.get_children():
            if self.tbl_materiais.item(item, "values")[0] == self.selected_material:
                self.tbl_materiais.selection_set(item)
                break
    if self.selected_espessura:
        for item in self.tbl_espessuras.get_children():
            if self.tbl_espessuras.item(item, "values")[0] == self.selected_espessura:
                self.tbl_espessuras.selection_set(item)
                break

def add_material_encomenda(self):
    _ensure_configured()
    enc = self.get_selected_encomenda()
    if not enc:
        return
    if _is_orc_based_encomenda(enc):
        messagebox.showinfo("Info", "Encomenda originada de orÃ§amento: material bloqueado. Crie uma encomenda manual para editar estrutura.")
        return
    use_custom = self.encomendas_use_custom and CUSTOM_TK_AVAILABLE
    parent_win = getattr(self, "_enc_editor_win", None) or self.root
    win = ctk.CTkToplevel(parent_win) if use_custom else Toplevel(parent_win)
    win.title("Adicionar material")
    try:
        win.geometry("460x190" if use_custom else "420x170")
        win.resizable(False, False)
    except Exception:
        pass
    try:
        win.transient(parent_win)
        win.grab_set()
        win.lift()
        win.focus_force()
    except Exception:
        pass
    Lbl = ctk.CTkLabel if use_custom else ttk.Label
    Cmb = ctk.CTkComboBox if use_custom else ttk.Combobox
    Btn = ctk.CTkButton if use_custom else ttk.Button
    Lbl(win, text="Material").grid(row=0, column=0, sticky="w", padx=8, pady=6)
    mat_var = StringVar()
    if use_custom:
        mat_cb = Cmb(win, variable=mat_var, values=MATERIAIS_PRESET, width=260)
    else:
        mat_cb = Cmb(win, textvariable=mat_var, values=MATERIAIS_PRESET, width=30, state="normal")
    mat_cb.grid(row=0, column=1, padx=8, pady=6, sticky="w")

    Lbl(win, text="Acabamento inox").grid(row=1, column=0, sticky="w", padx=8, pady=6)
    acab_var = StringVar()
    if use_custom:
        acab_cb = Cmb(win, variable=acab_var, values=INOX_ACABAMENTOS, width=260)
    else:
        acab_cb = Cmb(win, textvariable=acab_var, values=INOX_ACABAMENTOS, width=30, state="normal")
    acab_cb.grid(row=1, column=1, padx=8, pady=6, sticky="w")

    def on_save():
        mat = mat_var.get().strip()
        if not mat:
            messagebox.showerror("Erro", "Material obrigatorio")
            return
        if mat in ("AISI 304L", "AISI 316L"):
            acabamento = acab_var.get().strip()
            if acabamento in INOX_ACABAMENTOS:
                mat_final = f"{mat} - {acabamento}"
            else:
                mat_final = mat
        else:
            mat_final = mat
        if any(m.get("material") == mat_final for m in enc.get("materiais", [])):
            messagebox.showerror("Erro", "Material ja existe")
            return
        enc.setdefault("materiais", []).append({
            "material": mat_final,
            "estado": "Preparacao",
            "espessuras": [],
        })
        save_data(self.data)
        self.refresh_materiais(enc)
        win.destroy()

    Btn(win, text="Guardar", command=on_save, width=140 if use_custom else None).grid(row=2, column=0, columnspan=2, pady=10)
    try:
        mat_cb.focus_set()
    except Exception:
        pass
    return win

def remove_material_encomenda(self):
    _ensure_configured()
    enc = self.get_selected_encomenda()
    if not enc:
        return
    if _is_orc_based_encomenda(enc):
        messagebox.showinfo("Info", "Encomenda originada de orÃ§amento: material bloqueado.")
        return
    sel = self.tbl_materiais.selection()
    if not sel:
        messagebox.showerror("Erro", "Selecione um material")
        return
    mat = self.tbl_materiais.item(sel[0], "values")[0]
    enc["materiais"] = [m for m in enc.get("materiais", []) if m.get("material") != mat]
    save_data(self.data)
    self.refresh_materiais(enc)
    self.refresh_espessuras(enc, self.selected_material)
    self.refresh_pecas(enc, None)

def add_espessura(self, enc_override=None, material_override=None):
    _ensure_configured()
    enc = enc_override or self.get_selected_encomenda()
    if not enc:
        return
    if _is_orc_based_encomenda(enc):
        messagebox.showinfo("Info", "Encomenda originada de orÃ§amento: espessuras bloqueadas.")
        return
    selected_mat = str(material_override or getattr(self, "selected_material", "") or "").strip()
    if not selected_mat:
        messagebox.showerror("Erro", "Selecione um material")
        return
    use_custom = self.encomendas_use_custom and CUSTOM_TK_AVAILABLE
    parent_win = getattr(self, "_enc_editor_win", None) or self.root
    win = ctk.CTkToplevel(parent_win) if use_custom else Toplevel(parent_win)
    win.title("Adicionar espessura")
    try:
        win.geometry("390x150" if use_custom else "360x140")
        win.resizable(False, False)
    except Exception:
        pass
    try:
        win.transient(parent_win)
        win.grab_set()
        win.lift()
        win.focus_force()
    except Exception:
        pass
    Lbl = ctk.CTkLabel if use_custom else ttk.Label
    Cmb = ctk.CTkComboBox if use_custom else ttk.Combobox
    Btn = ctk.CTkButton if use_custom else ttk.Button
    Lbl(win, text="Espessura (mm)").grid(row=0, column=0, sticky="w", padx=8, pady=6)
    esp_var = StringVar()
    esp_opts = [str(v).rstrip("0").rstrip(".") if isinstance(v, float) else str(v) for v in ESPESSURAS_PRESET]
    if use_custom:
        esp_cb = Cmb(win, variable=esp_var, values=esp_opts, width=220)
    else:
        esp_cb = Cmb(win, textvariable=esp_var, values=esp_opts, width=24, state="normal")
    esp_cb.grid(row=0, column=1, padx=8, pady=6, sticky="w")

    def on_save():
        esp = esp_var.get().strip()
        if not esp:
            messagebox.showerror("Erro", "Espessura obrigatoria")
            return
        mat_obj = None
        sel_mat = str(material_override or getattr(self, "selected_material", "") or "").strip()
        for m in enc.get("materiais", []):
            if _match_material(str(m.get("material", "") or "").strip(), sel_mat):
                mat_obj = m
                sel_mat = str(m.get("material", "") or "").strip()
                break
        # Fallback robusto para fluxos em janela editor custom
        if mat_obj is None:
            try:
                if hasattr(self, "tbl_materiais"):
                    s = self.tbl_materiais.selection()
                    if s:
                        vals = self.tbl_materiais.item(s[0], "values")
                        mat_name = str(vals[0] if vals else "").strip()
                        for m in enc.get("materiais", []):
                            if _match_material(str(m.get("material", "") or "").strip(), mat_name):
                                mat_obj = m
                                self.selected_material = str(m.get("material", "") or "").strip()
                                break
            except Exception:
                pass
        if mat_obj is None and len(enc.get("materiais", [])) == 1:
            mat_obj = enc.get("materiais", [])[0]
            try:
                self.selected_material = str(mat_obj.get("material", "") or "").strip()
            except Exception:
                pass
        if not mat_obj:
            messagebox.showerror("Erro", "Material nao encontrado")
            return
        if any(e.get("espessura") == esp for e in mat_obj.get("espessuras", [])):
            messagebox.showerror("Erro", "Espessura ja existe")
            return
        mat_obj.setdefault("espessuras", []).append({
            "espessura": esp,
            "tempo_min": "",
            "estado": "Preparacao",
            "pecas": [],
        })
        save_data(self.data)
        try:
            self.selected_material = str(mat_obj.get("material", "") or "").strip()
        except Exception:
            pass
        self.refresh_espessuras(enc, self.selected_material)
        win.destroy()

    Btn(win, text="Guardar", command=on_save, width=140 if use_custom else None).grid(row=1, column=0, columnspan=2, pady=10)
    try:
        esp_cb.focus_set()
    except Exception:
        pass
    return win

def remove_espessura(self):
    _ensure_configured()
    enc = self.get_selected_encomenda()
    if not enc:
        return
    if _is_orc_based_encomenda(enc):
        messagebox.showinfo("Info", "Encomenda originada de orÃ§amento: espessuras bloqueadas.")
        return
    sel = self.tbl_espessuras.selection()
    if not sel:
        messagebox.showerror("Erro", "Selecione uma espessura")
        return
    esp = self.tbl_espessuras.item(sel[0], "values")[0]
    for m in enc.get("materiais", []):
        if m.get("material") == self.selected_material:
            m["espessuras"] = [e for e in m.get("espessuras", []) if e.get("espessura") != esp]
            break
    save_data(self.data)
    self.refresh_espessuras(enc, self.selected_material)
    self.refresh_pecas(enc, None)

def remove_peca(self):
    _ensure_configured()
    enc = self.get_selected_encomenda()
    if not enc:
        return
    sel = self.tbl_pecas.selection()
    if not sel:
        messagebox.showerror("Erro", "Selecione uma peÃ§a")
        return
    ref = self.tbl_pecas.item(sel[0], "values")[0]
    for m in enc.get("materiais", []):
        for e in m.get("espessuras", []):
            e["pecas"] = [p for p in e.get("pecas", []) if p.get("ref_interna") != ref]
    save_data(self.data)
    self.refresh_pecas(enc, self.selected_espessura)

def _find_peca_in_encomenda(enc, ref_interna="", ref_externa=""):
    ref_int = str(ref_interna or "").strip()
    ref_ext = str(ref_externa or "").strip()
    if not enc:
        return None
    for p in encomenda_pecas(enc):
        p_ref_int = str(p.get("ref_interna", "") or "").strip()
        p_ref_ext = str(p.get("ref_externa", "") or "").strip()
        if ref_int and p_ref_int == ref_int:
            return p
        if (not ref_int) and ref_ext and p_ref_ext == ref_ext:
            return p
        if ref_int and ref_ext and p_ref_int == ref_int and p_ref_ext == ref_ext:
            return p
    return None

def _open_peca_desenho_path(path):
    desenho = str(path or "").strip()
    if not desenho:
        messagebox.showinfo("Info", "Esta peça não tem desenho associado.")
        return False
    if not os.path.exists(desenho):
        messagebox.showerror("Erro", f"Ficheiro não encontrado:\n{desenho}")
        return False
    try:
        os.startfile(desenho)
        return True
    except Exception as ex:
        messagebox.showerror("Erro", f"Não foi possível abrir o desenho.\n{ex}")
        return False

def open_peca_desenho_by_refs(self, numero=None, ref_interna="", ref_externa="", enc=None, silent=False):
    _ensure_configured()
    enc_obj = enc or self.get_encomenda_by_numero(str(numero or "").strip())
    if not enc_obj:
        if not silent:
            messagebox.showerror("Erro", "Encomenda não encontrada.")
        return False
    peca = _find_peca_in_encomenda(enc_obj, ref_interna=ref_interna, ref_externa=ref_externa)
    if not peca:
        if not silent:
            messagebox.showerror("Erro", "Peça não encontrada.")
        return False
    return _open_peca_desenho_path(peca.get("desenho", ""))

def open_encomenda_by_numero(self, numero, open_editor=False):
    _ensure_configured()
    numero_txt = str(numero or "").strip()
    if not numero_txt:
        return False
    enc = self.get_encomenda_by_numero(numero_txt)
    if not enc:
        messagebox.showerror("Erro", f"Encomenda não encontrada: {numero_txt}")
        return False
    try:
        if hasattr(self, "e_filter"):
            self.e_filter.set("")
        if hasattr(self, "e_estado_filter"):
            self.e_estado_filter.set("Todas")
        if hasattr(self, "e_cliente_filter"):
            self.e_cliente_filter.set("Todos")
        enc_year = _enc_extract_year(
            enc.get("data_criacao", ""),
            enc.get("data_entrega", ""),
            enc.get("numero", ""),
            enc.get("ano"),
        )
        if enc_year and hasattr(self, "e_year_filter"):
            self.e_year_filter.set(enc_year)
        if hasattr(self, "tab_encomendas"):
            self.navigate_to_tab(self.tab_encomendas)
    except Exception:
        pass
    try:
        self.refresh_encomendas()
    except Exception:
        pass
    item_found = None
    try:
        for item in self.tbl_encomendas.get_children():
            vals = self.tbl_encomendas.item(item, "values")
            if vals and str(vals[0] or "").strip() == numero_txt:
                item_found = item
                break
    except Exception:
        item_found = None
    if item_found is not None:
        try:
            self.tbl_encomendas.selection_set(item_found)
            self.tbl_encomendas.focus(item_found)
            self.tbl_encomendas.see(item_found)
        except Exception:
            pass
    self.selected_encomenda_numero = numero_txt
    try:
        self.on_select_encomenda(None)
    except Exception:
        pass
    if open_editor:
        try:
            _open_encomenda_header_dialog(self, enc)
        except Exception:
            try:
                self.edit_encomenda()
            except Exception:
                return False
    return True

def open_encomenda_info_by_numero(self, numero, highlight_ref=None):
    _ensure_configured()
    numero_txt = str(numero or "").strip()
    if not numero_txt:
        return False
    enc = self.get_encomenda_by_numero(numero_txt)
    if not enc:
        messagebox.showerror("Erro", f"Encomenda não encontrada: {numero_txt}")
        return False

    old = getattr(self, "_enc_info_view_win", None)
    try:
        if old is not None and old.winfo_exists():
            old.destroy()
    except Exception:
        pass

    use_custom = self.encomendas_use_custom and CUSTOM_TK_AVAILABLE
    Win = ctk.CTkToplevel if use_custom else Toplevel
    Frame = ctk.CTkFrame if use_custom else ttk.Frame
    Label = ctk.CTkLabel if use_custom else ttk.Label
    Btn = ctk.CTkButton if use_custom else ttk.Button

    def _btn(parent, **kwargs):
        if use_custom:
            kwargs.setdefault("height", 36)
            kwargs.setdefault("corner_radius", 10)
            kwargs.setdefault("fg_color", CTK_PRIMARY_RED)
            kwargs.setdefault("hover_color", CTK_PRIMARY_RED_HOVER)
            kwargs.setdefault("text_color", "#ffffff")
            kwargs.setdefault("font", ("Segoe UI", 12, "bold"))
        return Btn(parent, **kwargs)

    def _info_box(parent, title, value, width=200):
        card = Frame(
            parent,
            fg_color="#f8fbff",
            corner_radius=10,
            border_width=1,
            border_color="#d8dee8",
            width=width,
        ) if use_custom else Frame(parent)
        if use_custom:
            card.pack_propagate(False)
        Label(
            card,
            text=title,
            font=("Segoe UI", 11, "bold") if use_custom else ("Segoe UI", 10, "bold"),
            text_color="#64748b" if use_custom else None,
            anchor="w",
        ).pack(fill="x", padx=10, pady=(8, 2))
        Label(
            card,
            text=str(value or "-"),
            font=("Segoe UI", 13, "bold") if use_custom else ("Segoe UI", 10),
            text_color="#0f172a" if use_custom else None,
            anchor="w",
            wraplength=max(120, width - 24),
            justify="left",
        ).pack(fill="x", padx=10, pady=(0, 8))
        return card

    win = Win(self.root)
    win.title(f"Informações da Encomenda {numero_txt}")
    try:
        sw = int(win.winfo_screenwidth() or 1366)
        sh = int(win.winfo_screenheight() or 768)
        ww = max(1280, min(1580, sw - 50))
        wh = max(820, min(960, sh - 70))
        pos_x = max(0, (sw - ww) // 2)
        pos_y = max(0, (sh - wh) // 2)
        win.geometry(f"{ww}x{wh}+{pos_x}+{pos_y}")
        win.minsize(1220, 760)
    except Exception:
        win.geometry("1440x880")
    try:
        if use_custom:
            win.configure(fg_color="#f5f7fb")
        win.transient(self.root)
        win.grab_set()
        win.lift()
        win.focus_force()
    except Exception:
        pass
    self._enc_info_view_win = win

    hero = Frame(
        win,
        fg_color="#eaf2ff",
        corner_radius=12,
        border_width=1,
        border_color="#c9d8ef",
    ) if use_custom else Frame(win)
    hero.pack(fill="x", padx=12, pady=(12, 8))
    Label(
        hero,
        text=f"Encomenda {numero_txt}",
        font=("Segoe UI", 20, "bold") if use_custom else ("Segoe UI", 14, "bold"),
        text_color="#0b1f4d" if use_custom else None,
    ).pack(side="left", padx=12, pady=8)
    Label(
        hero,
        text=f"Cliente: {enc.get('cliente', '-')}",
        font=("Segoe UI", 13, "bold") if use_custom else ("Segoe UI", 10, "bold"),
        text_color="#334155" if use_custom else None,
    ).pack(side="left", padx=6, pady=8)
    _btn(hero, text="Abrir editor", width=130 if use_custom else None, command=lambda: open_encomenda_by_numero(self, numero_txt, open_editor=True)).pack(side="right", padx=(6, 12), pady=8)
    _btn(hero, text="Fechar", width=120 if use_custom else None, command=win.destroy).pack(side="right", padx=6, pady=8)

    cards = Frame(win, fg_color="transparent") if use_custom else Frame(win)
    cards.pack(fill="x", padx=12, pady=(0, 8))
    summary = [
        ("Estado", enc.get("estado", "-")),
        ("Entrega", enc.get("data_entrega", "-")),
        ("Nota Cliente", enc.get("nota_cliente", "-")),
        ("Tempo Planeado", f"{parse_float(enc.get('tempo_estimado', 0), 0.0):g} h"),
        ("Cativar MP", "SIM" if bool(enc.get("cativar")) else "NAO"),
        ("Orçamento", enc.get("numero_orcamento", "-") or "-"),
    ]
    for title, value in summary:
        box = _info_box(cards, title, value, width=230)
        box.pack(side="left", fill="y", padx=(0, 8), pady=2)

    mid = Frame(win, fg_color="transparent") if use_custom else Frame(win)
    mid.pack(fill="x", padx=12, pady=(0, 8))

    obs_card = Frame(
        mid,
        fg_color="#ffffff",
        corner_radius=10,
        border_width=1,
        border_color="#d8dee8",
    ) if use_custom else Frame(mid)
    obs_card.pack(side="left", fill="both", expand=True, padx=(0, 6))
    Label(obs_card, text="Observações", font=("Segoe UI", 13, "bold") if use_custom else ("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=(8, 4))
    obs_txt = ctk.CTkTextbox(obs_card, height=90, wrap="word", fg_color="#f8fbff", border_width=1, border_color="#d8dee8") if use_custom else Text(obs_card, height=5, wrap="word")
    obs_txt.pack(fill="both", expand=True, padx=10, pady=(0, 8))
    try:
        obs_txt.insert("1.0", str(enc.get("Observacoes", "") or enc.get("ObservaÃ§Ãµes", "") or "Sem observações."))
        obs_txt.configure(state="disabled")
    except Exception:
        pass

    res_card = Frame(
        mid,
        fg_color="#ffffff",
        corner_radius=10,
        border_width=1,
        border_color="#d8dee8",
    ) if use_custom else Frame(mid)
    res_card.pack(side="left", fill="both", expand=True, padx=(6, 0))
    Label(res_card, text="Cativações / Reservas", font=("Segoe UI", 13, "bold") if use_custom else ("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=(8, 4))
    res_txt = ctk.CTkTextbox(res_card, height=90, wrap="word", fg_color="#f8fbff", border_width=1, border_color="#d8dee8") if use_custom else Text(res_card, height=5, wrap="word")
    res_txt.pack(fill="both", expand=True, padx=10, pady=(0, 8))
    reservas_lines = []
    for r in (enc.get("reservas", []) or []):
        reservas_lines.append(
            f"{r.get('material','-')} {r.get('espessura','-')} | Qtd: {parse_float(r.get('quantidade', 0), 0.0):g}"
        )
    if not reservas_lines:
        reservas_lines = ["Sem reservas/cativações."]
    try:
        res_txt.insert("1.0", "\n".join(reservas_lines))
        res_txt.configure(state="disabled")
    except Exception:
        pass

    grid_card = Frame(
        win,
        fg_color="#ffffff",
        corner_radius=10,
        border_width=1,
        border_color="#d8dee8",
    ) if use_custom else Frame(win)
    grid_card.pack(fill="both", expand=True, padx=12, pady=(0, 12))

    header = Frame(grid_card, fg_color="transparent") if use_custom else Frame(grid_card)
    header.pack(fill="x", padx=10, pady=(8, 4))
    Label(header, text="Referências da Encomenda", font=("Segoe UI", 14, "bold") if use_custom else ("Segoe UI", 11, "bold")).pack(side="left")
    Label(header, text="Duplo clique numa referência para abrir o desenho técnico.", font=("Segoe UI", 11) if use_custom else ("Segoe UI", 9), text_color="#64748b" if use_custom else None).pack(side="right")

    try:
        sty = ttk.Style(win)
        sty.theme_use("clam")
    except Exception:
        sty = ttk.Style(win)
    sty.configure(
        "EncInfo.Treeview",
        font=("Segoe UI", 10),
        rowheight=28,
        background="#f8fbff",
        fieldbackground="#f8fbff",
        borderwidth=0,
        relief="flat",
    )
    sty.configure(
        "EncInfo.Treeview.Heading",
        font=("Segoe UI", 10, "bold"),
        background=THEME_HEADER_BG,
        foreground="#ffffff",
        relief="flat",
    )
    sty.map("EncInfo.Treeview.Heading", background=[("active", THEME_HEADER_ACTIVE)])
    sty.map(
        "EncInfo.Treeview",
        background=[("selected", THEME_SELECT_BG)],
        foreground=[("selected", THEME_SELECT_FG)],
    )

    table_wrap = Frame(grid_card, fg_color="transparent") if use_custom else Frame(grid_card)
    table_wrap.pack(fill="both", expand=True, padx=10, pady=(0, 8))

    tbl = ttk.Treeview(
        table_wrap,
        columns=("refi", "refe", "material", "esp", "qtd", "ops", "estado", "desenho"),
        show="headings",
        style="EncInfo.Treeview",
    )
    headings = {
        "refi": ("Ref. Interna", 150),
        "refe": ("Ref. Externa", 200),
        "material": ("Material", 120),
        "esp": ("Esp.", 70),
        "qtd": ("Qtd", 70),
        "ops": ("Operações", 300),
        "estado": ("Estado", 110),
        "desenho": ("Desenho", 90),
    }
    for col, (txt, width) in headings.items():
        tbl.heading(col, text=txt)
        tbl.column(col, width=width, anchor="w" if col not in {"qtd", "esp", "desenho"} else "center", stretch=True)
    tbl.pack(side="left", fill="both", expand=True)
    sb_y = ctk.CTkScrollbar(table_wrap, orientation="vertical", command=tbl.yview) if use_custom else ttk.Scrollbar(table_wrap, orient="vertical", command=tbl.yview)
    sb_x = ctk.CTkScrollbar(grid_card, orientation="horizontal", command=tbl.xview) if use_custom else ttk.Scrollbar(grid_card, orient="horizontal", command=tbl.xview)
    tbl.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
    sb_y.pack(side="right", fill="y")
    sb_x.pack(fill="x", padx=10, pady=(0, 8))

    def _open_selected_desenho():
        sel = tbl.selection()
        if not sel:
            messagebox.showerror("Erro", "Selecione uma referência.")
            return
        vals = tbl.item(sel[0], "values")
        refi = str(vals[0] if len(vals) > 0 else "").strip()
        refe = str(vals[1] if len(vals) > 1 else "").strip()
        open_peca_desenho_by_refs(self, enc=enc, ref_interna=refi, ref_externa=refe)

    action_bar = Frame(grid_card, fg_color="transparent") if use_custom else Frame(grid_card)
    action_bar.pack(fill="x", padx=10, pady=(0, 10))
    _btn(action_bar, text="Ver desenho selecionado", width=190 if use_custom else None, command=_open_selected_desenho).pack(side="left", padx=(0, 6))
    _btn(action_bar, text="Abrir editor da encomenda", width=190 if use_custom else None, command=lambda: open_encomenda_by_numero(self, numero_txt, open_editor=True)).pack(side="left", padx=6)

    selected_iid = None
    row_index = 0
    for m in (enc.get("materiais", []) or []):
        for e in (m.get("espessuras", []) or []):
            for p in (e.get("pecas", []) or []):
                refi = str(p.get("ref_interna", "") or "")
                iid = tbl.insert(
                    "",
                    END,
                    values=(
                        refi,
                        str(p.get("ref_externa", "") or ""),
                        str(m.get("material", "") or p.get("material", "") or ""),
                        str(e.get("espessura", "") or p.get("espessura", "") or ""),
                        f"{parse_float(p.get('quantidade_pedida', 0), 0.0):g}",
                        str(p.get("Operacoes", "") or p.get("OperaÃ§Ãµes", "") or ""),
                        str(p.get("estado", "Preparacao") or "Preparacao"),
                        "SIM" if str(p.get("desenho", "") or "").strip() else "NAO",
                    ),
                    tags=("even" if row_index % 2 == 0 else "odd",),
                )
                if highlight_ref and str(highlight_ref).strip() == refi:
                    selected_iid = iid
                row_index += 1

    def _on_close():
        try:
            self._enc_info_view_win = None
        except Exception:
            pass
        win.destroy()

    try:
        tbl.bind("<Double-1>", lambda _e=None: _open_selected_desenho())
        win.protocol("WM_DELETE_WINDOW", _on_close)
    except Exception:
        pass
    if selected_iid:
        try:
            tbl.selection_set(selected_iid)
            tbl.see(selected_iid)
        except Exception:
            pass
    return True

def open_selected_peca_desenho(self):
    _ensure_configured()
    enc = self.get_selected_encomenda()
    if not enc:
        messagebox.showerror("Erro", "Selecione uma encomenda.")
        return
    sel = self.tbl_pecas.selection()
    if not sel:
        messagebox.showerror("Erro", "Selecione uma peÃ§a.")
        return
    vals = self.tbl_pecas.item(sel[0], "values")
    ref_int = str(vals[0] if len(vals) > 0 else "").strip()
    ref_ext = str(vals[1] if len(vals) > 1 else "").strip()
    open_peca_desenho_by_refs(self, enc=enc, ref_interna=ref_int, ref_externa=ref_ext)

def add_peca(self, esp_override=None):
    _ensure_configured()
    enc = self.get_selected_encomenda()
    if not enc:
        messagebox.showerror("Erro", "Selecione uma encomenda.")
        return

    sel_mat_default = self.selected_material or ""
    if not sel_mat_default and hasattr(self, "tbl_materiais"):
        sel_mat = self.tbl_materiais.selection()
        if sel_mat:
            vals = self.tbl_materiais.item(sel_mat[0], "values")
            sel_mat_default = str(vals[0] if vals else "").strip()

    sel_esp_default = str(esp_override or self.selected_espessura or "").strip()
    if not sel_esp_default and hasattr(self, "tbl_espessuras"):
        sel_esp = self.tbl_espessuras.selection()
        if sel_esp:
            vals = self.tbl_espessuras.item(sel_esp[0], "values")
            sel_esp_default = str(vals[0] if vals else "").strip()

    refs_db = self.data.get("orc_refs", {})
    use_custom = self.encomendas_use_custom and CUSTOM_TK_AVAILABLE
    parent_win = getattr(self, "_enc_editor_win", None) or self.root
    win = ctk.CTkToplevel(parent_win) if use_custom else Toplevel(parent_win)
    win.title("Adicionar peca")
    win.geometry("820x560")
    try:
        win.transient(parent_win)
        win.grab_set()
        win.lift()
        win.focus_force()
    except Exception:
        pass
    win.grid_columnconfigure(1, weight=1)

    Lbl = ctk.CTkLabel if use_custom else ttk.Label
    Ent = ctk.CTkEntry if use_custom else ttk.Entry
    Btn = ctk.CTkButton if use_custom else ttk.Button
    Cmb = ctk.CTkComboBox if use_custom else ttk.Combobox
    Chk = ctk.CTkCheckBox if use_custom else ttk.Checkbutton

    existing = [p.get("ref_interna", "") for p in encomenda_pecas(enc)]
    vars_ = {
        "ref_int": StringVar(value=next_ref_interna_unique(self.data, enc.get("cliente", ""), existing)),
        "ref_ext": StringVar(),
        "descricao": StringVar(),
        "material": StringVar(value=sel_mat_default),
        "espessura": StringVar(value=sel_esp_default),
        "operacao": StringVar(value=OFF_OPERACAO_OBRIGATORIA),
        "qtd": DoubleVar(value=1),
        "preco": DoubleVar(value=0),
        "desenho": StringVar(),
    }

    mat_opts = list(dict.fromkeys(MATERIAIS_PRESET + self.data.get("materiais_hist", []) + list_unique(self.data, "material")))
    esp_opts = [str(v).rstrip("0").rstrip(".") if isinstance(v, float) else str(v) for v in ESPESSURAS_PRESET]
    esp_hist = [str(v) for v in self.data.get("espessuras_hist", [])]
    esp_values = list(dict.fromkeys(esp_opts + esp_hist))

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
            messagebox.showerror("Erro", f"Ficheiro nao encontrado:\n{path}")
            return
        try:
            os.startfile(path)
        except Exception as ex:
            messagebox.showerror("Erro", f"Nao foi possivel abrir o desenho.\n{ex}")

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
        win2.title("Historico de Referencias")
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
            text="Historico de Referencias",
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
            ("desc", "Descricao", 220),
            ("mat", "Material", 100),
            ("esp", "Espessura", 90),
            ("preco", "Preco", 90),
            ("op", "Operacao", 140),
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
                    k,
                    r.get("ref_interna", ""),
                    r.get("descricao", ""),
                    r.get("material", ""),
                    r.get("espessura", ""),
                    r.get("preco_unit", 0),
                    " + ".join(parse_operacoes_lista(r.get("operacao", ""))),
                )
                if query and not any(query in str(v).lower() for v in row):
                    continue
                tag = "even" if row_i % 2 == 0 else "odd"
                tbl.insert("", END, values=row, tags=(tag,))
                row_i += 1

        def choose():
            sel = tbl.selection()
            if not sel:
                messagebox.showerror("Erro", "Selecione uma referencia")
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
                enc.get("cliente", ""),
                [p.get("ref_interna", "") for p in encomenda_pecas(enc)],
            )
        ),
        width=120 if use_custom else None,
    ).pack(side="left", padx=4)

    Lbl(win, text="Ref. Externa").grid(row=1, column=0, sticky="w", padx=8, pady=6)
    Ent(win, textvariable=vars_["ref_ext"], width=(320 if use_custom else 36)).grid(row=1, column=1, padx=8, pady=6, sticky="w")
    btn_ref = ctk.CTkFrame(win, fg_color="transparent") if use_custom else ttk.Frame(win)
    btn_ref.grid(row=1, column=2, columnspan=2, sticky="w", padx=4, pady=6)
    Btn(btn_ref, text="Historico", command=open_ref_history, width=120 if use_custom else None).pack(side="left", padx=4)
    Btn(btn_ref, text="Buscar", command=on_ref_pick, width=120 if use_custom else None).pack(side="left", padx=4)
    Btn(btn_ref, text="Materia-Prima", command=on_materia_pick, width=130 if use_custom else None).pack(side="left", padx=4)

    Lbl(win, text="Desenho Cliente").grid(row=2, column=0, sticky="w", padx=8, pady=6)
    Ent(win, textvariable=vars_["desenho"], width=(500 if use_custom else 48)).grid(row=2, column=1, columnspan=2, padx=8, pady=6, sticky="w")
    draw_btn = ctk.CTkFrame(win, fg_color="transparent") if use_custom else ttk.Frame(win)
    draw_btn.grid(row=2, column=3, sticky="w", padx=4, pady=6)
    Btn(draw_btn, text="Selecionar desenho", command=pick_desenho, width=150 if use_custom else None).pack(side="left", padx=(0, 4))
    Btn(draw_btn, text="Abrir", command=open_desenho, width=90 if use_custom else None).pack(side="left")

    Lbl(win, text="Descricao").grid(row=3, column=0, sticky="w", padx=8, pady=6)
    Ent(win, textvariable=vars_["descricao"], width=(500 if use_custom else 48)).grid(row=3, column=1, columnspan=3, padx=8, pady=6, sticky="w")

    Lbl(win, text="Material").grid(row=4, column=0, sticky="w", padx=8, pady=6)
    if use_custom:
        Cmb(win, variable=vars_["material"], values=mat_opts, width=210).grid(row=4, column=1, padx=8, pady=6, sticky="w")
    else:
        Cmb(win, textvariable=vars_["material"], values=mat_opts, width=18, state="normal").grid(row=4, column=1, padx=8, pady=6, sticky="w")

    Lbl(win, text="Espessura").grid(row=4, column=2, sticky="w", padx=8, pady=6)
    if use_custom:
        Cmb(win, variable=vars_["espessura"], values=esp_values, width=130).grid(row=4, column=3, padx=8, pady=6, sticky="w")
    else:
        Cmb(win, textvariable=vars_["espessura"], values=esp_values, width=10, state="normal").grid(row=4, column=3, padx=8, pady=6, sticky="w")

    Lbl(win, text="Operacao").grid(row=5, column=0, sticky="w", padx=8, pady=6)
    Ent(win, textvariable=vars_["operacao"], width=(420 if use_custom else 44)).grid(row=5, column=1, columnspan=2, padx=8, pady=6, sticky="w")
    op_btns = ctk.CTkFrame(win, fg_color="transparent") if use_custom else ttk.Frame(win)
    op_btns.grid(row=5, column=3, sticky="w", padx=4, pady=6)
    Btn(
        op_btns,
        text="Selecionar",
        command=lambda: (
            lambda val: vars_["operacao"].set(val)
            if val is not None
            else None
        )(self.escolher_operacoes_fluxo(vars_["operacao"].get(), parent=win)),
        width=110 if use_custom else None,
    ).pack(side="left", padx=(0, 4))
    Btn(op_btns, text="Padrao", command=lambda: vars_["operacao"].set(OFF_OPERACAO_OBRIGATORIA), width=90 if use_custom else None).pack(side="left")

    Lbl(win, text="Quantidade").grid(row=6, column=0, sticky="w", padx=8, pady=6)
    Ent(win, textvariable=vars_["qtd"], width=(140 if use_custom else None)).grid(row=6, column=1, padx=8, pady=6, sticky="w")
    Lbl(win, text="Preco Unitario (EUR)").grid(row=7, column=0, sticky="w", padx=8, pady=6)
    Ent(win, textvariable=vars_["preco"], width=(140 if use_custom else None)).grid(row=7, column=1, padx=8, pady=6, sticky="w")

    keep_ref = BooleanVar(value=True)
    Chk(win, text="Guardar referencia na base", variable=keep_ref).grid(row=8, column=0, columnspan=2, sticky="w", padx=8, pady=6)

    def save_piece():
        ref_int_val = vars_["ref_int"].get().strip()
        ref_ext_val = vars_["ref_ext"].get().strip()
        desc_txt = vars_["descricao"].get().strip()
        sel_mat = vars_["material"].get().strip()
        sel_esp = vars_["espessura"].get().strip()
        ops_txt = " + ".join(parse_operacoes_lista(vars_["operacao"].get()))
        qtd_val = parse_float(vars_["qtd"].get(), 0.0)
        preco_val = parse_float(vars_["preco"].get(), 0.0)
        desenho_val = vars_["desenho"].get().strip()

        if not sel_mat or not sel_esp:
            messagebox.showerror("Erro", "Material e espessura sao obrigatorios.")
            return
        if qtd_val <= 0:
            messagebox.showerror("Erro", "Quantidade invalida.")
            return

        if ref_int_val:
            for px in encomenda_pecas(enc):
                if px.get("ref_interna") == ref_int_val:
                    suggested = next_ref_interna_unique(
                        self.data,
                        enc.get("cliente", ""),
                        [p.get("ref_interna", "") for p in encomenda_pecas(enc)],
                    )
                    vars_["ref_int"].set(suggested)
                    messagebox.showerror("Erro", f"Referencia interna ja existe nesta encomenda. Nova sugerida: {suggested}")
                    return

        mat_obj = None
        for m in enc.get("materiais", []):
            if str(m.get("material", "")).strip() == sel_mat:
                mat_obj = m
                break
        if mat_obj is None:
            mat_obj = {"material": sel_mat, "estado": "Preparacao", "espessuras": []}
            enc.setdefault("materiais", []).append(mat_obj)

        esp_obj = None
        for e in mat_obj.get("espessuras", []):
            if str(e.get("espessura", "")).strip() == sel_esp:
                esp_obj = e
                break
        if esp_obj is None:
            esp_obj = {"espessura": sel_esp, "tempo_min": "", "estado": "Preparacao", "pecas": []}
            mat_obj.setdefault("espessuras", []).append(esp_obj)

        peca = {
            "id": f"PEC{len(encomenda_pecas(enc))+1:05d}",
            "ref_interna": ref_int_val,
            "ref_externa": ref_ext_val,
            "material": sel_mat,
            "espessura": sel_esp,
            "descricao": desc_txt,
            "quantidade_pedida": qtd_val,
            "Operacoes": ops_txt,
            "Observacoes": desc_txt,
            "desenho": desenho_val,
            "of": next_of_numero(self.data),
            "opp": next_opp_numero(self.data),
            "estado": "Preparacao",
            "produzido_ok": 0.0,
            "produzido_nok": 0.0,
            "inicio_producao": "",
            "fim_producao": "",
        }
        peca["operacoes_fluxo"] = build_operacoes_fluxo(ops_txt)
        esp_obj.setdefault("pecas", []).append(peca)

        push_unique(self.data.setdefault("materiais_hist", []), peca["material"])
        push_unique(self.data.setdefault("espessuras_hist", []), peca["espessura"])

        if ref_ext_val:
            self.data.setdefault("peca_hist", {})[ref_ext_val] = {
                "ref_interna": ref_int_val,
                "descricao": desc_txt,
                "material": sel_mat,
                "espessura": sel_esp,
                "Operacoes": ops_txt,
                "Observacoes": desc_txt,
                "desenho": desenho_val,
            }

            if keep_ref.get():
                self.data.setdefault("orc_refs", {})[ref_ext_val] = {
                    "ref_interna": ref_int_val,
                    "ref_externa": ref_ext_val,
                    "descricao": desc_txt,
                    "material": sel_mat,
                    "espessura": sel_esp,
                    "preco_unit": preco_val,
                    "operacao": ops_txt,
                    "desenho": desenho_val,
                }
                try:
                    mysql_upsert_orc_referencia(
                        ref_externa=ref_ext_val,
                        ref_interna=ref_int_val,
                        descricao=desc_txt,
                        material=sel_mat,
                        espessura=sel_esp,
                        preco_unit=preco_val,
                        operacao=ops_txt,
                        desenho_path=desenho_val,
                    )
                except Exception:
                    pass

        ensure_peca_operacoes(peca)
        update_refs(self.data, ref_int_val, ref_ext_val)
        # Não recalcular reservas automaticamente ao editar peças.
        # Mantemos as reservas manuais existentes por chapa/lote.
        save_data(self.data)

        self.refresh_encomendas()
        self.selected_encomenda_numero = enc.get("numero")
        self.selected_material = sel_mat
        self.selected_espessura = sel_esp
        self.refresh_materiais(enc)
        self.refresh_espessuras(enc, self.selected_material)
        self.refresh_pecas(enc, sel_esp)
        self.reselect_encomenda_material_espessura()
        win.destroy()

    Btn(win, text="Guardar", command=save_piece, width=150 if use_custom else None).grid(row=9, column=0, columnspan=4, pady=12)
    win.grab_set()
    return win


def preview_encomenda(self):
    _ensure_configured()
    enc = self.get_selected_encomenda()
    if not enc:
        return

    def cliente_display():
        codigo = enc.get("cliente", "")
        nome = ""
        c = find_cliente(self.data, codigo)
        if c:
            nome = c.get("nome", "")
        return f"{codigo} - {nome}".strip(" - ")

    def reserva_dim_lote(r):
        dim = ""
        lote = ""
        if r.get("material_id"):
            for m in self.data.get("materiais", []):
                if m.get("id") == r.get("material_id"):
                    dim = f"{m.get('comprimento','')}x{m.get('largura','')}"
                    lote = m.get("lote_fornecedor", "")
                    break
        if not dim:
            for m in self.data.get("materiais", []):
                if m.get("material") == r.get("material") and m.get("espessura") == r.get("espessura"):
                    dim = f"{m.get('comprimento','')}x{m.get('largura','')}"
                    lote = m.get("lote_fornecedor", "")
                    break
        return (dim or "-"), (lote or "-")

    def dim_for_piece(p):
        for r in enc.get("reservas", []):
            if r.get("material") == p.get("material") and r.get("espessura") == p.get("espessura"):
                dim, _ = reserva_dim_lote(r)
                return dim
        return "-"

    def cativacao_for_piece(p):
        for r in enc.get("reservas", []):
            if r.get("material") == p.get("material") and r.get("espessura") == p.get("espessura"):
                return "Sim"
        return "Nao"

    def opp_for_piece(p):
        return p.get("opp") or "-"

    def render_pdf(path):
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas as pdf_canvas
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        from reportlab.lib.utils import ImageReader
        width, height = A4
        c = pdf_canvas.Canvas(path, pagesize=A4)

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

        def fit_text(text, max_w, font_name, font_size):
            s = pdf_normalize_text(text)
            if pdfmetrics.stringWidth(s, font_name, font_size) <= max_w:
                return s
            ell = "..."
            max_w = max(0, max_w - pdfmetrics.stringWidth(ell, font_name, font_size))
            while s and pdfmetrics.stringWidth(s, font_name, font_size) > max_w:
                s = s[:-1]
            return (s + ell) if s else ell

        def draw_logo(x, y, w, h):
            logo = get_orc_logo_path()
            if not logo or not os.path.exists(logo):
                return
            try:
                img = ImageReader(logo)
                c.drawImage(img, x, yinv(y + h), width=w, height=h, preserveAspectRatio=True, mask="auto")
            except Exception:
                pass

        def fmt_num(val, decimals=2):
            try:
                v = float(val)
            except Exception:
                return str(val) if val is not None else ""
            s = f"{v:.{decimals}f}"
            if "." in s:
                s = s.rstrip("0").rstrip(".")
            return s

        margin = 20
        color_primary = (0.09, 0.27, 0.56)      # azul principal
        color_primary_dark = (0.05, 0.18, 0.40) # azul escuro
        color_border = (0.77, 0.83, 0.91)       # borda suave
        color_row_even = (0.95, 0.97, 1.00)     # zebra clara
        color_row_odd = (0.91, 0.95, 0.99)
        color_text_on_primary = (1, 1, 1)
        headers = ["Codigo", "Material", "Esp.", "Dimensao", "Qtd", "Cativacao", "Ordem de Fabrico"]
        xs = [margin + 4, 150, 230, 285, 365, 415, 485]
        x_ends = xs[1:] + [width - margin]
        col_w = [x_ends[i] - xs[i] for i in range(len(xs))]
        header_h = 18
        header_fs = 9
        row_h = 22
        row_fs = 9
        text_offset = (row_h / 2) + (row_fs / 2 - 1)

        def draw_frame():
            c.setStrokeColorRGB(*color_border)
            c.rect(margin, yinv(height - margin), width - margin * 2, height - margin * 2, stroke=1, fill=0)
            draw_logo(margin + 4, margin + 4, 90, 40)

        def draw_header():
            # Reserve a fixed right column for delivery/client note, preventing overlaps
            # with the title box when note text is long.
            right_col_w = 230
            right_col_left = width - margin - right_col_w
            logo_right = margin + 98
            title_left = logo_right + 10
            title_right = right_col_left - 12
            title_box_w = min(300, max(190, title_right - title_left))
            title_box_x = title_left + max(0.0, ((title_right - title_left) - title_box_w) / 2.0)
            title_box_y = margin + 10
            title_box_h = 24

            c.setStrokeColorRGB(*color_primary_dark)
            c.setLineWidth(1.1)
            c.setFillColorRGB(0.95, 0.97, 1.00)
            c.roundRect(title_box_x, yinv(title_box_y + title_box_h), title_box_w, title_box_h, 6, stroke=1, fill=1)
            c.setFillColorRGB(*color_primary_dark)
            set_font(True, 8.4)
            c.drawCentredString(title_box_x + (title_box_w / 2), yinv(title_box_y - 2), "ENCOMENDA")
            set_font(True, 12)
            title_txt = fit_text(
                f"N {enc.get('numero', '')}",
                title_box_w - 14,
                font_bold,
                12,
            )
            c.drawCentredString(title_box_x + (title_box_w / 2), yinv(title_box_y + 15), title_txt)

            c.setLineWidth(1)
            c.setFillColorRGB(0, 0, 0)
            set_font(False, 9)
            entrega_box_y = margin + 6
            entrega_box_h = 16
            entrega_box_x = right_col_left + 4
            entrega_box_w = right_col_w - 8
            c.setStrokeColorRGB(*color_border)
            c.roundRect(entrega_box_x, yinv(entrega_box_y + entrega_box_h), entrega_box_w, entrega_box_h, 4, stroke=1, fill=0)
            set_font(True, 9)
            entrega_txt = fit_text(
                f"ENTREGA: {enc.get('data_entrega','')}",
                entrega_box_w - 10,
                font_bold,
                9,
            )
            c.drawString(entrega_box_x + 5, yinv(entrega_box_y + 11), entrega_txt)
            nota_cli = str(enc.get("nota_cliente", "") or "").strip()
            if nota_cli:
                nota_box_y = entrega_box_y + entrega_box_h + 4
                nota_box_h = 24
                nota_box_x = right_col_left + 4
                nota_box_w = right_col_w - 8
                c.setStrokeColorRGB(*color_border)
                c.roundRect(nota_box_x, yinv(nota_box_y + nota_box_h), nota_box_w, nota_box_h, 4, stroke=1, fill=0)
                set_font(True, 8.6)
                c.drawString(nota_box_x + 5, yinv(nota_box_y + 10), "NOTA CLIENTE")
                set_font(True, 9)
                nota_txt = fit_text(
                    nota_cli,
                    nota_box_w - 10,
                    font_bold,
                    9,
                )
                c.drawString(nota_box_x + 5, yinv(nota_box_y + 20), nota_txt)
            set_font(True, 9)
            c.setFillColorRGB(*color_primary_dark)
            cli_txt = fit_text(
                f"Cliente: {cliente_display()}",
                title_box_w - 16,
                font_bold,
                9,
            )
            c.drawCentredString(title_box_x + (title_box_w / 2), yinv(margin + 50), cli_txt)
            c.setFillColorRGB(0, 0, 0)
            c.setStrokeColorRGB(*color_border)
            c.line(margin, yinv(84), width - margin, yinv(84))

        def draw_obs_chapas():
            # Observacoes
            set_font(True, 9)
            c.setFillColorRGB(*color_primary_dark)
            c.drawString(margin + 4, yinv(104), "Observacoes:")
            c.setFillColorRGB(0, 0, 0)
            set_font(False, 9)
            c.setStrokeColorRGB(*color_border)
            c.line(margin + 90, yinv(106), width - margin - 6, yinv(106))

            # Chapas cativadas box
            y = 124
            box_h = 90
            c.setStrokeColorRGB(*color_border)
            c.rect(margin, yinv(y + box_h), width - margin * 2, box_h, stroke=1, fill=0)
            set_font(True, 9)
            c.setFillColorRGB(*color_primary_dark)
            c.drawString(margin + 4, yinv(y + 14), "Chapas cativadas:")
            c.setFillColorRGB(0, 0, 0)
            set_font(False, 9)
            ly = y + 28
            if enc.get("reservas"):
                for r in enc.get("reservas", []):
                    dim, lote = reserva_dim_lote(r)
                    linha = f"{r.get('material','')} {r.get('espessura','')} | {dim} | Lote: {lote} | Qtd: {fmt_num(r.get('quantidade',0))}"
                    c.drawString(
                        margin + 10,
                        yinv(ly),
                        fit_text(linha, (width - margin * 2) - 20, font_regular, 9),
                    )
                    ly += 12
                    if ly > y + box_h - 8:
                        break
            else:
                c.drawString(margin + 10, yinv(ly), "Sem chapas cativadas")
            return y + box_h + 18

        def draw_table_header(y):
            header_text_y = y + (header_h / 2) + (header_fs / 2 - 1)
            set_font(True, 9)
            c.setFillColorRGB(*color_primary)
            c.rect(margin, yinv(y + header_h), width - margin * 2, header_h, fill=1, stroke=0)
            c.setFillColorRGB(*color_text_on_primary)
            for h, x, w in zip(headers, xs, col_w):
                c.drawCentredString(x + w / 2, yinv(header_text_y), fit_text(h, w - 6, font_bold, header_fs))
            c.setStrokeColorRGB(*color_border)
            c.line(margin, yinv(y), width - margin, yinv(y))
            c.line(margin, yinv(y + header_h), width - margin, yinv(y + header_h))
            c.setFillColorRGB(0, 0, 0)
            return y + header_h

        def draw_rows(y, rows):
            set_font(False, 9)
            for idx, p in enumerate(rows):
                c.setFillColorRGB(*(color_row_even if idx % 2 == 0 else color_row_odd))
                c.rect(margin, yinv(y + row_h), width - margin * 2, row_h, fill=1, stroke=0)
                c.setFillColorRGB(0, 0, 0)
                row = [
                    p.get("ref_interna", ""),
                    p.get("material", ""),
                    fmt_num(p.get("espessura", "")),
                    dim_for_piece(p),
                    fmt_num(p.get("quantidade_pedida", 0)),
                    cativacao_for_piece(p),
                    opp_for_piece(p),
                ]
                for val, x, w in zip(row, xs, col_w):
                    c.drawCentredString(
                        x + w / 2,
                        yinv(y + text_offset),
                        fit_text(val, w - 6, font_regular, row_fs),
                    )
                c.setStrokeColorRGB(*color_border)
                c.line(margin, yinv(y + row_h), width - margin, yinv(y + row_h))
                y += row_h
            return y

        rows = list(encomenda_pecas(enc))
        table_y_first = 232  # from draw_obs_chapas on page 1
        table_y_next = 110
        max_y = height - margin - 20
        rows_fit_first = max(1, int((max_y - (table_y_first + header_h)) // row_h))
        rows_fit_next = max(1, int((max_y - (table_y_next + header_h)) // row_h))
        if len(rows) <= rows_fit_first:
            total_pages = 1
        else:
            rem = len(rows) - rows_fit_first
            total_pages = 1 + int((rem + rows_fit_next - 1) // rows_fit_next)
        idx = 0
        page = 1
        while idx < len(rows) or idx == 0:
            draw_frame()
            draw_header()
            set_font(False, 8.5)
            c.setFillColorRGB(0, 0, 0)
            c.drawRightString(width - margin - 4, yinv(margin - 4), f"Pag. {page}/{total_pages}")
            if page == 1:
                table_y = draw_obs_chapas()
            else:
                table_y = 110
            y = draw_table_header(table_y)
            start_y = y
            max_y = height - margin - 20
            rows_fit = max(0, int((max_y - start_y) // row_h))
            page_rows = rows[idx: idx + rows_fit] if rows_fit > 0 else []
            y = draw_rows(start_y, page_rows)
            table_bottom = y
            # Vertical grid lines
            x_lines = [margin] + xs[1:] + [width - margin]
            for x in x_lines:
                c.line(x, yinv(table_y), x, yinv(table_bottom))
            idx += len(page_rows)
            if idx >= len(rows):
                break
            c.showPage()
            page += 1
        c.save()


    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(tempfile.gettempdir(), f"lugest_encomenda_{enc.get('numero','')}_{ts}.pdf")
    try:
        render_pdf(path)
    except Exception as exc:
        messagebox.showerror("Erro", f"Falha ao gerar PDF: {exc}")
        return
    try:
        os.startfile(path)
    except Exception:
        messagebox.showerror("Erro", "Nao foi possivel abrir o PDF.")

def _get_selected_cativar_context(self, enc_override=None, mat_override=None, esp_override=None):
    _ensure_configured()
    enc = enc_override if enc_override else self.get_selected_encomenda()
    if not enc:
        return None, "", ""

    sel_mat = str(mat_override or self.selected_material or "").strip()
    sel_esp = str(esp_override or self.selected_espessura or "").strip()

    if not sel_mat and hasattr(self, "tbl_materiais"):
        sel = self.tbl_materiais.selection()
        if sel:
            vals = self.tbl_materiais.item(sel[0], "values")
            sel_mat = str(vals[0] if vals else "").strip()

    if not sel_esp and hasattr(self, "tbl_espessuras"):
        sel = self.tbl_espessuras.selection()
        if sel:
            vals = self.tbl_espessuras.item(sel[0], "values")
            sel_esp = str(vals[0] if vals else "").strip()

    if sel_mat and not sel_esp:
        for m in enc.get("materiais", []):
            if not _match_material(m.get("material"), sel_mat):
                continue
            esps = [str(e.get("espessura", "")).strip() for e in m.get("espessuras", []) if str(e.get("espessura", "")).strip()]
            if len(esps) == 1:
                sel_esp = esps[0]
            break

    return enc, sel_mat, sel_esp

def cativar_stock_selecao(self):
    _ensure_configured()
    opened = cativar_stock(self)
    if not opened:
        return False
    return True


def cativar_stock(self, enc_override=None, mat_override=None, esp_override=None, parent_window=None):
    _ensure_configured()
    enc, sel_mat, sel_esp = _get_selected_cativar_context(
        self,
        enc_override=enc_override,
        mat_override=mat_override,
        esp_override=esp_override,
    )
    if not enc:
        return False
    if not sel_mat or not sel_esp:
        messagebox.showerror("Erro", "Selecione encomenda, material e espessura antes de cativar.")
        return False

    self.selected_material = sel_mat
    self.selected_espessura = sel_esp

    use_custom = CUSTOM_TK_AVAILABLE and getattr(self, "encomendas_use_custom", False)
    Win = ctk.CTkToplevel if use_custom else Toplevel
    Lbl = ctk.CTkLabel if use_custom else ttk.Label
    Ent = ctk.CTkEntry if use_custom else ttk.Entry
    Btn = ctk.CTkButton if use_custom else ttk.Button
    Frm = ctk.CTkFrame if use_custom else ttk.Frame

    parent_ref = parent_window or self.root
    win = Win(parent_ref)
    win.title("Cativar stock")
    try:
        if use_custom:
            win.geometry("900x520")
            win.configure(fg_color="#f7f8fb")
        win.transient(parent_ref)
        win.grab_set()
    except Exception:
        pass

    Lbl(
        win,
        text=f"Defina quantidades por chapa/retalho | {sel_mat} | esp. {sel_esp}",
        font=("Segoe UI", 13, "bold") if use_custom else None,
    ).grid(row=0, column=0, sticky="w", padx=8, pady=6)

    if use_custom:
        list_frame = ctk.CTkScrollableFrame(win, width=860, height=380, fg_color="#ffffff")
        list_frame.grid(row=1, column=0, padx=8, pady=6, sticky="nsew")
        win.grid_columnconfigure(0, weight=1)
        win.grid_rowconfigure(1, weight=1)
        headers = ["Dimensao", "Disponivel", "Local", "Lote", "Reservar"]
        for i, h in enumerate(headers):
            ctk.CTkLabel(list_frame, text=h, font=("Segoe UI", 12, "bold")).grid(row=0, column=i, sticky="w", padx=8, pady=(4, 6))
    else:
        list_frame = Frm(win)
        list_frame.grid(row=1, column=0, padx=6, pady=6, sticky="nsew")
        win.grid_columnconfigure(0, weight=1)
        win.grid_rowconfigure(1, weight=1)
        headers = ["Dimensao", "Disponivel", "Local", "Lote", "Reservar"]
        for i, h in enumerate(headers):
            ttk.Label(list_frame, text=h).grid(row=0, column=i, sticky="w", padx=4, pady=2)

    rows = []
    r = 1
    for m in self.data.get("materiais", []):
        disponivel = parse_float(m.get("quantidade"), 0.0) - parse_float(m.get("reservado"), 0.0)
        if disponivel <= 0:
            continue
        if not _match_material(m.get("material"), sel_mat):
            continue
        if _norm_espessura(m.get("espessura")) != _norm_espessura(sel_esp):
            continue
        dimensao = f"{m.get('comprimento','')}x{m.get('largura','')}"
        Lbl(list_frame, text=dimensao).grid(row=r, column=0, sticky="w", padx=4, pady=2)
        Lbl(list_frame, text=str(disponivel)).grid(row=r, column=1, sticky="w", padx=4, pady=2)
        Lbl(list_frame, text=m.get("LocalizaÃ§Ã£o", m.get("Localizacao", m.get("Localiza?o", "")))).grid(row=r, column=2, sticky="w", padx=4, pady=2)
        Lbl(list_frame, text=m.get("lote_fornecedor", "")).grid(row=r, column=3, sticky="w", padx=4, pady=2)
        v = StringVar()
        Ent(list_frame, textvariable=v, width=90 if use_custom else 8).grid(row=r, column=4, sticky="w", padx=4, pady=2)
        rows.append((m, v))
        r += 1

    if not rows:
        messagebox.showinfo("Info", f"Sem stock disponivel para {sel_mat} esp. {sel_esp}.")
        try:
            win.destroy()
        except Exception:
            pass
        return False

    def on_save():
        any_saved = False
        for m, v in rows:
            val = (v.get() or "").strip()
            if not val:
                continue
            try:
                qtd = float(val)
            except ValueError:
                messagebox.showerror("Erro", f"Quantidade invalida para {m.get('id','')}")
                return
            if qtd <= 0:
                messagebox.showerror("Erro", f"Quantidade invalida para {m.get('id','')}")
                return
            disponivel = parse_float(m.get("quantidade"), 0.0) - parse_float(m.get("reservado"), 0.0)
            if qtd > disponivel:
                messagebox.showerror("Erro", f"Quantidade maior que disponivel para {m.get('id','')}")
                return
            m["reservado"] = parse_float(m.get("reservado"), 0.0) + qtd
            m["atualizado_em"] = now_iso()
            enc.setdefault("reservas", []).append(
                {
                    "material_id": m.get("id"),
                    "material": m.get("material"),
                    "espessura": m.get("espessura"),
                    "quantidade": qtd,
                }
            )
            log_stock(self.data, "CATIVAR", f"{m.get('id','')} qtd={qtd} encomenda={enc.get('numero','')}")
            any_saved = True

        if not any_saved:
            messagebox.showerror("Erro", "Nenhuma quantidade definida")
            return

        enc["cativar"] = True
        try:
            self.e_cativar.set(True)
        except Exception:
            pass
        save_data(self.data)
        self.refresh()
        if self.selected_encomenda_numero == enc.get("numero"):
            self.e_chapa.set(self.get_chapa_reservada(enc.get("numero", ""), sel_mat, sel_esp))
            self.update_cativadas_display(enc)
        win.destroy()

    Btn(win, text="Reservar", command=on_save, width=160 if use_custom else None).grid(row=2, column=0, pady=8)
    return win


def descativar_stock_selecao(self, enc_override=None, mat_override=None, esp_override=None):
    _ensure_configured()
    enc, sel_mat, sel_esp = _get_selected_cativar_context(
        self,
        enc_override=enc_override,
        mat_override=mat_override,
        esp_override=esp_override,
    )
    if not enc:
        return False
    if not sel_mat or not sel_esp:
        messagebox.showerror("Erro", "Selecione material e espessura para descativar.")
        return False

    alvo = []
    manter = []
    for r in enc.get("reservas", []):
        if _match_material(r.get("material"), sel_mat) and _norm_espessura(r.get("espessura")) == _norm_espessura(sel_esp):
            alvo.append(r)
        else:
            manter.append(r)

    if not alvo:
        messagebox.showinfo("Info", f"Sem reservas para {sel_mat} esp. {sel_esp}.")
        return False

    if not messagebox.askyesno("Confirmar", f"Libertar reservas de {sel_mat} esp. {sel_esp}?"):
        return False

    for r in alvo:
        log_stock(self.data, "LIBERTAR", f"{r.get('material_id','')} qtd={r.get('quantidade',0)} encomenda={enc.get('numero','')}")
    aplicar_reserva_em_stock(self.data, alvo, -1)
    enc["reservas"] = manter
    enc["cativar"] = bool(enc.get("reservas"))
    try:
        self.e_cativar.set(enc["cativar"])
    except Exception:
        pass

    save_data(self.data)
    self.refresh()
    if self.selected_encomenda_numero == enc.get("numero"):
        self.e_chapa.set(self.get_chapa_reservada(enc.get("numero", ""), sel_mat, sel_esp))
        self.update_cativadas_display(enc)
    return True
def registar_producao(self, enc=None, preselect_ref=None, default_ok=None):
    _ensure_configured()
    if enc is None:
        enc = self.get_selected_encomenda()
    if not enc:
        return
    if not encomenda_pecas(enc):
        messagebox.showerror("Erro", "Sem peÃ§as")
        return
    win = Toplevel(self.root)
    win.title("Registar produÃ§Ã£o")

    peca_var = StringVar()
    valores = [p["ref_interna"] or p["ref_externa"] or p["id"] for p in encomenda_pecas(enc)]
    ttk.Label(win, text="PeÃ§a").grid(row=0, column=0, sticky="w")
    cb = ttk.Combobox(win, textvariable=peca_var, values=valores)
    cb.grid(row=0, column=1)

    if preselect_ref:
        peca_var.set(preselect_ref)
    else:
        sel = self.tbl_pecas.selection()
        if sel:
            selected_ref = self.tbl_pecas.item(sel[0], "values")[0] or self.tbl_pecas.item(sel[0], "values")[1]
            if selected_ref:
                peca_var.set(selected_ref)

    ok_var = DoubleVar()
    nok_var = DoubleVar()
    motivo_var = StringVar()
    if default_ok is not None:
        ok_var.set(default_ok)
    ttk.Label(win, text="OK").grid(row=1, column=0, sticky="w")
    ttk.Entry(win, textvariable=ok_var).grid(row=1, column=1)
    ttk.Label(win, text="Rejeitadas (NOK)").grid(row=2, column=0, sticky="w")
    ttk.Entry(win, textvariable=nok_var).grid(row=2, column=1)
    ttk.Label(win, text="Motivo NOK").grid(row=3, column=0, sticky="w")
    ttk.Entry(win, textvariable=motivo_var).grid(row=3, column=1)

    def on_save():
        ref = peca_var.get().strip()
        if not ref:
            messagebox.showerror("Erro", "PeÃ§a obrigatoria")
            return
        peca = None
        peca_esp = None
        for m in enc.get("materiais", []):
            for esp in m.get("espessuras", []):
                for p in esp.get("pecas", []):
                    r = p["ref_interna"] or p["ref_externa"] or p["id"]
                    if r == ref:
                        peca = p
                        peca_esp = esp
                        break
                if peca:
                    break
            if peca:
                break
        if not peca:
            messagebox.showerror("Erro", "PeÃ§a nao encontrada")
            return
        ok = float(ok_var.get() or 0)
        nok = float(nok_var.get() or 0)
        if ok < 0 or nok < 0:
            messagebox.showerror("Erro", "Valores invalidos")
            return
        peca["produzido_ok"] += ok
        peca["produzido_nok"] += nok
        if (ok + nok) > 0 and not peca.get("inicio_producao"):
            peca["inicio_producao"] = now_iso()
        atualizar_estado_peca(peca)
        if peca["estado"] == "Concluida":
            peca["fim_producao"] = now_iso()
        if peca_esp:
            if any(px["produzido_ok"] + px["produzido_nok"] > 0 for px in peca_esp.get("pecas", [])):
                peca_esp["estado"] = "Em producao"
            if all(px["estado"] == "Concluida" for px in peca_esp.get("pecas", [])):
                peca_esp["estado"] = "Concluida"
        update_estado_encomenda_por_espessuras(enc)
        if nok > 0:
            self.data["qualidade"].append({
                "encomenda": enc["numero"],
                "peca": peca["id"],
                "ok": ok,
                "nok": nok,
                "motivo": motivo_var.get().strip(),
                "data": now_iso(),
            })
        save_data(self.data)
        self.refresh_encomendas()
        self.refresh_pecas(enc, self.selected_espessura)
        self.refresh_qualidade()
        if hasattr(self, "refresh_operador"):
            self.refresh_operador()
        if hasattr(self, "restore_operador_selection"):
            self.restore_operador_selection()
        win.destroy()

    ttk.Button(win, text="Guardar", command=on_save).grid(row=4, column=0, columnspan=2, pady=10)

def edit_tempo_espessura(self):
    _ensure_configured()
    enc = self.get_selected_encomenda()
    if not enc:
        return
    if not enc.get("materiais"):
        messagebox.showinfo("Info", "Sem materiais")
        return
    use_custom = self.encomendas_use_custom and CUSTOM_TK_AVAILABLE
    Win = ctk.CTkToplevel if use_custom else Toplevel
    Lbl = ctk.CTkLabel if use_custom else ttk.Label
    Ent = ctk.CTkEntry if use_custom else ttk.Entry
    Btn = ctk.CTkButton if use_custom else ttk.Button
    parent_win = getattr(self, "_enc_editor_win", None) or self.root
    win = Win(parent_win)
    win.title("Tempo por espessura")
    if use_custom:
        win.geometry("520x560")
        try:
            win.configure(fg_color="#f7f8fb")
        except Exception:
            pass
    else:
        win.geometry("420x520")
    try:
        win.transient(parent_win)
        win.grab_set()
        win.lift()
        win.focus_force()
    except Exception:
        pass
    vars_map = {}
    if use_custom:
        inner = ctk.CTkScrollableFrame(win, fg_color="#ffffff")
        inner.pack(fill="both", expand=True, padx=10, pady=10)
    else:
        # scrollable content
        canvas = Canvas(win, borderwidth=0, highlightthickness=0)
        scroll = ttk.Scrollbar(win, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

    row = 0
    for m in enc.get("materiais", []):
        for e in m.get("espessuras", []):
            esp = e.get("espessura", "")
            Lbl(inner, text=f"{m.get('material','')} {esp} mm (min)").grid(row=row, column=0, sticky="w", padx=6, pady=4)
            v = StringVar(value=str(e.get("tempo_min", "")))
            Ent(inner, textvariable=v, width=120 if use_custom else 10).grid(row=row, column=1, padx=6, pady=4)
            vars_map[(m.get("material",""), esp)] = v
            row += 1

    def on_save():
        for (mat, esp), v in vars_map.items():
            val = v.get().strip()
            if not val:
                continue
            try:
                for m in enc.get("materiais", []):
                    if m.get("material") != mat:
                        continue
                    for e in m.get("espessuras", []):
                        if e.get("espessura") == esp:
                            e["tempo_min"] = int(val)
            except ValueError:
                messagebox.showerror("Erro", f"Tempo invalido para {mat} {esp}")
                return
        save_data(self.data)
        win.destroy()

    if use_custom:
        Btn(
            inner,
            text="Guardar",
            command=on_save,
            width=220,
            height=42,
            fg_color="#f59e0b",
            hover_color="#d97706",
        ).grid(row=row, column=0, columnspan=2, pady=(14, 10), padx=8, sticky="ew")
    else:
        Btn(inner, text="Guardar", command=on_save, width=20).grid(row=row, column=0, columnspan=2, pady=(10, 8))

def reabrir_encomenda(self):
    _ensure_configured()
    enc = self.get_selected_encomenda()
    if not enc:
        return
    win = Toplevel(self.root)
    win.title("Reabrir encomenda")
    ttk.Label(win, text="Motivo da reabertura").pack(padx=10, pady=(10, 6))
    reason = StringVar(value="Erro")
    for opt in ["Erro", "Avaria", "Pausa", "Outro"]:
        ttk.Radiobutton(win, text=opt, variable=reason, value=opt).pack(anchor="w", padx=10)
    other_var = StringVar()
    ttk.Entry(win, textvariable=other_var, width=30).pack(padx=10, pady=(6, 10))

    def on_ok():
        motivo = reason.get()
        if motivo == "Outro":
            motivo = other_var.get().strip() or "Outro"
        enc["estado"] = "Em producao"
        enc["obs_interrupcao"] = motivo
        save_data(self.data)
        self.refresh()
        win.destroy()

    ttk.Button(win, text="Confirmar", command=on_ok).pack(pady=(0, 10))

def _update_nota_cliente_visual(self):
    _ensure_configured()
    nota = (self.e_nota_cliente.get() if hasattr(self, "e_nota_cliente") else "") or ""
    filled = bool(str(nota).strip())
    if not hasattr(self, "e_nota_status_lbl"):
        return
    if self.encomendas_use_custom and CUSTOM_TK_AVAILABLE:
        txt = "Com nota cliente" if filled else "Sem nota cliente"
        color = "#1f7a4d" if filled else "#b3261e"
        try:
            self.e_nota_status_lbl.configure(text=txt, text_color=color)
        except Exception:
            pass
        try:
            if filled:
                self.e_nota_cliente_entry.configure(border_color="#1f7a4d")
            else:
                self.e_nota_cliente_entry.configure(border_color="#b3261e")
        except Exception:
            pass
    else:
        txt = "Com nota cliente" if filled else "Sem nota cliente"
        color = "#1f7a4d" if filled else "#b3261e"
        try:
            self.e_nota_status_lbl.configure(text=txt, foreground=color)
        except Exception:
            pass

def update_cativadas_display(self, enc):
    _ensure_configured()
    if not hasattr(self, "e_chapas_txt"):
        return
    self.e_chapas_txt.configure(state="normal")
    self.e_chapas_txt.delete("1.0", "end")
    if not enc or not enc.get("reservas"):
        self.e_chapas_txt.insert("1.0", "Sem chapas cativadas")
    else:
        linhas = []
        for r in enc.get("reservas", []):
            dim = ""
            lote = ""
            if r.get("material_id"):
                for m in self.data.get("materiais", []):
                    if m.get("id") == r.get("material_id"):
                        dim = f"{m.get('comprimento','')}x{m.get('largura','')}"
                        lote = m.get("lote_fornecedor", "")
                        break
            dim_txt = f"{dim}" if dim else "-"
            lote_txt = lote or "-"
            linhas.append(
                f"{r.get('material','')} {r.get('espessura','')} | {dim_txt} | Lote: {lote_txt} | Qtd: {r.get('quantidade',0)}"
            )
        self.e_chapas_txt.insert("1.0", "\n".join(linhas))
    self.e_chapas_txt.configure(state="disabled")

def open_ref_search(self, target_var, values, on_pick=None):
    _ensure_configured()
    use_custom = self.encomendas_use_custom and CUSTOM_TK_AVAILABLE
    win = ctk.CTkToplevel(self.root) if use_custom else Toplevel(self.root)
    win.title("Pesquisar referencia externa")
    win.geometry("360x300")
    Lbl = ctk.CTkLabel if use_custom else ttk.Label
    Ent = ctk.CTkEntry if use_custom else ttk.Entry
    Btn = ctk.CTkButton if use_custom else ttk.Button
    Lbl(win, text="Pesquisa").pack(anchor="w", padx=8, pady=(8, 2))
    q = StringVar()
    entry = Ent(win, textvariable=q)
    entry.pack(fill="x", padx=8)
    lst = Listbox(win, height=12)
    lst.pack(fill="both", expand=True, padx=8, pady=8)

    def refresh_list():
        lst.delete(0, END)
        query = q.get().strip().lower()
        for v in values:
            if not query or query in str(v).lower():
                lst.insert(END, v)

    def choose(_=None):
        sel = lst.curselection()
        if not sel:
            return
        target_var.set(lst.get(sel[0]))
        if on_pick:
            on_pick(None)
        win.destroy()

    entry.bind("<KeyRelease>", lambda _: refresh_list())
    lst.bind("<Double-Button-1>", choose)
    Btn(win, text="Selecionar", command=choose, width=140 if use_custom else None).pack(pady=(0, 8))
    refresh_list()
    entry.focus_set()


