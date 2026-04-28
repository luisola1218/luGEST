from module_context import configure_module, ensure_module

_CONFIGURED = False

def configure(main_globals):
    configure_module(globals(), main_globals)

def _ensure_configured():
    ensure_module(globals(), "operador_ordens_actions")

COR_OP_PRODUCAO = "#ffd8a8"
COR_OP_PRODUCAO_BLINK = "#ffbf80"
COR_OP_CONCLUIDA = "#b7e4c7"
COR_OP_INCOMPLETA = "#fff3bf"
COR_OP_INTERROMPIDA = "#fecaca"
COR_OP_PAUSADA = "#fff7cc"
COR_OP_AVARIA = "#fca5a5"
COR_OP_HEADER = "#04145f"
COR_OP_HEADER_ALT = "#0f172a"


def _op_base_bg_from_tag(tag, row_i):
    if tag == "op_em_curso":
        return COR_OP_PRODUCAO
    if tag == "op_pausada":
        return COR_OP_PAUSADA
    if tag == "op_concluida":
        return COR_OP_CONCLUIDA
    if tag == "op_avaria":
        return COR_OP_AVARIA
    if tag == "op_incompleta":
        return COR_OP_INCOMPLETA
    if tag == "op_interrompida":
        return COR_OP_INTERROMPIDA
    return "#f8fbff" if row_i % 2 == 0 else "#edf4ff"


def _op_badge_style(tag):
    if tag == "op_em_curso":
        return "#c2410c", "#fff7ed"
    if tag == "op_pausada":
        return "#b45309", "#fffbeb"
    if tag == "op_concluida":
        return "#166534", "#f0fdf4"
    if tag == "op_avaria":
        return "#b91c1c", "#fef2f2"
    if tag == "op_incompleta":
        return "#854d0e", "#fefce8"
    if tag == "op_interrompida":
        return "#9f1239", "#fff1f2"
    return "#334155", "#f8fafc"


def _op_bind_click(widget, callback):
    try:
        widget.bind("<Button-1>", lambda _e: callback())
    except Exception:
        return


def _op_flow_chip_style(state):
    state = norm_text(state)
    if "curso" in state:
        return "#1d4ed8", "#eff6ff"
    if "concl" in state:
        return "#166534", "#f0fdf4"
    if "paus" in state or "interromp" in state:
        return "#92400e", "#fffbeb"
    if "avari" in state:
        return "#b91c1c", "#fef2f2"
    return "#475569", "#f8fafc"


OP_ROW_FONT = ("Segoe UI", 10)
OP_ROW_FONT_BOLD = ("Segoe UI", 10, "bold")
OP_ROW_FONT_SMALL = ("Segoe UI", 9, "bold")
OP_PIECE_COLUMNS = [
    ("", 28),
    ("Ref. Int.", 108),
    ("Ref. Ext.", 145),
    ("Em curso", 160),
    ("Pend.", 400),
    ("Oper.", 78),
    ("P", 34),
    ("R", 34),
    ("Estado", 92),
]


def _op_make_fixed_header(parent):
    hdr = ctk.CTkFrame(parent, fg_color=COR_OP_HEADER, corner_radius=10)
    hdr.pack(fill="x", padx=8, pady=(0, 6))
    for title, width in OP_PIECE_COLUMNS:
        cell = ctk.CTkFrame(hdr, fg_color="transparent", width=width, height=40)
        cell.pack(side="left", fill="y", padx=(0, 0))
        cell.pack_propagate(False)
        ctk.CTkLabel(
            cell,
            text=title,
            font=("Segoe UI", 11, "bold"),
            text_color="#ffffff",
        ).pack(expand=True, fill="both", padx=4, pady=6)
    return hdr


def _op_make_fixed_cell(parent, width, text, font, anchor="w", text_color="#0f172a"):
    cell = ctk.CTkFrame(parent, fg_color="transparent", width=width, height=32)
    cell.pack(side="left", fill="y")
    cell.pack_propagate(False)
    lbl = ctk.CTkLabel(cell, text=text, font=font, text_color=text_color, anchor=anchor)
    padx = (5, 4) if anchor == "w" else (4, 5)
    lbl.pack(expand=True, fill="both", padx=padx, pady=2)
    return lbl


def _op_short_text(value, max_len=18):
    txt = str(value or "").strip()
    if len(txt) <= max_len:
        return txt
    return txt[: max_len - 1] + "…"


def _op_short_state(value):
    n = norm_text(value)
    if "curso" in n:
        return "Curso"
    if "paus" in n:
        return "Pausa"
    if "concl" in n:
        return "OK"
    if "avari" in n:
        return "Avaria"
    if "incomplet" in n:
        return "Incomp."
    return str(value or "")


def _op_runtime_state_sync(self, force=False, ttl_sec=4.0):
    now_ts = time.time()
    last_ts = float(getattr(self, "_op_runtime_sync_ts", 0) or 0)
    if (not force) and last_ts and ((now_ts - last_ts) < ttl_sec):
        return False
    try:
        sync_fn = getattr(self, "refresh_runtime_impulse_data", None)
        if callable(sync_fn):
            sync_fn(cleanup_orphans=False)
    except Exception:
        pass
    changed = False
    try:
        encomendas = list((getattr(self, "data", {}) or {}).get("encomendas", []) or [])
    except Exception:
        encomendas = []
    selected_num = str(getattr(self, "op_sel_enc_num", "") or "").strip()
    for enc in encomendas:
        if not isinstance(enc, dict):
            continue
        enc_num = str(enc.get("numero", "") or "").strip()
        if selected_num and (not force) and enc_num and (enc_num != selected_num):
            continue
        avaria_index = _op_open_avaria_index(getattr(self, "data", {}), enc_num) if enc_num else {}
        enc_changed = False
        for m in enc.get("materiais", []) or []:
            for e in m.get("espessuras", []) or []:
                for p in e.get("pecas", []) or []:
                    if not isinstance(p, dict):
                        continue
                    live_row = _op_live_avaria_row_for_piece(avaria_index, p)
                    prev_state = str(p.get("estado", "") or "")
                    prev_avaria = bool(p.get("avaria_ativa"))
                    prev_avaria_motivo = str(p.get("avaria_motivo", "") or "").strip()
                    prev_inter_motivo = str(p.get("interrupcao_peca_motivo", "") or "").strip()
                    prev_state_norm = norm_text(prev_state)
                    if live_row:
                        if _op_sync_piece_live_avaria(p, live_row):
                            enc_changed = True
                    else:
                        if prev_avaria:
                            p["avaria_ativa"] = False
                            enc_changed = True
                        if prev_avaria_motivo and "avari" in prev_state_norm:
                            p["avaria_motivo"] = ""
                            enc_changed = True
                        if prev_inter_motivo and prev_avaria_motivo and (prev_inter_motivo == prev_avaria_motivo):
                            p["interrupcao_peca_motivo"] = ""
                            p["interrupcao_peca_ts"] = ""
                            enc_changed = True
                        atualizar_estado_peca(p)
                        if str(p.get("estado", "") or "") != prev_state:
                            enc_changed = True
        prev_enc_state = str(enc.get("estado", "") or "")
        update_estado_encomenda_por_espessuras(enc)
        if str(enc.get("estado", "") or "") != prev_enc_state:
            enc_changed = True
        if enc_changed:
            changed = True
    self._op_runtime_sync_ts = now_ts
    return changed


def _is_laser_op_nome(nome):
    return "laser" in norm_text(nome)


def _peca_tem_laser(p):
    try:
        for op in ensure_peca_operacoes(p):
            if _is_laser_op_nome(op.get("nome", "")):
                return True
    except Exception:
        return False
    return False


def _peca_laser_concluido(p):
    try:
        for op in ensure_peca_operacoes(p):
            if _is_laser_op_nome(op.get("nome", "")):
                return "concl" in norm_text(op.get("estado", ""))
    except Exception:
        return False
    return False


def _esp_laser_concluido(esp_obj):
    pecas = list((esp_obj or {}).get("pecas", []) or [])
    if not pecas:
        return False
    flags = []
    for p in pecas:
        if _peca_tem_laser(p):
            flags.append(_peca_laser_concluido(p))
    if not flags:
        return False
    return all(flags)


def _esp_tem_laser(esp_obj):
    try:
        for p in list((esp_obj or {}).get("pecas", []) or []):
            if _peca_tem_laser(p):
                return True
    except Exception:
        return False
    return False


def _esp_baixa_laser_resolvida(esp_obj):
    if not esp_obj:
        return False
    return bool(esp_obj.get("baixa_laser_feita")) or bool(esp_obj.get("baixa_laser_confirmada_sem_baixa"))


def _dar_baixa_laser_espessura(self, enc, mat, esp, esp_obj):
    # Baixa de stock da matéria-prima no momento em que o LASER termina na espessura.
    if not esp_obj:
        return False
    if _esp_baixa_laser_resolvida(esp_obj):
        return True

    total = 0.0
    for p in esp_obj.get("pecas", []):
        total += (
            float(p.get("produzido_ok", 0) or 0)
            + float(p.get("produzido_nok", 0) or 0)
            + float(p.get("produzido_qualidade", 0) or 0)
        )
    if total <= 0:
        # Sem quantidade ainda, não força baixa.
        return True

    lote_sel = str(esp_obj.get("lote_baixa", "") or "")
    consumo_cativado = 0.0

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

    reservas_restantes = []
    for r in enc.get("reservas", []):
        mat_ok = _norm_mat(r.get("material")) == _norm_mat(mat)
        esp_ok = _norm_esp(r.get("espessura")) == _norm_esp(esp)
        if not (mat_ok and esp_ok):
            reservas_restantes.append(r)
            continue
        qtd_res = parse_float(r.get("quantidade"), 0.0)
        if qtd_res <= 0:
            continue
        stock_sel = None
        mid = str(r.get("material_id", "") or "").strip()
        if mid:
            for stock in self.data.get("materiais", []):
                if str(stock.get("id", "") or "") == mid:
                    stock_sel = stock
                    break
        if stock_sel is None:
            for stock in self.data.get("materiais", []):
                if _norm_mat(stock.get("material")) == _norm_mat(mat) and _norm_esp(stock.get("espessura")) == _norm_esp(esp):
                    stock_sel = stock
                    break
        if stock_sel is None:
            reservas_restantes.append(r)
            continue
        stock_sel["quantidade"] = max(0, parse_float(stock_sel.get("quantidade"), 0.0) - qtd_res)
        stock_sel["reservado"] = max(0, parse_float(stock_sel.get("reservado"), 0.0) - qtd_res)
        stock_sel["atualizado_em"] = now_iso()
        if not lote_sel:
            lote_sel = str(stock_sel.get("lote_fornecedor", "") or "")
        consumo_cativado += qtd_res
        log_stock(self.data, "BAIXA CATIVADA", f"{stock_sel.get('id','')} qtd={qtd_res} encomenda={enc.get('numero','')}")

    enc["reservas"] = reservas_restantes
    enc["cativar"] = bool(reservas_restantes)

    if consumo_cativado > 0:
        msg = (
            f"Foi dada baixa automatica de {fmt_num(consumo_cativado)} quantidades cativadas "
            f"para {mat} esp. {esp}.\n\nDeseja dar baixa adicional?"
        )
        if messagebox.askyesno("Baixa automatica (Laser)", msg):
            baixa = self.confirmar_baixa_stock(mat, esp, total)
            if baixa is not None:
                qtd_baixa, material_id, lote_manual = baixa
                for stock in self.data.get("materiais", []):
                    if material_id and stock.get("id") == material_id:
                        stock["quantidade"] = max(0, parse_float(stock.get("quantidade"), 0.0) - qtd_baixa)
                        stock["atualizado_em"] = now_iso()
                        break
                log_stock(self.data, "BAIXA", f"{material_id or ''} qtd={qtd_baixa} encomenda={enc.get('numero','')}")
                if lote_manual:
                    lote_sel = lote_manual
    else:
        baixa = self.confirmar_baixa_stock(mat, esp, total)
        if baixa is None:
            if messagebox.askyesno(
                "Sem baixa de stock",
                (
                    f"Não foi possível registar baixa para {mat} {esp}.\n\n"
                    "Pretende concluir o Laser mesmo assim (sem baixa)?"
                ),
            ):
                try:
                    log_stock(
                        self.data,
                        "SEM_BAIXA",
                        f"encomenda={enc.get('numero','')} mat={mat} esp={esp} motivo=laser_sem_stock",
                    )
                except Exception:
                    pass
            else:
                return False
        else:
            qtd_baixa, material_id, lote_manual = baixa
            for stock in self.data.get("materiais", []):
                if material_id and stock.get("id") == material_id:
                    stock["quantidade"] = max(0, parse_float(stock.get("quantidade"), 0.0) - qtd_baixa)
                    stock["atualizado_em"] = now_iso()
                    break
            log_stock(self.data, "BAIXA", f"{material_id or ''} qtd={qtd_baixa} encomenda={enc.get('numero','')}")
            if lote_manual:
                lote_sel = lote_manual

    if lote_sel:
        esp_obj["lote_baixa"] = lote_sel
        for p in esp_obj.get("pecas", []):
            p["lote_baixa"] = lote_sel
    esp_obj["baixa_laser_feita"] = True
    esp_obj["baixa_laser_confirmada_sem_baixa"] = False
    esp_obj["baixa_laser_em"] = now_iso()
    return True


def _prompt_sobras_laser(self, mat, esp, lote_sel=""):
    if not messagebox.askyesno("Sobras", "Existem sobras a adicionar ao stock?"):
        return
    use_custom = CUSTOM_TK_AVAILABLE and getattr(self, "op_use_custom", False)
    Win = ctk.CTkToplevel if use_custom else Toplevel
    Lbl = ctk.CTkLabel if use_custom else ttk.Label
    Ent = ctk.CTkEntry if use_custom else ttk.Entry
    Btn = ctk.CTkButton if use_custom else ttk.Button
    win = Win(self.root)
    win.title("Retalho")
    win.transient(self.root)
    win.lift()
    win.grab_set()
    if use_custom:
        try:
            win.geometry("460x280")
            win.resizable(False, False)
        except Exception:
            pass
    comp_var = StringVar()
    larg_var = StringVar()
    qtd_var = StringVar()
    Lbl(win, text=f"Material: {mat} | Espessura: {esp}", font=("Segoe UI", 12, "bold") if use_custom else None).grid(row=0, column=0, columnspan=2, sticky="w", padx=8, pady=(10, 8))
    Lbl(win, text="Comprimento (mm)").grid(row=1, column=0, sticky="w", padx=8, pady=5)
    Ent(win, textvariable=comp_var, width=170 if use_custom else None).grid(row=1, column=1, padx=8, pady=5)
    Lbl(win, text="Largura (mm)").grid(row=2, column=0, sticky="w", padx=8, pady=5)
    Ent(win, textvariable=larg_var, width=170 if use_custom else None).grid(row=2, column=1, padx=8, pady=5)
    Lbl(win, text="Quantidade").grid(row=3, column=0, sticky="w", padx=8, pady=5)
    Ent(win, textvariable=qtd_var, width=170 if use_custom else None).grid(row=3, column=1, padx=8, pady=5)

    def on_save():
        try:
            sc = float(comp_var.get())
            sl = float(larg_var.get())
            sq = float(qtd_var.get())
        except Exception:
            messagebox.showerror("Erro", "Valores inválidos de sobra")
            return
        self.data["materiais"].append({
            "id": f"MAT{len(self.data['materiais'])+1:05d}",
            "material": mat,
            "espessura": esp,
            "comprimento": sc,
            "largura": sl,
            "quantidade": sq,
            "reservado": 0.0,
            "Localizacao": "RETALHO",
            "lote_fornecedor": lote_sel or "",
            "is_sobra": True,
            "atualizado_em": now_iso(),
        })
        win.destroy()

    Btn(win, text="Guardar", command=on_save, width=130 if use_custom else None).grid(row=4, column=0, padx=8, pady=10)
    Btn(win, text="Cancelar", command=win.destroy, width=130 if use_custom else None).grid(row=4, column=1, padx=8, pady=10)
    self.root.wait_window(win)


def _trigger_laser_completion_actions(self, enc, mat, esp):
    esp_obj = self.get_operador_esp_obj(enc, mat, esp)
    if not esp_obj:
        return True
    if not _esp_laser_concluido(esp_obj):
        return True
    if not esp_obj.get("laser_concluido"):
        esp_obj["laser_concluido"] = True
        esp_obj["laser_concluido_em"] = now_iso()
    if _esp_baixa_laser_resolvida(esp_obj):
        return True
    ok = bool(_dar_baixa_laser_espessura(self, enc, mat, esp, esp_obj))
    if ok:
        _prompt_sobras_laser(self, mat, esp, str(esp_obj.get("lote_baixa", "") or ""))
    return ok


def _flush_piece_elapsed_minutes(piece_obj, ts_end=None):
    """Accumulate elapsed production time segment and stop active timing."""
    try:
        p = piece_obj if isinstance(piece_obj, dict) else {}
        ini = p.get("inicio_producao")
        if not ini:
            p["inicio_producao"] = ""
            return 0.0
        ts_ref = str(ts_end or now_iso())
        seg = iso_diff_minutes(ini, ts_ref)
        seg_val = parse_float(seg, 0.0)
        if seg_val > 0:
            base = parse_float(p.get("tempo_producao_min", 0.0), 0.0)
            p["tempo_producao_min"] = round(base + seg_val, 4)
        p["inicio_producao"] = ""
        return seg_val
    except Exception:
        try:
            piece_obj["inicio_producao"] = ""
        except Exception:
            pass
        return 0.0


def _mysql_ops_schema_ensure(cur):
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS peca_operacoes_execucao (
            id INT AUTO_INCREMENT PRIMARY KEY,
            encomenda_numero VARCHAR(30) NOT NULL,
            peca_id VARCHAR(30) NOT NULL,
            operacao VARCHAR(80) NOT NULL,
            estado VARCHAR(20) NOT NULL DEFAULT 'Livre',
            operador_atual VARCHAR(120) NULL,
            inicio_ts DATETIME NULL,
            fim_ts DATETIME NULL,
            ok_qty DECIMAL(10,2) NULL,
            nok_qty DECIMAL(10,2) NULL,
            qual_qty DECIMAL(10,2) NULL,
            updated_at DATETIME NULL,
            UNIQUE KEY uq_peca_operacao (peca_id, operacao),
            INDEX idx_poe_enc (encomenda_numero),
            INDEX idx_poe_estado (estado),
            INDEX idx_poe_operador (operador_atual)
        )
        """
    )
    # Escopo correto do lock: por encomenda + peça + operação.
    # Evita que uma peça "CL0003-0001" de outra encomenda herde estado em produção.
    try:
        cur.execute(
            """
            SELECT COUNT(*)
            FROM information_schema.STATISTICS
            WHERE TABLE_SCHEMA = DATABASE()
              AND TABLE_NAME = 'peca_operacoes_execucao'
              AND INDEX_NAME = 'uq_poe_enc_peca_operacao'
            """
        )
        row = cur.fetchone()
        has_new_uq = int((row[0] if isinstance(row, (list, tuple)) else row.get("COUNT(*)", 0)) or 0) > 0
        if not has_new_uq:
            cur.execute(
                """
                ALTER TABLE peca_operacoes_execucao
                ADD UNIQUE KEY uq_poe_enc_peca_operacao (encomenda_numero, peca_id, operacao)
                """
            )
    except Exception:
        pass
    # Remove unique antigo (peca_id,operacao) para permitir a mesma referência em encomendas diferentes.
    try:
        cur.execute(
            """
            SELECT COUNT(*)
            FROM information_schema.STATISTICS
            WHERE TABLE_SCHEMA = DATABASE()
              AND TABLE_NAME = 'peca_operacoes_execucao'
              AND INDEX_NAME = 'uq_peca_operacao'
            """
        )
        row = cur.fetchone()
        has_old_uq = int((row[0] if isinstance(row, (list, tuple)) else row.get("COUNT(*)", 0)) or 0) > 0
        if has_old_uq:
            cur.execute("ALTER TABLE peca_operacoes_execucao DROP INDEX uq_peca_operacao")
    except Exception:
        pass
    # Limpeza de locks antigos sem dono: evita bloqueios "fantasma".
    try:
        cur.execute(
            """
            UPDATE peca_operacoes_execucao
            SET estado='Livre', inicio_ts=NULL, fim_ts=NULL, updated_at=NOW()
            WHERE estado='Em producao' AND (operador_atual IS NULL OR operador_atual='')
            """
        )
    except Exception:
        pass


def _same_operator_name(a, b):
    try:
        na = norm_text(a)
        nb = norm_text(b)
        return bool(na) and bool(nb) and na == nb
    except Exception:
        sa = str(a or "").strip().lower()
        sb = str(b or "").strip().lower()
        return bool(sa) and bool(sb) and sa == sb


def _normalize_operator_set(valid_operators):
    out = set()
    for o in list(valid_operators or []):
        try:
            n = norm_text(o)
        except Exception:
            n = str(o or "").strip().lower()
        if n:
            out.add(n)
    return out


def _mysql_ops_retryable_error(ex):
    try:
        args = list(getattr(ex, "args", []) or [])
    except Exception:
        args = []
    if args:
        try:
            code = int(args[0])
        except Exception:
            code = None
        if code in (1205, 1213):
            return True
    text = str(ex or "").strip().lower()
    return ("deadlock" in text) or ("lock wait timeout" in text)


def _mysql_ops_acquire(enc_num, peca_id, operacoes, operador, valid_operators=None):
    if not (USE_MYSQL_STORAGE and MYSQL_AVAILABLE):
        return {"acquired": list(operacoes or []), "blocked": []}
    valid_ops_norm = _normalize_operator_set(valid_operators)
    last_error = None
    for attempt in range(3):
        conn = None
        acquired = []
        blocked = []
        try:
            conn = _mysql_connect()
            with conn.cursor() as cur:
                _mysql_ops_schema_ensure(cur)
                for op in list(operacoes or []):
                    opn = normalize_operacao_nome(op)
                    if not opn:
                        continue
                    cur.execute(
                        """
                        INSERT INTO peca_operacoes_execucao (
                            encomenda_numero, peca_id, operacao, estado, updated_at
                        ) VALUES (%s, %s, %s, 'Livre', NOW())
                        ON DUPLICATE KEY UPDATE
                            encomenda_numero = VALUES(encomenda_numero),
                            updated_at = NOW()
                        """,
                        (enc_num, peca_id, opn),
                    )
                    cur.execute(
                        """
                        UPDATE peca_operacoes_execucao
                        SET estado='Em producao',
                            operador_atual=%s,
                            inicio_ts=COALESCE(inicio_ts, NOW()),
                            fim_ts=NULL,
                            updated_at=NOW()
                        WHERE encomenda_numero=%s
                          AND peca_id=%s
                          AND operacao=%s
                          AND (estado='Livre' OR (estado='Em producao' AND operador_atual=%s))
                        """,
                        (operador, enc_num, peca_id, opn, operador),
                    )
                    if int(getattr(cur, "rowcount", 0) or 0) > 0:
                        acquired.append(opn)
                    else:
                        cur.execute(
                            """
                            SELECT operador_atual, estado
                            FROM peca_operacoes_execucao
                            WHERE encomenda_numero=%s
                              AND peca_id=%s
                              AND operacao=%s
                            LIMIT 1
                            """,
                            (enc_num, peca_id, opn),
                        )
                        row = cur.fetchone() or {}
                        owner = str((row.get("operador_atual") if isinstance(row, dict) else "") or "").strip()
                        st = str((row.get("estado") if isinstance(row, dict) else "") or "").strip()
                        owner_norm = norm_text(owner) if owner else ""
                        owner_missing = bool(owner_norm) and bool(valid_ops_norm) and (owner_norm not in valid_ops_norm)
                        if (not owner) or owner_missing or _same_operator_name(owner, operador):
                            cur.execute(
                                """
                                UPDATE peca_operacoes_execucao
                                SET estado='Em producao',
                                    operador_atual=%s,
                                    inicio_ts=COALESCE(inicio_ts, NOW()),
                                    fim_ts=NULL,
                                    updated_at=NOW()
                                WHERE encomenda_numero=%s
                                  AND peca_id=%s
                                  AND operacao=%s
                                """,
                                (operador, enc_num, peca_id, opn),
                            )
                            if int(getattr(cur, "rowcount", 0) or 0) > 0:
                                acquired.append(opn)
                                continue
                        blocked.append({"operacao": opn, "operador": owner, "estado": st})
            conn.commit()
            return {"acquired": acquired, "blocked": blocked}
        except Exception as ex:
            last_error = ex
            if conn:
                try:
                    conn.rollback()
                except Exception:
                    pass
            if (attempt + 1) >= 3 or not _mysql_ops_retryable_error(ex):
                raise RuntimeError(f"Falha MySQL ao iniciar operacao: {ex}") from ex
            time.sleep(0.25 * (attempt + 1))
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass
    if last_error is not None:
        raise RuntimeError(f"Falha MySQL ao iniciar operacao: {last_error}") from last_error
    return {"acquired": [], "blocked": []}


def _mysql_ops_finish(enc_num, peca_id, operacoes, operador, ok, nok, qual, valid_operators=None, complete=True):
    if not (USE_MYSQL_STORAGE and MYSQL_AVAILABLE):
        return {"done": list(operacoes or []), "blocked": []}
    valid_ops_norm = _normalize_operator_set(valid_operators)
    target_state = "Concluida" if complete else "Livre"
    target_owner = operador if complete else None
    last_error = None
    for attempt in range(3):
        conn = None
        done = []
        blocked = []
        try:
            conn = _mysql_connect()
            with conn.cursor() as cur:
                _mysql_ops_schema_ensure(cur)
                for op in list(operacoes or []):
                    opn = normalize_operacao_nome(op)
                    if not opn:
                        continue
                    cur.execute(
                        """
                        INSERT INTO peca_operacoes_execucao (
                            encomenda_numero, peca_id, operacao, estado, updated_at
                        ) VALUES (%s, %s, %s, 'Livre', NOW())
                        ON DUPLICATE KEY UPDATE
                            updated_at = NOW()
                        """,
                        (enc_num, peca_id, opn),
                    )
                    cur.execute(
                        """
                        UPDATE peca_operacoes_execucao
                        SET estado=%s,
                            operador_atual=%s,
                            fim_ts=NOW(),
                            ok_qty=%s,
                            nok_qty=%s,
                            qual_qty=%s,
                            updated_at=NOW()
                        WHERE encomenda_numero=%s
                          AND peca_id=%s
                          AND operacao=%s
                          AND estado='Em producao'
                          AND (operador_atual=%s OR operador_atual IS NULL OR operador_atual='')
                        """,
                        (target_state, target_owner, ok, nok, qual, enc_num, peca_id, opn, operador),
                    )
                    if int(getattr(cur, "rowcount", 0) or 0) > 0:
                        done.append(opn)
                    else:
                        cur.execute(
                            "SELECT operador_atual, estado FROM peca_operacoes_execucao WHERE encomenda_numero=%s AND peca_id=%s AND operacao=%s LIMIT 1",
                            (enc_num, peca_id, opn),
                        )
                        row = cur.fetchone() or {}
                        owner = str((row.get("operador_atual") if isinstance(row, dict) else "") or "").strip()
                        st = str((row.get("estado") if isinstance(row, dict) else "") or "").strip()
                        owner_norm = norm_text(owner) if owner else ""
                        owner_missing = bool(owner_norm) and bool(valid_ops_norm) and (owner_norm not in valid_ops_norm)
                        st_norm = norm_text(st)
                        if ("produ" in st_norm) and ((not owner) or owner_missing or _same_operator_name(owner, operador)):
                            cur.execute(
                                """
                                UPDATE peca_operacoes_execucao
                                SET estado=%s,
                                    operador_atual=%s,
                                    fim_ts=NOW(),
                                    ok_qty=%s,
                                    nok_qty=%s,
                                    qual_qty=%s,
                                    updated_at=NOW()
                                WHERE encomenda_numero=%s AND peca_id=%s AND operacao=%s
                                """,
                                (target_state, target_owner, ok, nok, qual, enc_num, peca_id, opn),
                            )
                            if int(getattr(cur, "rowcount", 0) or 0) > 0:
                                done.append(opn)
                                continue
                        blocked.append({"operacao": opn, "operador": owner, "estado": st})
            conn.commit()
            return {"done": done, "blocked": blocked}
        except Exception as ex:
            last_error = ex
            if conn:
                try:
                    conn.rollback()
                except Exception:
                    pass
            if (attempt + 1) >= 3 or not _mysql_ops_retryable_error(ex):
                raise RuntimeError(f"Falha MySQL ao concluir operacao: {ex}") from ex
            time.sleep(0.25 * (attempt + 1))
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass
    if last_error is not None:
        raise RuntimeError(f"Falha MySQL ao concluir operacao: {last_error}") from last_error
    return {"done": [], "blocked": []}


def _mysql_ops_reset_piece(enc_num, peca_id):
    if not (USE_MYSQL_STORAGE and MYSQL_AVAILABLE):
        return
    conn = None
    try:
        conn = _mysql_connect()
        with conn.cursor() as cur:
            _mysql_ops_schema_ensure(cur)
            cur.execute(
                """
                UPDATE peca_operacoes_execucao
                SET estado='Livre',
                    operador_atual=NULL,
                    inicio_ts=NULL,
                    fim_ts=NULL,
                    ok_qty=NULL,
                    nok_qty=NULL,
                    qual_qty=NULL,
                    updated_at=NOW()
                WHERE encomenda_numero=%s
                  AND peca_id=%s
                """,
                (enc_num, peca_id),
            )
        conn.commit()
    except Exception:
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
    finally:
        try:
            if conn:
                conn.close()
        except Exception:
            pass


def _mysql_ops_delete_order(enc_num):
    if not (USE_MYSQL_STORAGE and MYSQL_AVAILABLE):
        return
    enc_txt = str(enc_num or "").strip()
    if not enc_txt:
        return
    conn = None
    try:
        conn = _mysql_connect()
        with conn.cursor() as cur:
            _mysql_ops_schema_ensure(cur)
            cur.execute(
                """
                DELETE FROM peca_operacoes_execucao
                WHERE encomenda_numero=%s
                """,
                (enc_txt,),
            )
        conn.commit()
    except Exception:
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
    finally:
        try:
            if conn:
                conn.close()
        except Exception:
            pass


def _mysql_ops_status_for_piece(enc_num, peca_id):
    if not (USE_MYSQL_STORAGE and MYSQL_AVAILABLE):
        return []
    conn = None
    out = []
    try:
        conn = _mysql_connect()
        with conn.cursor() as cur:
            _mysql_ops_schema_ensure(cur)
            cur.execute(
                """
                SELECT operacao, estado, operador_atual
                FROM peca_operacoes_execucao
                WHERE encomenda_numero=%s
                  AND peca_id=%s
                ORDER BY operacao
                """,
                (enc_num, peca_id),
            )
            rows = cur.fetchall() or []
            for r in rows:
                if isinstance(r, dict):
                    out.append(
                        {
                            "operacao": normalize_operacao_nome(r.get("operacao", "")),
                            "estado": str(r.get("estado", "") or ""),
                            "operador": str(r.get("operador_atual", "") or ""),
                        }
                    )
    except Exception:
        return []
    finally:
        try:
            if conn:
                conn.close()
        except Exception:
            pass
    return out


def _mysql_ops_status_for_order(enc_num, cache_owner=None, force=False, ttl_sec=1.2):
    if not (USE_MYSQL_STORAGE and MYSQL_AVAILABLE):
        return {}
    enc_key = str(enc_num or "").strip()
    if not enc_key:
        return {}
    cache = None
    now_ts = time.time()
    if cache_owner is not None:
        try:
            cache = getattr(cache_owner, "_op_mysql_ops_status_cache", None)
            if not isinstance(cache, dict):
                cache = {}
                cache_owner._op_mysql_ops_status_cache = cache
            cached = cache.get(enc_key)
            if (
                (not force)
                and isinstance(cached, dict)
                and ((now_ts - float(cached.get("ts", 0) or 0)) < ttl_sec)
            ):
                return dict(cached.get("data", {}) or {})
        except Exception:
            cache = None
    conn = None
    out = {}
    try:
        conn = _mysql_connect()
        with conn.cursor() as cur:
            _mysql_ops_schema_ensure(cur)
            cur.execute(
                """
                SELECT peca_id, operacao, estado, operador_atual
                FROM peca_operacoes_execucao
                WHERE encomenda_numero=%s
                ORDER BY peca_id, operacao
                """,
                (enc_key,),
            )
            rows = cur.fetchall() or []
            for r in rows:
                if not isinstance(r, dict):
                    continue
                pid = str(r.get("peca_id", "") or "").strip()
                if not pid:
                    continue
                out.setdefault(pid, []).append(
                    {
                        "operacao": normalize_operacao_nome(r.get("operacao", "")),
                        "estado": str(r.get("estado", "") or ""),
                        "operador": str(r.get("operador_atual", "") or ""),
                    }
                )
    except Exception:
        return {}
    finally:
        try:
            if conn:
                conn.close()
        except Exception:
            pass
    if cache is not None:
        try:
            cache[enc_key] = {"ts": now_ts, "data": dict(out)}
        except Exception:
            pass
    return out


def _mysql_ops_release_piece(enc_num, peca_id):
    if not (USE_MYSQL_STORAGE and MYSQL_AVAILABLE):
        return
    conn = None
    try:
        conn = _mysql_connect()
        with conn.cursor() as cur:
            _mysql_ops_schema_ensure(cur)
            cur.execute(
                """
                UPDATE peca_operacoes_execucao
                SET estado='Livre',
                    operador_atual=NULL,
                    inicio_ts=NULL,
                    fim_ts=NULL,
                    updated_at=NOW()
                WHERE encomenda_numero=%s
                  AND peca_id=%s
                  AND estado='Em producao'
                """,
                (enc_num, peca_id),
            )
        conn.commit()
    except Exception:
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
    finally:
        try:
            if conn:
                conn.close()
        except Exception:
            pass


def _op_is_open_avaria_row(row):
    if not isinstance(row, dict):
        return False
    state_norm = str(row.get("estado", "") or "").strip().lower()
    if bool(str(row.get("fechada_at", "") or "").strip()) or ("fech" in state_norm):
        return False
    try:
        dur = float(row.get("duracao_min", 0) or 0)
    except Exception:
        dur = 0.0
    return dur <= 0.0


def _op_open_avaria_index(data, enc_num):
    out = {}
    if not isinstance(data, dict):
        return out
    rows = sorted(
        list(data.get("op_paragens", []) or []),
        key=lambda row: str((row or {}).get("created_at", "") or ""),
        reverse=True,
    )
    enc_key = str(enc_num or "").strip()
    for row in rows:
        if not _op_is_open_avaria_row(row):
            continue
        if str(row.get("encomenda_numero", "") or "").strip() != enc_key:
            continue
        pid = str(row.get("peca_id", "") or "").strip()
        ref = str(row.get("ref_interna", "") or "").strip()
        if pid and pid not in out:
            out[pid] = row
        if ref and ref not in out:
            out[ref] = row
    return out


def _op_live_avaria_row_for_piece(avaria_index, peca):
    if not isinstance(avaria_index, dict) or not isinstance(peca, dict):
        return {}
    pid = str(peca.get("id", "") or "").strip()
    ref = str(peca.get("ref_interna", "") or "").strip()
    row = avaria_index.get(pid) or avaria_index.get(ref)
    return row if isinstance(row, dict) else {}


def _op_piece_lookup_key(peca):
    if not isinstance(peca, dict):
        return ""
    return str(peca.get("id", "") or peca.get("ref_interna", "") or "").strip()


def _op_float(value, default=0.0):
    parser = globals().get("parse_float")
    if callable(parser):
        try:
            return float(parser(value, default))
        except Exception:
            pass
    try:
        return float(str(value).strip().replace(",", "."))
    except Exception:
        return float(default)


def _op_make_avaria_group_id(enc_num="", motivo="", operador="", ts_now=None):
    enc_key = str(enc_num or "").strip() or "ENC"
    raw_ts = str(ts_now or now_iso()).strip()
    ts_key = "".join(ch for ch in raw_ts if ch.isdigit())[-14:] or str(int(time.time() * 1000))
    try:
        token = str(uuid.uuid4().hex[:8])
    except Exception:
        token = str(int(time.time_ns() % 100000000)).zfill(8)
    return f"AVG|{enc_key}|{ts_key}|{token}"


def _op_closed_avaria_minutes_index(data, enc_num):
    index = {}
    if not isinstance(data, dict):
        return index
    enc_key = str(enc_num or "").strip()
    for row in list(data.get("op_paragens", []) or []):
        if not isinstance(row, dict):
            continue
        if enc_key and str(row.get("encomenda_numero", "") or "").strip() != enc_key:
            continue
        dur = max(0.0, _op_float(row.get("duracao_min", 0), 0.0))
        if dur <= 0:
            continue
        piece_id = str(row.get("peca_id", "") or "").strip()
        ref_int = str(row.get("ref_interna", "") or "").strip()
        if piece_id:
            index[("id", piece_id)] = index.get(("id", piece_id), 0.0) + dur
        if ref_int:
            index[("ref", ref_int)] = index.get(("ref", ref_int), 0.0) + dur
    return index


def _op_piece_closed_avaria_minutes(data, enc_num, peca, closed_index=None):
    if not isinstance(peca, dict):
        return 0.0
    index = closed_index if isinstance(closed_index, dict) else _op_closed_avaria_minutes_index(data, enc_num)
    piece_id = str(peca.get("id", "") or "").strip()
    ref_int = str(peca.get("ref_interna", "") or "").strip()
    if piece_id and ("id", piece_id) in index:
        return max(0.0, _op_float(index.get(("id", piece_id), 0), 0.0))
    if ref_int and ("ref", ref_int) in index:
        return max(0.0, _op_float(index.get(("ref", ref_int), 0), 0.0))
    return 0.0


def _op_piece_current_avaria_minutes(data, enc_num, peca, live_row=None, ts_ref=None):
    if not isinstance(peca, dict):
        return 0.0
    row = live_row if isinstance(live_row, dict) else {}
    if not row:
        row = _op_live_avaria_row_for_piece(_op_open_avaria_index(data, enc_num), peca)
    active = bool(row) or bool(peca.get("avaria_ativa"))
    if not active:
        return 0.0
    start_txt = str(
        (row or {}).get("created_at", "")
        or peca.get("avaria_inicio_ts", "")
        or peca.get("interrupcao_peca_ts", "")
        or ""
    ).strip()
    if not start_txt:
        return 0.0
    mins = iso_diff_minutes(start_txt, str(ts_ref or now_iso()))
    return max(0.0, _op_float(mins, 0.0))


def _op_piece_total_avaria_minutes(data, enc_num, peca, *, include_open=True, live_row=None, ts_ref=None, closed_index=None):
    total = _op_piece_closed_avaria_minutes(data, enc_num, peca, closed_index=closed_index)
    if include_open:
        total += _op_piece_current_avaria_minutes(data, enc_num, peca, live_row=live_row, ts_ref=ts_ref)
    return round(max(0.0, total), 4)


def _op_total_avaria_minutes_for_pieces(data, enc, pecas, *, include_open=True, ts_ref=None):
    enc_num = str((enc or {}).get("numero", "") or "").strip()
    closed_index = _op_closed_avaria_minutes_index(data, enc_num)
    avaria_index = _op_open_avaria_index(data, enc_num) if include_open else {}
    total = 0.0
    seen = set()
    for p in list(pecas or []):
        if not isinstance(p, dict):
            continue
        key = _op_piece_lookup_key(p)
        if not key or key in seen:
            continue
        seen.add(key)
        live_row = _op_live_avaria_row_for_piece(avaria_index, p) if include_open else {}
        total += _op_piece_total_avaria_minutes(
            data,
            enc_num,
            p,
            include_open=include_open,
            live_row=live_row,
            ts_ref=ts_ref,
            closed_index=closed_index,
        )
    return round(max(0.0, total), 4)


def _op_avaria_group_key(row):
    if not isinstance(row, dict):
        return ""
    explicit = str(
        row.get("grupo_id", "")
        or row.get("group_id", "")
        or row.get("batch_id", "")
        or row.get("paragem_grupo", "")
        or ""
    ).strip()
    if explicit:
        return explicit
    enc_num = str(row.get("encomenda_numero", "") or "").strip()
    created_at = str(row.get("created_at", "") or "").strip()
    causa = str(row.get("causa", "") or "").strip().lower()
    operador = str(row.get("operador", "") or "").strip().lower()
    detalhe = str(row.get("detalhe", "") or "").strip()
    piece_marker = str(row.get("peca_id", "") or row.get("ref_interna", "") or "").strip()
    if created_at:
        return "|".join(part for part in (enc_num, created_at, causa, operador, detalhe) if part)
    return "|".join(part for part in (enc_num, causa, operador, detalhe, piece_marker) if part)


def _op_unique_avaria_minutes_for_rows(rows, *, include_open=True, ts_ref=None):
    grouped = {}
    ts_end = str(ts_ref or now_iso())
    for row in list(rows or []):
        if not isinstance(row, dict):
            continue
        state_norm = norm_text(row.get("estado", ""))
        is_closed = bool(str(row.get("fechada_at", "") or "").strip()) or ("fech" in state_norm)
        dur = max(0.0, _op_float(row.get("duracao_min", 0), 0.0))
        if not is_closed:
            if not include_open:
                continue
            start_txt = str(row.get("created_at", "") or "").strip()
            if start_txt:
                dur = max(0.0, _op_float(iso_diff_minutes(start_txt, ts_end), 0.0))
        key = _op_avaria_group_key(row) or str(id(row))
        grouped[key] = max(grouped.get(key, 0.0), dur)
    return round(sum(grouped.values()), 4)


def _op_unique_avaria_minutes_for_encomenda(data, enc, *, include_open=True, ts_ref=None):
    if not isinstance(data, dict):
        return 0.0
    enc_num = str((enc or {}).get("numero", "") or "").strip()
    rows = [
        row
        for row in list(data.get("op_paragens", []) or [])
        if isinstance(row, dict) and str(row.get("encomenda_numero", "") or "").strip() == enc_num
    ]
    return _op_unique_avaria_minutes_for_rows(rows, include_open=include_open, ts_ref=ts_ref)


def _op_sync_piece_live_avaria(peca, avaria_row):
    if not isinstance(peca, dict) or not isinstance(avaria_row, dict):
        return False
    changed = False
    motivo = str(avaria_row.get("causa", "") or avaria_row.get("motivo", "") or "").strip()
    created_at = str(avaria_row.get("created_at", "") or "").strip()
    group_id = str(avaria_row.get("grupo_id", "") or avaria_row.get("group_id", "") or "").strip()
    if not bool(peca.get("avaria_ativa")):
        peca["avaria_ativa"] = True
        changed = True
    if motivo and str(peca.get("avaria_motivo", "") or "").strip() != motivo:
        peca["avaria_motivo"] = motivo
        changed = True
    if group_id and str(peca.get("avaria_grupo_id", "") or "").strip() != group_id:
        peca["avaria_grupo_id"] = group_id
        changed = True
    if created_at and not str(peca.get("avaria_inicio_ts", "") or "").strip():
        peca["avaria_inicio_ts"] = created_at
        changed = True
    if motivo and not str(peca.get("interrupcao_peca_motivo", "") or "").strip():
        peca["interrupcao_peca_motivo"] = motivo
        changed = True
    if created_at and not str(peca.get("interrupcao_peca_ts", "") or "").strip():
        peca["interrupcao_peca_ts"] = created_at
        changed = True
    if str(peca.get("estado", "") or "") != "Avaria":
        peca["estado"] = "Avaria"
        changed = True
    return changed


def _op_piece_has_open_avaria(data, enc_num, peca, avaria_index=None):
    if not isinstance(peca, dict):
        return False
    live_row = _op_live_avaria_row_for_piece(avaria_index or _op_open_avaria_index(data, enc_num), peca)
    return bool(live_row) or bool(peca.get("avaria_ativa"))


def _op_selected_open_avaria_lines(self, enc, pecas):
    enc_num = str((enc or {}).get("numero", "") or "").strip()
    avaria_index = _op_open_avaria_index(getattr(self, "data", {}), enc_num)
    out = []
    seen = set()
    for p in list(pecas or []):
        if not isinstance(p, dict):
            continue
        live_row = _op_live_avaria_row_for_piece(avaria_index, p)
        active = bool(live_row) or bool(p.get("avaria_ativa"))
        if not active:
            continue
        ref = str(p.get("ref_interna", "") or p.get("id", "") or "-").strip() or "-"
        motivo = str(
            (live_row or {}).get("causa", "")
            or p.get("avaria_motivo", "")
            or p.get("interrupcao_peca_motivo", "")
            or "Avaria em aberto"
        ).strip()
        key = (ref, motivo)
        if key in seen:
            continue
        seen.add(key)
        out.append(f"{ref}: {motivo}")
    return out


def _op_block_if_selected_avaria_open(self, enc, pecas, action_label):
    lines = _op_selected_open_avaria_lines(self, enc, pecas)
    if not lines:
        return False
    msg = f"Existe uma avaria aberta. Feche a avaria antes de {action_label}."
    msg += "\n\n" + "\n".join(lines[:6])
    if len(lines) > 6:
        msg += f"\n... +{len(lines) - 6}"
    messagebox.showerror("Avaria aberta", msg)
    return True


def _mark_piece_ops_in_progress(peca, ops, operador):
    try:
        fluxo = ensure_peca_operacoes(peca)
    except Exception:
        return
    sel = {normalize_operacao_nome(x) for x in list(ops or []) if normalize_operacao_nome(x)}
    if not sel:
        return
    for op in fluxo:
        nome = normalize_operacao_nome(op.get("nome", ""))
        if nome in sel and "concl" not in norm_text(op.get("estado", "")):
            op["estado"] = "Em producao"
            if not op.get("inicio"):
                op["inicio"] = now_iso()
            if operador:
                op["user"] = operador
    peca["operacoes_fluxo"] = fluxo


def _lock_owner_label(owner):
    o = str(owner or "").strip()
    return o if o else "Sem operador (lock antigo)"


def _active_operator_names(self):
    names = set()
    try:
        for o in list((self.data or {}).get("operadores", []) or []):
            if isinstance(o, dict):
                for k in ("nome", "name", "user", "utilizador", "username", "id"):
                    v = str(o.get(k, "") or "").strip()
                    if v:
                        names.add(v)
            else:
                v = str(o or "").strip()
                if v:
                    names.add(v)
    except Exception:
        pass
    try:
        cur = str(self.op_user.get() or "").strip() if hasattr(self, "op_user") else ""
        if cur:
            names.add(cur)
    except Exception:
        pass
    return names


def _current_operator_posto(self):
    try:
        role = str((getattr(self, "user", {}) or {}).get("role", "") or "").strip().lower()
        login_user = str((getattr(self, "user", {}) or {}).get("username", "") or "").strip()
    except Exception:
        role = ""
        login_user = ""
    try:
        user = str(self.op_user.get() or "").strip()
    except Exception:
        user = ""
    if role == "operador" and login_user:
        user = login_user
    if not user:
        return ""
    if role != "operador":
        try:
            posto = str(self.op_posto.get() or "").strip() if hasattr(self, "op_posto") else ""
        except Exception:
            posto = ""
        if posto:
            return posto
    try:
        return str((self.data or {}).get("operador_posto_map", {}).get(user, "") or "").strip()
    except Exception:
        return ""


def _format_event_info_with_posto(self, base_text):
    posto = _current_operator_posto(self)
    base = str(base_text or "").strip()
    if posto:
        if base:
            return f"[POSTO:{posto}] {base}"
        return f"[POSTO:{posto}]"
    return base


def _on_operador_user_change(self):
    try:
        role = str((getattr(self, "user", {}) or {}).get("role", "") or "").strip().lower()
        login_user = str((getattr(self, "user", {}) or {}).get("username", "") or "").strip()
    except Exception:
        role = ""
        login_user = ""
    try:
        user = str(self.op_user.get() or "").strip()
    except Exception:
        return
    if role == "operador" and login_user and user != login_user:
        try:
            self.op_user.set(login_user)
        except Exception:
            pass
        user = login_user
    try:
        postos = list((self.data or {}).get("postos_trabalho", []) or [])
    except Exception:
        postos = []
    if not postos:
        postos = ["Geral"]
    try:
        if hasattr(self, "op_posto_cb") and self.op_posto_cb is not None:
            self.op_posto_cb.configure(values=postos)
    except Exception:
        pass
    try:
        pmap = (self.data or {}).get("operador_posto_map", {}) or {}
    except Exception:
        pmap = {}
    current = str(pmap.get(user, "") or "").strip()
    if not current:
        current = str(postos[0])
    try:
        if hasattr(self, "op_posto"):
            self.op_posto.set(current)
    except Exception:
        pass


def _on_operador_posto_change(self):
    try:
        role = str((getattr(self, "user", {}) or {}).get("role", "") or "").strip().lower()
        login_user = str((getattr(self, "user", {}) or {}).get("username", "") or "").strip()
    except Exception:
        role = ""
        login_user = ""
    try:
        user = str(self.op_user.get() or "").strip()
        posto = str(self.op_posto.get() or "").strip() if hasattr(self, "op_posto") else ""
    except Exception:
        return
    if role == "operador":
        if not login_user:
            return
        try:
            forced = str((self.data or {}).get("operador_posto_map", {}).get(login_user, "") or "").strip() or "Geral"
            if hasattr(self, "op_posto") and str(self.op_posto.get() or "").strip() != forced:
                self.op_posto.set(forced)
        except Exception:
            pass
        return
    if not user:
        return
    try:
        if not isinstance((self.data or {}).get("operador_posto_map", None), dict):
            self.data["operador_posto_map"] = {}
        self.data["operador_posto_map"][user] = posto or "Geral"
        save_data(self.data)
    except Exception:
        pass


def _custom_prompt_text(parent, title, prompt, initial_value="", options=None):
    _ensure_configured()
    use_custom = CUSTOM_TK_AVAILABLE and (
        getattr(parent, "op_use_custom", False)
        if hasattr(parent, "op_use_custom")
        else True
    )
    if not (use_custom and ctk is not None):
        try:
            root = parent.root if hasattr(parent, "root") else parent
        except Exception:
            root = parent
        return simple_input(root, title, prompt)

    host = parent.root if hasattr(parent, "root") else parent
    out = {"value": None}
    win = ctk.CTkToplevel(host)
    win.title(title)
    try:
        win.geometry("520x220")
        win.resizable(False, False)
        win.transient(host)
        win.grab_set()
        win.configure(fg_color="#f7f8fb")
    except Exception:
        pass
    frame = ctk.CTkFrame(win, fg_color="#ffffff", corner_radius=10, border_width=1, border_color="#e7cfd3")
    frame.pack(fill="both", expand=True, padx=10, pady=10)
    ctk.CTkLabel(frame, text=prompt, font=("Segoe UI", 14, "bold"), text_color="#1e293b", anchor="w").pack(fill="x", padx=12, pady=(12, 8))
    var = StringVar(value=str(initial_value or ""))
    opts = [str(o or "").strip() for o in list(options or []) if str(o or "").strip()]
    if opts:
        combo_var = StringVar(value=opts[0])
        ctk.CTkComboBox(frame, variable=combo_var, values=opts, height=36, font=("Segoe UI", 13)).pack(fill="x", padx=12, pady=(0, 6))
        ctk.CTkLabel(
            frame,
            text="Pode confirmar a causa acima ou escrever um motivo personalizado:",
            font=("Segoe UI", 11),
            text_color="#475569",
            anchor="w",
        ).pack(fill="x", padx=12, pady=(0, 4))
    else:
        combo_var = None
    ent = ctk.CTkEntry(frame, textvariable=var, height=38, font=("Segoe UI", 14))
    ent.pack(fill="x", padx=12, pady=(0, 10))
    try:
        ent.focus_force()
    except Exception:
        pass

    btns = ctk.CTkFrame(frame, fg_color="transparent")
    btns.pack(fill="x", padx=12, pady=(0, 12))

    def _ok():
        txt = str(var.get() or "").strip()
        if not txt and combo_var is not None:
            txt = str(combo_var.get() or "").strip()
        out["value"] = txt
        win.destroy()

    def _cancel():
        out["value"] = None
        win.destroy()

    ctk.CTkButton(btns, text="Cancelar", command=_cancel, width=140).pack(side="right", padx=(8, 0))
    ctk.CTkButton(
        btns,
        text="Confirmar",
        command=_ok,
        width=160,
        fg_color="#f59e0b",
        hover_color="#d97706",
        text_color="#ffffff",
    ).pack(side="right")
    win.bind("<Return>", lambda _e: _ok())
    win.bind("<Escape>", lambda _e: _cancel())
    try:
        host.wait_window(win)
    except Exception:
        pass
    return out["value"]


def _metalurgica_paragem_options():
    return [
        "Avaria Laser - Fonte",
        "Avaria Laser - Cabeca de corte",
        "Troca de lente/bocal",
        "Falta de gas (O2/N2/Ar)",
        "Pressao de gas instavel",
        "Falta de chapa",
        "Chapa empenada/defeituosa",
        "Programa NC incorreto",
        "Setup / Afinacao",
        "Troca de ferramenta (quinagem)",
        "Avaria Quinadora",
        "Avaria Soldadura",
        "Avaria Roscagem/Furacao",
        "Falta de operador",
        "Falta de ponte rolante/empilhador",
        "Falha eletrica",
        "Falha pneumatica",
        "Paragem por controlo de qualidade",
        "Paragem de seguranca",
        "Manutencao preventiva",
        "Outro",
    ]


def _interrupcao_operacional_options():
    return [
        "Mudanca de Turno",
        "Alteracao de Prioridades",
        "Aguardar material/documentacao",
        "Outros (Registar Manual)",
    ]


def _prompt_interrupcao_operacional(self):
    motivo = _custom_prompt_text(
        self,
        "Interromper Peca",
        "Selecione o motivo da interrupcao:",
        options=_interrupcao_operacional_options(),
    )
    if motivo is None:
        return None
    motivo = str(motivo or "").strip()
    if not motivo:
        motivo = "Pausa operacional"
    if "registar manual" in norm_text(motivo) or norm_text(motivo) == "outros":
        manual = _custom_prompt_text(
            self,
            "Motivo Manual",
            "Descreva o motivo da interrupcao:",
        )
        if manual is None:
            return None
        manual = str(manual or "").strip()
        motivo = manual or "Outros"
    return motivo

def _ordens_select_opp_custom(self, opp):
    _ensure_configured()
    self.ordens_sel_opp = opp
    for k, btn in self._ordens_btns.items():
        try:
            if k == opp:
                btn.configure(fg_color="#b42318", text_color="white")
            else:
                btn.configure(fg_color=btn._default_fg_color if hasattr(btn, "_default_fg_color") else "#fbecee", text_color="black")
        except Exception:
            pass

def _ordens_get_selected_opp(self):
    _ensure_configured()
    if self.ordens_use_custom:
        return (self.ordens_sel_opp or "").strip()
    if not hasattr(self, "tbl_ordens"):
        return ""
    sel = self.tbl_ordens.selection()
    if not sel:
        return ""
    try:
        return str(self.tbl_ordens.item(sel[0], "values")[0]).strip()
    except Exception:
        return ""

def _ordens_extract_year(enc):
    _ensure_configured()
    if not isinstance(enc, dict):
        return ""
    try:
        return str(
            encomendas_actions._enc_extract_year(
                enc.get("data_criacao", ""),
                enc.get("data_entrega", ""),
                enc.get("numero", ""),
            )
            or ""
        ).strip()
    except Exception:
        for raw in (enc.get("data_criacao", ""), enc.get("data_entrega", ""), enc.get("numero", "")):
            txt = str(raw or "").strip()
            if len(txt) >= 4 and txt[:4].isdigit():
                return txt[:4]
        return ""

def refresh_ordens_year_options(self, keep_selection=True):
    _ensure_configured()
    current_year = str(datetime.now().year)
    selected = (
        str(self.ordens_year_filter.get() or "").strip()
        if keep_selection and hasattr(self, "ordens_year_filter")
        else current_year
    )
    years = {current_year}
    for enc in self.data.get("encomendas", []):
        y = _ordens_extract_year(enc)
        if y:
            years.add(y)
    try:
        for y in encomendas_actions._mysql_encomendas_years_cached(self):
            ys = str(y or "").strip()
            if ys:
                years.add(ys)
    except Exception:
        pass
    year_values = sorted(
        years,
        key=lambda x: int(x) if str(x).isdigit() else 0,
        reverse=True,
    )
    values = year_values + ["Todos"]
    if selected not in values:
        selected = current_year if current_year in values else values[0]
    if hasattr(self, "ordens_year_filter"):
        self.ordens_year_filter.set(selected)
    try:
        cb = getattr(self, "ordens_year_cb", None)
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

def _on_ordens_year_change(self, value=None):
    _ensure_configured()
    try:
        if value is not None and hasattr(self, "ordens_year_filter"):
            self.ordens_year_filter.set(str(value))
    except Exception:
        pass
    year_value = (
        str(self.ordens_year_filter.get() or "").strip()
        if hasattr(self, "ordens_year_filter")
        else ""
    )
    if not year_value:
        year_value = str(datetime.now().year)
    cache_key = f"ordens:{str(year_value).strip().lower()}"
    try:
        now_ts = time.time()
    except Exception:
        now_ts = 0
    last_key = str(getattr(self, "_enc_mysql_reload_key", "") or "")
    last_ts = float(getattr(self, "_enc_mysql_reload_ts", 0) or 0)
    ttl_sec = 45
    if (cache_key != last_key) or ((now_ts - last_ts) > ttl_sec):
        try:
            encomendas_actions._reload_encomendas_from_mysql(self, year_value=year_value)
            self._enc_mysql_reload_key = cache_key
            self._enc_mysql_reload_ts = now_ts
        except Exception:
            pass
    try:
        self.refresh_ordens_year_options(keep_selection=True)
    except Exception:
        pass
    self.refresh_ordens()

def refresh_operador_year_options(self, keep_selection=True):
    _ensure_configured()
    current_year = str(datetime.now().year)
    selected = (
        str(self.op_year_filter.get() or "").strip()
        if keep_selection and hasattr(self, "op_year_filter")
        else current_year
    )
    years = {current_year}
    for enc in self.data.get("encomendas", []):
        y = _ordens_extract_year(enc)
        if y:
            years.add(y)
    try:
        for y in encomendas_actions._mysql_encomendas_years_cached(self):
            ys = str(y or "").strip()
            if ys:
                years.add(ys)
    except Exception:
        pass
    year_values = sorted(
        years,
        key=lambda x: int(x) if str(x).isdigit() else 0,
        reverse=True,
    )
    values = year_values + ["Todos"]
    if selected not in values:
        selected = current_year if current_year in values else values[0]
    if hasattr(self, "op_year_filter"):
        self.op_year_filter.set(selected)
    try:
        cb = getattr(self, "op_year_cb", None)
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

def _on_operador_year_change(self, value=None):
    _ensure_configured()
    try:
        if value is not None and hasattr(self, "op_year_filter"):
            self.op_year_filter.set(str(value))
    except Exception:
        pass
    year_value = (
        str(self.op_year_filter.get() or "").strip()
        if hasattr(self, "op_year_filter")
        else ""
    )
    if not year_value:
        year_value = str(datetime.now().year)
    cache_key = f"operador:{str(year_value).strip().lower()}"
    try:
        now_ts = time.time()
    except Exception:
        now_ts = 0
    last_key = str(getattr(self, "_enc_mysql_reload_key", "") or "")
    last_ts = float(getattr(self, "_enc_mysql_reload_ts", 0) or 0)
    ttl_sec = 45
    if (cache_key != last_key) or ((now_ts - last_ts) > ttl_sec):
        try:
            encomendas_actions._reload_encomendas_from_mysql(self, year_value=year_value)
            self._enc_mysql_reload_key = cache_key
            self._enc_mysql_reload_ts = now_ts
        except Exception:
            pass
    try:
        self.refresh_operador_year_options(keep_selection=True)
    except Exception:
        pass
    self.refresh_operador()

def refresh_operador(self):
    _ensure_configured()
    if getattr(self, "_refresh_operador_busy", False):
        self._refresh_operador_pending = True
        return
    self._refresh_operador_busy = True
    self._refresh_operador_pending = False
    try:
        force_sync = False
        try:
            force_sync = bool(getattr(self, "_op_force_runtime_sync", False))
            self._op_force_runtime_sync = False
            if force_sync:
                self._op_force_ops_status_refresh = True
            _op_runtime_state_sync(self, force=force_sync, ttl_sec=4.0)
        except Exception:
            pass

        def _estado_tag_val(estado):
            try:
                en = norm_text(estado)
            except Exception:
                en = str(estado or "").strip().lower()
            if "avari" in en:
                return "op_avaria"
            if "produ" in en:
                return "op_em_curso"
            if "paus" in en:
                return "op_pausada"
            if "concl" in en:
                return "op_concluida"
            return None

        def _match_estado(estado):
            f = (self.op_status_filter.get() if hasattr(self, "op_status_filter") else "Ativas") or "Ativas"
            fn = f.strip().lower()
            try:
                en = norm_text(estado)
            except Exception:
                en = str(estado or "").strip().lower()
            if fn in ("todas", "todos", "all"):
                return True
            if "ativ" in fn:
                return ("prepar" in en) or ("produ" in en) or ("paus" in en) or ("avari" in en)
            if "prepar" in fn:
                return "prepar" in en
            if "produ" in fn:
                return "produ" in en
            if "paus" in fn:
                return "paus" in en
            if "avari" in fn:
                return "avari" in en
            if "concl" in fn:
                return "concl" in en
            return True

        def _match_ano(enc):
            ano_filtro = (
                str(self.op_year_filter.get() or "").strip()
                if hasattr(self, "op_year_filter")
                else "Todos"
            )
            ano_norm = ano_filtro.lower()
            if ano_norm in ("todos", "todas", "all") or not ano_filtro:
                return True
            enc_year = _ordens_extract_year(enc)
            try:
                est_norm = norm_text(enc.get("estado", ""))
            except Exception:
                est_norm = str(enc.get("estado", "") or "").strip().lower()
            if "concl" in est_norm:
                return enc_year == ano_filtro
            return True

        def _match_query(enc):
            query = (
                str(self.op_quick_filter.get() or "").strip().lower()
                if hasattr(self, "op_quick_filter")
                else ""
            )
            if not query:
                return True
            numero = str(enc.get("numero", "") or "")
            cliente = str(enc.get("cliente", "") or "")
            estado = str(enc.get("estado", "") or "")
            haystack = f"{numero} {cliente} {estado}".lower()
            return query in haystack

        try:
            if hasattr(self, "op_status_segment") and self.op_status_segment.winfo_exists():
                desired = self.op_status_filter.get() or "Ativas"
                current = ""
                try:
                    current = self.op_status_segment.get() or ""
                except Exception:
                    current = ""
                if current != desired:
                    self._suppress_operador_filter_cb = True
                    self.op_status_segment.set(desired)
        except Exception:
            pass
        finally:
            self._suppress_operador_filter_cb = False

        selected_num = self.op_sel_enc_num
        if not self.op_use_full_custom and hasattr(self, "op_tbl_enc"):
            sel = self.op_tbl_enc.selection()
            if sel:
                try:
                    selected_num = self.op_tbl_enc.item(sel[0], "values")[0]
                except Exception:
                    pass
        if self.op_use_full_custom:
            for w in self.op_enc_list.winfo_children():
                w.destroy()
            for idx, e in enumerate(self.data["encomendas"]):
                if not _match_estado(e.get("estado", "")):
                    continue
                if not _match_ano(e):
                    continue
                if not _match_query(e):
                    continue
                tag = _estado_tag_val(e.get("estado", ""))
                bg = (
                    COR_OP_AVARIA
                    if tag == "op_avaria"
                    else COR_OP_PRODUCAO
                    if tag == "op_em_curso"
                    else COR_OP_PAUSADA
                    if tag == "op_pausada"
                    else COR_OP_CONCLUIDA
                    if tag == "op_concluida"
                    else "#fff8f9"
                    if idx % 2 == 0
                    else "#fff0f2"
                )
                row = ctk.CTkFrame(self.op_enc_list, fg_color=bg, corner_radius=6)
                row.pack(fill="x", pady=2, padx=2)
                txt = f"{e.get('numero','')}   {e.get('cliente','')}   {e.get('estado','')}"
                btn = ctk.CTkButton(
                    row,
                    text=txt,
                    anchor="w",
                    fg_color=bg,
                    hover_color="#f8dfe3",
                    text_color="black",
                    height=50,
                    font=("Segoe UI", 14, "bold"),
                    command=lambda num=e.get("numero"): self._op_select_enc_custom(num),
                )
                btn.pack(fill="x", padx=4, pady=4)
        else:
            children = self.op_tbl_enc.get_children()
            if children:
                self.op_tbl_enc.delete(*children)
            target_iid = None
            for e in self.data["encomendas"]:
                if not _match_estado(e.get("estado", "")):
                    continue
                if not _match_ano(e):
                    continue
                if not _match_query(e):
                    continue
                tag = _estado_tag_val(e.get("estado", ""))
                iid = self.op_tbl_enc.insert("", END, values=(e["numero"], e["cliente"], e["estado"]), tags=(tag,) if tag else ())
                if selected_num and e.get("numero") == selected_num:
                    target_iid = iid
            self.op_tbl_enc.tag_configure("op_em_curso", background=COR_OP_PRODUCAO)
            self.op_tbl_enc.tag_configure("op_pausada", background=COR_OP_PAUSADA)
            self.op_tbl_enc.tag_configure("op_concluida", background=COR_OP_CONCLUIDA)
            self.op_tbl_enc.tag_configure("op_avaria", background=COR_OP_AVARIA)
            if target_iid:
                try:
                    self.op_tbl_enc.selection_set(target_iid)
                    self.op_tbl_enc.focus(target_iid)
                    self.op_tbl_enc.see(target_iid)
                    self.op_sel_enc_num = selected_num
                except Exception:
                    pass
    finally:
        self._refresh_operador_busy = False
        if getattr(self, "_refresh_operador_pending", False):
            self._refresh_operador_pending = False
            try:
                self.root.after(20, self.refresh_operador)
            except Exception:
                pass

def refresh_ordens(self):
    _ensure_configured()
    if getattr(self, "_refresh_ordens_busy", False):
        self._refresh_ordens_pending = True
        return
    self._refresh_ordens_busy = True
    self._refresh_ordens_pending = False
    try:
        if self.ordens_use_custom and hasattr(self, "ordens_list"):
            for w in self.ordens_list.winfo_children():
                w.destroy()
            self._ordens_btns = {}
        elif hasattr(self, "tbl_ordens") and self.tbl_ordens:
            children = self.tbl_ordens.get_children()
            if children:
                self.tbl_ordens.delete(*children)
        else:
            return
        query = self.opp_search.get().strip().lower() if hasattr(self, "opp_search") else ""
        estado_filtro = (
            self.ordens_status_filter.get().strip().lower()
            if hasattr(self, "ordens_status_filter")
            else "ativas"
        )
        ano_filtro = (
            self.ordens_year_filter.get().strip()
            if hasattr(self, "ordens_year_filter")
            else "Todos"
        )
        ano_norm = ano_filtro.lower()
        try:
            if hasattr(self, "ordens_status_segment") and self.ordens_status_segment.winfo_exists():
                desired = self.ordens_status_filter.get() or "Ativas"
                current = ""
                try:
                    current = self.ordens_status_segment.get() or ""
                except Exception:
                    current = ""
                if current != desired:
                    self._suppress_ordens_filter_cb = True
                    self.ordens_status_segment.set(desired)
        except Exception:
            pass
        finally:
            self._suppress_ordens_filter_cb = False
        row_i = 0
        changed = False
        for enc in self.data.get("encomendas", []):
            enc_year = _ordens_extract_year(enc)
            try:
                enc_est_norm = norm_text(enc.get("estado", ""))
            except Exception:
                enc_est_norm = str(enc.get("estado", "") or "").strip().lower()
            enc_is_concl = "concl" in enc_est_norm
            if ano_norm not in ("todos", "todas", "all") and ano_filtro and enc_is_concl and enc_year != ano_filtro:
                continue
            for p in encomenda_pecas(enc):
                if not p.get("opp"):
                    p["opp"] = next_opp_numero(self.data)
                    changed = True
                row = (
                    p.get("opp", ""),
                    enc.get("numero", ""),
                    p.get("ref_interna", ""),
                    p.get("ref_externa", ""),
                    p.get("material", ""),
                    p.get("espessura", ""),
                    fmt_num(p.get("quantidade_pedida", 0)),
                    p.get("estado", ""),
                )
                est_norm = norm_text(p.get("estado", ""))
                if estado_filtro and estado_filtro not in ("todas", "todos", "all"):
                    if "ativ" in estado_filtro:
                        if ("prepar" not in est_norm) and ("produ" not in est_norm) and ("incomplet" not in est_norm) and ("interromp" not in est_norm):
                            continue
                    elif "prepar" in estado_filtro and "prepar" not in est_norm:
                        continue
                    elif "produ" in estado_filtro and "produ" not in est_norm:
                        continue
                    elif "concl" in estado_filtro and "concl" not in est_norm:
                        continue
                if query and not any(query in str(v).lower() for v in row):
                    continue
                tag = "ord_even" if row_i % 2 == 0 else "ord_odd"
                est_txt = norm_text(p.get("estado", ""))
                if ("produ" in est_txt) or ("incomplet" in est_txt) or ("interromp" in est_txt):
                    tag = "ord_em_producao"
                elif "concl" in est_txt:
                    tag = "ord_concluida"
                if self.ordens_use_custom and hasattr(self, "ordens_list"):
                    bg = COR_OP_PRODUCAO if tag == "ord_em_producao" else COR_OP_CONCLUIDA if tag == "ord_concluida" else "#f8fbff" if row_i % 2 == 0 else "#edf4ff"
                    card = ctk.CTkFrame(self.ordens_list, fg_color=bg, corner_radius=7)
                    card.pack(fill="x", padx=2, pady=2)
                    txt = (
                        f"{row[0]}   |   {row[1]}   |   {row[2]}   |   {row[3]}   |   "
                        f"{row[4]} {row[5]}mm   |   Qtd: {row[6]}   |   {row[7]}"
                    )
                    btn = ctk.CTkButton(
                        card,
                        text=txt,
                        anchor="w",
                        fg_color=bg,
                        hover_color="#f8dfe3",
                        text_color="black",
                        height=42,
                        font=("Segoe UI", 12, "bold"),
                        command=lambda opp=row[0]: self._ordens_select_opp_custom(opp),
                    )
                    btn.pack(fill="x", padx=4, pady=4)
                    self._ordens_btns[row[0]] = btn
                    if self.ordens_sel_opp and self.ordens_sel_opp == row[0]:
                        self._ordens_select_opp_custom(row[0])
                else:
                    self.tbl_ordens.insert("", END, values=row, tags=(tag,))
                row_i += 1
        if changed:
            save_data(self.data)
    finally:
        self._refresh_ordens_busy = False
        if getattr(self, "_refresh_ordens_pending", False):
            self._refresh_ordens_pending = False
            try:
                self.root.after(20, self.refresh_ordens)
            except Exception:
                pass

def on_operador_select(self, _):
    _ensure_configured()
    if self.op_use_full_custom:
        numero = self.op_sel_enc_num
    else:
        sel = self.op_tbl_enc.selection()
        if not sel:
            return
        numero = self.op_tbl_enc.item(sel[0], "values")[0]
    if not numero:
        return
    self.op_sel_enc_num = numero
    enc = self.get_encomenda_by_numero(numero)
    prev_mat = str(getattr(self, "op_material", "") or "")
    prev_esp = str(getattr(self, "op_espessura", "") or "")
    prev_enc = str(getattr(self, "op_selected_enc_ctx", "") or "")
    self.op_selected_enc_ctx = str(numero)
    if prev_enc != str(numero):
        prev_mat = ""
        prev_esp = ""
        self.op_sel_pecas_ids = set()
    if self.op_use_full_custom:
        for w in self.op_esp_list.winfo_children():
            w.destroy()
        for w in self.op_pecas_list.winfo_children():
            w.destroy()
    else:
        esp_children = self.op_tbl_esp.get_children()
        if esp_children:
            self.op_tbl_esp.delete(*esp_children)
        pecas_children = self.op_tbl_pecas.get_children()
        if pecas_children:
            self.op_tbl_pecas.delete(*pecas_children)
    if not enc:
        return
    estado_dirty = False
    self.op_espessura = prev_esp or None
    self.op_material = prev_mat or None
    if self.op_use_full_custom:
        try:
            self.op_ctx_operador_lbl.configure(text=str(self.op_user.get() or "-"))
            self.op_ctx_posto_lbl.configure(text=str(self.op_posto.get() or "-"))
            self.op_ctx_estado_lbl.configure(text=str(enc.get("estado") or "Preparacao"))
        except Exception:
            pass
    row_i = 0
    esp_values = []
    selected_found = False
    total_esp = 0
    for m in enc.get("materiais", []):
        for e in m.get("espessuras", []):
            mat_val = str(m.get("material", "") or "")
            esp_val = str(e.get("espessura", "") or "")
            total_esp += 1
            esp_values.append(self._op_esp_label(mat_val, esp_val))
            is_sel = bool(prev_mat and prev_esp) and mat_val == prev_mat and fmt_num(esp_val) == fmt_num(prev_esp)
            if is_sel:
                selected_found = True
                self.op_material = mat_val
                self.op_espessura = esp_val
            prev_estado_esp = str(e.get("estado", "") or "")
            planeado, produzido = self.atualizar_estado_espessura(e)
            if str(e.get("estado", "") or "") != prev_estado_esp:
                estado_dirty = True
            tag = "op_even" if row_i % 2 == 0 else "op_odd"
            if e.get("estado") == "Em producao":
                tag = "op_em_curso"
            elif e.get("estado") == "Avaria":
                tag = "op_avaria"
            elif e.get("estado") == "Em pausa":
                tag = "op_pausada"
            elif e.get("estado") == "Concluida":
                tag = "op_concluida"
            if self.op_use_full_custom:
                bg = _op_base_bg_from_tag(tag, row_i)
                if is_sel:
                    bg = "#dbeafe"
                badge_fg, badge_txt = _op_badge_style(tag)
                row = ctk.CTkFrame(self.op_esp_list, fg_color=bg, corner_radius=6)
                row.pack(fill="x", pady=2, padx=2)
                for idx, weight in enumerate((4, 2, 1, 1, 2)):
                    row.grid_columnconfigure(idx, weight=weight)
                click_cmd = lambda mat=m.get('material',''), esp=e.get('espessura',''): self._op_select_esp_custom(mat, esp)
                values = [
                    (str(m.get("material", "") or "-"), 0, "w", OP_ROW_FONT_BOLD),
                    (f"{fmt_num(e.get('espessura', ''))} mm", 1, "w", OP_ROW_FONT_BOLD),
                    (fmt_num(planeado), 2, "e", OP_ROW_FONT_BOLD),
                    (fmt_num(produzido), 3, "e", OP_ROW_FONT_BOLD),
                ]
                for text, col, sticky, font in values:
                    lbl = ctk.CTkLabel(row, text=text, font=font, text_color="#0f172a")
                    lbl.grid(row=0, column=col, sticky=sticky, padx=8, pady=5)
                    _op_bind_click(lbl, click_cmd)
                badge = ctk.CTkLabel(
                    row,
                    text=str(e.get("estado", "") or "Preparacao"),
                    fg_color=badge_fg,
                    text_color=badge_txt,
                    corner_radius=999,
                    font=OP_ROW_FONT_SMALL,
                    height=24,
                )
                badge.grid(row=0, column=4, sticky="w", padx=8, pady=4)
                _op_bind_click(badge, click_cmd)
                _op_bind_click(row, click_cmd)
            else:
                self.op_tbl_esp.insert(
                    "",
                    END,
                    values=(m.get("material",""), e.get("espessura", ""), fmt_num(planeado), fmt_num(produzido), e.get("estado", "Preparacao")),
                    tags=(tag,) if tag else (),
                )
            row_i += 1
    if self.op_use_full_custom and hasattr(self, "op_esp_cb"):
        try:
            self.op_esp_cb.configure(values=esp_values or [""])
            if selected_found and self.op_material and self.op_espessura:
                label = self._op_esp_label(self.op_material, self.op_espessura)
                if label in esp_values:
                    self.op_esp_var.set(label)
                else:
                    self.op_esp_var.set("")
            else:
                self.op_esp_var.set("")
        except Exception:
            pass
    if not selected_found:
        self.op_espessura = None
        self.op_material = None
    if self.op_use_full_custom:
        try:
            active = self._op_esp_label(self.op_material, self.op_espessura) if self.op_material and self.op_espessura else "-"
            self.op_ctx_esp_sel_lbl.configure(text=active)
        except Exception:
            pass
    prev_enc_estado = str(enc.get("estado", "") or "")
    update_estado_encomenda_por_espessuras(enc)
    if str(enc.get("estado", "") or "") != prev_enc_estado:
        estado_dirty = True
    self.preencher_pecas_operador(enc, None, None)
    if estado_dirty:
        save_data(self.data)

def clear_operador_selection(self):
    _ensure_configured()
    self.op_espessura = None
    self.op_material = None
    if self.op_use_full_custom:
        for w in self.op_esp_list.winfo_children():
            w.destroy()
        for w in self.op_pecas_list.winfo_children():
            w.destroy()
        self.op_sel_peca = None
        self.op_sel_pecas_ids = set()
    else:
        for t in [self.op_tbl_esp, self.op_tbl_pecas]:
            children = t.get_children()
            if children:
                t.delete(*children)

def clear_operador_esp(self):
    _ensure_configured()
    self.op_espessura = None
    self.op_material = None
    if self.op_use_full_custom:
        for w in self.op_pecas_list.winfo_children():
            w.destroy()
        self.op_sel_peca = None
        self.op_sel_pecas_ids = set()
    else:
        children = self.op_tbl_pecas.get_children()
        if children:
            self.op_tbl_pecas.delete(*children)
    self.op_sel_peca = None
    self.op_sel_pecas_ids = set()

def on_operador_select_espessura(self, _):
    _ensure_configured()
    enc = self.get_operador_encomenda(show_error=False)
    if not enc:
        return
    if self.op_use_full_custom:
        mat = self.op_material
        esp = self.op_espessura
    else:
        sel = self.op_tbl_esp.selection()
        if not sel:
            return
        mat = self.op_tbl_esp.item(sel[0], "values")[0]
        esp = self.op_tbl_esp.item(sel[0], "values")[1]
    self.op_espessura = esp
    self.op_material = mat
    try:
        if self.op_use_full_custom and hasattr(self, "op_esp_var"):
            self.op_esp_var.set(self._op_esp_label(mat, esp))
    except Exception:
        pass
    self.preencher_pecas_operador(enc, mat, esp)

def preencher_pecas_operador(self, enc, mat, esp):
    _ensure_configured()
    if self.op_use_full_custom:
        for w in self.op_pecas_list.winfo_children():
            w.destroy()
    else:
        children = self.op_tbl_pecas.get_children()
        if children:
            self.op_tbl_pecas.delete(*children)
    row_i = 0
    if self.op_use_full_custom and (not hasattr(self, "op_sel_pecas_ids") or not isinstance(getattr(self, "op_sel_pecas_ids", None), set)):
        self.op_sel_pecas_ids = set()
    total_visible = 0
    total_exec = 0
    total_pausa = 0
    total_avaria = 0
    total_concl = 0
    first_open_avaria_txt = ""
    estado_dirty = False
    enc_num = str((enc or {}).get("numero", "") or "").strip()
    force_ops_refresh = bool(getattr(self, "_op_force_ops_status_refresh", False))
    ops_status_map = _mysql_ops_status_for_order(enc_num, cache_owner=self, force=force_ops_refresh) if enc_num else {}
    if force_ops_refresh:
        self._op_force_ops_status_refresh = False
    avaria_index = _op_open_avaria_index(getattr(self, "data", {}), enc_num) if enc_num else {}
    for m in enc.get("materiais", []):
        if mat and m.get("material") != mat:
            continue
        for e in m.get("espessuras", []):
            if esp and e.get("espessura") != esp:
                continue
            for p in e.get("pecas", []):
                produzido = float(p.get("produzido_ok", 0)) + float(p.get("produzido_nok", 0)) + float(p.get("produzido_qualidade", 0))
                estado_peca = norm_text(p.get("estado", ""))
                pid = str(p.get("id", "") or "").strip()
                live_avaria = _op_live_avaria_row_for_piece(avaria_index, p)
                if live_avaria and _op_sync_piece_live_avaria(p, live_avaria):
                    estado_dirty = True
                avaria_ativa = bool(p.get("avaria_ativa")) or bool(live_avaria)
                avaria_motivo = str(
                    (live_avaria or {}).get("causa", "")
                    or p.get("avaria_motivo", "")
                    or p.get("interrupcao_peca_motivo", "")
                    or ""
                ).strip()
                if avaria_ativa and not first_open_avaria_txt:
                    ref_txt = str(p.get("ref_interna", "") or p.get("id", "") or "-").strip() or "-"
                    first_open_avaria_txt = f"Avaria aberta em {ref_txt}: {avaria_motivo or 'Sem motivo registado'}"
                pend_ops = " + ".join(peca_operacoes_pendentes(p)) or "-"
                op_status_rows = list(ops_status_map.get(pid, []) or [])
                em_curso = []
                owners = []
                current_ops = set()
                posto_map = dict(self.data.get("operador_posto_map", {}) or {})
                for r in op_status_rows:
                    st = norm_text(r.get("estado", ""))
                    if "produ" in st:
                        opn = str(r.get("operacao", "") or "").strip()
                        own = _lock_owner_label(r.get("operador", ""))
                        own_key = str(r.get("operador", "") or "").strip()
                        posto = str(posto_map.get(own_key, "") or "").strip()
                        own_show = f"{own}@{posto}" if posto and own_key else own
                        if opn:
                            current_ops.add(normalize_operacao_nome(opn))
                            em_curso.append(f"{opn}({own_show})")
                            if own_show not in owners:
                                owners.append(own_show)
                em_curso_txt = ", ".join(em_curso) if em_curso else "-"
                dono_txt = ", ".join(owners) if owners else "-"
                fluxo = ensure_peca_operacoes(p)
                has_ops_progress = any("concl" in norm_text(op.get("estado", "")) for op in fluxo)
                qty_started = bool(parse_float(p.get("produzido_ok", 0), 0.0) + parse_float(p.get("produzido_nok", 0), 0.0) + parse_float(p.get("produzido_qualidade", 0), 0.0) > 0)
                started = qty_started or bool(p.get("inicio_producao")) or has_ops_progress
                # Saneamento de estado: sem operação ativa não pode permanecer "Em producao".
                try:
                    estado_ant = str(p.get("estado", "") or "")
                    if em_curso:
                        if estado_ant != "Em producao":
                            p["estado"] = "Em producao"
                            estado_dirty = True
                    elif avaria_ativa:
                        if estado_ant != "Avaria":
                            p["estado"] = "Avaria"
                            estado_dirty = True
                    else:
                        st_norm_now = norm_text(p.get("estado", ""))
                        if ("produ" in st_norm_now) and ("pausad" not in st_norm_now) and ("avari" not in st_norm_now):
                            if qty_started or has_ops_progress:
                                if estado_ant != "Incompleta":
                                    p["estado"] = "Incompleta"
                                    estado_dirty = True
                            else:
                                if estado_ant != "Preparacao":
                                    p["estado"] = "Preparacao"
                                    estado_dirty = True
                                if not p.get("fim_producao"):
                                    p["inicio_producao"] = ""
                except Exception:
                    pass
                estado_peca = norm_text(p.get("estado", ""))
                pausada = bool(started and (not em_curso) and pend_ops != "-" and ("interromp" not in estado_peca) and ("conclu" not in estado_peca))
                # Estado visual em tempo real: se existe operação ativa, prevalece "Em producao".
                estado_visual = str(p.get("estado", "") or "")
                if em_curso:
                    estado_visual = "Em producao"
                elif avaria_ativa:
                    estado_visual = "Avaria"
                elif pausada:
                    estado_visual = "Em producao/Pausada"
                elif "interromp" in estado_peca or "paus" in estado_peca:
                    estado_visual = "Em pausa"
                elif "incomplet" in estado_peca:
                    estado_visual = "Incompleta"
                elif "conclu" in estado_peca:
                    estado_visual = "Concluida"
                total_visible += 1
                if em_curso:
                    total_exec += 1
                elif avaria_ativa:
                    total_avaria += 1
                elif "conclu" in estado_peca:
                    total_concl += 1
                elif pausada or "interromp" in estado_peca or "paus" in estado_peca:
                    total_pausa += 1
                # Mantém consistência do objeto em memória com o estado atual da linha.
                try:
                    if em_curso and norm_text(p.get("estado", "")) != "em producao":
                        p["estado"] = "Em producao"
                        estado_dirty = True
                except Exception:
                    pass
                tag = "op_even" if row_i % 2 == 0 else "op_odd"
                if em_curso:
                    tag = "op_em_curso"
                elif pausada:
                    tag = "op_pausada"
                elif "conclu" in estado_peca:
                    tag = "op_concluida"
                elif avaria_ativa:
                    tag = "op_avaria"
                elif "interromp" in estado_peca or "paus" in estado_peca:
                    tag = "op_interrompida"
                elif "incomplet" in estado_peca:
                    tag = "op_incompleta"
                elif "produ" in estado_peca:
                    tag = "op_em_curso"
                if self.op_use_full_custom:
                    sel = getattr(self, "op_sel_peca", None) or {}
                    pid = str(p.get("id", "") or "").strip()
                    is_checked = bool(pid and pid in getattr(self, "op_sel_pecas_ids", set()))
                    is_selected = (
                        str(sel.get("id", "") or "") == str(p.get("id", "") or "")
                        and str(sel.get("id", "") or "").strip() != ""
                    ) or (
                        str(sel.get("ref_interna", "") or "") == str(p.get("ref_interna", "") or "")
                        and str(sel.get("ref_externa", "") or "") == str(p.get("ref_externa", "") or "")
                        and (str(sel.get("ref_interna", "") or "").strip() or str(sel.get("ref_externa", "") or "").strip())
                    )
                    base_bg = _op_base_bg_from_tag(tag, row_i)
                    bg = "#cfe8ff" if is_selected else base_bg
                    badge_fg, badge_txt = _op_badge_style(tag)
                    row = ctk.CTkFrame(
                        self.op_pecas_list,
                        fg_color=bg,
                        corner_radius=8,
                        border_width=1 if is_selected else 0,
                        border_color="#60a5fa" if is_selected else bg,
                    )
                    row.pack(fill="x", pady=1, padx=1)
                    chk_var = BooleanVar(value=is_checked)

                    def _toggle_piece(_p=p, _v=chk_var):
                        pid_i = str((_p or {}).get("id", "") or "").strip()
                        if not hasattr(self, "op_sel_pecas_ids") or not isinstance(getattr(self, "op_sel_pecas_ids", None), set):
                            self.op_sel_pecas_ids = set()
                        if bool(_v.get()):
                            if pid_i:
                                self.op_sel_pecas_ids.add(pid_i)
                            self.op_sel_peca = _p
                        else:
                            if pid_i:
                                self.op_sel_pecas_ids.discard(pid_i)
                            if str((getattr(self, "op_sel_peca", {}) or {}).get("id", "") or "") == pid_i:
                                self.op_sel_peca = None
                        try:
                            self.preencher_pecas_operador(enc, mat, esp)
                        except Exception:
                            pass

                    click_cmd = lambda px=p: self._op_select_peca_custom(px)
                    main_line = ctk.CTkFrame(row, fg_color="transparent")
                    main_line.pack(fill="x", padx=2, pady=(1, 0))
                    chk_cell = ctk.CTkFrame(main_line, fg_color="transparent", width=OP_PIECE_COLUMNS[0][1], height=32)
                    chk_cell.pack(side="left", fill="y")
                    chk_cell.pack_propagate(False)
                    chk = ctk.CTkCheckBox(
                        chk_cell,
                        text="",
                        variable=chk_var,
                        width=18,
                        checkbox_width=17,
                        checkbox_height=17,
                        command=_toggle_piece,
                    )
                    chk.pack(anchor="w", padx=(5, 0), pady=4)
                    detail_text = ""
                    if avaria_ativa:
                        if avaria_motivo:
                            detail_text = f"Avaria ativa: {avaria_motivo}"
                    elif "interromp" in estado_peca or "paus" in estado_peca:
                        motivo = str(p.get("interrupcao_peca_motivo", "") or "").strip()
                        if motivo:
                            detail_text = f"Interrupção: {motivo}"

                    cells = [
                        (OP_PIECE_COLUMNS[1][1], _op_short_text(str(p.get("ref_interna", "") or "-"), 14), OP_ROW_FONT_BOLD, "w"),
                        (OP_PIECE_COLUMNS[2][1], _op_short_text(str(p.get("ref_externa", "") or "-"), 24), OP_ROW_FONT_BOLD, "w"),
                        (OP_PIECE_COLUMNS[3][1], _op_short_text(em_curso_txt, 18), OP_ROW_FONT, "w"),
                        (OP_PIECE_COLUMNS[5][1], _op_short_text(dono_txt, 10), OP_ROW_FONT, "w"),
                        (OP_PIECE_COLUMNS[6][1], fmt_num(p.get("quantidade_pedida", 0)), OP_ROW_FONT_BOLD, "e"),
                        (OP_PIECE_COLUMNS[7][1], fmt_num(produzido), OP_ROW_FONT_BOLD, "e"),
                    ]
                    for width, text, font, anchor in cells[:3]:
                        lbl = _op_make_fixed_cell(main_line, width, text, font, anchor=anchor)
                        _op_bind_click(lbl, click_cmd)
                    pend_cell = ctk.CTkFrame(main_line, fg_color="transparent", width=OP_PIECE_COLUMNS[4][1], height=32)
                    pend_cell.pack(side="left", fill="y")
                    pend_cell.pack_propagate(False)
                    flow = ensure_peca_operacoes(p)
                    if flow:
                        pend_set = set(peca_operacoes_pendentes(p))
                        flow_line = ctk.CTkFrame(pend_cell, fg_color="transparent")
                        flow_line.pack(fill="both", expand=True, padx=4, pady=4)
                        for op in flow:
                            op_nome = normalize_operacao_nome(op.get("nome", ""))
                            op_state = "Pendente"
                            if op_nome in current_ops:
                                op_state = "Em curso"
                            elif "concl" in norm_text(op.get("estado", "")):
                                op_state = "Concluida"
                            elif avaria_ativa and op_nome in pend_set:
                                op_state = "Avaria"
                            fg_chip, txt_chip = _op_flow_chip_style(op_state)
                            chip = ctk.CTkLabel(
                                flow_line,
                                text=op_nome,
                                fg_color=fg_chip,
                                text_color=txt_chip,
                                corner_radius=999,
                                font=("Segoe UI", 8, "bold"),
                                height=17,
                            )
                            chip.pack(side="left", padx=(0, 5))
                            _op_bind_click(chip, click_cmd)
                    else:
                        lbl = _op_make_fixed_cell(pend_cell, OP_PIECE_COLUMNS[4][1], _op_short_text(pend_ops, 26), OP_ROW_FONT, anchor="w")
                        _op_bind_click(lbl, click_cmd)
                    for width, text, font, anchor in cells[3:]:
                        lbl = _op_make_fixed_cell(main_line, width, text, font, anchor=anchor)
                        _op_bind_click(lbl, click_cmd)
                    badge_cell = ctk.CTkFrame(main_line, fg_color="transparent", width=OP_PIECE_COLUMNS[8][1], height=32)
                    badge_cell.pack(side="left", fill="y")
                    badge_cell.pack_propagate(False)
                    badge = ctk.CTkLabel(
                        badge_cell,
                        text=_op_short_state(estado_visual),
                        fg_color=badge_fg,
                        text_color=badge_txt,
                        corner_radius=999,
                        font=OP_ROW_FONT_SMALL,
                        height=18,
                    )
                    badge.pack(anchor="w", padx=(3, 0), pady=3)
                    _op_bind_click(badge, click_cmd)
                    if detail_text:
                        detail_lbl = ctk.CTkLabel(
                            row,
                            text=detail_text,
                            font=("Segoe UI", 9, "bold"),
                            text_color="#7f1d1d" if avaria_ativa else "#92400e",
                            anchor="w",
                        )
                        detail_lbl.pack(anchor="w", padx=(26, 4), pady=(0, 1))
                        _op_bind_click(detail_lbl, click_cmd)
                    _op_bind_click(row, click_cmd)
                else:
                    self.op_tbl_pecas.insert(
                        "",
                        END,
                        values=(
                            p.get("ref_interna") or "",
                            p.get("ref_externa") or "",
                            pend_ops,
                            em_curso_txt,
                            fmt_num(p.get("quantidade_pedida", 0)),
                            fmt_num(produzido),
                            p.get("estado", ""),
                        ),
                        tags=(tag,) if tag else (),
                    )
                row_i += 1
    if not self.op_use_full_custom:
        self.op_tbl_pecas.tag_configure("op_em_curso", background=COR_OP_PRODUCAO)
        self.op_tbl_pecas.tag_configure("op_pausada", background=COR_OP_PAUSADA)
        self.op_tbl_pecas.tag_configure("op_concluida", background=COR_OP_CONCLUIDA)
        self.op_tbl_pecas.tag_configure("op_avaria", background=COR_OP_AVARIA)
        self.op_tbl_pecas.tag_configure("op_incompleta", background=COR_OP_INCOMPLETA)
        self.op_tbl_pecas.tag_configure("op_interrompida", background=COR_OP_INTERROMPIDA)
        self.op_tbl_pecas.tag_configure("op_even", background="#fff5f6")
        self.op_tbl_pecas.tag_configure("op_odd", background="#fbecee")
    if self.op_use_full_custom:
        try:
            self.op_ctx_total_lbl.configure(text=str(total_visible))
            self.op_ctx_exec_lbl.configure(text=str(total_exec))
            self.op_ctx_pausa_lbl.configure(text=str(total_pausa))
            self.op_ctx_avaria_lbl.configure(text=str(total_avaria))
            self.op_ctx_conc_lbl.configure(text=str(total_concl))
        except Exception:
            pass
    if estado_dirty:
        prev_enc_estado = str(enc.get("estado", "") or "")
        update_estado_encomenda_por_espessuras(enc)
        if str(enc.get("estado", "") or "") != prev_enc_estado:
            estado_dirty = True
    if self.op_use_full_custom:
        try:
            self.op_ctx_estado_lbl.configure(text=str(enc.get("estado") or "Preparacao"))
        except Exception:
            pass
        try:
            if hasattr(self, "op_ctx_alerta_lbl") and self.op_ctx_alerta_lbl.winfo_exists():
                if first_open_avaria_txt:
                    try:
                        if hasattr(self, "op_ctx_alerta_box") and self.op_ctx_alerta_box.winfo_exists():
                            self.op_ctx_alerta_box.configure(
                                fg_color="#fee2e2",
                                border_color="#f87171",
                            )
                    except Exception:
                        pass
                    self.op_ctx_alerta_lbl.configure(
                        text=first_open_avaria_txt,
                        text_color="#7f1d1d",
                    )
                else:
                    try:
                        if hasattr(self, "op_ctx_alerta_box") and self.op_ctx_alerta_box.winfo_exists():
                            self.op_ctx_alerta_box.configure(
                                fg_color="#ecfccb",
                                border_color="#84cc16",
                            )
                    except Exception:
                        pass
                    self.op_ctx_alerta_lbl.configure(
                        text="Sem avarias abertas na seleção.",
                        text_color="#365314",
                    )
        except Exception:
            pass
    # Nao gravar no disco durante render para manter fluidez do menu Operador.

def operador_ajustar_quantidade(self, p, delta):
    _ensure_configured()
    if not self.op_user.get().strip():
        messagebox.showerror("Erro", "Selecione o operador")
        return False
    atual = float(p.get("produzido_ok", 0))
    novo = atual + delta
    if novo < 0:
        novo = 0
    if novo > float(p.get("quantidade_pedida", 0)):
        if not messagebox.askyesno("Confirmar", "A quantidade excede o planeado. Pretende continuar?"):
            return False
    p["produzido_ok"] = novo
    if delta != 0:
        p.setdefault("hist", []).append({
            "ts": now_iso(),
            "user": self.op_user.get().strip(),
            "delta_ok": delta,
        })
    atualizar_estado_peca(p)
    return True

def get_operador_esp_obj(self, enc, mat, esp):
    _ensure_configured()
    for m in enc.get("materiais", []):
        if m.get("material") != mat:
            continue
        for e in m.get("espessuras", []):
            if e.get("espessura") == esp:
                return e
    return None

def get_operador_encomenda(self, show_error=True):
    _ensure_configured()
    if self.op_use_full_custom:
        numero = self.op_sel_enc_num
        if not numero and show_error:
            messagebox.showerror("Erro", "Selecione uma encomenda")
        return self.get_encomenda_by_numero(numero) if numero else None
    sel = self.op_tbl_enc.selection()
    if not sel:
        if show_error:
            messagebox.showerror("Erro", "Selecione uma encomenda")
        return None
    numero = self.op_tbl_enc.item(sel[0], "values")[0]
    return self.get_encomenda_by_numero(numero)

def _op_force_refresh_page(self):
    _ensure_configured()
    detail_restored = False
    try:
        self._op_force_runtime_sync = True
        self._op_force_ops_status_refresh = True
        self.refresh_operador()
    except Exception:
        pass
    try:
        self.restore_operador_selection()
        detail_restored = bool(
            getattr(self, "op_use_full_custom", False)
            and getattr(self, "op_detail_win", None)
            and self.op_detail_win.winfo_exists()
        )
    except Exception:
        pass
    allow_detail_refresh = not bool(getattr(self, "op_use_full_custom", False))
    if getattr(self, "op_use_full_custom", False):
        try:
            allow_detail_refresh = bool(
                getattr(self, "op_detail_win", None)
                and self.op_detail_win.winfo_exists()
            )
        except Exception:
            allow_detail_refresh = False
    if detail_restored and allow_detail_refresh:
        return
    if not allow_detail_refresh:
        return
    try:
        self.on_operador_select(None)
    except Exception:
        pass
    try:
        self.on_operador_select_espessura(None)
    except Exception:
        pass

def _op_action_with_refresh(self, action_fn):
    _ensure_configured()
    result = None
    if callable(action_fn):
        result = action_fn()
    # Refresh unico, com debounce, para nao bloquear a fluidez do menu operador.
    try:
        self._op_force_runtime_sync = True
        self._op_force_ops_status_refresh = True
        prev = getattr(self, "_op_refresh_after_id", None)
        if prev:
            try:
                self.root.after_cancel(prev)
            except Exception:
                pass
        self._op_refresh_after_id = self.root.after(260, lambda _self=self: _op_force_refresh_page(_self))
    except Exception:
        try:
            _op_force_refresh_page(self)
        except Exception:
            pass
    return result

def restore_operador_selection(self):
    _ensure_configured()
    if not getattr(self, "op_last_enc", None):
        return
    numero = self.op_last_enc
    if self.op_use_full_custom:
        self.op_sel_enc_num = numero
        if hasattr(self, "op_detail_win") and self.op_detail_win and self.op_detail_win.winfo_exists():
            try:
                self.on_operador_select(None)
            except Exception:
                pass
            mat = getattr(self, "op_last_mat", None)
            esp = getattr(self, "op_last_esp", None)
            if mat and esp:
                try:
                    self._op_select_esp_custom(mat, esp)
                except Exception:
                    pass
        return
    for item in self.op_tbl_enc.get_children():
        if self.op_tbl_enc.item(item, "values")[0] == numero:
            self.op_tbl_enc.selection_set(item)
            self.on_operador_select(None)
            break
    mat = getattr(self, "op_last_mat", None)
    esp = getattr(self, "op_last_esp", None)
    if mat and esp:
        for item in self.op_tbl_esp.get_children():
            vals = self.op_tbl_esp.item(item, "values")
            if vals[0] == mat and vals[1] == esp:
                self.op_tbl_esp.selection_set(item)
                self.on_operador_select_espessura(None)
                break

def get_operador_peca(self, enc):
    _ensure_configured()
    if not enc:
        return None
    if self.op_use_full_custom:
        if not self.op_sel_peca:
            messagebox.showerror("Erro", "Selecione uma peça")
            return None
        ref_int = self.op_sel_peca.get("ref_interna", "")
        ref_ext = self.op_sel_peca.get("ref_externa", "")
    else:
        sel = self.op_tbl_pecas.selection()
        if not sel:
            messagebox.showerror("Erro", "Selecione uma peça")
            return None
        ref_int = self.op_tbl_pecas.item(sel[0], "values")[0]
        ref_ext = self.op_tbl_pecas.item(sel[0], "values")[1]
    for m in enc.get("materiais", []):
        for e in m.get("espessuras", []):
            for p in e.get("pecas", []):
                if (
                    (ref_int and (p.get("ref_interna") == ref_int))
                    or (ref_ext and (p.get("ref_externa") == ref_ext))
                    or (p.get("id") == ref_int)
                ):
                    if not self.op_material:
                        self.op_material = p.get("material", "")
                    if not self.op_espessura:
                        self.op_espessura = p.get("espessura", "")
                    return p
    return None


def get_operador_pecas(self, enc, show_error=True):
    _ensure_configured()
    if not enc:
        return []
    if self.op_use_full_custom:
        ids = set(getattr(self, "op_sel_pecas_ids", set()) or set())
        if ids:
            out = []
            for m in enc.get("materiais", []):
                for e in m.get("espessuras", []):
                    for p in e.get("pecas", []):
                        pid = str(p.get("id", "") or "").strip()
                        if pid and pid in ids:
                            out.append(p)
            if out:
                return out
        if (not show_error) and not getattr(self, "op_sel_peca", None):
            return []
        p = self.get_operador_peca(enc)
        return [p] if p else []
    if (not show_error) and (not getattr(self, "op_tbl_pecas", None) or not self.op_tbl_pecas.selection()):
        return []
    p = self.get_operador_peca(enc)
    return [p] if p else []


def _op_visible_pecas(self, enc):
    _ensure_configured()
    out = []
    if not enc:
        return out
    mat_sel = str(getattr(self, "op_material", "") or "").strip()
    esp_sel = str(getattr(self, "op_espessura", "") or "").strip()
    esp_sel_n = fmt_num(esp_sel) if esp_sel else ""
    for m in enc.get("materiais", []):
        mat_ok = (not mat_sel) or str(m.get("material", "") or "").strip() == mat_sel
        if not mat_ok:
            continue
        for e in m.get("espessuras", []):
            esp_val = str(e.get("espessura", "") or "").strip()
            esp_ok = (not esp_sel_n) or fmt_num(esp_val) == esp_sel_n
            if not esp_ok:
                continue
            for p in e.get("pecas", []):
                out.append(p)
    return out


def _op_set_select_all_pecas(self, checked):
    _ensure_configured()
    enc = self.get_operador_encomenda(show_error=False)
    if not enc:
        return
    vis = _op_visible_pecas(self, enc)
    if not hasattr(self, "op_sel_pecas_ids") or not isinstance(getattr(self, "op_sel_pecas_ids", None), set):
        self.op_sel_pecas_ids = set()
    ids_vis = {str((p or {}).get("id", "") or "").strip() for p in vis if str((p or {}).get("id", "") or "").strip()}
    if checked:
        self.op_sel_pecas_ids.update(ids_vis)
        if vis:
            self.op_sel_peca = vis[0]
    else:
        self.op_sel_pecas_ids = {pid for pid in self.op_sel_pecas_ids if pid not in ids_vis}
        cur_id = str((getattr(self, "op_sel_peca", {}) or {}).get("id", "") or "").strip()
        if cur_id and cur_id in ids_vis:
            self.op_sel_peca = None
    try:
        self.on_operador_select_espessura(None)
    except Exception:
        pass

def operador_concluir_peca(self):
    _ensure_configured()
    enc = self.get_operador_encomenda()
    pecas = get_operador_pecas(self, enc, show_error=True)
    if not enc or not pecas:
        return
    if len(pecas) > 1 and not bool(getattr(self, "_op_multi_running", False)):
        if not messagebox.askyesno("Fim em seleção", f"Foram selecionadas {len(pecas)} peças.\nAbrir registo de fim para cada peça?"):
            return
        old_ids = set(getattr(self, "op_sel_pecas_ids", set()) or set())
        try:
            self._op_multi_running = True
            total_batch = len(pecas)
            for idx_batch, p_it in enumerate(pecas, start=1):
                self.op_sel_peca = p_it
                self.op_sel_pecas_ids = {str((p_it or {}).get("id", "") or "").strip()}
                self._op_batch_idx = idx_batch
                self._op_batch_total = total_batch
                ok_saved = self.operador_fim_peca()
                if not ok_saved:
                    break
        finally:
            self._op_multi_running = False
            self._op_batch_idx = 0
            self._op_batch_total = 0
            self.op_sel_pecas_ids = old_ids
            try:
                self.on_operador_select_espessura(None)
            except Exception:
                pass
        return
    p = pecas[0]
    if not self.operador_ajustar_quantidade(p, 1):
        return
    update_estado_encomenda_por_espessuras(enc)
    save_data(self.data)
    self.on_operador_select(None)
    self.on_operador_select_espessura(None)

def operador_inserir_qtd(self):
    _ensure_configured()
    enc = self.get_operador_encomenda()
    p = self.get_operador_peca(enc)
    if not enc or not p:
        return
    val = simple_input(self.root, "Quantidade", "Inserir quantidade a adicionar:")
    if val is None:
        return
    try:
        delta = float(val)
    except Exception:
        messagebox.showerror("Erro", "Quantidade inválida")
        return
    if delta <= 0:
        messagebox.showerror("Erro", "Quantidade inválida")
        return
    if not self.operador_ajustar_quantidade(p, delta):
        return
    update_estado_encomenda_por_espessuras(enc)
    save_data(self.data)
    self.on_operador_select(None)
    self.on_operador_select_espessura(None)

def operador_subtrair_qtd(self):
    _ensure_configured()
    enc = self.get_operador_encomenda()
    p = self.get_operador_peca(enc)
    if not enc or not p:
        return
    val = simple_input(self.root, "Ajuste", "Inserir quantidade a subtrair:")
    if val is None:
        return
    try:
        delta = float(val)
    except Exception:
        messagebox.showerror("Erro", "Quantidade inválida")
        return
    if delta <= 0:
        messagebox.showerror("Erro", "Quantidade inválida")
        return
    if not self.operador_ajustar_quantidade(p, -delta):
        return
    update_estado_encomenda_por_espessuras(enc)
    save_data(self.data)
    self.on_operador_select(None)
    self.on_operador_select_espessura(None)

def operador_reabrir_peca(self):
    _ensure_configured()
    enc = self.get_operador_encomenda()
    p = self.get_operador_peca(enc)
    if not enc or not p:
        return
    hist = p.get("hist", [])
    if not hist:
        messagebox.showerror("Erro", "Sem registos para reabrir")
        return
    last = hist.pop()
    delta = float(last.get("delta_ok", 0))
    if not self.operador_ajustar_quantidade(p, -delta):
        return
    update_estado_encomenda_por_espessuras(enc)
    save_data(self.data)
    self.on_operador_select(None)
    self.on_operador_select_espessura(None)

def operador_inicio_peca(self):
    _ensure_configured()
    try:
        if hasattr(self, "op_lote_var") and bool(self.op_lote_var.get()):
            return self.operador_inicio_todas_pecas_espessura()
    except Exception:
        pass
    enc = self.get_operador_encomenda()
    pecas = get_operador_pecas(self, enc, show_error=True)
    if not enc or not pecas:
        return False
    if _op_block_if_selected_avaria_open(self, enc, pecas, "iniciar a peça"):
        return False
    if len(pecas) > 1:
        self.op_sel_peca = pecas[0]
        return self.operador_inicio_todas_pecas_espessura()
    p = pecas[0]
    try:
        if "interromp" in norm_text(p.get("estado", "")):
            _mysql_ops_release_piece(str(enc.get("numero", "") or ""), str(p.get("id", "") or ""))
    except Exception:
        pass
    if not self.op_user.get().strip():
        messagebox.showerror("Erro", "Selecione o operador")
        return False
    pend = peca_operacoes_pendentes(p)
    if not pend:
        messagebox.showinfo("Info", "Esta peça não tem operações pendentes.")
        return False
    sel_ops = self.escolher_operacoes_concluir(p, parent=self.root)
    if sel_ops is None:
        return False
    sel_norm = []
    pend_set = set(pend)
    for op in sel_ops:
        opn = normalize_operacao_nome(op)
        if opn and opn in pend_set and opn not in sel_norm:
            sel_norm.append(opn)
    if not sel_norm:
        messagebox.showerror("Erro", "Selecione pelo menos uma operação pendente.")
        return False
    if len(sel_norm) > 1:
        # Regra de chão de fábrica: 1 operação ativa por ação para medir tempos por operação.
        sel_norm = [sel_norm[0]]
    lock_res = _mysql_ops_acquire(
        enc.get("numero", ""),
        str(p.get("id", "") or ""),
        sel_norm,
        self.op_user.get().strip(),
        valid_operators=_active_operator_names(self),
    )
    acquired = list(lock_res.get("acquired", []) or [])
    blocked = list(lock_res.get("blocked", []) or [])
    if not acquired:
        if blocked:
            ops_txt = ", ".join(f"{b.get('operacao','')} ({_lock_owner_label(b.get('operador',''))})" for b in blocked)
            messagebox.showerror("Ocupada", f"Operação(ões) em uso por outro operador:\n{ops_txt}")
        else:
            messagebox.showerror("Erro", "Não foi possível iniciar as operações selecionadas.")
        return False
    if blocked:
        ops_txt = ", ".join(f"{b.get('operacao','')} ({_lock_owner_label(b.get('operador',''))})" for b in blocked)
        messagebox.showwarning("Parcial", f"Foram iniciadas: {', '.join(acquired)}\nBloqueadas: {ops_txt}")
    if not p.get("inicio_producao"):
        p["inicio_producao"] = now_iso()
    p["interrupcao_peca_motivo"] = ""
    p["interrupcao_peca_ts"] = ""
    p["avaria_ativa"] = False
    p["avaria_motivo"] = ""
    p["avaria_fim_ts"] = ""
    _mark_piece_ops_in_progress(p, acquired, self.op_user.get().strip())
    atualizar_estado_peca(p)
    # Ao iniciar nova operação, a peça deve voltar a "Em producao".
    p["estado"] = "Em producao"
    try:
        log_fn = globals().get("mysql_log_production_event")
        if callable(log_fn):
            for opn in acquired:
                log_fn(
                    evento="START_OP",
                    encomenda_numero=enc.get("numero", ""),
                    peca_id=str(p.get("id", "") or ""),
                    ref_interna=p.get("ref_interna", ""),
                    material=p.get("material", ""),
                    espessura=p.get("espessura", ""),
                    operacao=opn,
                    operador=self.op_user.get().strip(),
                    info=_format_event_info_with_posto(self, "Operacao iniciada no menu Operador"),
                )
    except Exception:
        pass
    save_data(self.data)
    self.op_last_enc = enc.get("numero")
    self.op_last_mat = self.op_material
    self.op_last_esp = self.op_espessura
    self._op_force_runtime_sync = True
    self._op_force_ops_status_refresh = True
    self.refresh_operador()
    self.restore_operador_selection()

def operador_fim_peca(self):
    _ensure_configured()
    try:
        if hasattr(self, "op_lote_var") and bool(self.op_lote_var.get()):
            return self.operador_fim_todas_pecas_espessura()
    except Exception:
        pass
    enc = self.get_operador_encomenda()
    pecas = get_operador_pecas(self, enc, show_error=True)
    if not enc or not pecas:
        return
    if _op_block_if_selected_avaria_open(self, enc, pecas, "concluir a peça"):
        return False
    if len(pecas) > 1 and not bool(getattr(self, "_op_multi_running", False)):
        if not messagebox.askyesno(
            "Fim em seleção",
            f"Foram selecionadas {len(pecas)} peças.\nAbrir registo para cada peça (1 a 1)?",
        ):
            return
        old_ids = set(getattr(self, "op_sel_pecas_ids", set()) or set())
        try:
            self._op_multi_running = True
            total_batch = len(pecas)
            for idx_batch, p_it in enumerate(pecas, start=1):
                self.op_sel_peca = p_it
                self.op_sel_pecas_ids = {str((p_it or {}).get("id", "") or "").strip()}
                self._op_batch_idx = idx_batch
                self._op_batch_total = total_batch
                ok_saved = self.operador_fim_peca()
                if not ok_saved:
                    break
        finally:
            self._op_multi_running = False
            self._op_batch_idx = 0
            self._op_batch_total = 0
            self.op_sel_pecas_ids = old_ids
            try:
                self.on_operador_select_espessura(None)
            except Exception:
                pass
        return True
    p = pecas[0]
    if not self.op_user.get().strip():
        messagebox.showerror("Erro", "Selecione o operador")
        return
    self.op_last_enc = enc.get("numero")
    self.op_last_mat = self.op_material
    self.op_last_esp = self.op_espessura
    qtd_total = float(p.get("quantidade_pedida", 0))
    use_custom = CUSTOM_TK_AVAILABLE and (getattr(self, "op_use_custom", False) or self.op_use_full_custom)
    Win = ctk.CTkToplevel if use_custom else Toplevel
    Frm = ctk.CTkFrame if use_custom else ttk.Frame
    Lbl = ctk.CTkLabel if use_custom else ttk.Label
    Ent = ctk.CTkEntry if use_custom else ttk.Entry
    Btn = ctk.CTkButton if use_custom else ttk.Button
    win = Win(self.root)
    batch_idx = int(getattr(self, "_op_batch_idx", 0) or 0)
    batch_total = int(getattr(self, "_op_batch_total", 0) or 0)
    if batch_idx > 0 and batch_total > 0:
        win.title(f"Registo de Produção ({batch_idx}/{batch_total})")
    else:
        win.title("Registo de Produção")
    try:
        win.transient(self.root)
        win.grab_set()
        if use_custom:
            win.geometry("660x400")
            win.resizable(False, False)
        else:
            win.geometry("620x380")
    except Exception:
        pass
    try:
        win.grid_columnconfigure(1, weight=1)
    except Exception:
        pass
    cur_ok = parse_float(p.get("produzido_ok", 0), 0.0)
    cur_nok = parse_float(p.get("produzido_nok", 0), 0.0)
    cur_qual = parse_float(p.get("produzido_qualidade", 0), 0.0)
    if cur_ok <= 0 and cur_nok <= 0 and cur_qual <= 0:
        cur_ok = qtd_total
    ref_in = str(p.get("ref_interna", "") or "-")
    ref_ex = str(p.get("ref_externa", "") or "-")
    Lbl(
        win,
        text=f"Referência: {ref_in}  |  {ref_ex}",
        font=("Segoe UI", 13, "bold") if use_custom else None,
        text_color="#0f172a" if use_custom else None,
    ).grid(row=0, column=0, columnspan=3, sticky="w", padx=6, pady=(6, 4))
    if batch_idx > 0 and batch_total > 0:
        Lbl(
            win,
            text=f"Peça {batch_idx} de {batch_total}",
            font=("Segoe UI", 11, "bold") if use_custom else None,
            text_color="#475569" if use_custom else None,
        ).grid(row=0, column=2, sticky="e", padx=6, pady=(6, 4))
    Lbl(win, text="Peças Boas").grid(row=1, column=0, sticky="w", padx=6, pady=4)
    ok_var = StringVar(value=fmt_num(cur_ok))
    Ent(win, textvariable=ok_var, width=170 if use_custom else None).grid(row=1, column=1, padx=6, pady=4)
    Lbl(win, text="Peças NOK").grid(row=2, column=0, sticky="w", padx=6, pady=4)
    nok_var = StringVar(value=fmt_num(cur_nok))
    Ent(win, textvariable=nok_var, width=170 if use_custom else None).grid(row=2, column=1, padx=6, pady=4)
    Lbl(win, text="Avaliação Qualidade").grid(row=3, column=0, sticky="w", padx=6, pady=4)
    qual_var = StringVar(value=fmt_num(cur_qual))
    Ent(win, textvariable=qual_var, width=170 if use_custom else None).grid(row=3, column=1, padx=6, pady=4)
    Lbl(win, text="Operações concluídas").grid(row=4, column=0, sticky="w", padx=6, pady=4)
    default_ops = peca_operacoes_pendentes(p)
    ops_var = StringVar(value=" + ".join(default_ops))
    Ent(win, textvariable=ops_var, width=320 if use_custom else 42).grid(row=4, column=1, padx=6, pady=4, sticky="we")
    Btn(
        win,
        text="Selecionar",
        command=lambda: (
            lambda sel_ops: ops_var.set(" + ".join(sel_ops))
            if sel_ops else None
        )(self.escolher_operacoes_concluir(p, parent=win)),
        width=130 if use_custom else None,
    ).grid(row=4, column=2, padx=6, pady=4, sticky="w")
    dur_txt = "-"
    try:
        mins_base = parse_float(p.get("tempo_producao_min", 0), 0.0)
        mins_open = max(0.0, iso_diff_minutes(p.get("inicio_producao"), now_iso()) or 0.0)
        mins = mins_base + mins_open
        if mins > 0:
            h = int(max(0, mins) // 60)
            m = int(max(0, mins) % 60)
            dur_txt = f"{h:02d}:{m:02d} ({mins:.1f} min)"
    except Exception:
        dur_txt = "-"
    Lbl(
        win,
        text=f"Tempo decorrido em produção: {dur_txt}",
        text_color="#334155" if use_custom else None,
    ).grid(row=5, column=0, columnspan=3, sticky="w", padx=6, pady=(0, 2))
    concl_info = ", ".join(peca_operacoes_concluidas(p)) or "-"
    Lbl(
        win,
        text=f"Operações já concluídas: {concl_info}",
        text_color="#5b6f84" if use_custom else None,
    ).grid(row=6, column=0, columnspan=3, sticky="w", padx=6, pady=(0, 2))

    outcome = {"saved": False}

    def on_save():
        try:
            ok = float(ok_var.get() or 0)
            nok = float(nok_var.get() or 0)
            qual = float(qual_var.get() or 0)
        except Exception:
            messagebox.showerror("Erro", "Valores inválidos")
            return
        if ok < 0 or nok < 0 or qual < 0:
            messagebox.showerror("Erro", "Valores inválidos")
            return
        total = ok + nok + qual
        if total > float(p.get("quantidade_pedida", 0)):
            if not messagebox.askyesno("Confirmar", "Ultrapassa a quantidade pedida. Continuar?"):
                return

        # Validacao obrigatoria no momento de guardar.
        selected_ops = self.escolher_operacoes_concluir(p, parent=win)
        if selected_ops is None:
            return
        if selected_ops:
            ops_var.set(" + ".join(selected_ops))
        pend = set(peca_operacoes_pendentes(p))
        if pend:
            filtered_ops = []
            for op in selected_ops:
                nome = normalize_operacao_nome(op)
                if nome and nome in pend and nome not in filtered_ops:
                    filtered_ops.append(nome)
            selected_ops = filtered_ops
            if not selected_ops:
                messagebox.showerror("Erro", "Selecione pelo menos uma operação para concluir.")
                return
        else:
            selected_ops = []

        if len(selected_ops) > 1:
            # Regra: concluir uma operação de cada vez para KPI de tempo por operação.
            selected_ops = [selected_ops[0]]
            ops_var.set(" + ".join(selected_ops))
        lock_finish = _mysql_ops_finish(
            str(enc.get("numero", "") or ""),
            str(p.get("id", "") or ""),
            selected_ops,
            self.op_user.get().strip(),
            ok,
            nok,
            qual,
            valid_operators=_active_operator_names(self),
        )
        blocked_finish = list(lock_finish.get("blocked", []) or [])
        if blocked_finish:
            iniciar_antes = [b for b in blocked_finish if "produ" not in norm_text(b.get("estado", ""))]
            if iniciar_antes:
                ops_txt = ", ".join(str(b.get("operacao", "") or "") for b in iniciar_antes)
                messagebox.showerror(
                    "Operação não iniciada",
                    "Para medir tempo por setor, deve iniciar a operação antes de concluir.\n"
                    f"Operações por iniciar: {ops_txt}",
                )
                return
            ops_txt = ", ".join(f"{b.get('operacao','')} ({_lock_owner_label(b.get('operador',''))})" for b in blocked_finish)
            messagebox.showerror("Ocupada", f"Não pode concluir operações que estão com outro operador:\n{ops_txt}")
            return

        old_state = {
            "produzido_ok": p.get("produzido_ok", 0.0),
            "produzido_nok": p.get("produzido_nok", 0.0),
            "produzido_qualidade": p.get("produzido_qualidade", 0.0),
            "inicio_producao": p.get("inicio_producao", ""),
            "fim_producao": p.get("fim_producao", ""),
            "tempo_producao_min": p.get("tempo_producao_min", 0.0),
            "estado": p.get("estado", "Preparacao"),
            "Operacoes": p.get("Operacoes", p.get("Operações", "")),
            "operacoes_fluxo": [dict(op) for op in p.get("operacoes_fluxo", []) if isinstance(op, dict)],
        }

        ts_fim = now_iso()
        p["produzido_ok"] = ok
        p["produzido_nok"] = nok
        p["produzido_qualidade"] = qual
        concluir_operacoes_peca(p, selected_ops, user=self.op_user.get().strip())
        atualizar_estado_peca(p)
        if any(_is_laser_op_nome(op) for op in list(selected_ops or [])):
            ok_laser = _trigger_laser_completion_actions(self, enc, p.get("material", ""), p.get("espessura", ""))
            if not ok_laser:
                return
        _flush_piece_elapsed_minutes(p, ts_fim)
        if p.get("estado") == "Concluida":
            p["fim_producao"] = ts_fim
        else:
            p["fim_producao"] = ""
            p["estado"] = "Em producao/Pausada"
        # Finalizar espessura apenas quando TODAS as pecas estiverem concluidas
        mat = p.get("material", "")
        esp = p.get("espessura", "")
        esp_obj = self.get_operador_esp_obj(enc, mat, esp)
        if esp_obj and esp_obj.get("pecas") and all(px.get("estado") == "Concluida" for px in esp_obj.get("pecas", [])):
            if messagebox.askyesno("Finalizar", f"Todas as peças de {mat} {esp} estão concluídas. Finalizar?"):
                ok_fin = self.finalizar_espessura(enc, mat, esp, auto=False)
                if not ok_fin:
                    # Se o fluxo for cancelado, reverte a peca para nao concluir indevidamente.
                    p["produzido_ok"] = old_state["produzido_ok"]
                    p["produzido_nok"] = old_state["produzido_nok"]
                    p["produzido_qualidade"] = old_state["produzido_qualidade"]
                    p["inicio_producao"] = old_state["inicio_producao"]
                    p["fim_producao"] = old_state["fim_producao"]
                    p["tempo_producao_min"] = old_state["tempo_producao_min"]
                    p["estado"] = old_state["estado"]
                    p["Operacoes"] = old_state["Operacoes"]
                    p["operacoes_fluxo"] = old_state["operacoes_fluxo"]
                    return
        p.setdefault("hist", []).append({
            "ts": ts_fim,
            "user": self.op_user.get().strip(),
            "acao": "Fim Peça" if p.get("estado") == "Concluida" else "Registo Operação",
            "operacoes": selected_ops,
            "ok": ok,
            "nok": nok,
            "qual": qual,
            "tempo_min": p.get("tempo_producao_min", 0),
        })
        if nok > 0:
            self.data.setdefault("rejeitadas_hist", []).append({
                "data": now_iso(),
                "operador": self.op_user.get().strip(),
                "encomenda": enc.get("numero", ""),
                "material": p.get("material", ""),
                "espessura": p.get("espessura", ""),
                "ref_interna": p.get("ref_interna", ""),
                "ref_externa": p.get("ref_externa", ""),
                "nok": nok,
            })
        try:
            log_fn = globals().get("mysql_log_production_event")
            if callable(log_fn):
                for opn in (selected_ops or []):
                    log_fn(
                        evento="FINISH_OP",
                        encomenda_numero=enc.get("numero", ""),
                        peca_id=str(p.get("id", "") or ""),
                        ref_interna=p.get("ref_interna", ""),
                        material=p.get("material", ""),
                        espessura=p.get("espessura", ""),
                        operacao=opn,
                        operador=self.op_user.get().strip(),
                        qtd_ok=ok,
                        qtd_nok=nok,
                        info=_format_event_info_with_posto(self, "Operacao concluida no menu Operador"),
                    )
                if nok > 0:
                    log_fn(
                        evento="SCRAP",
                        encomenda_numero=enc.get("numero", ""),
                        peca_id=str(p.get("id", "") or ""),
                        ref_interna=p.get("ref_interna", ""),
                        material=p.get("material", ""),
                        espessura=p.get("espessura", ""),
                        operador=self.op_user.get().strip(),
                        qtd_nok=nok,
                        info=_format_event_info_with_posto(self, "Registo NOK no fim de peca"),
                    )
        except Exception:
            pass
        update_estado_encomenda_por_espessuras(enc)
        save_data(self.data)
        outcome["saved"] = True
        win.destroy()
        try:
            if p.get("estado") != "Concluida":
                pend_txt = ", ".join(peca_operacoes_pendentes(p)) or "-"
                messagebox.showinfo("Operações pendentes", f"Registo guardado.\nA peça ficou em pausa.\nPendentes: {pend_txt}")
        except Exception:
            pass
        self._op_force_runtime_sync = True
        self._op_force_ops_status_refresh = True
        self.refresh_operador()
        self.restore_operador_selection()
        try:
            if hasattr(self, "mark_tab_dirty"):
                self.mark_tab_dirty("expedicao", "plano")
        except Exception:
            pass

    def on_cancel():
        outcome["saved"] = False
        win.destroy()

    btns = Frm(win, fg_color="transparent") if use_custom else Frm(win)
    btns.grid(row=7, column=0, columnspan=3, sticky="ew", padx=6, pady=10)
    if use_custom:
        Btn(
            btns,
            text="Guardar",
            command=on_save,
            width=160,
            fg_color="#f59e0b",
            hover_color="#d97706",
            text_color="#ffffff",
        ).pack(side="left", padx=4)
    else:
        Btn(btns, text="Guardar", command=on_save, width=160 if use_custom else None).pack(side="left", padx=4)
    Btn(btns, text="Cancelar", command=on_cancel, width=140 if use_custom else None).pack(side="right", padx=4)
    try:
        self.root.wait_window(win)
    except Exception:
        pass
    return bool(outcome.get("saved", False))

def operador_reabrir_peca_total(self):
    _ensure_configured()
    enc = self.get_operador_encomenda()
    p = self.get_operador_peca(enc)
    if not enc or not p:
        return
    p["produzido_ok"] = 0.0
    p["produzido_nok"] = 0.0
    p["produzido_qualidade"] = 0.0
    p["inicio_producao"] = ""
    p["fim_producao"] = ""
    p["estado"] = "Preparacao"
    p["operacoes_fluxo"] = build_operacoes_fluxo(p.get("Operacoes", p.get("Operações", "")))
    _mysql_ops_reset_piece(str(enc.get("numero", "") or ""), str(p.get("id", "") or ""))
    save_data(self.data)
    self.op_last_enc = enc.get("numero")
    self.op_last_mat = self.op_material
    self.op_last_esp = self.op_espessura
    self.refresh_operador()
    self.restore_operador_selection()


def _selecionar_operacao_lote(self, operacoes, parent=None):
    _ensure_configured()
    ops = [normalize_operacao_nome(x) for x in list(operacoes or []) if normalize_operacao_nome(x)]
    ops = list(dict.fromkeys(ops))
    if not ops:
        return None
    use_custom = CUSTOM_TK_AVAILABLE and (getattr(self, "op_use_custom", False) or self.op_use_full_custom)
    if len(ops) == 1:
        return ops[0]
    if not use_custom:
        txt = "\n".join(f"- {o}" for o in ops)
        val = simple_input(self.root, "Operação em lote", f"Operações disponíveis:\n{txt}\n\nEscreva a operação:")
        if val is None:
            return None
        op = normalize_operacao_nome(val)
        return op if op in ops else None

    host = parent if parent is not None else self.root
    win = ctk.CTkToplevel(host)
    win.title("Iniciar em lote")
    try:
        win.geometry("560x230")
        win.resizable(False, False)
        win.transient(host)
        win.grab_set()
        win.configure(fg_color="#f7f8fb")
    except Exception:
        pass
    frame = ctk.CTkFrame(win, fg_color="#ffffff", corner_radius=10, border_width=1, border_color="#dbe4f0")
    frame.pack(fill="both", expand=True, padx=10, pady=10)
    ctk.CTkLabel(
        frame,
        text="Selecione a operação para iniciar em todas as peças da espessura",
        font=("Segoe UI", 14, "bold"),
        text_color="#0f172a",
    ).pack(anchor="w", padx=12, pady=(12, 8))
    var = StringVar(value=ops[0])
    cb = ctk.CTkComboBox(frame, variable=var, values=ops, width=360)
    cb.pack(anchor="w", padx=12, pady=(0, 10))
    out = {"op": None}

    def _ok():
        op = normalize_operacao_nome(var.get())
        out["op"] = op if op in ops else None
        win.destroy()

    def _cancel():
        out["op"] = None
        win.destroy()

    btns = ctk.CTkFrame(frame, fg_color="transparent")
    btns.pack(fill="x", padx=12, pady=(2, 12))
    ctk.CTkButton(btns, text="Cancelar", command=_cancel, width=140).pack(side="right", padx=(8, 0))
    ctk.CTkButton(
        btns,
        text="Confirmar",
        command=_ok,
        width=160,
        fg_color="#f59e0b",
        hover_color="#d97706",
        text_color="#ffffff",
    ).pack(side="right")
    try:
        host.wait_window(win)
    except Exception:
        pass
    return out["op"]


def operador_inicio_todas_pecas_espessura(self):
    _ensure_configured()
    enc = self.get_operador_encomenda()
    if not enc:
        return
    operador = self.op_user.get().strip() if hasattr(self, "op_user") else ""
    if not operador:
        messagebox.showerror("Erro", "Selecione o operador")
        return
    mat, esp = self._get_operador_mat_esp()
    if not mat or not esp:
        messagebox.showerror("Erro", "Selecione uma espessura")
        return
    esp_obj = self.get_operador_esp_obj(enc, mat, esp)
    if not esp_obj:
        # fallback por normalização de espessura
        esp_key = fmt_num(esp)
        for m in enc.get("materiais", []):
            if str(m.get("material", "")) != str(mat):
                continue
            for e in m.get("espessuras", []):
                if fmt_num(e.get("espessura", "")) == esp_key:
                    esp_obj = e
                    break
            if esp_obj:
                break
    if not esp_obj:
        messagebox.showerror("Erro", "Espessura não encontrada.")
        return
    pecas = list(esp_obj.get("pecas", []) or [])
    sel_ids = set(getattr(self, "op_sel_pecas_ids", set()) or set())
    if sel_ids:
        pecas = [p for p in pecas if str(p.get("id", "") or "").strip() in sel_ids]
    if not pecas:
        messagebox.showinfo("Info", "Sem peças nesta espessura.")
        return
    avaria_index = _op_open_avaria_index(getattr(self, "data", {}), str(enc.get("numero", "") or ""))

    ops_disponiveis = []
    for p in pecas:
        for op in peca_operacoes_pendentes(p):
            opn = normalize_operacao_nome(op)
            if opn and opn not in ops_disponiveis:
                ops_disponiveis.append(opn)
    if not ops_disponiveis:
        messagebox.showinfo("Info", "Não existem operações pendentes nas peças desta espessura.")
        return

    op_sel = _selecionar_operacao_lote(self, ops_disponiveis, parent=self.root)
    if not op_sel:
        return
    if not messagebox.askyesno(
        "Confirmar lote",
        f"Iniciar '{op_sel}' em todas as peças elegíveis de {mat} {esp}?",
    ):
        return

    started_pieces = 0
    blocked_pieces = 0
    skipped_pieces = 0
    blocked_info = []
    ts_now = now_iso()

    for p in pecas:
        if _op_piece_has_open_avaria(getattr(self, "data", {}), str(enc.get("numero", "") or ""), p, avaria_index=avaria_index):
            skipped_pieces += 1
            continue
        pend = set(peca_operacoes_pendentes(p))
        if op_sel not in pend:
            skipped_pieces += 1
            continue
        lock_res = _mysql_ops_acquire(
            enc.get("numero", ""),
            str(p.get("id", "") or ""),
            [op_sel],
            operador,
            valid_operators=_active_operator_names(self),
        )
        acquired = list(lock_res.get("acquired", []) or [])
        blocked = list(lock_res.get("blocked", []) or [])
        if acquired:
            if not p.get("inicio_producao"):
                p["inicio_producao"] = ts_now
            p["interrupcao_peca_motivo"] = ""
            p["interrupcao_peca_ts"] = ""
            p["avaria_ativa"] = False
            p["avaria_motivo"] = ""
            p["avaria_fim_ts"] = ""
            _mark_piece_ops_in_progress(p, acquired, operador)
            atualizar_estado_peca(p)
            started_pieces += 1
            try:
                log_fn = globals().get("mysql_log_production_event")
                if callable(log_fn):
                    log_fn(
                        evento="START_OP",
                        encomenda_numero=enc.get("numero", ""),
                        peca_id=str(p.get("id", "") or ""),
                        ref_interna=p.get("ref_interna", ""),
                        material=p.get("material", ""),
                        espessura=p.get("espessura", ""),
                        operacao=op_sel,
                        operador=operador,
                        info=_format_event_info_with_posto(self, "Operacao iniciada em lote (espessura)"),
                    )
            except Exception:
                pass
        elif blocked:
            blocked_pieces += 1
            b = blocked[0]
            blocked_info.append(
                f"{p.get('ref_interna','-')}: {_lock_owner_label(b.get('operador',''))}"
            )
        else:
            blocked_pieces += 1

    update_estado_encomenda_por_espessuras(enc)
    save_data(self.data)
    try:
        if hasattr(self, "mark_tab_dirty"):
            self.mark_tab_dirty("plano")
    except Exception:
        pass
    self.op_last_enc = enc.get("numero")
    self.op_last_mat = self.op_material
    self.op_last_esp = self.op_espessura
    self._op_force_runtime_sync = True
    self._op_force_ops_status_refresh = True
    self.refresh_operador()
    self.restore_operador_selection()

    msg = (
        f"Operação: {op_sel}\n"
        f"Iniciadas: {started_pieces}\n"
        f"Bloqueadas: {blocked_pieces}\n"
        f"Sem essa operação: {skipped_pieces}"
    )
    if blocked_info:
        msg += "\n\nBloqueadas por:\n" + "\n".join(blocked_info[:8])
        if len(blocked_info) > 8:
            msg += f"\n... +{len(blocked_info)-8}"
    messagebox.showinfo("Início em lote", msg)


def operador_fim_todas_pecas_espessura(self):
    _ensure_configured()
    enc = self.get_operador_encomenda()
    if not enc:
        return
    operador = self.op_user.get().strip() if hasattr(self, "op_user") else ""
    if not operador:
        messagebox.showerror("Erro", "Selecione o operador")
        return
    mat, esp = self._get_operador_mat_esp()
    if not mat or not esp:
        messagebox.showerror("Erro", "Selecione uma espessura")
        return
    esp_obj = self.get_operador_esp_obj(enc, mat, esp)
    if not esp_obj:
        esp_key = fmt_num(esp)
        for m in enc.get("materiais", []):
            if str(m.get("material", "")) != str(mat):
                continue
            for e in m.get("espessuras", []):
                if fmt_num(e.get("espessura", "")) == esp_key:
                    esp_obj = e
                    break
            if esp_obj:
                break
    if not esp_obj:
        messagebox.showerror("Erro", "Espessura não encontrada.")
        return
    pecas = list(esp_obj.get("pecas", []) or [])
    sel_ids = set(getattr(self, "op_sel_pecas_ids", set()) or set())
    if sel_ids:
        pecas = [p for p in pecas if str(p.get("id", "") or "").strip() in sel_ids]
    if not pecas:
        messagebox.showinfo("Info", "Sem peças nesta espessura.")
        return
    avaria_index = _op_open_avaria_index(getattr(self, "data", {}), str(enc.get("numero", "") or ""))

    ops_disponiveis = []
    for p in pecas:
        for op in peca_operacoes_pendentes(p):
            opn = normalize_operacao_nome(op)
            if opn and opn not in ops_disponiveis:
                ops_disponiveis.append(opn)
    if not ops_disponiveis:
        messagebox.showinfo("Info", "Não existem operações pendentes nas peças desta espessura.")
        return
    op_sel = _selecionar_operacao_lote(self, ops_disponiveis, parent=self.root)
    if not op_sel:
        return
    if not messagebox.askyesno(
        "Confirmar lote",
        f"Concluir '{op_sel}' em todas as peças elegíveis de {mat} {esp}?",
    ):
        return

    done_pieces = 0
    blocked_pieces = 0
    skipped_pieces = 0
    blocked_info = []
    ts_now = now_iso()

    for p in pecas:
        if _op_piece_has_open_avaria(getattr(self, "data", {}), str(enc.get("numero", "") or ""), p, avaria_index=avaria_index):
            skipped_pieces += 1
            continue
        pend = set(peca_operacoes_pendentes(p))
        if op_sel not in pend:
            skipped_pieces += 1
            continue
        ok = parse_float(p.get("produzido_ok", 0), 0.0)
        nok = parse_float(p.get("produzido_nok", 0), 0.0)
        qual = parse_float(p.get("produzido_qualidade", 0), 0.0)
        lock_res = _mysql_ops_finish(
            str(enc.get("numero", "") or ""),
            str(p.get("id", "") or ""),
            [op_sel],
            operador,
            ok,
            nok,
            qual,
            valid_operators=_active_operator_names(self),
        )
        done = list(lock_res.get("done", []) or [])
        blocked = list(lock_res.get("blocked", []) or [])
        if done:
            concluir_operacoes_peca(p, [op_sel], user=operador)
            atualizar_estado_peca(p)
            _flush_piece_elapsed_minutes(p, ts_now)
            if p.get("estado") == "Concluida":
                p["fim_producao"] = ts_now
            else:
                p["fim_producao"] = ""
                p["estado"] = "Em producao/Pausada"
            p.setdefault("hist", []).append(
                {
                    "ts": ts_now,
                    "user": operador,
                    "acao": "Fim Operação (Lote)",
                    "operacoes": [op_sel],
                    "ok": ok,
                    "nok": nok,
                    "qual": qual,
                }
            )
            done_pieces += 1
            try:
                log_fn = globals().get("mysql_log_production_event")
                if callable(log_fn):
                    log_fn(
                        evento="FINISH_OP",
                        encomenda_numero=enc.get("numero", ""),
                        peca_id=str(p.get("id", "") or ""),
                        ref_interna=p.get("ref_interna", ""),
                        material=p.get("material", ""),
                        espessura=p.get("espessura", ""),
                        operacao=op_sel,
                        operador=operador,
                        qtd_ok=ok,
                        qtd_nok=nok,
                        info=_format_event_info_with_posto(self, "Operacao concluida em lote (espessura)"),
                    )
            except Exception:
                pass
        elif blocked:
            blocked_pieces += 1
            b = blocked[0]
            blocked_info.append(
                f"{p.get('ref_interna','-')}: {_lock_owner_label(b.get('operador',''))}"
            )
        else:
            blocked_pieces += 1

    if done_pieces > 0 and _is_laser_op_nome(op_sel):
        ok_laser = _trigger_laser_completion_actions(self, enc, mat, esp)
        if not ok_laser:
            return False

    update_estado_encomenda_por_espessuras(enc)
    save_data(self.data)
    self.op_last_enc = enc.get("numero")
    self.op_last_mat = self.op_material
    self.op_last_esp = self.op_espessura
    self._op_force_runtime_sync = True
    self._op_force_ops_status_refresh = True
    self.refresh_operador()
    self.restore_operador_selection()

    msg = (
        f"Operação: {op_sel}\n"
        f"Concluídas: {done_pieces}\n"
        f"Bloqueadas: {blocked_pieces}\n"
        f"Sem essa operação: {skipped_pieces}"
    )
    if blocked_info:
        msg += "\n\nBloqueadas por:\n" + "\n".join(blocked_info[:8])
        if len(blocked_info) > 8:
            msg += f"\n... +{len(blocked_info)-8}"
    messagebox.showinfo("Fim em lote", msg)
    return True

def operador_retomar_peca(self):
    _ensure_configured()
    enc = self.get_operador_encomenda()
    pecas = get_operador_pecas(self, enc, show_error=True)
    if not enc or not pecas:
        return
    if _op_block_if_selected_avaria_open(self, enc, pecas, "retomar a peça"):
        return False
    operador = (self.op_user.get().strip() if hasattr(self, "op_user") else "").strip()
    resumed = 0
    for p in pecas:
        est = norm_text(p.get("estado", ""))
        if ("interromp" not in est) and ("incomplet" not in est):
            continue
        try:
            _mysql_ops_release_piece(str(enc.get("numero", "") or ""), str(p.get("id", "") or ""))
        except Exception:
            pass
        motivo = str(p.get("interrupcao_peca_motivo", "") or "").strip()
        p.setdefault("hist", []).append(
            {
                "ts": now_iso(),
                "user": operador,
                "acao": "Retomar Peça",
                "motivo": motivo,
            }
        )
        p["interrupcao_peca_motivo"] = ""
        p["interrupcao_peca_ts"] = ""
        # Retomar não inicia produção automaticamente.
        # A peça fica pronta para novo "Iniciar Peça".
        p["estado"] = "Em producao/Pausada"
        resumed += 1
        try:
            log_fn = globals().get("mysql_log_production_event")
            if callable(log_fn):
                log_fn(
                    evento="RESUME_PIECE",
                    encomenda_numero=enc.get("numero", ""),
                    peca_id=str(p.get("id", "") or ""),
                    ref_interna=p.get("ref_interna", ""),
                    material=p.get("material", ""),
                    espessura=p.get("espessura", ""),
                    operador=operador,
                    info=_format_event_info_with_posto(self, f"Retoma de peça. Motivo anterior: {motivo or '-'}"),
                )
        except Exception:
            pass
    if resumed == 0:
        messagebox.showinfo("Info", "Nenhuma peça interrompida/incompleta na seleção.")
        return
    update_estado_encomenda_por_espessuras(enc)
    save_data(self.data)
    self.refresh_operador()
    self.restore_operador_selection()
    messagebox.showinfo("Retomar", "Peça(s) retomada(s). Agora use 'Iniciar Peça' para voltar a contar tempo.")
    return True

def operador_interromper_peca(self):
    _ensure_configured()
    try:
        if hasattr(self, "op_lote_var") and bool(self.op_lote_var.get()):
            return self.operador_interromper_todas_pecas_espessura()
    except Exception:
        pass
    enc = self.get_operador_encomenda()
    pecas = get_operador_pecas(self, enc, show_error=True)
    if not enc or not pecas:
        return
    if _op_block_if_selected_avaria_open(self, enc, pecas, "interromper a peça"):
        return False
    operador = (self.op_user.get().strip() if hasattr(self, "op_user") else "").strip()
    if not operador:
        messagebox.showerror("Erro", "Selecione o operador")
        return
    motivo = _prompt_interrupcao_operacional(self)
    if motivo is None:
        return
    ts_pause = now_iso()
    for p in pecas:
        p["estado"] = "Interrompida"
        p["interrupcao_peca_motivo"] = motivo
        p["interrupcao_peca_ts"] = ts_pause
        _flush_piece_elapsed_minutes(p, ts_pause)
        try:
            fluxo = ensure_peca_operacoes(p)
            for op in fluxo:
                if "concl" not in norm_text(op.get("estado", "")):
                    op["estado"] = "Preparacao"
                    op["fim"] = ""
            p["operacoes_fluxo"] = fluxo
        except Exception:
            pass
        p.setdefault("hist", []).append(
            {
                "ts": ts_pause,
                "user": operador,
                "acao": "Interromper Peça",
                "motivo": motivo,
            }
        )
        try:
            _mysql_ops_release_piece(str(enc.get("numero", "") or ""), str(p.get("id", "") or ""))
        except Exception:
            pass
        try:
            log_fn = globals().get("mysql_log_production_event")
            if callable(log_fn):
                log_fn(
                    evento="PAUSE_PIECE",
                    encomenda_numero=enc.get("numero", ""),
                    peca_id=str(p.get("id", "") or ""),
                    ref_interna=p.get("ref_interna", ""),
                    material=p.get("material", ""),
                    espessura=p.get("espessura", ""),
                    operador=operador,
                    info=_format_event_info_with_posto(self, f"Interrupcao da peca. Motivo: {motivo}"),
                )
        except Exception:
            pass
    update_estado_encomenda_por_espessuras(enc)
    save_data(self.data)
    self.op_last_enc = enc.get("numero")
    self.op_last_mat = self.op_material
    self.op_last_esp = self.op_espessura
    self._op_force_runtime_sync = True
    self._op_force_ops_status_refresh = True
    self.refresh_operador()
    self.restore_operador_selection()


def operador_interromper_todas_pecas_espessura(self):
    _ensure_configured()
    enc = self.get_operador_encomenda()
    if not enc:
        return
    operador = (self.op_user.get().strip() if hasattr(self, "op_user") else "").strip()
    if not operador:
        messagebox.showerror("Erro", "Selecione o operador")
        return
    motivo = _prompt_interrupcao_operacional(self)
    if motivo is None:
        return
    mat, esp = self._get_operador_mat_esp()
    if not mat or not esp:
        messagebox.showerror("Erro", "Selecione uma espessura")
        return
    ts_pause = now_iso()

    esp_obj = self.get_operador_esp_obj(enc, mat, esp)
    if not esp_obj:
        esp_key = fmt_num(esp)
        for m in enc.get("materiais", []):
            if str(m.get("material", "")) != str(mat):
                continue
            for e in m.get("espessuras", []):
                if fmt_num(e.get("espessura", "")) == esp_key:
                    esp_obj = e
                    break
            if esp_obj:
                break
    if not esp_obj:
        messagebox.showerror("Erro", "Espessura não encontrada.")
        return

    afetadas = 0
    pecas = list(esp_obj.get("pecas", []) or [])
    sel_ids = set(getattr(self, "op_sel_pecas_ids", set()) or set())
    if sel_ids:
        pecas = [p for p in pecas if str(p.get("id", "") or "").strip() in sel_ids]
    avaria_index = _op_open_avaria_index(getattr(self, "data", {}), str(enc.get("numero", "") or ""))
    for p in pecas:
        if _op_piece_has_open_avaria(getattr(self, "data", {}), str(enc.get("numero", "") or ""), p, avaria_index=avaria_index):
            continue
        est = norm_text(p.get("estado", ""))
        if "conclu" in est:
            continue
        p["estado"] = "Interrompida"
        p["interrupcao_peca_motivo"] = motivo
        p["interrupcao_peca_ts"] = ts_pause
        _flush_piece_elapsed_minutes(p, ts_pause)
        try:
            fluxo = ensure_peca_operacoes(p)
            for op in fluxo:
                if "concl" not in norm_text(op.get("estado", "")):
                    op["estado"] = "Preparacao"
                    op["fim"] = ""
            p["operacoes_fluxo"] = fluxo
        except Exception:
            pass
        p.setdefault("hist", []).append(
            {
                "ts": ts_pause,
                "user": operador,
                "acao": "Interromper Peça (Lote)",
                "motivo": motivo,
            }
        )
        try:
            _mysql_ops_release_piece(str(enc.get("numero", "") or ""), str(p.get("id", "") or ""))
        except Exception:
            pass
        try:
            log_fn = globals().get("mysql_log_production_event")
            if callable(log_fn):
                log_fn(
                    evento="PAUSE_PIECE",
                    encomenda_numero=enc.get("numero", ""),
                    peca_id=str(p.get("id", "") or ""),
                    ref_interna=p.get("ref_interna", ""),
                    material=p.get("material", ""),
                    espessura=p.get("espessura", ""),
                    operador=operador,
                    info=_format_event_info_with_posto(self, f"Interrupcao em lote. Motivo: {motivo}"),
                )
        except Exception:
            pass
        afetadas += 1

    update_estado_encomenda_por_espessuras(enc)
    save_data(self.data)
    self.op_last_enc = enc.get("numero")
    self.op_last_mat = self.op_material
    self.op_last_esp = self.op_espessura
    self._op_force_runtime_sync = True
    self._op_force_ops_status_refresh = True
    self.refresh_operador()
    self.restore_operador_selection()
    messagebox.showinfo("Interromper em lote", f"Peças interrompidas: {afetadas}")


def operador_registar_avaria(self):
    _ensure_configured()
    enc = self.get_operador_encomenda()
    pecas = get_operador_pecas(self, enc, show_error=True)
    if not enc or not pecas:
        return
    if _op_block_if_selected_avaria_open(self, enc, pecas, "registar nova avaria"):
        return False
    operador = (self.op_user.get().strip() if hasattr(self, "op_user") else "").strip()
    if not operador:
        messagebox.showerror("Erro", "Selecione o operador")
        return
    motivo = _custom_prompt_text(
        self,
        "Registar Avaria",
        "Selecione o motivo da avaria:",
        options=_metalurgica_paragem_options(),
    )
    if motivo is None:
        return
    motivo = (str(motivo or "").strip() or "Avaria não especificada")
    ts_now = now_iso()
    group_id = _op_make_avaria_group_id(str(enc.get("numero", "") or ""), motivo, operador, ts_now=ts_now)
    total = 0
    for p in pecas:
        if bool(p.get("avaria_ativa")):
            continue
        p["estado"] = "Avaria"
        p["avaria_ativa"] = True
        p["avaria_motivo"] = motivo
        p["avaria_grupo_id"] = group_id
        p["avaria_inicio_ts"] = ts_now
        p["avaria_fim_ts"] = ""
        p["interrupcao_peca_motivo"] = motivo
        p["interrupcao_peca_ts"] = ts_now
        _flush_piece_elapsed_minutes(p, ts_now)
        try:
            fluxo = ensure_peca_operacoes(p)
            for op in fluxo:
                if "concl" not in norm_text(op.get("estado", "")):
                    op["estado"] = "Preparacao"
                    op["fim"] = ""
            p["operacoes_fluxo"] = fluxo
        except Exception:
            pass
        p.setdefault("hist", []).append(
            {
                "ts": ts_now,
                "user": operador,
                "acao": "Registar Avaria",
                "motivo": motivo,
            }
        )
        try:
            _mysql_ops_release_piece(str(enc.get("numero", "") or ""), str(p.get("id", "") or ""))
        except Exception:
            pass
        total += 1
        try:
            log_fn = globals().get("mysql_log_production_event")
            if callable(log_fn):
                log_fn(
                    evento="PARAGEM",
                    encomenda_numero=enc.get("numero", ""),
                    peca_id=str(p.get("id", "") or ""),
                    ref_interna=p.get("ref_interna", ""),
                    material=p.get("material", ""),
                    espessura=p.get("espessura", ""),
                    operador=operador,
                    causa_paragem=motivo,
                    info=_format_event_info_with_posto(self, "Registo de avaria no menu Operador"),
                    created_at=ts_now,
                    grupo_id=group_id,
                )
        except Exception:
            pass
    update_estado_encomenda_por_espessuras(enc)
    save_data(self.data)
    self._op_force_runtime_sync = True
    self._op_force_ops_status_refresh = True
    self.refresh_operador()
    self.restore_operador_selection()
    messagebox.showinfo("Avaria", f"Avaria registada em {total} peça(s).")


def operador_fechar_avaria(self):
    _ensure_configured()
    enc = self.get_operador_encomenda()
    if not enc:
        return
    operador = (self.op_user.get().strip() if hasattr(self, "op_user") else "").strip()
    if not operador:
        messagebox.showerror("Erro", "Selecione o operador")
        return
    ts_now = now_iso()
    closed = 0
    closed_group_minutes = {}
    avaria_index = _op_open_avaria_index(getattr(self, "data", {}), str(enc.get("numero", "") or ""))
    selected_pecas = get_operador_pecas(self, enc, show_error=False)
    target_pecas = []
    seen = set()

    def _push_targets(candidates):
        for p in list(candidates or []):
            if not isinstance(p, dict):
                continue
            live_row = _op_live_avaria_row_for_piece(avaria_index, p)
            if live_row:
                _op_sync_piece_live_avaria(p, live_row)
            if not bool(p.get("avaria_ativa")) and not live_row:
                continue
            key = str(p.get("id", "") or p.get("ref_interna", "") or "").strip()
            if not key or key in seen:
                continue
            seen.add(key)
            target_pecas.append(p)

    _push_targets(selected_pecas)
    if not target_pecas:
        _push_targets(_op_visible_pecas(self, enc))
    if not target_pecas:
        messagebox.showinfo("Avaria", "Nenhuma avaria aberta na seleção atual.")
        return

    for p in target_pecas:
        live_row = _op_live_avaria_row_for_piece(avaria_index, p)
        if live_row:
            _op_sync_piece_live_avaria(p, live_row)
        dur_segment = _op_piece_current_avaria_minutes(
            getattr(self, "data", {}),
            str(enc.get("numero", "") or ""),
            p,
            live_row=live_row,
            ts_ref=ts_now,
        )
        motivo = str(p.get("avaria_motivo", "") or p.get("interrupcao_peca_motivo", "") or "").strip() or "Avaria não especificada"
        stored_group_key = str(p.get("avaria_grupo_id", "") or "").strip()
        p["avaria_ativa"] = False
        p["avaria_fim_ts"] = ts_now
        p["interrupcao_peca_motivo"] = ""
        p["interrupcao_peca_ts"] = ""
        p["avaria_motivo"] = ""
        p["avaria_grupo_id"] = ""
        atualizar_estado_peca(p)
        p.setdefault("hist", []).append(
            {
                "ts": ts_now,
                "user": operador,
                "acao": "Fechar Avaria",
                "motivo": motivo,
                "inicio": str(p.get("avaria_inicio_ts", "") or ""),
                "duracao_min": round(dur_segment, 2),
            }
        )
        try:
            log_fn = globals().get("mysql_log_production_event")
            if callable(log_fn):
                log_fn(
                    evento="CLOSE_AVARIA",
                    encomenda_numero=enc.get("numero", ""),
                    peca_id=str(p.get("id", "") or ""),
                    ref_interna=p.get("ref_interna", ""),
                    material=p.get("material", ""),
                    espessura=p.get("espessura", ""),
                    operador=operador,
                    causa_paragem=motivo,
                    info=_format_event_info_with_posto(self, f"Avaria fechada no menu Operador. Motivo: {motivo}"),
                )
        except Exception:
            pass
        closed += 1
        group_key = _op_avaria_group_key(live_row or {})
        if not group_key:
            group_key = stored_group_key or _op_piece_lookup_key(p)
        closed_group_minutes[group_key] = max(closed_group_minutes.get(group_key, 0.0), max(0.0, dur_segment))
    if closed <= 0:
        messagebox.showinfo("Avaria", "Nenhuma avaria aberta nas peças selecionadas.")
        return
    update_estado_encomenda_por_espessuras(enc)
    save_data(self.data)
    self._op_force_runtime_sync = True
    self._op_force_ops_status_refresh = True
    self.refresh_operador()
    self.restore_operador_selection()
    total_paragem = round(sum(closed_group_minutes.values()), 2)
    messagebox.showinfo(
        "Avaria",
        f"Avaria fechada em {closed} peça(s).\nTempo total de paragem: {fmt_num(total_paragem)} min",
    )


def operador_alertar_chefia(self):
    _ensure_configured()
    enc = self.get_operador_encomenda()
    if not enc:
        return False
    operador = (self.op_user.get().strip() if hasattr(self, "op_user") else "").strip()
    if not operador:
        messagebox.showerror("Erro", "Selecione o operador")
        return False
    enc_num = str(enc.get("numero", "") or "").strip()
    pecas_sel = get_operador_pecas(self, enc, show_error=False)
    p = pecas_sel[0] if pecas_sel else None
    posto = _current_operator_posto(self) or "Geral"
    peca_id = str((p or {}).get("id", "") or "").strip()
    ref_interna = str((p or {}).get("ref_interna", "") or "").strip()
    material = str((p or {}).get("material", "") or "").strip()
    espessura = str((p or {}).get("espessura", "") or "").strip()
    detalhe = f"Chefia solicitada para deslocacao imediata ao colaborador no posto {posto}."
    alert_local = {
        "created_at": now_iso(),
        "tipo": "POKE_CHEFIA",
        "encomenda_numero": enc_num,
        "peca_id": peca_id,
        "ref_interna": ref_interna,
        "material": material,
        "espessura": espessura,
        "operador": operador,
        "posto": posto,
        "mensagem": detalhe,
    }
    try:
        log_fn = globals().get("mysql_log_production_event")
        if callable(log_fn):
            log_fn(
                evento="POKE_CHEFIA",
                encomenda_numero=enc_num,
                peca_id=peca_id,
                ref_interna=ref_interna,
                material=material,
                espessura=espessura,
                operador=operador,
                info=_format_event_info_with_posto(self, detalhe),
            )
    except Exception:
        pass
    try:
        self.data.setdefault("chefia_alertas", []).append(alert_local)
        self.data["chefia_alertas"] = list(self.data.get("chefia_alertas", []) or [])[-200:]
        save_data(self.data)
    except Exception:
        pass
    msg = (
        "Poke enviado para a chefia.\n"
        f"Encomenda: {enc_num or '-'}\n"
        f"Operador: {operador or '-'}\n"
        f"Posto: {posto or '-'}"
    )
    messagebox.showinfo("Alertar Chefia", msg)
    return True


def operador_ver_desenho(self):
    _ensure_configured()
    enc = self.get_operador_encomenda()
    pecas = get_operador_pecas(self, enc, show_error=True)
    if not enc or not pecas:
        return False
    p = pecas[0]
    ref_int = str(p.get("ref_interna", "") or "").strip()
    ref_ext = str(p.get("ref_externa", "") or "").strip()
    try:
        ok = bool(self.open_peca_desenho_by_refs(enc=enc, ref_interna=ref_int, ref_externa=ref_ext, silent=False))
    except Exception:
        ok = False
    if not ok:
        messagebox.showinfo("Ver desenho", "Não foi possível abrir o desenho da peça selecionada.")
    return ok


def _get_operador_mat_esp(self):
    _ensure_configured()
    if self.op_use_full_custom:
        mat = (self.op_material or "").strip()
        esp = str(self.op_espessura).strip() if self.op_espessura is not None else ""
        if (not mat or not esp) and hasattr(self, "op_esp_var"):
            val = (self.op_esp_var.get() or "").strip()
            if "|" in val:
                m, e = val.split("|", 1)
                mat = mat or m.strip()
                esp = esp or e.replace("mm", "").strip()
        return mat, esp
    sel = self.op_tbl_esp.selection()
    if not sel:
        return "", ""
    return self.op_tbl_esp.item(sel[0], "values")[0], self.op_tbl_esp.item(sel[0], "values")[1]

def operador_iniciar_espessura(self):
    _ensure_configured()
    enc = self.get_operador_encomenda()
    if not enc:
        return
    mat, esp = self._get_operador_mat_esp()
    if not mat or not esp:
        messagebox.showerror("Erro", "Selecione uma espessura")
        return
    esp_key = fmt_num(esp)
    # Apenas uma espessura em producao
    for m in enc.get("materiais", []):
        for e in m.get("espessuras", []):
            if e.get("estado") == "Em producao" and not (m.get("material") == mat and fmt_num(e.get("espessura")) == esp_key):
                messagebox.showerror("Erro", "Já existe uma espessura em produção.")
                return
    found = False
    ts_inicio = now_iso()
    for m in enc.get("materiais", []):
        if m.get("material") != mat:
            continue
        for e in m.get("espessuras", []):
            if fmt_num(e.get("espessura")) == esp_key:
                if not e.get("inicio_producao"):
                    e["inicio_producao"] = ts_inicio
                e["estado"] = "Em producao"
                for p in e.get("pecas", []):
                    if p.get("estado") == "Preparacao":
                        p["estado"] = "Em producao"
                    if not p.get("inicio_producao"):
                        p["inicio_producao"] = ts_inicio
                found = True
                break
    if not found:
        messagebox.showerror("Erro", "Espessura não encontrada na encomenda.")
        return
    if not enc.get("inicio_producao"):
        enc["inicio_producao"] = ts_inicio
    update_estado_encomenda_por_espessuras(enc)
    save_data(self.data)
    try:
        if hasattr(self, "mark_tab_dirty"):
            self.mark_tab_dirty("encomendas")
    except Exception:
        pass
    try:
        self.refresh_operador()
    except Exception:
        pass
    self.on_operador_select(None)
    self.on_operador_select_espessura(None)
    messagebox.showinfo("OK", f"Espessura {esp} iniciada")

def operador_finalizar_espessura(self):
    _ensure_configured()
    enc = self.get_operador_encomenda()
    if not enc:
        return
    mat, esp = self._get_operador_mat_esp()
    if not mat or not esp:
        messagebox.showerror("Erro", "Selecione uma espessura")
        return
    if not self.finalizar_espessura(enc, mat, esp, auto=False):
        return
    self.on_operador_select(None)
    messagebox.showinfo("OK", f"Espessura {esp} finalizada")

def operador_iniciar_encomenda(self):
    _ensure_configured()
    enc = self.get_operador_encomenda()
    if not enc:
        return
    obs = simple_input(self.root, "Início Encomenda", "Observação:")
    enc["inicio_encomenda"] = now_iso()
    enc["estado_operador"] = "Em curso"
    enc["obs_inicio"] = obs or ""
    try:
        log_fn = globals().get("mysql_log_production_event")
        if callable(log_fn):
            log_fn(
                evento="START_ENC",
                encomenda_numero=enc.get("numero", ""),
                operador=self.op_user.get().strip() if hasattr(self, "op_user") else "",
                info=_format_event_info_with_posto(self, f"Inicio encomenda. Obs: {obs or ''}"),
            )
    except Exception:
        pass
    save_data(self.data)
    self.refresh_operador()

def operador_interromper(self):
    _ensure_configured()
    enc = self.get_operador_encomenda()
    if not enc:
        return
    motivo = _custom_prompt_text(
        self,
        "Interromper",
        "Selecione a causa de paragem:",
        options=_metalurgica_paragem_options(),
    )
    enc["estado_operador"] = "Interrompida"
    enc["obs_interrupcao"] = motivo or ""
    try:
        log_fn = globals().get("mysql_log_production_event")
        if callable(log_fn):
            log_fn(
                evento="PARAGEM",
                encomenda_numero=enc.get("numero", ""),
                operador=self.op_user.get().strip() if hasattr(self, "op_user") else "",
                causa_paragem=(motivo or "Sem motivo"),
                info=_format_event_info_with_posto(self, f"Interrupcao de encomenda: {motivo or 'Sem motivo'}"),
            )
    except Exception:
        pass
    save_data(self.data)
    self.refresh_operador()

def operador_finalizar_encomenda(self):
    _ensure_configured()
    enc = self.get_operador_encomenda()
    if not enc:
        return
    enc["fim_encomenda"] = now_iso()
    enc["estado_operador"] = "Concluida"
    try:
        log_fn = globals().get("mysql_log_production_event")
        if callable(log_fn):
            log_fn(
                evento="FINISH_ENC",
                encomenda_numero=enc.get("numero", ""),
                operador=self.op_user.get().strip() if hasattr(self, "op_user") else "",
                info=_format_event_info_with_posto(self, "Encomenda finalizada no menu Operador"),
            )
    except Exception:
        pass
    save_data(self.data)
    self.refresh_operador()

def gerir_operadores(self):
    _ensure_configured()
    if hasattr(self, "op_manage_win") and self.op_manage_win and self.op_manage_win.winfo_exists():
        try:
            self.op_manage_win.deiconify()
            self.op_manage_win.lift()
            self.op_manage_win.focus_force()
            self.op_manage_win.attributes("-topmost", True)
            self.op_manage_win.after(
                250,
                lambda: self.op_manage_win
                and self.op_manage_win.winfo_exists()
                and self.op_manage_win.attributes("-topmost", False),
            )
        except Exception:
            pass
        return

    use_custom = CUSTOM_TK_AVAILABLE and os.environ.get("USE_CUSTOM_OP", "1") != "0"
    win = ctk.CTkToplevel(self.root) if use_custom else Toplevel(self.root)
    self.op_manage_win = win
    win.title("Operadores")
    try:
        win.transient(self.root)
    except Exception:
        pass
    try:
        win.lift()
        win.focus_force()
        win.attributes("-topmost", True)
        win.after(250, lambda: win.winfo_exists() and win.attributes("-topmost", False))
    except Exception:
        pass

    def on_close():
        try:
            win.destroy()
        finally:
            self.op_manage_win = None

    win.protocol("WM_DELETE_WINDOW", on_close)
    try:
        if not isinstance(self.data.get("operador_posto_map", None), dict):
            self.data["operador_posto_map"] = {}
        if not isinstance(self.data.get("postos_trabalho", None), list) or not self.data.get("postos_trabalho"):
            self.data["postos_trabalho"] = ["Geral", "Laser 1", "Quinagem 1", "Soldadura 1", "Embalamento 1"]
        operator_accounts = [
            str(u.get("username", "") or "").strip()
            for u in list(self.data.get("users", []) or [])
            if str(u.get("role", "") or "").strip().lower() == "operador" and str(u.get("username", "") or "").strip()
        ]
        existing_ops = [str(o).strip() for o in list(self.data.get("operadores", []) or []) if str(o).strip()]
        merged_ops = list(dict.fromkeys(existing_ops + operator_accounts))
        self.data["operadores"] = merged_ops
    except Exception:
        pass

    if use_custom:
        try:
            win.geometry("860x560")
        except Exception:
            pass
        root_wrap = ctk.CTkFrame(win, fg_color="#f7fafc")
        root_wrap.pack(fill="both", expand=True, padx=10, pady=10)
        title = ctk.CTkFrame(root_wrap, fg_color="white", corner_radius=10, border_width=1, border_color="#d8e1ec")
        title.pack(fill="x", pady=(0, 8))
        ctk.CTkLabel(title, text="Operadores e Postos Atribuídos", font=("Segoe UI", 20, "bold"), text_color="#0f172a").pack(side="left", padx=14, pady=12)
    else:
        root_wrap = ttk.Frame(win)
        root_wrap.pack(fill="both", expand=True, padx=10, pady=10)

    grid_wrap = ctk.CTkFrame(root_wrap, fg_color="white", corner_radius=10, border_width=1, border_color="#d8e1ec") if use_custom else ttk.Frame(root_wrap)
    grid_wrap.pack(fill="both", expand=True)

    cols = ("operador", "posto")
    tree_style = "OperadorManage.Treeview"
    try:
        style = ttk.Style()
        style.configure(
            tree_style,
            font=("Segoe UI", 11),
            rowheight=30,
            background="#f8fbff",
            fieldbackground="#f8fbff",
            borderwidth=0,
        )
        style.configure(
            f"{tree_style}.Heading",
            font=("Segoe UI", 11, "bold"),
            background=THEME_HEADER_BG,
            foreground="white",
            relief="flat",
        )
        style.map(f"{tree_style}.Heading", background=[("active", THEME_HEADER_ACTIVE)])
        style.map(tree_style, background=[("selected", THEME_SELECT_BG)], foreground=[("selected", THEME_SELECT_FG)])
    except Exception:
        pass

    tree = ttk.Treeview(grid_wrap, columns=cols, show="headings", style=tree_style)
    tree.heading("operador", text="Operador")
    tree.heading("posto", text="Posto atribuído")
    tree.column("operador", width=260, anchor="w")
    tree.column("posto", width=360, anchor="w")
    tree.tag_configure("even", background="#f8fbff")
    tree.tag_configure("odd", background="#eef4fb")
    sy = ttk.Scrollbar(grid_wrap, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=sy.set)
    tree.grid(row=0, column=0, sticky="nsew", padx=(10, 0), pady=(10, 6))
    sy.grid(row=0, column=1, sticky="ns", padx=(0, 10), pady=(10, 6))
    grid_wrap.grid_rowconfigure(0, weight=1)
    grid_wrap.grid_columnconfigure(0, weight=1)

    form = ctk.CTkFrame(root_wrap, fg_color="white", corner_radius=10, border_width=1, border_color="#d8e1ec") if use_custom else ttk.Frame(root_wrap)
    form.pack(fill="x", pady=(8, 0))

    op_nome_var = StringVar()
    op_posto_var = StringVar()

    def _refresh_rows(select_name=None):
        try:
            rows = [str(o).strip() for o in list(self.data.get("operadores", []) or []) if str(o).strip()]
        except Exception:
            rows = []
        current_map = dict(self.data.get("operador_posto_map", {}) or {})
        for iid in tree.get_children():
            tree.delete(iid)
        target = None
        for idx, nome in enumerate(rows):
            posto = str(current_map.get(nome, "") or "Geral")
            iid = tree.insert("", "end", values=(nome, posto), tags=("even" if idx % 2 == 0 else "odd",))
            if select_name and nome == select_name:
                target = iid
        if target:
            tree.selection_set(target)
            tree.focus(target)
            tree.see(target)

    def _load_selected(_event=None):
        sel = tree.selection()
        if not sel:
            return
        vals = tree.item(sel[0], "values") or []
        op_nome_var.set(str(vals[0]) if len(vals) > 0 else "")
        op_posto_var.set(str(vals[1]) if len(vals) > 1 else "Geral")

    tree.bind("<<TreeviewSelect>>", _load_selected)

    if use_custom:
        ctk.CTkLabel(form, text="Operador", font=("Segoe UI", 12, "bold")).grid(row=0, column=0, sticky="w", padx=12, pady=(10, 4))
        ctk.CTkEntry(form, textvariable=op_nome_var, width=240).grid(row=1, column=0, sticky="w", padx=12, pady=(0, 10))
        ctk.CTkLabel(form, text="Posto atribuído", font=("Segoe UI", 12, "bold")).grid(row=0, column=1, sticky="w", padx=12, pady=(10, 4))
        posto_cb = ctk.CTkComboBox(form, variable=op_posto_var, values=list(self.data.get("postos_trabalho", []) or ["Geral"]), width=240)
        posto_cb.grid(row=1, column=1, sticky="w", padx=12, pady=(0, 10))
        btn_row = ctk.CTkFrame(form, fg_color="transparent")
        btn_row.grid(row=1, column=2, sticky="e", padx=12, pady=(0, 10))
        form.grid_columnconfigure(2, weight=1)
    else:
        ttk.Label(form, text="Operador").grid(row=0, column=0, sticky="w", padx=12, pady=(10, 4))
        ttk.Entry(form, textvariable=op_nome_var, width=30).grid(row=1, column=0, sticky="w", padx=12, pady=(0, 10))
        ttk.Label(form, text="Posto atribuído").grid(row=0, column=1, sticky="w", padx=12, pady=(10, 4))
        posto_cb = ttk.Combobox(form, textvariable=op_posto_var, values=list(self.data.get("postos_trabalho", []) or ["Geral"]), width=28, state="readonly")
        posto_cb.grid(row=1, column=1, sticky="w", padx=12, pady=(0, 10))
        btn_row = ttk.Frame(form)
        btn_row.grid(row=1, column=2, sticky="e", padx=12, pady=(0, 10))
        form.grid_columnconfigure(2, weight=1)

    def _sync_operator_values():
        values = list(self.data.get("operadores", []) or [])
        try:
            if getattr(self, "op_user_cb", None) is not None:
                self.op_user_cb.configure(values=values)
        except Exception:
            pass
        try:
            current_user = str(self.op_user.get() or "").strip()
            if current_user:
                _on_operador_user_change(self)
        except Exception:
            pass

    def add_op():
        nome = str(op_nome_var.get() or "").strip()
        posto = str(op_posto_var.get() or "").strip() or "Geral"
        if not nome:
            messagebox.showerror("Erro", "Indique o nome/utilizador do operador.")
            return
        if nome in list(self.data.get("operadores", []) or []):
            messagebox.showerror("Erro", "Operador já existe.")
            return
        self.data.setdefault("operadores", []).append(nome)
        self.data.setdefault("operador_posto_map", {})[nome] = posto
        save_data(self.data)
        _sync_operator_values()
        _refresh_rows(select_name=nome)

    def save_assignment():
        nome = str(op_nome_var.get() or "").strip()
        posto = str(op_posto_var.get() or "").strip() or "Geral"
        if not nome:
            messagebox.showerror("Erro", "Selecione ou indique um operador.")
            return
        if nome not in list(self.data.get("operadores", []) or []):
            self.data.setdefault("operadores", []).append(nome)
        self.data.setdefault("operador_posto_map", {})[nome] = posto
        save_data(self.data)
        _sync_operator_values()
        _refresh_rows(select_name=nome)

    def remove_op():
        nome = str(op_nome_var.get() or "").strip()
        if not nome:
            sel = tree.selection()
            if sel:
                vals = tree.item(sel[0], "values") or []
                nome = str(vals[0]) if vals else ""
        if not nome:
            return
        self.data["operadores"] = [o for o in list(self.data.get("operadores", []) or []) if str(o).strip() != nome]
        try:
            self.data.setdefault("operador_posto_map", {}).pop(nome, None)
        except Exception:
            pass
        if self.op_user.get() == nome:
            self.op_user.set("")
        save_data(self.data)
        op_nome_var.set("")
        op_posto_var.set("Geral")
        _sync_operator_values()
        _refresh_rows()

    Btn = ctk.CTkButton if use_custom else ttk.Button
    close_parent = btn_row
    Btn(btn_row, text="Novo", command=lambda: (op_nome_var.set(""), op_posto_var.set("Geral")), width=120 if use_custom else None).pack(side="left", padx=(0, 8))
    Btn(btn_row, text="Guardar", command=save_assignment, width=120 if use_custom else None).pack(side="left", padx=(0, 8))
    Btn(btn_row, text="Adicionar", command=add_op, width=120 if use_custom else None).pack(side="left", padx=(0, 8))
    Btn(btn_row, text="Remover", command=remove_op, width=120 if use_custom else None).pack(side="left", padx=(0, 8))
    Btn(btn_row, text="Fechar", command=on_close, width=120 if use_custom else None).pack(side="left")

    _refresh_rows()

def preview_operador_pdf(self):
    _ensure_configured()
    op_name = self.op_user.get().strip()
    enc = self.get_operador_encomenda(show_error=False)
    if not enc and hasattr(self, "op_tbl_enc"):
        sel = self.op_tbl_enc.selection()
        if sel:
            numero = self.op_tbl_enc.item(sel[0], "values")[0]
            enc = self.get_encomenda_by_numero(numero)
    if not enc and self.op_use_full_custom and getattr(self, "op_sel_enc_num", None):
        enc = self.get_encomenda_by_numero(self.op_sel_enc_num)
    if not enc:
        messagebox.showerror("Erro", "Selecione uma encomenda")
        return
    import tempfile
    path = os.path.join(tempfile.gettempdir(), f"lugest_producao_{enc.get('numero','')}.pdf")
    self.render_operador_producao_pdf(path, enc, op_name)
    try:
        os.startfile(path)
    except Exception:
        messagebox.showerror("Erro", "Não foi possível abrir a pré-visualização em PDF.")

def render_operador_producao_pdf(self, path, enc, op_name=""):
    _ensure_configured()
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as pdf_canvas
    from reportlab.lib import colors
    from reportlab.pdfbase import pdfmetrics
    from reportlab.lib.utils import ImageReader

    rows = []
    for m in enc.get("materiais", []):
        for e in m.get("espessuras", []):
            for p in e.get("pecas", []):
                rows.append(
                    {
                        "opp": p.get("opp", ""),
                        "ref_int": p.get("ref_interna", ""),
                        "ref_ext": p.get("ref_externa", ""),
                        "mat": p.get("material", m.get("material", "")),
                        "esp": p.get("espessura", e.get("espessura", "")),
                        "plan": parse_float(p.get("quantidade_pedida", 0), 0),
                        "ok": parse_float(p.get("produzido_ok", 0), 0),
                        "nok": parse_float(p.get("produzido_nok", 0), 0),
                        "qual": parse_float(p.get("produzido_qualidade", 0), 0),
                        "tempo_peca": parse_float(p.get("tempo_producao_min", 0), 0),
                        "estado": p.get("estado", ""),
                    }
                )
    rows.sort(key=lambda r: (str(r.get("mat", "")), parse_float(r.get("esp", 0), 0), str(r.get("ref_int", ""))))

    cliente_code = enc.get("cliente", "")
    cliente_nome = ""
    for cobj in self.data.get("clientes", []):
        if cobj.get("codigo") == cliente_code:
            cliente_nome = cobj.get("nome", "")
            break
    cliente_txt = f"{cliente_code} - {cliente_nome}".strip(" - ")

    reservas = enc.get("reservas", []) or []
    if reservas:
        res_txt = "; ".join(
            [f"{r.get('material','')} {r.get('espessura','')} x {fmt_num(r.get('quantidade', 0))}" for r in reservas if isinstance(r, dict)]
        )
        if not res_txt:
            res_txt = "-"
    else:
        res_txt = "-"

    tempo_total = parse_float(enc.get("tempo_producao_min", 0), 0)
    if tempo_total <= 0:
        tcalc = iso_diff_minutes(enc.get("inicio_producao"), enc.get("fim_producao"))
        tempo_total = parse_float(tcalc, 0) if tcalc is not None else 0
    tempo_esp = parse_float(enc.get("tempo_espessuras_min", 0), 0)
    tempo_pecas = parse_float(enc.get("tempo_pecas_min", 0), 0)
    tempo_avarias = _op_unique_avaria_minutes_for_encomenda(self.data, enc, include_open=True)
    tempo_operacional = round(max(0.0, tempo_pecas + tempo_avarias), 2)

    c = pdf_canvas.Canvas(path, pagesize=A4)
    width, height = A4
    margin = 20
    table_x = margin
    cols = [
        ("OPP", 60, "w"),
        ("Ref. Int.", 74, "w"),
        ("Ref. Ext.", 78, "w"),
        ("Material", 68, "w"),
        ("Esp.", 34, "center"),
        ("Plan.", 38, "right"),
        ("OK", 34, "right"),
        ("NOK", 34, "right"),
        ("Qual.", 36, "right"),
        ("Tempo", 48, "right"),
        ("Estado", 51, "w"),
    ]
    table_w = sum(w for _, w, _ in cols)
    row_h = 16
    header_h = 18

    def yinv(y):
        return height - y

    def clip_text(txt, max_w, size=8, bold=False):
        s = "" if txt is None else str(txt)
        fn = "Helvetica-Bold" if bold else "Helvetica"
        if pdfmetrics.stringWidth(s, fn, size) <= max_w:
            return s
        ell = "..."
        while s and pdfmetrics.stringWidth(s + ell, fn, size) > max_w:
            s = s[:-1]
        return (s + ell) if s else ""

    def draw_logo(x, y_top, max_w=82, max_h=28):
        logo = get_orc_logo_path()
        if not logo or not os.path.exists(logo):
            return
        try:
            img = ImageReader(logo)
            c.drawImage(img, x, height - y_top - max_h, width=max_w, height=max_h, preserveAspectRatio=True, mask="auto")
        except Exception:
            pass

    def draw_page_header(page_no, first_page):
        c.setStrokeColor(colors.HexColor("#c6cfdb"))
        c.rect(margin - 4, margin - 4, width - (2 * (margin - 4)), height - (2 * (margin - 4)), stroke=1, fill=0)
        c.setStrokeColor(colors.black)

        draw_logo(margin, 24, 82, 28)
        title_left = margin + 94
        title_right = width - margin - 160
        title_w = min(300, max(200, title_right - title_left))
        title_x = title_left + max(0.0, ((title_right - title_left) - title_w) / 2.0)
        title_y = 20
        title_h = 22
        c.setLineWidth(1.1)
        c.setStrokeColor(colors.HexColor("#8b1e2d"))
        c.roundRect(title_x, yinv(title_y + title_h), title_w, title_h, 6, stroke=1, fill=0)
        c.setFillColor(colors.HexColor("#8b1e2d"))
        c.setFont("Helvetica", 8)
        c.drawCentredString(title_x + (title_w / 2), yinv(title_y - 1), "Producao Operador")
        c.setFont("Helvetica-Bold", 11)
        c.drawCentredString(
            title_x + (title_w / 2),
            yinv(title_y + 15),
            clip_text(f"ENCOMENDA {enc.get('numero', '-')}", title_w - 12, size=11, bold=True),
        )

        c.setFillColor(colors.black)
        c.setFont("Helvetica", 8.5)
        c.drawRightString(width - margin, yinv(24), f"Data: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        c.drawRightString(width - margin, yinv(36), f"Pagina {page_no}")
        c.drawRightString(width - margin, yinv(48), f"Cliente: {clip_text(cliente_txt, 150, size=8.5)}")

        if first_page:
            box_y = 56
            box_h = 48
            c.setStrokeColor(colors.HexColor("#c6cfdb"))
            c.setLineWidth(0.8)
            c.rect(margin, yinv(box_y + box_h), table_w, box_h, stroke=1, fill=0)
            c.setStrokeColor(colors.black)
            c.setFont("Helvetica", 8.2)
            c.drawString(margin + 6, yinv(box_y + 12), f"Operador: {clip_text(op_name or '-', 150, size=8.2)}")
            c.drawString(margin + 250, yinv(box_y + 12), f"Inicio: {enc.get('inicio_producao', '-')}")
            c.drawString(margin + 6, yinv(box_y + 24), f"Fim: {enc.get('fim_producao', '-')}")
            c.drawString(margin + 250, yinv(box_y + 24), f"Tempo Encomenda: {fmt_num(tempo_total)} min")
            c.drawString(
                margin + 6,
                yinv(box_y + 36),
                f"Espessuras {fmt_num(tempo_esp)} | Pecas {fmt_num(tempo_pecas)} | Avarias {fmt_num(tempo_avarias)} | Total refs {fmt_num(tempo_operacional)} min",
            )
            c.drawString(margin + 6, yinv(box_y + 48), f"Chapas cativadas: {clip_text(res_txt, table_w - 14, size=8)}")
            return 112

        c.setFont("Helvetica", 8.5)
        c.drawString(margin, yinv(68), f"Cliente: {clip_text(cliente_txt, table_w - 8, size=8.5)}")
        c.line(margin, yinv(74), width - margin, yinv(74))
        return 84

    def draw_table_header(y_top):
        c.setFillColor(colors.HexColor("#8b1e2d"))
        c.rect(table_x, yinv(y_top + header_h), table_w, header_h, stroke=1, fill=1)
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 8.5)
        xx = table_x
        for name, cw, align in cols:
            if align == "right":
                c.drawRightString(xx + cw - 3, yinv(y_top + 12), name)
            elif align == "center":
                c.drawCentredString(xx + cw / 2, yinv(y_top + 12), name)
            else:
                c.drawString(xx + 3, yinv(y_top + 12), name)
            xx += cw
        c.setFillColor(colors.black)

    page = 1
    idx = 0
    while idx < len(rows) or idx == 0:
        first_page = page == 1
        table_top = draw_page_header(page, first_page)
        draw_table_header(table_top)
        y_rows = table_top + header_h
        max_y = height - margin - 24
        fit = max(1, int((max_y - y_rows) // row_h))
        count = min(fit, len(rows) - idx) if rows else 0

        if count == 0:
            c.setFont("Helvetica-Oblique", 9)
            c.drawString(table_x + 4, yinv(y_rows + 16), "Sem registos de producao para esta encomenda.")

        c.setFont("Helvetica", 8)
        for local_i in range(count):
            row = rows[idx + local_i]
            y = y_rows + (local_i * row_h)
            fill = colors.HexColor("#fff8f9") if ((idx + local_i) % 2 == 0) else colors.HexColor("#f6fbff")
            c.setFillColor(fill)
            c.rect(table_x, yinv(y + row_h), table_w, row_h, stroke=1, fill=1)
            c.setFillColor(colors.black)

            values = [
                str(row.get("opp", "") or ""),
                str(row.get("ref_int", "") or ""),
                str(row.get("ref_ext", "") or ""),
                str(row.get("mat", "") or ""),
                fmt_num(row.get("esp", "")),
                fmt_num(row.get("plan", 0)),
                fmt_num(row.get("ok", 0)),
                fmt_num(row.get("nok", 0)),
                fmt_num(row.get("qual", 0)),
                fmt_num(row.get("tempo_peca", 0)),
                str(row.get("estado", "") or ""),
            ]
            xx = table_x
            for (_, cw, align), val in zip(cols, values):
                txt = clip_text(val, cw - 6, 8)
                if align == "right":
                    c.drawRightString(xx + cw - 3, yinv(y + 11), txt)
                elif align == "center":
                    c.drawCentredString(xx + cw / 2, yinv(y + 11), txt)
                else:
                    c.drawString(xx + 3, yinv(y + 11), txt)
                xx += cw

        idx += count
        if idx < len(rows):
            c.showPage()
            page += 1

    c.save()

def _get_produtos_mov_operador(self, operador=""):
    _ensure_configured()
    rows = []
    for r in self.data.get("produtos_mov", []):
        norm = normalize_produto_mov_row(r)
        if norm:
            rows.append(norm)
    self.data["produtos_mov"] = rows
    if operador and operador != "Todos":
        rows = [r for r in rows if str(r.get("operador", "")).strip().lower() == operador.strip().lower()]
    rows.sort(key=lambda r: r.get("data", ""))
    return rows

def render_produtos_mov_operador_pdf(self, path, operador="Todos"):
    _ensure_configured()
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as pdf_canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.utils import ImageReader
    from reportlab.lib import colors

    c = pdf_canvas.Canvas(path, pagesize=A4)
    w, h = A4
    m = 20

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
        fn = font_bold if bold else font_regular
        if pdfmetrics.stringWidth(s, fn, size) <= max_w:
            return s
        ell = "..."
        while s and pdfmetrics.stringWidth(s + ell, fn, size) > max_w:
            s = s[:-1]
        return s + ell if s else ""

    def draw_logo(x, y_top, max_w=78, max_h=28):
        logo = get_orc_logo_path()
        if not logo or not os.path.exists(logo):
            return
        try:
            img = ImageReader(logo)
            c.drawImage(img, x, h - y_top - max_h, width=max_w, height=max_h, preserveAspectRatio=True, mask="auto")
        except Exception:
            pass

    rows = self._get_produtos_mov_operador(operador)
    # Ajustado para caber em A4 sem cortar a coluna de observacoes.
    cols = [
        ("Data", 84, "w"),
        ("Operador", 74, "w"),
        ("Codigo", 50, "w"),
        ("Descricao", 118, "w"),
        ("Qtd", 38, "e"),
        ("Antes", 38, "e"),
        ("Depois", 38, "e"),
        ("Obs", 115, "w"),
    ]
    table_w = sum(cw for _, cw, _ in cols)
    row_h = 16
    header_h = 18
    reserved_bottom = 36
    top_first = 70
    top_next = 58

    def draw_header(page):
        c.setStrokeColor(colors.HexColor("#c6cfdb"))
        c.rect(m - 4, m - 4, w - (2 * (m - 4)), h - (2 * (m - 4)), stroke=1, fill=0)
        c.setStrokeColor(colors.black)
        draw_logo(m + 2, 24)
        c.setFillColor(colors.HexColor("#8b1e2d"))
        set_font(True, 15)
        c.drawString(m + 88, h - 36, "Movimentos de Stock por Operador")
        c.setFillColor(colors.black)
        set_font(False, 9)
        c.drawRightString(w - m, h - 30, f"Data: {datetime.now().strftime('%d/%m/%Y')}")
        c.drawRightString(w - m, h - 44, f"Operador: {operador or 'Todos'}")
        c.drawRightString(w - m, h - 58, f"Página {page}")

    def draw_table_header(y_top):
        c.setFillColor(colors.HexColor("#8b1e2d"))
        c.rect(m, h - y_top - header_h, table_w, header_h, stroke=1, fill=1)
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

    idx = 0
    page = 1
    while idx < len(rows) or idx == 0:
        first = page == 1
        draw_header(page)
        top = top_first if first else top_next
        draw_table_header(top)
        fit = max(1, int((h - m - reserved_bottom - (top + header_h)) // row_h))
        count = min(fit, len(rows) - idx) if rows else 0

        set_font(False, 8)
        y_top = top + header_h
        for j in range(count):
            r = rows[idx + j]
            y_row = y_top + j * row_h
            fill = colors.HexColor("#f8fbff") if ((idx + j) % 2 == 0) else colors.HexColor("#eef4ff")
            c.setFillColor(fill)
            c.rect(m, h - y_row - row_h, table_w, row_h, stroke=1, fill=1)
            c.setFillColor(colors.black)
            vals = [
                str(r.get("data", ""))[:19].replace("T", " "),
                r.get("operador", ""),
                r.get("codigo", ""),
                r.get("descricao", ""),
                fmt_num(r.get("qtd", 0)),
                fmt_num(r.get("antes", 0)),
                fmt_num(r.get("depois", 0)),
                r.get("obs", ""),
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
        if idx >= len(rows):
            set_font(False, 8)
            c.drawString(m, 24, "Relatório luGEST")
            c.drawRightString(w - m, 24, f"Total movimentos: {len(rows)}")
            break
        c.showPage()
        page += 1

    c.save()

def produto_mov_operador_dialog(self, default_operador=""):
    _ensure_configured()
    use_custom = CUSTOM_TK_AVAILABLE and os.environ.get("USE_CUSTOM_PROD", "1") != "0"
    Dlg = ctk.CTkToplevel if use_custom else Toplevel
    Frm = ctk.CTkFrame if use_custom else ttk.Frame
    Lbl = ctk.CTkLabel if use_custom else ttk.Label
    Btn = ctk.CTkButton if use_custom else ttk.Button

    dlg = Dlg(self.root)
    dlg.title("Movimentos por Operador")
    dlg.geometry("1120x620")
    try:
        dlg.minsize(1060, 560)
        dlg.transient(self.root)
    except Exception:
        pass
    dlg.grab_set()

    op = StringVar(value=(default_operador or "Todos"))
    ops = ["Todos"] + sorted({str(o) for o in self.data.get("operadores", []) if str(o).strip()})

    host = Frm(dlg, fg_color="#f5f7fb", corner_radius=10) if use_custom else Frm(dlg)
    host.pack(fill="both", expand=True, padx=8, pady=8)

    top = Frm(host, fg_color="#eef2f8", corner_radius=8) if use_custom else Frm(host)
    top.pack(fill="x", padx=8, pady=8)
    Lbl(top, text="Operador", font=("Segoe UI", 13, "bold") if use_custom else None).pack(side="left", padx=8, pady=4)
    if use_custom:
        cb = ctk.CTkComboBox(
            top,
            variable=op,
            values=(ops or ["Todos"]),
            width=220,
            state="readonly",
            fg_color="white",
            border_color="#c8d2e3",
            button_color="#8b1e2d",
            button_hover_color="#992233",
            dropdown_fg_color="#ffffff",
            dropdown_hover_color="#f3d9df",
        )
    else:
        cb = ttk.Combobox(top, textvariable=op, values=ops, width=24, state="normal")
    cb.pack(side="left", padx=4, pady=4)

    cols = ("data", "operador", "codigo", "descricao", "qtd", "antes", "depois", "obs")
    tree_wrap = Frm(host, fg_color="#ffffff", corner_radius=8) if use_custom else Frm(host)
    tree_wrap.pack(fill="both", expand=True, padx=8, pady=6)

    tree_style = ""
    if use_custom:
        try:
            style = ttk.Style(dlg)
            style.configure(
                "MovOp.Treeview",
                background="#f8fbff",
                fieldbackground="#f8fbff",
                foreground="#111827",
                rowheight=34,
                borderwidth=0,
                relief="flat",
                font=("Segoe UI", 12),
            )
            style.configure(
                "MovOp.Treeview.Heading",
                background="#8b1e2d",
                foreground="white",
                font=("Segoe UI", 12, "bold"),
                relief="flat",
                borderwidth=0,
            )
            style.map(
                "MovOp.Treeview",
                background=[("selected", "#f8d7da")],
                foreground=[("selected", "#7a1222")],
            )
            style.map("MovOp.Treeview.Heading", background=[("active", "#992233")])
            tree_style = "MovOp.Treeview"
        except Exception:
            tree_style = ""

    tree = ttk.Treeview(tree_wrap, columns=cols, show="headings", style=tree_style)
    headers = {
        "data": "Data",
        "operador": "Operador",
        "codigo": "Código",
        "descricao": "Descrição",
        "qtd": "Qtd",
        "antes": "Antes",
        "depois": "Depois",
        "obs": "Obs",
    }
    for ccol in cols:
        tree.heading(ccol, text=headers[ccol])
        if ccol == "data":
            tree.column(ccol, width=145, minwidth=120, anchor="w", stretch=False)
        elif ccol == "operador":
            tree.column(ccol, width=130, minwidth=110, anchor="w", stretch=False)
        elif ccol == "codigo":
            tree.column(ccol, width=95, minwidth=80, anchor="w", stretch=False)
        elif ccol == "descricao":
            tree.column(ccol, width=230, minwidth=180, anchor="w", stretch=True)
        elif ccol in ("qtd", "antes", "depois"):
            tree.column(ccol, width=85, minwidth=70, anchor="e", stretch=False)
        else:
            tree.column(ccol, width=320, minwidth=220, anchor="w", stretch=True)
    if use_custom:
        vsb = ctk.CTkScrollbar(
            tree_wrap,
            orientation="vertical",
            command=tree.yview,
            fg_color="#e7edf8",
            button_color="#8b1e2d",
            button_hover_color="#992233",
        )
        hsb = ctk.CTkScrollbar(
            tree_wrap,
            orientation="horizontal",
            command=tree.xview,
            fg_color="#e7edf8",
            button_color="#8b1e2d",
            button_hover_color="#992233",
        )
    else:
        vsb = ttk.Scrollbar(tree_wrap, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(tree_wrap, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")
    tree_wrap.rowconfigure(0, weight=1)
    tree_wrap.columnconfigure(0, weight=1)
    tree.tag_configure("even", background="#f8fbff")
    tree.tag_configure("odd", background="#edf3ff")

    def refresh_rows(*_):
        for iid in tree.get_children():
            tree.delete(iid)
        rows = self._get_produtos_mov_operador(op.get().strip())
        for idx, r in enumerate(rows):
            tag = "odd" if idx % 2 else "even"
            tree.insert(
                "",
                END,
                values=(
                    str(r.get("data", ""))[:19].replace("T", " "),
                    r.get("operador", ""),
                    r.get("codigo", ""),
                    r.get("descricao", ""),
                    fmt_num(r.get("qtd", 0)),
                    fmt_num(r.get("antes", 0)),
                    fmt_num(r.get("depois", 0)),
                    r.get("obs", ""),
                ),
                tags=(tag,),
            )

    def preview_pdf():
        operador = op.get().strip()
        prev_dir = os.path.join(BASE_DIR, "previews")
        try:
            os.makedirs(prev_dir, exist_ok=True)
        except Exception:
            prev_dir = tempfile.gettempdir()
        path = os.path.join(prev_dir, f"mov_operador_{(operador or 'todos').replace(' ', '_')}.pdf")
        self.render_produtos_mov_operador_pdf(path, operador or "Todos")
        ok = self._open_pdf_default(path)
        if not ok:
            messagebox.showwarning("Aviso", f"PDF gerado, mas não abriu automaticamente.\n{path}")

    def save_pdf():
        operador = op.get().strip() or "todos"
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile=f"mov_operador_{operador.replace(' ', '_')}.pdf",
        )
        if not path:
            return
        self.render_produtos_mov_operador_pdf(path, op.get().strip() or "Todos")
        messagebox.showinfo("OK", "PDF guardado")

    def print_pdf():
        operador = op.get().strip()
        path = os.path.join(tempfile.gettempdir(), "lugest_mov_operador.pdf")
        self.render_produtos_mov_operador_pdf(path, operador or "Todos")
        try:
            os.startfile(path, "print")
        except Exception:
            self._open_pdf_default(path)

    if use_custom:
        cb.configure(command=lambda _v=None: refresh_rows())
    else:
        cb.bind("<<ComboboxSelected>>", refresh_rows)
    btns = Frm(host, fg_color="#eef2f8", corner_radius=8) if use_custom else Frm(host)
    btns.pack(fill="x", padx=8, pady=8)
    Btn(btns, text="Atualizar", command=refresh_rows).pack(side="left", padx=4)
    Btn(btns, text="Pré-visualizar", command=preview_pdf).pack(side="left", padx=4)
    Btn(btns, text="Guardar PDF", command=save_pdf).pack(side="left", padx=4)
    Btn(btns, text="Imprimir", command=print_pdf).pack(side="left", padx=4)
    Btn(btns, text="Fechar", command=dlg.destroy).pack(side="right", padx=4)

    refresh_rows()

def _open_operador_encomenda_detail(self, numero):
    _ensure_configured()
    enc = self.get_encomenda_by_numero(numero)
    if not enc:
        messagebox.showerror("Erro", "Encomenda não encontrada")
        return
    if hasattr(self, "op_detail_win") and self.op_detail_win and self.op_detail_win.winfo_exists():
        try:
            self.op_detail_win.destroy()
        except Exception:
            pass
    self.op_sel_enc_num = numero
    self.op_material = None
    self.op_espessura = None
    self.op_sel_peca = None
    self.op_sel_pecas_ids = set()

    use_custom = CUSTOM_TK_AVAILABLE and getattr(self, "op_use_custom", False)
    Win = ctk.CTkToplevel if use_custom else Toplevel
    Frame = ctk.CTkFrame if use_custom else ttk.Frame
    Label = ctk.CTkLabel if use_custom else ttk.Label
    Button = ctk.CTkButton if use_custom else ttk.Button

    def ActionBtn(parent, **kwargs):
        if use_custom:
            kwargs.setdefault("height", 44)
            kwargs.setdefault("corner_radius", 10)
            kwargs.setdefault("font", ("Segoe UI", 13, "bold"))
        return Button(parent, **kwargs)
    win = Win(self.root)
    win.title(f"Operador - Encomenda {numero}")
    try:
        win.geometry("1320x860")
        if use_custom:
            win.minsize(1200, 760)
        try:
            win.state("zoomed")
        except Exception:
            pass
        win.transient(self.root)
    except Exception:
        pass
    self.op_detail_win = win

    wrap = Frame(win, fg_color="#f8fafc") if use_custom else Frame(win)
    wrap.pack(fill="both", expand=True, padx=8, pady=8)
    if use_custom:
        def _info_box(parent, title, value):
            box = ctk.CTkFrame(parent, fg_color="#eef4ff", corner_radius=10, border_width=1, border_color="#d7e3f3")
            box.pack(side="left", fill="both", expand=True, padx=5, pady=4)
            ctk.CTkLabel(box, text=title, font=("Segoe UI", 11, "bold"), text_color="#475569").pack(anchor="w", padx=10, pady=(7, 0))
            lbl = ctk.CTkLabel(box, text=value, font=("Segoe UI", 15, "bold"), text_color="#0f172a")
            lbl.pack(anchor="w", padx=10, pady=(2, 8))
            return lbl

        def _section_card(parent, title, subtitle=""):
            card = ctk.CTkFrame(parent, fg_color="#ffffff", border_width=1, border_color="#dbe4f0", corner_radius=14)
            title_row = ctk.CTkFrame(card, fg_color="transparent")
            title_row.pack(fill="x", padx=10, pady=(8, 6))
            ctk.CTkLabel(title_row, text=title, font=("Segoe UI", 14, "bold"), text_color="#0f172a").pack(side="left")
            if subtitle:
                ctk.CTkLabel(title_row, text=subtitle, font=("Segoe UI", 11, "bold"), text_color="#64748b").pack(side="right")
            return card

        def _header(parent, columns):
            hdr = ctk.CTkFrame(parent, fg_color=COR_OP_HEADER, corner_radius=10)
            hdr.pack(fill="x", padx=8, pady=(0, 6))
            for idx, (title, weight, anchor) in enumerate(columns):
                hdr.grid_columnconfigure(idx, weight=weight)
                lbl = ctk.CTkLabel(
                    hdr,
                    text=title,
                    font=("Segoe UI", 12, "bold"),
                    text_color="#ffffff",
                    anchor=anchor,
                )
                sticky = "ew"
                if anchor == "w":
                    sticky = "w"
                elif anchor == "e":
                    sticky = "e"
                lbl.grid(row=0, column=idx, sticky=sticky, padx=10, pady=8)
            return hdr
    else:
        _info_box = None
        _section_card = None
        _header = None

    top = Frame(wrap, fg_color="white", border_width=1, border_color="#dbe4f0", corner_radius=14) if use_custom else Frame(wrap)
    top.pack(fill="x", padx=4, pady=(4, 6))
    top_left = Frame(top, fg_color="transparent") if use_custom else Frame(top)
    top_left.pack(side="left", fill="x", expand=True, padx=10, pady=8)
    cli = enc.get("cliente", "")
    Label(top_left, text=f"Operador / Produção", font=("Segoe UI", 18, "bold") if use_custom else None, text_color="#0f172a" if use_custom else None).pack(anchor="w")
    Label(top_left, text=f"Encomenda {numero}  |  Cliente {cli}", font=("Segoe UI", 13, "bold") if use_custom else None, text_color="#475569" if use_custom else None).pack(anchor="w", pady=(2, 0))
    ActionBtn(top, text="Fechar", command=win.destroy, width=120 if use_custom else None).pack(side="right", padx=10, pady=10)

    ctx_card = Frame(
        wrap,
        fg_color="#ffffff",
        border_width=1,
        border_color="#dbe4f0",
        corner_radius=14,
    ) if use_custom else Frame(wrap)
    ctx_card.pack(fill="x", padx=4, pady=(0, 6))
    ctx_line_1 = Frame(ctx_card, fg_color="transparent") if use_custom else Frame(ctx_card)
    ctx_line_1.pack(fill="x", padx=8, pady=(6, 2))
    ctx_line_2 = Frame(ctx_card, fg_color="transparent") if use_custom else Frame(ctx_card)
    ctx_line_2.pack(fill="x", padx=8, pady=(0, 6))
    ctx_line_3 = Frame(ctx_card, fg_color="transparent") if use_custom else Frame(ctx_card)
    ctx_line_3.pack(fill="x", padx=8, pady=(0, 8))
    if use_custom:
        self.op_ctx_operador_lbl = _info_box(ctx_line_1, "Operador", str(self.op_user.get() or "-"))
        self.op_ctx_posto_lbl = _info_box(ctx_line_1, "Posto", str(self.op_posto.get() or "-"))
        self.op_ctx_estado_lbl = _info_box(ctx_line_1, "Estado da encomenda", str(enc.get("estado") or "Preparacao"))
        self.op_ctx_esp_sel_lbl = _info_box(ctx_line_1, "Espessura ativa", "-")
        self.op_ctx_total_lbl = _info_box(ctx_line_2, "Total de peças", "0")
        self.op_ctx_exec_lbl = _info_box(ctx_line_2, "Em curso", "0")
        self.op_ctx_pausa_lbl = _info_box(ctx_line_2, "Pausa / Interr.", "0")
        self.op_ctx_avaria_lbl = _info_box(ctx_line_2, "Avarias", "0")
        self.op_ctx_conc_lbl = _info_box(ctx_line_2, "Concluídas", "0")
        self.op_ctx_alerta_box = ctk.CTkFrame(
            ctx_line_3,
            corner_radius=12,
            fg_color="#ecfccb",
            border_width=1,
            border_color="#84cc16",
        )
        self.op_ctx_alerta_box.pack(fill="x", padx=2)
        self.op_ctx_alerta_lbl = ctk.CTkLabel(
            self.op_ctx_alerta_box,
            text="Sem avarias abertas na seleção.",
            anchor="w",
            fg_color="transparent",
            text_color="#365314",
            font=("Segoe UI", 11, "bold"),
            height=34,
        )
        self.op_ctx_alerta_lbl.pack(fill="x", padx=10)

    actions_card = Frame(
        wrap,
        fg_color="#ffffff",
        border_width=1,
        border_color="#dbe4f0",
        corner_radius=14,
    ) if use_custom else Frame(wrap)
    actions_card.pack(fill="x", padx=4, pady=(0, 6))
    if use_custom:
        Label(actions_card, text="Quadro de Produção", font=("Segoe UI", 14, "bold"), text_color="#0f172a").pack(anchor="w", padx=10, pady=(8, 2))
    btn_line_1 = Frame(actions_card, fg_color="transparent") if use_custom else Frame(actions_card)
    btn_line_1.pack(fill="x", padx=8, pady=(2, 4))
    ActionBtn(btn_line_1, text="Pré-visualizar Produção", command=self.preview_operador_pdf, width=180 if use_custom else None).pack(side="left", padx=(0, 8))
    ActionBtn(btn_line_1, text="Imprimir Ordem de Fabrico", command=self.print_of_selected_encomenda, width=200 if use_custom else None).pack(side="left", padx=(0, 8))
    ActionBtn(btn_line_1, text="Histórico Rejeitadas", command=self.show_rejeitadas_hist, width=165 if use_custom else None).pack(side="left", padx=(0, 8))
    if use_custom:
        Label(
            actions_card,
            text="Fluxo: selecionar espessura -> selecionar peça -> executar operação",
            text_color="#64748b",
            font=("Segoe UI", 11, "bold"),
        ).pack(anchor="w", padx=10, pady=(0, 8))

    self.op_esp_var = StringVar(value="")
    self.op_lote_var = BooleanVar(value=False)
    select_card = Frame(
        wrap,
        fg_color="#ffffff",
        border_width=1,
        border_color="#dbe4f0",
        corner_radius=14,
    ) if use_custom else Frame(wrap)
    select_card.pack(fill="x", padx=4, pady=(0, 6))
    select_line = Frame(select_card, fg_color="transparent") if use_custom else Frame(select_card)
    select_line.pack(fill="x", padx=10, pady=8)
    Label(select_line, text="Selecionar espessura", font=("Segoe UI", 13, "bold") if use_custom else None).pack(side="left", padx=(0, 8))
    if use_custom:
        self.op_esp_cb = ctk.CTkComboBox(
            select_line,
            variable=self.op_esp_var,
            values=[],
            width=320,
            command=lambda _v=None: self._op_select_esp_combo(),
        )
        self.op_esp_cb.pack(side="left", padx=(0, 8))
        ActionBtn(select_line, text="Carregar", command=self._op_select_esp_combo, width=110).pack(side="left", padx=(0, 10))
        ctk.CTkCheckBox(
            select_line,
            text="Modo lote (espessura)",
            variable=self.op_lote_var,
            font=("Segoe UI", 12, "bold"),
            text_color="#334155",
        ).pack(side="left")
    else:
        self.op_esp_cb = ttk.Combobox(select_line, textvariable=self.op_esp_var, values=[], width=38, state="readonly")
        self.op_esp_cb.pack(side="left", padx=(0, 8))
        self.op_esp_cb.bind("<<ComboboxSelected>>", lambda _e=None: self._op_select_esp_combo())
        ttk.Checkbutton(select_line, text="Modo lote (espessura)", variable=self.op_lote_var).pack(side="left", padx=(14, 0))

    body = Frame(wrap, fg_color="transparent") if use_custom else Frame(wrap)
    body.pack(fill="both", expand=True, padx=4, pady=(0, 4))
    if use_custom:
        body.grid_columnconfigure(0, weight=2)
        body.grid_columnconfigure(1, weight=5)
        body.grid_rowconfigure(0, weight=1)

    esp_section = _section_card(body, "Material / Espessuras", "Seleção rápida") if use_custom else Frame(body)
    if use_custom:
        esp_section.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
    else:
        esp_section.pack(side="left", fill="both", expand=False)
    if use_custom:
        _header(
            esp_section,
            [
                ("Material", 4, "w"),
                ("Esp.", 2, "w"),
                ("Plan.", 1, "e"),
                ("Prod.", 1, "e"),
                ("Estado", 2, "w"),
            ],
        )
    self.op_esp_list = ctk.CTkScrollableFrame(
        esp_section,
        fg_color="white",
        corner_radius=12,
        border_width=0,
        height=540,
    ) if use_custom else ttk.Frame(esp_section)
    self.op_esp_list.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    right = Frame(body, fg_color="transparent") if use_custom else Frame(body)
    if use_custom:
        right.grid(row=0, column=1, sticky="nsew")
        right.grid_rowconfigure(1, weight=1)
    else:
        right.pack(side="left", fill="both", expand=True)

    pecas_actions = Frame(
        right,
        fg_color="#ffffff",
        border_width=1,
        border_color="#dbe4f0",
        corner_radius=14,
    ) if use_custom else Frame(right)
    pecas_actions.pack(fill="x", pady=(0, 6))
    hdr_line = Frame(pecas_actions, fg_color="transparent") if use_custom else Frame(pecas_actions)
    hdr_line.pack(fill="x", padx=10, pady=(8, 2))
    Label(hdr_line, text="Ações da Produção", font=("Segoe UI", 14, "bold") if use_custom else None, text_color="#0f172a" if use_custom else None).pack(side="left")
    self.op_select_all_var = BooleanVar(value=False)
    if use_custom:
        ctk.CTkCheckBox(
            hdr_line,
            text="Selecionar todas",
            variable=self.op_select_all_var,
            font=("Segoe UI", 12, "bold"),
            text_color="#334155",
            command=lambda: _op_set_select_all_pecas(self, bool(self.op_select_all_var.get())),
        ).pack(side="right")
    pecas_btns = Frame(pecas_actions, fg_color="transparent") if use_custom else Frame(pecas_actions)
    pecas_btns.pack(fill="x", padx=10, pady=(2, 8))
    if use_custom:
        ActionBtn(
            pecas_btns,
            text="Iniciar",
            command=self.operador_inicio_peca,
            width=130,
            fg_color="#f59e0b",
            hover_color="#d97706",
            text_color="#ffffff",
        ).pack(side="left", padx=(0, 8))
    else:
        ActionBtn(pecas_btns, text="Iniciar", command=self.operador_inicio_peca, width=130).pack(side="left", padx=(0, 8))
    ActionBtn(pecas_btns, text="Retomar", command=self.operador_retomar_peca, width=120 if use_custom else None).pack(side="left", padx=(0, 8))
    ActionBtn(pecas_btns, text="Finalizar", command=self.operador_fim_peca, width=120 if use_custom else None).pack(side="left", padx=(0, 8))
    ActionBtn(pecas_btns, text="Interromper", command=self.operador_interromper_peca, width=130 if use_custom else None).pack(side="left", padx=(0, 8))
    ActionBtn(pecas_btns, text="Registar Avaria", command=self.operador_registar_avaria, width=150 if use_custom else None).pack(side="left", padx=(0, 8))
    ActionBtn(pecas_btns, text="Fim Avaria", command=self.operador_fechar_avaria, width=130 if use_custom else None).pack(side="left", padx=(0, 8))
    ActionBtn(pecas_btns, text="Atualizar", command=lambda: _op_force_refresh_page(self), width=115 if use_custom else None).pack(side="left", padx=(0, 8))
    ActionBtn(pecas_btns, text="Ver Etiqueta", command=self.preview_peca_pdf, width=130 if use_custom else None).pack(side="left", padx=(0, 8))
    ActionBtn(pecas_btns, text="Ver desenho", command=lambda: operador_ver_desenho(self), width=130 if use_custom else None).pack(side="left", padx=(0, 8))
    ActionBtn(pecas_btns, text="Reabrir", command=self.operador_reabrir_peca_total, width=115 if use_custom else None).pack(side="left", padx=(0, 8))
    if use_custom:
        ActionBtn(
            pecas_btns,
            text="Alertar Chefia",
            command=lambda: operador_alertar_chefia(self),
            width=150,
            fg_color="#b91c1c",
            hover_color="#991b1b",
            text_color="#ffffff",
        ).pack(side="left", padx=(0, 8))
    else:
        ActionBtn(pecas_btns, text="Alertar Chefia", command=lambda: operador_alertar_chefia(self)).pack(side="left", padx=(0, 8))

    peca_section = _section_card(right, "Quadro peça a peça", "Leitura por operação e estado") if use_custom else Frame(right)
    peca_section.pack(fill="both", expand=True)
    if use_custom:
        _op_make_fixed_header(peca_section)
    self.op_pecas_list = ctk.CTkScrollableFrame(
        peca_section,
        fg_color="white",
        corner_radius=12,
        border_width=0,
    ) if use_custom else ttk.Frame(peca_section)
    self.op_pecas_list.pack(fill="both", expand=True, padx=8, pady=(0, 8))
    if use_custom:
        try:
            self.op_pecas_list.configure(height=660)
        except Exception:
            pass

    def _on_close():
        try:
            win.destroy()
        finally:
            self.op_detail_win = None

    try:
        win.protocol("WM_DELETE_WINDOW", _on_close)
    except Exception:
        pass

    self.on_operador_select(None)

def finalizar_espessura(self, enc, mat, esp, auto=False):
    _ensure_configured()
    esp_obj = self.get_operador_esp_obj(enc, mat, esp)
    if not esp_obj:
        return False
    pecas = esp_obj.get("pecas", [])
    if not pecas or not all(p.get("estado") == "Concluida" for p in pecas):
        if not auto:
            messagebox.showerror("Erro", "Ainda existem peças por concluir")
        return False
    esp_obj["estado"] = "Concluida"
    for p in esp_obj.get("pecas", []):
        atualizar_estado_peca(p)
    ts_fim_esp = now_iso()
    esp_obj["fim_producao"] = ts_fim_esp
    dur_esp = iso_diff_minutes(esp_obj.get("inicio_producao"), ts_fim_esp)
    if dur_esp is not None:
        esp_obj["tempo_producao_min"] = dur_esp
    # NUNCA pedir baixa no Embalamento/fecho final.
    # A baixa de MP e sobras são EXCLUSIVAS do fecho do Corte Laser.
    if _esp_tem_laser(esp_obj) and (not _esp_baixa_laser_resolvida(esp_obj)):
        esp_obj["baixa_laser_confirmada_sem_baixa"] = True
        esp_obj["baixa_laser_em"] = now_iso()
        try:
            log_stock(
                self.data,
                "SEM_BAIXA_CONFIRMADA",
                f"encomenda={enc.get('numero','')} mat={mat} esp={esp} motivo=confirmacao_no_fecho",
            )
        except Exception:
            pass
    update_estado_encomenda_por_espessuras(enc)
    save_data(self.data)
    try:
        if hasattr(self, "mark_tab_dirty"):
            self.mark_tab_dirty("materia", "encomendas", "plano", "expedicao")
    except Exception:
        pass
    return True

def _preview_piece_pdf(self, enc, p):
    _ensure_configured()
    if not p.get("opp"):
        p["opp"] = next_opp_numero(self.data)
        save_data(self.data)
    opp = p.get("opp", "")
    codigo = p.get("ref_interna", "")
    qtd = p.get("quantidade_pedida", 0)
    material = p.get("material", "")
    lote = p.get("lote_baixa") or "-"
    for r in enc.get("reservas", []):
        if r.get("material") == material and r.get("espessura") == p.get("espessura"):
            if r.get("material_id"):
                for m in self.data.get("materiais", []):
                    if m.get("id") == r.get("material_id"):
                        lote = m.get("lote_fornecedor", "") or "-"
                        break
            break

    use_custom = CUSTOM_TK_AVAILABLE and (getattr(self, "op_use_custom", False) or self.op_use_full_custom)
    owner = self.root
    try:
        if hasattr(self, "op_detail_win") and self.op_detail_win and self.op_detail_win.winfo_exists():
            owner = self.op_detail_win
    except Exception:
        owner = self.root
    win = ctk.CTkToplevel(owner) if use_custom else Toplevel(owner)
    win.title("Pré-visualização Peça")
    try:
        win.transient(owner)
        if use_custom:
            win.configure(fg_color="#f3f6fb")
    except Exception:
        pass
    try:
        win.lift()
        win.focus_force()
        win.attributes("-topmost", True)
        win.after(220, lambda w=win: (w.winfo_exists() and w.attributes("-topmost", False)))
    except Exception:
        pass
    wrap = ctk.CTkFrame(win, fg_color="#f3f6fb") if use_custom else ttk.Frame(win)
    wrap.pack(fill="both", expand=True, padx=10, pady=10)
    canvas = Canvas(wrap, width=420, height=260, bg="white")
    canvas.pack(fill="both", expand=True)

    # header band
    canvas.create_rectangle(0, 0, 420, 28, fill="#8b1e2d", outline="#8b1e2d")
    canvas.create_text(12, 6, anchor="nw", text="Etiqueta de Produção", font=("Segoe UI", 9, "bold"), fill="white")

    # logo
    self._piece_logo = None
    try:
        logo_path = get_orc_logo_path()
        if logo_path and os.path.exists(logo_path):
            try:
                from PIL import Image, ImageTk
                img = Image.open(logo_path)
                img = img.resize((90, 30))
                self._piece_logo = ImageTk.PhotoImage(img)
                canvas.create_image(10, 34, anchor="nw", image=self._piece_logo)
            except Exception:
                pass
    except Exception:
        pass

    canvas.create_text(110, 35, anchor="nw", text=f"Peça: {codigo}", font=("Segoe UI", 12, "bold"))
    canvas.create_text(110, 55, anchor="nw", text=f"Ordem de Fabrico: {opp}", font=("Segoe UI", 10))
    canvas.create_text(110, 73, anchor="nw", text=f"Material: {material}  |  Lote: {lote}", font=("Segoe UI", 9))
    canvas.create_text(110, 90, anchor="nw", text=f"Quantidade: {qtd}", font=("Segoe UI", 9))
    canvas.create_line(10, 112, 410, 112, width=1)
    canvas.create_text(10, 120, anchor="nw", text="Código de Barras (Ordem de Fabrico):", font=("Segoe UI", 9, "bold"))
    canvas.create_rectangle(10, 138, 410, 210, outline="#111827")
    canvas.create_text(20, 160, anchor="nw", text=opp, font=("Segoe UI", 10, "bold"))

    def render_pdf(path):
        from reportlab.lib.units import mm
        from reportlab.pdfgen import canvas as pdf_canvas
        from reportlab.graphics.barcode import code128
        from reportlab.lib.utils import ImageReader
        width, height = (110 * mm, 50 * mm)
        c = pdf_canvas.Canvas(path, pagesize=(width, height))
        def yinv(y):
            return height - y
        # header band
        c.setFillColorRGB(0.12, 0.31, 0.48)
        c.rect(0, yinv(16), width, 16, stroke=0, fill=1)
        c.setFillColorRGB(1, 1, 1)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(8, yinv(12), "Etiqueta de Producao")

        # logo (top-left)
        logo = get_orc_logo_path()
        if logo and os.path.exists(logo):
            try:
                img = ImageReader(logo)
                c.drawImage(img, 6, yinv(42), width=18*mm, height=8*mm, preserveAspectRatio=True, mask="auto")
            except Exception:
                pass

        c.setFillColorRGB(0, 0, 0)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(28, yinv(32), f"Peca: {codigo}")
        c.setFont("Helvetica", 9)
        c.drawString(28, yinv(44), f"Ordem de Fabrico: {opp}")
        c.drawString(28, yinv(54), f"Material: {material}  Lote: {lote}")
        c.drawString(28, yinv(64), f"Quantidade: {qtd}")
        c.line(6, yinv(70), width - 6, yinv(70))
        c.setFont("Helvetica-Bold", 8)
        c.drawString(6, yinv(78), "Codigo de Barras (Ordem de Fabrico):")
        barcode = code128.Code128(opp, barHeight=12*mm, barWidth=0.6)
        barcode.drawOn(c, 6, yinv(102))
        c.save()

    def save_pdf():
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        render_pdf(path)
        messagebox.showinfo("OK", "PDF guardado com sucesso.")

    def print_pdf():
        import tempfile
        path = os.path.join(tempfile.gettempdir(), f"lugest_peca_{codigo}.pdf")
        render_pdf(path)
        try:
            os.startfile(path, "print")
        except Exception:
            try:
                os.startfile(path)
            except Exception:
                messagebox.showerror("Erro", "Não foi possível abrir a impressão.")

    def open_with_pdf():
        import tempfile
        path = os.path.join(tempfile.gettempdir(), f"lugest_peca_{codigo}.pdf")
        render_pdf(path)
        try:
            subprocess.Popen(["rundll32", "shell32.dll,OpenAs_RunDLL", path])
        except Exception:
            messagebox.showerror("Erro", "Não foi possível abrir o seletor de aplicações.")

    btns = ctk.CTkFrame(win, fg_color="#f3f6fb") if use_custom else ttk.Frame(win)
    btns.pack(fill="x", padx=10, pady=(0, 10))
    Btn = ctk.CTkButton if use_custom else ttk.Button
    Btn(btns, text="Guardar PDF", command=save_pdf, width=130 if use_custom else None).pack(side="left", padx=6)
    Btn(btns, text="Imprimir PDF", command=print_pdf, width=130 if use_custom else None).pack(side="left", padx=6)
    Btn(btns, text="Abrir com...", command=open_with_pdf, width=130 if use_custom else None).pack(side="left", padx=6)

def print_of_selected_encomenda(self):
    _ensure_configured()
    if self.op_use_full_custom:
        numero = (self.op_sel_enc_num or "").strip()
        if not numero:
            messagebox.showerror("Erro", "Selecione uma encomenda")
            return
    else:
        sel = self.op_tbl_enc.selection() if hasattr(self, "op_tbl_enc") else None
        if not sel:
            messagebox.showerror("Erro", "Selecione uma encomenda")
            return
        numero = self.op_tbl_enc.item(sel[0], "values")[0]
    enc = self.get_encomenda_by_numero(numero)
    if not enc:
        messagebox.showerror("Erro", "Encomenda não encontrada")
        return
    import tempfile
    path = os.path.join(tempfile.gettempdir(), f"lugest_of_{numero}.pdf")
    self.render_of_pdf(path, enc)
    try:
        os.startfile(path, "print")
    except Exception:
        try:
            os.startfile(path)
        except Exception:
            messagebox.showerror("Erro", "Não foi possível imprimir a Ordem de Fabrico.")

def render_of_pdf(self, path, enc):
    _ensure_configured()
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as pdf_canvas
    from reportlab.pdfbase import pdfmetrics
    from reportlab.lib import colors
    from reportlab.lib.utils import ImageReader

    w, h = A4
    m = 20
    c = pdf_canvas.Canvas(path, pagesize=A4)

    def yinv(y):
        return h - y

    def clip_text(txt, max_w, size=8, bold=False):
        s = "" if txt is None else str(txt)
        fn = "Helvetica-Bold" if bold else "Helvetica"
        if pdfmetrics.stringWidth(s, fn, size) <= max_w:
            return s
        ell = "..."
        while s and pdfmetrics.stringWidth(s + ell, fn, size) > max_w:
            s = s[:-1]
        return (s + ell) if s else ""

    def draw_logo(x, y_top, max_w=82, max_h=28):
        logo = get_orc_logo_path()
        if not logo or not os.path.exists(logo):
            return
        try:
            img = ImageReader(logo)
            c.drawImage(img, x, h - y_top - max_h, width=max_w, height=max_h, preserveAspectRatio=True, mask="auto")
        except Exception:
            pass

    cli_code = str(enc.get("cliente", "") or "")
    cli_nome = ""
    for cl in self.data.get("clientes", []):
        if str(cl.get("codigo", "")) == cli_code:
            cli_nome = str(cl.get("nome", "") or "")
            break
    cli_txt = f"{cli_code} - {cli_nome}".strip(" - ")
    tempo_total = parse_float(enc.get("tempo_producao_min", 0), 0)
    if tempo_total <= 0:
        tcalc = iso_diff_minutes(enc.get("inicio_producao"), enc.get("fim_producao"))
        tempo_total = parse_float(tcalc, 0) if tcalc is not None else 0
    tempo_esp = parse_float(enc.get("tempo_espessuras_min", 0), 0)
    tempo_pecas = parse_float(enc.get("tempo_pecas_min", 0), 0)
    tempo_avarias = _op_unique_avaria_minutes_for_encomenda(self.data, enc, include_open=True)
    tempo_operacional = round(max(0.0, tempo_pecas + tempo_avarias), 2)

    reservas = enc.get("reservas", []) or []
    res_list = []
    for r in reservas:
        if not isinstance(r, dict):
            continue
        res_list.append(
            f"{r.get('material', '')} {r.get('espessura', '')} x {fmt_num(r.get('quantidade', 0))}"
        )
    res_txt = "; ".join([x for x in res_list if x.strip()]) or "-"

    rows = list(encomenda_pecas(enc))
    cols = [
        ("OF", 86, "w"),
        ("Ref. Int.", 86, "w"),
        ("Ref. Ext.", 100, "w"),
        ("Material", 88, "w"),
        ("Esp.", 40, "center"),
        ("Qtd", 44, "right"),
        ("Observações", 111, "w"),
    ]
    table_w = sum(cw for _, cw, _ in cols)
    row_h = 16
    header_h = 18

    def draw_page_header(page_no, first_page):
        c.setStrokeColor(colors.HexColor("#c6cfdb"))
        c.rect(m - 4, m - 4, w - (2 * (m - 4)), h - (2 * (m - 4)), stroke=1, fill=0)
        c.setStrokeColor(colors.black)

        draw_logo(m, 24, 82, 28)
        title_left = m + 94
        title_right = w - m - 160
        title_w = min(300, max(190, title_right - title_left))
        title_x = title_left + max(0.0, ((title_right - title_left) - title_w) / 2.0)
        title_y = 20
        title_h = 22
        c.setLineWidth(1.1)
        c.setStrokeColor(colors.HexColor("#8b1e2d"))
        c.roundRect(title_x, yinv(title_y + title_h), title_w, title_h, 6, stroke=1, fill=0)
        c.setFillColor(colors.HexColor("#8b1e2d"))
        c.setFont("Helvetica", 8)
        c.drawCentredString(title_x + (title_w / 2), yinv(title_y - 1), "Ordem de Fabrico")
        c.setFont("Helvetica-Bold", 11)
        c.drawCentredString(
            title_x + (title_w / 2),
            yinv(title_y + 15),
            clip_text(f"ENCOMENDA {enc.get('numero', '')}", title_w - 12, size=11, bold=True),
        )

        c.setFillColor(colors.black)
        c.setFont("Helvetica", 8.5)
        c.drawRightString(w - m, yinv(24), f"Data: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        c.drawRightString(w - m, yinv(36), f"Pagina {page_no}")
        c.drawRightString(w - m, yinv(48), f"Cliente: {clip_text(cli_txt, 150, size=8.5)}")

        if first_page:
            box_y = 56
            box_h = 56
            c.setStrokeColor(colors.HexColor("#c6cfdb"))
            c.setLineWidth(0.8)
            c.rect(m, yinv(box_y + box_h), table_w, box_h, stroke=1, fill=0)
            c.setStrokeColor(colors.black)
            c.setFont("Helvetica", 8.5)
            c.drawString(m + 6, yinv(box_y + 12), f"Cliente: {clip_text(cli_txt, table_w - 14, size=8.5)}")
            c.drawString(m + 6, yinv(box_y + 25), f"Tempo Encomenda: {fmt_num(tempo_total)} min")
            c.drawString(m + 250, yinv(box_y + 25), f"Tempo Espessuras: {fmt_num(tempo_esp)} min")
            c.drawString(
                m + 6,
                yinv(box_y + 38),
                f"Tempo Pecas: {fmt_num(tempo_pecas)} min | Avarias: {fmt_num(tempo_avarias)} min | Total refs: {fmt_num(tempo_operacional)} min",
            )
            c.drawString(m + 6, yinv(box_y + 51), f"Chapas cativadas: {clip_text(res_txt, table_w - 14, size=8)}")
            return 120

        c.setFont("Helvetica", 8.5)
        c.drawString(m, yinv(68), f"Cliente: {clip_text(cli_txt, table_w - 8, size=8.5)}")
        c.line(m, yinv(74), w - m, yinv(74))
        return 82

    def draw_table_header(y_top):
        c.setFillColor(colors.HexColor("#8b1e2d"))
        c.rect(m, yinv(y_top + header_h), table_w, header_h, stroke=1, fill=1)
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 8.5)
        xx = m
        for name, cw, align in cols:
            if align == "right":
                c.drawRightString(xx + cw - 3, yinv(y_top + 12), name)
            elif align == "center":
                c.drawCentredString(xx + cw / 2, yinv(y_top + 12), name)
            else:
                c.drawString(xx + 3, yinv(y_top + 12), name)
            xx += cw
        c.setFillColor(colors.black)

    idx = 0
    page = 1
    while idx < len(rows) or idx == 0:
        first_page = page == 1
        table_top = draw_page_header(page, first_page)
        draw_table_header(table_top)
        y_rows = table_top + header_h
        max_y = h - m - 24
        fit = max(1, int((max_y - y_rows) // row_h))
        count = min(fit, len(rows) - idx) if rows else 0

        c.setFont("Helvetica", 8)
        if count == 0:
            c.setFont("Helvetica-Oblique", 9)
            c.drawString(m + 4, yinv(y_rows + 16), "Sem peças registadas.")
        for local_i in range(count):
            p = rows[idx + local_i]
            y = y_rows + (local_i * row_h)
            fill = colors.HexColor("#f8fbff") if ((idx + local_i) % 2 == 0) else colors.HexColor("#fff0f2")
            c.setFillColor(fill)
            c.rect(m, yinv(y + row_h), table_w, row_h, stroke=1, fill=1)
            c.setFillColor(colors.black)

            obs = p.get("observacoes", p.get("Observacoes", p.get("Observações", ""))) or ""
            vals = [
                str(p.get("of_codigo", "") or p.get("opp", "")),
                str(p.get("ref_interna", "") or ""),
                str(p.get("ref_externa", "") or ""),
                str(p.get("material", "") or ""),
                fmt_num(p.get("espessura", "")),
                fmt_num(p.get("quantidade_pedida", 0)),
                str(obs),
            ]
            xx = m
            for (_, cw, align), val in zip(cols, vals):
                txt = clip_text(val, cw - 6, 8)
                if align == "right":
                    c.drawRightString(xx + cw - 3, yinv(y + 11), txt)
                elif align == "center":
                    c.drawCentredString(xx + cw / 2, yinv(y + 11), txt)
                else:
                    c.drawString(xx + 3, yinv(y + 11), txt)
                xx += cw

        idx += count
        if idx >= len(rows):
            break
        c.setFont("Helvetica-Oblique", 8)
        c.drawRightString(w - m, 26, "Continua na próxima página...")
        c.showPage()
        page += 1

    c.save()

def show_rejeitadas_hist(self):
    _ensure_configured()
    win = Toplevel(self.root)
    win.title("Histórico de Rejeições")
    tbl = ttk.Treeview(win, columns=("data", "operador", "encomenda", "material", "espessura", "ref_int", "ref_ext", "nok"), show="headings", height=12)
    headings = {
        "data": "Data",
        "operador": "Operador",
        "encomenda": "Encomenda",
        "material": "Material",
        "espessura": "Espessura",
        "ref_int": "Ref. Interna",
        "ref_ext": "Ref. Externa",
        "nok": "NOK",
    }
    for key, text in headings.items():
        tbl.heading(key, text=text)
        tbl.column(key, width=160)
    tbl.column("data", width=180)
    tbl.column("ref_int", width=180)
    tbl.column("ref_ext", width=220)
    wrap = ttk.Frame(win)
    wrap.pack(fill="both", expand=True, padx=10, pady=10)
    tbl.pack(in_=wrap, side="top", fill="both", expand=True)
    sbx = ttk.Scrollbar(wrap, orient="horizontal", command=tbl.xview)
    tbl.configure(xscrollcommand=sbx.set)
    sbx.pack(side="bottom", fill="x")
    for item in self.data.get("rejeitadas_hist", []):
        tbl.insert("", END, values=(
            item.get("data",""),
            item.get("operador",""),
            item.get("encomenda",""),
            item.get("material",""),
            item.get("espessura",""),
            item.get("ref_interna",""),
            item.get("ref_externa",""),
            item.get("nok",0),
        ))

