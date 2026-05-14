import os
import time
import os
import re
from datetime import date, datetime
from tkinter import BooleanVar, Button, Canvas, StringVar, Toplevel, messagebox
from tkinter import ttk

try:
    import customtkinter as ctk  # type: ignore
except Exception:
    ctk = None  # type: ignore


def _safe_date(raw):
    txt = str(raw or "").strip()
    if not txt:
        return None
    for fm in ("%Y-%m-%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(txt[:19], fm).date()
        except Exception:
            pass
    try:
        return datetime.fromisoformat(txt).date()
    except Exception:
        return None


def _bind_widget_click(widget, callback):
    try:
        widget.bind("<Button-1>", lambda _event: callback())
    except Exception:
        return


def _fmt_num(value):
    try:
        num = float(value)
    except Exception:
        return "0"
    if abs(num - int(num)) < 0.001:
        return str(int(num))
    return f"{num:.2f}"


def _norm_text(value):
    txt = str(value or "").strip().lower()
    for a, b in (
        ("Ã¡", "a"),
        ("Ã ", "a"),
        ("Ã¢", "a"),
        ("Ã£", "a"),
        ("Ã©", "e"),
        ("Ãª", "e"),
        ("Ã­", "i"),
        ("Ã³", "o"),
        ("Ã´", "o"),
        ("Ãµ", "o"),
        ("Ãº", "u"),
        ("Ã§", "c"),
    ):
        txt = txt.replace(a, b)
    return txt


def _parse_float(value, default=0.0):
    try:
        return float(str(value).replace(",", "."))
    except Exception:
        return float(default)


def _safe_dt(raw):
    txt = str(raw or "").strip()
    if not txt:
        return None
    try:
        return datetime.fromisoformat(txt.replace("Z", ""))
    except Exception:
        pass
    for fm in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(txt[:19], fm)
        except Exception:
            pass
    return None


def _diff_min(start_raw, end_raw=None):
    d1 = _safe_dt(start_raw)
    if d1 is None:
        return None
    d2 = _safe_dt(end_raw) if end_raw else datetime.now()
    if d2 is None:
        d2 = datetime.now()
    try:
        return max(0.0, (d2 - d1).total_seconds() / 60.0)
    except Exception:
        return None


def _extract_posto_from_info(info_txt):
    try:
        txt = str(info_txt or "")
    except Exception:
        return ""
    if not txt:
        return ""
    m = re.search(r"\[POSTO:\s*([^\]]+)\]", txt, flags=re.IGNORECASE)
    if not m:
        return ""
    return str(m.group(1) or "").strip()


def _extract_encomenda_from_text(txt):
    s = str(txt or "")
    if not s:
        return ""
    for pat in (
        r"\bencomenda\s*=\s*([A-Za-z0-9\-_]+)",
        r"\benc\s*=\s*([A-Za-z0-9\-_]+)",
        r"\b(BARCELBAL[0-9]{3,6})\b",
        r"\b([A-Z][A-Z0-9]{2,15}[0-9]{4,6})\b",
    ):
        try:
            m = re.search(pat, s, flags=re.IGNORECASE)
        except Exception:
            m = None
        if m:
            return str(m.group(1) or "").strip().upper()
    return ""


def _center_window_on_screen(win, width, height, min_w=None, min_h=None, margin=40):
    try:
        sw = int(win.winfo_screenwidth() or 1280)
        sh = int(win.winfo_screenheight() or 720)
        max_w = max(640, sw - margin)
        max_h = max(480, sh - margin)
        ww = min(int(width), max_w)
        wh = min(int(height), max_h)
        if min_w is not None or min_h is not None:
            try:
                win.minsize(int(min_w or ww), int(min_h or wh))
            except Exception:
                pass
        x = max(0, int((sw - ww) / 2))
        y = max(0, int((sh - wh) / 2))
        win.geometry(f"{ww}x{wh}+{x}+{y}")
        try:
            win.update_idletasks()
        except Exception:
            pass
    except Exception:
        pass


def _bring_window_front(win, parent=None, *, modal=False, keep_topmost=False):
    try:
        if parent is not None:
            win.transient(parent)
    except Exception:
        pass
    try:
        win.lift()
        win.focus_force()
    except Exception:
        pass
    if modal:
        try:
            win.grab_set()
        except Exception:
            pass
    try:
        win.attributes("-topmost", True)
        if not keep_topmost:
            win.after(250, lambda: win.winfo_exists() and win.attributes("-topmost", False))
    except Exception:
        pass


def _calc_encomenda_total(app, enc):
    total = _parse_float(enc.get("total", 0), 0)
    if total > 0:
        return total
    num_orc = str(enc.get("numero_orcamento", "") or "").strip()
    num_enc = str(enc.get("numero", "") or "").strip()
    if num_orc:
        for o in app.data.get("orcamentos", []):
            if str(o.get("numero", "") or "").strip() == num_orc:
                total_orc = _parse_float(o.get("total", 0), 0)
                if total_orc > 0:
                    return total_orc
                subtotal_orc = _parse_float(o.get("subtotal", 0), 0)
                iva_perc = _parse_float(o.get("iva_perc", 0), 0)
                if subtotal_orc > 0:
                    return subtotal_orc * (1.0 + (iva_perc / 100.0))
                break
    if num_enc:
        for o in app.data.get("orcamentos", []):
            if str(o.get("numero_encomenda", "") or "").strip() == num_enc:
                total_orc = _parse_float(o.get("total", 0), 0)
                if total_orc > 0:
                    return total_orc
                subtotal_orc = _parse_float(o.get("subtotal", 0), 0)
                iva_perc = _parse_float(o.get("iva_perc", 0), 0)
                if subtotal_orc > 0:
                    return subtotal_orc * (1.0 + (iva_perc / 100.0))
                break
    pecas = enc.get("pecas", [])
    if not isinstance(pecas, list) or not pecas:
        pecas = []
        for m in enc.get("materiais", []):
            for e in m.get("espessuras", []):
                for p in e.get("pecas", []):
                    pecas.append(p)
    refs_db = {}
    try:
        refs_db = app.data.get("orc_refs", {}) or {}
    except Exception:
        refs_db = {}
    orc_linhas = []
    try:
        num_orc = str(enc.get("numero_orcamento", "") or "").strip()
        if num_orc:
            for o in app.data.get("orcamentos", []):
                if str(o.get("numero", "") or "").strip() == num_orc:
                    orc_linhas = list(o.get("linhas", []) or [])
                    break
    except Exception:
        orc_linhas = []

    def _find_preco_ref(ref_int, ref_ext):
        ref_int = str(ref_int or "").strip()
        ref_ext = str(ref_ext or "").strip()
        if ref_ext and ref_ext in refs_db:
            return _parse_float((refs_db.get(ref_ext) or {}).get("preco_unit", 0), 0)
        if ref_int and ref_int in refs_db:
            return _parse_float((refs_db.get(ref_int) or {}).get("preco_unit", 0), 0)
        for l in orc_linhas:
            if not isinstance(l, dict):
                continue
            if ref_ext and str(l.get("ref_externa", "") or "").strip() == ref_ext:
                return _parse_float(l.get("preco_unit", 0), 0)
            if ref_int and str(l.get("ref_interna", "") or "").strip() == ref_int:
                return _parse_float(l.get("preco_unit", 0), 0)
        return 0.0

    calc = 0.0
    for p in pecas:
        if not isinstance(p, dict):
            continue
        qtd = _parse_float(p.get("quantidade_pedida", p.get("qtd", p.get("quantidade", 0))), 0)
        preco = _parse_float(
            p.get("preco", p.get("preco_unit", p.get("preco_unitario", p.get("valor_unitario", 0)))),
            0,
        )
        if preco <= 0:
            preco = _find_preco_ref(p.get("ref_interna", ""), p.get("ref_externa", ""))
        if qtd > 0 and preco > 0:
            calc += (qtd * preco)
    return calc


def _produto_preco_unit(app, prod):
    fn = globals().get("produto_preco_unitario")
    if callable(fn):
        try:
            v = _parse_float(fn(prod), 0)
            if v > 0:
                return v
        except Exception:
            pass
    for k in ("preco_unid", "preco_unit", "preco_compra_un", "p_compra"):
        v = _parse_float(prod.get(k, 0), 0)
        if v > 0:
            return v
    return 0.0


def _materia_preco_unit(app, mat):
    for k in ("preco_unid", "preco_unit", "preco_compra_un"):
        v = _parse_float((mat or {}).get(k, 0), 0)
        if v > 0:
            return v
    fn = globals().get("materia_preco_unitario")
    if callable(fn):
        try:
            v = _parse_float(fn(mat), 0)
            if v > 0:
                return v
        except Exception:
            pass
    return _parse_float(mat.get("p_compra", 0), 0)


def _ne_line_total(linha, qtd_ref=None):
    qtd = _parse_float(linha.get("qtd", 0), 0)
    preco = _parse_float(linha.get("preco", 0), 0)
    desc = _parse_float(linha.get("desconto", 0), 0)
    iva = _parse_float(linha.get("iva", 0), 0)
    if qtd <= 0:
        return 0.0
    bruto = qtd * max(0.0, preco)
    sem_iva = bruto * (1.0 - max(0.0, min(100.0, desc)) / 100.0)
    total_qtd = sem_iva * (1.0 + max(0.0, min(100.0, iva)) / 100.0)
    try:
        t_linha = _parse_float(linha.get("total", 0), 0)
        if t_linha > 0:
            total_qtd = t_linha
    except Exception:
        pass
    if qtd_ref is None:
        return total_qtd
    try:
        qtd_ref = max(0.0, float(qtd_ref))
    except Exception:
        qtd_ref = 0.0
    if qtd_ref <= 0:
        return 0.0
    return total_qtd * (qtd_ref / qtd)


def _compute_finance_metrics(app):
    produtos = list(app.data.get("produtos", []))
    materias = list(app.data.get("materiais", []))
    notas = list(app.data.get("notas_encomenda", []))

    val_produtos = 0.0
    for p in produtos:
        qtd = _parse_float(p.get("qty", p.get("quantidade", 0)), 0)
        if qtd <= 0:
            continue
        val_produtos += qtd * _produto_preco_unit(app, p)

    val_materias = 0.0
    for m in materias:
        qtd = _parse_float(m.get("quantidade", 0), 0)
        if qtd <= 0:
            continue
        val_materias += qtd * _materia_preco_unit(app, m)

    val_ne_aprovadas = 0.0
    hist_rows = []
    total_comprado_hist = 0.0
    for ne in notas:
        estado_txt = str(ne.get("estado", "") or "")
        est = _norm_text(estado_txt)
        if "aprov" in est:
            total_ne = _parse_float(ne.get("total", 0), 0)
            if total_ne <= 0:
                total_ne = sum(_ne_line_total(l) for l in list(ne.get("linhas", []) or []))
            val_ne_aprovadas += max(0.0, total_ne)

        ne_num = str(ne.get("numero", "") or "")
        forn = str(ne.get("fornecedor", "") or "")
        for l in list(ne.get("linhas", []) or []):
            qtd_tot = _parse_float(l.get("qtd", 0), 0)
            qtd_ent = _parse_float(l.get("qtd_entregue", 0), 0)
            entregue = bool(l.get("entregue") or l.get("_stock_in"))
            qtd_hist = qtd_ent if qtd_ent > 0 else (qtd_tot if entregue else 0.0)
            if qtd_hist <= 0:
                continue
            preco = _parse_float(l.get("preco", 0), 0)
            total_l = _ne_line_total(l, qtd_ref=qtd_hist)
            if total_l <= 0 and preco > 0:
                total_l = qtd_hist * preco
            data_l = (
                str(l.get("data_doc_entrega", "") or "").strip()
                or str(l.get("data_entrega_real", "") or "").strip()
                or str(ne.get("data_doc_ultima", "") or "").strip()
                or str(ne.get("data_entrega", "") or "").strip()
            )
            hist_rows.append(
                {
                    "data": data_l,
                    "ne": ne_num,
                    "fornecedor": forn,
                    "artigo": str(l.get("descricao", "") or l.get("ref", "") or ""),
                    "qtd": qtd_hist,
                    "preco": preco,
                    "total": total_l,
                    "estado": estado_txt,
                }
            )
            total_comprado_hist += total_l

    hist_rows.sort(key=lambda r: str(r.get("data", "")), reverse=True)
    return {
        "valor_stock_produtos": val_produtos,
        "valor_stock_materias": val_materias,
        "valor_ne_aprovadas": val_ne_aprovadas,
        "valor_total_stock": val_produtos + val_materias,
        "valor_total_comprado": total_comprado_hist,
        "compras_count": len(hist_rows),
        "compras_rows": hist_rows,
    }


def refresh_dashboard_finance_window(app):
    win = getattr(app, "dashboard_fin_win", None)
    if win is None:
        return
    try:
        if not win.winfo_exists():
            app.dashboard_fin_win = None
            return
    except Exception:
        app.dashboard_fin_win = None
        return

    data = _compute_finance_metrics(app)
    vars_map = getattr(app, "dashboard_fin_vars", {})
    mapping = {
        "valor_stock_produtos": f"{data['valor_stock_produtos']:.2f} EUR",
        "valor_stock_materias": f"{data['valor_stock_materias']:.2f} EUR",
        "valor_ne_aprovadas": f"{data['valor_ne_aprovadas']:.2f} EUR",
        "valor_total_stock": f"{data['valor_total_stock']:.2f} EUR",
        "valor_total_comprado": f"{data['valor_total_comprado']:.2f} EUR",
        "compras_count": str(data["compras_count"]),
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    for k, v in mapping.items():
        var = vars_map.get(k)
        if var is not None:
            try:
                var.set(v)
            except Exception:
                pass

    rows = data.get("compras_rows", [])
    rows_host = getattr(app, "dashboard_fin_rows_host", None)
    if rows_host is not None and ctk is not None:
        try:
            for child in rows_host.winfo_children():
                child.destroy()
            col_w = (0.09, 0.10, 0.21, 0.29, 0.06, 0.08, 0.08, 0.09)
            for idx, r in enumerate(rows[:250]):
                bg = "#eef4ff" if idx % 2 else "#f8fbff"
                row = ctk.CTkFrame(rows_host, fg_color=bg, corner_radius=6, height=32)
                row.pack(fill="x", padx=4, pady=1)
                row.pack_propagate(False)
                vals = (
                    r.get("data", ""),
                    r.get("ne", ""),
                    r.get("fornecedor", ""),
                    r.get("artigo", ""),
                    _fmt_num(r.get("qtd", 0)),
                    f"{_parse_float(r.get('preco', 0), 0):.2f}",
                    f"{_parse_float(r.get('total', 0), 0):.2f}",
                    r.get("estado", ""),
                )
                for cidx, (txt, rw) in enumerate(zip(vals, col_w)):
                    anchor = "w" if cidx < 4 else ("e" if cidx in (4, 5, 6) else "center")
                    lbl = ctk.CTkLabel(row, text=str(txt), anchor=anchor, font=("Segoe UI", 13), text_color="#0f172a")
                    lbl.place(relx=sum(col_w[:cidx]), rely=0.0, relwidth=rw, relheight=1.0)
        except Exception:
            pass

    tree = getattr(app, "dashboard_fin_tree", None)
    if tree is not None:
        try:
            for iid in tree.get_children():
                tree.delete(iid)
            for idx, r in enumerate(rows[:250]):
                tag = "odd" if idx % 2 else "even"
                tree.insert(
                    "",
                    "end",
                    values=(
                        r.get("data", ""),
                        r.get("ne", ""),
                        r.get("fornecedor", ""),
                        r.get("artigo", ""),
                        _fmt_num(r.get("qtd", 0)),
                        f"{_parse_float(r.get('preco', 0), 0):.2f}",
                        f"{_parse_float(r.get('total', 0), 0):.2f}",
                        r.get("estado", ""),
                    ),
                    tags=(tag,),
                )
        except Exception:
            pass


def open_dashboard_finance_window(app):
    use_custom = ctk is not None and os.environ.get("USE_CUSTOM_DASH_FIN", "1") != "0"
    old = getattr(app, "dashboard_fin_win", None)
    try:
        if old is not None and old.winfo_exists():
            old.lift()
            old.focus_force()
            refresh_dashboard_finance_window(app)
            return
    except Exception:
        pass

    Win = ctk.CTkToplevel if use_custom else Toplevel
    Frame = ctk.CTkFrame if use_custom else ttk.Frame
    Label = ctk.CTkLabel if use_custom else ttk.Label
    Btn = ctk.CTkButton if use_custom else ttk.Button

    win = Win(app.root)
    app.dashboard_fin_win = win
    try:
        win.title("Dashboard Financeiro - Stock e Compras")
        win.transient(app.root)
        win.geometry("1380x860")
        win.minsize(1160, 740)
        if use_custom:
            win.configure(fg_color="#eef2f7")
        _bring_window_front(win, app.root, modal=True, keep_topmost=False)
    except Exception:
        pass

    app.dashboard_fin_vars = {k: StringVar(value="-") for k in (
        "valor_stock_produtos",
        "valor_stock_materias",
        "valor_ne_aprovadas",
        "valor_total_stock",
        "valor_total_comprado",
        "compras_count",
        "updated_at",
    )}

    top = Frame(win, fg_color="transparent") if use_custom else Frame(win)
    top.pack(fill="x", padx=14, pady=(12, 8))
    Label(top, text="Dashboard Financeiro", font=("Segoe UI", 24, "bold") if use_custom else ("Segoe UI", 16, "bold"), text_color="#0f172a" if use_custom else None).pack(side="left")
    Label(top, text="Stock | Compras | Notas de Encomenda", font=("Segoe UI", 13, "bold") if use_custom else ("Segoe UI", 10, "bold"), text_color="#475569" if use_custom else None).pack(side="left", padx=(12, 0), pady=(6, 0))
    Label(top, textvariable=app.dashboard_fin_vars["updated_at"], text_color="#64748b" if use_custom else None, font=("Segoe UI", 12, "bold") if use_custom else None).pack(side="right")

    actions = Frame(win, fg_color="#ffffff" if use_custom else "transparent", corner_radius=10 if use_custom else None, border_width=1 if use_custom else 0, border_color="#d7dee9" if use_custom else None) if use_custom else Frame(win)
    actions.pack(fill="x", padx=14, pady=(0, 10))
    Btn(actions, text="Atualizar", command=lambda: refresh_dashboard_finance_window(app), width=130 if use_custom else None).pack(side="left", padx=6, pady=8)
    Btn(actions, text="Fechar", command=win.destroy, width=120 if use_custom else None).pack(side="right", padx=6, pady=8)

    cards_wrap = Frame(win, fg_color="transparent") if use_custom else Frame(win)
    cards_wrap.pack(fill="x", padx=14, pady=(0, 8))
    for i in range(3):
        try:
            cards_wrap.grid_columnconfigure(i, weight=1)
        except Exception:
            pass
    cards = [
        ("Stock Produtos", "valor_stock_produtos"),
        ("Stock Materia-Prima", "valor_stock_materias"),
        ("NE Aprovadas", "valor_ne_aprovadas"),
        ("Stock Total", "valor_total_stock"),
        ("Historico Comprado", "valor_total_comprado"),
        ("Movimentos Compra", "compras_count"),
    ]
    for idx, (title, key) in enumerate(cards):
        r, c = divmod(idx, 3)
        card = Frame(cards_wrap, fg_color="#ffffff" if use_custom else None, corner_radius=12 if use_custom else None, border_width=1 if use_custom else 0, border_color="#dce3ee" if use_custom else None) if use_custom else ttk.LabelFrame(cards_wrap, text=title)
        card.grid(row=r, column=c, padx=6, pady=6, sticky="nsew")
        if use_custom:
            Label(card, text=title, font=("Segoe UI", 12, "bold"), text_color="#1e293b").pack(anchor="w", padx=10, pady=(8, 2))
            Label(card, textvariable=app.dashboard_fin_vars[key], font=("Segoe UI", 18, "bold"), text_color="#0f172a").pack(anchor="w", padx=10, pady=(0, 8))
        else:
            ttk.Label(card, textvariable=app.dashboard_fin_vars[key], font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=10, pady=8)

    hist_box = Frame(win, fg_color="#ffffff" if use_custom else None, corner_radius=12 if use_custom else None, border_width=1 if use_custom else 0, border_color="#dce3ee" if use_custom else None) if use_custom else ttk.LabelFrame(win, text="Historico de Compras (Notas de Encomenda)")
    hist_box.pack(fill="both", expand=True, padx=14, pady=(0, 10))
    if use_custom:
        Label(hist_box, text="Historico de Compras (Notas de Encomenda)", font=("Segoe UI", 13, "bold")).pack(anchor="w", padx=10, pady=(8, 4))

    app.dashboard_fin_tree = None
    app.dashboard_fin_rows_host = None
    if use_custom and ctk is not None:
        hdr = ctk.CTkFrame(hist_box, fg_color="#0b1f4d", corner_radius=8, height=34)
        hdr.pack(fill="x", padx=8, pady=(2, 4))
        hdr.pack_propagate(False)
        headers = ("Data", "NE", "Fornecedor", "Artigo/Descricao", "Qtd", "Preco", "Total", "Estado")
        col_w = (0.09, 0.10, 0.21, 0.29, 0.06, 0.08, 0.08, 0.09)
        for cidx, (txt, rw) in enumerate(zip(headers, col_w)):
            anchor = "w" if cidx < 4 else ("e" if cidx in (4, 5, 6) else "center")
            lbl = ctk.CTkLabel(hdr, text=txt, anchor=anchor, font=("Segoe UI", 13, "bold"), text_color="#ffffff")
            lbl.place(relx=sum(col_w[:cidx]), rely=0.0, relwidth=rw, relheight=1.0)
        body = ctk.CTkScrollableFrame(hist_box, fg_color="#ffffff", corner_radius=8)
        body.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        app.dashboard_fin_rows_host = body
    else:
        cols = ("data", "ne", "fornecedor", "artigo", "qtd", "preco", "total", "estado")
        tree = ttk.Treeview(hist_box, columns=cols, show="headings", height=18)
        app.dashboard_fin_tree = tree
        head = {
            "data": "Data",
            "ne": "NE",
            "fornecedor": "Fornecedor",
            "artigo": "Artigo/Descricao",
            "qtd": "Qtd",
            "preco": "Preco",
            "total": "Total",
            "estado": "Estado",
        }
        widths = {"data": 100, "ne": 120, "fornecedor": 220, "artigo": 320, "qtd": 70, "preco": 90, "total": 100, "estado": 130}
        for c in cols:
            an = "center" if c in ("data", "ne", "qtd", "preco", "total", "estado") else "w"
            tree.heading(c, text=head[c])
            tree.column(c, width=widths[c], anchor=an, stretch=(c in ("fornecedor", "artigo")))
        tree.tag_configure("even", background="#f8fbff")
        tree.tag_configure("odd", background="#eef4ff")
        sy = ttk.Scrollbar(hist_box, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sy.set)
        tree.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=(4, 8))
        sy.pack(side="right", fill="y", padx=(0, 8), pady=(4, 8))

    def _on_close():
        try:
            try:
                win.grab_release()
            except Exception:
                pass
            win.destroy()
        finally:
            app.dashboard_fin_win = None
            app.dashboard_fin_tree = None
            app.dashboard_fin_rows_host = None

    try:
        win.protocol("WM_DELETE_WINDOW", _on_close)
    except Exception:
        pass

    refresh_dashboard_finance_window(app)


def _compute_production_pulse_metrics(app, period_days=7, year_filter=None):
    try:
        if hasattr(app, "refresh_runtime_impulse_data"):
            app.refresh_runtime_impulse_data(cleanup_orphans=True)
    except Exception:
        pass
    eventos = list(app.data.get("op_eventos", []) or [])
    paragens = list(app.data.get("op_paragens", []) or [])
    encomendas = list(app.data.get("encomendas", []) or [])
    qualidade_rows = list(app.data.get("qualidade", []) or [])
    posto_map = dict(app.data.get("operador_posto_map", {}) or {})
    op_plan_map_raw = dict(app.data.get("tempos_operacao_planeada_min", {}) or {})
    op_plan_map = {str(k or "").strip().lower(): max(0.0, _parse_float(v, 0)) for k, v in op_plan_map_raw.items()}

    def _piece_qty(px):
        return max(
            0.0,
            _parse_float(
                px.get("quantidade_pedida", px.get("qtd", px.get("quantidade", 0))),
                0,
            ),
        )

    piece_meta = {}
    valid_encs = set()
    valid_piece_keys = set()
    for enc in encomendas:
        enc_num = str(enc.get("numero", "") or "").strip()
        if enc_num:
            valid_encs.add(enc_num)
        for m in list(enc.get("materiais", []) or []):
            mat = str(m.get("material", "") or "").strip()
            for esp in list(m.get("espessuras", []) or []):
                esp_txt = str(esp.get("espessura", "") or "").strip()
                laser_plan = max(0.0, _parse_float(esp.get("tempo_min", 0), 0))
                for p in list(esp.get("pecas", []) or []):
                    p_id = str(p.get("id", "") or "").strip()
                    if not p_id:
                        continue
                    valid_piece_keys.add((enc_num, p_id))
                    piece_meta[(enc_num, p_id)] = {
                        "ref_interna": str(p.get("ref_interna", "") or "").strip(),
                        "material": str(p.get("material", "") or mat),
                        "espessura": str(p.get("espessura", "") or esp_txt),
                        "qtd": _piece_qty(p),
                        "laser_plan": laser_plan,
                    }

    cutoff = None
    try:
        pd = int(period_days or 0)
    except Exception:
        pd = 0
    if pd > 0:
        try:
            from datetime import timedelta
            cutoff = date.today() - timedelta(days=pd - 1)
        except Exception:
            cutoff = None
    try:
        yf = int(str(year_filter or "").strip()) if str(year_filter or "").strip().isdigit() else None
    except Exception:
        yf = None

    def _match_year(dt_obj):
        if yf is None:
            return True
        if dt_obj is None:
            return True
        try:
            return int(dt_obj.year) == int(yf)
        except Exception:
            return True

    month_ref = int(date.today().month)
    year_ref = int(yf) if yf is not None else int(date.today().year)

    def _match_month_scope(dt_obj):
        if dt_obj is None:
            return False
        try:
            return int(dt_obj.year) == int(year_ref) and int(dt_obj.month) == int(month_ref)
        except Exception:
            return False

    def _match_period_scope(dt_obj):
        if dt_obj is None:
            return False
        if cutoff is not None and dt_obj < cutoff:
            return False
        return _match_year(dt_obj)

    start_ops = 0
    finish_ops = 0
    qtd_ok = 0.0
    qtd_nok = 0.0
    ultimos_eventos = []
    eventos_validos = []
    for ev in eventos:
        d_ev = _safe_date(ev.get("created_at"))
        if cutoff is not None and d_ev is not None and d_ev < cutoff:
            continue
        if not _match_year(d_ev):
            continue
        enc_num_ev = str(ev.get("encomenda_numero", "") or "").strip()
        peca_id_ev = str(ev.get("peca_id", "") or "").strip()
        if enc_num_ev and enc_num_ev not in valid_encs:
            continue
        if peca_id_ev and enc_num_ev and (enc_num_ev, peca_id_ev) not in valid_piece_keys:
            continue
        evento = str(ev.get("evento", "") or "").strip().upper()
        if evento == "START_OP":
            start_ops += 1
        elif evento == "FINISH_OP":
            finish_ops += 1
        # NOTE: qtd_ok/qtd_nok nao deve ser somada por evento de operacao,
        # porque o mesmo registo de peca pode gerar varios FINISH_OP (1 por operacao).
        # A fonte correta para qualidade e a tabela/colecao de registos de qualidade.
        ultimos_eventos.append(
            {
                "created_at": str(ev.get("created_at", "") or ""),
                "evento": str(ev.get("evento", "") or ""),
                "encomenda": str(ev.get("encomenda_numero", "") or ""),
                "peca": str(ev.get("ref_interna", "") or ev.get("peca_id", "") or ""),
                "operador": str(ev.get("operador", "") or ""),
                "posto": _extract_posto_from_info(ev.get("info", "")),
                "info": str(ev.get("info", "") or ""),
            }
        )
        eventos_validos.append(
            {
                "dt": _safe_dt(ev.get("created_at")),
                "created_at": str(ev.get("created_at", "") or ""),
                "evento": evento,
                "encomenda": str(ev.get("encomenda_numero", "") or "").strip(),
                "peca_id": str(ev.get("peca_id", "") or "").strip(),
                "ref_interna": str(ev.get("ref_interna", "") or "").strip(),
                "operacao": str(ev.get("operacao", "") or "").strip(),
                "operador": str(ev.get("operador", "") or "").strip(),
            }
        )

    paragem_piece_totals = {}
    paragem_order_totals = {}
    avaria_incidents = {}
    _pulse_now_iso = datetime.now().isoformat(sep=" ", timespec="seconds")
    for p in paragens:
        d_pa = _safe_date(p.get("created_at"))
        if not _match_period_scope(d_pa):
            continue
        enc_num_pa = str(p.get("encomenda_numero", "") or "").strip()
        peca_id_pa = str(p.get("peca_id", "") or "").strip()
        if enc_num_pa and enc_num_pa not in valid_encs:
            continue
        if peca_id_pa and enc_num_pa and (enc_num_pa, peca_id_pa) not in valid_piece_keys:
            continue
        origem = str(p.get("origem", "") or "").strip().upper()
        causa_norm = _norm_text(p.get("causa", ""))
        mins = max(0.0, _parse_float(p.get("duracao_min", 0), 0))
        if origem != "AVARIA" and "avaria" not in causa_norm:
            continue
        if mins <= 0 and str(p.get("estado", "") or "").strip().upper() == "ABERTA":
            mins = max(0.0, _diff_min(p.get("created_at"), _pulse_now_iso) or 0.0)
        if mins <= 0:
            continue
        incident_key = (
            enc_num_pa or "-",
            str(p.get("operador", "") or "").strip() or "-",
            str(p.get("causa", "") or "").strip() or "Sem causa",
            str(p.get("created_at", "") or "").strip(),
            str(p.get("fechada_at", "") or "").strip(),
            str(p.get("estado", "") or "").strip().upper(),
        )
        row = avaria_incidents.setdefault(
            incident_key,
            {
                "encomenda": enc_num_pa or "-",
                "operador": str(p.get("operador", "") or "").strip() or "-",
                "causa": str(p.get("causa", "") or "").strip() or "Sem causa",
                "created_at": str(p.get("created_at", "") or "").strip(),
                "fechada_at": str(p.get("fechada_at", "") or "").strip(),
                "estado": str(p.get("estado", "") or "").strip().upper(),
                "mins": 0.0,
                "piece_keys": set(),
            },
        )
        row["mins"] = max(_parse_float(row.get("mins", 0), 0), mins)
        if enc_num_pa and peca_id_pa:
            row["piece_keys"].add((enc_num_pa, peca_id_pa))

    def _pulse_batch_token(ts_txt):
        dt_obj = _safe_dt(ts_txt)
        if dt_obj is not None:
            try:
                return dt_obj.strftime("%Y-%m-%d %H:%M")
            except Exception:
                pass
        return str(ts_txt or "").strip()[:16]

    def _aggregate_laser_pulse_rows(rows):
        grouped = {}
        out = []
        for row in list(rows or []):
            if not bool(row.get("laser_batch")):
                clean = dict(row)
                ref_txt = str(clean.get("ref_interna", "") or clean.get("peca", "") or "").strip()
                clean["ref_interna"] = ref_txt
                clean["refs"] = [ref_txt] if ref_txt else []
                clean.pop("laser_batch", None)
                clean.pop("batch_token", None)
                clean.pop("start_at", None)
                clean.pop("finish_at", None)
                clean.pop("material", None)
                clean.pop("espessura", None)
                out.append(clean)
                continue
            key = (
                str(row.get("encomenda", "") or "").strip(),
                str(row.get("material", "") or "").strip(),
                str(row.get("espessura", "") or "").strip(),
                str(row.get("operacao", "") or "").strip(),
                str(row.get("operador", "") or "").strip(),
                str(row.get("posto", "") or "").strip(),
                str(row.get("batch_token", "") or "").strip(),
            )
            grp = grouped.setdefault(
                key,
                {
                    "encomenda": str(row.get("encomenda", "") or "").strip(),
                    "operacao": str(row.get("operacao", "") or "").strip(),
                    "operador": str(row.get("operador", "") or "").strip(),
                    "posto": str(row.get("posto", "") or "").strip(),
                    "material": str(row.get("material", "") or "").strip(),
                    "espessura": str(row.get("espessura", "") or "").strip(),
                    "plan_min": max(0.0, _parse_float(row.get("plan_min", 0), 0)),
                    "start_at": str(row.get("start_at", "") or "").strip(),
                    "finish_at": str(row.get("finish_at", "") or "").strip(),
                    "created_at": str(row.get("created_at", "") or "").strip(),
                    "is_live": bool(row.get("is_live")),
                    "elapsed_live_max": max(0.0, _parse_float(row.get("elapsed_min", 0), 0)),
                    "refs": [],
                    "count": 0,
                },
            )
            grp["count"] += 1
            grp["is_live"] = bool(grp.get("is_live")) or bool(row.get("is_live"))
            ref_txt = str(row.get("peca", "") or "").strip()
            if ref_txt and ref_txt not in grp["refs"]:
                grp["refs"].append(ref_txt)
            grp["plan_min"] = max(grp["plan_min"], max(0.0, _parse_float(row.get("plan_min", 0), 0)))
            grp["elapsed_live_max"] = max(grp.get("elapsed_live_max", 0.0), max(0.0, _parse_float(row.get("elapsed_min", 0), 0)))
            start_txt = str(row.get("start_at", "") or "").strip()
            finish_txt = str(row.get("finish_at", "") or "").strip()
            created_txt = str(row.get("created_at", "") or "").strip()
            if start_txt and (not grp["start_at"] or (_safe_dt(start_txt) and _safe_dt(grp["start_at"]) and _safe_dt(start_txt) < _safe_dt(grp["start_at"]))):
                grp["start_at"] = start_txt
            if finish_txt and (not grp["finish_at"] or (_safe_dt(finish_txt) and _safe_dt(grp["finish_at"]) and _safe_dt(finish_txt) > _safe_dt(grp["finish_at"]))):
                grp["finish_at"] = finish_txt
            if created_txt and (not grp["created_at"] or created_txt > grp["created_at"]):
                grp["created_at"] = created_txt

        for grp in grouped.values():
            elapsed_min = 0.0
            if bool(grp.get("is_live")):
                elapsed_min = max(0.0, _parse_float(grp.get("elapsed_live_max", 0), 0))
            else:
                ini_dt = _safe_dt(grp.get("start_at"))
                fim_dt = _safe_dt(grp.get("finish_at"))
                if ini_dt is not None and fim_dt is not None:
                    try:
                        elapsed_min = max(0.0, (fim_dt - ini_dt).total_seconds() / 60.0)
                    except Exception:
                        elapsed_min = 0.0
            plan_min = max(0.0, _parse_float(grp.get("plan_min", 0), 0))
            delta_min = elapsed_min - plan_min if plan_min > 0 else 0.0
            refs = list(grp.get("refs", []) or [])
            qtd_refs = int(grp.get("count", 0) or len(refs) or 1)
            ref_base = refs[0] if refs else "-"
            ref_txt = f"{ref_base} (+{qtd_refs - 1})" if qtd_refs > 1 else ref_base
            out.append(
                {
                    "created_at": str(grp.get("created_at", "") or ""),
                    "encomenda": str(grp.get("encomenda", "") or ""),
                    "peca": f"Lote Laser | {ref_txt}",
                    "operacao": str(grp.get("operacao", "") or "Corte Laser"),
                    "operador": str(grp.get("operador", "") or ""),
                    "posto": str(grp.get("posto", "") or ""),
                    "elapsed_min": round(elapsed_min, 1),
                    "plan_min": round(plan_min, 1),
                    "delta_min": round(delta_min, 1),
                    "fora": bool(plan_min > 0 and delta_min > 0.01),
                    "ref_interna": refs[0] if refs else "",
                    "refs": refs,
                }
            )
        out.sort(key=lambda r: str(r.get("created_at", "") or ""), reverse=True)
        return out

    historico_tempo_raw = []
    open_ops = {}
    for ev in sorted(eventos_validos, key=lambda r: (r.get("dt") or datetime.min, r.get("created_at", ""))):
        evento = str(ev.get("evento", "") or "").strip().upper()
        enc_num = str(ev.get("encomenda", "") or "").strip()
        peca_id = str(ev.get("peca_id", "") or "").strip()
        operacao = str(ev.get("operacao", "") or "").strip()
        op_norm = _norm_text(operacao)
        if not enc_num or not peca_id or not op_norm:
            continue
        key = (enc_num, peca_id, op_norm)
        if evento == "START_OP":
            open_ops.setdefault(key, []).append(
                {
                    "dt": ev.get("dt"),
                    "operador": str(ev.get("operador", "") or "").strip(),
                    "created_at": ev.get("created_at", ""),
                }
            )
            continue
        if evento != "FINISH_OP":
            continue
        stack = list(open_ops.get(key, []) or [])
        if not stack:
            continue
        fin_operador = str(ev.get("operador", "") or "").strip()
        start_row = None
        if fin_operador:
            for i in range(len(stack) - 1, -1, -1):
                if str(stack[i].get("operador", "") or "").strip() == fin_operador:
                    start_row = stack.pop(i)
                    break
        if start_row is None:
            start_row = stack.pop(-1)
        open_ops[key] = stack
        ini_dt = start_row.get("dt")
        fim_dt = ev.get("dt")
        if ini_dt is None or fim_dt is None:
            continue
        try:
            elapsed_min = max(0.0, (fim_dt - ini_dt).total_seconds() / 60.0)
        except Exception:
            continue
        meta = piece_meta.get((enc_num, peca_id), {}) or {}
        qtd_piece = max(0.0, _parse_float(meta.get("qtd", 0), 0))
        plan_min = 0.0
        if "laser" in op_norm:
            plan_min = max(0.0, _parse_float(meta.get("laser_plan", 0), 0))
        else:
            plan_per_piece = max(0.0, _parse_float(op_plan_map.get(op_norm, 0), 0))
            if plan_per_piece > 0 and qtd_piece > 0:
                plan_min = plan_per_piece * qtd_piece
        delta_min = elapsed_min - plan_min if plan_min > 0 else 0.0
        historico_tempo_raw.append(
            {
                "created_at": str(ev.get("created_at", "") or ""),
                "encomenda": enc_num,
                "peca": str(meta.get("ref_interna", "") or ev.get("ref_interna", "") or peca_id),
                "operacao": operacao or "-",
                "operador": fin_operador or str(start_row.get("operador", "") or ""),
                "posto": str(posto_map.get(fin_operador or str(start_row.get("operador", "") or "").strip(), "") or ""),
                "material": str(meta.get("material", "") or ""),
                "espessura": str(meta.get("espessura", "") or ""),
                "laser_batch": bool("laser" in op_norm and plan_min > 0),
                "batch_token": _pulse_batch_token(start_row.get("created_at", "")),
                "start_at": str(start_row.get("created_at", "") or ""),
                "finish_at": str(ev.get("created_at", "") or ""),
                "is_live": False,
                "elapsed_min": round(elapsed_min, 1),
                "plan_min": round(plan_min, 1),
                "delta_min": round(delta_min, 1),
                "fora": bool(plan_min > 0 and delta_min > 0.01),
            }
        )

    # Qualidade/quantidades sempre calculadas pelos registos de qualidade.
    for q in qualidade_rows:
        d_q = _safe_date(q.get("data"))
        if not _match_month_scope(d_q):
            continue
        qtd_ok += max(0.0, _parse_float(q.get("ok", 0), 0))
        qtd_nok += max(0.0, _parse_float(q.get("nok", 0), 0))

    for inc in list(avaria_incidents.values()):
        mins = max(0.0, _parse_float(inc.get("mins", 0), 0))
        if mins <= 0:
            continue
        enc_num_pa = str(inc.get("encomenda", "") or "").strip()
        paragem_order_totals[enc_num_pa] = paragem_order_totals.get(enc_num_pa, 0.0) + mins
        piece_keys = list(inc.get("piece_keys", set()) or [])
        if len(piece_keys) == 1:
            only_key = piece_keys[0]
            paragem_piece_totals[only_key] = paragem_piece_totals.get(only_key, 0.0) + mins

    down_min = 0.0
    causa_count = {}
    causa_time = {}
    causa_last_dt = {}
    avaria_count = 0
    avaria_min = 0.0
    now_iso = _pulse_now_iso
    for inc in list(avaria_incidents.values()):
        mins = max(0.0, _parse_float(inc.get("mins", 0), 0))
        if mins <= 0:
            continue
        causa = str(inc.get("causa", "") or "").strip() or "Sem causa"
        encomenda = str(inc.get("encomenda", "") or "").strip() or "-"
        operador = str(inc.get("operador", "") or "").strip() or "-"
        down_min += mins
        k = (causa, encomenda, operador)
        causa_count[k] = causa_count.get(k, 0) + 1
        causa_time[k] = causa_time.get(k, 0.0) + mins
        dtp = _safe_dt(inc.get("created_at"))
        prev_dt = causa_last_dt.get(k)
        if dtp is not None and (prev_dt is None or dtp > prev_dt):
            causa_last_dt[k] = dtp
        avaria_count += 1
        avaria_min += mins

    andon_prod = 0
    andon_setup = 0
    andon_wait = 0
    andon_stop = 0
    interrompidas = []
    pecas_tempo_raw = []
    pecas_em_curso = 0
    pecas_fora_tempo = 0
    for e in encomendas:
        d_enc = _safe_date(e.get("data_criacao") or e.get("inicio_producao") or e.get("data_entrega"))
        if not _match_year(d_enc):
            continue
        est = _norm_text(e.get("estado", ""))
        if "concl" in est or "cancel" in est:
            continue
        if "produ" in est:
            andon_prod += 1
        elif "paus" in est or "interromp" in est:
            andon_stop += 1
        elif "prepar" in est:
            andon_setup += 1
        elif "parad" in est:
            andon_stop += 1
        else:
            andon_wait += 1
        for m in list(e.get("materiais", []) or []):
            for esp in list(m.get("espessuras", []) or []):
                esp_plan_min = max(0.0, _parse_float(esp.get("tempo_min", 0), 0))
                esp_pecas = list(esp.get("pecas", []) or [])
                esp_qtd_total = 0.0
                for px in esp_pecas:
                    esp_qtd_total += _piece_qty(px)
                for p in list(esp.get("pecas", []) or []):
                    p_est = _norm_text(p.get("estado", ""))
                    fluxo = list(p.get("operacoes_fluxo", []) or [])
                    op_em_curso = ""
                    op_user = ""
                    for op in fluxo:
                        op_st = _norm_text(op.get("estado", ""))
                        if "produ" in op_st:
                            op_em_curso = str(op.get("nome", "") or "").strip()
                            op_user = str(op.get("user", "") or "").strip()
                            break
                    avaria_ativa = bool(p.get("avaria_ativa"))
                    running = bool(op_em_curso) or ("produ" in p_est and "interromp" not in p_est and "concl" not in p_est and "paus" not in p_est and "avari" not in p_est)
                    if running or avaria_ativa:
                        qtd_piece = _piece_qty(p)
                        plan_min = 0.0
                        op_for_plan = str(op_em_curso or "").strip()
                        if not op_for_plan:
                            for op in fluxo:
                                if "concl" not in _norm_text(op.get("estado", "")):
                                    op_for_plan = str(op.get("nome", "") or "").strip()
                                    break
                        op_norm = _norm_text(op_for_plan)
                        is_laser = "laser" in op_norm
                        if is_laser and esp_plan_min > 0:
                            plan_min = esp_plan_min
                        elif op_norm:
                            # Preparacao para proximos postos (quinagem, soldadura, etc.).
                            # Quando configurado, tempo por operacao e por peca.
                            plan_per_piece = max(0.0, _parse_float(op_plan_map.get(op_norm, 0), 0))
                            if plan_per_piece > 0 and qtd_piece > 0:
                                plan_min = plan_per_piece * qtd_piece
                        elapsed_base = max(0.0, _parse_float(p.get("tempo_producao_min", 0), 0))
                        elapsed_open = max(0.0, _diff_min(p.get("inicio_producao"), now_iso) or 0.0)
                        elapsed_min = elapsed_base + elapsed_open
                        if avaria_ativa:
                            elapsed_min += max(0.0, _diff_min(p.get("avaria_inicio_ts"), now_iso) or 0.0)
                        delta_min = elapsed_min - plan_min if plan_min > 0 else 0.0
                        fora = bool(plan_min > 0 and delta_min > 0.01)
                        if fora:
                            pecas_fora_tempo += 1
                        pecas_tempo_raw.append(
                            {
                                "encomenda": str(e.get("numero", "") or ""),
                                "peca": str(p.get("ref_interna", "") or p.get("id", "") or ""),
                                "operacao": ("Avaria" if avaria_ativa else (op_for_plan or "Em curso")),
                                "operador": str(op_user or ""),
                                "posto": str(posto_map.get(op_user, "") or ""),
                                "material": str(p.get("material", "") or m.get("material", "") or ""),
                                "espessura": str(p.get("espessura", "") or esp.get("espessura", "") or ""),
                                "laser_batch": bool(is_laser and not avaria_ativa and plan_min > 0),
                                "batch_token": _pulse_batch_token(p.get("inicio_producao", "")),
                                "start_at": str(p.get("inicio_producao", "") or ""),
                                "finish_at": str(now_iso),
                                "is_live": True,
                                "elapsed_min": round(elapsed_min, 1),
                                "plan_min": round(plan_min, 1),
                                "delta_min": round(delta_min, 1),
                                "fora": fora,
                            }
                        )
                    if ("interromp" not in p_est) and ("paus" not in p_est) and (not bool(p.get("avaria_ativa"))):
                        continue
                    interrompidas.append(
                        {
                            "encomenda": str(e.get("numero", "") or ""),
                            "peca": str(p.get("ref_interna", "") or p.get("id", "") or ""),
                            "operador": str((p.get("hist", [{}])[-1].get("user", "") if isinstance(p.get("hist", []), list) and p.get("hist") else "") or ""),
                            "posto": str(
                                posto_map.get(
                                    str((p.get("hist", [{}])[-1].get("user", "") if isinstance(p.get("hist", []), list) and p.get("hist") else "") or "").strip(),
                                    "",
                                )
                                or ""
                            ),
                            "motivo": str(p.get("avaria_motivo", "") or p.get("interrupcao_peca_motivo", "") or ""),
                            "ts": str(p.get("interrupcao_peca_ts", "") or ""),
                        }
                    )

    historico_tempo = _aggregate_laser_pulse_rows(historico_tempo_raw)
    pecas_tempo = _aggregate_laser_pulse_rows(pecas_tempo_raw)

    live_plan_total = 0.0
    live_elapsed_total = 0.0
    live_rows_total = len(list(pecas_tempo or []))
    live_avaria_rows = 0
    for row in list(pecas_tempo or []):
        plan_v = max(0.0, _parse_float(row.get("plan_min", 0), 0))
        elapsed_v = max(0.0, _parse_float(row.get("elapsed_min", 0), 0))
        op_norm = _norm_text(row.get("operacao", ""))
        if "avaria" in op_norm:
            live_avaria_rows += 1
        if plan_v <= 0 or elapsed_v <= 0:
            continue
        live_plan_total += plan_v
        live_elapsed_total += elapsed_v

    if live_rows_total > 0:
        disp = max(0.0, ((live_rows_total - live_avaria_rows) / max(1, live_rows_total)) * 100.0)
    else:
        disp = 100.0
    if live_plan_total > 0 and live_elapsed_total > 0:
        perf = (live_plan_total / live_elapsed_total) * 100.0
    elif live_rows_total > 0:
        perf = 100.0
    else:
        perf = 100.0
    perf = max(0.0, min(130.0, perf))
    qual = (qtd_ok / (qtd_ok + qtd_nok) * 100.0) if (qtd_ok + qtd_nok) > 0 else 100.0
    qual = max(0.0, min(100.0, qual))
    oee = (disp * min(perf, 100.0) * qual) / 10000.0

    pecas_em_curso = len(pecas_tempo)
    pecas_fora_tempo = sum(1 for r in pecas_tempo if bool(r.get("fora")))

    top_paragens = []
    for key, count in causa_count.items():
        causa, encomenda, operador = key
        dt_last = causa_last_dt.get(key)
        dt_txt = dt_last.strftime("%Y-%m-%d %H:%M:%S") if dt_last is not None else "-"
        top_paragens.append(
            {
                "causa": causa,
                "encomenda": encomenda,
                "operador": operador,
                "ocorrencias": count,
                "minutos": round(causa_time.get(key, 0.0), 1),
                "data_ultima": dt_txt,
            }
        )
    top_paragens.sort(key=lambda r: (r["ocorrencias"], r["minutos"]), reverse=True)

    ultimos_eventos.sort(key=lambda r: r.get("created_at", ""), reverse=True)
    ultimos_eventos = ultimos_eventos[:18]
    interrompidas.sort(key=lambda r: r.get("ts", ""), reverse=True)
    pecas_tempo.sort(
        key=lambda r: (
            1 if bool(r.get("fora")) else 0,
            str(r.get("encomenda", "") or ""),
            _parse_float(r.get("delta_min", 0), 0),
            _parse_float(r.get("elapsed_min", 0), 0),
        ),
        reverse=True,
    )
    desvio_max_min = 0.0
    for r in pecas_tempo:
        desvio_max_min = max(desvio_max_min, _parse_float(r.get("delta_min", 0), 0))
    historico_tempo.sort(key=lambda r: r.get("created_at", ""), reverse=True)

    thr_disp = 85.0
    thr_qual = 97.0
    thr_perf = 80.0
    thr_down = 480.0

    alertas = []
    if live_rows_total > 0 and disp < thr_disp:
        alertas.append(f"- Disponibilidade baixa ({disp:.1f}%).")
    if qual < thr_qual:
        alertas.append(f"- Qualidade abaixo da meta ({qual:.1f}%).")
    if live_rows_total > 0 and perf < thr_perf:
        alertas.append(f"- Performance baixa ({perf:.1f}%).")
    if down_min >= thr_down:
        alertas.append(f"- Paragens elevadas no mes ({down_min:.1f} min).")
    if pecas_fora_tempo > 0:
        alertas.append(f"- {pecas_fora_tempo} peca(s) em curso fora do tempo planeado.")
    if not alertas:
        alertas.append("- Sem alertas criticos.")

    return {
        "oee": oee,
        "disponibilidade": disp,
        "performance": perf,
        "qualidade": qual,
        "start_ops": start_ops,
        "finish_ops": finish_ops,
        "qtd_ok": qtd_ok,
        "qtd_nok": qtd_nok,
        "pecas_em_curso": pecas_em_curso,
        "pecas_fora_tempo": pecas_fora_tempo,
        "desvio_max_min": desvio_max_min,
        "down_min": down_min,
        "andon_prod": andon_prod,
        "andon_setup": andon_setup,
        "andon_wait": andon_wait,
        "andon_stop": andon_stop,
        "top_paragens": top_paragens,
        "ultimos_eventos": ultimos_eventos,
        "interrompidas": interrompidas[:40],
        "pecas_tempo": pecas_tempo[:60],
        "historico_tempo": historico_tempo[:160],
        "paragens_encomenda": {str(k): round(v, 1) for k, v in paragem_order_totals.items() if _parse_float(v, 0) > 0},
        "avaria_count": int(avaria_count),
        "avaria_min": round(avaria_min, 1),
        "alertas": "\n".join(alertas),
        "period_days": pd,
        "quality_scope": f"{year_ref:04d}-{month_ref:02d}",
    }


def _collect_operator_history_rows(app, period_days=0, year_filter=None):
    try:
        if hasattr(app, "refresh_runtime_impulse_data"):
            app.refresh_runtime_impulse_data(cleanup_orphans=True)
    except Exception:
        pass
    eventos = list(app.data.get("op_eventos", []) or [])
    paragens = list(app.data.get("op_paragens", []) or [])
    stock_log = list(app.data.get("stock_log", []) or [])
    produtos_mov = list(app.data.get("produtos_mov", []) or [])
    encomendas = list(app.data.get("encomendas", []) or [])

    def _piece_qty(px):
        return max(
            0.0,
            _parse_float(
                px.get("quantidade_pedida", px.get("qtd", px.get("quantidade", 0))),
                0,
            ),
        )

    piece_meta = {}
    valid_encs = set()
    valid_piece_keys = set()
    for enc in encomendas:
        enc_num = str(enc.get("numero", "") or "").strip()
        if enc_num:
            valid_encs.add(enc_num)
        for m in list(enc.get("materiais", []) or []):
            mat = str(m.get("material", "") or "").strip()
            for esp in list(m.get("espessuras", []) or []):
                esp_txt = str(esp.get("espessura", "") or "").strip()
                for p in list(esp.get("pecas", []) or []):
                    p_id = str(p.get("id", "") or "").strip()
                    if not p_id:
                        continue
                    valid_piece_keys.add((enc_num, p_id))
                    piece_meta[(enc_num, p_id)] = {
                        "ref_interna": str(p.get("ref_interna", "") or "").strip(),
                        "material": str(p.get("material", "") or mat),
                        "espessura": str(p.get("espessura", "") or esp_txt),
                        "qtd": _piece_qty(p),
                    }

    cutoff = None
    try:
        pd = int(period_days or 0)
    except Exception:
        pd = 0
    if pd > 0:
        try:
            from datetime import timedelta
            cutoff = date.today() - timedelta(days=pd - 1)
        except Exception:
            cutoff = None
    try:
        yf = int(str(year_filter or "").strip()) if str(year_filter or "").strip().isdigit() else None
    except Exception:
        yf = None

    def _in_period(raw_dt):
        d = _safe_date(raw_dt)
        if cutoff is not None and d is not None and d < cutoff:
            return False
        if yf is not None and d is not None:
            try:
                if int(d.year) != int(yf):
                    return False
            except Exception:
                pass
        return True

    rows = []
    ev_resume_index = {}
    open_ops = {}
    ev_sorted = sorted(
        list(eventos or []),
        key=lambda r: (_safe_dt(r.get("created_at")) or datetime.min, str(r.get("created_at", "") or "")),
    )
    for ev in ev_sorted:
        ts = str(ev.get("created_at", "") or "")
        if not _in_period(ts):
            continue
        enc_num = str(ev.get("encomenda_numero", "") or "").strip()
        peca_id = str(ev.get("peca_id", "") or "").strip()
        if enc_num and enc_num not in valid_encs:
            continue
        if peca_id and enc_num and (enc_num, peca_id) not in valid_piece_keys:
            continue
        meta = piece_meta.get((enc_num, peca_id), {}) or {}
        evento = str(ev.get("evento", "") or "").strip().upper()
        # PARAGEM/STOP ja entram pela tabela dedicada op_paragens.
        # Se mantivermos aqui tambem, aparecem duplicadas no historico.
        if evento in ("PARAGEM", "STOP"):
            continue
        operacao = str(ev.get("operacao", "") or "").strip()
        operador = str(ev.get("operador", "") or "").strip()
        material = str(ev.get("material", "") or meta.get("material", "") or "").strip()
        esp = str(ev.get("espessura", "") or meta.get("espessura", "") or "").strip()
        ref_int = str(ev.get("ref_interna", "") or meta.get("ref_interna", "") or peca_id or "").strip()
        qtd_ok = _parse_float(ev.get("qtd_ok", 0), 0)
        qtd_nok = _parse_float(ev.get("qtd_nok", 0), 0)
        info = str(ev.get("info", "") or "").strip()
        tempo_min = 0.0
        op_norm = _norm_text(operacao)
        key = (enc_num, peca_id, op_norm, operador or "-")
        if evento in ("RESUME_PIECE", "START_OP", "FINISH_OP"):
            k_resume = (enc_num, peca_id)
            ev_resume_index.setdefault(k_resume, []).append(_safe_dt(ts))
        if evento == "START_OP":
            open_ops.setdefault(key, []).append(_safe_dt(ts))
        elif evento == "FINISH_OP":
            st_list = list(open_ops.get(key, []) or [])
            if st_list:
                ini = st_list.pop(-1)
                open_ops[key] = st_list
                try:
                    if ini is not None and _safe_dt(ts) is not None:
                        tempo_min = max(0.0, (_safe_dt(ts) - ini).total_seconds() / 60.0)
                except Exception:
                    tempo_min = 0.0
        rows.append(
            {
                "created_at": ts,
                "origem": "Operacao",
                "evento": evento or "-",
                "encomenda": enc_num,
                "material": material,
                "espessura": esp,
                "peca": ref_int,
                "operacao": operacao or "-",
                "operador": operador or "-",
                "qtd_ok": qtd_ok,
                "qtd_nok": qtd_nok,
                "tempo_min": round(tempo_min, 1),
                "detalhe": info,
                "causa": "",
            }
        )

    now_iso = datetime.now().isoformat(sep=" ", timespec="seconds")
    paragem_groups = {}
    for pa in paragens:
        ts = str(pa.get("created_at", "") or "")
        if not _in_period(ts):
            continue
        enc_num = str(pa.get("encomenda_numero", "") or "").strip()
        peca_id = str(pa.get("peca_id", "") or "").strip()
        if enc_num and enc_num not in valid_encs:
            continue
        if peca_id and enc_num and (enc_num, peca_id) not in valid_piece_keys:
            continue
        mins = _parse_float(pa.get("duracao_min", 0), 0)
        estado_pa = str(pa.get("estado", "") or "").strip().upper()
        fechada_at = str(pa.get("fechada_at", "") or "").strip()
        if mins <= 0 and estado_pa == "ABERTA":
            mins = max(0.0, _diff_min(ts, now_iso) or 0.0)
        if mins <= 0:
            t0 = _safe_dt(ts)
            if t0 is not None:
                cands = []
                for dt in list(ev_resume_index.get((enc_num, peca_id), []) or []):
                    if dt is not None and dt > t0:
                        cands.append(dt)
                if cands:
                    t1 = min(cands)
                    try:
                        mins = max(0.0, (t1 - t0).total_seconds() / 60.0)
                    except Exception:
                        mins = 0.0
        if mins <= 0:
            continue
        ref_int = str(pa.get("ref_interna", "") or (piece_meta.get((enc_num, peca_id), {}) or {}).get("ref_interna", "") or peca_id or "").strip()
        key = (
            enc_num or "-",
            str(pa.get("operador", "") or "").strip() or "-",
            str(pa.get("causa", "") or "").strip() or "Sem causa",
            ts,
            fechada_at,
            estado_pa,
        )
        grp = paragem_groups.setdefault(
            key,
            {
                "created_at": ts,
                "origem": "Paragem",
                "evento": "PARAGEM",
                "encomenda": enc_num,
                "materials": set(),
                "espessuras": set(),
                "pecas": [],
                "operacao": "",
                "operador": str(pa.get("operador", "") or "").strip() or "-",
                "qtd_ok": 0.0,
                "qtd_nok": 0.0,
                "tempo_min": 0.0,
                "detalhes": [],
                "causa": str(pa.get("causa", "") or "").strip(),
            },
        )
        grp["tempo_min"] = max(_parse_float(grp.get("tempo_min", 0), 0), mins)
        material_txt = str(pa.get("material", "") or "").strip()
        esp_txt = str(pa.get("espessura", "") or "").strip()
        if material_txt:
            grp["materials"].add(material_txt)
        if esp_txt:
            grp["espessuras"].add(esp_txt)
        if ref_int and ref_int not in grp["pecas"]:
            grp["pecas"].append(ref_int)
        det = str(pa.get("detalhe", "") or "").strip()
        if det and det not in grp["detalhes"]:
            grp["detalhes"].append(det)

    for grp in list(paragem_groups.values()):
        pecas = list(grp.get("pecas", []) or [])
        if len(pecas) <= 1:
            peca_txt = str(pecas[0] if pecas else "")
        else:
            peca_txt = f"{pecas[0]} (+{len(pecas) - 1})"
        detalhe_parts = []
        if len(pecas) > 1:
            detalhe_parts.append("Pecas afetadas: " + ", ".join(pecas))
        detalhe_parts.extend(list(grp.get("detalhes", []) or []))
        rows.append(
            {
                "created_at": str(grp.get("created_at", "") or ""),
                "origem": "Paragem",
                "evento": "PARAGEM",
                "encomenda": str(grp.get("encomenda", "") or "").strip(),
                "material": " | ".join(sorted(list(grp.get("materials", set()) or []))),
                "espessura": " | ".join(sorted(list(grp.get("espessuras", set()) or []))),
                "peca": peca_txt,
                "operacao": "",
                "operador": str(grp.get("operador", "") or "").strip() or "-",
                "qtd_ok": 0.0,
                "qtd_nok": 0.0,
                "tempo_min": round(_parse_float(grp.get("tempo_min", 0), 0), 1),
                "detalhe": " | ".join([x for x in detalhe_parts if str(x or "").strip()]),
                "causa": str(grp.get("causa", "") or "").strip(),
            }
        )

    for s in stock_log:
        ts = str(s.get("data", "") or "")
        if not _in_period(ts):
            continue
        acao = str(s.get("acao", "") or "").strip()
        det = str(s.get("detalhes", "") or "").strip()
        enc_num = _extract_encomenda_from_text(det)
        rows.append(
            {
                "created_at": ts,
                "origem": "Stock MP",
                "evento": acao or "-",
                "encomenda": enc_num,
                "material": "",
                "espessura": "",
                "peca": "",
                "operacao": "",
                "operador": "-",
                "qtd_ok": 0.0,
                "qtd_nok": 0.0,
                "tempo_min": 0.0,
                "detalhe": det,
                "causa": "",
            }
        )

    for mv in produtos_mov:
        ts = str(mv.get("data", "") or "")
        if not _in_period(ts):
            continue
        detalhe = (
            f"Codigo: {str(mv.get('codigo', '') or '')} | Desc: {str(mv.get('descricao', '') or '')} | "
            f"Antes: {_parse_float(mv.get('antes', 0), 0):.2f} | Depois: {_parse_float(mv.get('depois', 0), 0):.2f} | "
            f"Ref: {str(mv.get('ref_doc', '') or '')} | Obs: {str(mv.get('obs', '') or '')}"
        )
        rows.append(
            {
                "created_at": ts,
                "origem": "Produtos",
                "evento": str(mv.get("tipo", "") or "").strip() or "-",
                "encomenda": _extract_encomenda_from_text(str(mv.get("ref_doc", "") or "")),
                "material": "",
                "espessura": "",
                "peca": str(mv.get("codigo", "") or "").strip(),
                "operacao": "",
                "operador": str(mv.get("operador", "") or "").strip() or "-",
                "qtd_ok": _parse_float(mv.get("qtd", 0), 0),
                "qtd_nok": 0.0,
                "tempo_min": 0.0,
                "detalhe": detalhe,
                "causa": "",
            }
        )

    # Dedupe defensivo: evita repeticoes por sincronizacao/backfill.
    dedup = {}
    for r in rows:
        k = (
            str(r.get("created_at", "") or ""),
            _norm_text(r.get("origem", "")),
            _norm_text(r.get("evento", "")),
            str(r.get("encomenda", "") or ""),
            str(r.get("peca", "") or ""),
            _norm_text(r.get("operacao", "")),
            str(r.get("operador", "") or ""),
            _norm_text(r.get("causa", "")),
            _norm_text(r.get("detalhe", "")),
        )
        if k not in dedup:
            dedup[k] = r
    rows = list(dedup.values())
    rows.sort(key=lambda r: (_safe_dt(r.get("created_at")) or datetime.min, str(r.get("created_at", "") or "")), reverse=True)
    return rows


def open_production_ops_history_detail_window(app, row):
    if not isinstance(row, dict):
        return
    use_custom = ctk is not None and os.environ.get("USE_CUSTOM_DASH_PULSE", "1") != "0"
    old = getattr(app, "dashboard_ops_hist_detail_win", None)
    try:
        if old is not None and old.winfo_exists():
            try:
                old.destroy()
            except Exception:
                pass
    except Exception:
        pass

    Win = ctk.CTkToplevel if use_custom else Toplevel
    Frame = ctk.CTkFrame if use_custom else ttk.Frame
    Label = ctk.CTkLabel if use_custom else ttk.Label
    Btn = ctk.CTkButton if use_custom else ttk.Button

    win = Win(app.root)
    app.dashboard_ops_hist_detail_win = win
    try:
        win.title("Detalhe do Registo")
        win.geometry("980x640")
        win.minsize(860, 560)
        if use_custom:
            win.configure(fg_color="#eef2f7")
    except Exception:
        pass
    _bring_window_front(win, getattr(app, "dashboard_ops_hist_win", app.root), modal=True, keep_topmost=False)

    head = Frame(win, fg_color="#ffffff" if use_custom else "transparent", corner_radius=12 if use_custom else None, border_width=1 if use_custom else 0, border_color="#d7dee9" if use_custom else None) if use_custom else Frame(win)
    head.pack(fill="x", padx=14, pady=(14, 8))
    left = Frame(head, fg_color="transparent") if use_custom else Frame(head)
    left.pack(side="left", fill="x", expand=True, padx=12, pady=10)
    Label(left, text="Detalhe do Registo", font=("Segoe UI", 22, "bold") if use_custom else ("Segoe UI", 15, "bold"), text_color="#0f172a" if use_custom else None).pack(anchor="w")
    Label(left, text="Visualizacao completa do evento selecionado.", font=("Segoe UI", 12) if use_custom else ("Segoe UI", 9), text_color="#64748b" if use_custom else None).pack(anchor="w", pady=(2, 0))
    Btn(head, text="Fechar", width=110 if use_custom else None, command=win.destroy).pack(side="right", padx=12, pady=12)

    body = Frame(win, fg_color="#ffffff" if use_custom else "transparent", corner_radius=12 if use_custom else None, border_width=1 if use_custom else 0, border_color="#d7dee9" if use_custom else None) if use_custom else Frame(win)
    body.pack(fill="both", expand=True, padx=14, pady=(0, 14))

    grid = Frame(body, fg_color="transparent") if use_custom else Frame(body)
    grid.pack(fill="x", padx=12, pady=(12, 8))

    field_specs = [
        ("Data", "created_at"),
        ("Origem", "origem"),
        ("Evento", "evento"),
        ("Encomenda", "encomenda"),
        ("Material", "material"),
        ("Espessura", "espessura"),
        ("Peca", "peca"),
        ("Operacao", "operacao"),
        ("Operador", "operador"),
        ("OK", "qtd_ok"),
        ("NOK", "qtd_nok"),
        ("Tempo (min)", "tempo_min"),
        ("Causa", "causa"),
    ]

    for i in range(3):
        try:
            grid.grid_columnconfigure(i, weight=1, uniform="hist_detail")
        except Exception:
            pass

    for idx, (label_txt, key) in enumerate(field_specs):
        r, c = divmod(idx, 3)
        card = Frame(grid, fg_color="#f8fbff" if use_custom else "white", corner_radius=10 if use_custom else None, border_width=1 if use_custom else 0, border_color="#dbe3ef" if use_custom else None) if use_custom else Frame(grid)
        card.grid(row=r, column=c, padx=6, pady=6, sticky="nsew")
        Label(card, text=label_txt, font=("Segoe UI", 10, "bold") if use_custom else ("Segoe UI", 9, "bold"), text_color="#64748b" if use_custom else None).pack(anchor="w", padx=10, pady=(8, 0))
        raw = row.get(key, "")
        if key in ("qtd_ok", "qtd_nok"):
            val = _fmt_num(raw)
        elif key == "tempo_min":
            val = f"{_parse_float(raw, 0):.1f}"
        else:
            val = str(raw or "-")
        Label(card, text=val, font=("Segoe UI", 14, "bold") if use_custom else ("Segoe UI", 10, "bold"), text_color="#0f172a" if use_custom else None, anchor="w", justify="left", wraplength=280).pack(fill="x", padx=10, pady=(2, 10))

    detail_wrap = Frame(body, fg_color="#f8fafc" if use_custom else "white", corner_radius=10 if use_custom else None, border_width=1 if use_custom else 0, border_color="#dbe3ef" if use_custom else None) if use_custom else Frame(body)
    detail_wrap.pack(fill="both", expand=True, padx=12, pady=(4, 12))
    Label(detail_wrap, text="Detalhe completo", font=("Segoe UI", 12, "bold") if use_custom else ("Segoe UI", 10, "bold"), text_color="#334155" if use_custom else None).pack(anchor="w", padx=10, pady=(8, 4))
    detail_text = str(row.get("detalhe", "") or "").strip()
    if not detail_text:
        detail_text = "-"
    if use_custom and ctk is not None:
        box = ctk.CTkTextbox(detail_wrap, fg_color="#ffffff", border_width=1, border_color="#dbe3ef", corner_radius=8)
        box.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        try:
            box.insert("1.0", detail_text)
            box.configure(state="disabled")
        except Exception:
            pass
    else:
        ttk.Label(detail_wrap, text=detail_text, justify="left", wraplength=900).pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def _on_close():
        try:
            try:
                win.grab_release()
            except Exception:
                pass
            win.destroy()
        finally:
            app.dashboard_ops_hist_detail_win = None

    try:
        win.protocol("WM_DELETE_WINDOW", _on_close)
    except Exception:
        pass


def refresh_production_ops_history_window(app):
    win = getattr(app, "dashboard_ops_hist_win", None)
    if win is None:
        return
    try:
        if not win.winfo_exists():
            app.dashboard_ops_hist_win = None
            return
    except Exception:
        app.dashboard_ops_hist_win = None
        return

    period_txt = str(getattr(app, "dashboard_ops_hist_period_var", StringVar(value="7 dias")).get() or "7 dias").strip()
    period_map = {"Hoje": 1, "7 dias": 7, "30 dias": 30, "Tudo": 0}
    period_days = period_map.get(period_txt, 7)
    year_txt = str(getattr(app, "dashboard_ops_hist_year_var", StringVar(value="Todos")).get() or "Todos").strip()
    year_filter = None if _norm_text(year_txt) == "todos" else year_txt

    rows = _collect_operator_history_rows(app, period_days=period_days, year_filter=year_filter)
    app.dashboard_ops_hist_rows_all = rows

    enc_vals = ["Todas"] + sorted({str(r.get("encomenda", "") or "").strip() for r in rows if str(r.get("encomenda", "") or "").strip()})
    op_vals = ["Todos"] + sorted({str(r.get("operador", "") or "").strip() for r in rows if str(r.get("operador", "") or "").strip() and str(r.get("operador", "") or "").strip() != "-"})
    mat_vals = ["Todos"] + sorted({str(r.get("material", "") or "").strip() for r in rows if str(r.get("material", "") or "").strip()})
    esp_vals = ["Todas"] + sorted({str(r.get("espessura", "") or "").strip() for r in rows if str(r.get("espessura", "") or "").strip()})

    for attr, vals in (
        ("dashboard_ops_hist_enc_menu", enc_vals),
        ("dashboard_ops_hist_op_menu", op_vals),
        ("dashboard_ops_hist_mat_menu", mat_vals),
        ("dashboard_ops_hist_esp_menu", esp_vals),
    ):
        menu = getattr(app, attr, None)
        if menu is None:
            continue
        try:
            if ctk is not None and menu.__class__.__name__.lower().startswith("ctk"):
                menu.configure(values=vals)
            else:
                menu["values"] = vals
        except Exception:
            pass

    enc_filter = str(getattr(app, "dashboard_ops_hist_enc_var", StringVar(value="Todas")).get() or "Todas").strip()
    op_filter = str(getattr(app, "dashboard_ops_hist_operador_var", StringVar(value="Todos")).get() or "Todos").strip()
    mat_filter = str(getattr(app, "dashboard_ops_hist_mat_var", StringVar(value="Todos")).get() or "Todos").strip()
    esp_filter = str(getattr(app, "dashboard_ops_hist_esp_var", StringVar(value="Todas")).get() or "Todas").strip()
    origem_filter = _norm_text(str(getattr(app, "dashboard_ops_hist_origem_var", StringVar(value="Todos")).get() or "Todos"))
    texto_filter = _norm_text(str(getattr(app, "dashboard_ops_hist_text_var", StringVar(value="")).get() or ""))
    causa_filter = _norm_text(str(getattr(app, "dashboard_ops_hist_causa_var", StringVar(value="")).get() or ""))
    col_filters = getattr(app, "dashboard_ops_hist_col_vars", {}) or {}

    out = []
    for r in rows:
        if enc_filter.lower() != "todas" and str(r.get("encomenda", "") or "").strip() != enc_filter:
            continue
        if op_filter.lower() != "todos" and str(r.get("operador", "") or "").strip() != op_filter:
            continue
        if mat_filter.lower() != "todos" and str(r.get("material", "") or "").strip() != mat_filter:
            continue
        if esp_filter.lower() != "todas" and str(r.get("espessura", "") or "").strip() != esp_filter:
            continue
        orig = _norm_text(r.get("origem", ""))
        if origem_filter.startswith("oper") and "oper" not in orig:
            continue
        if origem_filter.startswith("parag") and "parag" not in orig:
            continue
        if origem_filter.startswith("stock") and "stock" not in orig:
            continue
        if origem_filter.startswith("produt") and "produt" not in orig:
            continue
        if texto_filter:
            blob = _norm_text(
                " | ".join(
                    [
                        str(r.get("encomenda", "") or ""),
                        str(r.get("peca", "") or ""),
                        str(r.get("operacao", "") or ""),
                        str(r.get("operador", "") or ""),
                        str(r.get("detalhe", "") or ""),
                    ]
                )
            )
            if texto_filter not in blob:
                continue
        if causa_filter and causa_filter not in _norm_text(r.get("causa", "")):
            continue
        if col_filters:
            col_map = {
                "data": str(r.get("created_at", "") or ""),
                "origem": str(r.get("origem", "") or ""),
                "evento": str(r.get("evento", "") or ""),
                "encomenda": str(r.get("encomenda", "") or ""),
                "material": str(r.get("material", "") or ""),
                "esp": str(r.get("espessura", "") or ""),
                "peca": str(r.get("peca", "") or ""),
                "operacao": str(r.get("operacao", "") or ""),
                "operador": str(r.get("operador", "") or ""),
                "causa": str(r.get("causa", "") or ""),
                "detalhe": str(r.get("detalhe", "") or ""),
            }
            reject = False
            for ck, cvar in col_filters.items():
                try:
                    want = _norm_text(cvar.get())
                except Exception:
                    want = ""
                if not want:
                    continue
                got = _norm_text(col_map.get(ck, ""))
                if want not in got:
                    reject = True
                    break
            if reject:
                continue
        out.append(r)

    app.dashboard_ops_hist_rows_filtered = out
    tree = getattr(app, "dashboard_ops_hist_tree", None)
    app.dashboard_ops_hist_row_map = {}
    if tree is not None:
        try:
            for iid in tree.get_children():
                tree.delete(iid)
            for idx, r in enumerate(out):
                origem = str(r.get("origem", "") or "")
                tag = "odd" if idx % 2 else "even"
                if "paragem" in _norm_text(origem):
                    tag = "pause"
                elif "stock" in _norm_text(origem):
                    tag = "stock"
                elif "produt" in _norm_text(origem):
                    tag = "produto"
                iid = tree.insert(
                    "",
                    "end",
                    values=(
                        str(r.get("created_at", "") or ""),
                        origem,
                        str(r.get("evento", "") or ""),
                        str(r.get("encomenda", "") or ""),
                        str(r.get("material", "") or ""),
                        str(r.get("espessura", "") or ""),
                        str(r.get("peca", "") or ""),
                        str(r.get("operacao", "") or ""),
                        str(r.get("operador", "") or ""),
                        _fmt_num(r.get("qtd_ok", 0)),
                        _fmt_num(r.get("qtd_nok", 0)),
                        f"{_parse_float(r.get('tempo_min', 0), 0):.1f}",
                        str(r.get("causa", "") or ""),
                        str(r.get("detalhe", "") or ""),
                    ),
                    tags=(tag,),
                )
                app.dashboard_ops_hist_row_map[iid] = dict(r)
        except Exception:
            pass

    info_var = getattr(app, "dashboard_ops_hist_info_var", None)
    if info_var is not None:
        try:
            info_var.set(f"Registos: {len(out)} (filtro) / {len(rows)} (total)")
        except Exception:
            pass
    try:
        summary_vars = getattr(app, "dashboard_ops_hist_summary_vars", {}) or {}
        total_tempo = sum(max(0.0, _parse_float(r.get("tempo_min", 0), 0)) for r in out)
        total_paragens = sum(1 for r in out if "paragem" in _norm_text(r.get("origem", "")))
        total_oper = sum(1 for r in out if "oper" in _norm_text(r.get("origem", "")))
        total_stock_prod = sum(1 for r in out if "stock" in _norm_text(r.get("origem", "")) or "produt" in _norm_text(r.get("origem", "")))
        if summary_vars:
            for k, v in (
                ("total", str(len(out))),
                ("operacoes", str(total_oper)),
                ("paragens", str(total_paragens)),
                ("stock_prod", str(total_stock_prod)),
                ("tempo", f"{total_tempo:.1f} min"),
            ):
                try:
                    summary_vars.get(k).set(v)
                except Exception:
                    pass
    except Exception:
        pass

    # Auto-refresh para manter tempos e eventos sempre atualizados.
    try:
        prev_after = getattr(app, "dashboard_ops_hist_after_id", None)
        if prev_after:
            try:
                app.root.after_cancel(prev_after)
            except Exception:
                pass
        if getattr(app, "dashboard_ops_hist_win", None) is not None:
            app.dashboard_ops_hist_after_id = app.root.after(3000, lambda: refresh_production_ops_history_window(app))
    except Exception:
        pass

def open_production_ops_history_window(app, preset=None):
    use_custom = ctk is not None and os.environ.get("USE_CUSTOM_DASH_PULSE", "1") != "0"
    old = getattr(app, "dashboard_ops_hist_win", None)
    try:
        if old is not None and old.winfo_exists():
            old.lift()
            old.focus_force()
            if isinstance(preset, dict):
                enc = str(preset.get("encomenda", "") or "").strip()
                if enc:
                    try:
                        app.dashboard_ops_hist_enc_var.set(enc)
                    except Exception:
                        pass
                causa = str(preset.get("causa", "") or "").strip()
                if causa:
                    try:
                        app.dashboard_ops_hist_causa_var.set(causa)
                    except Exception:
                        pass
                op = str(preset.get("operador", "") or "").strip()
                if op:
                    try:
                        app.dashboard_ops_hist_operador_var.set(op)
                    except Exception:
                        pass
            refresh_production_ops_history_window(app)
            return
    except Exception:
        pass

    Win = ctk.CTkToplevel if use_custom else Toplevel
    Frame = ctk.CTkFrame if use_custom else ttk.Frame
    Label = ctk.CTkLabel if use_custom else ttk.Label
    Btn = ctk.CTkButton if use_custom else ttk.Button

    win = Win(app.root)
    app.dashboard_ops_hist_win = win
    try:
        win.title("Historico de Operacoes")
        win.geometry("1460x860")
        win.minsize(1260, 700)
        if use_custom:
            win.configure(fg_color="#eef2f7")
    except Exception:
        pass
    _bring_window_front(win, app.root, modal=True, keep_topmost=False)

    app.dashboard_ops_hist_period_var = StringVar(value=str(getattr(app, "dashboard_pulse_period_var", StringVar(value="7 dias")).get() or "7 dias"))
    app.dashboard_ops_hist_year_var = StringVar(value=str(getattr(app, "dashboard_pulse_year_var", StringVar(value="Todos")).get() or "Todos"))
    app.dashboard_ops_hist_enc_var = StringVar(value="Todas")
    app.dashboard_ops_hist_operador_var = StringVar(value="Todos")
    app.dashboard_ops_hist_mat_var = StringVar(value="Todos")
    app.dashboard_ops_hist_esp_var = StringVar(value="Todas")
    app.dashboard_ops_hist_origem_var = StringVar(value="Todos")
    app.dashboard_ops_hist_text_var = StringVar(value="")
    app.dashboard_ops_hist_causa_var = StringVar(value="")
    app.dashboard_ops_hist_info_var = StringVar(value="Registos: 0")
    app.dashboard_ops_hist_col_vars = {
        "data": StringVar(value=""),
        "origem": StringVar(value=""),
        "evento": StringVar(value=""),
        "encomenda": StringVar(value=""),
        "material": StringVar(value=""),
        "esp": StringVar(value=""),
        "peca": StringVar(value=""),
        "operacao": StringVar(value=""),
        "operador": StringVar(value=""),
        "causa": StringVar(value=""),
        "detalhe": StringVar(value=""),
    }
    app.dashboard_ops_hist_rows_all = []
    app.dashboard_ops_hist_after_id = None
    app.dashboard_ops_hist_enc_menu = None
    app.dashboard_ops_hist_op_menu = None
    app.dashboard_ops_hist_mat_menu = None
    app.dashboard_ops_hist_esp_menu = None

    app.dashboard_ops_hist_summary_vars = {
        "total": StringVar(value="0"),
        "operacoes": StringVar(value="0"),
        "paragens": StringVar(value="0"),
        "stock_prod": StringVar(value="0"),
        "tempo": StringVar(value="0.0 min"),
    }

    top = Frame(
        win,
        fg_color="#ffffff" if use_custom else "transparent",
        corner_radius=12 if use_custom else None,
        border_width=1 if use_custom else 0,
        border_color="#d7dee9" if use_custom else None,
    ) if use_custom else Frame(win)
    top.pack(fill="x", padx=14, pady=(12, 8))
    top_left = Frame(top, fg_color="transparent") if use_custom else Frame(top)
    top_left.pack(side="left", fill="x", expand=True, padx=12, pady=10)
    Label(top_left, text="Historico de Operacoes", font=("Segoe UI", 24, "bold") if use_custom else ("Segoe UI", 15, "bold"), text_color="#0f172a" if use_custom else None).pack(anchor="w")
    Label(top_left, text="Rastreabilidade de operacoes, avarias, stock e movimentos por encomenda.", font=("Segoe UI", 12) if use_custom else ("Segoe UI", 9), text_color="#64748b" if use_custom else None).pack(anchor="w", pady=(2, 0))
    top_right = Frame(top, fg_color="transparent") if use_custom else Frame(top)
    top_right.pack(side="right", padx=12, pady=10)
    Label(top_right, textvariable=app.dashboard_ops_hist_info_var, font=("Segoe UI", 12, "bold") if use_custom else ("Segoe UI", 10, "bold"), text_color="#475569" if use_custom else None).pack(anchor="e", pady=(0, 6))
    Btn(top_right, text="Fechar", width=110 if use_custom else None, command=win.destroy).pack(anchor="e")

    filters = Frame(
        win,
        fg_color="#ffffff" if use_custom else "transparent",
        corner_radius=10 if use_custom else None,
        border_width=1 if use_custom else 0,
        border_color="#d7dee9" if use_custom else None,
    ) if use_custom else Frame(win)
    filters.pack(fill="x", padx=14, pady=(0, 8))

    def _opt(parent, values, var, width, cb):
        if use_custom and ctk is not None:
            w = ctk.CTkOptionMenu(parent, values=values, variable=var, width=width, command=lambda _v=None: cb())
            return w
        w = ttk.Combobox(parent, values=values, width=max(8, int(width / 12)), state="readonly", textvariable=var)
        w.bind("<<ComboboxSelected>>", lambda _e: cb())
        return w

    Label(filters, text="Filtros principais", font=("Segoe UI", 12, "bold") if use_custom else ("Segoe UI", 10, "bold"), text_color="#0f172a" if use_custom else None).pack(anchor="w", padx=12, pady=(8, 0))

    filter_row_1 = Frame(filters, fg_color="transparent") if use_custom else Frame(filters)
    filter_row_1.pack(fill="x", padx=10, pady=(8, 2))
    filter_row_2 = Frame(filters, fg_color="transparent") if use_custom else Frame(filters)
    filter_row_2.pack(fill="x", padx=10, pady=(2, 8))

    def _filter_cell(parent, label_txt, card_width=None):
        if use_custom:
            cell = ctk.CTkFrame(
                parent,
                fg_color="#f8fbff",
                corner_radius=10,
                border_width=1,
                border_color="#dbe3ef",
            )
        else:
            cell = Frame(parent)
        cell.pack(side="left", padx=(0, 10), pady=2)
        if card_width:
            try:
                cell.configure(width=card_width)
            except Exception:
                pass
        Label(
            cell,
            text=label_txt,
            font=("Segoe UI", 11, "bold") if use_custom else ("Segoe UI", 10, "bold"),
            text_color="#334155" if use_custom else None,
        ).pack(anchor="w", padx=8 if use_custom else 0, pady=(6 if use_custom else 0, 2))
        return cell

    cell = _filter_cell(filter_row_1, "Periodo", 132)
    _opt(cell, ["Hoje", "7 dias", "30 dias", "Tudo"], app.dashboard_ops_hist_period_var, 110, lambda: refresh_production_ops_history_window(app)).pack(anchor="w", padx=8 if use_custom else 0, pady=(0, 8 if use_custom else 0))
    cell = _filter_cell(filter_row_1, "Ano", 106)
    years = ["Todos"] + [str(y) for y in range(datetime.now().year - 3, datetime.now().year + 2)]
    _opt(cell, years, app.dashboard_ops_hist_year_var, 90, lambda: refresh_production_ops_history_window(app)).pack(anchor="w", padx=8 if use_custom else 0, pady=(0, 8 if use_custom else 0))
    cell = _filter_cell(filter_row_1, "Encomenda", 188)
    app.dashboard_ops_hist_enc_menu = _opt(cell, ["Todas"], app.dashboard_ops_hist_enc_var, 170, lambda: refresh_production_ops_history_window(app))
    app.dashboard_ops_hist_enc_menu.pack(anchor="w", padx=8 if use_custom else 0, pady=(0, 8 if use_custom else 0))
    cell = _filter_cell(filter_row_1, "Operador", 178)
    app.dashboard_ops_hist_op_menu = _opt(cell, ["Todos"], app.dashboard_ops_hist_operador_var, 160, lambda: refresh_production_ops_history_window(app))
    app.dashboard_ops_hist_op_menu.pack(anchor="w", padx=8 if use_custom else 0, pady=(0, 8 if use_custom else 0))
    cell = _filter_cell(filter_row_1, "Origem", 158)
    _opt(cell, ["Todos", "Operacoes", "Paragens", "Stock", "Produtos"], app.dashboard_ops_hist_origem_var, 140, lambda: refresh_production_ops_history_window(app)).pack(anchor="w", padx=8 if use_custom else 0, pady=(0, 8 if use_custom else 0))
    Btn(filter_row_1, text="Atualizar", width=120 if use_custom else None, command=lambda: refresh_production_ops_history_window(app)).pack(side="right", padx=6, pady=(20, 0))

    cell = _filter_cell(filter_row_2, "Material", 178)
    app.dashboard_ops_hist_mat_menu = _opt(cell, ["Todos"], app.dashboard_ops_hist_mat_var, 160, lambda: refresh_production_ops_history_window(app))
    app.dashboard_ops_hist_mat_menu.pack(anchor="w", padx=8 if use_custom else 0, pady=(0, 8 if use_custom else 0))
    cell = _filter_cell(filter_row_2, "Espessura", 148)
    app.dashboard_ops_hist_esp_menu = _opt(cell, ["Todas"], app.dashboard_ops_hist_esp_var, 130, lambda: refresh_production_ops_history_window(app))
    app.dashboard_ops_hist_esp_menu.pack(anchor="w", padx=8 if use_custom else 0, pady=(0, 8 if use_custom else 0))
    cell = _filter_cell(filter_row_2, "Filtro livre", 280)
    ent_txt = ctk.CTkEntry(cell, textvariable=app.dashboard_ops_hist_text_var, width=260) if use_custom and ctk is not None else ttk.Entry(cell, textvariable=app.dashboard_ops_hist_text_var, width=28)
    ent_txt.pack(anchor="w", padx=8 if use_custom else 0, pady=(0, 8 if use_custom else 0))
    cell = _filter_cell(filter_row_2, "Causa", 300)
    ent_causa = ctk.CTkEntry(cell, textvariable=app.dashboard_ops_hist_causa_var, width=280) if use_custom and ctk is not None else ttk.Entry(cell, textvariable=app.dashboard_ops_hist_causa_var, width=28)
    ent_causa.pack(anchor="w", padx=8 if use_custom else 0, pady=(0, 8 if use_custom else 0))
    try:
        ent_txt.bind("<KeyRelease>", lambda _e: refresh_production_ops_history_window(app))
        ent_causa.bind("<KeyRelease>", lambda _e: refresh_production_ops_history_window(app))
    except Exception:
        pass

    summary = Frame(
        win,
        fg_color="#ffffff" if use_custom else "transparent",
        corner_radius=10 if use_custom else None,
        border_width=1 if use_custom else 0,
        border_color="#d7dee9" if use_custom else None,
    ) if use_custom else Frame(win)
    summary.pack(fill="x", padx=14, pady=(0, 8))

    def _summary_chip(parent, title, var_name, width=170):
        chip = Frame(parent, fg_color="#f8fbff" if use_custom else "white", corner_radius=10 if use_custom else None, border_width=1 if use_custom else 0, border_color="#dbe3ef" if use_custom else None) if use_custom else Frame(parent)
        chip.pack(side="left", padx=(0, 8), pady=8)
        Label(chip, text=title, font=("Segoe UI", 10, "bold") if use_custom else ("Segoe UI", 9, "bold"), text_color="#64748b" if use_custom else None).pack(anchor="w", padx=10, pady=(6, 0))
        Label(chip, textvariable=app.dashboard_ops_hist_summary_vars[var_name], font=("Segoe UI", 16, "bold") if use_custom else ("Segoe UI", 11, "bold"), text_color="#0f172a" if use_custom else None).pack(anchor="w", padx=10, pady=(0, 8))
        try:
            chip.configure(width=width)
        except Exception:
            pass

    _summary_chip(summary, "Registos filtrados", "total", 170)
    _summary_chip(summary, "Operacoes", "operacoes", 150)
    _summary_chip(summary, "Paragens", "paragens", 150)
    _summary_chip(summary, "Stock/Produtos", "stock_prod", 170)
    _summary_chip(summary, "Tempo acumulado", "tempo", 180)

    body = Frame(
        win,
        fg_color="#ffffff" if use_custom else None,
        corner_radius=10 if use_custom else None,
        border_width=1 if use_custom else 0,
        border_color="#d7dee9" if use_custom else None,
    ) if use_custom else Frame(win)
    body.pack(fill="both", expand=True, padx=14, pady=(0, 12))

    body_title = Frame(body, fg_color="transparent") if use_custom else Frame(body)
    body_title.pack(fill="x", padx=10, pady=(10, 4))
    Label(body_title, text="Registos detalhados", font=("Segoe UI", 13, "bold") if use_custom else ("Segoe UI", 10, "bold"), text_color="#0f172a" if use_custom else None).pack(side="left")
    Label(body_title, text="Duplo clique na linha para abrir o detalhe completo.", font=("Segoe UI", 11) if use_custom else ("Segoe UI", 9), text_color="#64748b" if use_custom else None).pack(side="right")

    excel_filters = Frame(
        body,
        fg_color="#f8fafc" if use_custom else "white",
        corner_radius=8 if use_custom else None,
        border_width=1 if use_custom else 0,
        border_color="#dbe3ef" if use_custom else None,
    ) if use_custom else Frame(body)
    excel_filters.pack(fill="x", padx=8, pady=(4, 4))
    excel_head = Frame(excel_filters, fg_color="transparent") if use_custom else Frame(excel_filters)
    excel_head.pack(fill="x", padx=8, pady=(6, 2))
    Label(
        excel_head,
        text="Filtros por coluna (tipo Excel):",
        font=("Segoe UI", 11, "bold") if use_custom else ("Segoe UI", 9, "bold"),
        text_color="#334155" if use_custom else None,
    ).pack(side="left")
    Label(
        excel_head,
        text="Pesquisa fina por campo.",
        font=("Segoe UI", 10) if use_custom else ("Segoe UI", 8),
        text_color="#64748b" if use_custom else None,
    ).pack(side="right")

    def _mk_col_filter(parent, key, width):
        var = app.dashboard_ops_hist_col_vars.get(key)
        if use_custom and ctk is not None:
            w = ctk.CTkEntry(parent, textvariable=var, width=width, height=28, placeholder_text=key)
        else:
            w = ttk.Entry(parent, textvariable=var, width=max(8, int(width / 10)))
        try:
            w.bind("<KeyRelease>", lambda _e: refresh_production_ops_history_window(app))
        except Exception:
            pass
        return w

    filters_row_1 = Frame(excel_filters, fg_color="transparent") if use_custom else Frame(excel_filters)
    filters_row_1.pack(fill="x", padx=8, pady=(0, 2))
    filters_row_2 = Frame(excel_filters, fg_color="transparent") if use_custom else Frame(excel_filters)
    filters_row_2.pack(fill="x", padx=8, pady=(0, 6))

    def _mk_labeled(parent, key, label, width):
        cell = Frame(parent, fg_color="transparent") if use_custom else Frame(parent)
        cell.pack(side="left", padx=(0, 6), pady=2)
        Label(
            cell,
            text=label,
            font=("Segoe UI", 10, "bold") if use_custom else ("Segoe UI", 9, "bold"),
            text_color="#334155" if use_custom else None,
        ).pack(anchor="w")
        ent = _mk_col_filter(cell, key, width)
        ent.pack(anchor="w", pady=(1, 0))

    for key, label, wd in (
        ("data", "Data", 120),
        ("origem", "Origem", 100),
        ("evento", "Evento", 100),
        ("encomenda", "Encomenda", 130),
        ("material", "Material", 120),
        ("esp", "Esp.", 90),
    ):
        _mk_labeled(filters_row_1, key, label, wd)

    for key, label, wd in (
        ("peca", "Peça", 170),
        ("operacao", "Operação", 130),
        ("operador", "Operador", 120),
        ("causa", "Causa", 190),
        ("detalhe", "Detalhe", 220),
    ):
        _mk_labeled(filters_row_2, key, label, wd)

    def _clear_col_filters():
        for _k, _v in (app.dashboard_ops_hist_col_vars or {}).items():
            try:
                _v.set("")
            except Exception:
                pass
        refresh_production_ops_history_window(app)

    Btn(
        filters_row_2,
        text="Limpar colunas",
        width=120 if use_custom else None,
        command=_clear_col_filters,
    ).pack(side="right", padx=8, pady=(18, 0))

    cols = ("data", "origem", "evento", "enc", "mat", "esp", "peca", "op", "operador", "ok", "nok", "tempo", "causa", "det")
    display_cols = ("data", "origem", "evento", "enc", "mat", "esp", "peca", "op", "operador", "ok", "nok", "tempo", "causa")
    tree = ttk.Treeview(body, columns=cols, displaycolumns=display_cols, show="headings", height=20, style="PulseHist.Treeview")
    app.dashboard_ops_hist_tree = tree
    headers = {
        "data": "Data",
        "origem": "Origem",
        "evento": "Evento",
        "enc": "Encomenda",
        "mat": "Material",
        "esp": "Esp.",
        "peca": "Peca",
        "op": "Operacao",
        "operador": "Operador",
        "ok": "OK",
        "nok": "NOK",
        "tempo": "Tempo(min)",
        "causa": "Causa",
        "det": "Detalhe",
    }
    widths = {
        "data": 132,
        "origem": 82,
        "evento": 88,
        "enc": 118,
        "mat": 105,
        "esp": 58,
        "peca": 210,
        "op": 128,
        "operador": 102,
        "ok": 52,
        "nok": 52,
        "tempo": 88,
        "causa": 300,
        "det": 520,
    }
    anchors = {
        "data": "center", "origem": "center", "evento": "center", "enc": "center",
        "mat": "w", "esp": "center", "peca": "w", "op": "center", "operador": "center",
        "ok": "center", "nok": "center", "tempo": "center", "causa": "w", "det": "w",
    }
    for c in cols:
        tree.heading(c, text=headers[c], anchor=anchors[c])
        tree.column(c, width=widths[c], anchor=anchors[c], stretch=(c in ("peca", "causa")))
    tree.tag_configure("even", background="#f8fbff")
    tree.tag_configure("odd", background="#eef4ff")
    tree.tag_configure("pause", background="#fff7ed")
    tree.tag_configure("stock", background="#eff6ff")
    tree.tag_configure("produto", background="#ecfeff")
    sy = ttk.Scrollbar(body, orient="vertical", command=tree.yview)
    sx = ttk.Scrollbar(body, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
    tree.pack(side="top", fill="both", expand=True, padx=(8, 0), pady=(8, 0))
    sy.pack(side="right", fill="y", padx=(0, 8), pady=(8, 0))
    sx.pack(side="bottom", fill="x", padx=(8, 8), pady=(0, 8))

    try:
        st = ttk.Style()
        st.configure("PulseHist.Treeview", rowheight=32, font=("Segoe UI", 10), background="#f8fbff", fieldbackground="#f8fbff")
        st.configure("PulseHist.Treeview.Heading", background="#0b1f66", foreground="#ffffff", font=("Segoe UI", 10, "bold"), relief="flat")
        st.map("PulseHist.Treeview", background=[("selected", "#cfe4ff")], foreground=[("selected", "#0f172a")])
        st.map(
            "PulseHist.Treeview.Heading",
            background=[
                ("active", "#102a7a"),
                ("pressed", "#102a7a"),
                ("selected", "#0b1f66"),
                ("focus", "#0b1f66"),
            ],
            foreground=[
                ("active", "#ffffff"),
                ("pressed", "#ffffff"),
                ("selected", "#ffffff"),
                ("focus", "#ffffff"),
            ],
            relief=[("pressed", "flat"), ("active", "flat"), ("selected", "flat")],
        )
    except Exception:
        pass

    def _open_selected_history_row(_event=None):
        try:
            sel = tree.selection()
            if not sel:
                return
            row = (getattr(app, "dashboard_ops_hist_row_map", {}) or {}).get(sel[0])
            if not isinstance(row, dict):
                return
            open_production_ops_history_detail_window(app, row)
        except Exception:
            pass

    try:
        tree.bind("<Double-1>", _open_selected_history_row)
        tree.bind("<Return>", _open_selected_history_row)
    except Exception:
        pass

    if isinstance(preset, dict):
        enc = str(preset.get("encomenda", "") or "").strip()
        if enc:
            app.dashboard_ops_hist_enc_var.set(enc)
        causa = str(preset.get("causa", "") or "").strip()
        if causa:
            app.dashboard_ops_hist_causa_var.set(causa)
        op = str(preset.get("operador", "") or "").strip()
        if op:
            app.dashboard_ops_hist_operador_var.set(op)

    def _on_close():
        try:
            try:
                win.grab_release()
            except Exception:
                pass
            try:
                if getattr(app, "dashboard_ops_hist_after_id", None):
                    app.root.after_cancel(app.dashboard_ops_hist_after_id)
            except Exception:
                pass
            app.dashboard_ops_hist_after_id = None
            win.destroy()
        finally:
            app.dashboard_ops_hist_win = None
            app.dashboard_ops_hist_tree = None

    try:
        win.protocol("WM_DELETE_WINDOW", _on_close)
    except Exception:
        pass

    refresh_production_ops_history_window(app)


def refresh_production_pulse_window(app):
    win = getattr(app, "dashboard_pulse_win", None)
    if win is None:
        return
    try:
        if not win.winfo_exists():
            app.dashboard_pulse_win = None
            return
    except Exception:
        app.dashboard_pulse_win = None
        return

    period_var = getattr(app, "dashboard_pulse_period_var", None)
    period_txt = str(period_var.get()).strip() if period_var is not None else "7 dias"
    period_map = {
        "Hoje": 1,
        "7 dias": 7,
        "30 dias": 30,
        "Tudo": 0,
    }
    period_days = period_map.get(period_txt, 7)
    year_var = getattr(app, "dashboard_pulse_year_var", None)
    year_txt = str(year_var.get()).strip() if year_var is not None else "Todos"
    year_filter = None if year_txt.lower() == "todos" else year_txt
    data = _compute_production_pulse_metrics(app, period_days=period_days, year_filter=year_filter)
    enc_filter_var = getattr(app, "dashboard_pulse_enc_var", None)
    enc_filter = str(enc_filter_var.get()).strip() if enc_filter_var is not None else "Todas"
    enc_vals = ["Todas"]
    try:
        nums = sorted(
            {
                str(e.get("numero", "") or "").strip()
                for e in list(app.data.get("encomendas", []) or [])
                if str(e.get("numero", "") or "").strip()
            }
        )
        enc_vals.extend(nums)
    except Exception:
        pass
    if enc_filter not in enc_vals:
        enc_filter = "Todas"
        if enc_filter_var is not None:
            try:
                enc_filter_var.set(enc_filter)
            except Exception:
                pass
    enc_menu = getattr(app, "dashboard_pulse_enc_menu", None)
    if enc_menu is not None:
        try:
            if ctk is not None and enc_menu.__class__.__name__.lower().startswith("ctk"):
                enc_menu.configure(values=enc_vals)
            else:
                enc_menu["values"] = enc_vals
        except Exception:
            pass

    view_events = list(data.get("ultimos_eventos", []) or [])
    view_tempo = list(data.get("pecas_tempo", []) or [])
    view_int = list(data.get("interrompidas", []) or [])
    view_hist = list(data.get("historico_tempo", []) or [])
    view_mode_var = getattr(app, "dashboard_pulse_view_var", None)
    view_mode = str(view_mode_var.get()).strip() if view_mode_var is not None else "So desvio"
    origem_var = getattr(app, "dashboard_pulse_origem_var", None)
    origem_mode = _norm_text(origem_var.get()) if origem_var is not None else "ambos"
    if enc_filter and enc_filter.lower() != "todas":
        view_events = [r for r in view_events if str(r.get("encomenda", "") or "").strip() == enc_filter]
        view_tempo = [r for r in view_tempo if str(r.get("encomenda", "") or "").strip() == enc_filter]
        view_int = [r for r in view_int if str(r.get("encomenda", "") or "").strip() == enc_filter]
        view_hist = [r for r in view_hist if str(r.get("encomenda", "") or "").strip() == enc_filter]
    if _norm_text(view_mode).startswith("so"):
        view_tempo = [r for r in view_tempo if bool(r.get("fora"))]
        view_hist = [r for r in view_hist if bool(r.get("fora"))]
    if "curso" in origem_mode:
        view_hist = []
    elif "histor" in origem_mode:
        view_tempo = []

    def _match_enc_year(enc_obj):
        if year_filter in (None, "", "Todos", "todos"):
            return True
        try:
            y_txt = str(year_filter).strip()
            d_ref = _safe_date(enc_obj.get("data_criacao") or enc_obj.get("inicio_producao") or enc_obj.get("data_entrega"))
            if d_ref is not None:
                return str(d_ref.year) == y_txt
            numero = str(enc_obj.get("numero", "") or "").strip()
            if len(numero) >= 9:
                parts = numero.split("-")
                if len(parts) >= 2 and str(parts[1]).isdigit():
                    return str(parts[1]) == y_txt
        except Exception:
            pass
        return True

    snapshot_hist = {}
    for enc_obj in list(app.data.get("encomendas", []) or []):
        enc_num = str(enc_obj.get("numero", "") or "").strip()
        if not enc_num:
            continue
        if enc_filter and enc_filter.lower() != "todas" and enc_num != enc_filter:
            continue
        if not _match_enc_year(enc_obj):
            continue
        snap_elapsed = 0.0
        snap_plan = 0.0
        snap_ops = 0
        for mat in list(enc_obj.get("materiais", []) or []):
            for esp in list(mat.get("espessuras", []) or []):
                plan_laser = max(0.0, _parse_float(esp.get("tempo_min", 0), 0))
                pecas = list(esp.get("pecas", []) or [])
                if plan_laser > 0:
                    laser_elapsed = 0.0
                    laser_has_work = False
                    for p in pecas:
                        elapsed_piece = max(0.0, _parse_float(p.get("tempo_producao_min", 0), 0))
                        if elapsed_piece <= 0 and not list(p.get("hist", []) or []):
                            continue
                        laser_has_work = True
                        laser_elapsed = max(laser_elapsed, elapsed_piece)
                    if laser_has_work:
                        snap_elapsed += laser_elapsed
                        snap_plan += plan_laser
                        snap_ops += 1
        if snap_elapsed <= 0 and snap_plan <= 0:
            continue
        delta_snap = snap_elapsed - snap_plan if snap_plan > 0 else 0.0
        snapshot_hist[enc_num] = {
            "encomenda": enc_num,
            "ops": snap_ops,
            "elapsed_min": round(snap_elapsed, 1),
            "plan_min": round(snap_plan, 1),
            "delta_min": round(delta_snap, 1),
            "fora": bool(snap_plan > 0 and delta_snap > 0.01),
        }
    if "curso" in origem_mode:
        snapshot_hist = {}
    elif _norm_text(view_mode).startswith("so"):
        snapshot_hist = {k: v for k, v in list(snapshot_hist.items()) if bool(v.get("fora"))}

    # Histórico consolidado por encomenda (para gestão de atraso global).
    hist_enc = {}
    for r in list(view_hist or []):
        enc = str(r.get("encomenda", "") or "").strip() or "-"
        g = hist_enc.setdefault(
            enc,
            {
                "encomenda": enc,
                "ops": 0,
                "elapsed_min": 0.0,
                "plan_min": 0.0,
                "delta_min": 0.0,
                "fora": False,
            },
        )
        g["ops"] += 1
        g["elapsed_min"] += _parse_float(r.get("elapsed_min", 0), 0)
        g["plan_min"] += _parse_float(r.get("plan_min", 0), 0)
        g["delta_min"] += _parse_float(r.get("delta_min", 0), 0)
        if bool(r.get("fora")):
            g["fora"] = True
    for enc, snap in list(snapshot_hist.items()):
        g = hist_enc.setdefault(
            enc,
            {
                "encomenda": enc,
                "ops": 0,
                "elapsed_min": 0.0,
                "plan_min": 0.0,
                "delta_min": 0.0,
                "fora": False,
            },
        )
        snap_elapsed = max(0.0, _parse_float(snap.get("elapsed_min", 0), 0))
        snap_plan = max(0.0, _parse_float(snap.get("plan_min", 0), 0))
        curr_elapsed = max(0.0, _parse_float(g.get("elapsed_min", 0), 0))
        curr_plan = max(0.0, _parse_float(g.get("plan_min", 0), 0))
        if snap_elapsed > curr_elapsed + 0.01 or (abs(snap_elapsed - curr_elapsed) <= 0.01 and snap_plan > curr_plan):
            g["elapsed_min"] = snap_elapsed
            g["plan_min"] = snap_plan
            g["ops"] = max(int(g.get("ops", 0) or 0), int(snap.get("ops", 0) or 0))
            g["delta_min"] = snap_elapsed - snap_plan if snap_plan > 0 else 0.0
            g["fora"] = bool(snap_plan > 0 and g["delta_min"] > 0.01)
    paragens_encomenda = dict(data.get("paragens_encomenda", {}) or {})
    for enc, g in list(hist_enc.items()):
        extra_stop = max(0.0, _parse_float(paragens_encomenda.get(enc, 0), 0))
        if extra_stop <= 0:
            continue
        g["elapsed_min"] += extra_stop
        if _parse_float(g.get("plan_min", 0), 0) > 0:
            g["delta_min"] = g["elapsed_min"] - _parse_float(g.get("plan_min", 0), 0)
            if _parse_float(g.get("delta_min", 0), 0) > 0.01:
                g["fora"] = True
    view_hist_encomenda = list(hist_enc.values())
    view_hist_encomenda.sort(
        key=lambda r: (
            1 if bool(r.get("fora")) else 0,
            _parse_float(r.get("delta_min", 0), 0),
            _parse_float(r.get("elapsed_min", 0), 0),
        ),
        reverse=True,
    )

    def _pulse_open_encomenda(numero, open_editor=False, highlight_ref=None):
        enc_num = str(numero or "").strip()
        if not enc_num:
            return
        try:
            if open_editor:
                app.open_encomenda_by_numero(enc_num, open_editor=True)
            else:
                app.open_encomenda_info_by_numero(enc_num, highlight_ref=highlight_ref)
        except Exception:
            pass

    def _pulse_extract_ref(row):
        refs = list(row.get("refs", []) or [])
        ref_txt = str((refs[0] if refs else (row.get("ref_interna", "") or row.get("peca", ""))) or "").strip()
        if ref_txt.lower().startswith("lote laser |"):
            ref_txt = ref_txt.split("|", 1)[1].strip()
        if " (+" in ref_txt:
            ref_txt = ref_txt.split(" (+", 1)[0].strip()
        return ref_txt

    def _pulse_open_desenho(row):
        enc_num = str(row.get("encomenda", "") or "").strip()
        ref_txt = _pulse_extract_ref(row)
        if not enc_num:
            return
        opened = False
        if ref_txt:
            try:
                opened = bool(app.open_peca_desenho_by_refs(numero=enc_num, ref_interna=ref_txt, silent=True))
            except Exception:
                opened = False
        if not opened:
            _pulse_open_encomenda(enc_num, open_editor=False, highlight_ref=ref_txt)

    # KPIs de pecas em curso/furo de tempo devem refletir o mesmo filtro visual.
    pecas_em_curso_view = len(view_tempo)
    pecas_fora_view = sum(1 for r in view_tempo if bool(r.get("fora"))) + sum(1 for r in view_hist if bool(r.get("fora")))
    top_paragens_view = list(data.get("top_paragens", []) or [])
    if enc_filter and enc_filter.lower() != "todas":
        top_paragens_view = [r for r in top_paragens_view if str(r.get("encomenda", "") or "").strip() == enc_filter]
    top_paragens_view.sort(key=lambda r: (_parse_float(r.get("ocorrencias", 0), 0), _parse_float(r.get("minutos", 0), 0)), reverse=True)
    down_min_view = sum(max(0.0, _parse_float(r.get("minutos", 0), 0)) for r in top_paragens_view)
    desvio_max_view = 0.0
    for r in view_tempo:
        desvio_max_view = max(desvio_max_view, _parse_float(r.get("delta_min", 0), 0))
    for r in view_hist:
        desvio_max_view = max(desvio_max_view, _parse_float(r.get("delta_min", 0), 0))
    for r in view_hist_encomenda:
        desvio_max_view = max(desvio_max_view, _parse_float(r.get("delta_min", 0), 0))

    vars_map = getattr(app, "dashboard_pulse_vars", {})
    mapping = {
        "oee": f"{data['oee']:.1f}%",
        "disponibilidade": f"{data['disponibilidade']:.1f}%",
        "performance": f"{data['performance']:.1f}%",
        "qualidade": f"{data['qualidade']:.1f}%",
        "start_ops": str(data["start_ops"]),
        "finish_ops": str(data["finish_ops"]),
        "qtd_ok": _fmt_num(data["qtd_ok"]),
        "qtd_nok": _fmt_num(data["qtd_nok"]),
        "pecas_em_curso": str(int(pecas_em_curso_view)),
        "pecas_fora_tempo": str(int(pecas_fora_view)),
        "desvio_max_min": f"{desvio_max_view:.1f} min",
        "down_min": f"{down_min_view:.1f} min",
        "andon": f"Prod: {data['andon_prod']} | Setup: {data['andon_setup']} | Espera: {data['andon_wait']} | Parada/Pausa: {data['andon_stop']}",
        "alertas": data["alertas"],
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "period": period_txt,
    }
    for k, v in mapping.items():
        var = vars_map.get(k)
        if var is not None:
            try:
                var.set(v)
            except Exception:
                pass

    # Semaforo visual por KPI (CustomTkinter only)
    def _metric_num(raw):
        txt = str(raw or "").strip().lower()
        txt = txt.replace("%", "").replace("eur", "").replace("min", "").replace(",", ".")
        txt = txt.strip()
        return _parse_float(txt, 0)

    def _pulse_status(key, val):
        v = _metric_num(val)
        if key == "oee":
            if v >= 75:
                return ("ok", "#16a34a")
            if v >= 60:
                return ("warn", "#d97706")
            return ("crit", "#dc2626")
        if key == "disponibilidade":
            if v >= 90:
                return ("ok", "#16a34a")
            if v >= 85:
                return ("warn", "#d97706")
            return ("crit", "#dc2626")
        if key == "performance":
            if v >= 90:
                return ("ok", "#16a34a")
            if v >= 80:
                return ("warn", "#d97706")
            return ("crit", "#dc2626")
        if key == "qualidade":
            if v >= 98:
                return ("ok", "#16a34a")
            if v >= 95:
                return ("warn", "#d97706")
            return ("crit", "#dc2626")
        if key == "down_min":
            lim = 480.0
            av_count = int(_parse_float(data.get("avaria_count", 0), 0))
            av_min = _parse_float(data.get("avaria_min", 0), 0)
            if av_count >= 2 or av_min >= 20:
                return ("crit", "#dc2626")
            if av_count > 0 and v <= (lim * 0.6):
                return ("warn", "#d97706")
            if v <= (lim * 0.6):
                return ("ok", "#16a34a")
            if v <= lim:
                return ("warn", "#d97706")
            return ("crit", "#dc2626")
        if key == "desvio_max_min":
            if v <= 0:
                return ("ok", "#16a34a")
            if v <= 30:
                return ("warn", "#d97706")
            return ("crit", "#dc2626")
        return ("na", "#334155")

    card_widgets = getattr(app, "dashboard_pulse_card_widgets", {}) or {}
    severity_rank = {"ok": 1, "warn": 2, "crit": 3}
    priority_rules = {
        "oee": "OEE global baixo. Rever disponibilidade/performance/qualidade.",
        "disponibilidade": "Disponibilidade baixa. Atacar paragens e setups.",
        "performance": "Performance baixa. Ver tempos de ciclo e microparagens.",
        "qualidade": "Qualidade critica. Investigar causas de NOK/refugo.",
        "down_min": "Tempo de paragens elevado no mes.",
        "desvio_max_min": "Pecas fora de tempo planeado. Ver gargalo da operacao em curso.",
    }
    best_key = None
    best_status = "ok"
    best_value = "0"
    for key in ("oee", "disponibilidade", "performance", "qualidade", "down_min", "desvio_max_min"):
        row = card_widgets.get(key) or {}
        value_lbl = row.get("value")
        chip_lbl = row.get("chip")
        card = row.get("card")
        if value_lbl is None or chip_lbl is None or card is None:
            continue
        status, color = _pulse_status(key, mapping.get(key, "0"))
        if best_key is None or severity_rank.get(status, 0) > severity_rank.get(best_status, 0):
            best_key = key
            best_status = status
            best_value = mapping.get(key, "0")
        try:
            value_lbl.configure(text_color=color)
        except Exception:
            pass
        try:
            chip_text = "OK" if status == "ok" else ("ATENCAO" if status == "warn" else ("CRITICO" if status == "crit" else "-"))
            chip_lbl.configure(text=chip_text, text_color="#ffffff", fg_color=color)
        except Exception:
            pass
        try:
            bg = "#ecfdf3" if status == "ok" else ("#fff7ed" if status == "warn" else "#fef2f2")
            card.configure(fg_color=bg)
        except Exception:
            pass

    priority_var = vars_map.get("prioridade")
    pri_widgets = getattr(app, "dashboard_pulse_priority_widgets", {}) or {}
    pri_label = pri_widgets.get("label")
    pri_frame = pri_widgets.get("frame")
    if best_key is None:
        best_key = "oee"
    if best_status == "crit":
        pri_color = "#dc2626"
        pri_bg = "#fef2f2"
        prefix = "PRIORIDADE CRITICA"
    elif best_status == "warn":
        pri_color = "#d97706"
        pri_bg = "#fff7ed"
        prefix = "PRIORIDADE ALTA"
    else:
        pri_color = "#16a34a"
        pri_bg = "#ecfdf3"
        prefix = "OPERACAO ESTAVEL"
    pri_msg = priority_rules.get(best_key, "Acompanhar operacao.")
    if best_key == "desvio_max_min":
        fonte = list(view_tempo or []) + list(view_hist or [])
        if fonte:
            gargalo = max(fonte, key=lambda r: _parse_float(r.get("delta_min", 0), 0))
            pri_msg = (
                f"{priority_rules.get(best_key)} "
                f"Gargalo: {gargalo.get('encomenda','-')} / {gargalo.get('peca','-')} / {gargalo.get('operacao','-')}."
            )
    pri_txt = f"{prefix}: {pri_msg} ({best_key.upper()}: {best_value})"
    if priority_var is not None:
        try:
            priority_var.set(pri_txt)
        except Exception:
            pass
    try:
        if pri_frame is not None:
            pri_frame.configure(fg_color=pri_bg)
    except Exception:
        pass
    try:
        if pri_label is not None:
            pri_label.configure(text_color=pri_color)
    except Exception:
        pass

    # Alerta critico (popup + som) com cooldown para evitar repeticao excessiva.
    try:
        cfg = getattr(app, "dashboard_pulse_alert_cfg", {}) or {}
        enable_popup = bool(cfg.get("popup", True))
        enable_sound = bool(cfg.get("sound", True))
        cooldown_sec = int(cfg.get("cooldown_sec", 90) or 90)
    except Exception:
        enable_popup = True
        enable_sound = True
        cooldown_sec = 90
    last_alert = getattr(app, "dashboard_pulse_last_alert", {}) or {}
    now_ts = time.time()
    last_msg = str(last_alert.get("msg", "") or "")
    last_ts = float(last_alert.get("ts", 0.0) or 0.0)
    if best_status == "crit":
        should_alert = (pri_txt != last_msg) or ((now_ts - last_ts) >= max(10, cooldown_sec))
        if should_alert:
            if enable_sound:
                try:
                    import winsound  # type: ignore
                    winsound.MessageBeep(winsound.MB_ICONHAND)
                except Exception:
                    pass
            if enable_popup:
                try:
                    parent_win = getattr(app, "dashboard_pulse_win", None)
                    if parent_win is not None and parent_win.winfo_exists():
                        try:
                            parent_win.lift()
                            parent_win.focus_force()
                            parent_win.attributes("-topmost", True)
                        except Exception:
                            pass
                        messagebox.showwarning("Production Pulse - Alerta Critico", pri_txt, parent=parent_win)
                        try:
                            parent_win.attributes("-topmost", False)
                            parent_win.lift()
                            parent_win.focus_force()
                        except Exception:
                            pass
                    else:
                        messagebox.showwarning("Production Pulse - Alerta Critico", pri_txt)
                except Exception:
                    pass
            app.dashboard_pulse_last_alert = {"msg": pri_txt, "ts": now_ts}

    host = getattr(app, "dashboard_pulse_paragens_host", None)
    if host is not None and ctk is not None:
        try:
            for child in host.winfo_children():
                child.destroy()
            rows = list(top_paragens_view[:8])
            if not rows:
                ctk.CTkLabel(
                    host,
                    text="Sem causas no filtro atual.",
                    text_color="#64748b",
                    font=("Segoe UI", 12),
                    anchor="w",
                ).pack(fill="x", padx=8, pady=6)
                rows = []
            if rows:
                hdr = ctk.CTkFrame(host, fg_color="#0b1f66", corner_radius=6, height=28)
                hdr.pack(fill="x", padx=4, pady=(2, 2))
                hdr.pack_propagate(False)
                col_cfg = (
                    ("Causa", 220, "w"),
                    ("Encomenda", 120, "center"),
                    ("Operador", 110, "center"),
                    ("Data", 145, "center"),
                    ("Ocorr.", 60, "center"),
                    ("Min", 70, "center"),
                )
                for txt, width, anc in col_cfg:
                    l = ctk.CTkLabel(hdr, text=txt, anchor=anc, width=width, font=("Segoe UI", 11, "bold"), text_color="#ffffff")
                    l.pack(side="left", padx=(8, 2), pady=3)
            for idx, row in enumerate(rows):
                bg = "#f8fbff" if idx % 2 == 0 else "#eef4ff"
                line = ctk.CTkFrame(host, fg_color=bg, corner_radius=6, height=30)
                line.pack(fill="x", padx=4, pady=1)
                line.pack_propagate(False)
                vals = (
                    (str(row.get("causa", "") or "-"), "w"),
                    (str(row.get("encomenda", "") or "-"), "center"),
                    (str(row.get("operador", "") or "-"), "center"),
                    (str(row.get("data_ultima", "") or "-"), "center"),
                    (str(int(_parse_float(row.get("ocorrencias", 0), 0))), "center"),
                    (f"{_parse_float(row.get('minutos', 0), 0):.1f}", "center"),
                )
                for (txt, width, _hdr_anc), (val, anc) in zip(col_cfg, vals):
                    l = ctk.CTkLabel(line, text=val, anchor=anc, width=width, font=("Segoe UI", 11), text_color="#0f172a")
                    l.pack(side="left", padx=(8, 2), pady=3)
                    _bind_widget_click(
                        l,
                        lambda rw=row: open_production_ops_history_window(
                            app,
                            {
                                "encomenda": str(rw.get("encomenda", "") or "").strip(),
                                "operador": str(rw.get("operador", "") or "").strip(),
                                "causa": str(rw.get("causa", "") or "").strip(),
                            },
                        ),
                    )
                _bind_widget_click(
                    line,
                    lambda rw=row: open_production_ops_history_window(
                        app,
                        {
                            "encomenda": str(rw.get("encomenda", "") or "").strip(),
                            "operador": str(rw.get("operador", "") or "").strip(),
                            "causa": str(rw.get("causa", "") or "").strip(),
                        },
                    ),
                )
        except Exception:
            pass

    tree_par = getattr(app, "dashboard_pulse_paragens_tree", None)
    if tree_par is not None:
        try:
            for iid in tree_par.get_children():
                tree_par.delete(iid)
            for idx, row in enumerate(top_paragens_view[:8]):
                tag = "odd" if idx % 2 else "even"
                tree_par.insert(
                    "",
                    "end",
                    values=(
                        row.get("causa", ""),
                        row.get("encomenda", "-"),
                        row.get("operador", "-"),
                        row.get("data_ultima", "-"),
                        row.get("ocorrencias", 0),
                        f"{_parse_float(row.get('minutos', 0), 0):.1f}",
                    ),
                    tags=(tag,),
                )
            def _on_open_from_top(_event=None):
                try:
                    sel = tree_par.selection()
                    if not sel:
                        return
                    vals = tree_par.item(sel[0], "values")
                    enc = str(vals[1] if len(vals) > 1 else "" or "").strip()
                    op = str(vals[2] if len(vals) > 2 else "" or "").strip()
                    causa = str(vals[0] if len(vals) > 0 else "" or "").strip()
                    open_production_ops_history_window(
                        app,
                        {"encomenda": enc, "operador": op, "causa": causa},
                    )
                except Exception:
                    pass
            tree_par.bind("<Double-1>", _on_open_from_top)
        except Exception:
            pass

    host_events = getattr(app, "dashboard_pulse_events_host", None)
    if host_events is not None and ctk is not None:
        try:
            for child in host_events.winfo_children():
                child.destroy()
            for idx, row in enumerate(view_events[:18]):
                bg = "#f8fbff" if idx % 2 == 0 else "#eef4ff"
                ev = ctk.CTkFrame(host_events, fg_color=bg, corner_radius=6, height=30)
                ev.pack(fill="x", padx=4, pady=1)
                ev.pack_propagate(False)
                posto_txt = str(row.get("posto", "") or "").strip()
                posto_part = f" | {posto_txt}" if posto_txt else ""
                txt = f"{row.get('created_at', '')} | {row.get('evento', '')} | {row.get('encomenda', '')} | {row.get('peca', '')} | {row.get('operador', '')}{posto_part}"
                ctk.CTkLabel(ev, text=txt, anchor="w", font=("Consolas", 11), text_color="#0f172a").pack(fill="both", expand=True, padx=8)
        except Exception:
            pass

    tree_ev = getattr(app, "dashboard_pulse_events_tree", None)
    if tree_ev is not None:
        try:
            for iid in tree_ev.get_children():
                tree_ev.delete(iid)
            for idx, row in enumerate(view_events[:18]):
                tag = "odd" if idx % 2 else "even"
                tree_ev.insert(
                    "",
                    "end",
                    values=(
                        row.get("created_at", ""),
                        row.get("evento", ""),
                        row.get("encomenda", ""),
                        row.get("peca", ""),
                        row.get("operador", ""),
                        row.get("posto", ""),
                    ),
                    tags=(tag,),
                )
        except Exception:
            pass

    host_int = getattr(app, "dashboard_pulse_interrupted_host", None)
    if host_int is not None and ctk is not None:
        try:
            for child in host_int.winfo_children():
                child.destroy()
            for idx, row in enumerate(view_int[:40]):
                bg = "#fef2f2" if idx % 2 == 0 else "#fee2e2"
                it = ctk.CTkFrame(host_int, fg_color=bg, corner_radius=6, height=30)
                it.pack(fill="x", padx=4, pady=1)
                it.pack_propagate(False)
                posto_txt = str(row.get("posto", "") or "").strip() or "-"
                txt = f"{row.get('encomenda','')} | {row.get('peca','')} | {row.get('operador','-')} @ {posto_txt} | {row.get('motivo','-')}"
                ctk.CTkLabel(it, text=txt, anchor="w", font=("Consolas", 11), text_color="#7f1d1d").pack(fill="both", expand=True, padx=8)
        except Exception:
            pass

    host_rt = getattr(app, "dashboard_pulse_running_host", None)
    if host_rt is not None and ctk is not None:
        for child in host_rt.winfo_children():
            child.destroy()
        try:
            if not view_tempo and not view_hist:
                ctk.CTkLabel(
                    host_rt,
                    text="Sem pecas em curso para o filtro atual.",
                    text_color="#64748b",
                    font=("Segoe UI", 12),
                    anchor="w",
                ).pack(fill="x", padx=8, pady=6)
            if view_tempo:
                hdr_run = ctk.CTkFrame(host_rt, fg_color="#dbeafe", corner_radius=6, height=26)
                hdr_run.pack(fill="x", padx=4, pady=(4, 1))
                hdr_run.pack_propagate(False)
                ctk.CTkLabel(hdr_run, text="Em curso", anchor="w", font=("Segoe UI", 12, "bold"), text_color="#1e3a8a").pack(fill="both", expand=True, padx=8)
            last_enc = None
            for idx, row in enumerate(view_tempo[:50]):
                enc_num = str(row.get("encomenda", "") or "")
                if enc_filter.lower() == "todas" and enc_num and enc_num != last_enc:
                    sep = ctk.CTkFrame(host_rt, fg_color="#dbeafe", corner_radius=6, height=26)
                    sep.pack(fill="x", padx=4, pady=(4, 1))
                    sep.pack_propagate(False)
                    sep_lbl = ctk.CTkLabel(sep, text=f"Encomenda: {enc_num}", anchor="w", font=("Segoe UI", 12, "bold"), text_color="#1e3a8a")
                    sep_lbl.pack(fill="both", expand=True, padx=8)
                    _bind_widget_click(sep, lambda n=enc_num: _pulse_open_encomenda(n, open_editor=False))
                    _bind_widget_click(sep_lbl, lambda n=enc_num: _pulse_open_encomenda(n, open_editor=False))
                    last_enc = enc_num
                fora = bool(row.get("fora"))
                elapsed = _parse_float(row.get("elapsed_min", 0), 0)
                plan = _parse_float(row.get("plan_min", 0), 0)
                near = bool((not fora) and plan > 0 and elapsed >= (plan * 0.9))
                bg = "#fee2e2" if fora else ("#fffbeb" if near else ("#f8fbff" if idx % 2 == 0 else "#eef4ff"))
                fg = "#991b1b" if fora else ("#a16207" if near else "#0f172a")
                it = ctk.CTkFrame(host_rt, fg_color=bg, corner_radius=6, height=30)
                it.pack(fill="x", padx=4, pady=1)
                it.pack_propagate(False)
                d = _parse_float(row.get("delta_min", 0), 0)
                desvio_txt = f"+{d:.1f}m" if d > 0 else f"{d:.1f}m"
                posto_txt = str(row.get("posto", "") or "").strip()
                posto_part = f" | Posto: {posto_txt}" if posto_txt else ""
                elapsed_v = _parse_float(row.get("elapsed_min", 0), 0)
                plan_v = _parse_float(row.get("plan_min", 0), 0)
                txt = (
                    f"{row.get('encomenda','')} | {row.get('peca','')} | {row.get('operacao','-')} | "
                    f"T: {elapsed_v:.1f}m / Plan: {plan_v:.1f}m | Desvio: {desvio_txt}{posto_part}"
                )
                it_lbl = ctk.CTkLabel(it, text=txt, anchor="w", font=("Consolas", 11), text_color=fg)
                it_lbl.pack(fill="both", expand=True, padx=8)
                _bind_widget_click(
                    it,
                    lambda rw=row: _pulse_open_encomenda(
                        str(rw.get("encomenda", "") or "").strip(),
                        open_editor=False,
                        highlight_ref=_pulse_extract_ref(rw),
                    ),
                )
                _bind_widget_click(
                    it_lbl,
                    lambda rw=row: _pulse_open_encomenda(
                        str(rw.get("encomenda", "") or "").strip(),
                        open_editor=False,
                        highlight_ref=_pulse_extract_ref(rw),
                    ),
                )
                try:
                    it.bind("<Double-1>", lambda _e, rw=row: _pulse_open_desenho(rw))
                    it_lbl.bind("<Double-1>", lambda _e, rw=row: _pulse_open_desenho(rw))
                except Exception:
                    pass
            if view_hist_encomenda:
                hdr_hist = ctk.CTkFrame(host_rt, fg_color="#e2e8f0", corner_radius=6, height=26)
                hdr_hist.pack(fill="x", padx=4, pady=(6, 1))
                hdr_hist.pack_propagate(False)
                ctk.CTkLabel(hdr_hist, text="Historico consolidado por encomenda (desvio)", anchor="w", font=("Segoe UI", 12, "bold"), text_color="#334155").pack(fill="both", expand=True, padx=8)
            for idx, row in enumerate(view_hist_encomenda[:80]):
                fora = bool(row.get("fora"))
                bg = "#fee2e2" if fora else ("#f8fafc" if idx % 2 == 0 else "#eef2f7")
                fg = "#991b1b" if fora else "#1e293b"
                it = ctk.CTkFrame(host_rt, fg_color=bg, corner_radius=6, height=30)
                it.pack(fill="x", padx=4, pady=1)
                it.pack_propagate(False)
                d = _parse_float(row.get("delta_min", 0), 0)
                desvio_txt = f"+{d:.1f}m" if d > 0 else f"{d:.1f}m"
                txt = (
                    f"Encomenda: {row.get('encomenda','-')} | Ops concluidas: {int(_parse_float(row.get('ops',0),0))} | "
                    f"T total: {_parse_float(row.get('elapsed_min',0),0):.1f}m / "
                    f"Plan total: {_parse_float(row.get('plan_min',0),0):.1f}m | Desvio: {desvio_txt}"
                )
                it_lbl = ctk.CTkLabel(it, text=txt, anchor="w", font=("Consolas", 11), text_color=fg)
                it_lbl.pack(fill="both", expand=True, padx=8)
                _bind_widget_click(
                    it,
                    lambda rw=row: _pulse_open_encomenda(str(rw.get("encomenda", "") or "").strip(), open_editor=False),
                )
                _bind_widget_click(
                    it_lbl,
                    lambda rw=row: _pulse_open_encomenda(str(rw.get("encomenda", "") or "").strip(), open_editor=False),
                )
                try:
                    it.bind(
                        "<Double-1>",
                        lambda _e, rw=row: open_production_ops_history_window(
                            app,
                            {"encomenda": str(rw.get("encomenda", "") or "").strip()},
                        ),
                    )
                    it_lbl.bind(
                        "<Double-1>",
                        lambda _e, rw=row: open_production_ops_history_window(
                            app,
                            {"encomenda": str(rw.get("encomenda", "") or "").strip()},
                        ),
                    )
                except Exception:
                    pass
        except Exception as ex:
            ctk.CTkLabel(
                host_rt,
                text=f"Erro ao renderizar lista de pecas em curso: {ex}",
                text_color="#991b1b",
                font=("Segoe UI", 12),
                anchor="w",
            ).pack(fill="x", padx=8, pady=6)

    tree_rt = getattr(app, "dashboard_pulse_running_tree", None)
    if tree_rt is not None:
        try:
            row_map = {}
            for iid in tree_rt.get_children():
                tree_rt.delete(iid)
            for idx, row in enumerate(view_tempo[:50]):
                elapsed = _parse_float(row.get("elapsed_min", 0), 0)
                plan = _parse_float(row.get("plan_min", 0), 0)
                near = bool((not bool(row.get("fora"))) and plan > 0 and elapsed >= (plan * 0.9))
                tag = "warn" if bool(row.get("fora")) else ("near" if near else ("odd" if idx % 2 else "even"))
                d = _parse_float(row.get("delta_min", 0), 0)
                desvio_txt = f"+{d:.1f}" if d > 0 else f"{d:.1f}"
                iid = tree_rt.insert(
                    "",
                    "end",
                    values=(
                        row.get("encomenda", ""),
                        row.get("peca", ""),
                        row.get("operacao", ""),
                        f"{_parse_float(row.get('elapsed_min', 0), 0):.1f}",
                        f"{_parse_float(row.get('plan_min', 0), 0):.1f}",
                        desvio_txt,
                    ),
                    tags=(tag,),
                )
                row_map[iid] = dict(row)
            app.dashboard_pulse_running_row_map = row_map

            def _on_open_from_running(_event=None):
                try:
                    sel = tree_rt.selection()
                    if not sel:
                        return
                    row = (getattr(app, "dashboard_pulse_running_row_map", {}) or {}).get(sel[0])
                    if isinstance(row, dict):
                        _pulse_open_desenho(row)
                except Exception:
                    pass

            tree_rt.bind("<Double-1>", _on_open_from_running)
            tree_rt.bind("<Return>", _on_open_from_running)
        except Exception:
            pass

    # Auto-refresh em tempo real (janela aberta): evita clicar manualmente.
    try:
        prev_after = getattr(app, "dashboard_pulse_after_id", None)
        if prev_after:
            try:
                app.root.after_cancel(prev_after)
            except Exception:
                pass
        if getattr(app, "dashboard_pulse_win", None) is not None:
            app.dashboard_pulse_after_id = app.root.after(3000, lambda: refresh_production_pulse_window(app))
    except Exception:
        pass

    tree_int = getattr(app, "dashboard_pulse_interrupted_tree", None)
    if tree_int is not None:
        try:
            for iid in tree_int.get_children():
                tree_int.delete(iid)
            for idx, row in enumerate(view_int[:40]):
                tag = "odd" if idx % 2 else "even"
                tree_int.insert(
                    "",
                    "end",
                    values=(
                        row.get("encomenda", ""),
                        row.get("peca", ""),
                        row.get("operador", ""),
                        row.get("motivo", ""),
                    ),
                    tags=(tag,),
                )
        except Exception:
            pass


def open_production_pulse_window(app):
    use_custom = ctk is not None and os.environ.get("USE_CUSTOM_DASH_PULSE", "1") != "0"
    old = getattr(app, "dashboard_pulse_win", None)
    try:
        if old is not None and old.winfo_exists():
            old.lift()
            old.focus_force()
            refresh_production_pulse_window(app)
            return
    except Exception:
        pass

    Win = ctk.CTkToplevel if use_custom else Toplevel
    Frame = ctk.CTkFrame if use_custom else ttk.Frame
    Label = ctk.CTkLabel if use_custom else ttk.Label
    Btn = ctk.CTkButton if use_custom else ttk.Button

    win = Win(app.root)
    app.dashboard_pulse_win = win
    try:
        win.title("Production Pulse - OEE e Andon")
        win.transient(app.root)
        _center_window_on_screen(win, 1420, 900, min_w=1240, min_h=780)
        if use_custom:
            win.configure(fg_color="#eef2f7")
        win.lift()
        win.focus_force()
        try:
            win.grab_set()
        except Exception:
            pass
        try:
            win.attributes("-topmost", True)
        except Exception:
            pass
    except Exception:
        pass

    app.dashboard_pulse_vars = {k: StringVar(value="-") for k in (
        "oee", "disponibilidade", "performance", "qualidade",
        "start_ops", "finish_ops", "qtd_ok", "qtd_nok", "pecas_em_curso", "pecas_fora_tempo", "desvio_max_min", "down_min",
        "andon", "alertas", "updated_at", "period", "prioridade",
    )}
    app.dashboard_pulse_card_widgets = {}
    app.dashboard_pulse_priority_widgets = {}
    app.dashboard_pulse_last_alert = {"msg": "", "ts": 0.0}
    app.dashboard_pulse_alert_cfg = {"popup": True, "sound": True, "cooldown_sec": 90}
    app.dashboard_pulse_popup_var = BooleanVar(value=True)
    app.dashboard_pulse_sound_var = BooleanVar(value=True)
    app.dashboard_pulse_cooldown_var = StringVar(value="90s")
    app.dashboard_pulse_period_var = StringVar(value="7 dias")
    app.dashboard_pulse_enc_var = StringVar(value="Todas")
    app.dashboard_pulse_view_var = StringVar(value="So desvio")
    app.dashboard_pulse_origem_var = StringVar(value="Ambos")
    app.dashboard_pulse_year_var = StringVar(value=str(datetime.now().year))
    app.dashboard_pulse_enc_menu = None

    def _sync_alert_cfg(*_args):
        try:
            pop = bool(app.dashboard_pulse_popup_var.get())
        except Exception:
            pop = True
        try:
            snd = bool(app.dashboard_pulse_sound_var.get())
        except Exception:
            snd = True
        try:
            txt = str(app.dashboard_pulse_cooldown_var.get() or "90s")
            sec = int(str(txt).replace("s", "").strip() or 90)
        except Exception:
            sec = 90
        app.dashboard_pulse_alert_cfg = {"popup": pop, "sound": snd, "cooldown_sec": sec}

    top = Frame(win, fg_color="transparent") if use_custom else Frame(win)
    top.pack(fill="x", padx=14, pady=(12, 8))
    Label(top, text="Production Pulse", font=("Segoe UI", 24, "bold") if use_custom else ("Segoe UI", 16, "bold"), text_color="#0f172a" if use_custom else None).pack(side="left")
    Label(top, text="OEE | Andon | Paragens | Alertas", font=("Segoe UI", 13, "bold") if use_custom else ("Segoe UI", 10, "bold"), text_color="#475569" if use_custom else None).pack(side="left", padx=(12, 0), pady=(6, 0))
    Label(top, textvariable=app.dashboard_pulse_vars["updated_at"], text_color="#64748b" if use_custom else None, font=("Segoe UI", 12, "bold") if use_custom else None).pack(side="right")

    actions = Frame(win, fg_color="#ffffff" if use_custom else "transparent", corner_radius=10 if use_custom else None, border_width=1 if use_custom else 0, border_color="#d7dee9" if use_custom else None) if use_custom else Frame(win)
    actions.pack(fill="x", padx=14, pady=(0, 10))
    Label(actions, text="Periodo:", font=("Segoe UI", 12, "bold") if use_custom else ("Segoe UI", 10, "bold")).pack(side="left", padx=(4, 2))
    if use_custom and ctk is not None:
        ctk.CTkOptionMenu(
            actions,
            values=["Hoje", "7 dias", "30 dias", "Tudo"],
            variable=app.dashboard_pulse_period_var,
            width=120,
            command=lambda _v=None: refresh_production_pulse_window(app),
        ).pack(side="left", padx=(0, 8))
    else:
        cmb = ttk.Combobox(actions, values=["Hoje", "7 dias", "30 dias", "Tudo"], width=12, state="readonly", textvariable=app.dashboard_pulse_period_var)
        cmb.pack(side="left", padx=(0, 8))
        cmb.bind("<<ComboboxSelected>>", lambda _e: refresh_production_pulse_window(app))
    Label(actions, text="Ano:", font=("Segoe UI", 12, "bold") if use_custom else ("Segoe UI", 10, "bold")).pack(side="left", padx=(8, 2))
    year_values = ["Todos"] + [str(y) for y in range(datetime.now().year - 3, datetime.now().year + 2)]
    if use_custom and ctk is not None:
        ctk.CTkOptionMenu(
            actions,
            values=year_values,
            variable=app.dashboard_pulse_year_var,
            width=100,
            command=lambda _v=None: refresh_production_pulse_window(app),
        ).pack(side="left", padx=(0, 8))
    else:
        cmb_year = ttk.Combobox(actions, values=year_values, width=8, state="readonly", textvariable=app.dashboard_pulse_year_var)
        cmb_year.pack(side="left", padx=(0, 8))
        cmb_year.bind("<<ComboboxSelected>>", lambda _e: refresh_production_pulse_window(app))
    Label(actions, text="Encomenda:", font=("Segoe UI", 12, "bold") if use_custom else ("Segoe UI", 10, "bold")).pack(side="left", padx=(8, 2))
    if use_custom and ctk is not None:
        app.dashboard_pulse_enc_menu = ctk.CTkOptionMenu(
            actions,
            values=["Todas"],
            variable=app.dashboard_pulse_enc_var,
            width=170,
            command=lambda _v=None: refresh_production_pulse_window(app),
        )
        app.dashboard_pulse_enc_menu.pack(side="left", padx=(0, 8))
    else:
        cmb_enc = ttk.Combobox(actions, values=["Todas"], width=16, state="readonly", textvariable=app.dashboard_pulse_enc_var)
        cmb_enc.pack(side="left", padx=(0, 8))
        cmb_enc.bind("<<ComboboxSelected>>", lambda _e: refresh_production_pulse_window(app))
        app.dashboard_pulse_enc_menu = cmb_enc
    Label(actions, text="Visao:", font=("Segoe UI", 12, "bold") if use_custom else ("Segoe UI", 10, "bold")).pack(side="left", padx=(8, 2))
    if use_custom and ctk is not None:
        ctk.CTkOptionMenu(
            actions,
            values=["So desvio", "Todas"],
            variable=app.dashboard_pulse_view_var,
            width=120,
            command=lambda _v=None: refresh_production_pulse_window(app),
        ).pack(side="left", padx=(0, 8))
    else:
        cmb_view = ttk.Combobox(actions, values=["So desvio", "Todas"], width=11, state="readonly", textvariable=app.dashboard_pulse_view_var)
        cmb_view.pack(side="left", padx=(0, 8))
        cmb_view.bind("<<ComboboxSelected>>", lambda _e: refresh_production_pulse_window(app))
    Label(actions, text="Origem:", font=("Segoe UI", 12, "bold") if use_custom else ("Segoe UI", 10, "bold")).pack(side="left", padx=(8, 2))
    if use_custom and ctk is not None:
        ctk.CTkOptionMenu(
            actions,
            values=["Ambos", "Em curso", "Historico"],
            variable=app.dashboard_pulse_origem_var,
            width=120,
            command=lambda _v=None: refresh_production_pulse_window(app),
        ).pack(side="left", padx=(0, 8))
    else:
        cmb_orig = ttk.Combobox(actions, values=["Ambos", "Em curso", "Historico"], width=11, state="readonly", textvariable=app.dashboard_pulse_origem_var)
        cmb_orig.pack(side="left", padx=(0, 8))
        cmb_orig.bind("<<ComboboxSelected>>", lambda _e: refresh_production_pulse_window(app))
    Label(actions, text="Alertas:", font=("Segoe UI", 12, "bold") if use_custom else ("Segoe UI", 10, "bold")).pack(side="left", padx=(8, 2))
    if use_custom and ctk is not None:
        ctk.CTkSwitch(
            actions,
            text="Popup",
            variable=app.dashboard_pulse_popup_var,
            command=_sync_alert_cfg,
            onvalue=True,
            offvalue=False,
            width=80,
        ).pack(side="left", padx=(0, 4))
        ctk.CTkSwitch(
            actions,
            text="Som",
            variable=app.dashboard_pulse_sound_var,
            command=_sync_alert_cfg,
            onvalue=True,
            offvalue=False,
            width=70,
        ).pack(side="left", padx=(0, 6))
        ctk.CTkOptionMenu(
            actions,
            values=["30s", "60s", "90s", "120s", "180s"],
            variable=app.dashboard_pulse_cooldown_var,
            width=90,
            command=lambda _v=None: _sync_alert_cfg(),
        ).pack(side="left", padx=(0, 8))
    else:
        cb_pop = ttk.Checkbutton(actions, text="Popup", variable=app.dashboard_pulse_popup_var, command=_sync_alert_cfg)
        cb_pop.pack(side="left", padx=(0, 4))
        cb_snd = ttk.Checkbutton(actions, text="Som", variable=app.dashboard_pulse_sound_var, command=_sync_alert_cfg)
        cb_snd.pack(side="left", padx=(0, 6))
        cmb_cd = ttk.Combobox(actions, values=["30s", "60s", "90s", "120s", "180s"], width=7, state="readonly", textvariable=app.dashboard_pulse_cooldown_var)
        cmb_cd.pack(side="left", padx=(0, 8))
        cmb_cd.bind("<<ComboboxSelected>>", lambda _e: _sync_alert_cfg())
    Btn(actions, text="Atualizar", command=lambda: refresh_production_pulse_window(app), width=130 if use_custom else None).pack(side="left", padx=6, pady=8)
    Btn(
        actions,
        text="Historico Operacoes",
        command=lambda: open_production_ops_history_window(app),
        width=175 if use_custom else None,
    ).pack(side="left", padx=6, pady=8)
    Btn(actions, text="Fechar", command=win.destroy, width=120 if use_custom else None).pack(side="right", padx=6, pady=8)
    _sync_alert_cfg()

    pri = Frame(win, fg_color="#ecfdf3" if use_custom else None, corner_radius=10 if use_custom else None, border_width=1 if use_custom else 0, border_color="#bbf7d0" if use_custom else None) if use_custom else Frame(win)
    pri.pack(fill="x", padx=14, pady=(0, 10))
    pri_lbl = Label(
        pri,
        textvariable=app.dashboard_pulse_vars["prioridade"],
        font=("Segoe UI", 13, "bold") if use_custom else ("Segoe UI", 10, "bold"),
        text_color="#166534" if use_custom else None,
        anchor="w",
    )
    if use_custom:
        pri_lbl.pack(fill="x", padx=12, pady=8)
    else:
        pri_lbl.pack(fill="x", padx=8, pady=6)
    app.dashboard_pulse_priority_widgets = {"frame": pri, "label": pri_lbl}

    cards_wrap = Frame(win, fg_color="transparent") if use_custom else Frame(win)
    cards_wrap.pack(fill="x", padx=14, pady=(0, 8))
    for i in range(5):
        try:
            cards_wrap.grid_columnconfigure(i, weight=1)
        except Exception:
            pass
    cards = [
        ("OEE", "oee"),
        ("Disponibilidade", "disponibilidade"),
        ("Performance", "performance"),
        ("Qualidade", "qualidade"),
        ("Paragens", "down_min"),
        ("Start Op", "start_ops"),
        ("Finish Op", "finish_ops"),
        ("Pecas OK", "qtd_ok"),
        ("Pecas NOK", "qtd_nok"),
        ("Pecas em Curso", "pecas_em_curso"),
        ("Fora do Tempo", "pecas_fora_tempo"),
        ("Desvio Max", "desvio_max_min"),
        ("Andon", "andon"),
    ]
    for idx, (title, key) in enumerate(cards):
        r, c = divmod(idx, 5)
        card = Frame(cards_wrap, fg_color="#ffffff" if use_custom else None, corner_radius=12 if use_custom else None, border_width=1 if use_custom else 0, border_color="#dce3ee" if use_custom else None) if use_custom else ttk.LabelFrame(cards_wrap, text=title)
        card.grid(row=r, column=c, padx=6, pady=6, sticky="nsew")
        if use_custom:
            head = Frame(card, fg_color="transparent")
            head.pack(fill="x", padx=10, pady=(8, 2))
            Label(head, text=title, font=("Segoe UI", 12, "bold"), text_color="#1e293b").pack(side="left")
            chip = ctk.CTkLabel(head, text="-", width=74, height=22, corner_radius=999, fg_color="#334155", text_color="#ffffff", font=("Segoe UI", 11, "bold"))
            chip.pack(side="right")
            value_lbl = Label(card, textvariable=app.dashboard_pulse_vars[key], font=("Segoe UI", 16, "bold"), text_color="#0f172a")
            value_lbl.pack(anchor="w", padx=10, pady=(0, 8))
            app.dashboard_pulse_card_widgets[key] = {"card": card, "chip": chip, "value": value_lbl}
        else:
            ttk.Label(card, textvariable=app.dashboard_pulse_vars[key], font=("Segoe UI", 11, "bold")).pack(anchor="w", padx=10, pady=8)

    middle = Frame(win, fg_color="transparent") if use_custom else Frame(win)
    middle.pack(fill="both", expand=True, padx=14, pady=(0, 10))
    try:
        middle.grid_columnconfigure(0, weight=1)
        middle.grid_columnconfigure(1, weight=2)
        middle.grid_rowconfigure(0, weight=2)
        middle.grid_rowconfigure(1, weight=1)
        middle.grid_rowconfigure(2, weight=1)
    except Exception:
        pass

    alert_box = Frame(middle, fg_color="#ffffff" if use_custom else None, corner_radius=12 if use_custom else None, border_width=1 if use_custom else 0, border_color="#dce3ee" if use_custom else None) if use_custom else ttk.LabelFrame(middle, text="Alertas")
    alert_box.grid(row=0, column=0, sticky="nsew", padx=(0, 6), pady=(0, 6))
    Label(alert_box, text="Alertas", font=("Segoe UI", 13, "bold") if use_custom else ("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=(8, 4))
    Label(alert_box, textvariable=app.dashboard_pulse_vars["alertas"], justify="left", anchor="w", wraplength=360).pack(fill="both", expand=True, padx=10, pady=(0, 10))

    par_box = Frame(middle, fg_color="#ffffff" if use_custom else None, corner_radius=12 if use_custom else None, border_width=1 if use_custom else 0, border_color="#dce3ee" if use_custom else None) if use_custom else ttk.LabelFrame(middle, text="Top Causas de Paragem")
    par_box.grid(row=1, column=0, sticky="nsew", padx=(0, 6))
    Label(par_box, text="Top Causas de Paragem", font=("Segoe UI", 13, "bold") if use_custom else ("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=(8, 4))
    app.dashboard_pulse_paragens_host = None
    app.dashboard_pulse_paragens_tree = None
    if use_custom and ctk is not None:
        body = ctk.CTkScrollableFrame(par_box, fg_color="#ffffff", corner_radius=8)
        body.pack(fill="both", expand=True, padx=8, pady=(2, 8))
        app.dashboard_pulse_paragens_host = body
        app.dashboard_pulse_paragens_tree = None
    else:
        style_name = "PulseTop.Treeview"
        heading_style = "PulseTop.Treeview.Heading"
        try:
            st = ttk.Style()
            st.configure(style_name, rowheight=30, background="#f8fbff", fieldbackground="#f8fbff", borderwidth=0, relief="flat")
            st.map(style_name, background=[("selected", "#dbeafe")], foreground=[("selected", "#0f172a")])
            st.configure(heading_style, background="#0b1f66", foreground="#ffffff", font=("Segoe UI", 10, "bold"), relief="flat")
        except Exception:
            style_name = ""
        tree = ttk.Treeview(par_box, columns=("causa", "enc", "operador", "data", "ocor", "min"), show="headings", height=9, style=style_name)
        tree.heading("causa", text="Causa")
        tree.heading("enc", text="Encomenda")
        tree.heading("operador", text="Operador")
        tree.heading("data", text="Data")
        tree.heading("ocor", text="Ocorr.")
        tree.heading("min", text="Min")
        tree.column("causa", width=240, anchor="w")
        tree.column("enc", width=110, anchor="center")
        tree.column("operador", width=100, anchor="center")
        tree.column("data", width=140, anchor="center")
        tree.column("ocor", width=72, anchor="center")
        tree.column("min", width=72, anchor="center")
        tree.tag_configure("even", background="#f8fbff")
        tree.tag_configure("odd", background="#eef4ff")
        sy = ttk.Scrollbar(par_box, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sy.set)
        tree.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=(4, 8))
        sy.pack(side="right", fill="y", padx=(0, 8), pady=(4, 8))
        app.dashboard_pulse_paragens_tree = tree

    int_box = Frame(middle, fg_color="#ffffff" if use_custom else None, corner_radius=12 if use_custom else None, border_width=1 if use_custom else 0, border_color="#dce3ee" if use_custom else None) if use_custom else ttk.LabelFrame(middle, text="Pecas Interrompidas Ativas")
    int_box.grid(row=2, column=0, sticky="nsew", padx=(0, 6), pady=(6, 0))
    Label(int_box, text="Pecas Interrompidas Ativas", font=("Segoe UI", 13, "bold") if use_custom else ("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=(8, 4))
    app.dashboard_pulse_interrupted_host = None
    app.dashboard_pulse_interrupted_tree = None
    if use_custom and ctk is not None:
        body = ctk.CTkScrollableFrame(int_box, fg_color="#ffffff", corner_radius=8)
        body.pack(fill="both", expand=True, padx=8, pady=(2, 8))
        app.dashboard_pulse_interrupted_host = body
    else:
        cols = ("enc", "peca", "operador", "motivo")
        tree = ttk.Treeview(int_box, columns=cols, show="headings", height=8)
        tree.heading("enc", text="Enc")
        tree.heading("peca", text="Peca")
        tree.heading("operador", text="Operador")
        tree.heading("motivo", text="Motivo")
        tree.column("enc", width=110, anchor="center")
        tree.column("peca", width=170, anchor="w")
        tree.column("operador", width=110, anchor="center")
        tree.column("motivo", width=220, anchor="w", stretch=True)
        tree.tag_configure("even", background="#fef2f2")
        tree.tag_configure("odd", background="#fee2e2")
        sy = ttk.Scrollbar(int_box, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sy.set)
        tree.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=(4, 8))
        sy.pack(side="right", fill="y", padx=(0, 8), pady=(4, 8))
        app.dashboard_pulse_interrupted_tree = tree

    run_box = Frame(middle, fg_color="#ffffff" if use_custom else None, corner_radius=12 if use_custom else None, border_width=1 if use_custom else 0, border_color="#dce3ee" if use_custom else None) if use_custom else ttk.LabelFrame(middle, text="Pecas em Curso (Tempo)")
    run_box.grid(row=0, column=1, rowspan=3, sticky="nsew", padx=(6, 0), pady=(0, 0))
    Label(run_box, text="Pecas em Curso / Trabalhos (Tempo vs Planeado)", font=("Segoe UI", 13, "bold") if use_custom else ("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=(8, 4))
    app.dashboard_pulse_running_host = None
    app.dashboard_pulse_running_tree = None
    if use_custom and ctk is not None:
        body = ctk.CTkScrollableFrame(run_box, fg_color="#ffffff", corner_radius=8)
        body.pack(fill="both", expand=True, padx=8, pady=(2, 8))
        app.dashboard_pulse_running_host = body
    else:
        cols = ("enc", "peca", "op", "elapsed", "plan", "desvio")
        tree = ttk.Treeview(run_box, columns=cols, show="headings", height=7)
        for c, t, w, an in (
            ("enc", "Enc", 110, "center"),
            ("peca", "Peca", 210, "w"),
            ("op", "Operacao", 140, "center"),
            ("elapsed", "Tempo (m)", 90, "center"),
            ("plan", "Planeado (m)", 110, "center"),
            ("desvio", "Desvio (m)", 90, "center"),
        ):
            tree.heading(c, text=t)
            tree.column(c, width=w, anchor=an, stretch=(c == "peca"))
        tree.tag_configure("even", background="#f8fbff")
        tree.tag_configure("odd", background="#eef4ff")
        tree.tag_configure("warn", background="#fee2e2")
        tree.tag_configure("near", background="#fffbeb")
        sy = ttk.Scrollbar(run_box, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sy.set)
        tree.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=(4, 8))
        sy.pack(side="right", fill="y", padx=(0, 8), pady=(4, 8))
        app.dashboard_pulse_running_tree = tree

    app.dashboard_pulse_events_host = None
    app.dashboard_pulse_events_tree = None

    def _on_close():
        try:
            try:
                win.grab_release()
            except Exception:
                pass
            try:
                win.attributes("-topmost", False)
            except Exception:
                pass
            win.destroy()
        finally:
            app.dashboard_pulse_win = None
            app.dashboard_pulse_vars = {}
            app.dashboard_pulse_card_widgets = {}
            app.dashboard_pulse_priority_widgets = {}
            app.dashboard_pulse_last_alert = {"msg": "", "ts": 0.0}
            app.dashboard_pulse_alert_cfg = {"popup": True, "sound": True, "cooldown_sec": 90}
            app.dashboard_pulse_enc_var = None
            app.dashboard_pulse_origem_var = None
            app.dashboard_pulse_enc_menu = None
            try:
                if getattr(app, "dashboard_pulse_after_id", None):
                    app.root.after_cancel(app.dashboard_pulse_after_id)
            except Exception:
                pass
            app.dashboard_pulse_after_id = None
            app.dashboard_pulse_paragens_host = None
            app.dashboard_pulse_paragens_tree = None
            app.dashboard_pulse_interrupted_host = None
            app.dashboard_pulse_interrupted_tree = None
            app.dashboard_pulse_running_host = None
            app.dashboard_pulse_running_tree = None
            app.dashboard_pulse_events_host = None
            app.dashboard_pulse_events_tree = None

    try:
        win.protocol("WM_DELETE_WINDOW", _on_close)
    except Exception:
        pass

    refresh_production_pulse_window(app)


def _compute_dashboard_metrics(app):
    hoje = date.today()
    encomendas = list(app.data.get("encomendas", []))
    orcamentos = list(app.data.get("orcamentos", []))
    notas = list(app.data.get("notas_encomenda", []))
    produtos = list(app.data.get("produtos", []))
    plano_rows = list(app.data.get("plano", []))

    enc_total = len(encomendas)
    enc_concluidas = 0
    enc_em_producao = 0
    enc_atrasadas = 0
    total_valor_ativo = 0.0
    total_valor_concluido = 0.0
    for e in encomendas:
        estado_n = _norm_text(e.get("estado", ""))
        total_e = _calc_encomenda_total(app, e)
        if "concl" in estado_n:
            enc_concluidas += 1
            total_valor_concluido += total_e
        else:
            total_valor_ativo += total_e
        if ("produ" in estado_n) and ("concl" not in estado_n):
            enc_em_producao += 1
        if "concl" not in estado_n and "cancel" not in estado_n:
            dt = _safe_date(e.get("data_entrega"))
            if dt and dt < hoje:
                enc_atrasadas += 1

    enc_ativas = max(0, enc_total - enc_concluidas)
    taxa_conclusao = (enc_concluidas / enc_total * 100.0) if enc_total else 0.0

    orc_abertos = 0
    for o in orcamentos:
        estado = _norm_text(o.get("estado", ""))
        if ("edi" in estado) or ("enviado" in estado) or ("ativo" in estado):
            orc_abertos += 1

    ne_pend = 0
    for n in notas:
        estado = _norm_text(n.get("estado", ""))
        if not any(x in estado for x in ("entregue", "cancelada", "cancelado", "concl")):
            ne_pend += 1

    stock_alerta = 0
    for p in produtos:
        alerta = _parse_float(p.get("alerta", p.get("aviso_stock", 0)), 0)
        qtd = _parse_float(p.get("qty", p.get("quantidade", 0)), 0)
        if alerta > 0 and qtd <= alerta:
            stock_alerta += 1

    hoje_week = hoje.isocalendar()
    plano_semana = 0
    for row in plano_rows:
        d = _safe_date(row.get("data") or row.get("dia"))
        if not d:
            continue
        wk = d.isocalendar()
        if (wk[0], wk[1]) == (hoje_week[0], hoje_week[1]):
            plano_semana += 1

    alertas = [
        f"- Encomendas atrasadas: {enc_atrasadas}",
        f"- Notas de encomenda pendentes: {ne_pend}",
        f"- Produtos com stock abaixo do alerta: {stock_alerta}",
        f"- Orcamentos ativos: {orc_abertos}",
    ]

    urgentes = []
    for e in encomendas:
        estado = _norm_text(e.get("estado", ""))
        if "concl" in estado or "cancel" in estado:
            continue
        numero = str(e.get("numero", "") or "").strip() or "(sem numero)"
        cliente = str(e.get("cliente", "") or "").strip() or "sem cliente"
        dt = _safe_date(e.get("data_entrega"))
        if not dt:
            continue
        delta = (dt - hoje).days
        if delta < 0:
            urgentes.append((120 + abs(delta), f"Encomenda {numero} atrasada ({cliente})"))
        elif delta <= 1:
            when_txt = "hoje" if delta == 0 else "amanha"
            urgentes.append((95 - delta, f"Encomenda {numero} entrega {when_txt}"))

    for n in notas:
        estado = _norm_text(n.get("estado", ""))
        if any(x in estado for x in ("entregue", "cancelada", "cancelado", "concl")):
            continue
        num = str(n.get("numero", "") or "").strip() or "(sem numero)"
        dt = _safe_date(n.get("data_entrega") or n.get("data_prevista"))
        if not dt:
            continue
        delta = (dt - hoje).days
        if delta < 0:
            urgentes.append((105 + abs(delta), f"NE {num} atrasada"))
        elif delta <= 1:
            when_txt = "hoje" if delta == 0 else "amanha"
            urgentes.append((82 - delta, f"NE {num} prevista para {when_txt}"))

    urgentes.sort(key=lambda x: x[0], reverse=True)
    urgentes_txt = []
    seen = set()
    for _, txt in urgentes:
        key = txt.lower().strip()
        if not key or key in seen:
            continue
        seen.add(key)
        urgentes_txt.append(f"- {txt}")
        if len(urgentes_txt) >= 8:
            break
    if not urgentes_txt:
        urgentes_txt = ["- Sem itens urgentes."]

    monthly = {}
    for e in encomendas:
        d = _safe_date(e.get("data_criacao") or e.get("data_entrega"))
        if not d:
            continue
        key = d.strftime("%Y-%m")
        rec = monthly.setdefault(key, {"total": 0, "concluidas": 0})
        rec["total"] += 1
        if "concl" in _norm_text(e.get("estado", "")):
            rec["concluidas"] += 1
    keys = sorted(monthly.keys())[-6:]
    chart_data = [(k, monthly[k]["total"], monthly[k]["concluidas"]) for k in keys]

    return {
        "enc_total": enc_total,
        "enc_ativas": enc_ativas,
        "enc_em_producao": enc_em_producao,
        "enc_concluidas": enc_concluidas,
        "enc_atrasadas": enc_atrasadas,
        "orc_abertos": orc_abertos,
        "ne_pend": ne_pend,
        "stock_alerta": stock_alerta,
        "plano_semana": plano_semana,
        "taxa_conclusao": taxa_conclusao,
        "valor_ativo": total_valor_ativo,
        "valor_concluido": total_valor_concluido,
        "alertas": "\n".join(alertas),
        "urgentes": "\n".join(urgentes_txt),
        "chart_data": chart_data,
    }


def _draw_dashboard_chart(canvas, chart_data):
    try:
        canvas.delete("all")
    except Exception:
        return
    w = max(620, int(canvas.winfo_width() or 980))
    h = max(220, int(canvas.winfo_height() or 260))
    try:
        canvas.configure(width=w, height=h)
    except Exception:
        pass
    left, right = 48, w - 20
    top, bottom = 20, h - 34
    canvas.create_line(left, top, left, bottom, fill="#b8c3d6")
    canvas.create_line(left, bottom, right, bottom, fill="#b8c3d6")
    if not chart_data:
        canvas.create_text(
            14,
            20,
            anchor="nw",
            text="Sem dados de tendencia para os ultimos meses.",
            fill="#64748b",
            font=("Segoe UI", 10, "italic"),
        )
        return
    max_v = 1
    for _, total, concl in chart_data:
        max_v = max(max_v, int(total), int(concl))
    group_w = max(72, int((right - left) / max(1, len(chart_data))))
    bar_w = max(12, int(group_w * 0.25))
    for i, (month, total, concl) in enumerate(chart_data):
        x0 = left + i * group_w + 12
        ht = int((bottom - top - 10) * (float(total) / max_v))
        hc = int((bottom - top - 10) * (float(concl) / max_v))
        yt = bottom - ht
        yc = bottom - hc
        canvas.create_rectangle(x0, yt, x0 + bar_w, bottom, fill="#60a5fa", outline="")
        canvas.create_rectangle(x0 + bar_w + 5, yc, x0 + (bar_w * 2) + 5, bottom, fill="#22c55e", outline="")
        canvas.create_text(x0 + bar_w, bottom + 12, text=month[5:], fill="#334155", font=("Segoe UI", 9, "bold"))
        canvas.create_text(x0 + bar_w, min(yt - 10, bottom - 8), text=str(int(total)), fill="#1d4ed8", font=("Segoe UI", 8))
        canvas.create_text(x0 + bar_w + 5, min(yc - 10, bottom - 8), text=str(int(concl)), fill="#15803d", font=("Segoe UI", 8))
    canvas.create_text(right - 210, top + 4, text="Total", anchor="w", fill="#1d4ed8", font=("Segoe UI", 9, "bold"))
    canvas.create_text(right - 140, top + 4, text="Concluidas", anchor="w", fill="#15803d", font=("Segoe UI", 9, "bold"))


def refresh_dashboard_window(app):
    win = getattr(app, "dashboard_win", None)
    if win is None:
        return
    try:
        if not win.winfo_exists():
            app.dashboard_win = None
            return
    except Exception:
        app.dashboard_win = None
        return
    data = _compute_dashboard_metrics(app)
    vars_map = getattr(app, "dashboard_vars", {})
    set_map = {
        "enc_total": str(data["enc_total"]),
        "enc_ativas": str(data["enc_ativas"]),
        "enc_em_producao": str(data["enc_em_producao"]),
        "enc_concluidas": str(data["enc_concluidas"]),
        "enc_atrasadas": str(data["enc_atrasadas"]),
        "orc_abertos": str(data["orc_abertos"]),
        "ne_pend": str(data["ne_pend"]),
        "stock_alerta": str(data["stock_alerta"]),
        "plano_semana": str(data["plano_semana"]),
        "taxa_conclusao": f"{data['taxa_conclusao']:.1f}%",
        "valor_ativo": f"{data['valor_ativo']:.2f} EUR",
        "valor_concluido": f"{data['valor_concluido']:.2f} EUR",
        "alertas": data["alertas"],
        "urgentes": data["urgentes"],
    }
    for k, val in set_map.items():
        v = vars_map.get(k)
        if v is not None:
            try:
                v.set(val)
            except Exception:
                pass
    cv = getattr(app, "dashboard_chart_canvas", None)
    if cv is not None:
        _draw_dashboard_chart(cv, data.get("chart_data", []))
    try:
        ts_var = vars_map.get("updated_at")
        if ts_var is not None:
            ts_var.set(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    except Exception:
        pass


def open_dashboard_window(app):
    use_custom = getattr(app, "menu_use_custom", False) and ctk is not None
    old = getattr(app, "dashboard_win", None)
    try:
        if old is not None and old.winfo_exists():
            old.lift()
            old.focus_force()
            refresh_dashboard_window(app)
            return
    except Exception:
        pass

    Win = ctk.CTkToplevel if use_custom else Toplevel
    Frame = ctk.CTkFrame if use_custom else ttk.Frame
    Label = ctk.CTkLabel if use_custom else ttk.Label
    Btn = ctk.CTkButton if use_custom else ttk.Button

    win = Win(app.root)
    app.dashboard_win = win
    try:
        win.title("Dashboard ERP - luGEST")
        win.transient(app.root)
        _center_window_on_screen(win, 1380, 880, min_w=1180, min_h=760)
        if use_custom:
            win.configure(fg_color="#eef2f7")
        _bring_window_front(win, app.root, modal=True, keep_topmost=False)
    except Exception:
        pass
    try:
        win.grid_columnconfigure(0, weight=1)
        win.grid_rowconfigure(3, weight=1)
        win.grid_rowconfigure(4, weight=1)
    except Exception:
        pass

    app.dashboard_vars = {k: StringVar(value="-") for k in (
        "enc_total", "enc_ativas", "enc_em_producao", "enc_concluidas",
        "enc_atrasadas", "orc_abertos", "ne_pend", "stock_alerta",
        "plano_semana", "taxa_conclusao", "valor_ativo", "valor_concluido",
        "alertas", "urgentes", "updated_at",
    )}

    top = Frame(win, fg_color="transparent") if use_custom else Frame(win)
    top.grid(row=0, column=0, sticky="ew", padx=14, pady=(12, 8))
    try:
        top.grid_columnconfigure(0, weight=1)
    except Exception:
        pass
    Label(top, text="Dashboard Operacional ERP", font=("Segoe UI", 24, "bold") if use_custom else ("Segoe UI", 16, "bold"), text_color="#0f172a" if use_custom else None).grid(row=0, column=0, sticky="w")
    Label(top, textvariable=app.dashboard_vars["updated_at"], text_color="#64748b" if use_custom else None, font=("Segoe UI", 12, "bold") if use_custom else None).grid(row=0, column=1, sticky="e", padx=(6, 0))

    actions = Frame(win, fg_color="#ffffff" if use_custom else "transparent", corner_radius=10 if use_custom else None, border_width=1 if use_custom else 0, border_color="#d7dee9" if use_custom else None) if use_custom else Frame(win)
    actions.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 8))
    Btn(actions, text="Atualizar", command=lambda: refresh_dashboard_window(app), width=130 if use_custom else None).pack(side="left", padx=6, pady=8)
    Btn(actions, text="Financeiro Stock", command=lambda: open_dashboard_finance_window(app), width=160 if use_custom else None).pack(side="left", padx=6, pady=8)
    Btn(actions, text="Production Pulse", command=lambda: open_production_pulse_window(app), width=160 if use_custom else None).pack(side="left", padx=6, pady=8)
    Btn(actions, text="Encomendas", command=lambda: app.navigate_to_tab(app.tab_encomendas), width=130 if use_custom else None).pack(side="left", padx=6, pady=8)
    Btn(actions, text="Plano", command=lambda: app.navigate_to_tab(app.tab_plano), width=130 if use_custom else None).pack(side="left", padx=6, pady=8)
    Btn(actions, text="Notas Encomenda", command=lambda: app.navigate_to_tab(app.tab_ne), width=150 if use_custom else None).pack(side="left", padx=6, pady=8)
    Btn(actions, text="Fechar", command=win.destroy, width=120 if use_custom else None).pack(side="right", padx=6, pady=8)

    cards_wrap = Frame(win, fg_color="transparent") if use_custom else Frame(win)
    cards_wrap.grid(row=2, column=0, sticky="ew", padx=14, pady=(0, 8))
    for c in range(4):
        try:
            cards_wrap.grid_columnconfigure(c, weight=1)
        except Exception:
            pass

    card_specs = [
        ("Encomendas Totais", "enc_total"),
        ("Encomendas Ativas", "enc_ativas"),
        ("Em Producao", "enc_em_producao"),
        ("Concluidas", "enc_concluidas"),
        ("Atrasadas", "enc_atrasadas"),
        ("Orcamentos Ativos", "orc_abertos"),
        ("NE Pendentes", "ne_pend"),
        ("Alerta Stock", "stock_alerta"),
        ("Plano Semana", "plano_semana"),
        ("Taxa Conclusao", "taxa_conclusao"),
        ("Valor Ativo", "valor_ativo"),
        ("Valor Concluido", "valor_concluido"),
    ]
    for i, (title, key) in enumerate(card_specs):
        r, c = divmod(i, 4)
        card = Frame(cards_wrap, fg_color="#ffffff" if use_custom else None, corner_radius=12 if use_custom else None, border_width=1 if use_custom else 0, border_color="#dce3ee" if use_custom else None) if use_custom else ttk.LabelFrame(cards_wrap, text=title)
        card.grid(row=r, column=c, padx=6, pady=6, sticky="nsew")
        if use_custom:
            Label(card, text=title, font=("Segoe UI", 12, "bold"), text_color="#1e293b").pack(anchor="w", padx=10, pady=(8, 2))
            Label(card, textvariable=app.dashboard_vars[key], font=("Segoe UI", 19, "bold"), text_color="#0f172a").pack(anchor="w", padx=10, pady=(0, 8))
        else:
            ttk.Label(card, textvariable=app.dashboard_vars[key], font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=10, pady=6)

    mid = Frame(win, fg_color="transparent") if use_custom else Frame(win)
    mid.grid(row=3, column=0, sticky="nsew", padx=14, pady=(0, 8))
    try:
        mid.grid_columnconfigure(0, weight=1)
        mid.grid_columnconfigure(1, weight=1)
        mid.grid_rowconfigure(0, weight=1)
    except Exception:
        pass

    alerts_box = Frame(mid, fg_color="#ffffff" if use_custom else None, corner_radius=12 if use_custom else None, border_width=1 if use_custom else 0, border_color="#dce3ee" if use_custom else None) if use_custom else ttk.LabelFrame(mid, text="Alertas")
    alerts_box.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
    Label(alerts_box, text="Alertas", font=("Segoe UI", 13, "bold") if use_custom else ("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=(8, 4))
    Label(alerts_box, textvariable=app.dashboard_vars["alertas"], justify="left", anchor="w", wraplength=560).pack(fill="both", expand=True, padx=10, pady=(0, 10))

    urg_box = Frame(mid, fg_color="#ffffff" if use_custom else None, corner_radius=12 if use_custom else None, border_width=1 if use_custom else 0, border_color="#dce3ee" if use_custom else None) if use_custom else ttk.LabelFrame(mid, text="Top Urgentes")
    urg_box.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
    Label(urg_box, text="Top Urgentes", font=("Segoe UI", 13, "bold") if use_custom else ("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=(8, 4))
    Label(urg_box, textvariable=app.dashboard_vars["urgentes"], justify="left", anchor="w", wraplength=560).pack(fill="both", expand=True, padx=10, pady=(0, 10))

    chart_wrap = Frame(win, fg_color="#ffffff" if use_custom else None, corner_radius=12 if use_custom else None, border_width=1 if use_custom else 0, border_color="#dce3ee" if use_custom else None) if use_custom else ttk.LabelFrame(win, text="Tendencia (6 meses)")
    chart_wrap.grid(row=4, column=0, sticky="nsew", padx=14, pady=(0, 12))
    Label(chart_wrap, text="Tendencia de Encomendas (Total vs Concluidas)", font=("Segoe UI", 13, "bold") if use_custom else ("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=(8, 4))
    app.dashboard_chart_canvas = Canvas(chart_wrap, height=250, bg="#ffffff", highlightthickness=0 if use_custom else 1)
    app.dashboard_chart_canvas.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    def _on_close():
        try:
            try:
                win.grab_release()
            except Exception:
                pass
            win.destroy()
        finally:
            app.dashboard_win = None
            app.dashboard_chart_canvas = None

    try:
        win.protocol("WM_DELETE_WINDOW", _on_close)
    except Exception:
        pass

    refresh_dashboard_window(app)


def build_menu_dashboard(app, *, custom_tk_available, load_logo_fn, get_orc_logo_path_fn):
    app.menu_use_custom = custom_tk_available and ctk is not None and os.environ.get("USE_CUSTOM_MENU", "1") != "0"
    container = ttk.Frame(app.tab_menu, style="Menu.TFrame")
    container.pack(fill="both", expand=True, padx=20, pady=16)

    top = ttk.Frame(container, style="Menu.TFrame")
    top.pack(fill="x", pady=(0, 16))

    app.menu_logo_img = None
    try:
        logo_path = get_orc_logo_path_fn()
        if logo_path:
            try:
                from PIL import Image, ImageTk
                img = Image.open(logo_path)
                max_w = 340
                if img.width > max_w:
                    scale = max_w / float(img.width)
                    img = img.resize((int(img.width * scale), int(img.height * scale)))
                app.menu_logo_img = ImageTk.PhotoImage(img)
            except Exception:
                app.menu_logo_img = load_logo_fn(340)
    except Exception:
        app.menu_logo_img = load_logo_fn(340)

    if app.menu_logo_img:
        ttk.Label(top, image=app.menu_logo_img, style="Menu.TLabel").pack(pady=(0, 4))
    ttk.Label(top, text="luGEST", style="Menu.Title.TLabel").pack(pady=(6, 0))
    ttk.Label(top, text="Sistema de Gestao de Producao", style="Menu.Sub.TLabel").pack(pady=(2, 6))

    grid = ttk.Frame(container, style="Menu.TFrame")
    grid.pack(anchor="n")

    specs = [
        ("Clientes", app.tab_clientes),
        ("Materia-Prima", app.tab_materia),
        ("Encomendas", app.tab_encomendas),
        ("Plano de Producao", app.tab_plano),
        ("Qualidade", app.tab_qualidade),
        ("Orcamentacao", app.tab_orc),
        ("Ordens de Fabrico", app.tab_ordens),
        ("Exportacoes", app.tab_export),
        ("Expedicao", app.tab_expedicao),
        ("Produtos", app.tab_produtos),
        ("Notas Encomenda", app.tab_ne),
        ("Operador", app.tab_operador),
        ("Dashboard", "dashboard"),
        ("Encerrar", None),
    ]
    app.menu_nav_cols = 3
    app.menu_nav = []
    for idx, (txt, tab) in enumerate(specs):
        r = idx // app.menu_nav_cols
        c = idx % app.menu_nav_cols
        if tab is None:
            cmd = app.root.destroy
        elif tab == "dashboard":
            cmd = lambda: open_dashboard_window(app)
        else:
            cmd = (lambda t=tab: app.navigate_to_tab(t))
        btn = app.create_menu_button(
            grid,
            txt,
            command=cmd,
            danger=(tab is None),
            compact=False,
        )
        btn.grid(row=r, column=c, padx=10, pady=8, sticky="")
        app.menu_nav.append((btn, tab))
    for i in range(app.menu_nav_cols):
        grid.columnconfigure(i, minsize=210)

    app.menu_stats_vars = {}
    app.menu_alerts_var = StringVar(value="")
    app.menu_urgentes_var = StringVar(value="")
    app.menu_chart_canvas = None


def create_menu_button(app, parent, text, command, danger=False, compact=True):
    if getattr(app, "menu_use_custom", False) and ctk is not None:
        txt_norm = str(text or "").strip().lower()
        is_encerrar = "encerrar" in txt_norm
        theme_primary = getattr(app, "CTK_PRIMARY_RED", None) or globals().get("CTK_PRIMARY_RED") or "#ba2d3d"
        theme_hover = getattr(app, "CTK_PRIMARY_RED_HOVER", None) or globals().get("CTK_PRIMARY_RED_HOVER") or "#a32035"
        if is_encerrar:
            fg = "#f59e0b"
            hover = "#d97706"
        else:
            fg = theme_primary
            hover = theme_hover
        return ctk.CTkButton(
            parent,
            text=text,
            command=command,
            width=(220 if compact else 230),
            corner_radius=(16 if compact else 22),
            height=(42 if compact else 56),
            font=("Segoe UI", (13 if compact else 15), "bold"),
            fg_color=fg,
            hover_color=hover,
            text_color="#ffffff",
            bg_color="#ffffff",
        )
    btn = Button(
        parent,
        text=text,
        width=(18 if compact else 24),
        height=(1 if compact else 2),
        command=command,
        **app.btn_cfg,
    )
    if danger:
        btn.configure(bg="#f3dede", activebackground="#ecc5c5")
    return btn


def refresh_menu_dashboard(app):
    if not hasattr(app, "menu_nav"):
        return
    cols = max(1, int(getattr(app, "menu_nav_cols", 2) or 2))
    visibles = []
    for btn, tab in app.menu_nav:
        if tab is None or tab == "dashboard":
            visibles.append((btn, tab))
            continue
        state = app.nb.tab(tab, "state")
        if state == "hidden":
            btn.grid_remove()
        else:
            visibles.append((btn, tab))
    pos = 0
    for btn, tab in visibles:
        r = pos // cols
        c = pos % cols
        btn.grid(row=r, column=c, padx=10, pady=8, sticky="")
        pos += 1
    refresh_dashboard_window(app)


def refresh_menu_dashboard_stats(app, *, norm_text_fn, parse_float_fn):
    refresh_dashboard_window(app)
    refresh_dashboard_finance_window(app)


def refresh_menu_dashboard_chart(app, *, norm_text_fn, parse_float_fn):
    refresh_dashboard_window(app)
    refresh_dashboard_finance_window(app)


def navigate_to_tab(app, tab):
    try:
        app.nb.select(tab)
    except Exception:
        return
    app.update_menu_back_button()


def go_to_main_menu(app):
    try:
        if app.nb.tab(app.tab_menu, "state") == "hidden":
            return
        app.nb.select(app.tab_menu)
    except Exception:
        return
    app.update_menu_back_button()


def on_notebook_tab_changed(app, _event=None):
    try:
        selected = app.nb.select()
        def _deferred_tab_load(sel=selected):
            try:
                if hasattr(app, "ensure_tab_built_by_widget"):
                    app.ensure_tab_built_by_widget(sel)
                if hasattr(app, "refresh_selected_tab_if_dirty"):
                    app.refresh_selected_tab_if_dirty(sel)
                if hasattr(app, "tab_menu") and str(sel) == str(app.tab_menu):
                    app.refresh_menu_dashboard()
            except Exception:
                pass
        try:
            app.root.after(1, _deferred_tab_load)
        except Exception:
            _deferred_tab_load()
    except Exception:
        pass
    app.update_menu_back_button()


def update_menu_back_button(app):
    if app.user.get("role") == "Operador":
        return
    try:
        if app.nb.tab(app.tab_menu, "state") == "hidden":
            if app.menu_back_bar:
                app.menu_back_bar.pack_forget()
            return
    except Exception:
        return
    selected = app.nb.select()
    show = selected != str(app.tab_menu)
    if show:
        if not app.menu_back_bar:
            app.menu_back_bar = ttk.Frame(app.root, style="Menu.TFrame")
            btn = app.create_menu_button(app.menu_back_bar, "< Menu Principal", app.go_to_main_menu, danger=False, compact=True)
            btn.pack(side="left", padx=6, pady=4)
        app.menu_back_bar.pack(fill="x", padx=10, pady=(6, 0), before=app.nb)
    else:
        if app.menu_back_bar:
            app.menu_back_bar.pack_forget()


def apply_menu_only_mode(app):
    if not getattr(app, "menu_only_mode", False):
        return
    try:
        app.nb.configure(style="MenuOnly.TNotebook")
        style = ttk.Style()
        style.layout("MenuOnly.TNotebook.Tab", [])
        style.configure("MenuOnly.TNotebook", tabmargins=(0, 0, 0, 0))
        style.configure("MenuOnly.TNotebook.Tab", padding=(0, 0), font=("Segoe UI", 1))
    except Exception:
        pass


def show_operador_menu(app):
    if app.op_menu:
        app.op_menu.pack_forget()
    app.op_menu = ttk.Frame(app.root, style="Menu.TFrame")
    app.op_menu.pack(fill="x", padx=10, pady=(10, 0), before=app.nb)
    app.create_menu_button(app.op_menu, "Operador", lambda: app.navigate_to_tab(app.tab_operador), compact=True).pack(side="left", padx=6)
    app.create_menu_button(app.op_menu, "Plano de Producao", lambda: app.navigate_to_tab(app.tab_plano), compact=True).pack(side="left", padx=6)

