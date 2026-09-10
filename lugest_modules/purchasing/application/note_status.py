"""Normalize delivery progress without discarding explicit purchasing states."""
def last_delivery_date(ne):
    dates = []

    def _push(raw):
        txt = str(raw or "").strip()
        if txt:
            dates.append(txt[:10])

    if not isinstance(ne, dict):
        return ""
    _push(ne.get("data_ultima_entrega"))
    _push(ne.get("data_entregue"))
    for ent in ne.get("entregas", []):
        if not isinstance(ent, dict):
            continue
        _push(ent.get("data_entrega"))
        _push(ent.get("data_documento"))
    for line in ne.get("linhas", []):
        if not isinstance(line, dict):
            continue
        _push(line.get("data_entrega_real"))
        _push(line.get("data_doc_entrega"))
        for ent_l in line.get("entregas_linha", []):
            if not isinstance(ent_l, dict):
                continue
            _push(ent_l.get("data_entrega"))
            _push(ent_l.get("data_documento"))
    return max(dates) if dates else ""


def normalize_status(ne, parse_float, norm_text):
    if not isinstance(ne, dict):
        return False
    changed = False
    estado_atual = str(ne.get("estado", "") or "").strip()
    estado_norm = norm_text(estado_atual)

    if "cancel" in estado_norm:
        return False

    if "convert" in estado_norm or bool(ne.get("oculta")):
        if estado_atual != "Convertida":
            ne["estado"] = "Convertida"
            changed = True
        return changed

    if bool(ne.get("_draft")):
        desired = estado_atual or "Em edicao"
        if desired != estado_atual:
            ne["estado"] = desired
            changed = True
        return changed

    linhas = [line for line in ne.get("linhas", []) if isinstance(line, dict)]
    if not linhas:
        desired = estado_atual or "Em edicao"
        if desired != estado_atual:
            ne["estado"] = desired
            changed = True
        return changed

    any_started = False
    all_done = True
    meaningful_lines = 0

    for line in linhas:
        qtd_total = max(0.0, parse_float(line.get("qtd", 0), 0))
        qtd_ent_raw = parse_float(
            line.get("qtd_entregue", qtd_total if line.get("entregue") else 0),
            0,
        )
        qtd_ent = max(0.0, qtd_ent_raw)
        flag_done = bool(line.get("entregue"))
        flag_stock = bool(line.get("_stock_in")) or bool(line.get("stock_in"))

        if qtd_total > 0:
            meaningful_lines += 1
            qtd_ent = min(qtd_total, qtd_ent)
            if flag_done and qtd_ent < (qtd_total - 1e-9):
                qtd_ent = qtd_total
            if abs(qtd_ent - qtd_ent_raw) > 1e-9:
                line["qtd_entregue"] = qtd_ent
                changed = True

            line_done = flag_done or qtd_ent >= (qtd_total - 1e-9)
            if bool(line.get("entregue")) != line_done:
                line["entregue"] = line_done
                changed = True

            any_started = any_started or flag_done or flag_stock or qtd_ent > 0
            all_done = all_done and line_done
        else:
            any_started = any_started or flag_done or flag_stock or qtd_ent > 0

    if meaningful_lines and all_done:
        desired = "Entregue"
    elif any_started:
        desired = "Parcialmente Entregue"
    elif estado_norm in {"enviada", "cotacao aprovada"}:
        desired = estado_atual
    elif "edi" in estado_norm and "apro" not in estado_norm:
        desired = estado_atual or "Em edicao"
    else:
        desired = "Aprovada"

    if desired != estado_atual:
        ne["estado"] = desired
        changed = True

    if desired == "Entregue":
        last_dt = last_delivery_date(ne)
        if last_dt and not str(ne.get("data_ultima_entrega", "") or "").strip():
            ne["data_ultima_entrega"] = last_dt
            changed = True
        if last_dt and not str(ne.get("data_entregue", "") or "").strip():
            ne["data_entregue"] = last_dt
            changed = True

    return changed
