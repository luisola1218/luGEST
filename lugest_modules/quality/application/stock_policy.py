"""Quality stock policy operating exclusively on the supplied item."""

def status_code(value):
    raw = str(value or "").strip().casefold()
    if "devol" in raw:
        return "DEVOLVER_FORNECEDOR"
    if "averig" in raw or "analise" in raw or "análise" in raw:
        return "EM_AVERIGUACAO"
    if "rejeit" in raw:
        return "REJEITADO"
    if "aprov" in raw:
        return "APROVADO"
    return "EM_INSPECAO"


def quarantine(item, *, kind, parse_float, now_iso, max_qty=None):
    if not isinstance(item, dict):
        return False
    status = status_code(item.get("quality_status", item.get("inspection_status", "")))
    if status == "APROVADO":
        return False
    qty_key = "qty" if str(kind or "").casefold().startswith("prod") else "quantidade"
    current_qty = parse_float(item.get(qty_key, 0), 0)
    pending = parse_float(item.get("quality_pending_qty", 0), 0)
    if current_qty <= 0:
        return False
    quarantine_qty = current_qty
    if max_qty is not None:
        quarantine_qty = min(current_qty, max(0.0, parse_float(max_qty, 0) - pending))
    if quarantine_qty <= 0:
        return False
    item["quality_pending_qty"] = pending + quarantine_qty
    item["quality_received_qty"] = max(parse_float(item.get("quality_received_qty", 0), 0), pending + quarantine_qty)
    item[qty_key] = max(0.0, current_qty - quarantine_qty)
    item["quality_blocked"] = True
    item["logistic_status"] = str(item.get("logistic_status", "") or "RECEBIDO").strip()
    item["atualizado_em"] = now_iso()
    return True

