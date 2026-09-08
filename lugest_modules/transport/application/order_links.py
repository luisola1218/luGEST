"""Derive order transport links from the active routes without mutating inputs."""
from copy import deepcopy


def synchronize_orders(trips, orders, norm_text):
    assigned = {}
    for route in trips:
        if not isinstance(route, dict):
            continue
        number = str(route.get("numero", "") or "").strip()
        state = str(route.get("estado", "") or "Planeado").strip() or "Planeado"
        if not number or "anulad" in norm_text(state):
            continue
        for stop in route.get("paragens", []) or []:
            if not isinstance(stop, dict):
                continue
            order_number = str(stop.get("encomenda_numero", stop.get("encomenda", "")) or "").strip()
            if order_number:
                assigned[order_number] = (number, str(stop.get("estado", "") or "").strip() or state)
    result = deepcopy(orders)
    for order in result:
        if not isinstance(order, dict):
            continue
        number = str(order.get("numero", "") or "").strip()
        if number:
            order["transporte_numero"], order["estado_transporte"] = assigned.get(number, ("", ""))
    return result
