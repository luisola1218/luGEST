"""Business routing decisions independently of the desktop runtime."""
from copy import deepcopy
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from lugest_modules.quotes.domain.routing import LineRouting, RoutingRules


def main():
    routing = LineRouting(RoutingRules(
        "piece", "product", "service", lambda value: value or "piece",
        lambda value: value, str.lower, lambda value, default=0: float(value or default),
        lambda values: list(dict.fromkeys(values)), lambda values: "; ".join(values),
    ))
    cases = [
        ({"tipo_item": "product"}, "montagem"),
        ({"tipo_item": "service"}, "montagem"),
        ({}, "conjunto"),
        ({"desenho": "peca.dxf", "operacao": "Corte Laser"}, "laser"),
        ({"material": "Tubo", "espessura": "3", "operacao": "Serralharia"}, "serralharia"),
        ({"material": "Aco", "espessura": "3", "operacao": "Corte Laser"}, "laser"),
    ]
    original = deepcopy(cases)
    for line, route in cases:
        assert routing.production_route(line) == route
    assert routing.raw_material({"stock_material_id": "MAT1"})
    assert routing.raw_material({"ref_externa": "MAT123"})
    assert not routing.raw_material({"ref_externa": "MAT123", "desenho": "part.dxf"})
    assert not routing.raw_material({"ref_externa": "MAT123", "tempo_peca_min": 2})
    assert not routing.raw_material({"tipo_item": "product", "stock_material_id": "MAT1"})
    assert routing.operations({"operacao": "Corte Laser", "tempos_operacao": {"Corte Laser": 2, "Quinagem": 3}}) == ["Corte Laser", "Quinagem"]
    assert cases == original and "main" not in sys.modules
    print("quote-routing-ok laser=yes fabrication=yes assembly=yes raw-material=yes detached=yes")


if __name__ == "__main__":
    main()
