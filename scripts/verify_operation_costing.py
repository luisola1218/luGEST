"""Cost calculations independent of Qt, MySQL and configuration files."""
from copy import deepcopy
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from lugest_core.operation_costing import OperationCostingEngine


def parse_number(value, default=0):
    try:
        return float(str(value).replace(",", "."))
    except (ValueError, TypeError):
        return float(default)


def main():
    engine = OperationCostingEngine(
        available_operations=["Cut", "Paint"],
        normalize_operation=lambda value: str(value or "").strip(),
        parse_number=parse_number,
        parse_operations=lambda value: [v.strip() for v in str(value).split(";") if v.strip()],
    )
    settings = engine.merge_settings({"active_profile": "Test", "profiles": {"Test": {
        "Cut": {"pricing_mode": "per_piece", "setup_min": 10, "unit_time_min": 2, "hour_rate_eur": 60},
        "Paint": {"pricing_mode": "per_area_m2", "fixed_unit_eur": 12},
    }}})
    payload = {"operacao": "Cut; Paint", "qtd": 5, "area_m2": 2}
    before = deepcopy((payload, settings))
    result = engine.estimate(payload, settings)
    assert result["summary"]["complete"]
    assert result["summary"]["tempo_total_min"] == 20
    assert result["summary"]["custo_total_eur"] == 140
    assert (payload, settings) == before
    settings["profiles"]["Test"]["Cut"]["requires_driver_input"] = True
    result = engine.estimate({"operacao": "Cut", "qtd": 5}, settings)
    assert not result["summary"]["complete"]
    assert result["operations"][0]["missing_driver_input"]
    confirmed = {"operacao": "Cut", "qtd": 5, "operacoes_detalhe": [{"nome": "Cut", "driver_units_confirmed": True}]}
    assert engine.estimate(confirmed, settings)["summary"]["complete"]
    manual = {"operacao": "Manual", "qtd": 2, "operacoes_detalhe": [
        {"nome": "Manual", "tempo_unit_min": 3, "custo_unit_eur": 5}]}
    result = engine.estimate(manual, settings)
    assert result["summary"]["custo_total_eur"] == 10
    assert result["summary"]["tempo_total_min"] == 6
    assert not engine.estimate({"operacao": ""}, settings)["summary"]["complete"]
    print("operation-costing-ok setup=yes area=yes manual=yes confirmation=yes inputs-preserved=yes")


if __name__ == "__main__":
    main()
