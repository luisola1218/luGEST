"""Regression for quote conversion totals, routing and detached output."""
from copy import deepcopy
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from lugest_modules.quotes.application.order_lines import OrderLinePorts, build_order_lines


def main():
    allocated = []

    def next_reference(used):
        ref = f"NEW{len(allocated) + 1}"
        assert ref not in used
        allocated.append(ref)
        return ref

    ports = OrderLinePorts(
        parse_float=lambda value, default=0: float(value or default),
        normalize_orc_line_type=lambda value: value or "piece",
        normalize_operacao_nome=lambda value: value,
        production_route=lambda line: line.get("route", "laser"),
        operations=lambda line: line.get("ops", ["Corte Laser"]),
        operations_text=lambda line: "; ".join(line.get("ops", ["Corte Laser"])),
        default_resource=lambda operation, preferred: f"{operation}:{preferred}",
        next_reference=next_reference,
        next_piece_order=lambda: "OPP-FALLBACK",
        build_operacoes_fluxo=lambda value: [{"nome": value}],
        now_iso=lambda: "2026-09-08T11:00:00",
    )
    lines = [
        {"material": "Aco", "espessura": "3", "qtd": 2, "tempo_peca_min": 3,
         "ref_interna": "R1", "ops": ["Corte Laser", "Quinagem"],
         "tempos_operacao": {"Corte Laser": 1, "Quinagem": 2}},
        {"material": "Aco", "espessura": "3", "qtd": 3, "tempo_peca_min": 4,
         "ref_interna": "R1", "tempos_operacao": {"Corte Laser": 4}},
        {"tipo_item": "service", "route": "montagem", "qtd": 2, "tempo_peca_min": 1},
    ]
    original = deepcopy(lines)
    result = build_order_lines(ports, lines, "OF-0001", "Laser1")
    # Previously per-operation time overwrote the order accumulator: 14 instead of 20.
    assert result.estimated_minutes == 20
    bucket = result.materials[0]["espessuras"][0]
    assert bucket["tempos_operacao"] == {"Corte Laser": 14, "Quinagem": 4}
    assert bucket["tempo_min"] == 14
    assert [piece["ref_interna"] for piece in bucket["pecas"]] == ["R1", "NEW1"]
    assert [piece["opp"] for piece in bucket["pecas"]] == ["OPP-0001-01", "OPP-0001-02"]
    assert result.assembly_items[0]["estado"] == "Pendente"
    assert result.references == [("R1", ""), ("NEW1", "")]
    bucket["pecas"][0]["tempos_operacao"]["Corte Laser"] = 99
    assert lines == original
    metalwork = [{"material": "Aco", "espessura": "5", "qtd": 1, "tempo_peca_min": 6,
                  "route": "serralharia", "ops": ["Corte Laser", "Serralharia"]}]
    result = build_order_lines(ports, metalwork, "", "Bancada")
    bucket = result.materials[0]["espessuras"][0]
    assert bucket["tempos_operacao"] == {"Serralharia": 6} and bucket["tempo_min"] == 0
    assert bucket["pecas"][0]["opp"] == "OPP-FALLBACK"
    component = build_order_lines(ports, [{"route": "conjunto", "qtd": 1}], "", "")
    assert component.assembly_items[0]["estado"] == "Componente" and component.materials == []
    try:
        build_order_lines(ports, [{"qtd": 1}], "", "")
    except ValueError:
        pass
    else:
        raise AssertionError("Missing production material accepted")
    assert "main" not in sys.modules and not any(name.startswith("PySide6") for name in sys.modules)
    print("quote-order-lines-ok total-time=yes operations=yes routing=yes references=yes detached=yes")


if __name__ == "__main__":
    main()
