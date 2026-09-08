"""Quote commands with explicit reference capabilities and aggregate persistence."""
from __future__ import annotations
from dataclasses import dataclass
from typing import Any, Callable, Protocol

class QuoteWriteRepository(Protocol):
    def get(self, number: str) -> dict | None: ...
    def next_number(self) -> str: ...
    def next_reference(self, client: str, reserved: list[str]) -> str: ...
    def save(self, quote: dict, *, direct: bool = True) -> None: ...
    def remove(self, number: str) -> None: ...

@dataclass(frozen=True)
class QuoteCommandRules:
    _active_client_ref_usage: Callable[..., Any]
    _known_client_ref_for_external: Callable[..., Any]
    _known_client_ref_pairs: Callable[..., Any]
    _normalize_orc_client: Callable[..., Any]
    _normalize_orc_line: Callable[..., Any]
    _normalize_quote_discount_groups: Callable[..., Any]
    _normalize_quote_discount_mode: Callable[..., Any]
    _normalize_supplier_reference: Callable[..., Any]
    _normalize_workcenter_value: Callable[..., Any]
    _parse_float: Callable[..., Any]
    _quote_default_delivery_text: Callable[..., Any]
    _quote_line_is_raw_material: Callable[..., Any]
    _quote_standard_iva_perc: Callable[..., Any]
    _ref_client_code: Callable[..., Any]
    _repair_orc_ref_history: Callable[..., Any]
    current_year: Callable[..., Any]
    now_iso: Callable[..., Any]
    orc_line_is_piece: Callable[..., Any]

class QuoteCommands:
    def __init__(self, repository: QuoteWriteRepository, rules: QuoteCommandRules):
        self.repository = repository
        self.rules = rules

    def save(self, payload: dict[str, Any]) -> str:
        numero = str(payload.get("numero", "") or "").strip()
        existing = self.repository.get(numero)
        if existing is None:
            year = self.rules.current_year()
            automatic_number = not numero or numero.upper().startswith(f"ORC-{year}-")
            if automatic_number:
                numero = str(self.repository.next_number())
        existing = self.repository.get(numero)
        posto_trabalho = self.rules._normalize_workcenter_value(payload.get("posto_trabalho", "") or (existing or {}).get("posto_trabalho", ""))
        client_payload = dict(payload.get("cliente", {}) or {})
        client = self.rules._normalize_orc_client(client_payload)
        client_code = self.rules._ref_client_code(client.get("codigo", ""))
        if not any(str(client.get(key, "") or "").strip() for key in ("codigo", "nome", "empresa")):
            raise ValueError("Cliente obrigatorio.")
        lines = [self.rules._normalize_orc_line(row) for row in list(payload.get("linhas", []) or [])]
        if client_code:
            self.rules._repair_orc_ref_history(client_code)
            # Repairs can save and replace the active snapshot. Subsequent
            # reference allocation must use that snapshot, not the old mapping.
            taken_refs, _pairs = self.rules._active_client_ref_usage(client_code, exclude_orc_numero=numero)
            reusable_pairs = self.rules._known_client_ref_pairs(client_code)
            seen_refs: set[str] = set()
            seen_pairs: set[tuple[str, str]] = set()
            reserved_refs = set(taken_refs)
            for line in lines:
                if not self.rules.orc_line_is_piece(line):
                    line["ref_interna"] = ""
                    continue
                if self.rules._quote_line_is_raw_material(line):
                    line["ref_interna"] = ""
                    continue
                ref_externa = str(line.get("ref_externa", "") or "").strip()
                known_ref = self.rules._known_client_ref_for_external(client_code, ref_externa)
                if known_ref:
                    line["ref_interna"] = known_ref
                ref_interna = str(line.get("ref_interna", "") or "").strip().upper()
                pair = (ref_externa, ref_interna)
                can_reuse_known = bool(ref_interna and pair in reusable_pairs)
                can_reuse_current = bool(ref_interna and pair in seen_pairs)
                if not ref_interna or ((ref_interna in seen_refs or ref_interna in reserved_refs) and not can_reuse_known and not can_reuse_current):
                    ref_interna = str(self.repository.next_reference(client_code, list(reserved_refs | seen_refs)))
                    line["ref_interna"] = ref_interna
                seen_refs.add(ref_interna)
                seen_pairs.add((ref_externa, ref_interna))
        iva_perc = self.rules._quote_standard_iva_perc()
        desconto_perc = round(max(0.0, min(100.0, self.rules._parse_float(payload.get("desconto_perc", (existing or {}).get("desconto_perc", 0)), 0))), 2)
        desconto_modo = self.rules._normalize_quote_discount_mode(payload.get("desconto_modo", (existing or {}).get("desconto_modo", "total")))
        desconto_grupos = self.rules._normalize_quote_discount_groups(payload.get("desconto_grupos", (existing or {}).get("desconto_grupos", [])))
        incremento_preco_perc = round(max(-100.0, self.rules._parse_float(payload.get("incremento_preco_perc", (existing or {}).get("incremento_preco_perc", 0)), 0)), 2)
        preco_transporte = round(self.rules._parse_float(payload.get("preco_transporte", 0), 0), 2)
        custo_transporte = round(self.rules._parse_float(payload.get("custo_transporte", (existing or {}).get("custo_transporte", 0)), 0), 2)
        paletes = round(self.rules._parse_float(payload.get("paletes", (existing or {}).get("paletes", 0)), 0), 2)
        peso_bruto_kg = round(self.rules._parse_float(payload.get("peso_bruto_kg", (existing or {}).get("peso_bruto_kg", 0)), 0), 2)
        volume_m3 = round(self.rules._parse_float(payload.get("volume_m3", (existing or {}).get("volume_m3", 0)), 0), 3)
        transportadora_id, transportadora_nome, _transportadora_contacto = self.rules._normalize_supplier_reference(
            payload.get("transportadora_id", (existing or {}).get("transportadora_id", "")),
            payload.get("transportadora_nome", (existing or {}).get("transportadora_nome", "")),
        )
        if "nota_transporte" in payload and not str(payload.get("nota_transporte", "") or "").strip() and preco_transporte <= 0:
            custo_transporte = 0.0
            transportadora_id = ""
            transportadora_nome = ""
            referencia_transporte = ""
            zona_transporte = ""
        else:
            referencia_transporte = str(payload.get("referencia_transporte", (existing or {}).get("referencia_transporte", "")) or "").strip()
            zona_transporte = str(payload.get("zona_transporte", (existing or {}).get("zona_transporte", "")) or "").strip()
        prazo_entrega_texto = str(
            payload.get(
                "prazo_entrega_texto",
                (existing or {}).get("prazo_entrega_texto", self.rules._quote_default_delivery_text()),
            )
            or self.rules._quote_default_delivery_text()
        ).strip()
        prazo_entrega_data = str(payload.get("prazo_entrega_data", (existing or {}).get("prazo_entrega_data", "")) or "").strip()[:10]
        subtotal_linhas = 0.0
        subtotal_com_desconto = 0.0
        desconto_valor = 0.0
        normalized_discount_groups = {item for item in desconto_grupos if item}
        for line in lines:
            line_total = round(self.rules._parse_float(line.get("total", 0), 0), 2)
            qtd = round(self.rules._parse_float(line.get("qtd", 0), 0), 2)
            preco_unit = round(self.rules._parse_float(line.get("preco_unit", 0), 0), 4)
            preco_unit_incrementado = round(max(0.0, preco_unit * (1.0 + (incremento_preco_perc / 100.0))), 4)
            line_total = round(qtd * preco_unit_incrementado, 2)
            line["preco_unit_incrementado"] = preco_unit_incrementado
            line["total"] = line_total
            subtotal_linhas = round(subtotal_linhas + line_total, 2)
            discount_key = str(line.get("discount_group_key", "") or "").strip()
            apply_discount = desconto_perc > 0 and (
                desconto_modo == "total"
                or not normalized_discount_groups
                or discount_key in normalized_discount_groups
            )
            if apply_discount:
                discounted_unit = round(preco_unit_incrementado * (1.0 - (desconto_perc / 100.0)), 4)
            else:
                discounted_unit = preco_unit_incrementado
            discounted_total = round(qtd * discounted_unit, 2)
            line_discount = round(max(0.0, line_total - discounted_total), 2)
            line["preco_unit_desconto"] = discounted_unit
            line["total_desconto"] = discounted_total
            line["desconto_aplicado"] = line_discount
            subtotal_com_desconto = round(subtotal_com_desconto + discounted_total, 2)
            desconto_valor = round(desconto_valor + line_discount, 2)
        subtotal_bruto = round(subtotal_linhas + preco_transporte, 2)
        subtotal = round(max(0.0, subtotal_com_desconto + preco_transporte), 2)
        total = round(subtotal * (1.0 + (iva_perc / 100.0)), 2)
        note = {
            "numero": numero,
            "data": str(payload.get("data", "") or existing.get("data", "") if isinstance(existing, dict) else "") or self.rules.now_iso(),
            "estado": str(payload.get("estado", "") or (existing or {}).get("estado", "") or "Em edição"),
            "cliente": client,
            "posto_trabalho": posto_trabalho,
            "linhas": lines,
            "iva_perc": iva_perc,
            "desconto_perc": desconto_perc,
            "desconto_modo": desconto_modo,
            "desconto_grupos": desconto_grupos,
            "incremento_preco_perc": incremento_preco_perc,
            "desconto_valor": desconto_valor,
            "preco_transporte": preco_transporte,
            "custo_transporte": custo_transporte,
            "paletes": paletes,
            "peso_bruto_kg": peso_bruto_kg,
            "volume_m3": volume_m3,
            "transportadora_id": transportadora_id,
            "transportadora_nome": transportadora_nome,
            "referencia_transporte": referencia_transporte,
            "zona_transporte": zona_transporte,
            "subtotal_linhas": subtotal_linhas,
            "subtotal_bruto": subtotal_bruto,
            "subtotal": subtotal,
            "total": total,
            "numero_encomenda": str(payload.get("numero_encomenda", "") or (existing or {}).get("numero_encomenda", "") or "").strip(),
            "ano": int(str(payload.get("ano", "") or (existing or {}).get("ano", "") or self.rules.current_year())),
            "executado_por": str(payload.get("executado_por", "") or (existing or {}).get("executado_por", "") or "").strip(),
            "nota_transporte": (
                str(payload.get("nota_transporte", "") or "").strip()
                if "nota_transporte" in payload
                else str((existing or {}).get("nota_transporte", "") or "").strip()
            ),
            "notas_pdf": str(payload.get("notas_pdf", "") or (existing or {}).get("notas_pdf", "") or "").strip(),
            "prazo_entrega_texto": prazo_entrega_texto,
            "prazo_entrega_data": prazo_entrega_data,
            "nota_cliente": str(payload.get("nota_cliente", "") or (existing or {}).get("nota_cliente", "") or "").strip(),
        }
        self.repository.save(note)
        return numero

    def remove(self, number: str) -> None:
        self.repository.remove(str(number or '').strip())

    def set_state(self, number: str, state: str) -> str:
        number = str(number or '').strip()
        quote = self.repository.get(number)
        if quote is None:
            raise ValueError('Orçamento não encontrado.')
        quote['estado'] = str(state or '').strip() or 'Em edição'
        self.repository.save(quote, direct=False)
        return number
