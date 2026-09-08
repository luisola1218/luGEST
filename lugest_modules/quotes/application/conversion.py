"""Coordinate quote conversion through an explicit staged repository."""
from dataclasses import dataclass
from typing import Any, Callable, Protocol
from lugest_modules.quotes.application.assemblies import technical_sheet
from lugest_modules.quotes.application.order_lines import OrderLinePorts, build_order_lines

class ConversionRepository(Protocol):
    def quote(self, number: str) -> dict[str, Any] | None: ...
    def clients(self) -> list[dict[str, Any]]: ...
    def allocate_client_code(self) -> str: ...
    def add_client(self, client: dict[str, Any]) -> None: ...
    def allocate_order_number(self) -> str: ...
    def allocate_order_code(self, order: dict[str, Any]) -> str: ...
    def line_ports(self, client_code: str) -> OrderLinePorts: ...
    def add_reference(self, internal: str, external: str) -> None: ...
    def save(self, quote: dict[str, Any], order: dict[str, Any]) -> None: ...

@dataclass(frozen=True)
class ConversionRules:
    normalize_client: Callable
    now_iso: Callable
    parse_float: Callable
    normalize_workcenter: Callable
    assembly_detail: Callable
    update_order_state: Callable

class QuoteConversion:
    def __init__(self, repository: ConversionRepository, rules: ConversionRules):
        self.repository = repository
        self.rules = rules

    def convert(self, numero: str, nota_cliente: str = "") -> str:
        numero = str(numero or "").strip()
        note = str(nota_cliente or "").strip()
        orc = self.repository.quote(numero)
        if orc is None:
            raise ValueError("Orçamento não encontrado.")
        if str(orc.get("numero_encomenda", "") or "").strip():
            raise ValueError("Orcamento ja convertido.")
        estado_norm = str(orc.get("estado", "") or "").strip().lower()
        if "aprovado" not in estado_norm:
            raise ValueError("Apenas orcamentos aprovados podem ser convertidos.")
        if not list(orc.get("linhas", []) or []):
            raise ValueError("Sem linhas para converter.")
        cli = self.rules.normalize_client(orc.get("cliente", {}))
        codigo = str(cli.get("codigo", "") or "").strip()
        if codigo and any(str(row.get("codigo", "") or "").strip() == codigo for row in self.repository.clients()):
            cliente_code = codigo
        else:
            cliente_code = ""
            for row in self.repository.clients():
                if not isinstance(row, dict):
                    continue
                if cli.get("nif") and str(row.get("nif", "") or "").strip() == str(cli.get("nif", "") or "").strip():
                    cliente_code = str(row.get("codigo", "") or "").strip()
                    break
                if cli.get("nome") and str(row.get("nome", "") or "").strip() == str(cli.get("nome", "") or "").strip():
                    cliente_code = str(row.get("codigo", "") or "").strip()
                    break
            if not cliente_code:
                cliente_code = str(self.repository.allocate_client_code())
                self.repository.add_client(
                    {
                        "codigo": cliente_code,
                        "nome": str(cli.get("nome", "") or "").strip(),
                        "nif": str(cli.get("nif", "") or "").strip(),
                        "morada": str(cli.get("morada", "") or "").strip(),
                        "contacto": str(cli.get("contacto", "") or "").strip(),
                        "email": str(cli.get("email", "") or "").strip(),
                        "observacoes": "",
                    }
                )
        alert_txt = (
            f"ALERTA: Encomenda gerada por conversao do orcamento {orc.get('numero')}. "
            "Confirmar dados de cliente, materiais, espessuras e prazos."
        )
        obs_txt = f"{alert_txt} | Origem: Orcamento {orc.get('numero')}"
        if note:
            obs_txt += f" | Nota cliente: {note}"
        enc = {
            "numero": self.repository.allocate_order_number(),
            "cliente": cliente_code,
            "nota_cliente": note,
            "nota_transporte": str(orc.get("nota_transporte", "") or "").strip(),
            "preco_transporte": round(self.rules.parse_float(orc.get("preco_transporte", 0), 0), 2),
            "custo_transporte": round(self.rules.parse_float(orc.get("custo_transporte", 0), 0), 2),
            "paletes": round(self.rules.parse_float(orc.get("paletes", 0), 0), 2),
            "peso_bruto_kg": round(self.rules.parse_float(orc.get("peso_bruto_kg", 0), 0), 2),
            "volume_m3": round(self.rules.parse_float(orc.get("volume_m3", 0), 0), 3),
            "transportadora_id": str(orc.get("transportadora_id", "") or "").strip(),
            "transportadora_nome": str(orc.get("transportadora_nome", "") or "").strip(),
            "referencia_transporte": str(orc.get("referencia_transporte", "") or "").strip(),
            "zona_transporte": str(orc.get("zona_transporte", "") or "").strip(),
            "local_descarga": str(cli.get("morada", "") or "").strip(),
            "transporte_numero": "",
            "estado_transporte": "",
            "data_criacao": self.rules.now_iso(),
            "data_entrega": str(orc.get("prazo_entrega_data", "") or "").strip()[:10],
            "tempo": 0.0,
            "tempo_estimado": 0.0,
            "cativar": False,
            "posto_trabalho": self.rules.normalize_workcenter(orc.get("posto_trabalho", "")),
            "observacoes": obs_txt,
            "alerta_conversao": True,
            "estado": "Preparacao",
            "materiais": [],
            "reservas": [],
            "montagem_itens": [],
            "numero_orcamento": orc.get("numero"),
            "tipo_encomenda": "Cliente",
            "produto_fichas": [],
        }
        sheet_groups: dict[str, dict[str, float]] = {}
        sheet_sources: dict[str, dict[str, Any]] = {}
        for source_line in list(orc.get("linhas", []) or []):
            sheet_code = str(source_line.get("conjunto_codigo", "") or "").strip()
            if not sheet_code:
                continue
            group_key = str(source_line.get("grupo_uuid", "") or "").strip() or sheet_code
            base_qty = self.rules.parse_float(source_line.get("qtd_base", 0), 0)
            line_qty = self.rules.parse_float(source_line.get("qtd", 0), 0)
            group_qty = (line_qty / base_qty) if base_qty > 0 and line_qty > 0 else 1.0
            code_groups = sheet_groups.setdefault(sheet_code, {})
            code_groups[group_key] = max(float(code_groups.get(group_key, 0) or 0), group_qty)
            if sheet_code in sheet_sources:
                continue
            try:
                stored_sheet = dict(self.rules.assembly_detail(sheet_code) or {})
            except Exception:
                stored_sheet = {}
            sheet_sources[sheet_code] = {
                "codigo": sheet_code,
                "param_codigo": str(stored_sheet.get("param_codigo", "") or source_line.get("conjunto_param_codigo", "") or "").strip(),
                "descricao": str(
                    stored_sheet.get("descricao", "")
                    or source_line.get("conjunto_nome", "")
                    or sheet_code
                ).strip(),
                "notas": str(stored_sheet.get("notas", "") or "").strip(),
                "ficha_tecnica": technical_sheet(
                    source_line.get("ficha_tecnica", {}) or stored_sheet.get("ficha_tecnica", {})
                ),
            }
        for sheet_code, snapshot in sheet_sources.items():
            snapshot["quantidade_conjuntos"] = round(
                max(1.0, sum(sheet_groups.get(sheet_code, {}).values())),
                2,
            )
            enc["produto_fichas"].append(snapshot)
        enc_of = str(self.repository.allocate_order_code(enc) or "").strip()
        prepared = build_order_lines(self.repository.line_ports(cliente_code),
                                     list(orc.get("linhas", []) or []), enc_of,
                                     enc.get("posto_trabalho", ""))
        enc["materiais"] = prepared.materials
        enc["montagem_itens"] = prepared.assembly_items
        enc["tempo_estimado"] = prepared.estimated_minutes
        enc["tempo"] = round(prepared.estimated_minutes / 60.0, 2) if prepared.estimated_minutes > 0 else 0.0
        for internal_ref, external_ref in prepared.references:
            self.repository.add_reference(internal_ref, external_ref)
        self.rules.update_order_state(enc)
        orc["numero_encomenda"] = enc["numero"]
        if note:
            orc["nota_cliente"] = note
        orc["estado"] = "Convertido em Encomenda"
        self.repository.save(orc, enc)
        return str(enc["numero"])
