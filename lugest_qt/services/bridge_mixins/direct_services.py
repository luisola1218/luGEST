from __future__ import annotations

import copy
from datetime import datetime, timedelta
from typing import Any


class DirectServicesBridgeMixin:
    """Short commercial flow for field services that bypasses production."""

    DIRECT_SERVICE_STATES = ("Rascunho", "Confirmado", "Faturado", "Anulado")

    def _direct_services(self) -> list[dict[str, Any]]:
        return self.ensure_data().setdefault("servicos_diretos", [])

    def _direct_service_find(self, numero: str) -> dict[str, Any] | None:
        target = str(numero or "").strip()
        if not target:
            return None
        return next(
            (
                row
                for row in self._direct_services()
                if str((row or {}).get("numero", "") or "").strip() == target
            ),
            None,
        )

    def _direct_service_next_number(self) -> str:
        year = str(self.desktop_main.datetime.now().year)
        highest = 0
        prefix = f"SRV-{year}-"
        for row in self._direct_services():
            raw = str((row or {}).get("numero", "") or "").strip().upper()
            if not raw.startswith(prefix):
                continue
            try:
                highest = max(highest, int(raw.rsplit("-", 1)[-1]))
            except Exception:
                continue
        return f"{prefix}{highest + 1:04d}"

    def _direct_service_client(self, code: str) -> dict[str, Any] | None:
        target = str(code or "").strip()
        return next(
            (
                row
                for row in list(self.ensure_data().get("clientes", []) or [])
                if str((row or {}).get("codigo", "") or "").strip() == target
            ),
            None,
        )

    def _direct_service_normalize_line(self, payload: dict[str, Any], index: int) -> dict[str, Any]:
        kind = str(payload.get("kind", "") or "service").strip().lower()
        if kind not in {"set", "product", "material", "service"}:
            kind = "service"
        qty = round(self._parse_float(payload.get("qty", 1), 1), 3)
        unit_price = round(self._parse_float(payload.get("unit_price", 0), 0), 4)
        iva_perc = round(self._parse_float(payload.get("iva_perc", 23), 23), 2)
        if qty <= 0:
            raise ValueError(f"A quantidade da linha {index} tem de ser superior a zero.")
        description = str(payload.get("description", "") or "").strip()
        if not description:
            raise ValueError(f"Indica a descrição da linha {index}.")
        subtotal = round(qty * unit_price, 2)
        tax_value = round(subtotal * iva_perc / 100.0, 2)
        return {
            "id": str(payload.get("id", "") or self.desktop_main.uuid.uuid4().hex[:12].upper()).strip(),
            "kind": kind,
            "ref": str(payload.get("ref", "") or "").strip(),
            "description": description,
            "qty": qty,
            "unit": str(payload.get("unit", "") or ("SV" if kind == "service" else "UN")).strip() or "UN",
            "unit_price": unit_price,
            "iva_perc": iva_perc,
            "subtotal": subtotal,
            "valor_iva": tax_value,
            "total": round(subtotal + tax_value, 2),
        }

    def _direct_service_recalculate(self, service: dict[str, Any]) -> dict[str, Any]:
        lines = [
            self._direct_service_normalize_line(dict(row or {}), index)
            for index, row in enumerate(list(service.get("linhas", []) or []), start=1)
        ]
        service["linhas"] = lines
        service["subtotal"] = round(sum(self._parse_float(row.get("subtotal", 0), 0) for row in lines), 2)
        service["valor_iva"] = round(sum(self._parse_float(row.get("valor_iva", 0), 0) for row in lines), 2)
        service["total"] = round(sum(self._parse_float(row.get("total", 0), 0) for row in lines), 2)
        return service

    def direct_service_dashboard(self) -> dict[str, Any]:
        rows = self.direct_service_rows("", "Todos")
        return {
            "draft_count": sum(1 for row in rows if row.get("estado") == "Rascunho"),
            "confirmed_count": sum(1 for row in rows if row.get("estado") == "Confirmado"),
            "pending_billing_count": sum(1 for row in rows if row.get("estado") == "Confirmado" and not row.get("faturacao_numero")),
            "billed_count": sum(1 for row in rows if row.get("estado") == "Faturado"),
            "active_total": round(sum(self._parse_float(row.get("total", 0), 0) for row in rows if row.get("estado") != "Anulado"), 2),
            "row_count": len(rows),
        }

    def direct_service_rows(self, filter_text: str = "", state_filter: str = "Todos") -> list[dict[str, Any]]:
        query = self.desktop_main.norm_text(filter_text)
        state = self.desktop_main.norm_text(state_filter)
        rows: list[dict[str, Any]] = []
        for raw in self._direct_services():
            if not isinstance(raw, dict):
                continue
            row = self._direct_service_recalculate(raw)
            estado = str(row.get("estado", "") or "Rascunho").strip() or "Rascunho"
            if state not in {"", "todos", "todas", "all"} and self.desktop_main.norm_text(estado) != state:
                continue
            search_values = " ".join(
                str(value or "")
                for value in (
                    row.get("numero"), row.get("cliente_codigo"), row.get("cliente_nome"),
                    row.get("local_servico"), row.get("responsavel"), row.get("faturacao_numero"),
                )
            )
            if query and query not in self.desktop_main.norm_text(search_values):
                continue
            rows.append(
                {
                    "numero": str(row.get("numero", "") or "").strip(),
                    "data_servico": str(row.get("data_servico", "") or "").strip()[:10],
                    "cliente_codigo": str(row.get("cliente_codigo", "") or "").strip(),
                    "cliente_nome": str(row.get("cliente_nome", "") or "").strip(),
                    "cliente": f"{str(row.get('cliente_codigo', '') or '').strip()} - {str(row.get('cliente_nome', '') or '').strip()}".strip(" -"),
                    "estado": estado,
                    "linhas": len(list(row.get("linhas", []) or [])),
                    "subtotal": round(self._parse_float(row.get("subtotal", 0), 0), 2),
                    "valor_iva": round(self._parse_float(row.get("valor_iva", 0), 0), 2),
                    "total": round(self._parse_float(row.get("total", 0), 0), 2),
                    "faturacao_numero": str(row.get("faturacao_numero", "") or "").strip(),
                    "stock_consumido": bool(row.get("stock_consumido", False)),
                }
            )
        rows.sort(key=lambda item: (item.get("data_servico", ""), item.get("numero", "")), reverse=True)
        return rows

    def direct_service_detail(self, numero: str) -> dict[str, Any]:
        row = self._direct_service_find(numero)
        if row is None:
            raise ValueError("Serviço direto não encontrado.")
        self._direct_service_recalculate(row)
        return copy.deepcopy(row)

    def direct_service_create(self) -> dict[str, Any]:
        today = str(self.desktop_main.now_iso())[:10]
        try:
            due = (datetime.strptime(today, "%Y-%m-%d") + timedelta(days=30)).strftime("%Y-%m-%d")
        except Exception:
            due = today
        row = {
            "numero": self._direct_service_next_number(),
            "cliente_codigo": "",
            "cliente_nome": "",
            "data_servico": today,
            "data_vencimento": due,
            "estado": "Rascunho",
            "local_servico": "",
            "responsavel": str((self.user or {}).get("username", "") or "").strip(),
            "obs": "",
            "linhas": [],
            "subtotal": 0.0,
            "valor_iva": 0.0,
            "total": 0.0,
            "stock_consumido": False,
            "faturacao_numero": "",
            "created_at": self.desktop_main.now_iso(),
            "updated_at": self.desktop_main.now_iso(),
            "confirmado_at": "",
            "anulado_at": "",
        }
        self._direct_services().append(row)
        self._save(force=True)
        return copy.deepcopy(row)

    def direct_service_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        numero = str(payload.get("numero", "") or "").strip()
        target = self._direct_service_find(numero)
        if target is None:
            raise ValueError("Serviço direto não encontrado.")
        if str(target.get("estado", "") or "Rascunho") != "Rascunho":
            raise ValueError("Depois de confirmado, o serviço fica fechado. Anula o documento e cria outro para corrigir.")
        client_code = str(payload.get("cliente_codigo", "") or "").strip()
        client = self._direct_service_client(client_code)
        if client_code and client is None:
            raise ValueError("Cliente não encontrado.")
        lines = [dict(row or {}) for row in list(payload.get("linhas", []) or [])]
        target.update(
            {
                "cliente_codigo": client_code,
                "cliente_nome": str((client or {}).get("nome", "") or payload.get("cliente_nome", "") or "").strip(),
                "data_servico": str(payload.get("data_servico", "") or target.get("data_servico", "") or self.desktop_main.now_iso())[:10],
                "data_vencimento": str(payload.get("data_vencimento", "") or target.get("data_vencimento", ""))[:10],
                "local_servico": str(payload.get("local_servico", "") or "").strip(),
                "responsavel": str(payload.get("responsavel", "") or "").strip(),
                "obs": str(payload.get("obs", "") or "").strip(),
                "linhas": lines,
                "updated_at": self.desktop_main.now_iso(),
            }
        )
        self._direct_service_recalculate(target)
        self._save(force=True)
        return copy.deepcopy(target)

    def direct_service_catalog(self, kind: str, filter_text: str = "") -> list[dict[str, Any]]:
        kind_txt = str(kind or "").strip().lower()
        if kind_txt == "set":
            return [
                {
                    "kind": "set", "ref": row.get("codigo", ""), "description": row.get("descricao", ""),
                    "unit": "UN", "unit_price": row.get("total_final", 0), "available": "Componentes validados na confirmação",
                }
                for row in self.conjunto_rows(filter_text)
                if bool(row.get("ativo", True))
            ]
        if kind_txt == "product":
            return [
                {
                    "kind": "product", "ref": row.get("codigo", ""), "description": row.get("descricao", ""),
                    "unit": row.get("unid", "UN"), "unit_price": row.get("preco_venda", row.get("preco_unid", 0)),
                    "available": row.get("available_qty", 0),
                }
                for row in self.product_rows(filter_text, False)
            ]
        if kind_txt == "material":
            result: list[dict[str, Any]] = []
            for item in self.material_rows(filter_text, False):
                record = dict(item.get("record", {}) or {})
                preview = self.material_price_preview(record)
                available = max(0.0, self._parse_float(record.get("quantidade", 0), 0) - self._parse_float(record.get("reservado", 0), 0))
                formato = str(record.get("formato", "") or "-").strip() or "-"
                comprimento = self._parse_float(preview.get("comprimento", record.get("comprimento", 0)), 0)
                largura = self._parse_float(preview.get("largura", record.get("largura", 0)), 0)
                metros = self._parse_float(preview.get("metros", record.get("metros", 0)), 0)
                format_key = formato.casefold()
                preview_dimensions = str(preview.get("dimension_label", "") or "-").strip() or "-"
                if format_key == "chapa" and comprimento > 0 and largura > 0:
                    dimensions = f"{self._fmt(comprimento)} x {self._fmt(largura)} mm"
                elif metros > 0:
                    dimensions = f"{self._fmt(metros)} m/barra"
                elif comprimento > 0 and largura > 0:
                    dimensions = f"{self._fmt(comprimento)} x {self._fmt(largura)} mm"
                elif comprimento > 0:
                    dimensions = f"{self._fmt(comprimento)} mm"
                else:
                    dimensions = preview_dimensions
                section = str(preview.get("secao_tipo", record.get("secao_tipo", "")) or "").strip().upper()
                thickness = self._parse_float(record.get("espessura", 0), 0)
                if section:
                    nominal = self._parse_float(preview.get("altura", record.get("altura", thickness)), 0)
                    if format_key == "perfil":
                        specification = f"{section} {self._fmt(nominal)}" if nominal > 0 else section
                    else:
                        specification = preview_dimensions if preview_dimensions != "-" else (
                            f"{section} {self._fmt(nominal)}" if nominal > 0 else section
                        )
                elif thickness > 0:
                    specification = f"{self._fmt(thickness)} mm"
                else:
                    specification = "-"
                result.append(
                    {
                        "kind": "material",
                        "ref": str(record.get("id", "") or "").strip(),
                        "material": str(record.get("material", "") or "-").strip() or "-",
                        "format": formato,
                        "dimensions": dimensions,
                        "specification": specification,
                        "lot": str(record.get("lote_interno", "") or "-").strip() or "-",
                        "description": " | ".join(
                            value for value in (
                                str(record.get("material", "") or "").strip(),
                                formato,
                                dimensions,
                                specification,
                                str(record.get("lote_interno", "") or "").strip(),
                            ) if value
                        ),
                        "unit": "UN",
                        "unit_price": preview.get("preco_unid", 0),
                        "available": available,
                    }
                )
            return result
        return []

    def _direct_service_stock_requirements(self, service: dict[str, Any]) -> tuple[dict[str, float], dict[str, float]]:
        products: dict[str, float] = {}
        materials: dict[str, float] = {}

        def add(target: dict[str, float], key: str, value: Any) -> None:
            clean_key = str(key or "").strip()
            qty = self._parse_float(value, 0)
            if clean_key and qty > 0:
                target[clean_key] = round(target.get(clean_key, 0.0) + qty, 4)

        for line in list(service.get("linhas", []) or []):
            kind = str(line.get("kind", "") or "").strip().lower()
            qty = self._parse_float(line.get("qty", 0), 0)
            ref = str(line.get("ref", "") or "").strip()
            if kind == "product":
                add(products, ref, qty)
            elif kind == "material":
                add(materials, ref, qty)
            elif kind == "set":
                for item in self.conjunto_expand(ref, qty):
                    item_qty = self._parse_float(item.get("qtd", 0), 0)
                    product_code = str(item.get("produto_codigo", "") or "").strip()
                    material_id = str(item.get("stock_material_id", "") or "").strip()
                    if product_code:
                        add(products, product_code, item_qty)
                    elif material_id:
                        add(materials, material_id, item_qty)
        return products, materials

    def direct_service_confirm(self, numero: str) -> dict[str, Any]:
        target = self._direct_service_find(numero)
        if target is None:
            raise ValueError("Serviço direto não encontrado.")
        if str(target.get("estado", "") or "Rascunho") != "Rascunho":
            raise ValueError("Apenas um rascunho pode ser confirmado.")
        self._direct_service_recalculate(target)
        if not str(target.get("cliente_codigo", "") or "").strip():
            raise ValueError("Seleciona o cliente antes de confirmar.")
        if not list(target.get("linhas", []) or []):
            raise ValueError("Adiciona pelo menos uma linha ao serviço.")
        products, materials = self._direct_service_stock_requirements(target)
        data = self.ensure_data()
        for code, qty in products.items():
            product = next((row for row in list(data.get("produtos", []) or []) if str(row.get("codigo", "") or "").strip() == code), None)
            if product is None:
                raise ValueError(f"O produto {code} do serviço/conjunto não existe no stock.")
            available = self._parse_float(product.get("qty", 0), 0)
            if qty > available + 1e-9:
                raise ValueError(f"Stock insuficiente do produto {code}: necessário {self._fmt(qty)}, disponível {self._fmt(available)}.")
        for material_id, qty in materials.items():
            material = self.material_by_id(material_id)
            if material is None:
                raise ValueError(f"O material {material_id} do serviço/conjunto não existe no stock.")
            if self._material_quality_is_blocked(material):
                raise ValueError(f"O material {material_id} está bloqueado pela Qualidade.")
            available = max(0.0, self._parse_float(material.get("quantidade", 0), 0) - self._parse_float(material.get("reservado", 0), 0))
            if qty > available + 1e-9:
                raise ValueError(f"Stock insuficiente do material {material_id}: necessário {self._fmt(qty)}, disponível {self._fmt(available)}.")

        backup = {
            key: copy.deepcopy(data.get(key, []))
            for key in ("produtos", "produtos_mov", "materiais", "stock_log")
        }
        try:
            for code, qty in products.items():
                self.product_consume(code, qty, f"Serviço direto {numero}", issue_mode="stock")
            if materials:
                self.consume_material_allocations(
                    [{"material_id": material_id, "quantidade": qty} for material_id, qty in materials.items()],
                    reason=f"Serviço direto {numero}",
                )
            target["stock_consumido"] = True
            target["estado"] = "Confirmado"
            target["confirmado_at"] = self.desktop_main.now_iso()
            target["updated_at"] = self.desktop_main.now_iso()
            self._save(force=True)
        except Exception:
            for key, value in backup.items():
                data[key] = value
            self._save(force=True)
            raise
        return copy.deepcopy(target)

    def direct_service_cancel(self, numero: str, reason: str = "") -> dict[str, Any]:
        target = self._direct_service_find(numero)
        if target is None:
            raise ValueError("Serviço direto não encontrado.")
        if str(target.get("estado", "") or "") == "Faturado":
            raise ValueError("Um serviço já enviado para faturação deve ser regularizado no módulo Faturação.")
        if str(target.get("estado", "") or "") == "Anulado":
            return copy.deepcopy(target)
        target["estado"] = "Anulado"
        target["anulado_at"] = self.desktop_main.now_iso()
        target["anulado_motivo"] = str(reason or "").strip()
        target["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return copy.deepcopy(target)

    def direct_service_remove(self, numero: str) -> None:
        target = self._direct_service_find(numero)
        if target is None:
            raise ValueError("Serviço direto não encontrado.")
        if str(target.get("estado", "") or "Rascunho") != "Rascunho":
            raise ValueError("Só é possível remover rascunhos. Usa Anular para documentos confirmados.")
        self.ensure_data()["servicos_diretos"] = [
            row for row in self._direct_services() if str((row or {}).get("numero", "") or "").strip() != str(numero or "").strip()
        ]
        self._save(force=True)

    def direct_service_send_to_billing(self, numero: str) -> dict[str, Any]:
        target = self._direct_service_find(numero)
        if target is None:
            raise ValueError("Serviço direto não encontrado.")
        if str(target.get("estado", "") or "") not in {"Confirmado", "Faturado"}:
            raise ValueError("Confirma o serviço antes de o enviar para Faturação.")
        record = self._billing_create_record_from_source("service", str(numero or "").strip())
        target["faturacao_numero"] = str(record.get("numero", "") or "").strip()
        target["estado"] = "Faturado"
        target["updated_at"] = self.desktop_main.now_iso()
        self._save(force=True)
        return {
            "service": copy.deepcopy(target),
            "billing": self.billing_detail(target["faturacao_numero"]),
        }
