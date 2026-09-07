from __future__ import annotations

import copy
from lugest_qt.services.bridge_helpers import _search_normalize, _search_terms
from typing import Any


class ClientsBackendMixin:
    """Legacy adapter for clients; see BACKEND_GUIDE.md."""

    def order_clients(self) -> list[dict[str, str]]:
        rows = []
        for client in list(self.ensure_data().get("clientes", []) or []):
            codigo = str(client.get("codigo", "") or "").strip()
            nome = str(client.get("nome", "") or "").strip()
            if not codigo:
                continue
            rows.append({"codigo": codigo, "nome": nome, "label": f"{codigo} - {nome}".strip(" -")})
        rows.sort(key=lambda item: item["codigo"])
        return rows

    def client_rows(self, filter_text: str = "") -> list[dict[str, Any]]:
        query_terms = _search_terms(filter_text)
        rows: list[dict[str, Any]] = []
        for raw in list(self.ensure_data().get("clientes", []) or []):
            row = {
                "codigo": str(raw.get("codigo", "") or "").strip(),
                "nome": str(raw.get("nome", "") or "").strip(),
                "nif": str(raw.get("nif", "") or "").strip(),
                "morada": str(raw.get("morada", "") or "").strip(),
                "latitude": str(raw.get("latitude", "") or "").strip(),
                "longitude": str(raw.get("longitude", "") or "").strip(),
                "contacto": str(raw.get("contacto", "") or "").strip(),
                "email": str(raw.get("email", "") or "").strip(),
                "observacoes": str(raw.get("observacoes", "") or "").strip(),
                "prazo_entrega": str(raw.get("prazo_entrega", "") or "").strip(),
                "cond_pagamento": str(raw.get("cond_pagamento", "") or "").strip(),
                "obs_tecnicas": str(raw.get("obs_tecnicas", "") or "").strip(),
            }
            haystack = _search_normalize(" ".join(str(value or "") for value in row.values()))
            if query_terms and not all(term in haystack for term in query_terms):
                continue
            rows.append(row)
        rows.sort(key=lambda item: (item.get("codigo") or "", item.get("nome") or ""))
        return rows

    def client_next_code(self) -> str:
        return str(self.desktop_main.next_cliente_codigo(self.ensure_data()))

    def client_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        data = self.ensure_data()
        codigo = str(payload.get("codigo", "") or "").strip() or str(self.desktop_main.next_cliente_codigo(data))
        nome = str(payload.get("nome", "") or "").strip()
        if not nome:
            raise ValueError("Nome do cliente obrigatorio.")
        row = {
            "codigo": codigo,
            "nome": nome,
            "nif": str(payload.get("nif", "") or "").strip(),
            "morada": str(payload.get("morada", "") or "").strip(),
            "latitude": str(payload.get("latitude", "") or "").strip(),
            "longitude": str(payload.get("longitude", "") or "").strip(),
            "contacto": str(payload.get("contacto", "") or "").strip(),
            "email": str(payload.get("email", "") or "").strip(),
            "observacoes": str(payload.get("observacoes", "") or "").strip(),
            "prazo_entrega": str(payload.get("prazo_entrega", "") or "").strip(),
            "cond_pagamento": str(payload.get("cond_pagamento", "") or "").strip(),
            "obs_tecnicas": str(payload.get("obs_tecnicas", "") or "").strip(),
        }
        rows = data.setdefault("clientes", [])
        existing = next((item for item in rows if str(item.get("codigo", "") or "").strip() == codigo), None)
        if existing is None:
            rows.append(row)
            target = row
        else:
            existing.update(row)
            target = existing
        upsert = getattr(self.desktop_main, "mysql_upsert_cliente", None)
        if bool(getattr(self.desktop_main, "_ASYNC_SAVE_ENABLED", False)):
            self._save(force=False)
        elif callable(upsert):
            upsert(target)
            if isinstance(self._base_data_snapshot, dict):
                base_rows = self._base_data_snapshot.setdefault("clientes", [])
                base_existing = next((item for item in base_rows if str(item.get("codigo", "") or "").strip() == codigo), None)
                if base_existing is None:
                    base_rows.append(copy.deepcopy(target))
                else:
                    base_existing.update(copy.deepcopy(target))
        else:
            self._save(force=True)
        return dict(target)

    def client_remove(self, codigo: str) -> None:
        data = self.ensure_data()
        code = str(codigo or "").strip()
        if not code:
            raise ValueError("Cliente inválido.")
        if any(str(enc.get("cliente", "") or "").strip() == code for enc in list(data.get("encomendas", []) or [])):
            raise ValueError("Nao e possivel remover um cliente usado em encomendas.")
        if any(str(self._normalize_orc_client(orc.get("cliente", {})).get("codigo", "") or "").strip() == code for orc in list(data.get("orcamentos", []) or [])):
            raise ValueError("Nao e possivel remover um cliente usado em orcamentos.")
        before = len(list(data.get("clientes", []) or []))
        data["clientes"] = [row for row in list(data.get("clientes", []) or []) if str(row.get("codigo", "") or "").strip() != code]
        if len(data["clientes"]) == before:
            raise ValueError("Cliente não encontrado.")
        self._save(force=True)
