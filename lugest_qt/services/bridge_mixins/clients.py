from __future__ import annotations

from lugest_qt.services.bridge_helpers import _search_normalize
from lugest_modules.clients.api import ClientService
from lugest_modules.clients.infrastructure.legacy_repository import LegacyClientRepository
from typing import Any


class ClientsBackendMixin:
    """Legacy adapter for clients; see BACKEND_GUIDE.md."""

    def order_clients(self) -> list[dict[str, str]]:
        return self._client_service().order_options()

    def client_rows(self, filter_text: str = "") -> list[dict[str, Any]]:
        return self._client_service().rows(filter_text)

    def client_next_code(self) -> str:
        return self._client_service().next_code()

    def client_save(self, payload: dict[str, Any]) -> dict[str, Any]:
        return self._client_service().save(payload)

    def client_remove(self, codigo: str) -> None:
        return self._client_service().remove(codigo)

    def _client_service(self) -> ClientService:
        upsert = getattr(self.desktop_main, "mysql_upsert_cliente", None)
        repository = LegacyClientRepository(
            get_data=self.ensure_data,
            get_baseline=lambda: self._base_data_snapshot,
            save_dataset=self._save,
            next_code=self.desktop_main.next_cliente_codigo,
            normalize_quote_client=self._normalize_orc_client,
            upsert=upsert if callable(upsert) else None,
            asynchronous=bool(getattr(self.desktop_main, "_ASYNC_SAVE_ENABLED", False)),
        )
        return ClientService(repository, normalize_search=_search_normalize)
