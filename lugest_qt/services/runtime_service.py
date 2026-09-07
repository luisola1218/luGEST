from __future__ import annotations

import os
import time
from collections import OrderedDict
from copy import deepcopy
from pathlib import Path
from typing import Callable


class RuntimeService:
    def __init__(self, runtime=None, *, clock: Callable[[], float] = time.monotonic, max_cache_entries: int = 128) -> None:
        if max_cache_entries < 1:
            raise ValueError("max_cache_entries must be positive")
        if runtime is None:
            from lugest_qt.services import pulse_runtime
            runtime = pulse_runtime
        self.runtime = runtime
        self._clock = clock
        self._max_cache_entries = max_cache_entries
        self._cache: OrderedDict[tuple, tuple[float, dict]] = OrderedDict()
        self._default_ttl_sec = 3.0

    def _cache_get(self, key: tuple, ttl_sec: float | None = None) -> dict | None:
        row = self._cache.get(key)
        if not row:
            return None
        loaded_at, payload = row
        ttl = self._default_ttl_sec if ttl_sec is None else max(0.0, float(ttl_sec or 0.0))
        if ttl <= 0 or (self._clock() - loaded_at) >= ttl:
            self._cache.pop(key, None)
            return None
        self._cache.move_to_end(key)
        return deepcopy(payload)

    def _cache_put(self, key: tuple, payload: dict) -> dict:
        data = deepcopy(dict(payload or {}))
        self._cache[key] = (self._clock(), data)
        self._cache.move_to_end(key)
        while len(self._cache) > self._max_cache_entries:
            self._cache.popitem(last=False)
        return deepcopy(data)

    def invalidate_cache(self) -> None:
        self._cache.clear()

    def dashboard(
        self,
        period: str = "7 dias",
        year: str | None = None,
        encomenda: str = "Todas",
        visao: str = "Todas",
        origem: str = "Ambos",
        *,
        force: bool = False,
    ) -> dict:
        key = ("dashboard", period, year or "", encomenda, visao, origem)
        if not force:
            cached = self._cache_get(key)
            if cached is not None:
                return cached
        return self._cache_put(
            key,
            self.runtime.get_dashboard(period=period, year=year, encomenda=encomenda, visao=visao, origem=origem),
        )

    def operator_board(
        self,
        year: str | None = None,
        username: str = "",
        role: str = "",
        *,
        force: bool = False,
    ) -> dict:
        key = ("operator_board", year or "", username, role)
        if not force:
            cached = self._cache_get(key)
            if cached is not None:
                return cached
        return self._cache_put(
            key,
            self.runtime.get_operator_board(year=year, username=username, role=role),
        )

    def planning_overview(
        self,
        year: str | None = None,
        week_start: str | None = None,
        operation: str | None = None,
        *,
        force: bool = False,
    ) -> dict:
        key = ("planning_overview", year or "", week_start or "", operation or "")
        if not force:
            cached = self._cache_get(key)
            if cached is not None:
                return cached
        return self._cache_put(
            key,
            self.runtime.get_planning_overview(year=year, week_start=week_start, operation=operation),
        )

    def planning_pdf(self, year: str | None = None, week_start: str | None = None, operation: str | None = None) -> Path:
        path = Path(self.runtime.build_planning_pdf(year=year, week_start=week_start, operation=operation))
        os.startfile(str(path))
        return path

    def avarias(self, *, force: bool = False) -> dict:
        key = ("avarias",)
        if not force:
            cached = self._cache_get(key, ttl_sec=5.0)
            if cached is not None:
                return cached
        return self._cache_put(key, self.runtime.get_mobile_avarias())

    def alerts(self, *, force: bool = False) -> dict:
        key = ("alerts",)
        if not force:
            cached = self._cache_get(key, ttl_sec=5.0)
            if cached is not None:
                return cached
        return self._cache_put(key, self.runtime.get_mobile_alerts())

    def pulse_plan_delay_set_reason(self, item_key: str, reason: str) -> dict:
        self.invalidate_cache()
        return self.runtime.set_plan_delay_reason(item_key=item_key, reason=reason)

    def pulse_plan_delay_clear_reason(self, item_key: str) -> dict:
        self.invalidate_cache()
        return self.runtime.clear_plan_delay_reason(item_key=item_key)
