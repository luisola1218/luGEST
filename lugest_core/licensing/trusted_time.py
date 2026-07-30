from __future__ import annotations

import os
import threading
import time
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
from urllib.request import Request, urlopen


class TrustedTimeError(RuntimeError):
    """Raised when an online, TLS-protected time cannot be obtained."""


_DEFAULT_TIME_URLS = (
    "https://www.google.com/generate_204",
    "https://www.cloudflare.com/cdn-cgi/trace",
    "https://www.microsoft.com/",
)
_CACHE_LOCK = threading.RLock()
_CACHE_UTC: datetime | None = None
_CACHE_MONOTONIC = 0.0
_CACHE_SOURCE = ""
_CACHE_SOURCES: tuple[str, ...] = ()


def _configured_urls() -> tuple[str, ...]:
    # Deliberately not configurable from the customer environment: otherwise a
    # locally controlled endpoint could return an arbitrary historic Date.
    return _DEFAULT_TIME_URLS


def _timeout_seconds() -> float:
    try:
        return max(0.5, min(5.0, float(os.environ.get("LUGEST_TRUSTED_TIME_TIMEOUT_SEC", "1.8") or 1.8)))
    except Exception:
        return 1.8


def _cache_ttl_seconds() -> float:
    try:
        return max(30.0, min(1800.0, float(os.environ.get("LUGEST_TRUSTED_TIME_CACHE_SEC", "300") or 300)))
    except Exception:
        return 300.0


def _http_date_utc(url: str, timeout: float) -> datetime:
    request = Request(
        url,
        headers={
            "User-Agent": "LuGEST-ERP/1.0 trusted-time",
            "Accept": "*/*",
            "Cache-Control": "no-cache, no-store, max-age=0",
            "Pragma": "no-cache",
        },
        method="GET",
    )
    with urlopen(request, timeout=timeout) as response:
        raw_date = str(response.headers.get("Date", "") or "").strip()
        if not raw_date:
            raise TrustedTimeError(f"{url}: resposta sem cabeçalho Date.")
        parsed = parsedate_to_datetime(raw_date)
        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=timezone.utc)
        return parsed.astimezone(timezone.utc)


def _last_sunday(year: int, month: int) -> datetime:
    if month == 12:
        next_month = datetime(year + 1, 1, 1, tzinfo=timezone.utc)
    else:
        next_month = datetime(year, month + 1, 1, tzinfo=timezone.utc)
    last_day = next_month - timedelta(days=1)
    return last_day - timedelta(days=(last_day.weekday() + 1) % 7)


def portugal_datetime(utc_value: datetime) -> datetime:
    """Convert aware UTC to mainland Portugal using the current EU WET/WEST rules."""

    aware_utc = utc_value if utc_value.tzinfo is not None else utc_value.replace(tzinfo=timezone.utc)
    aware_utc = aware_utc.astimezone(timezone.utc)
    year = aware_utc.year
    dst_start = _last_sunday(year, 3).replace(hour=1)
    dst_end = _last_sunday(year, 10).replace(hour=1)
    is_summer = dst_start <= aware_utc < dst_end
    portugal_tz = timezone(timedelta(hours=1 if is_summer else 0), name="WEST" if is_summer else "WET")
    return aware_utc.astimezone(portugal_tz)


def _select_consistent_time(observations: list[tuple[datetime, str]]) -> tuple[datetime, tuple[str, ...]]:
    if not observations:
        raise TrustedTimeError("Nenhuma fonte de hora HTTPS respondeu.")
    ordered = sorted(observations, key=lambda item: item[0])
    if len(ordered) >= 2:
        spread = (ordered[-1][0] - ordered[0][0]).total_seconds()
        if spread > 300:
            raise TrustedTimeError("As fontes de hora HTTPS devolveram valores incompatíveis.")
    selected = ordered[len(ordered) // 2][0]
    return selected, tuple(source for _, source in ordered)


def trusted_time_snapshot(*, force: bool = False) -> dict:
    """Return UTC and Portugal time derived from HTTPS response Date headers.

    A short monotonic cache avoids network access during every UI refresh. The
    cache never consults the computer wall clock after a successful validation.
    """

    global _CACHE_UTC, _CACHE_MONOTONIC, _CACHE_SOURCE, _CACHE_SOURCES
    monotonic_now = time.monotonic()
    with _CACHE_LOCK:
        if (
            not force
            and _CACHE_UTC is not None
            and _CACHE_MONOTONIC > 0
            and monotonic_now - _CACHE_MONOTONIC <= _cache_ttl_seconds()
        ):
            current_utc = _CACHE_UTC + timedelta(seconds=max(0.0, monotonic_now - _CACHE_MONOTONIC))
            return {
                "utc": current_utc,
                "portugal": portugal_datetime(current_utc),
                "source": _CACHE_SOURCE,
                "sources": list(_CACHE_SOURCES),
                "cached": True,
            }

    observations: list[tuple[datetime, str]] = []
    errors: list[str] = []
    timeout = _timeout_seconds()
    for url in _configured_urls():
        try:
            observations.append((_http_date_utc(url, timeout), url))
            if len(observations) >= 2:
                break
        except Exception as exc:
            errors.append(f"{url}: {exc}")
    try:
        selected_utc, sources = _select_consistent_time(observations)
    except Exception as exc:
        detail = " | ".join(errors[:3])
        raise TrustedTimeError(f"Não foi possível validar a hora online. {exc} {detail}".strip()) from exc

    with _CACHE_LOCK:
        _CACHE_UTC = selected_utc
        _CACHE_MONOTONIC = time.monotonic()
        _CACHE_SOURCES = sources
        _CACHE_SOURCE = " + ".join(source.split("/")[2] for source in sources)
        return {
            "utc": selected_utc,
            "portugal": portugal_datetime(selected_utc),
            "source": _CACHE_SOURCE,
            "sources": list(sources),
            "cached": False,
        }


def reset_trusted_time_cache() -> None:
    global _CACHE_UTC, _CACHE_MONOTONIC, _CACHE_SOURCE, _CACHE_SOURCES
    with _CACHE_LOCK:
        _CACHE_UTC = None
        _CACHE_MONOTONIC = 0.0
        _CACHE_SOURCE = ""
        _CACHE_SOURCES = ()
