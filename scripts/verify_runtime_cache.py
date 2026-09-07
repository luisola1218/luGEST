"""Cache regressions using an injected runtime: no MySQL or user data access."""
from pathlib import Path
import sys
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from lugest_qt.services.runtime_service import RuntimeService


def main() -> int:
    now = [100.0]
    calls = []
    source = {"rows": [{"quantity": 2}]}

    def dashboard(**kwargs):
        calls.append(kwargs)
        return source

    runtime = SimpleNamespace(get_dashboard=dashboard, get_mobile_alerts=lambda: source)
    service = RuntimeService(runtime, clock=lambda: now[0], max_cache_entries=2)
    first = service.dashboard()
    first["rows"][0]["quantity"] = 999
    source["rows"][0]["quantity"] = 7
    assert service.dashboard()["rows"][0]["quantity"] == 2
    assert len(calls) == 1
    now[0] += 3.0
    assert service.dashboard()["rows"][0]["quantity"] == 7
    assert len(calls) == 2, "The TTL boundary must expire"
    service.dashboard(force=True)
    assert len(calls) == 3
    service.dashboard(year="2025")
    service.dashboard()  # Refresh recency of default filters.
    service.dashboard(year="2026")
    assert len(service._cache) == 2
    previous = len(calls)
    service.dashboard(year="2025")
    assert len(calls) == previous + 1, "Least recently used filters must be evicted"
    service.alerts()
    now[0] += 4
    assert service._cache_get(("alerts",), ttl_sec=5) is not None
    now[0] += 1
    assert service._cache_get(("alerts",), ttl_sec=5) is None
    service.invalidate_cache()
    assert not service._cache
    try:
        RuntimeService(runtime, max_cache_entries=0)
    except ValueError:
        pass
    else:
        raise AssertionError("Unbounded cache configuration accepted")
    print("runtime-cache-ok nested-isolation=yes ttl=monotonic eviction=bounded force=yes")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
