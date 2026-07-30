from __future__ import annotations

import sys
from datetime import datetime, timezone
from pathlib import Path
from unittest.mock import patch

REPO_ROOT = Path(__file__).resolve().parents[1]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

import main as legacy
from lugest_core.licensing.trusted_time import TrustedTimeError, portugal_datetime, trusted_time_snapshot


def _config(**overrides):
    payload = legacy._trial_config_defaults()
    payload.update(
        {
            "enabled": True,
            "company_name": "Cliente Teste",
            "started_at": "2026-07-01T10:00:00+00:00",
            "duration_days": 30,
            "device_fingerprint": legacy.current_machine_fingerprint(),
        }
    )
    payload.update(overrides)
    return payload


def _snapshot(value: str):
    utc_value = datetime.fromisoformat(value).astimezone(timezone.utc)
    return {
        "utc": utc_value,
        "portugal": portugal_datetime(utc_value),
        "source": "time.test.invalid",
        "cached": False,
    }


def main() -> int:
    winter = portugal_datetime(datetime(2026, 1, 15, 12, tzinfo=timezone.utc))
    summer = portugal_datetime(datetime(2026, 7, 15, 12, tzinfo=timezone.utc))
    assert winter.utcoffset().total_seconds() == 0
    assert summer.utcoffset().total_seconds() == 3600

    saved = []
    with (
        patch.object(legacy, "load_trial_config", return_value=_config()),
        patch.object(legacy, "save_trial_config", side_effect=lambda payload: saved.append(dict(payload)) or payload),
        patch.object(legacy, "_trial_time_snapshot", return_value=_snapshot("2026-07-10T12:00:00+00:00")),
    ):
        status = legacy.get_trial_status(force_time=True)
    assert status["state"] == "active"
    assert status["blocking"] is False
    assert status["time_valid"] is True
    assert "+01:00" in status["portugal_time"]
    assert saved and saved[-1]["last_trusted_at"].startswith("2026-07-10T12:00:00")

    with (
        patch.object(
            legacy,
            "load_trial_config",
            return_value=_config(last_trusted_at="2026-07-15T12:00:00+00:00"),
        ),
        patch.object(legacy, "_trial_time_snapshot", return_value=_snapshot("2026-07-10T12:00:00+00:00")),
    ):
        rollback = legacy.get_trial_status(force_time=True)
    assert rollback["state"] == "time_rollback"
    assert rollback["blocking"] is True

    with (
        patch.object(legacy, "load_trial_config", return_value=_config()),
        patch.object(
            legacy,
            "_trial_time_snapshot",
            side_effect=TrustedTimeError("internet indisponível"),
        ),
    ):
        offline = legacy.get_trial_status(force_time=True)
    assert offline["state"] == "time_unavailable"
    assert offline["blocking"] is True

    with (
        patch.object(legacy, "load_trial_config", return_value=_config(duration_days=2)),
        patch.object(legacy, "_trial_time_snapshot", return_value=_snapshot("2026-07-10T12:00:00+00:00")),
    ):
        expired = legacy.get_trial_status(force_time=True)
    assert expired["state"] == "expired"
    assert expired["blocking"] is True

    online = trusted_time_snapshot(force=True)
    assert isinstance(online.get("utc"), datetime)
    assert str(online.get("source", "") or "").strip()
    print(
        "trial-trusted-time-ok",
        online["utc"].isoformat(),
        online["portugal"].isoformat(),
        online["source"],
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
