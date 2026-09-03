from __future__ import annotations

import hashlib
import os
import platform
import uuid


def current_machine_fingerprint() -> str:
    """Return the legacy-compatible workstation fingerprint.

    The format is kept stable for existing trials.  Commercial activation may
    later add a versioned fingerprint strategy without invalidating this value.
    """

    raw = "|".join(
        [
            str(platform.system() or "").strip().lower(),
            str(platform.machine() or "").strip().lower(),
            str(platform.node() or os.environ.get("COMPUTERNAME", "") or "").strip().lower(),
            str(uuid.getnode() or "").strip().lower(),
        ]
    )
    digest = hashlib.sha256(raw.encode("utf-8", errors="ignore")).hexdigest().upper()
    return f"{digest[:4]}-{digest[4:8]}-{digest[8:12]}-{digest[12:16]}"


__all__ = ["current_machine_fingerprint"]
