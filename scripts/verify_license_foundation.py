from __future__ import annotations

import sys
import tempfile
from datetime import datetime, timedelta, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey

from lugest_core.licensing import (
    LicenseClaims,
    LicenseState,
    LicenseValidationError,
    build_license_token,
    canonical_payload,
    current_machine_fingerprint,
    decode_license_token,
    evaluate_license,
)
from lugest_infra.licensing import LicenseStore


def _expect_invalid(token: str, public_key_pem: bytes) -> None:
    try:
        decode_license_token(token, public_key_pem)
    except LicenseValidationError:
        return
    raise AssertionError("Uma licença adulterada foi aceite.")


def main() -> int:
    fingerprint = current_machine_fingerprint()
    assert len(fingerprint) == 19 and fingerprint.count("-") == 3
    assert current_machine_fingerprint() == fingerprint
    private_key = Ed25519PrivateKey.generate()
    public_key_pem = private_key.public_key().public_bytes(
        encoding=serialization.Encoding.PEM,
        format=serialization.PublicFormat.SubjectPublicKeyInfo,
    )
    now = datetime(2026, 9, 3, 12, 0, tzinfo=timezone.utc)
    claims = LicenseClaims(
        license_id="LIC-TEST-0001",
        customer_id="PT-509000000",
        customer_name="Cliente Industrial, Lda.",
        edition="Professional",
        issued_at=now - timedelta(days=1),
        not_before=now - timedelta(hours=1),
        expires_at=now + timedelta(days=365),
        max_workstations=3,
        max_users=25,
        features=("orcamentos", "producao", "stock"),
        device_fingerprints=("AAAA-BBBB-CCCC-DDDD",),
        metadata={"contract": "annual"},
    )
    payload = claims.to_mapping()
    signature = private_key.sign(canonical_payload(payload))
    token = build_license_token(payload, signature)
    decoded = decode_license_token(token, public_key_pem)
    assert decoded == claims

    active = evaluate_license(
        decoded,
        now_utc=now,
        device_fingerprint="aaaa-bbbb-cccc-dddd",
        required_features=("Orcamentos", "stock"),
        requested_workstations=3,
    )
    assert active.allowed and active.state is LicenseState.ACTIVE

    assert evaluate_license(decoded, now_utc=now - timedelta(days=3)).state is LicenseState.NOT_YET_VALID
    assert evaluate_license(decoded, now_utc=now + timedelta(days=366)).state is LicenseState.EXPIRED
    assert (
        evaluate_license(decoded, now_utc=now, device_fingerprint="OUTRO-EQUIPAMENTO").state
        is LicenseState.DEVICE_MISMATCH
    )
    assert (
        evaluate_license(decoded, now_utc=now, device_fingerprint="AAAA-BBBB-CCCC-DDDD", required_features=("faturacao",)).state
        is LicenseState.FEATURE_MISSING
    )
    assert (
        evaluate_license(decoded, now_utc=now, device_fingerprint="AAAA-BBBB-CCCC-DDDD", requested_workstations=4).state
        is LicenseState.WORKSTATION_LIMIT
    )

    parts = token.split(".")
    changed_payload = ("A" if parts[1][0] != "A" else "B") + parts[1][1:]
    _expect_invalid(".".join((parts[0], changed_payload, parts[2])), public_key_pem)
    # Change a significant Base64 character, not the trailing padding bits.
    changed_signature = ("A" if parts[2][0] != "A" else "B") + parts[2][1:]
    _expect_invalid(".".join((parts[0], parts[1], changed_signature)), public_key_pem)

    with tempfile.TemporaryDirectory(prefix="lugest-license-") as tmp:
        store = LicenseStore(Path(tmp) / "license.lugest")
        assert store.load() == ""
        saved_path = store.save(token)
        assert saved_path.exists()
        assert store.load() == token
        assert not list(saved_path.parent.glob("*.tmp"))

    print("license-foundation-ok signed=yes tamper=blocked policy=validated storage=atomic")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
