"""Licensing rules independent from UI, storage and customer infrastructure."""

from .license import (
    LicenseClaims,
    LicenseDecision,
    LicenseState,
    LicenseValidationError,
    build_license_token,
    canonical_payload,
    decode_license_token,
    evaluate_license,
)
from .device import current_machine_fingerprint
from .trusted_time import TrustedTimeError, portugal_datetime, trusted_time_snapshot

__all__ = [
    "LicenseClaims",
    "LicenseDecision",
    "LicenseState",
    "LicenseValidationError",
    "TrustedTimeError",
    "build_license_token",
    "canonical_payload",
    "current_machine_fingerprint",
    "decode_license_token",
    "evaluate_license",
    "portugal_datetime",
    "trusted_time_snapshot",
]
