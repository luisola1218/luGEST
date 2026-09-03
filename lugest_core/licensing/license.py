from __future__ import annotations

import base64
import json
from dataclasses import dataclass, field
from datetime import datetime, timezone
from enum import Enum
from typing import Any, Iterable, Mapping

from cryptography.exceptions import InvalidSignature
from cryptography.hazmat.primitives.serialization import load_pem_public_key


TOKEN_PREFIX = "LUGEST1"
SCHEMA_VERSION = 1


class LicenseValidationError(ValueError):
    """Raised when a license envelope is malformed or its signature is invalid."""


class LicenseState(str, Enum):
    ACTIVE = "active"
    NOT_YET_VALID = "not_yet_valid"
    EXPIRED = "expired"
    DEVICE_MISMATCH = "device_mismatch"
    FEATURE_MISSING = "feature_missing"
    WORKSTATION_LIMIT = "workstation_limit"


def _utc_datetime(value: Any, field_name: str, *, required: bool = False) -> datetime | None:
    text = str(value or "").strip()
    if not text:
        if required:
            raise LicenseValidationError(f"O campo {field_name} é obrigatório.")
        return None
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except (TypeError, ValueError) as exc:
        raise LicenseValidationError(f"O campo {field_name} não contém uma data ISO válida.") from exc
    if parsed.tzinfo is None:
        raise LicenseValidationError(f"O campo {field_name} tem de incluir fuso horário.")
    return parsed.astimezone(timezone.utc)


def _utc_text(value: datetime | None) -> str:
    if value is None:
        return ""
    aware = value if value.tzinfo is not None else value.replace(tzinfo=timezone.utc)
    return aware.astimezone(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")


def _clean_unique(values: Iterable[Any]) -> tuple[str, ...]:
    if isinstance(values, (str, bytes)):
        values = (values,)
    result: list[str] = []
    seen: set[str] = set()
    for raw in values:
        value = str(raw or "").strip()
        key = value.casefold()
        if not value or key in seen:
            continue
        seen.add(key)
        result.append(value)
    return tuple(result)


@dataclass(frozen=True)
class LicenseClaims:
    license_id: str
    customer_id: str
    customer_name: str
    edition: str
    issued_at: datetime
    not_before: datetime
    expires_at: datetime | None = None
    max_workstations: int = 1
    max_users: int = 0
    features: tuple[str, ...] = ()
    device_fingerprints: tuple[str, ...] = ()
    metadata: Mapping[str, Any] = field(default_factory=dict)
    schema_version: int = SCHEMA_VERSION

    def __post_init__(self) -> None:
        if self.schema_version != SCHEMA_VERSION:
            raise LicenseValidationError(f"Versão de licença não suportada: {self.schema_version}.")
        for field_name in ("license_id", "customer_id", "customer_name", "edition"):
            if not str(getattr(self, field_name, "") or "").strip():
                raise LicenseValidationError(f"O campo {field_name} é obrigatório.")
        if self.issued_at.tzinfo is None or self.not_before.tzinfo is None:
            raise LicenseValidationError("As datas da licença têm de incluir fuso horário.")
        if self.expires_at is not None and self.expires_at.tzinfo is None:
            raise LicenseValidationError("A validade da licença tem de incluir fuso horário.")
        if self.expires_at is not None and self.expires_at <= self.not_before:
            raise LicenseValidationError("A data de expiração tem de ser posterior ao início da validade.")
        if int(self.max_workstations) < 1:
            raise LicenseValidationError("A licença tem de permitir pelo menos um posto.")
        if int(self.max_users) < 0:
            raise LicenseValidationError("O limite de utilizadores não pode ser negativo.")

    @classmethod
    def from_mapping(cls, payload: Mapping[str, Any]) -> "LicenseClaims":
        if not isinstance(payload, Mapping):
            raise LicenseValidationError("O conteúdo da licença não é um objeto válido.")
        try:
            schema_version = int(payload.get("schema_version", 0) or 0)
            max_workstations = int(payload.get("max_workstations", 1) or 1)
            max_users = int(payload.get("max_users", 0) or 0)
        except (TypeError, ValueError) as exc:
            raise LicenseValidationError("Os limites da licença não são numéricos.") from exc
        raw_metadata = payload.get("metadata", {})
        metadata = dict(raw_metadata) if isinstance(raw_metadata, Mapping) else {}
        return cls(
            schema_version=schema_version,
            license_id=str(payload.get("license_id", "") or "").strip(),
            customer_id=str(payload.get("customer_id", "") or "").strip(),
            customer_name=str(payload.get("customer_name", "") or "").strip(),
            edition=str(payload.get("edition", "") or "").strip(),
            issued_at=_utc_datetime(payload.get("issued_at"), "issued_at", required=True),  # type: ignore[arg-type]
            not_before=_utc_datetime(payload.get("not_before"), "not_before", required=True),  # type: ignore[arg-type]
            expires_at=_utc_datetime(payload.get("expires_at"), "expires_at"),
            max_workstations=max_workstations,
            max_users=max_users,
            features=_clean_unique(payload.get("features", []) or []),
            device_fingerprints=_clean_unique(payload.get("device_fingerprints", []) or []),
            metadata=metadata,
        )

    def to_mapping(self) -> dict[str, Any]:
        return {
            "schema_version": self.schema_version,
            "license_id": self.license_id,
            "customer_id": self.customer_id,
            "customer_name": self.customer_name,
            "edition": self.edition,
            "issued_at": _utc_text(self.issued_at),
            "not_before": _utc_text(self.not_before),
            "expires_at": _utc_text(self.expires_at),
            "max_workstations": int(self.max_workstations),
            "max_users": int(self.max_users),
            "features": list(self.features),
            "device_fingerprints": list(self.device_fingerprints),
            "metadata": dict(self.metadata),
        }


@dataclass(frozen=True)
class LicenseDecision:
    state: LicenseState
    message: str
    claims: LicenseClaims
    missing_features: tuple[str, ...] = ()

    @property
    def allowed(self) -> bool:
        return self.state is LicenseState.ACTIVE


def canonical_payload(payload: Mapping[str, Any]) -> bytes:
    return json.dumps(
        dict(payload),
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")


def _b64url_encode(value: bytes) -> str:
    return base64.urlsafe_b64encode(value).decode("ascii").rstrip("=")


def _b64url_decode(value: str) -> bytes:
    text = str(value or "").strip()
    try:
        return base64.urlsafe_b64decode(text + ("=" * (-len(text) % 4)))
    except Exception as exc:
        raise LicenseValidationError("A licença contém Base64 inválido.") from exc


def build_license_token(payload: Mapping[str, Any], signature: bytes) -> str:
    """Build an envelope from a payload and externally produced signature.

    Signing deliberately remains outside the shipped application: commercial
    tooling may hold the private key, while customer builds only receive the
    public verification key.
    """

    encoded_payload = _b64url_encode(canonical_payload(payload))
    return f"{TOKEN_PREFIX}.{encoded_payload}.{_b64url_encode(signature)}"


def decode_license_token(token: str, public_key_pem: bytes | str) -> LicenseClaims:
    parts = str(token or "").strip().split(".")
    if len(parts) != 3 or parts[0] != TOKEN_PREFIX:
        raise LicenseValidationError("Formato de licença desconhecido.")
    payload_bytes = _b64url_decode(parts[1])
    signature = _b64url_decode(parts[2])
    try:
        public_key = load_pem_public_key(
            public_key_pem.encode("utf-8") if isinstance(public_key_pem, str) else public_key_pem
        )
        public_key.verify(signature, payload_bytes)  # type: ignore[attr-defined]
    except (InvalidSignature, TypeError, ValueError, AttributeError) as exc:
        raise LicenseValidationError("Assinatura da licença inválida.") from exc
    try:
        payload = json.loads(payload_bytes.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise LicenseValidationError("Conteúdo da licença inválido.") from exc
    return LicenseClaims.from_mapping(payload)


def evaluate_license(
    claims: LicenseClaims,
    *,
    now_utc: datetime,
    device_fingerprint: str = "",
    required_features: Iterable[str] = (),
    requested_workstations: int = 1,
) -> LicenseDecision:
    if now_utc.tzinfo is None:
        raise LicenseValidationError("A hora de validação tem de incluir fuso horário.")
    current = now_utc.astimezone(timezone.utc)
    not_before = claims.not_before.astimezone(timezone.utc)
    expires_at = claims.expires_at.astimezone(timezone.utc) if claims.expires_at is not None else None
    if current < not_before:
        return LicenseDecision(LicenseState.NOT_YET_VALID, "A licença ainda não entrou em vigor.", claims)
    if expires_at is not None and current >= expires_at:
        return LicenseDecision(LicenseState.EXPIRED, "A licença expirou.", claims)

    allowed_devices = {value.casefold() for value in claims.device_fingerprints}
    current_device = str(device_fingerprint or "").strip().casefold()
    if allowed_devices and current_device not in allowed_devices:
        return LicenseDecision(
            LicenseState.DEVICE_MISMATCH,
            "Este equipamento não está autorizado pela licença.",
            claims,
        )

    licensed_features = {value.casefold() for value in claims.features}
    requested = _clean_unique(required_features)
    missing = tuple(value for value in requested if value.casefold() not in licensed_features)
    if missing:
        return LicenseDecision(
            LicenseState.FEATURE_MISSING,
            "A licença não inclui todos os módulos solicitados.",
            claims,
            missing_features=missing,
        )

    try:
        workstation_count = max(1, int(requested_workstations or 1))
    except (TypeError, ValueError):
        workstation_count = 1
    if workstation_count > claims.max_workstations:
        return LicenseDecision(
            LicenseState.WORKSTATION_LIMIT,
            "Foi ultrapassado o número de postos autorizado.",
            claims,
        )
    return LicenseDecision(LicenseState.ACTIVE, "Licença válida.", claims)


__all__ = [
    "LicenseClaims",
    "LicenseDecision",
    "LicenseState",
    "LicenseValidationError",
    "TOKEN_PREFIX",
    "build_license_token",
    "canonical_payload",
    "decode_license_token",
    "evaluate_license",
]
