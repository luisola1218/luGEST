from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping
from urllib.parse import urlparse


SHA256_PATTERN = re.compile(r"^[0-9a-f]{64}$", flags=re.IGNORECASE)
ALLOWED_REFERENCE_SCHEMES = {"https", "file"}


class UpdateManifestError(ValueError):
    """Raised when a release manifest cannot be trusted by the updater."""


def _required_text(payload: Mapping[str, Any], key: str, *, max_length: int = 2048) -> str:
    value = str(payload.get(key, "") or "").strip()
    if not value:
        raise UpdateManifestError(f"Manifest sem campo '{key}'.")
    if len(value) > max_length:
        raise UpdateManifestError(f"O campo '{key}' excede o tamanho permitido.")
    return value


def _validate_reference(value: str, key: str) -> str:
    parsed = urlparse(value)
    if parsed.scheme and len(parsed.scheme) > 1 and parsed.scheme.casefold() not in ALLOWED_REFERENCE_SCHEMES:
        raise UpdateManifestError(f"O campo '{key}' usa um protocolo não autorizado: {parsed.scheme}.")
    return value


def _validate_sha256(value: Any, key: str) -> str:
    digest = str(value or "").strip().casefold()
    if not SHA256_PATTERN.fullmatch(digest):
        raise UpdateManifestError(f"O campo '{key}' tem de conter um SHA-256 completo de 64 caracteres.")
    return digest


@dataclass(frozen=True)
class UpdateManifest:
    version: str
    package_url: str
    package_sha256: str
    bootstrap_url: str
    bootstrap_sha256: str
    channel: str = "stable"
    notes: str = ""
    schema_version: int = 1


def validate_update_manifest(payload: Mapping[str, Any]) -> UpdateManifest:
    if not isinstance(payload, Mapping):
        raise UpdateManifestError("O manifesto de atualização não é um objeto JSON válido.")
    try:
        schema_version = int(payload.get("schema_version", 1) or 1)
    except (TypeError, ValueError) as exc:
        raise UpdateManifestError("A versão do formato do manifesto não é numérica.") from exc
    if schema_version != 1:
        raise UpdateManifestError(f"Versão de manifesto não suportada: {schema_version}.")
    version = _required_text(payload, "version", max_length=80)
    if not re.search(r"\d", version):
        raise UpdateManifestError("A versão da atualização não contém qualquer componente numérico.")
    package_url = _validate_reference(_required_text(payload, "package_url"), "package_url")
    bootstrap_url = _validate_reference(_required_text(payload, "bootstrap_url"), "bootstrap_url")
    channel = str(payload.get("channel", "stable") or "stable").strip().casefold()
    if channel not in {"stable", "beta", "pilot"}:
        raise UpdateManifestError("O canal tem de ser stable, beta ou pilot.")
    return UpdateManifest(
        schema_version=schema_version,
        version=version,
        package_url=package_url,
        package_sha256=_validate_sha256(payload.get("sha256"), "sha256"),
        bootstrap_url=bootstrap_url,
        bootstrap_sha256=_validate_sha256(payload.get("bootstrap_sha256"), "bootstrap_sha256"),
        channel=channel,
        notes=str(payload.get("notes", "") or "").strip()[:20_000],
    )


def sha256_file(path: Path | str, *, chunk_size: int = 1024 * 1024) -> str:
    source = Path(path)
    digest = hashlib.sha256()
    with source.open("rb") as handle:
        while chunk := handle.read(max(4096, int(chunk_size))):
            digest.update(chunk)
    return digest.hexdigest()
