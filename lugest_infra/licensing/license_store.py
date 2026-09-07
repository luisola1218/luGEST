from __future__ import annotations

import os
import tempfile
from pathlib import Path

from lugest_infra.app_paths import AppPaths


class LicenseStore:
    """Small atomic store for a signed license token.

    License state is intentionally separated from business data and from the
    editable Qt configuration.  The token is public/signed data; no private
    signing material is ever stored here.
    """

    MAX_TOKEN_BYTES = 64 * 1024

    def __init__(self, path: Path | str | None = None) -> None:
        self.path = Path(path) if path is not None else self.default_path()

    @staticmethod
    def default_path() -> Path:
        return AppPaths(Path.cwd()).license_file()

    def load(self) -> str:
        if not self.path.exists():
            return ""
        if self.path.stat().st_size > self.MAX_TOKEN_BYTES:
            raise ValueError("O ficheiro de licença excede o tamanho permitido.")
        return self.path.read_text(encoding="utf-8").strip()

    def save(self, token: str) -> Path:
        clean = str(token or "").strip()
        if not clean:
            raise ValueError("A licença não pode ficar vazia.")
        encoded = clean.encode("utf-8")
        if len(encoded) > self.MAX_TOKEN_BYTES:
            raise ValueError("A licença excede o tamanho permitido.")
        self.path.parent.mkdir(parents=True, exist_ok=True)
        descriptor, temporary_name = tempfile.mkstemp(
            prefix=f".{self.path.name}.",
            suffix=".tmp",
            dir=str(self.path.parent),
        )
        temporary_path = Path(temporary_name)
        try:
            with os.fdopen(descriptor, "wb") as handle:
                handle.write(encoded)
                handle.flush()
                os.fsync(handle.fileno())
            os.replace(temporary_path, self.path)
        except Exception:
            temporary_path.unlink(missing_ok=True)
            raise
        return self.path


__all__ = ["LicenseStore"]
