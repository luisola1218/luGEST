from __future__ import annotations

import json
import os
import shutil
import tempfile
from pathlib import Path
from typing import Any


class AtomicJsonStore:
    """JSON storage with bounded reads, atomic replacement and one recovery copy."""

    def __init__(self, path: Path | str, *, max_bytes: int = 4 * 1024 * 1024) -> None:
        self.path = Path(path)
        self.max_bytes = max(1024, int(max_bytes))

    @property
    def backup_path(self) -> Path:
        return self.path.with_suffix(self.path.suffix + ".bak")

    def _read(self, path: Path) -> Any:
        if path.stat().st_size > self.max_bytes:
            raise ValueError(f"O ficheiro JSON excede {self.max_bytes} bytes: {path}")
        return json.loads(path.read_text(encoding="utf-8-sig"))

    def load(self, default: Any = None) -> Any:
        for candidate in (self.path, self.backup_path):
            try:
                if candidate.is_file():
                    return self._read(candidate)
            except (OSError, UnicodeError, json.JSONDecodeError, ValueError):
                continue
        return default

    def save(self, payload: Any) -> Path:
        encoded = (json.dumps(payload, ensure_ascii=False, indent=2) + "\n").encode("utf-8")
        if len(encoded) > self.max_bytes:
            raise ValueError(f"O conteúdo JSON excede {self.max_bytes} bytes.")
        self.path.parent.mkdir(parents=True, exist_ok=True)
        if self.path.is_file():
            try:
                self._read(self.path)
                self._atomic_copy(self.path, self.backup_path)
            except (OSError, UnicodeError, json.JSONDecodeError, ValueError):
                pass
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

    @staticmethod
    def _atomic_copy(source: Path, target: Path) -> None:
        descriptor, temporary_name = tempfile.mkstemp(
            prefix=f".{target.name}.",
            suffix=".tmp",
            dir=str(target.parent),
        )
        os.close(descriptor)
        temporary_path = Path(temporary_name)
        try:
            shutil.copy2(source, temporary_path)
            os.replace(temporary_path, target)
        except Exception:
            temporary_path.unlink(missing_ok=True)
            raise


__all__ = ["AtomicJsonStore"]
