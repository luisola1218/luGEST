from __future__ import annotations

import os
import re
import sqlite3
import threading
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


_JOB_ID = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]{0,29}$")
_DIGEST = re.compile(r"^[a-f0-9]{64}$")
_MIME_EXTENSIONS = {
    "image/jpeg": ".jpg",
    "image/png": ".png",
    "image/webp": ".webp",
    "image/heic": ".heic",
}


def validate_job_id(value: Any) -> str:
    job_id = str(value or "").strip()
    if not _JOB_ID.fullmatch(job_id):
        raise ValueError("Número de serviço inválido.")
    return job_id


def detect_image_mime(data: bytes) -> str | None:
    if data.startswith(b"\xff\xd8\xff"):
        return "image/jpeg"
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return "image/png"
    if len(data) >= 12 and data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return "image/webp"
    if len(data) >= 12 and data[4:8] == b"ftyp" and data[8:12] in {
        b"heic",
        b"heix",
        b"hevc",
        b"mif1",
    }:
        return "image/heic"
    return None


class AttachmentStore:
    """Arquivo privado da app móvel, completamente separado da base do desktop."""

    def __init__(self, root: Path):
        self.root = root.resolve()
        self.files_root = self.root / "files"
        self.database_path = self.root / "attachments.sqlite3"
        self._lock = threading.RLock()

    def save(
        self,
        *,
        job_id: str,
        kind: str,
        original_name: str,
        declared_mime: str,
        digest: str,
        data: bytes,
        device_id: str,
    ) -> dict[str, Any]:
        job_id = validate_job_id(job_id)
        kind = str(kind or "").strip().lower()
        if kind not in {"photo", "signature"}:
            raise ValueError("Tipo de anexo inválido.")
        digest = str(digest or "").strip().lower()
        if not _DIGEST.fullmatch(digest):
            raise ValueError("Identificador do anexo inválido.")
        detected_mime = detect_image_mime(data)
        if detected_mime is None:
            raise ValueError("O anexo não é uma imagem JPEG, PNG, WebP ou HEIC válida.")
        if declared_mime and declared_mime != detected_mime:
            raise ValueError("O formato declarado não corresponde ao ficheiro enviado.")

        extension = _MIME_EXTENSIONS[detected_mime]
        relative_path = Path(job_id) / f"{digest}{extension}"
        target = self.files_root / relative_path
        created_at = datetime.now(timezone.utc).isoformat()
        safe_original = Path(str(original_name or "anexo")).name[:180]

        with self._lock:
            target.parent.mkdir(parents=True, exist_ok=True)
            if not target.exists():
                temporary = target.with_name(f".{target.name}.{os.getpid()}.tmp")
                temporary.write_bytes(data)
                os.replace(temporary, target)
            self._initialize()
            with sqlite3.connect(self.database_path, timeout=10) as connection:
                connection.execute(
                    """
                    INSERT INTO attachments (
                      job_id, digest, kind, original_name, mime_type,
                      relative_path, size_bytes, device_id, created_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(job_id, digest) DO UPDATE SET
                      kind=excluded.kind,
                      original_name=excluded.original_name,
                      mime_type=excluded.mime_type,
                      relative_path=excluded.relative_path,
                      size_bytes=excluded.size_bytes,
                      device_id=excluded.device_id
                    """,
                    (
                        job_id,
                        digest,
                        kind,
                        safe_original,
                        detected_mime,
                        relative_path.as_posix(),
                        len(data),
                        str(device_id or "")[:64],
                        created_at,
                    ),
                )
                connection.commit()
        return {
            "job_id": job_id,
            "sha256": digest,
            "kind": kind,
            "mime_type": detected_mime,
            "size": len(data),
            "stored": True,
        }

    def count_for_job(self, job_id: str) -> int:
        job_id = validate_job_id(job_id)
        with self._lock:
            self._initialize()
            with sqlite3.connect(self.database_path, timeout=10) as connection:
                row = connection.execute(
                    "SELECT COUNT(*) FROM attachments WHERE job_id=?", (job_id,)
                ).fetchone()
        return int(row[0] if row else 0)

    def _initialize(self) -> None:
        self.root.mkdir(parents=True, exist_ok=True)
        with sqlite3.connect(self.database_path, timeout=10) as connection:
            connection.execute("PRAGMA journal_mode=WAL")
            connection.execute(
                """
                CREATE TABLE IF NOT EXISTS attachments (
                  job_id TEXT NOT NULL,
                  digest TEXT NOT NULL,
                  kind TEXT NOT NULL,
                  original_name TEXT NOT NULL,
                  mime_type TEXT NOT NULL,
                  relative_path TEXT NOT NULL,
                  size_bytes INTEGER NOT NULL,
                  device_id TEXT NOT NULL,
                  created_at TEXT NOT NULL,
                  PRIMARY KEY (job_id, digest)
                )
                """
            )
            connection.execute(
                "CREATE INDEX IF NOT EXISTS idx_attachments_job ON attachments(job_id)"
            )
            connection.commit()
