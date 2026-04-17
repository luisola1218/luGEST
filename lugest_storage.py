import os
import re
import shutil
from pathlib import Path
from typing import Any


SHARED_STORAGE_ENV = "LUGEST_SHARED_STORAGE_ROOT"


def _text(value: Any) -> str:
    return str(value or "").strip()


def _expand_path(raw: Any, base_dir: str | Path | None = None) -> Path | None:
    text = _text(raw)
    if not text:
        return None
    expanded = os.path.expandvars(os.path.expanduser(text))
    path = Path(expanded)
    if path.is_absolute():
        return path
    if base_dir:
        return Path(base_dir) / path
    return path


def configured_shared_storage_root(base_dir: str | Path | None = None) -> Path | None:
    raw = _text(os.environ.get(SHARED_STORAGE_ENV, ""))
    if not raw:
        return None
    path = _expand_path(raw, base_dir=base_dir)
    if path is None:
        return None
    return path


def shared_storage_root(base_dir: str | Path | None = None) -> Path:
    configured = configured_shared_storage_root(base_dir=base_dir)
    if configured is not None:
        return configured
    base = Path(base_dir) if base_dir else Path.cwd()
    return base / "generated" / "shared"


def shared_storage_configured() -> bool:
    return bool(_text(os.environ.get(SHARED_STORAGE_ENV, "")))


def shared_storage_label(base_dir: str | Path | None = None) -> str:
    root = shared_storage_root(base_dir=base_dir)
    if shared_storage_configured():
        return str(root)
    return f"{root} (local)"


def _slug_filename(text: Any, fallback: str = "ficheiro") -> str:
    raw = _text(text)
    if not raw:
        raw = fallback
    sanitized = re.sub(r"[^\w.\-]+", "_", raw, flags=re.UNICODE).strip("._")
    return sanitized or fallback


def resolve_file_reference(raw: Any, base_dir: str | Path | None = None) -> Path | None:
    path = _expand_path(raw, base_dir=base_dir)
    if path is None:
        return None
    if path.is_absolute():
        return path
    storage_path = shared_storage_root(base_dir=base_dir) / path
    if storage_path.exists():
        return storage_path
    if base_dir:
        return Path(base_dir) / path
    return path


def allocate_storage_output_path(
    category: str,
    filename: str,
    *,
    base_dir: str | Path | None = None,
) -> Path:
    target_dir = shared_storage_root(base_dir=base_dir) / Path(str(category or "").strip())
    target_dir.mkdir(parents=True, exist_ok=True)
    safe_name = _slug_filename(filename, fallback="documento")
    return (target_dir / safe_name).resolve()


def import_file_to_storage(
    raw: Any,
    category: str,
    *,
    base_dir: str | Path | None = None,
    preferred_name: str = "",
) -> str:
    current = _text(raw)
    if not current:
        return ""
    source = resolve_file_reference(current, base_dir=base_dir)
    if source is None or not source.exists() or not source.is_file():
        return current
    source = source.resolve()
    root = shared_storage_root(base_dir=base_dir).resolve()
    try:
        if source.is_relative_to(root):
            return str(source)
    except Exception:
        pass
    name = _slug_filename(preferred_name or source.name, fallback=source.name or "ficheiro")
    target = allocate_storage_output_path(category, name, base_dir=base_dir)
    if target.exists():
        try:
            if target.samefile(source):
                return str(target)
        except Exception:
            pass
        stem = target.stem
        suffix = target.suffix
        counter = 2
        while target.exists():
            target = target.with_name(f"{stem}_{counter}{suffix}")
            counter += 1
    shutil.copy2(source, target)
    return str(target.resolve())


def file_reference_name(raw: Any, base_dir: str | Path | None = None) -> str:
    path = resolve_file_reference(raw, base_dir=base_dir)
    if path is None:
        return ""
    return path.name
