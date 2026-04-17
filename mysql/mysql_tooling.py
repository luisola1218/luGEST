from __future__ import annotations

import os
import shutil
from pathlib import Path


def load_env_candidates(paths: list[Path]) -> None:
    for path in paths:
        if not path.exists():
            continue
        for raw in path.read_text(encoding="utf-8").splitlines():
            line = str(raw or "").strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            clean_key = key.strip().lstrip("\ufeff")
            clean_value = value.strip().strip('"').strip("'")
            if clean_key and clean_key not in os.environ:
                os.environ[clean_key] = clean_value


def env_int(name: str, default: int) -> int:
    raw = str(os.environ.get(name, default) or "").strip()
    try:
        return int(raw)
    except Exception:
        return int(default)


def resolve_mysql_binary(
    explicit_path: str | None,
    env_var_name: str,
    executable_names: list[str],
) -> str | None:
    explicit = str(explicit_path or "").strip()
    if explicit:
        candidate = Path(explicit).expanduser()
        if candidate.exists():
            return str(candidate.resolve())

    env_value = str(os.environ.get(env_var_name, "") or "").strip()
    if env_value:
        candidate = Path(env_value).expanduser()
        if candidate.exists():
            return str(candidate.resolve())

    for executable_name in executable_names:
        found = shutil.which(executable_name)
        if found:
            return str(Path(found).resolve())

    roots = [
        Path(os.environ.get("ProgramFiles", "")),
        Path(os.environ.get("ProgramFiles(x86)", "")),
        Path(os.environ.get("ProgramW6432", "")),
        Path("C:/xampp/mysql/bin"),
        Path("C:/mariadb/bin"),
    ]
    patterns = (
        "MySQL/*/bin",
        "MySQL Server */bin",
        "MariaDB */bin",
        "MariaDB/*/bin",
    )
    for root in roots:
        if not root or not root.exists():
            continue
        if root.is_dir() and root.name.lower() == "bin":
            for executable_name in executable_names:
                candidate = root / executable_name
                if candidate.exists():
                    return str(candidate.resolve())
        for pattern in patterns:
            for bin_dir in root.glob(pattern):
                for executable_name in executable_names:
                    candidate = bin_dir / executable_name
                    if candidate.exists():
                        return str(candidate.resolve())
    return None


def build_mysql_env(password: str) -> dict[str, str]:
    env = dict(os.environ)
    if password:
        env["MYSQL_PWD"] = password
    return env
