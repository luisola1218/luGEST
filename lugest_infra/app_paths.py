from __future__ import annotations

import os
import shutil
import tempfile
from dataclasses import dataclass
from pathlib import Path


USER_DATA_ENV = "LUGEST_USER_DATA_DIR"
MACHINE_DATA_ENV = "LUGEST_MACHINE_DATA_DIR"


def _environment_path(name: str) -> Path | None:
    value = str(os.environ.get(name, "") or "").strip()
    return Path(os.path.expandvars(os.path.expanduser(value))) if value else None


@dataclass(frozen=True)
class AppPaths:
    """Resolve mutable application data outside the installation directory."""

    install_dir: Path
    application_name: str = "luGEST"

    def __init__(self, install_dir: Path | str, application_name: str = "luGEST") -> None:
        object.__setattr__(self, "install_dir", Path(install_dir).resolve())
        object.__setattr__(self, "application_name", str(application_name or "luGEST").strip() or "luGEST")

    @property
    def user_data_dir(self) -> Path:
        configured = _environment_path(USER_DATA_ENV)
        if configured is not None:
            return configured.resolve()
        local_app_data = str(os.environ.get("LOCALAPPDATA", "") or "").strip()
        base = Path(local_app_data) if local_app_data else Path.home() / "AppData" / "Local"
        # Keep mutable files outside both the legacy %LOCALAPPDATA%\luGEST
        # installation and the new %LOCALAPPDATA%\Programs\luGEST location.
        return (base / f"{self.application_name}-data").resolve()

    @property
    def machine_data_dir(self) -> Path:
        configured = _environment_path(MACHINE_DATA_ENV)
        if configured is not None:
            return configured.resolve()
        program_data = str(os.environ.get("PROGRAMDATA", "") or "").strip()
        if program_data:
            return (Path(program_data) / self.application_name).resolve()
        return self.user_data_dir

    @property
    def config_dir(self) -> Path:
        return self.user_data_dir / "config"

    @property
    def logs_dir(self) -> Path:
        return self.user_data_dir / "logs"

    @property
    def cache_dir(self) -> Path:
        return self.user_data_dir / "cache"

    @property
    def license_dir(self) -> Path:
        return self.machine_data_dir / "licensing"

    def config_file(self, filename: str = "lugest_qt_config.json") -> Path:
        return self.config_dir / Path(filename).name

    def log_file(self, filename: str = "lugest_runtime.log") -> Path:
        return self.logs_dir / Path(filename).name

    def license_file(self, filename: str = "license.lugest") -> Path:
        return self.license_dir / Path(filename).name

    def migrate_legacy_file(self, source: Path | str, target: Path | str) -> bool:
        source_path = Path(source)
        target_path = Path(target)
        if target_path.exists() or not source_path.is_file():
            return False
        try:
            if source_path.resolve() == target_path.resolve():
                return False
        except OSError:
            pass
        target_path.parent.mkdir(parents=True, exist_ok=True)
        descriptor, temporary_name = tempfile.mkstemp(
            prefix=f".{target_path.name}.",
            suffix=".migrate",
            dir=str(target_path.parent),
        )
        os.close(descriptor)
        temporary_path = Path(temporary_name)
        try:
            shutil.copy2(source_path, temporary_path)
            os.replace(temporary_path, target_path)
        except Exception:
            temporary_path.unlink(missing_ok=True)
            raise
        return True


__all__ = ["AppPaths", "MACHINE_DATA_ENV", "USER_DATA_ENV"]
