from __future__ import annotations

import os
import sys
import tempfile
from pathlib import Path
from types import SimpleNamespace

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_infra.app_paths import AppPaths, MACHINE_DATA_ENV, USER_DATA_ENV
from lugest_infra.config import AtomicJsonStore
from lugest_infra.diagnostics import RuntimeDiagnostics
from lugest_infra.licensing import LicenseStore
from lugest_qt.services.main_bridge import LegacyBackend


def main() -> int:
    previous_user = os.environ.get(USER_DATA_ENV)
    previous_machine = os.environ.get(MACHINE_DATA_ENV)
    try:
        with tempfile.TemporaryDirectory(prefix="lugest-storage-") as temporary:
            root = Path(temporary)
            install_dir = root / "install"
            user_dir = root / "user-data"
            machine_dir = root / "machine-data"
            install_dir.mkdir()
            os.environ[USER_DATA_ENV] = str(user_dir)
            os.environ[MACHINE_DATA_ENV] = str(machine_dir)

            paths = AppPaths(install_dir)
            assert paths.config_file() == (user_dir / "config" / "lugest_qt_config.json").resolve()
            assert paths.license_file() == (machine_dir / "licensing" / "license.lugest").resolve()

            legacy = install_dir / "lugest_qt_config.json"
            legacy.write_text('{"source": "legacy"}\n', encoding="utf-8")
            target = paths.config_file()
            assert paths.migrate_legacy_file(legacy, target)
            assert not paths.migrate_legacy_file(legacy, target)
            assert AtomicJsonStore(target).load() == {"source": "legacy"}

            store = AtomicJsonStore(target)
            store.save({"revision": 1})
            store.save({"revision": 2})
            assert store.load() == {"revision": 2}
            target.write_text("{corrupt", encoding="utf-8")
            assert store.load() == {"revision": 1}

            diagnostics = RuntimeDiagnostics(path=paths.log_file())
            assert diagnostics.log_path() == (user_dir / "logs" / "lugest_runtime.log").resolve()
            assert LicenseStore.default_path() == (machine_dir / "licensing" / "license.lugest").resolve()

            backend = LegacyBackend.__new__(LegacyBackend)
            backend.base_dir = install_dir
            backend.app_paths = paths
            backend.desktop_main = SimpleNamespace()
            backend._qt_config_cache = None
            backend._qt_config_last_error = ""
            backend._product_taxonomy_nodes_cache = None
            saved = backend._save_qt_config({"ui": {"density": "comfortable"}})
            assert saved == {"ui": {"density": "comfortable"}}
            backend._qt_config_cache = None
            assert backend._load_qt_config() == saved
            assert backend._qt_config_last_error == ""
            backend._qt_config_cache = {
                "update_settings": {
                    "manifest_url": "https://updates.example.invalid/latest.json",
                    "github_token": "must-not-be-persisted",
                }
            }
            update_config_path = backend._update_sync_installer_config()
            update_config = AtomicJsonStore(update_config_path).load(default={})
            assert update_config_path.parent == paths.config_dir
            assert "github_token" not in update_config
    finally:
        if previous_user is None:
            os.environ.pop(USER_DATA_ENV, None)
        else:
            os.environ[USER_DATA_ENV] = previous_user
        if previous_machine is None:
            os.environ.pop(MACHINE_DATA_ENV, None)
        else:
            os.environ[MACHINE_DATA_ENV] = previous_machine

    print("app-storage-ok paths=separated migration=preserved writes=atomic recovery=backup bridge=local")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
