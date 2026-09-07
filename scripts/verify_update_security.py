from __future__ import annotations

import copy
import hashlib
import json
import sys
import tempfile
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from lugest_core.updates import UpdateManifestError, sha256_file, validate_update_manifest
from lugest_infra.app_paths import AppPaths
from lugest_qt.services.main_bridge import LegacyBackend


def expect_rejected(payload: dict[str, object], field: str) -> None:
    try:
        validate_update_manifest(payload)
    except UpdateManifestError:
        return
    raise AssertionError(f"Manifesto inseguro aceite ({field}).")


def main() -> int:
    with tempfile.TemporaryDirectory(prefix="lugest-update-security-") as temporary:
        root = Path(temporary)
        package = root / "package.zip"
        bootstrap = root / "repair.ps1"
        package.write_bytes(b"release-package")
        bootstrap.write_bytes(b"Write-Host secure-bootstrap\n")
        payload: dict[str, object] = {
            "schema_version": 1,
            "version": "2026.09.04.1",
            "channel": "stable",
            "package_url": "https://updates.example.invalid/package.zip",
            "sha256": sha256_file(package),
            "bootstrap_url": "https://updates.example.invalid/repair.ps1",
            "bootstrap_sha256": sha256_file(bootstrap),
            "notes": "Teste local sem rede.",
        }
        validated = validate_update_manifest(payload)
        assert validated.package_sha256 == hashlib.sha256(package.read_bytes()).hexdigest()
        assert validated.bootstrap_sha256 == hashlib.sha256(bootstrap.read_bytes()).hexdigest()

        for field in ("sha256", "bootstrap_sha256"):
            candidate = copy.deepcopy(payload)
            candidate.pop(field)
            expect_rejected(candidate, field)
            candidate = copy.deepcopy(payload)
            candidate[field] = "abc123"
            expect_rejected(candidate, field)

        for field in ("package_url", "bootstrap_url"):
            candidate = copy.deepcopy(payload)
            candidate[field] = "http://updates.example.invalid/insecure"
            expect_rejected(candidate, field)

        install_dir = root / "install"
        install_dir.mkdir()
        (install_dir / "VERSION").write_text("2026.09.03.1\n", encoding="utf-8")
        local_payload = dict(payload)
        local_payload["package_url"] = package.name
        local_payload["bootstrap_url"] = bootstrap.name
        manifest = root / "latest.json"
        manifest.write_text(json.dumps(local_payload), encoding="utf-8")
        backend = LegacyBackend.__new__(LegacyBackend)
        backend.base_dir = install_dir
        backend.app_paths = AppPaths(install_dir)
        backend.desktop_main = SimpleNamespace()
        backend._qt_config_cache = {
            "update_settings": {"manifest_url": str(manifest), "channel": "stable"}
        }
        backend._qt_config_last_error = ""
        backend._product_taxonomy_nodes_cache = None
        command = backend.update_installer_command()
        installed_bootstrap = install_dir / "Reparar Atualizador Instalado.ps1"
        assert installed_bootstrap.read_bytes() == bootstrap.read_bytes()
        assert "-ConfigPath" in command
        assert command[command.index("-ManifestUrl") + 1] == str(manifest)

        bootstrap.write_bytes(b"tampered")
        try:
            backend.update_installer_command()
        except ValueError as exc:
            assert "Checksum" in str(exc)
        else:
            raise AssertionError("Bootstrap adulterado aceite pela ponte Qt.")
        assert installed_bootstrap.read_bytes() != bootstrap.read_bytes()

        staging = root / "failed-download"
        staging.mkdir()
        with patch("lugest_qt.services.bridge_mixins.updates.tempfile.mkdtemp", return_value=str(staging)):
            try:
                backend._update_download_ref_to_temp(str(root / "missing.ps1"))
            except ValueError:
                pass
            else:
                raise AssertionError("Missing update source accepted")
        assert not staging.exists(), "Failed downloads must clean their staging directory"
        try:
            backend._update_download_ref_to_temp("http://updates.example.invalid/file")
        except ValueError as exc:
            assert "HTTPS" in str(exc)
        else:
            raise AssertionError("Insecure direct download accepted")

    updater = (ROOT / "scripts" / "lugest_update.ps1").read_text(encoding="utf-8-sig")
    repair = (ROOT / "scripts" / "repair_installed_updater.ps1").read_text(encoding="utf-8-sig")
    mandatory_hash_guard = "if ($expectedHash -notmatch '^[A-Fa-f0-9]{64}$')"
    for name, source in (("updater", updater), ("repair", repair)):
        assert mandatory_hash_guard in source, f"{name}: falta validação SHA-256 obrigatória"
        assert "if ($expectedHash)" not in source, f"{name}: checksum ainda é opcional"
        assert "^http://" in source, f"{name}: falta bloquear HTTP"

    release_builder = (ROOT / "scripts" / "prepare_final_release.ps1").read_text(encoding="utf-8-sig")
    for required_name in (
        "Atualizar LuisGEST.ps1",
        "Atualizar LuisGEST.bat",
        "Reparar Atualizador Instalado.ps1",
        "Reparar Atualizador Instalado.bat",
    ):
        assert required_name in release_builder, f"release comercial não inclui {required_name}"

    print("update-security-ok manifest=strict package_hash=required bootstrap_hash=required transport=https bridge=tamper-blocked")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
