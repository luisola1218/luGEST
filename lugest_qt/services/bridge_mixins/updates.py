from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import tempfile
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any

from lugest_core.updates import sha256_file as _sha256_file
from lugest_core.updates import validate_update_manifest as _validate_update_manifest
from lugest_infra.config import AtomicJsonStore


class UpdatesBridgeMixin:
    """Desktop update adapter. The host supplies paths and configuration storage."""

    def app_version(self) -> str:
        candidates = [
            self.base_dir / "VERSION",
            Path.cwd() / "VERSION",
        ]
        for path in candidates:
            try:
                if path.exists():
                    value = path.read_text(encoding="utf-8").strip()
                    if value:
                        return value
            except Exception:
                continue
        return "0.0.0"

    def update_settings(self) -> dict[str, Any]:
        cfg = self._load_qt_config()
        stored = dict(cfg.get("update_settings", {}) or {})
        stored.pop("github_token", None)
        manifest_env = str(os.environ.get("LUGEST_UPDATE_MANIFEST_URL", "") or "").strip()
        github_token_env = str(os.environ.get("LUGEST_UPDATE_GITHUB_TOKEN", "") or "").strip()
        defaults = {
            "current_version": self.app_version(),
            "manifest_url": manifest_env or "..\\Atualizacoes\\latest.json",
            "channel": "stable",
            "github_token": github_token_env,
            "auto_check": False,
        }
        return {**defaults, **stored, "current_version": self.app_version()}

    def update_save_settings(self, payload: dict[str, Any]) -> dict[str, Any]:
        cfg = self._load_qt_config()
        current = dict(cfg.get("update_settings", {}) or {})
        current.pop("github_token", None)
        for key in ("manifest_url", "channel", "auto_check"):
            if key in dict(payload or {}):
                current[key] = payload.get(key)
        current["current_version"] = self.app_version()
        cfg["update_settings"] = current
        self._save_qt_config(cfg)
        return self.update_settings()

    def _update_version_parts(self, value: Any) -> tuple[int, int, int, int]:
        parts = [int(match.group(0)) for match in re.finditer(r"\d+", str(value or ""))]
        while len(parts) < 4:
            parts.append(0)
        return tuple(parts[:4])

    def _update_resolve_ref(self, value: Any, base: Path | None = None) -> str:
        txt = str(value or "").strip()
        if not txt:
            return ""
        if re.match(r"^https?://", txt, flags=re.IGNORECASE):
            return txt
        parsed = urllib.parse.urlparse(txt)
        if parsed.scheme.lower() == "file":
            return urllib.request.url2pathname(parsed.path)
        path = Path(txt)
        if path.is_absolute():
            return str(path)
        return str((base or self.base_dir) / path)

    def _update_github_headers(self, token: str = "", *, binary_asset: bool = False) -> dict[str, str]:
        headers: dict[str, str] = {}
        token_txt = str(token or "").strip()
        if token_txt:
            headers["Authorization"] = f"Bearer {token_txt}"
        headers["User-Agent"] = "LuisGEST-Updater"
        headers["Accept"] = "application/octet-stream" if binary_asset else "application/vnd.github+json"
        return headers

    def _update_resolve_github_release_asset_api_url(self, url: str, token: str = "") -> str:
        token_txt = str(token or "").strip()
        if not token_txt:
            return ""
        txt = str(url or "").strip()
        tag_match = re.match(
            r"^https://github\.com/(?P<owner>[^/]+)/(?P<repo>[^/]+)/releases/download/(?P<tag>[^/]+)/(?P<asset>[^/?#]+)$",
            txt,
            flags=re.IGNORECASE,
        )
        latest_match = re.match(
            r"^https://github\.com/(?P<owner>[^/]+)/(?P<repo>[^/]+)/releases/latest/download/(?P<asset>[^/?#]+)$",
            txt,
            flags=re.IGNORECASE,
        )
        match = tag_match or latest_match
        if match is None:
            return ""
        owner = str(match.group("owner") or "").strip()
        repo = str(match.group("repo") or "").strip()
        asset_name = urllib.parse.unquote(str(match.group("asset") or "").strip())
        if not owner or not repo or not asset_name:
            return ""
        api_url = f"https://api.github.com/repos/{owner}/{repo}/releases/latest"
        if tag_match is not None:
            tag = str(tag_match.group("tag") or "").strip()
            if not tag:
                return ""
            api_url = f"https://api.github.com/repos/{owner}/{repo}/releases/tags/{tag}"
        request = urllib.request.Request(api_url, headers=self._update_github_headers(token_txt))
        with urllib.request.urlopen(request, timeout=12) as response:
            payload = json.loads(response.read().decode("utf-8", errors="ignore"))
        if isinstance(payload, dict):
            for asset in list(payload.get("assets", []) or []):
                if str(dict(asset).get("name", "") or "") == asset_name:
                    return str(dict(asset).get("url", "") or "").strip()
        return ""

    def _update_read_json_ref(self, ref: str) -> tuple[dict[str, Any], Path | None]:
        resolved = self._update_resolve_ref(ref)
        if not resolved:
            raise ValueError("Configura o URL/caminho do manifest de atualizacao.")
        if re.match(r"^http://", resolved, flags=re.IGNORECASE):
            raise ValueError("O manifesto remoto tem de usar HTTPS.")
        if re.match(r"^https?://", resolved, flags=re.IGNORECASE):
            settings = self.update_settings()
            headers = {}
            token = str(settings.get("github_token", "") or "").strip()
            request_url = resolved
            if token:
                asset_api_url = self._update_resolve_github_release_asset_api_url(resolved, token)
                if asset_api_url:
                    request_url = asset_api_url
                    headers = self._update_github_headers(token, binary_asset=True)
                else:
                    headers["Authorization"] = f"Bearer {token}"
                    headers["User-Agent"] = "LuisGEST-Updater"
            request = urllib.request.Request(request_url, headers=headers)
            with urllib.request.urlopen(request, timeout=12) as response:
                payload = json.loads(response.read().decode("utf-8-sig", errors="ignore"))
            return (payload if isinstance(payload, dict) else {}, None)
        path = Path(resolved)
        if not path.exists():
            raise ValueError(f"Manifest nao encontrado: {path}")
        payload = json.loads(path.read_text(encoding="utf-8-sig"))
        return (payload if isinstance(payload, dict) else {}, path)

    def _update_resolve_relative_ref(self, ref: str, base_ref: str, manifest_path: Path | None = None) -> str:
        ref_txt = str(ref or "").strip()
        if not ref_txt:
            return ""
        if re.match(r"^https?://", ref_txt, flags=re.IGNORECASE):
            return ref_txt
        if ref_txt.lower().startswith("file:///"):
            return str(Path(urllib.request.url2pathname(urllib.parse.urlparse(ref_txt).path)))
        base_txt = str(base_ref or "").strip()
        if base_txt and re.match(r"^https?://", base_txt, flags=re.IGNORECASE):
            encoded_ref = urllib.parse.quote(ref_txt, safe="/:@?&=%#+,;~-._")
            return urllib.parse.urljoin(base_txt, encoded_ref)
        if manifest_path is not None:
            return str((manifest_path.parent / ref_txt).resolve())
        return self._update_resolve_ref(ref_txt)

    def _update_download_ref_to_temp(self, ref: str, suffix: str = ".tmp") -> Path:
        resolved = self._update_resolve_ref(ref)
        if not resolved:
            raise ValueError("Referencia de atualizacao vazia.")
        if re.match(r"^http://", resolved, flags=re.IGNORECASE):
            raise ValueError("O download remoto tem de usar HTTPS.")
        temp_path = Path(tempfile.mkdtemp(prefix="lugest_update_bootstrap_")) / f"asset{suffix}"
        try:
            if re.match(r"^https?://", resolved, flags=re.IGNORECASE):
                settings = self.update_settings()
                headers = {}
                token = str(settings.get("github_token", "") or "").strip()
                request_url = resolved
                if token:
                    asset_api_url = self._update_resolve_github_release_asset_api_url(resolved, token)
                    if asset_api_url:
                        request_url = asset_api_url
                        headers = self._update_github_headers(token, binary_asset=True)
                    else:
                        headers["Authorization"] = f"Bearer {token}"
                        headers["User-Agent"] = "LuisGEST-Updater"
                request = urllib.request.Request(request_url, headers=headers)
                with urllib.request.urlopen(request, timeout=20) as response, temp_path.open("wb") as handle:
                    shutil.copyfileobj(response, handle)
                return temp_path
            source = Path(resolved)
            if not source.exists():
                raise ValueError(f"Ficheiro de atualizacao nao encontrado: {source}")
            shutil.copy2(source, temp_path)
            return temp_path
        except BaseException:
            shutil.rmtree(temp_path.parent, ignore_errors=True)
            raise

    def _update_download_ref_to_path(self, ref: str, target: Path) -> Path:
        downloaded = self._update_download_ref_to_temp(ref, suffix=target.suffix or ".tmp")
        try:
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(downloaded, target)
        finally:
            shutil.rmtree(downloaded.parent, ignore_errors=True)
        return target

    def update_check(self) -> dict[str, Any]:
        settings = self.update_settings()
        current_version = self.app_version()
        manifest, manifest_path = self._update_read_json_ref(str(settings.get("manifest_url", "") or ""))
        validated = _validate_update_manifest(manifest)
        latest_version = validated.version
        available = self._update_version_parts(latest_version) > self._update_version_parts(current_version)
        return {
            "current_version": current_version,
            "latest_version": latest_version,
            "update_available": available,
            "manifest_url": str(settings.get("manifest_url", "") or ""),
            "manifest_path": str(manifest_path or ""),
            "package_url": validated.package_url,
            "bootstrap_url": validated.bootstrap_url,
            "sha256": validated.package_sha256,
            "bootstrap_sha256": validated.bootstrap_sha256,
            "notes": validated.notes,
            "channel": validated.channel,
        }

    def update_installer_command(self) -> list[str]:
        settings = dict(self.update_settings() or {})
        powershell_exe = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32" / "WindowsPowerShell" / "v1.0" / "powershell.exe"
        manifest_url = str(settings.get("manifest_url", "") or "").strip()
        if not manifest_url:
            raise ValueError("Configura primeiro o manifest de atualizacao.")
        manifest, manifest_path = self._update_read_json_ref(manifest_url)
        validated = _validate_update_manifest(manifest)
        bootstrap_ref = validated.bootstrap_url
        bootstrap_resolved = self._update_resolve_relative_ref(bootstrap_ref, manifest_url, manifest_path)
        local_repair_script = self.base_dir / "Reparar Atualizador Instalado.ps1"
        # Fluxo validado em cliente: renovar primeiro o reparador local e so depois
        # executa-lo. Foi este comportamento que substituiu com sucesso a copia manual
        # via TeamViewer que o utilizador fazia quando o update automatico falhava.
        downloaded_bootstrap = self._update_download_ref_to_temp(bootstrap_resolved, suffix=".ps1")
        try:
            actual_bootstrap_hash = _sha256_file(downloaded_bootstrap)
            if actual_bootstrap_hash.casefold() != validated.bootstrap_sha256.casefold():
                raise ValueError(
                    "Checksum do reparador invÃ¡lido. A atualizaÃ§Ã£o foi interrompida antes de substituir ficheiros."
                )
            staged_repair_script = local_repair_script.with_suffix(".ps1.download")
            try:
                shutil.copy2(downloaded_bootstrap, staged_repair_script)
                os.replace(staged_repair_script, local_repair_script)
            except Exception:
                staged_repair_script.unlink(missing_ok=True)
                raise
        finally:
            shutil.rmtree(downloaded_bootstrap.parent, ignore_errors=True)
        command = [
            str(powershell_exe),
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(local_repair_script),
            "-InstallDir",
            str(self.base_dir),
            "-ManifestUrl",
            manifest_url,
            "-ConfigPath",
            str(self.app_paths.config_file("update_config.json")),
            "-CurrentVersion",
            self.app_version(),
        ]
        token = str(settings.get("github_token", "") or "").strip()
        if token:
            command.extend(["-GitHubToken", token])
        return command

    def _update_sync_installer_config(self) -> Path:
        target = self.app_paths.config_file("update_config.json")
        settings = dict(self.update_settings() or {})
        payload = {
            "current_version": self.app_version(),
            "manifest_url": str(settings.get("manifest_url", "") or "").strip(),
            "channel": str(settings.get("channel", "stable") or "stable").strip() or "stable",
            "auto_check": bool(settings.get("auto_check", False)),
        }
        AtomicJsonStore(target, max_bytes=1024 * 1024).save(payload)
        return target

    def update_start_installer(self) -> dict[str, Any]:
        config_path = self._update_sync_installer_config()
        command = self.update_installer_command()
        creationflags = getattr(subprocess, "CREATE_NEW_CONSOLE", 0)
        subprocess.Popen(command, cwd=str(self.base_dir), close_fds=True, creationflags=creationflags)
        return {"started": True, "command": command, "config_path": str(config_path)}
