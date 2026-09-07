from __future__ import annotations

import time
from typing import Any


class UsersBackendMixin:
    """Legacy adapter for users; see BACKEND_GUIDE.md."""

    def _user_profiles(self) -> dict[str, Any]:
        cfg = self._load_qt_config()
        return dict(cfg.get("user_profiles", {}) or {})

    def _save_user_profiles(self, profiles: dict[str, Any]) -> dict[str, Any]:
        cfg = self._load_qt_config()
        cfg["user_profiles"] = dict(profiles or {})
        self._save_qt_config(cfg)
        return dict(cfg.get("user_profiles", {}) or {})

    def _user_profile(self, username: str) -> dict[str, Any]:
        return dict(self._user_profiles().get(str(username or "").strip().lower(), {}) or {})

    def _authenticate_error_message(self, data: dict[str, Any], username: str, password: str) -> str:
        username_txt = str(username or "").strip()
        if not username_txt:
            return "Indica um utilizador."

        owner_username = str(self.desktop_main.trial_owner_username() or "").strip()
        local_user = self.desktop_main.find_local_user(data, username_txt)
        local_users = [
            row
            for row in list(data.get("users", []) or [])
            if isinstance(row, dict) and str(row.get("username", "") or "").strip()
        ]

        if owner_username and username_txt.lower() == owner_username.lower():
            if not str(password or "").strip():
                return (
                    f"O login '{owner_username}' e reservado ao proprietario/licenciamento. "
                    "Indica a password real do owner ou entra com um utilizador local da aplicacao."
                )
            return (
                f"O login '{owner_username}' e reservado ao proprietario/licenciamento e esta password nao foi aceite. "
                "Nao uses o hash do lugest.env como password; entra com um utilizador local da aplicacao "
                "ou repoe o administrador local."
            )

        if not local_users:
            return (
                "Nao existem utilizadores locais configurados. "
                "Cria ou repoe o administrador inicial antes de entrar."
            )

        if not isinstance(local_user, dict):
            return "Utilizador nao encontrado."

        return "Password incorreta."

    def authenticate(self, username: str, password: str) -> dict[str, Any]:
        data = self.reload(force=False)
        owner_session = self.desktop_main.ensure_trial_login_session(username, password, allow_owner=True)
        if isinstance(owner_session, dict):
            merged = dict(owner_session)
            self.user = merged
            return self.user
        user = self.desktop_main.authenticate_local_user(data, username, password)
        if user is None:
            raise ValueError(self._authenticate_error_message(data, username, password))
        self.desktop_main.ensure_trial_login_session(username, password, allow_owner=False)
        merged = self.desktop_main.build_authenticated_user_session(user, password)
        profile = self._user_profile(str(merged.get("username", "") or ""))
        if profile and not bool(profile.get("active", True)):
            raise ValueError("Utilizador desativado.")
        for key in ("posto", "posto_trabalho", "work_center"):
            if str(profile.get("posto", "") or "").strip():
                merged[key] = str(profile.get("posto", "") or "").strip()
        merged["active"] = bool(profile.get("active", True))
        merged["menu_permissions"] = dict(profile.get("menu_permissions", {}) or {})
        self.user = merged
        try:
            self.desktop_main.touch_trial_success(str(merged.get("username", "") or "").strip(), owner=False)
        except Exception:
            pass
        return self.user

    def trial_status(self, *, force: bool = False) -> dict[str, Any]:
        if (
            not force
            and isinstance(self._trial_status_cache, dict)
            and self._trial_status_loaded_at > 0
            and (time.monotonic() - self._trial_status_loaded_at) <= self._trial_status_cache_ttl_sec
        ):
            return dict(self._trial_status_cache)
        try:
            payload = dict(self.desktop_main.get_trial_status(force_time=bool(force)) or {})
        except TypeError:
            payload = dict(self.desktop_main.get_trial_status() or {})
        payload["management_allowed"] = self.is_owner_session()
        self._trial_status_cache = dict(payload)
        self._trial_status_loaded_at = time.monotonic()
        return payload

    def is_owner_session(self) -> bool:
        return bool(dict(self.user or {}).get("owner_session", False))

    def _ensure_trial_management_access(self) -> None:
        if self.is_owner_session():
            return
        raise ValueError("A gestao de trial/licenca exige login pelo utilizador OWNER.")

    def activate_trial_license(self, company_name: str = "", duration_days: int = 60, notes: str = "") -> dict[str, Any]:
        self._ensure_trial_management_access()
        actor = str((self.user or {}).get("username", "") or "").strip()
        payload = dict(
            self.desktop_main.activate_trial_license(
                company_name=company_name,
                duration_days=duration_days,
                created_by=actor,
                notes=notes,
                reset_start=True,
            )
            or {}
        )
        payload["management_allowed"] = True
        self._trial_status_cache = dict(payload)
        self._trial_status_loaded_at = time.monotonic()
        return payload

    def extend_trial_license(self, extra_days: int = 30) -> dict[str, Any]:
        self._ensure_trial_management_access()
        actor = str((self.user or {}).get("username", "") or "").strip()
        payload = dict(self.desktop_main.extend_trial_license(extra_days=extra_days, updated_by=actor) or {})
        payload["management_allowed"] = True
        self._trial_status_cache = dict(payload)
        self._trial_status_loaded_at = time.monotonic()
        return payload

    def disable_trial_license(self) -> dict[str, Any]:
        self._ensure_trial_management_access()
        actor = str((self.user or {}).get("username", "") or "").strip()
        payload = dict(self.desktop_main.disable_trial_license(updated_by=actor) or {})
        payload["management_allowed"] = True
        self._trial_status_cache = dict(payload)
        self._trial_status_loaded_at = time.monotonic()
        return payload

    def verify_supervisor_password(self, password: str) -> bool:
        stored = str(dict(self._load_qt_config().get("ui_options", {}) or {}).get("operator_supervisor_password", "") or "").strip()
        if not stored:
            return False
        return bool(self.desktop_main.verify_password(str(password or "").strip(), stored))

    def available_menu_pages(self) -> list[dict[str, str]]:
        return [
            {"key": "stock_dashboard", "label": "Dashboard"},
            {"key": "materials", "label": "Matéria-Prima"},
            {"key": "products", "label": "Produtos"},
            {"key": "direct_services", "label": "Serviços"},
            {"key": "clients", "label": "Clientes"},
            {"key": "suppliers", "label": "Fornecedores"},
            {"key": "orders", "label": "Encomendas"},
            {"key": "quotes", "label": "Orçamentos"},
            {"key": "planning", "label": "Planeamento"},
            {"key": "transportes", "label": "Transportes"},
            {"key": "material_assistant", "label": "Assistente MP"},
            {"key": "operator", "label": "Operador"},
            {"key": "opp", "label": "OPP"},
            {"key": "shipping", "label": "Expedição"},
            {"key": "billing", "label": "Faturação"},
            {"key": "purchase_notes", "label": "Notas Encomenda"},
            {"key": "quality", "label": "Qualidade"},
            {"key": "diagnostics", "label": "Diagnóstico"},
            {"key": "pulse", "label": "Pulse"},
            {"key": "avarias", "label": "Avarias"},
            {"key": "home", "label": "Resumo"},
        ]

    def available_roles(self) -> list[str]:
        return ["Admin", "Producao", "Qualidade", "Planeamento", "Orcamentista", "Operador"]

    def user_rows(self) -> list[dict[str, Any]]:
        profiles = self._user_profiles()
        rows: list[dict[str, Any]] = []
        for user in list(self.ensure_data().get("users", []) or []):
            username = str(user.get("username", "") or "").strip()
            if not username:
                continue
            profile = dict(profiles.get(username.lower(), {}) or {})
            stored_password = str(user.get("password", "") or "").strip()
            rows.append(
                {
                    "username": username,
                    "password": "",
                    "password_set": bool(stored_password),
                    "role": str(user.get("role", "") or "").strip() or "Operador",
                    "posto": str(profile.get("posto", "") or user.get("posto", "") or "").strip(),
                    "active": bool(profile.get("active", True)),
                    "menu_permissions": dict(profile.get("menu_permissions", {}) or {}),
                }
            )
        rows.sort(key=lambda row: (str(row.get("role", "") or ""), str(row.get("username", "") or "").lower()))
        return rows

    def allowed_pages_for_user(self, user: dict[str, Any] | None = None) -> list[str]:
        current = dict(user or self.user or {})
        if str(current.get("role", "") or "").strip().lower() == "admin":
            return [row["key"] for row in self.available_menu_pages()]
        perms = dict(current.get("menu_permissions", {}) or {})
        if not perms:
            profile = self._user_profile(str(current.get("username", "") or ""))
            perms = dict(profile.get("menu_permissions", {}) or {})
        if not perms:
            return [row["key"] for row in self.available_menu_pages()]
        return [row["key"] for row in self.available_menu_pages() if bool(perms.get(row["key"], False))]

    def save_user(self, payload: dict[str, Any], current_username: str = "") -> dict[str, Any]:
        data = self.ensure_data()
        username = str(payload.get("username", "") or "").strip()
        password = str(payload.get("password", "") or "")
        role = str(payload.get("role", "") or "Operador").strip() or "Operador"
        posto = str(payload.get("posto", "") or "").strip()
        active = bool(payload.get("active", True))
        permissions = {str(key): bool(value) for key, value in dict(payload.get("menu_permissions", {}) or {}).items()}
        if not username:
            raise ValueError("Utilizador obrigatorio.")
        current_txt = str(current_username or "").strip()
        current_key = current_txt.lower()
        logged_key = str((self.user or {}).get("username", "") or "").strip().lower()
        target = None
        for row in list(data.get("users", []) or []):
            row_username = str(row.get("username", "") or "").strip()
            if row_username.lower() == current_key and current_key:
                target = row
                break
        if target is None:
            for row in list(data.get("users", []) or []):
                if str(row.get("username", "") or "").strip().lower() == username.lower():
                    target = row
                    break
        if target is None and any(str(row.get("username", "") or "").strip().lower() == username.lower() for row in list(data.get("users", []) or [])):
            raise ValueError("Ja existe um utilizador com esse username.")
        if target is None and not password.strip():
            raise ValueError("Password obrigatoria para um novo utilizador.")
        stored_password = ""
        if password.strip():
            self.desktop_main.validate_local_password(username, password)
            stored_password = self.desktop_main.normalize_password_for_storage(username, password, require_strong=True)
        if target is not None:
            old_username = str(target.get("username", "") or "").strip()
            if old_username.lower() != username.lower() and any(str(row.get("username", "") or "").strip().lower() == username.lower() for row in list(data.get("users", []) or [])):
                raise ValueError("Ja existe um utilizador com esse username.")
            if not stored_password:
                stored_password = str(target.get("password", "") or "").strip()
            target.update({"username": username, "password": stored_password, "role": role})
        else:
            data.setdefault("users", []).append({"username": username, "password": stored_password, "role": role})
        if logged_key and logged_key in {current_key, username.lower()} and not active:
            raise ValueError("Nao podes desativar o utilizador autenticado.")
        operadores = {str(v).strip() for v in list(data.get("operadores", []) or []) if str(v).strip()}
        operadores.discard(current_txt)
        operadores.discard(username)
        if role.lower() == "operador":
            operadores.add(username)
        data["operadores"] = sorted(operadores)
        orcamentistas = {str(v).strip() for v in list(data.get("orcamentistas", []) or []) if str(v).strip()}
        orcamentistas.discard(current_txt)
        orcamentistas.discard(username)
        if role.lower() in {"orcamentista", "orçamentista"}:
            orcamentistas.add(username)
        data["orcamentistas"] = sorted(orcamentistas)
        profiles = self._user_profiles()
        if current_key and current_key != username.lower() and current_key in profiles:
            profiles.pop(current_key, None)
        profiles[username.lower()] = {
            "posto": posto,
            "active": active,
            "menu_permissions": permissions,
        }
        self._save_user_profiles(profiles)
        self._save(force=True)
        if logged_key and logged_key in {current_key, username.lower()}:
            session_password = str(password or (self.user or {}).get("_session_password", "") or "").strip()
            if session_password:
                self.user = self.authenticate(username, session_password)
        return next((row for row in self.user_rows() if str(row.get("username", "") or "").strip().lower() == username.lower()), {})

    def remove_user(self, username: str) -> None:
        username_txt = str(username or "").strip()
        if not username_txt:
            raise ValueError("Utilizador inválido.")
        if self.user and str(self.user.get("username", "") or "").strip().lower() == username_txt.lower():
            raise ValueError("Nao e permitido remover o utilizador atualmente autenticado.")
        data = self.ensure_data()
        before = len(list(data.get("users", []) or []))
        data["users"] = [row for row in list(data.get("users", []) or []) if str(row.get("username", "") or "").strip().lower() != username_txt.lower()]
        if len(list(data.get("users", []) or [])) == before:
            raise ValueError("Utilizador não encontrado.")
        profiles = self._user_profiles()
        profiles.pop(username_txt.lower(), None)
        self._save_user_profiles(profiles)
        self._save(force=True)
