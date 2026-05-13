from __future__ import annotations

import os
import re
import sys
import tempfile
from datetime import date, datetime, timedelta
from dataclasses import dataclass
from pathlib import Path
from types import SimpleNamespace
from typing import Any


def _env_int(name: str, default: int) -> int:
    try:
        return int(str(os.environ.get(name, default) or default).strip())
    except Exception:
        return int(default)


@dataclass(frozen=True)
class _RuntimeSettings:
    db_host: str = os.environ.get("LUGEST_DB_HOST", "127.0.0.1")
    db_port: int = _env_int("LUGEST_DB_PORT", 3306)
    db_user: str = os.environ.get("LUGEST_DB_USER", "")
    db_pass: str = os.environ.get("LUGEST_DB_PASS", "")
    db_name: str = os.environ.get("LUGEST_DB_NAME", "lugest")


settings = _RuntimeSettings()


def _resolve_desktop_root() -> Path:
    explicit = str(os.environ.get("LUGEST_DESKTOP_ROOT", "") or "").strip()
    if explicit:
        candidate = Path(explicit).expanduser().resolve()
        if (candidate / "main.py").exists():
            return candidate
    fallback = Path(__file__).resolve().parents[2]
    if (fallback / "main.py").exists():
        return fallback
    return fallback


REPO_ROOT = _resolve_desktop_root()
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

os.environ.setdefault("LUGEST_DB_HOST", settings.db_host)
os.environ.setdefault("LUGEST_DB_PORT", str(settings.db_port))
os.environ.setdefault("LUGEST_DB_USER", settings.db_user)
os.environ.setdefault("LUGEST_DB_PASS", settings.db_pass)
os.environ.setdefault("LUGEST_DB_NAME", settings.db_name)

import main as desktop_main  # noqa: E402
from lugest_desktop.legacy import menu_rooting as desktop_pulse  # noqa: E402
from lugest_desktop.legacy import plan_actions as desktop_plan  # noqa: E402
from lugest_qt.services.main_bridge import LegacyBackend  # noqa: E402

PERIOD_MAP = {"hoje": 1, "1": 1, "7 dias": 7, "7": 7, "30 dias": 30, "30": 30, "tudo": 0, "0": 0}
_MATERIAL_BACKEND: LegacyBackend | None = None


def _period_to_days(raw: str | int | None) -> int:
    return PERIOD_MAP.get(str(raw or "7 dias").strip().lower(), 7)


def _material_backend() -> LegacyBackend:
    global _MATERIAL_BACKEND
    if _MATERIAL_BACKEND is None:
        _MATERIAL_BACKEND = LegacyBackend()
    return _MATERIAL_BACKEND


def _make_context() -> SimpleNamespace:
    data = desktop_main.load_data()
    ctx = SimpleNamespace()
    ctx.data = data

    def _refresh_runtime_impulse_data(cleanup_orphans: bool = True) -> None:
        desktop_main.mysql_refresh_runtime_impulse_data(ctx.data, cleanup_orphans=cleanup_orphans)

    ctx.refresh_runtime_impulse_data = _refresh_runtime_impulse_data
    return ctx


def _refresh_ctx_runtime(ctx: SimpleNamespace) -> None:
    try:
        if hasattr(ctx, "refresh_runtime_impulse_data"):
            ctx.refresh_runtime_impulse_data(cleanup_orphans=True)
    except Exception:
        pass


def _match_year(enc_obj: dict[str, Any], year_filter: str | None) -> bool:
    if not year_filter:
        return True
    try:
        d_ref = desktop_pulse._safe_date(enc_obj.get("data_criacao") or enc_obj.get("inicio_producao") or enc_obj.get("data_entrega"))
        if d_ref is not None:
            return str(d_ref.year) == str(year_filter)
    except Exception:
        pass
    numero = str(enc_obj.get("numero", "") or "").strip()
    parts = numero.split("-")
    if len(parts) >= 2 and parts[1].isdigit():
        return parts[1] == str(year_filter)
    return True


def _build_history_snapshot(ctx: SimpleNamespace, metrics: dict[str, Any], enc_filter: str, view_mode: str, origin_mode: str, year_filter: str | None) -> list[dict[str, Any]]:
    view_hist = list(metrics.get("historico_tempo", []) or [])
    origem_mode_norm = desktop_pulse._norm_text(origin_mode)
    if enc_filter and enc_filter.lower() != "todas":
        view_hist = [r for r in view_hist if str(r.get("encomenda", "") or "").strip() == enc_filter]
    if desktop_pulse._norm_text(view_mode).startswith("so"):
        view_hist = [r for r in view_hist if bool(r.get("fora"))]
    if "curso" in origem_mode_norm:
        view_hist = []

    snapshot_hist: dict[str, dict[str, Any]] = {}
    for enc_obj in list(ctx.data.get("encomendas", []) or []):
        enc_num = str(enc_obj.get("numero", "") or "").strip()
        if not enc_num:
            continue
        if enc_filter and enc_filter.lower() != "todas" and enc_num != enc_filter:
            continue
        if not _match_year(enc_obj, year_filter):
            continue
        snap_elapsed = 0.0
        snap_plan = 0.0
        snap_ops = 0
        for mat in list(enc_obj.get("materiais", []) or []):
            for esp in list(mat.get("espessuras", []) or []):
                plan_laser = max(0.0, desktop_pulse._parse_float(esp.get("tempo_min", 0), 0))
                if plan_laser <= 0:
                    continue
                laser_elapsed = 0.0
                laser_has_work = False
                for p in list(esp.get("pecas", []) or []):
                    elapsed_piece = max(0.0, desktop_pulse._parse_float(p.get("tempo_producao_min", 0), 0))
                    if elapsed_piece <= 0 and not list(p.get("hist", []) or []):
                        continue
                    laser_has_work = True
                    laser_elapsed = max(laser_elapsed, elapsed_piece)
                if laser_has_work:
                    snap_elapsed += laser_elapsed
                    snap_plan += plan_laser
                    snap_ops += 1
        if snap_elapsed <= 0 and snap_plan <= 0:
            continue
        delta_snap = snap_elapsed - snap_plan if snap_plan > 0 else 0.0
        snapshot_hist[enc_num] = {
            "encomenda": enc_num,
            "ops": snap_ops,
            "elapsed_min": round(snap_elapsed, 1),
            "plan_min": round(snap_plan, 1),
            "delta_min": round(delta_snap, 1),
            "fora": bool(snap_plan > 0 and delta_snap > 0.01),
        }
    if "curso" in origem_mode_norm:
        snapshot_hist = {}
    elif desktop_pulse._norm_text(view_mode).startswith("so"):
        snapshot_hist = {k: v for k, v in list(snapshot_hist.items()) if bool(v.get("fora"))}

    hist_enc: dict[str, dict[str, Any]] = {}
    for row in view_hist:
        enc = str(row.get("encomenda", "") or "").strip() or "-"
        g = hist_enc.setdefault(enc, {"encomenda": enc, "ops": 0, "elapsed_min": 0.0, "plan_min": 0.0, "delta_min": 0.0, "fora": False})
        g["ops"] += 1
        g["elapsed_min"] += desktop_pulse._parse_float(row.get("elapsed_min", 0), 0)
        g["plan_min"] += desktop_pulse._parse_float(row.get("plan_min", 0), 0)
        g["delta_min"] += desktop_pulse._parse_float(row.get("delta_min", 0), 0)
        if bool(row.get("fora")):
            g["fora"] = True
    for enc, snap in list(snapshot_hist.items()):
        g = hist_enc.setdefault(enc, {"encomenda": enc, "ops": 0, "elapsed_min": 0.0, "plan_min": 0.0, "delta_min": 0.0, "fora": False})
        snap_elapsed = max(0.0, desktop_pulse._parse_float(snap.get("elapsed_min", 0), 0))
        snap_plan = max(0.0, desktop_pulse._parse_float(snap.get("plan_min", 0), 0))
        curr_elapsed = max(0.0, desktop_pulse._parse_float(g.get("elapsed_min", 0), 0))
        curr_plan = max(0.0, desktop_pulse._parse_float(g.get("plan_min", 0), 0))
        if snap_elapsed > curr_elapsed + 0.01 or (abs(snap_elapsed - curr_elapsed) <= 0.01 and snap_plan > curr_plan):
            g["elapsed_min"] = snap_elapsed
            g["plan_min"] = snap_plan
            g["ops"] = max(int(g.get("ops", 0) or 0), int(snap.get("ops", 0) or 0))
            g["delta_min"] = snap_elapsed - snap_plan if snap_plan > 0 else 0.0
            g["fora"] = bool(snap_plan > 0 and g["delta_min"] > 0.01)
    for enc, extra_stop in dict(metrics.get("paragens_encomenda", {}) or {}).items():
        g = hist_enc.get(enc)
        if not g:
            continue
        extra_stop_num = max(0.0, desktop_pulse._parse_float(extra_stop, 0))
        if extra_stop_num <= 0:
            continue
        g["elapsed_min"] += extra_stop_num
        if desktop_pulse._parse_float(g.get("plan_min", 0), 0) > 0:
            g["delta_min"] = g["elapsed_min"] - desktop_pulse._parse_float(g.get("plan_min", 0), 0)
            g["fora"] = bool(g["delta_min"] > 0.01)

    rows = list(hist_enc.values())
    rows.sort(key=lambda r: (1 if bool(r.get("fora")) else 0, desktop_pulse._parse_float(r.get("delta_min", 0), 0), desktop_pulse._parse_float(r.get("elapsed_min", 0), 0)), reverse=True)
    if "histor" in origem_mode_norm:
        return rows
    if "curso" in origem_mode_norm:
        return []
    return rows


def get_users() -> list[dict[str, str]]:
    data = desktop_main.load_data()
    out = []
    for user in list(data.get("users", []) or []):
        if not isinstance(user, dict):
            continue
        out.append({
            "username": str(user.get("username", "") or "").strip(),
            "role": str(user.get("role", "") or "").strip(),
        })
    return out


def trial_status() -> dict[str, Any]:
    getter = getattr(desktop_main, "get_trial_status", None)
    if callable(getter):
        try:
            return dict(getter() or {})
        except Exception:
            return {}
    return {}


def mysql_connection_status() -> dict[str, Any]:
    errors_fn = getattr(desktop_main, "_mysql_runtime_errors", None)
    if callable(errors_fn):
        try:
            issues = list(errors_fn() or [])
        except Exception:
            issues = []
        if issues:
            return {
                "ok": False,
                "message": "; ".join(str(item or "").strip() for item in issues if str(item or "").strip()),
                "host": settings.db_host,
                "port": settings.db_port,
                "database": settings.db_name,
            }
    connect = getattr(desktop_main, "_mysql_connect", None)
    if not callable(connect):
        return {
            "ok": False,
            "message": "Ligacao MySQL indisponivel no runtime.",
            "host": settings.db_host,
            "port": settings.db_port,
            "database": settings.db_name,
        }
    conn = None
    try:
        conn = connect()
        with conn.cursor() as cur:
            cur.execute("SELECT 1 AS ok")
            cur.fetchone()
        return {
            "ok": True,
            "message": "Ligacao MySQL operacional.",
            "host": settings.db_host,
            "port": settings.db_port,
            "database": settings.db_name,
        }
    except Exception as exc:
        return {
            "ok": False,
            "message": str(exc),
            "host": settings.db_host,
            "port": settings.db_port,
            "database": settings.db_name,
        }
    finally:
        if conn is not None:
            try:
                conn.close()
            except Exception:
                pass


def _is_trial_owner(username: str) -> bool:
    getter = getattr(desktop_main, "trial_owner_username", None)
    owner = str(getter() if callable(getter) else "" or "").strip().lower()
    current = str(username or "").strip().lower()
    return bool(owner and current and owner == current)


def ensure_trial_access(username: str = "") -> dict[str, Any]:
    status = trial_status()
    if bool(status.get("blocking", False)) and not _is_trial_owner(username):
        raise RuntimeError(str(status.get("message", "") or "Licenca bloqueada."))
    return status


def authenticate(username: str, password: str) -> dict[str, str] | None:
    u_txt = str(username or "").strip().lower()
    p_txt = str(password or "")
    if not u_txt or not p_txt:
        return None
    data = desktop_main.load_data()
    owner_session = desktop_main.ensure_trial_login_session(username, password, allow_owner=True)
    if isinstance(owner_session, dict):
        return {
            "username": str(owner_session.get("username", "") or "").strip(),
            "role": str(owner_session.get("role", "") or "").strip() or "Admin",
        }
    user = desktop_main.authenticate_local_user(data, username, password)
    if not isinstance(user, dict):
        return None
    desktop_main.ensure_trial_login_session(username, password, allow_owner=False)
    try:
        desktop_main.touch_trial_success(str(user.get("username", "") or "").strip(), owner=False)
    except Exception:
        pass
    return {
        "username": str(user.get("username", "") or "").strip(),
        "role": str(user.get("role", "") or "").strip() or "Operador",
    }


def _piece_operation_snapshot(piece: dict[str, Any]) -> dict[str, Any]:
    state = str(piece.get("estado", "") or "").strip()
    state_norm = desktop_pulse._norm_text(state)
    fluxo = list(piece.get("operacoes_fluxo", []) or [])
    op_current = desktop_main.normalize_operacao_nome(piece.get("operacao_atual", "")) or ""
    if op_current and op_current != "-":
        current = op_current
    else:
        current = ""
        for op in fluxo:
            op_state = desktop_pulse._norm_text(op.get("estado", ""))
            if "curso" in op_state or "produc" in op_state or "avari" in op_state:
                current = desktop_main.normalize_operacao_nome(op.get("nome", "")) or ""
                break
    pending = ""
    for op in fluxo:
        op_state = desktop_pulse._norm_text(op.get("estado", ""))
        if "pend" in op_state or "prepar" in op_state:
            pending = desktop_main.normalize_operacao_nome(op.get("nome", "")) or ""
            break
    last_done = ""
    for op in fluxo:
        op_state = desktop_pulse._norm_text(op.get("estado", ""))
        if "concl" in op_state:
            last_done = desktop_main.normalize_operacao_nome(op.get("nome", "")) or last_done

    avaria_ativa = bool(piece.get("avaria_ativa"))
    has_running_flow = False
    has_paused_flow = False
    for op in fluxo:
        op_state = desktop_pulse._norm_text(op.get("estado", ""))
        if "curso" in op_state or "produc" in op_state:
            has_running_flow = True
        if "paus" in op_state or "interromp" in op_state:
            has_paused_flow = True

    if avaria_ativa or "avari" in state_norm:
        effective_state = "Avaria"
        location = current or pending or last_done
        resumo = f"Avaria {location}".strip()
    elif has_running_flow or "curso" in state_norm or "produc" in state_norm:
        effective_state = "Em curso"
        location = current or pending or last_done
        resumo = f"Em curso {location}".strip()
    elif has_paused_flow or "paus" in state_norm or "interromp" in state_norm:
        effective_state = "Pausa"
        location = current or pending or last_done
        resumo = f"Pausa {location}".strip()
    elif "concl" in state_norm:
        effective_state = "Concluida"
        location = last_done or current or pending
        resumo = f"Concluida {location}".strip()
    else:
        effective_state = "Preparacao"
        location = pending or current or last_done
        resumo = f"Preparacao {location}".strip()
    return {"state": state, "effective_state": effective_state, "location": location or "-", "summary": resumo.strip()}


def _order_production_snapshot(enc: dict[str, Any]) -> dict[str, Any]:
    running = 0
    paused = 0
    avaria = 0
    done = 0
    current_locations: list[str] = []
    paused_locations: list[str] = []
    avaria_locations: list[str] = []
    for mat in list(enc.get("materiais", []) or []):
        for esp in list(mat.get("espessuras", []) or []):
            for p in list(esp.get("pecas", []) or []):
                snap = _piece_operation_snapshot(p)
                state_norm = desktop_pulse._norm_text(snap.get("effective_state", snap.get("state", "")))
                location = str(snap["location"] or "-").strip()
                if "avari" in state_norm:
                    avaria += 1
                    if location and location != "-":
                        avaria_locations.append(location)
                elif "paus" in state_norm or "interromp" in state_norm:
                    paused += 1
                    if location and location != "-":
                        paused_locations.append(location)
                elif "curso" in state_norm or "produc" in state_norm:
                    running += 1
                    if location and location != "-":
                        current_locations.append(location)
                elif "concl" in state_norm:
                    done += 1
    unique_current = " | ".join(dict.fromkeys(current_locations))
    unique_paused = " | ".join(dict.fromkeys(paused_locations))
    unique_avaria = " | ".join(dict.fromkeys(avaria_locations))
    if avaria:
        resumo = f"Avaria{f' | {unique_avaria}' if unique_avaria else ''}"
    elif running:
        resumo = f"Em curso{f' | {unique_current}' if unique_current else ''}"
    elif paused:
        resumo = f"Em pausa{f' | {unique_paused}' if unique_paused else ''}"
    elif done and not (running or paused or avaria):
        resumo = "Concluida"
    else:
        resumo = "Por iniciar"
    return {
        "running": running,
        "paused": paused,
        "avaria": avaria,
        "done": done,
        "current_ops": unique_current,
        "paused_ops": unique_paused,
        "avaria_ops": unique_avaria,
        "resumo": resumo,
    }


def _compose_production_resumo(prod: dict[str, Any]) -> str:
    avaria = int(prod.get("avaria", 0) or 0)
    running = int(prod.get("running", 0) or 0)
    paused = int(prod.get("paused", 0) or 0)
    done = int(prod.get("done", 0) or 0)
    avaria_ops = str(prod.get("avaria_ops", "") or "").strip()
    current_ops = str(prod.get("current_ops", "") or "").strip()
    paused_ops = str(prod.get("paused_ops", "") or "").strip()
    if avaria:
        return f"Avaria{f' | {avaria_ops}' if avaria_ops else ''}"
    if running:
        return f"Em curso{f' | {current_ops}' if current_ops else ''}"
    if paused:
        return f"Em pausa{f' | {paused_ops}' if paused_ops else ''}"
    if done and not (running or paused or avaria):
        return "Concluida"
    return "Por iniciar"


def _runtime_live_orders_from_events(data: dict[str, Any]) -> dict[str, dict[str, Any]]:
    events = sorted(
        list(data.get("op_eventos", []) or []),
        key=lambda row: str(row.get("created_at", "") or ""),
    )
    paragens = list(data.get("op_paragens", []) or [])
    open_ops: dict[tuple[str, str, str], str] = {}
    for ev in events:
        enc = str(ev.get("encomenda_numero", "") or ev.get("encomenda", "") or "").strip()
        peca = str(ev.get("peca_id", "") or "").strip()
        if not enc or not peca:
            continue
        op_raw = str(ev.get("operacao", "") or "").strip()
        op = desktop_main.normalize_operacao_nome(op_raw) or op_raw or "Em curso"
        op_norm = desktop_pulse._norm_text(op)
        evento_norm = desktop_pulse._norm_text(ev.get("evento", ""))
        key = (enc, peca, op_norm)
        if evento_norm in ("start_op", "resume_piece"):
            open_ops[key] = op
        elif evento_norm == "finish_op":
            if key in open_ops:
                open_ops.pop(key, None)
            else:
                fallback_key = next((k for k in list(open_ops.keys()) if k[0] == enc and k[1] == peca), None)
                if fallback_key:
                    open_ops.pop(fallback_key, None)

    out: dict[str, dict[str, Any]] = {}
    for (enc, _peca, _op_norm), op_label in open_ops.items():
        row = out.setdefault(enc, {"running": 0, "paused": 0, "avaria": 0, "current_ops": set(), "paused_ops": set(), "avaria_ops": set()})
        row["running"] += 1
        if op_label:
            row["current_ops"].add(op_label)

    for pa in paragens:
        enc = str(pa.get("encomenda_numero", "") or "").strip()
        if not enc:
            continue
        state_norm = desktop_pulse._norm_text(pa.get("estado", ""))
        is_closed = bool(str(pa.get("fechada_at", "") or "").strip()) or bool("fech" in state_norm)
        if is_closed:
            continue
        dur = desktop_pulse._parse_float(pa.get("duracao_min", 0), 0)
        if dur > 0:
            continue
        row = out.setdefault(enc, {"running": 0, "paused": 0, "avaria": 0, "current_ops": set(), "paused_ops": set(), "avaria_ops": set()})
        row["avaria"] += 1
        causa = str(pa.get("causa", "") or "").strip() or "Paragem"
        row["avaria_ops"].add(causa)

    for enc, row in list(out.items()):
        row["current_ops"] = " | ".join(sorted(list(row["current_ops"])))
        row["paused_ops"] = " | ".join(sorted(list(row["paused_ops"])))
        row["avaria_ops"] = " | ".join(sorted(list(row["avaria_ops"])))
    return out


def _safe_dt(raw: Any) -> datetime | None:
    if raw is None:
        return None
    if isinstance(raw, datetime):
        dt = raw
    else:
        txt = str(raw or "").strip()
        if not txt:
            return None
        dt = None
        try:
            dt = desktop_main.iso_to_dt(txt)
        except Exception:
            dt = None
        if dt is None:
            try:
                dt = datetime.fromisoformat(txt.replace("Z", "+00:00"))
            except Exception:
                return None
    try:
        if getattr(dt, "tzinfo", None) is not None:
            return dt.astimezone().replace(tzinfo=None)
    except Exception:
        pass
    return dt


_POSTO_RE = re.compile(r"\[POSTO:([^\]]+)\]", flags=re.IGNORECASE)


def _extract_posto_tag(raw: Any) -> str:
    txt = str(raw or "").strip()
    if not txt:
        return ""
    m = _POSTO_RE.search(txt)
    if not m:
        return ""
    return str(m.group(1) or "").strip()


def _strip_posto_tag(raw: Any) -> str:
    txt = str(raw or "").strip()
    if not txt:
        return ""
    return _POSTO_RE.sub("", txt).strip()


def _resolve_posto(data: dict[str, Any], operador: str, *notes: Any) -> str:
    for n in notes:
        posto = _extract_posto_tag(n)
        if posto:
            return posto
    try:
        posto_map = dict(data.get("operador_posto_map", {}) or {})
    except Exception:
        posto_map = {}
    return str(posto_map.get(str(operador or "").strip(), "") or "").strip()


def _collect_open_avaria_alerts(data: dict[str, Any], limit: int = 8) -> list[dict[str, Any]]:
    rows = sorted(
        list(data.get("op_paragens", []) or []),
        key=lambda row: str(row.get("created_at", "") or ""),
        reverse=True,
    )
    out: list[dict[str, Any]] = []
    seen: set[tuple[str, str, str, str, str]] = set()
    for row in rows:
        if not isinstance(row, dict):
            continue
        state_norm = desktop_pulse._norm_text(row.get("estado", ""))
        is_closed = bool(str(row.get("fechada_at", "") or "").strip()) or ("fech" in state_norm)
        if is_closed:
            continue
        enc = str(row.get("encomenda_numero", "") or "").strip()
        peca = str(row.get("peca_id", "") or "").strip()
        operador = str(row.get("operador", "") or "").strip()
        causa = str(row.get("causa", "") or "").strip() or "Avaria"
        detalhe = str(row.get("detalhe", "") or "").strip()
        posto = _resolve_posto(data, operador, detalhe)
        dedup_key = (enc, peca, posto, operador, causa)
        if dedup_key in seen:
            continue
        seen.add(dedup_key)
        info_chunks = []
        if enc:
            info_chunks.append(f"Enc {enc}")
        if posto:
            info_chunks.append(f"Maquina/Posto {posto}")
        if operador:
            info_chunks.append(f"Operador {operador}")
        text = "AVARIA ABERTA"
        if info_chunks:
            text += " | " + " | ".join(info_chunks)
        if causa:
            text += f" | Motivo: {causa}"
        out.append(
            {
                "id": f"avaria|{str(row.get('created_at', '') or '').strip()}|{enc}|{peca}|{posto}|{operador}|{causa}",
                "kind": "avaria",
                "critical": True,
                "text": text,
                "encomenda": enc,
                "peca_id": peca,
                "posto": posto,
                "operador": operador,
                "motivo": causa,
                "created_at": str(row.get("created_at", "") or ""),
            }
        )
        if len(out) >= max(1, int(limit)):
            break
    return out


def _collect_poke_alerts(data: dict[str, Any], max_age_minutes: int = 30, limit: int = 8) -> list[dict[str, Any]]:
    rows = sorted(
        list(data.get("op_eventos", []) or []),
        key=lambda row: str(row.get("created_at", "") or ""),
        reverse=True,
    )
    now_dt = datetime.now()
    max_age = timedelta(minutes=max(1, int(max_age_minutes)))
    out: list[dict[str, Any]] = []
    seen: set[tuple[str, str, str, str]] = set()
    for row in rows:
        if not isinstance(row, dict):
            continue
        evento_norm = desktop_pulse._norm_text(row.get("evento", ""))
        if evento_norm not in ("poke_chefia", "alert_chefia", "poke"):
            continue
        created_txt = str(row.get("created_at", "") or "").strip()
        created_dt = _safe_dt(created_txt)
        if created_dt is not None and (now_dt - created_dt) > max_age:
            continue
        enc = str(row.get("encomenda_numero", "") or "").strip()
        operador = str(row.get("operador", "") or "").strip()
        info = str(row.get("info", "") or "").strip()
        posto = _resolve_posto(data, operador, info)
        info_clean = _strip_posto_tag(info)
        dedup_key = (enc, posto, operador, info_clean)
        if dedup_key in seen:
            continue
        seen.add(dedup_key)
        text = "POKE CHEFIA"
        info_chunks = []
        if enc:
            info_chunks.append(f"Enc {enc}")
        if posto:
            info_chunks.append(f"Posto {posto}")
        if operador:
            info_chunks.append(f"Operador {operador}")
        if info_chunks:
            text += " | " + " | ".join(info_chunks)
        if info_clean:
            text += f" | {info_clean}"
        out.append(
            {
                "id": f"poke|{created_txt}|{enc}|{posto}|{operador}|{info_clean}",
                "kind": "poke",
                "critical": True,
                "text": text,
                "encomenda": enc,
                "posto": posto,
                "operador": operador,
                "mensagem": info_clean,
                "created_at": created_txt,
            }
        )
        if len(out) >= max(1, int(limit)):
            break
    return out


def _collect_local_chefia_alerts(data: dict[str, Any], max_age_minutes: int = 60, limit: int = 8) -> list[dict[str, Any]]:
    rows = sorted(
        list(data.get("chefia_alertas", []) or []),
        key=lambda row: str((row or {}).get("created_at", "") or ""),
        reverse=True,
    )
    now_dt = datetime.now()
    max_age = timedelta(minutes=max(1, int(max_age_minutes)))
    out: list[dict[str, Any]] = []
    seen: set[tuple[str, str, str, str]] = set()
    for row in rows:
        if not isinstance(row, dict):
            continue
        created_txt = str(row.get("created_at", "") or "").strip()
        created_dt = _safe_dt(created_txt)
        if created_dt is not None and (now_dt - created_dt) > max_age:
            continue
        enc = str(row.get("encomenda_numero", "") or "").strip()
        operador = str(row.get("operador", "") or "").strip()
        posto = str(row.get("posto", "") or "").strip()
        msg = str(row.get("mensagem", "") or "").strip()
        dedup_key = (enc, posto, operador, msg)
        if dedup_key in seen:
            continue
        seen.add(dedup_key)
        text = "POKE CHEFIA"
        chunks = []
        if enc:
            chunks.append(f"Enc {enc}")
        if posto:
            chunks.append(f"Posto {posto}")
        if operador:
            chunks.append(f"Operador {operador}")
        if chunks:
            text += " | " + " | ".join(chunks)
        if msg:
            text += f" | {msg}"
        out.append(
            {
                "id": f"poke_local|{created_txt}|{enc}|{posto}|{operador}|{msg}",
                "kind": "poke",
                "critical": True,
                "text": text,
                "encomenda": enc,
                "posto": posto,
                "operador": operador,
                "mensagem": msg,
                "created_at": created_txt,
            }
        )
        if len(out) >= max(1, int(limit)):
            break
    return out


def _piece_output_qty(piece: dict[str, Any]) -> float:
    total = (
        desktop_pulse._parse_float(piece.get("produzido_ok", 0), 0)
        + desktop_pulse._parse_float(piece.get("produzido_nok", 0), 0)
        + desktop_pulse._parse_float(piece.get("produzido_qualidade", 0), 0)
    )
    if total > 0:
        return round(max(0.0, total), 1)
    fallback = desktop_pulse._parse_float(piece.get("quantidade_produzida", piece.get("produzido", 0)), 0)
    return round(max(0.0, fallback), 1)


def _progress_pct(planned: float, produced: float, done: bool = False) -> float:
    plan = max(0.0, float(planned or 0))
    prod = max(0.0, float(produced or 0))
    if plan <= 0:
        return 100.0 if done or prod > 0 else 0.0
    return round(min(100.0, max(0.0, (prod / plan) * 100.0)), 1)


def _runtime_live_piece_ops(data: dict[str, Any]) -> dict[tuple[str, str], dict[str, Any]]:
    events = sorted(
        list(data.get("op_eventos", []) or []),
        key=lambda row: str((row or {}).get("created_at", "") or ""),
    )
    open_ops: dict[tuple[str, str, str], dict[str, Any]] = {}
    for ev in events:
        if not isinstance(ev, dict):
            continue
        enc = str(ev.get("encomenda_numero", "") or "").strip()
        peca = str(ev.get("peca_id", "") or "").strip()
        if not enc or not peca:
            continue
        evento_norm = desktop_pulse._norm_text(ev.get("evento", ""))
        op_raw = str(ev.get("operacao", "") or "").strip()
        op_label = desktop_main.normalize_operacao_nome(op_raw) or op_raw or "Em curso"
        op_norm = desktop_pulse._norm_text(op_label)
        if evento_norm == "start_op" and op_norm:
            open_ops[(enc, peca, op_norm)] = {
                "operacao": op_label,
                "operador": str(ev.get("operador", "") or "").strip(),
                "created_at": str(ev.get("created_at", "") or "").strip(),
            }
            continue
        if evento_norm not in ("finish_op", "pause_piece", "paragem", "stop", "close_avaria"):
            continue
        if op_norm and (enc, peca, op_norm) in open_ops:
            open_ops.pop((enc, peca, op_norm), None)
            continue
        stale_keys = [key for key in list(open_ops.keys()) if key[0] == enc and key[1] == peca]
        for key in stale_keys:
            open_ops.pop(key, None)
    out: dict[tuple[str, str], dict[str, Any]] = {}
    for (enc, peca, _op_norm), row in open_ops.items():
        item = out.setdefault((enc, peca), {"ops": [], "operador": "", "created_at": ""})
        op_label = str(row.get("operacao", "") or "").strip()
        if op_label and op_label not in item["ops"]:
            item["ops"].append(op_label)
        operador = str(row.get("operador", "") or "").strip()
        if operador and not item["operador"]:
            item["operador"] = operador
        created_at = str(row.get("created_at", "") or "").strip()
        if created_at and (not item["created_at"] or created_at < item["created_at"]):
            item["created_at"] = created_at
    return out


def _runtime_open_avaria_piece_index(data: dict[str, Any]) -> dict[tuple[str, str], dict[str, Any]]:
    rows = sorted(
        list(data.get("op_paragens", []) or []),
        key=lambda row: str((row or {}).get("created_at", "") or ""),
        reverse=True,
    )
    out: dict[tuple[str, str], dict[str, Any]] = {}
    for row in rows:
        if not isinstance(row, dict):
            continue
        state_norm = desktop_pulse._norm_text(row.get("estado", ""))
        is_closed = bool(str(row.get("fechada_at", "") or "").strip()) or ("fech" in state_norm)
        if is_closed:
            continue
        enc = str(row.get("encomenda_numero", "") or "").strip()
        peca = str(row.get("peca_id", "") or "").strip()
        ref = str(row.get("ref_interna", "") or "").strip()
        if enc and peca and (enc, peca) not in out:
            out[(enc, peca)] = row
        if enc and ref and (enc, f"ref::{ref}") not in out:
            out[(enc, f"ref::{ref}")] = row
    return out


def _runtime_open_avaria_for_piece(index: dict[tuple[str, str], dict[str, Any]], enc: str, piece: dict[str, Any]) -> dict[str, Any]:
    piece_id = str(piece.get("id", "") or "").strip()
    ref = str(piece.get("ref_interna", "") or "").strip()
    return index.get((enc, piece_id)) or index.get((enc, f"ref::{ref}")) or {}


def _format_duration_minutes(minutes: float) -> str:
    mins = max(0.0, float(minutes or 0))
    if mins < 60:
        return f"{mins:.0f} min"
    hours = int(mins // 60)
    rem = int(round(mins % 60))
    return f"{hours:02d}h{rem:02d}"


def _group_live_state(pieces: list[dict[str, Any]], fallback: str = "") -> str:
    states = [desktop_pulse._norm_text((row or {}).get("estado", "")) for row in list(pieces or [])]
    if any(("produ" in st) or ("curso" in st) for st in states):
        return "Em producao"
    if any("avari" in st for st in states):
        return "Avaria"
    if any(("paus" in st) or ("interromp" in st) for st in states):
        return "Em pausa"
    if states and all("concl" in st for st in states):
        return "Concluida"
    if any("incomplet" in st for st in states):
        return "Incompleta"
    return fallback or "Preparacao"


def get_mobile_alerts() -> dict[str, Any]:
    ctx = _make_context()
    _refresh_ctx_runtime(ctx)
    data = ctx.data if isinstance(ctx.data, dict) else {}
    avarias = _collect_open_avaria_alerts(data, limit=8)
    pokes = _collect_poke_alerts(data, max_age_minutes=30, limit=8)
    local_pokes = _collect_local_chefia_alerts(data, max_age_minutes=60, limit=8)
    merged = avarias + pokes + local_pokes
    items: list[dict[str, Any]] = []
    seen: set[tuple[str, str, str]] = set()
    for row in merged:
        if not isinstance(row, dict):
            continue
        key = (
            str(row.get("kind", "") or "").strip(),
            str(row.get("created_at", "") or "").strip(),
            str(row.get("text", "") or "").strip(),
        )
        if key in seen:
            continue
        seen.add(key)
        items.append(row)
    banner = " | ".join([str(a.get("text", "") or "").strip() for a in items[:2] if str(a.get("text", "") or "").strip()])
    return {
        "critical": bool(items),
        "count": len(items),
        "banner": banner,
        "items": items,
        "updated_at": desktop_main.now_iso(),
    }


def get_mobile_avarias(history_limit: int = 80) -> dict[str, Any]:
    ctx = _make_context()
    _refresh_ctx_runtime(ctx)
    data = ctx.data if isinstance(ctx.data, dict) else {}
    rows = sorted(
        list(data.get("op_paragens", []) or []),
        key=lambda row: str((row or {}).get("created_at", "") or ""),
        reverse=True,
    )
    now_dt = datetime.now()
    open_items: list[dict[str, Any]] = []
    history_items: list[dict[str, Any]] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        created_txt = str(row.get("created_at", "") or "").strip()
        closed_txt = str(row.get("fechada_at", "") or "").strip()
        state_norm = desktop_pulse._norm_text(row.get("estado", ""))
        dur = max(0.0, desktop_pulse._parse_float(row.get("duracao_min", 0), 0))
        created_dt = _safe_dt(created_txt)
        closed_dt = _safe_dt(closed_txt)
        is_closed = bool(closed_txt) or ("fech" in state_norm)
        if not is_closed and created_dt is not None:
            dur = max(0.0, (now_dt - created_dt).total_seconds() / 60.0)
        enc = str(row.get("encomenda_numero", "") or "").strip()
        operador = str(row.get("operador", "") or "").strip()
        detalhe = str(row.get("detalhe", "") or "").strip()
        item = {
            "id": f"avaria_row|{created_txt}|{enc}|{str(row.get('peca_id', '') or '').strip()}|{str(row.get('causa', '') or '').strip()}",
            "encomenda": enc,
            "peca_id": str(row.get("peca_id", "") or "").strip(),
            "ref_interna": str(row.get("ref_interna", "") or "").strip(),
            "material": str(row.get("material", "") or "-").strip() or "-",
            "espessura": str(row.get("espessura", "") or "-").strip() or "-",
            "operador": operador,
            "posto": _resolve_posto(data, operador, detalhe),
            "motivo": str(row.get("causa", "") or "").strip() or "Avaria",
            "detalhe": _strip_posto_tag(detalhe),
            "created_at": created_txt,
            "fechada_at": closed_txt,
            "duracao_min": round(dur, 1),
            "duracao_txt": _format_duration_minutes(dur),
            "estado": "Fechada" if is_closed else "Em aberto",
            "critical": not is_closed,
        }
        if not is_closed:
            open_items.append(item)
        elif len(history_items) < max(1, int(history_limit)):
            history_items.append(item)
    headline = ""
    if open_items:
        top = open_items[0]
        headline = (
            f"Avaria em curso | Enc {top.get('encomenda', '-')}"
            f" | Posto {top.get('posto', '-')}"
            f" | {top.get('motivo', 'Avaria')}"
        )
    return {
        "summary": {
            "open_count": len(open_items),
            "history_count": len(history_items),
            "critical": bool(open_items),
            "headline": headline,
        },
        "open": open_items,
        "history": history_items,
        "updated_at": desktop_main.now_iso(),
    }


def get_material_separation(horizon_days: int = 4) -> dict[str, Any]:
    backend = _material_backend()
    backend.reload()
    rows = [dict(row or {}) for row in list(backend.material_assistant_separation_rows(horizon_days=horizon_days) or [])]
    checked_sep = 0
    postos: list[dict[str, Any]] = []
    current_group: dict[str, Any] | None = None
    for row in rows:
        if bool(row.get("visto_sep_checked")):
            checked_sep += 1
        posto_txt = str(row.get("posto_trabalho", "") or "Sem posto").strip() or "Sem posto"
        if current_group is None or str(current_group.get("posto_trabalho", "") or "") != posto_txt:
            current_group = {
                "posto_trabalho": posto_txt,
                "count": 0,
                "checked_sep": 0,
                "rows": [],
            }
            postos.append(current_group)
        current_group["count"] = int(current_group.get("count", 0) or 0) + 1
        if bool(row.get("visto_sep_checked")):
            current_group["checked_sep"] = int(current_group.get("checked_sep", 0) or 0) + 1
        current_group["rows"].append(
            {
                "check_key": str(row.get("check_key", "") or "").strip(),
                "numero": str(row.get("numero", "") or "-").strip() or "-",
                "cliente": str(row.get("cliente", "") or "-").strip() or "-",
                "material": str(row.get("material", "") or "-").strip() or "-",
                "espessura": str(row.get("espessura", "") or "-").strip() or "-",
                "dimensao": str(row.get("dimensao", "") or "-").strip() or "-",
                "lote": str(row.get("lote_sugerido", row.get("lote_atual", "-")) or "-").strip() or "-",
                "quantidade": float(row.get("quantidade", 0) or 0),
                "planeado_dia": str(row.get("planeamento_dia", "") or "-").strip() or "-",
                "planeado_hora": str(row.get("planeamento_hora", "") or "-").strip() or "-",
                "turno": str(row.get("planeamento_turno", "") or "-").strip() or "-",
                "acao_sugerida": str(row.get("acao_sugerida", "") or "-").strip() or "-",
                "alerta_texto": str(row.get("alerta_texto", "") or "").strip(),
                "reserva_estado": str(row.get("reserva_estado", "") or "").strip(),
                "priority_label": str(row.get("priority_label", "") or "Media").strip() or "Media",
                "visto_sep_checked": bool(row.get("visto_sep_checked")),
            }
        )
    return {
        "updated_at": desktop_main.now_iso(),
        "horizon_days": max(1, int(horizon_days or 4)),
        "horizon_label": f"{max(1, int(horizon_days or 4))} dias úteis",
        "summary": {
            "rows": len(rows),
            "checked_sep": checked_sep,
            "pending_sep": max(0, len(rows) - checked_sep),
            "postos": len(postos),
        },
        "postos": postos,
    }


def set_material_separation_check(check_key: str, checked: bool) -> dict[str, Any]:
    backend = _material_backend()
    backend.material_assistant_set_check(str(check_key or "").strip(), "sep", bool(checked))
    return {"ok": True, "check_key": str(check_key or "").strip(), "checked": bool(checked)}


def build_material_separation_pdf(horizon_days: int = 4) -> Path:
    backend = _material_backend()
    backend.reload()
    return Path(backend.material_assistant_render_separation_pdf(horizon_days=horizon_days))


def set_plan_delay_reason(item_key: str, reason: str) -> dict[str, Any]:
    backend = _material_backend()
    backend.reload()
    backend.pulse_plan_delay_set_reason(str(item_key or "").strip(), str(reason or "").strip())
    return {"ok": True, "item_key": str(item_key or "").strip(), "reason": str(reason or "").strip()}


def clear_plan_delay_reason(item_key: str) -> dict[str, Any]:
    backend = _material_backend()
    backend.reload()
    backend.pulse_plan_delay_clear_reason(str(item_key or "").strip())
    return {"ok": True, "item_key": str(item_key or "").strip()}


def get_dashboard(period: str | int = "7 dias", year: str | None = None, encomenda: str = "Todas", visao: str = "Todas", origem: str = "Ambos") -> dict[str, Any]:
    ctx = _make_context()
    metrics = desktop_pulse._compute_production_pulse_metrics(ctx, period_days=_period_to_days(period), year_filter=year)
    backend = _material_backend()
    backend.reload()
    enc_filter = str(encomenda or "Todas").strip()
    view_mode = str(visao or "Todas").strip()
    origin_mode = str(origem or "Ambos").strip()
    origem_norm = desktop_pulse._norm_text(origin_mode)

    running = list(metrics.get("pecas_tempo", []) or [])
    if enc_filter and enc_filter.lower() != "todas":
        running = [r for r in running if str(r.get("encomenda", "") or "").strip() == enc_filter]
    if desktop_pulse._norm_text(view_mode).startswith("so"):
        running = [r for r in running if bool(r.get("fora"))]
    if "histor" in origem_norm:
        running = []

    history_raw = list(metrics.get("historico_tempo", []) or [])
    if enc_filter and enc_filter.lower() != "todas":
        history_raw = [r for r in history_raw if str(r.get("encomenda", "") or "").strip() == enc_filter]
    if desktop_pulse._norm_text(view_mode).startswith("so"):
        history_raw = [r for r in history_raw if bool(r.get("fora"))]
    if "curso" in origem_norm:
        history_raw = []

    history = _build_history_snapshot(ctx, metrics, enc_filter, view_mode, origin_mode, year)
    if "curso" in origem_norm:
        history = []

    interrupted = list(metrics.get("interrompidas", []) or [])
    if enc_filter and enc_filter.lower() != "todas":
        interrupted = [r for r in interrupted if str(r.get("encomenda", "") or "").strip() == enc_filter]

    top_stop_rows = list(metrics.get("top_paragens", []) or [])
    if enc_filter and enc_filter.lower() != "todas":
        top_stop_rows = [r for r in top_stop_rows if str(r.get("encomenda", "") or "").strip() == enc_filter]
    paragens_min = sum(max(0.0, desktop_pulse._parse_float(r.get("minutos", 0), 0)) for r in top_stop_rows)
    top_stops = top_stop_rows[:8]

    pecas_em_curso = len(running)
    pecas_fora_tempo = sum(1 for r in running if bool(r.get("fora"))) + sum(1 for r in history_raw if bool(r.get("fora")))
    desvio_max_min = 0.0
    for row in list(running) + list(history):
        desvio_max_min = max(desvio_max_min, max(0.0, desktop_pulse._parse_float(row.get("delta_min", 0), 0)))

    perf_plan_total = 0.0
    perf_real_total = 0.0
    for row in list(running) + list(history_raw):
        plan_v = max(0.0, desktop_pulse._parse_float(row.get("plan_min", 0), 0))
        real_v = max(0.0, desktop_pulse._parse_float(row.get("elapsed_min", row.get("tempo_real_min", 0)), 0))
        if plan_v <= 0 or real_v <= 0:
            continue
        perf_plan_total += plan_v
        perf_real_total += real_v

    no_operational_data = not running and not history_raw and perf_plan_total <= 0 and perf_real_total <= 0

    oee = round(desktop_pulse._parse_float(metrics.get("oee", 0), 0), 1)
    disponibilidade = round(desktop_pulse._parse_float(metrics.get("disponibilidade", 0), 0), 1)
    performance_raw = 0.0
    if perf_plan_total > 0 and perf_real_total > 0:
        performance_raw = (perf_plan_total / perf_real_total) * 100.0
        performance = round(max(0.0, min(250.0, performance_raw)), 1)
    else:
        performance = round(desktop_pulse._parse_float(metrics.get("performance", 0), 0), 1)
    qualidade = round(desktop_pulse._parse_float(metrics.get("qualidade", 0), 0), 1)
    andon = {
        "prod": int(metrics.get("andon_prod", 0) or 0),
        "setup": int(metrics.get("andon_setup", 0) or 0),
        "espera": int(metrics.get("andon_wait", 0) or 0),
        "stop": int(metrics.get("andon_stop", 0) or 0),
    }
    if no_operational_data:
        oee = 0.0
        disponibilidade = 0.0
        performance = 0.0
        performance_raw = 0.0
        qualidade = 0.0
        andon = {"prod": 0, "setup": 0, "espera": 0, "stop": 0}
    plan_delay = backend.pulse_plan_delay_rows(
        period_days=_period_to_days(period),
        year_filter=year,
        encomenda=enc_filter,
    )
    plan_delay_open = int(plan_delay.get("open_count", 0) or 0)
    plan_delay_ack = int(plan_delay.get("acknowledged_count", 0) or 0)
    alertas: list[str] = []
    if no_operational_data:
        alertas.append("- Sem dados operacionais no período selecionado.")
    if (not no_operational_data) and pecas_em_curso > 0 and disponibilidade < 85.0:
        alertas.append(f"- Disponibilidade baixa ({disponibilidade:.1f}%).")
    if (not no_operational_data) and qualidade < 97.0:
        alertas.append(f"- Qualidade abaixo da meta ({qualidade:.1f}%).")
    if (not no_operational_data) and pecas_em_curso > 0 and performance < 80.0:
        alertas.append(f"- Performance baixa ({performance:.1f}%).")
    if (not no_operational_data) and paragens_min >= 480.0:
        alertas.append(f"- Paragens elevadas no filtro ({paragens_min:.1f} min).")
    if (not no_operational_data) and pecas_fora_tempo > 0:
        alertas.append(f"- {pecas_fora_tempo} registo(s) fora do tempo planeado.")
    if plan_delay_open > 0:
        alertas.append(f"- {plan_delay_open} grupo(s) de corte laser estão fora do horário planeado e sem baixa registada.")
    elif plan_delay_ack > 0:
        alertas.append(f"- {plan_delay_ack} grupo(s) com atraso ao planeamento foram justificados.")
    if not alertas:
        alertas.append("- Sem alertas criticos.")

    return {
        "summary": {
            "oee": oee,
            "disponibilidade": disponibilidade,
            "performance": performance,
            "performance_raw": round(performance_raw or performance, 1),
            "perf_plan_total": round(perf_plan_total, 1),
            "perf_real_total": round(perf_real_total, 1),
            "qualidade": qualidade,
            "paragens_min": round(paragens_min, 1),
            "pecas_em_curso": int(pecas_em_curso),
            "pecas_fora_tempo": int(pecas_fora_tempo),
            "desvio_max_min": round(desvio_max_min, 1),
            "andon": andon,
            "quality_scope": str(metrics.get("quality_scope", "") or ""),
            "alerts": "\n".join(alertas),
            "plan_delay_open": plan_delay_open,
            "plan_delay_acknowledged": plan_delay_ack,
        },
        "running": running,
        "history": history,
        "top_stops": top_stops,
        "interrupted": interrupted,
        "plan_delay": plan_delay,
        "updated_at": desktop_main.now_iso(),
    }


def list_encomendas(year: str | None = None) -> list[dict[str, Any]]:
    ctx = _make_context()
    _refresh_ctx_runtime(ctx)
    client_lookup = _client_lookup_map(ctx)
    runtime_live = _runtime_live_orders_from_events(ctx.data)
    rows = []
    for enc in list(ctx.data.get("encomendas", []) or []):
        if not isinstance(enc, dict):
            continue
        if year and not _match_year(enc, year):
            continue
        prod = _order_production_snapshot(enc)
        enc_num = str(enc.get("numero", "") or "").strip()
        client_meta = _resolve_client_identity(enc, client_lookup)
        live = runtime_live.get(enc_num, {})
        live_running = int(live.get("running", 0) or 0)
        live_avaria = int(live.get("avaria", 0) or 0)
        if live_running > 0:
            prod["running"] = max(int(prod.get("running", 0) or 0), live_running)
            if str(live.get("current_ops", "") or "").strip():
                prod["current_ops"] = str(live.get("current_ops", "") or "").strip()
            # Se ha operacao aberta em curso, nao queremos mostrar pausa como estado dominante.
            if int(prod.get("running", 0) or 0) > 0:
                prod["paused"] = 0
                prod["paused_ops"] = ""
        if live_avaria > 0:
            prod["avaria"] = max(int(prod.get("avaria", 0) or 0), live_avaria)
            if str(live.get("avaria_ops", "") or "").strip():
                prod["avaria_ops"] = str(live.get("avaria_ops", "") or "").strip()
        prod["resumo"] = _compose_production_resumo(prod)
        rows.append({
            "numero": enc_num,
            "cliente": str(client_meta.get("cliente", "") or "").strip(),
            "cliente_nome": str(client_meta.get("cliente_nome", "") or "").strip(),
            "cliente_display": str(client_meta.get("cliente_display", "") or "").strip(),
            "estado": str(enc.get("estado", "") or "").strip(),
            "data_entrega": str(enc.get("data_entrega", "") or "").strip(),
            "nota_cliente": str(enc.get("nota_cliente", "") or "").strip(),
            "tempo_h": desktop_pulse._parse_float(enc.get("tempo", 0), 0),
            "producao_resumo": prod["resumo"],
            "pecas_em_curso": prod["running"],
            "pecas_em_pausa": prod["paused"],
            "pecas_em_avaria": prod["avaria"],
            "ops_ativas": prod["current_ops"],
            "ops_pausadas": prod["paused_ops"],
            "ops_avaria": prod["avaria_ops"],
        })
    rows.sort(key=lambda r: r.get("numero", ""), reverse=True)
    return rows


def get_encomenda(numero: str) -> dict[str, Any] | None:
    ctx = _make_context()
    _refresh_ctx_runtime(ctx)
    numero_txt = str(numero or "").strip()
    if not numero_txt:
        return None
    for enc in list(ctx.data.get("encomendas", []) or []):
        if str(enc.get("numero", "") or "").strip() != numero_txt:
            continue
        return enc
    return None


def _match_plan_year(item: dict[str, Any], year_filter: str | None) -> bool:
    if not year_filter:
        return True
    raw = item.get("data") or item.get("data_planeada") or item.get("movido_em")
    try:
        d_ref = desktop_pulse._safe_date(raw)
        if d_ref is not None:
            return str(d_ref.year) == str(year_filter)
    except Exception:
        pass
    return True


def _resolve_week_start(year_filter: str | None, raw_week_start: str | None) -> date:
    if raw_week_start:
        try:
            parsed = desktop_pulse._safe_date(raw_week_start)
        except Exception:
            parsed = None
        if parsed is not None:
            return parsed - timedelta(days=parsed.weekday())
    today = datetime.now().date()
    default_start = today - timedelta(days=today.weekday())
    year_txt = str(year_filter or "").strip()
    if year_txt.isdigit():
        year_num = int(year_txt)
        if default_start.year != year_num:
            seed = date(year_num, 1, 1)
            default_start = seed - timedelta(days=seed.weekday())
    return default_start


def _client_lookup_map(ctx: SimpleNamespace) -> dict[str, str]:
    lookup: dict[str, str] = {}
    for row in list((ctx.data or {}).get("clientes", []) or []):
        if not isinstance(row, dict):
            continue
        code = str(row.get("codigo", "") or "").strip()
        name = str(row.get("nome", "") or "").strip()
        if code and name and code.lower() not in lookup:
            lookup[code.lower()] = name
    return lookup


def _resolve_client_identity(enc: dict[str, Any], client_lookup: dict[str, str] | None = None) -> dict[str, str]:
    code = str(enc.get("cliente", "") or "").strip()
    name = str(enc.get("cliente_nome", "") or "").strip()
    if not name and code and isinstance(client_lookup, dict):
        name = str(client_lookup.get(code.lower(), "") or "").strip()
    display = " | ".join(part for part in [code, name] if part).strip()
    return {
        "cliente": code,
        "cliente_nome": name,
        "cliente_display": display or code or name or "-",
    }


def _order_client_map(ctx: SimpleNamespace) -> dict[str, dict[str, str]]:
    client_lookup = _client_lookup_map(ctx)
    orders: dict[str, dict[str, str]] = {}
    for enc in list((ctx.data or {}).get("encomendas", []) or []):
        if not isinstance(enc, dict):
            continue
        numero = str(enc.get("numero", "") or "").strip()
        if not numero:
            continue
        orders[numero] = _resolve_client_identity(enc, client_lookup)
    return orders


def _plan_row_payload(row: dict[str, Any], *, history: bool = False, order_client_map: dict[str, dict[str, str]] | None = None) -> dict[str, Any]:
    board_color = str(desktop_plan._pdf_block_color_for_item(row) or "#dbeafe").strip() or "#dbeafe"
    enc_num = str(row.get("encomenda", "") or row.get("encomenda_numero", "") or "").strip()
    client_meta = dict((order_client_map or {}).get(enc_num, {}) or {})
    op_normalize = getattr(desktop_main, "normalize_planeamento_operacao", desktop_main.normalize_operacao_nome)
    payload = {
        "id": str(row.get("id", "") or row.get("bloco_id", "") or "").strip(),
        "encomenda": enc_num,
        "data": str(row.get("data", "") or row.get("data_planeada", "") or "").strip(),
        "inicio": str(row.get("inicio", "") or "").strip(),
        "duracao_min": round(max(0.0, desktop_pulse._parse_float(row.get("duracao_min", 0), 0)), 1),
        "material": str(row.get("material", "") or "-").strip() or "-",
        "espessura": str(row.get("espessura", "") or "-").strip() or "-",
        "operacao": str(op_normalize(row.get("operacao", "")) or "Corte Laser").strip(),
        "chapa": str(row.get("chapa", "") or "-").strip() or "-",
        "cliente": str(client_meta.get("cliente", "") or "").strip(),
        "cliente_nome": str(client_meta.get("cliente_nome", "") or "").strip(),
        "cliente_display": str(client_meta.get("cliente_display", "") or "").strip(),
        "color": board_color,
        "source_color": str(row.get("color", "") or "#dbeafe"),
    }
    if history:
        payload["duracao_min"] = round(
            max(0.0, desktop_pulse._parse_float(row.get("tempo_planeado_min", row.get("duracao_min", 0)), 0)),
            1,
        )
        payload["tempo_real_min"] = round(max(0.0, desktop_pulse._parse_float(row.get("tempo_real_min", 0), 0)), 1)
        payload["estado_final"] = str(row.get("estado_final", "") or "-").strip() or "-"
    return payload


def _status_rank(raw: str) -> int:
    txt = desktop_pulse._norm_text(raw)
    if "avari" in txt:
        return 0
    if "curso" in txt or "produc" in txt:
        return 1
    if "paus" in txt or "interromp" in txt:
        return 2
    if "prepar" in txt:
        return 3
    if "concl" in txt:
        return 5
    return 4


def _group_plan_index(data: dict[str, Any]) -> dict[tuple[str, str, str], dict[str, str]]:
    index: dict[tuple[str, str, str], dict[str, str]] = {}
    for row in list(data.get("plano", []) or []):
        if not isinstance(row, dict):
            continue
        enc_num = str(row.get("encomenda", "") or row.get("encomenda_numero", "") or "").strip()
        material = str(row.get("material", "") or "-").strip() or "-"
        espessura = str(row.get("espessura", "") or "-").strip() or "-"
        data_planeada = str(row.get("data", "") or row.get("data_planeada", "") or "").strip()
        inicio = str(row.get("inicio", "") or "").strip()
        if not enc_num or not data_planeada:
            continue
        key = (enc_num, material, espessura)
        candidate = {
            "data": data_planeada,
            "inicio": inicio,
        }
        current = index.get(key)
        if current is None:
            index[key] = candidate
            continue
        candidate_sort = (candidate.get("data", "9999-99-99"), candidate.get("inicio", "99:99") or "99:99")
        current_sort = (current.get("data", "9999-99-99"), current.get("inicio", "99:99") or "99:99")
        if candidate_sort < current_sort:
            index[key] = candidate
    return index


def get_operator_board(year: str | None = None, username: str = "", role: str = "") -> dict[str, Any]:
    ctx = _make_context()
    _refresh_ctx_runtime(ctx)
    client_lookup = _client_lookup_map(ctx)
    plan_index = _group_plan_index(ctx.data)
    items: list[dict[str, Any]] = []
    active_orders: set[str] = set()
    sum_running = 0
    sum_paused = 0
    sum_avaria = 0
    sum_done = 0
    sum_plan = 0.0
    sum_prod = 0.0
    live_piece_ops = _runtime_live_piece_ops(ctx.data)
    live_piece_avarias = _runtime_open_avaria_piece_index(ctx.data)

    for enc in list(ctx.data.get("encomendas", []) or []):
        if not isinstance(enc, dict):
            continue
        if year and not _match_year(enc, year):
            continue
        enc_num = str(enc.get("numero", "") or "").strip()
        if not enc_num:
            continue
        client_meta = _resolve_client_identity(enc, client_lookup)
        cliente_txt = str(client_meta.get("cliente_display", "") or "").strip()
        enc_estado = str(enc.get("estado", "") or "").strip()
        ordem = enc.get("ordem_fabrico", {}) if isinstance(enc.get("ordem_fabrico", {}), dict) else {}
        of_codigo = str(enc.get("of_codigo", "") or ordem.get("id", "") or ordem.get("codigo", "") or "").strip()
        entrega = str(enc.get("data_entrega", "") or "").strip()
        for mat in list(enc.get("materiais", []) or []):
            mat_nome = str(mat.get("material", "") or "-").strip() or "-"
            for esp in list(mat.get("espessuras", []) or []):
                pieces: list[dict[str, Any]] = []
                tempo_plan = max(0.0, desktop_pulse._parse_float(esp.get("tempo_min", 0), 0))
                tempo_real = 0.0
                group_plan = 0.0
                group_prod = 0.0
                group_avarias = 0
                group_running = 0
                group_paused = 0
                group_done = 0
                group_has_cut_laser = False
                for p in list(esp.get("pecas", []) or []):
                    ops = desktop_main.ensure_peca_operacoes(p)
                    piece_id = str(p.get("id", "") or "").strip()
                    piece_state = str(p.get("estado", "") or "").strip()
                    live_ops = live_piece_ops.get((enc_num, piece_id), {})
                    live_avaria = _runtime_open_avaria_for_piece(live_piece_avarias, enc_num, p)
                    current_op = " + ".join(list(live_ops.get("ops", []) or [])) or desktop_main.normalize_operacao_nome(p.get("operacao_atual", "")) or "-"
                    owner = str(live_ops.get("operador", "") or p.get("operador_atual", "") or p.get("bloqueada_por", "") or "").strip()
                    tempo_piece = max(0.0, desktop_pulse._parse_float(p.get("tempo_producao_min", 0), 0))
                    op_started_dt = _safe_dt(live_ops.get("created_at", ""))
                    op_elapsed = round(max(0.0, (datetime.now() - op_started_dt).total_seconds() / 60.0), 1) if op_started_dt is not None else 0.0
                    tempo_real = max(tempo_real, tempo_piece)
                    planned_qty = round(max(0.0, desktop_pulse._parse_float(p.get("quantidade_pedida", 0), 0)), 1)
                    produced_qty = _piece_output_qty(p)
                    group_plan += planned_qty
                    group_prod += produced_qty
                    if live_avaria:
                        piece_state = "Avaria"
                        group_avarias += 1
                    elif list(live_ops.get("ops", []) or []):
                        piece_state = "Em producao"
                    state_norm = desktop_pulse._norm_text(piece_state)
                    if "concl" in state_norm:
                        sum_done += 1
                        group_done += 1
                    elif "avari" in state_norm:
                        sum_avaria += 1
                    elif "paus" in state_norm or "interromp" in state_norm:
                        sum_paused += 1
                        group_paused += 1
                    elif "curso" in state_norm or "produc" in state_norm:
                        sum_running += 1
                        group_running += 1
                    sum_plan += planned_qty
                    sum_prod += produced_qty
                    ops_payload = []
                    live_ops_set = {desktop_pulse._norm_text(op) for op in list(live_ops.get("ops", []) or []) if str(op).strip()}
                    for op in list(ops or []):
                        nome = desktop_main.normalize_operacao_nome(op.get("nome", ""))
                        estado_op = str(op.get("estado", "") or "").strip()
                        nome_norm = desktop_pulse._norm_text(nome)
                        if nome_norm and nome_norm in live_ops_set:
                            estado_op = "Em producao"
                        elif live_avaria and "concl" not in desktop_pulse._norm_text(estado_op):
                            estado_op = "Avaria"
                        if ("laser" in nome_norm) or ("corte" in nome_norm):
                            group_has_cut_laser = True
                        ops_payload.append(
                            {
                                "nome": nome,
                                "estado": estado_op,
                                "qtd_ok": desktop_pulse._parse_float(op.get("qtd_ok", 0), 0),
                                "qtd_nok": desktop_pulse._parse_float(op.get("qtd_nok", 0), 0),
                                "qtd_qual": desktop_pulse._parse_float(op.get("qtd_qual", 0), 0),
                            }
                        )
                    piece_done = "concl" in state_norm
                    pieces.append(
                        {
                            "id": piece_id,
                            "of": str(p.get("of", "") or of_codigo).strip(),
                            "ref_interna": str(p.get("ref_interna", "") or "-").strip() or "-",
                            "ref_externa": str(p.get("ref_externa", "") or "-").strip() or "-",
                            "estado": piece_state,
                            "operacao_atual": current_op,
                            "operador": owner,
                            "tempo_min": round(op_elapsed if op_elapsed > 0 else tempo_piece, 1),
                            "tempo_total_min": round(tempo_piece, 1),
                            "tempo_operacao_min": round(op_elapsed, 1),
                            "planeado": planned_qty,
                            "produzido": produced_qty,
                            "progress_pct": _progress_pct(planned_qty, produced_qty, done=piece_done),
                            "pendentes": list(desktop_main.peca_operacoes_pendentes(p)),
                            "ops": ops_payload,
                            "motivo": str((live_avaria or {}).get("causa", "") or p.get("avaria_motivo", "") or p.get("interrupcao_peca_motivo", "") or "").strip(),
                            "avaria_aberta": bool(live_avaria),
                            "avaria_inicio": str((live_avaria or {}).get("created_at", "") or p.get("avaria_inicio_ts", "") or "").strip(),
                        }
                    )
                if not pieces:
                    continue
                group_state = _group_live_state(pieces, fallback=str(esp.get("estado", "") or "").strip())
                espessura_txt = str(esp.get("espessura", "") or "-").strip() or "-"
                plan_meta = dict(plan_index.get((enc_num, mat_nome, espessura_txt), {}) or {})
                active_orders.add(enc_num)
                items.append(
                    {
                        "encomenda": enc_num,
                        "of": of_codigo,
                        "of_codigo": of_codigo,
                        "cliente": str(client_meta.get("cliente", "") or "").strip(),
                        "cliente_nome": str(client_meta.get("cliente_nome", "") or "").strip(),
                        "cliente_display": cliente_txt,
                        "estado": enc_estado,
                        "data_entrega": entrega,
                        "material": mat_nome,
                        "espessura": espessura_txt,
                        "estado_espessura": group_state,
                        "tempo_plan_min": round(tempo_plan, 1),
                        "tempo_real_min": round(tempo_real, 1),
                        "desvio_min": round(tempo_real - tempo_plan, 1) if tempo_plan > 0 else 0.0,
                        "planeado_total": round(group_plan, 1),
                        "produzido_total": round(group_prod, 1),
                        "progress_pct": _progress_pct(group_plan, group_prod, done=bool(group_plan > 0 and group_prod >= group_plan)),
                        "pecas_em_avaria": group_avarias,
                        "pecas_em_curso": group_running,
                        "pecas_em_pausa": group_paused,
                        "pecas_concluidas": group_done,
                        "tem_corte_laser": bool(group_has_cut_laser),
                        "prazo_corte_laser_data": str(plan_meta.get("data", "") or "").strip(),
                        "prazo_corte_laser_inicio": str(plan_meta.get("inicio", "") or "").strip(),
                        "pieces": pieces,
                    }
                )

    items.sort(
        key=lambda row: (
            _status_rank(str(row.get("estado_espessura", "") or "")),
            str(row.get("encomenda", "") or ""),
            str(row.get("material", "") or ""),
            desktop_pulse._parse_float(row.get("espessura", 0), 0),
        )
    )

    return {
        "summary": {
            "encomendas_ativas": len(active_orders),
            "grupos": len(items),
            "pecas_em_curso": sum_running,
            "pecas_em_pausa": sum_paused,
            "pecas_em_avaria": sum_avaria,
            "pecas_concluidas": sum_done,
            "planeado_total": round(sum_plan, 1),
            "produzido_total": round(sum_prod, 1),
            "progress_pct": _progress_pct(sum_plan, sum_prod, done=bool(sum_plan > 0 and sum_prod >= sum_plan)),
            "username": username,
            "role": role,
        },
        "items": items,
    }


def get_planning_overview(year: str | None = None, week_start: str | None = None, operation: str | None = None) -> dict[str, Any]:
    ctx = _make_context()
    _refresh_ctx_runtime(ctx)
    client_lookup = _client_lookup_map(ctx)
    order_client_map = _order_client_map(ctx)
    op_normalize = getattr(desktop_main, "normalize_planeamento_operacao", desktop_main.normalize_operacao_nome)
    op_txt = str(op_normalize(operation or "") or "Corte Laser").strip() or "Corte Laser"
    active: list[dict[str, Any]] = []
    history: list[dict[str, Any]] = []
    backlog: list[dict[str, Any]] = []
    week_active: list[dict[str, Any]] = []
    active_minutes = 0.0
    week_minutes = 0.0
    week_start_dt = _resolve_week_start(year, week_start)
    week_dates = [week_start_dt + timedelta(days=i) for i in range(6)]
    week_end_dt = week_dates[-1]

    def row_operation(row: dict[str, Any]) -> str:
        return str(op_normalize(row.get("operacao", "")) or "Corte Laser").strip() or "Corte Laser"

    def esp_time_for_operation(esp_obj: dict[str, Any], target_op: str) -> float:
        if target_op == "Corte Laser":
            return max(0.0, desktop_pulse._parse_float(esp_obj.get("tempo_min", 0), 0))
        raw_map = dict(esp_obj.get("tempos_operacao", {}) or {})
        for op_name, raw_value in raw_map.items():
            if str(op_normalize(op_name) or "").strip() != target_op:
                continue
            return max(0.0, desktop_pulse._parse_float(raw_value, 0))
        return 0.0

    def esp_has_operation(esp_obj: dict[str, Any], target_op: str) -> bool:
        for piece in list(esp_obj.get("pecas", []) or []):
            for op in list(desktop_main.ensure_peca_operacoes(piece) or []):
                if str(op_normalize(op.get("nome", "")) or "").strip() == target_op:
                    return True
        return esp_time_for_operation(esp_obj, target_op) > 0

    for row in list(ctx.data.get("plano", []) or []):
        if not isinstance(row, dict):
            continue
        if year and not _match_plan_year(row, year):
            continue
        if row_operation(row) != op_txt:
            continue
        payload = _plan_row_payload(row, history=False, order_client_map=order_client_map)
        dur = max(0.0, desktop_pulse._parse_float(payload.get("duracao_min", 0), 0))
        active_minutes += dur
        active.append(payload)
        row_date = desktop_pulse._safe_date(payload.get("data"))
        if row_date is None or row_date < week_start_dt or row_date > week_end_dt:
            continue
        week_active.append(payload)
        week_minutes += dur

    for row in list(ctx.data.get("plano_hist", []) or []):
        if not isinstance(row, dict):
            continue
        if year and not _match_plan_year(row, year):
            continue
        if row_operation(row) != op_txt:
            continue
        history.append(_plan_row_payload(row, history=True, order_client_map=order_client_map))

    for enc in list(ctx.data.get("encomendas", []) or []):
        if not isinstance(enc, dict):
            continue
        if year and not _match_year(enc, year):
            continue
        estado = str(enc.get("estado", "") or "").strip()
        if "concl" in desktop_pulse._norm_text(estado):
            continue
        client_meta = _resolve_client_identity(enc, client_lookup)
        if op_txt == "Montagem":
            if not list(enc.get("montagem_itens", []) or []):
                continue
            tempo_plan = max(0.0, desktop_pulse._parse_float(desktop_main.encomenda_montagem_tempo_min(enc), 0))
            backlog.append(
                {
                    "numero": str(enc.get("numero", "") or "").strip(),
                    "cliente": str(client_meta.get("cliente", "") or "").strip(),
                    "cliente_nome": str(client_meta.get("cliente_nome", "") or "").strip(),
                    "cliente_display": str(client_meta.get("cliente_display", "") or "").strip(),
                    "estado": estado,
                    "data_entrega": str(enc.get("data_entrega", "") or "").strip(),
                    "tempo_plan_min": round(tempo_plan, 1),
                }
            )
            continue
        seen_any = False
        tempo_plan = 0.0
        for mat in list(enc.get("materiais", []) or []):
            for esp in list(mat.get("espessuras", []) or []):
                if not esp_has_operation(esp, op_txt):
                    continue
                seen_any = True
                tempo_plan += esp_time_for_operation(esp, op_txt)
        if not seen_any:
            continue
        backlog.append(
            {
                "numero": str(enc.get("numero", "") or "").strip(),
                "cliente": str(client_meta.get("cliente", "") or "").strip(),
                "cliente_nome": str(client_meta.get("cliente_nome", "") or "").strip(),
                "cliente_display": str(client_meta.get("cliente_display", "") or "").strip(),
                "estado": estado,
                "data_entrega": str(enc.get("data_entrega", "") or "").strip(),
                "tempo_plan_min": round(tempo_plan, 1),
            }
        )

    active.sort(key=lambda row: (row.get("data", ""), row.get("inicio", ""), row.get("encomenda", "")))
    week_active.sort(key=lambda row: (row.get("data", ""), row.get("inicio", ""), row.get("encomenda", "")))
    history.sort(key=lambda row: (row.get("data", ""), row.get("inicio", ""), row.get("encomenda", "")), reverse=True)
    backlog.sort(key=lambda row: (row.get("data_entrega", "") or "9999-99-99", row.get("numero", "")))

    month_token = f"{year or datetime.now().year}-{datetime.now().month:02d}"
    month_hist = [row for row in history if str(row.get("data", "")).startswith(month_token)]

    return {
        "summary": {
            "blocos_ativos": len(week_active),
            "backlog": len(backlog),
            "historico_mes": len(month_hist),
            "min_ativos": round(week_minutes, 1),
            "min_ativos_total": round(active_minutes, 1),
            "min_historico_mes": round(sum(max(0.0, desktop_pulse._parse_float(r.get("duracao_min", 0), 0)) for r in month_hist), 1),
            "week_start": week_start_dt.isoformat(),
            "week_end": week_end_dt.isoformat(),
            "week_label": f"{week_start_dt.strftime('%d/%m')} - {week_end_dt.strftime('%d/%m')}",
            "operacao": op_txt,
        },
        "week_dates": [d.isoformat() for d in week_dates],
        "active": week_active[:160],
        "history": history[:220],
        "backlog": backlog[:220],
    }


def build_planning_pdf(year: str | None = None, week_start: str | None = None, operation: str | None = None) -> Path:
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        from reportlab.pdfgen import canvas as pdf_canvas
    except Exception as exc:
        raise RuntimeError("ReportLab nao disponivel para gerar PDF") from exc

    overview = get_planning_overview(year=year, week_start=week_start, operation=operation)
    summary = dict(overview.get("summary", {}) or {})
    active = list(overview.get("active", []) or [])
    backlog = list(overview.get("backlog", []) or [])
    history = list(overview.get("history", []) or [])
    op_txt = str(summary.get("operacao", "") or operation or "Corte Laser").strip() or "Corte Laser"
    op_token = desktop_pulse._norm_text(op_txt).replace(" ", "_") or "corte_laser"
    path = Path(tempfile.gettempdir()) / f"lugest_impulse_planeamento_{op_token}_{str(year or datetime.now().year)}.pdf"

    c = pdf_canvas.Canvas(str(path), pagesize=landscape(A4))
    width, height = landscape(A4)
    margin = 28
    navy = colors.HexColor("#081D5C")
    steel = colors.HexColor("#223548")
    line = colors.HexColor("#D8E0EC")
    orange = colors.HexColor("#F59E0B")
    green = colors.HexColor("#16A34A")

    font_regular = "Helvetica"
    font_bold = "Helvetica-Bold"
    for name, file_name in (("SegoeUI", "segoeui.ttf"), ("SegoeUI-Bold", "segoeuib.ttf")):
        f_path = Path(r"C:\Windows\Fonts") / file_name
        if f_path.exists():
            pdfmetrics.registerFont(TTFont(name, str(f_path)))
            if "Bold" in name:
                font_bold = name
            else:
                font_regular = name

    def set_font(bold: bool, size: int) -> None:
        c.setFont(font_bold if bold else font_regular, size)

    def page_header(title: str, subtitle: str) -> float:
        c.setFillColor(navy)
        c.roundRect(margin, height - margin - 58, width - margin * 2, 58, 18, stroke=0, fill=1)
        c.setFillColor(colors.white)
        set_font(True, 18)
        c.drawString(margin + 16, height - margin - 22, title)
        set_font(False, 9)
        c.drawString(margin + 16, height - margin - 38, subtitle)
        return height - margin - 76

    def summary_band(y: float) -> float:
        boxes = [
            ("Blocos ativos", str(summary.get("blocos_ativos", 0)), navy),
            ("Backlog", str(summary.get("backlog", 0)), orange),
            ("Historico mes", str(summary.get("historico_mes", 0)), green),
            ("Min ativos", f"{summary.get('min_ativos', 0)}m", steel),
        ]
        box_w = (width - margin * 2 - 18) / 4
        x = margin
        for label, value, color in boxes:
            c.setFillColor(colors.white)
            c.setStrokeColor(line)
            c.roundRect(x, y - 46, box_w, 42, 14, stroke=1, fill=1)
            c.setFillColor(color)
            set_font(True, 9)
            c.drawString(x + 10, y - 16, label)
            set_font(True, 15)
            c.drawString(x + 10, y - 34, value)
            x += box_w + 6
        return y - 58

    def section_table(y: float, title: str, columns: list[str], rows: list[list[str]], min_y: float = 48) -> float:
        needed = 30 + 18 + (len(rows) * 16) + 16
        if y - needed < min_y:
            c.showPage()
            y = page_header(f"Plano {op_txt}", f"Ano {year or datetime.now().year} | Exportacao desktop")
        c.setFillColor(colors.white)
        c.setStrokeColor(line)
        box_h = 24 + 18 + max(1, len(rows)) * 16 + 12
        c.roundRect(margin, y - box_h, width - margin * 2, box_h, 18, stroke=1, fill=1)
        c.setFillColor(steel)
        set_font(True, 11)
        c.drawString(margin + 12, y - 16, title)
        col_ws = [0.18, 0.16, 0.14, 0.14, 0.14, 0.24]
        col_x = margin + 12
        c.setFillColor(navy)
        c.roundRect(margin + 10, y - 40, width - margin * 2 - 20, 18, 8, stroke=0, fill=1)
        c.setFillColor(colors.white)
        set_font(True, 8)
        x_positions: list[float] = []
        total_w = width - margin * 2 - 24
        for idx, col in enumerate(columns):
            x_positions.append(col_x)
            c.drawString(col_x + 4, y - 28, col)
            col_x += total_w * col_ws[idx]
        c.setFillColor(steel)
        set_font(False, 8)
        row_y = y - 54
        for row in rows:
            if row_y < min_y:
                break
            for idx, value in enumerate(row[: len(columns)]):
                c.drawString(x_positions[idx] + 4, row_y, str(value)[:34])
            c.setStrokeColor(line)
            c.line(margin + 12, row_y - 4, width - margin - 12, row_y - 4)
            row_y -= 16
        return y - box_h - 10

    y = page_header(f"Plano {op_txt}", f"Ano {year or datetime.now().year} | Exportacao desktop")
    y = summary_band(y)
    active_rows = [
        [
            str(row.get("encomenda", "") or "-"),
            str(row.get("data", "") or "-"),
            str(row.get("inicio", "") or "-"),
            str(row.get("material", "") or "-"),
            str(row.get("espessura", "") or "-"),
            f"{row.get('duracao_min', 0)}m",
        ]
        for row in active[:18]
    ]
    history_rows = [
        [
            str(row.get("encomenda", "") or "-"),
            str(row.get("data", "") or "-"),
            str(row.get("inicio", "") or "-"),
            str(row.get("material", "") or "-"),
            str(row.get("espessura", "") or "-"),
            f"{row.get('tempo_real_min', 0)} / {row.get('duracao_min', 0)}m",
        ]
        for row in history[:18]
    ]
    backlog_rows = [
        [
            str(row.get("numero", "") or "-"),
            str(row.get("cliente", "") or "-"),
            str(row.get("estado", "") or "-"),
            str(row.get("data_entrega", "") or "-"),
            f"{row.get('tempo_plan_min', 0)}m",
            "",
        ]
        for row in backlog[:18]
    ]
    y = section_table(y, "Blocos ativos", ["Encomenda", "Data", "Inicio", "Material", "Esp.", "Duracao"], active_rows)
    y = section_table(y, "Backlog", ["Encomenda", "Cliente", "Estado", "Entrega", "Plano", ""], backlog_rows)
    y = section_table(y, "Historico", ["Encomenda", "Data", "Inicio", "Material", "Esp.", "Real / Plano"], history_rows)
    c.save()
    return path
