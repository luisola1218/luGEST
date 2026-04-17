from __future__ import annotations

from typing import Annotated

from fastapi import Depends, FastAPI, Header, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel

from .config import settings
from .security import decode_token, issue_token
from .services.pulse_runtime import (
    authenticate,
    build_material_separation_pdf,
    build_planning_pdf,
    clear_plan_delay_reason,
    ensure_trial_access,
    get_dashboard,
    get_mobile_avarias,
    get_encomenda,
    get_mobile_alerts,
    get_material_separation,
    get_operator_board,
    get_planning_overview,
    list_encomendas,
    mysql_connection_status,
    set_plan_delay_reason,
    set_material_separation_check,
    trial_status,
)


class LoginIn(BaseModel):
    username: str
    password: str


class MaterialSeparationCheckIn(BaseModel):
    check_key: str
    checked: bool


class PlanDelayReasonIn(BaseModel):
    item_key: str
    reason: str = ""


app = FastAPI(
    title="LUGEST Impulse Mobile API",
    version="0.1.0",
    docs_url="/docs",
    redoc_url="/redoc",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=list(settings.api_allowed_origins),
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


def require_auth(
    authorization: Annotated[str | None, Header()] = None,
    access_token: str | None = Query(default=None),
) -> dict:
    token = ""
    if authorization:
        parts = authorization.strip().split(" ", 1)
        token = parts[1] if len(parts) == 2 and parts[0].lower() == "bearer" else authorization.strip()
    if not token and access_token:
        token = str(access_token or "").strip()
    if not token:
        raise HTTPException(status_code=401, detail="Token em falta")
    try:
        payload = decode_token(token)
    except RuntimeError as exc:
        raise HTTPException(status_code=503, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=401, detail=str(exc)) from exc
    try:
        ensure_trial_access(str(payload.get("username", "") or ""))
    except Exception as exc:
        raise HTTPException(status_code=403, detail=str(exc)) from exc
    return payload


@app.get("/api/v1/health")
def health() -> dict:
    status = dict(trial_status() or {})
    db_status = dict(mysql_connection_status() or {})
    overall = "ok"
    if bool(status.get("blocking", False)):
        overall = "blocked"
    if not bool(db_status.get("ok", False)):
        overall = "degraded"
    return {
        "status": overall,
        "trial": {
            "enabled": bool(status.get("enabled", False)),
            "state": str(status.get("state", "") or "").strip(),
            "company_name": str(status.get("company_name", "") or "").strip(),
            "days_remaining": status.get("days_remaining"),
            "message": str(status.get("message", "") or "").strip(),
            "owner_configured": bool(status.get("owner_configured", False)),
        },
        "security": {
            "api_secret_configured": bool(settings.api_secret_configured()),
            "cors_origins": list(settings.api_allowed_origins),
        },
        "database": db_status,
    }


@app.post("/api/v1/auth/login")
def login(payload: LoginIn) -> dict:
    try:
        auth = authenticate(payload.username, payload.password)
    except ValueError as exc:
        raise HTTPException(status_code=403, detail=str(exc)) from exc
    if not auth:
        raise HTTPException(status_code=401, detail="Credenciais invalidas")
    try:
        token = issue_token(auth["username"], auth["role"])
    except RuntimeError as exc:
        raise HTTPException(status_code=503, detail=str(exc)) from exc
    return {"token": token, "user": auth}


@app.get("/api/v1/pulse/dashboard")
def pulse_dashboard(
    _user: dict = Depends(require_auth),
    period: str = Query(default="7 dias"),
    year: str | None = Query(default=None),
    encomenda: str = Query(default="Todas"),
    visao: str = Query(default="Todas"),
    origem: str = Query(default="Ambos"),
) -> dict:
    return get_dashboard(period=period, year=year, encomenda=encomenda, visao=visao, origem=origem)


@app.get("/api/v1/pulse/encomendas")
def pulse_encomendas(
    _user: dict = Depends(require_auth),
    year: str | None = Query(default=None),
) -> dict:
    return {"items": list_encomendas(year=year)}


@app.get("/api/v1/pulse/encomendas/{numero}")
def pulse_encomenda_detail(
    numero: str,
    _user: dict = Depends(require_auth),
) -> dict:
    enc = get_encomenda(numero)
    if not enc:
        raise HTTPException(status_code=404, detail="Encomenda nao encontrada")
    return {"item": enc}


@app.get("/api/v1/mobile/operator-board")
def mobile_operator_board(
    _user: dict = Depends(require_auth),
    year: str | None = Query(default=None),
) -> dict:
    return get_operator_board(year=year, username=str(_user.get("username", "") or ""), role=str(_user.get("role", "") or ""))


@app.get("/api/v1/mobile/alerts")
def mobile_alerts(
    _user: dict = Depends(require_auth),
) -> dict:
    return get_mobile_alerts()


@app.get("/api/v1/mobile/avarias")
def mobile_avarias(
    _user: dict = Depends(require_auth),
) -> dict:
    return get_mobile_avarias()


@app.get("/api/v1/mobile/planning")
def mobile_planning(
    _user: dict = Depends(require_auth),
    year: str | None = Query(default=None),
    week_start: str | None = Query(default=None),
) -> dict:
    return get_planning_overview(year=year, week_start=week_start)


@app.get("/api/v1/mobile/planning/pdf")
def mobile_planning_pdf(
    _user: dict = Depends(require_auth),
    year: str | None = Query(default=None),
    week_start: str | None = Query(default=None),
) -> FileResponse:
    path = build_planning_pdf(year=year, week_start=week_start)
    filename = f"lugest_planeamento_{str(year or 'atual').strip() or 'atual'}.pdf"
    return FileResponse(path=path, media_type="application/pdf", filename=filename)


@app.get("/api/v1/mobile/material-separation")
def mobile_material_separation(
    _user: dict = Depends(require_auth),
    horizon_days: int = Query(default=4),
) -> dict:
    return get_material_separation(horizon_days=horizon_days)


@app.post("/api/v1/mobile/material-separation/check")
def mobile_material_separation_check(
    payload: MaterialSeparationCheckIn,
    _user: dict = Depends(require_auth),
) -> dict:
    return set_material_separation_check(payload.check_key, payload.checked)


@app.get("/api/v1/mobile/material-separation/pdf")
def mobile_material_separation_pdf(
    _user: dict = Depends(require_auth),
    horizon_days: int = Query(default=4),
) -> FileResponse:
    path = build_material_separation_pdf(horizon_days=horizon_days)
    filename = f"lugest_separacao_mp_{int(horizon_days or 4)}d.pdf"
    return FileResponse(path=path, media_type="application/pdf", filename=filename)


@app.post("/api/v1/pulse/plan-delay/reason")
def pulse_plan_delay_reason(
    payload: PlanDelayReasonIn,
    _user: dict = Depends(require_auth),
) -> dict:
    return set_plan_delay_reason(payload.item_key, payload.reason)


@app.post("/api/v1/pulse/plan-delay/clear")
def pulse_plan_delay_clear(
    payload: PlanDelayReasonIn,
    _user: dict = Depends(require_auth),
) -> dict:
    return clear_plan_delay_reason(payload.item_key)
