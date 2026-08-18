from __future__ import annotations

import hmac
import base64
import binascii
import hashlib
import json
import os
from datetime import date, datetime
from decimal import Decimal
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, urlparse

import pymysql
from pymysql.cursors import DictCursor

from .auth import DeviceRegistry, RateLimiter
from .attachments import AttachmentStore, validate_job_id


ROOT = Path(__file__).resolve().parents[1]


def _load_env(path: Path) -> None:
    if not path.exists():
        return
    for raw_line in path.read_text(encoding="utf-8-sig").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key:
            os.environ.setdefault(key, value)


_load_env(ROOT / "lugest.env")


def _device_store_path() -> Path:
    configured = str(os.environ.get("LUGEST_MOBILE_DEVICE_STORE", "") or "").strip()
    return (Path(configured) if configured else ROOT / "mobile_gateway_data" / "devices.json").resolve()


DEVICE_REGISTRY = DeviceRegistry(_device_store_path())
ATTACHMENT_STORE = AttachmentStore(ROOT / "mobile_gateway_data" / "attachments")
RATE_LIMITER = RateLimiter(
    requests=int(os.environ.get("LUGEST_MOBILE_RATE_LIMIT", "180") or 180),
    window_seconds=60,
)
AUTH_FAILURE_LIMITER = RateLimiter(requests=60, window_seconds=60)


def _required_env(name: str) -> str:
    value = str(os.environ.get(name, "") or "").strip()
    if not value:
        raise RuntimeError(f"Falta definir {name} no lugest.env.")
    return value


def _database_connection():
    mobile_user = str(os.environ.get("LUGEST_MOBILE_DB_USER", "") or "").strip()
    mobile_password = str(os.environ.get("LUGEST_MOBILE_DB_PASS", "") or "").strip()
    allow_desktop_credentials = str(
        os.environ.get("LUGEST_MOBILE_ALLOW_DESKTOP_DB_CREDENTIALS", "1") or "1"
    ).strip().lower() in {"1", "true", "yes"}
    if mobile_user and mobile_password:
        user = mobile_user
        password = mobile_password
    elif allow_desktop_credentials:
        user = _required_env("LUGEST_DB_USER")
        password = _required_env("LUGEST_DB_PASS")
    else:
        raise RuntimeError("As credenciais restritas do gateway móvel não estão configuradas.")
    return pymysql.connect(
        host=_required_env("LUGEST_DB_HOST"),
        port=int(os.environ.get("LUGEST_DB_PORT", "3306")),
        user=user,
        password=password,
        database=_required_env("LUGEST_DB_NAME"),
        charset="utf8mb4",
        cursorclass=DictCursor,
        connect_timeout=5,
        read_timeout=10,
        write_timeout=10,
        autocommit=True,
    )


def _json_default(value: Any):
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, Decimal):
        return float(value)
    raise TypeError(f"Tipo não serializável: {type(value)!r}")


def _number(value: Any) -> float:
    try:
        return float(value or 0)
    except (TypeError, ValueError):
        return 0.0


class MobileGatewayHandler(BaseHTTPRequestHandler):
    server_version = "LuGEST-Mobile-Gateway/0.3"

    def do_GET(self) -> None:  # noqa: N802 - API do BaseHTTPRequestHandler
        try:
            if not self._authorized():
                self._send_auth_failure()
                return
            parsed = urlparse(self.path)
            query = parse_qs(parsed.query)
            routes = {
                "/health": self._health,
                "/v1/stock/summary": self._stock_summary,
                "/v1/stock/products": self._stock_products,
                "/v1/stock/materials": self._stock_materials,
            }
            handler = routes.get(parsed.path)
            if handler is None:
                self._send_json(HTTPStatus.NOT_FOUND, {"error": "Rota não encontrada."})
                return
            self._send_json(HTTPStatus.OK, handler(query))
        except (pymysql.MySQLError, RuntimeError):
            self._send_json(
                HTTPStatus.SERVICE_UNAVAILABLE,
                {"error": "A base de dados LuGEST não está disponível."},
            )
        except Exception:
            self._send_json(
                HTTPStatus.INTERNAL_SERVER_ERROR,
                {"error": "Ocorreu um erro interno no gateway móvel."},
            )

    def do_POST(self) -> None:  # noqa: N802 - API do BaseHTTPRequestHandler
        try:
            if not self._authorized():
                self._send_auth_failure()
                return
            parsed = urlparse(self.path)
            routes = {
                "/v1/service-jobs/sync": self._sync_service_jobs,
                "/v1/service-jobs/attachments": self._save_attachment,
            }
            handler = routes.get(parsed.path)
            if handler is None:
                self._send_json(HTTPStatus.NOT_FOUND, {"error": "Rota não encontrada."})
                return
            payload = self._read_json(max_bytes=15 * 1024 * 1024)
            self._send_json(HTTPStatus.OK, handler(payload))
        except ValueError as exc:
            self._send_json(HTTPStatus.BAD_REQUEST, {"error": str(exc)})
        except (pymysql.MySQLError, RuntimeError):
            self._send_json(
                HTTPStatus.SERVICE_UNAVAILABLE,
                {"error": "A base de dados LuGEST não está disponível."},
            )
        except Exception:
            self._send_json(
                HTTPStatus.INTERNAL_SERVER_ERROR,
                {"error": "Ocorreu um erro interno no gateway móvel."},
            )

    def _authorized(self) -> bool:
        supplied = self.headers.get("Authorization", "")
        if supplied.startswith("Bearer "):
            supplied = supplied[7:]
        self.device = DEVICE_REGISTRY.authenticate(supplied)
        self._auth_status = HTTPStatus.UNAUTHORIZED
        if self.device is None:
            if not AUTH_FAILURE_LIMITER.allow(self._client_key()):
                self._auth_status = HTTPStatus.TOO_MANY_REQUESTS
                return False
            allow_legacy = str(
                os.environ.get("LUGEST_MOBILE_ALLOW_LEGACY_TOKEN", "1") or "1"
            ).strip().lower() in {"1", "true", "yes"}
            expected = str(os.environ.get("LUGEST_MOBILE_API_TOKEN", "") or "").strip()
            if not allow_legacy or not expected or not hmac.compare_digest(
                supplied.encode(), expected.encode()
            ):
                return False
            self.device = {"id": "legacy", "name": "Ligação de testes"}
        client_key = self._client_key()
        if not RATE_LIMITER.allow(f"{self.device['id']}:{client_key}"):
            self._auth_status = HTTPStatus.TOO_MANY_REQUESTS
            return False
        return True

    def _send_auth_failure(self) -> None:
        status = getattr(self, "_auth_status", HTTPStatus.UNAUTHORIZED)
        if status == HTTPStatus.TOO_MANY_REQUESTS:
            self._send_json(status, {"error": "Demasiados pedidos. Tenta novamente dentro de um minuto."})
        else:
            self._send_json(status, {"error": "Chave de acesso inválida ou dispositivo revogado."})

    def _client_key(self) -> str:
        remote = str(self.client_address[0] or "")
        if remote in {"127.0.0.1", "::1"}:
            forwarded = str(self.headers.get("CF-Connecting-IP", "") or "").strip()
            if forwarded:
                return forwarded[:64]
        return remote[:64]

    def _health(self, _query: dict[str, list[str]]) -> dict[str, Any]:
        with _database_connection() as connection:
            with connection.cursor() as cursor:
                cursor.execute("SELECT 1 AS ok")
                cursor.fetchone()
        return {
            "ok": True,
            "service": "lugest-mobile-gateway",
            "version": "0.3.0",
            "device": str(getattr(self, "device", {}).get("name", "")),
            "server_time": datetime.now().astimezone().isoformat(),
        }

    def _stock_summary(self, _query: dict[str, list[str]]) -> dict[str, Any]:
        with _database_connection() as connection:
            with connection.cursor() as cursor:
                cursor.execute(
                    """
                    SELECT
                      (SELECT COUNT(*) FROM produtos) AS products,
                      (SELECT COUNT(*) FROM materiais) AS materials,
                      (SELECT COUNT(*) FROM produtos
                         WHERE COALESCE(qty, 0) <= 0
                            OR (COALESCE(alerta, 0) > 0 AND COALESCE(qty, 0) <= alerta))
                      +
                      (SELECT COUNT(*) FROM materiais
                         WHERE COALESCE(quantidade, 0) - COALESCE(reservado, 0) <= 0)
                      AS critical,
                      GREATEST(
                        COALESCE((SELECT MAX(atualizado_em) FROM produtos), '1970-01-01'),
                        COALESCE((SELECT MAX(atualizado_em) FROM materiais), '1970-01-01')
                      ) AS updated_at
                    """
                )
                row = cursor.fetchone() or {}
        return {
            "products": int(row.get("products") or 0),
            "materials": int(row.get("materials") or 0),
            "critical": int(row.get("critical") or 0),
            "updated_at": row.get("updated_at"),
        }

    def _stock_products(self, query: dict[str, list[str]]) -> dict[str, Any]:
        term = self._query_text(query)
        in_stock = self._in_stock_only(query)
        clauses: list[str] = []
        params: list[Any] = []
        if term:
            like = f"%{term}%"
            clauses.append(
                "(codigo LIKE %s OR descricao LIKE %s OR categoria LIKE %s "
                "OR subcat LIKE %s OR tipo LIKE %s)"
            )
            params.extend([like] * 5)
        if in_stock:
            clauses.append("COALESCE(qty, 0) > 0")
        where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
        with _database_connection() as connection:
            with connection.cursor() as cursor:
                cursor.execute(
                    f"""
                    SELECT codigo, descricao, categoria, subcat, tipo, unid,
                           qty, alerta, p_compra, atualizado_em
                      FROM produtos
                      {where}
                     ORDER BY descricao, codigo
                     LIMIT 300
                    """,
                    params,
                )
                rows = cursor.fetchall()
        items = [
            {
                "kind": "product",
                "code": str(row.get("codigo") or ""),
                "description": str(row.get("descricao") or row.get("codigo") or ""),
                "available": _number(row.get("qty")),
                "unit": str(row.get("unid") or "UN"),
                "sale_price": _number(row.get("p_compra")),
                "location": "",
                "category": " / ".join(
                    str(row.get(field) or "").strip()
                    for field in ("categoria", "subcat", "tipo")
                    if str(row.get(field) or "").strip()
                ),
                "alert_level": _number(row.get("alerta")),
                "updated_at": row.get("atualizado_em"),
            }
            for row in rows
        ]
        return {"items": items, "count": len(items)}

    def _stock_materials(self, query: dict[str, list[str]]) -> dict[str, Any]:
        term = self._query_text(query)
        in_stock = self._in_stock_only(query)
        clauses: list[str] = []
        params: list[Any] = []
        if term:
            like = f"%{term}%"
            clauses.append(
                "(id LIKE %s OR lote_interno LIKE %s OR lote_fornecedor LIKE %s "
                "OR formato LIKE %s OR material LIKE %s OR localizacao LIKE %s)"
            )
            params.extend([like] * 6)
        if in_stock:
            clauses.append("COALESCE(quantidade, 0) - COALESCE(reservado, 0) > 0")
        where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
        with _database_connection() as connection:
            with connection.cursor() as cursor:
                cursor.execute(
                    f"""
                    SELECT id, lote_interno, lote_fornecedor, formato, material,
                           espessura, comprimento, largura, metros, quantidade,
                           reservado, localizacao, preco_unid, atualizado_em
                      FROM materiais
                      {where}
                     ORDER BY material, formato, espessura, id
                     LIMIT 300
                    """,
                    params,
                )
                rows = cursor.fetchall()
        items = []
        for row in rows:
            details = [str(row.get("formato") or "").strip(), str(row.get("material") or "").strip()]
            if str(row.get("espessura") or "").strip():
                details.append(f"{row['espessura']} mm")
            if _number(row.get("comprimento")) and _number(row.get("largura")):
                details.append(f"{_number(row['comprimento']):g}×{_number(row['largura']):g} mm")
            elif _number(row.get("metros")):
                details.append(f"{_number(row['metros']):g} m")
            available = _number(row.get("quantidade")) - _number(row.get("reservado"))
            items.append(
                {
                    "kind": "material",
                    "code": str(row.get("id") or ""),
                    "description": " · ".join(value for value in details if value),
                    "available": available,
                    "unit": "UN",
                    "sale_price": _number(row.get("preco_unid")),
                    "location": str(row.get("localizacao") or ""),
                    "category": str(row.get("formato") or ""),
                    "alert_level": 0,
                    "updated_at": row.get("atualizado_em"),
                    "lot": str(row.get("lote_interno") or row.get("lote_fornecedor") or ""),
                }
            )
        return {"items": items, "count": len(items)}

    def _sync_service_jobs(self, payload: dict[str, Any]) -> dict[str, Any]:
        raw_jobs = payload.get("jobs", [])
        if not isinstance(raw_jobs, list):
            raise ValueError("O campo jobs tem de ser uma lista.")
        if len(raw_jobs) > 250:
            raise ValueError("Foram enviados demasiados serviços num só pedido.")

        results: list[dict[str, Any]] = []
        with _database_connection() as connection:
            connection.begin()
            try:
                with connection.cursor() as cursor:
                    for raw in raw_jobs:
                        if not isinstance(raw, dict):
                            raise ValueError("Foi recebido um serviço inválido.")
                        job_id = validate_job_id(raw.get("id"))
                        client_name = str(raw.get("clientName") or "").strip()[:150]
                        title = str(raw.get("title") or "").strip()
                        if not job_id or not title or not client_name:
                            raise ValueError("Cada serviço exige número, título e cliente.")

                        cursor.execute(
                            "SELECT estado FROM servicos_diretos WHERE numero=%s FOR UPDATE",
                            (job_id,),
                        )
                        existing = cursor.fetchone() or {}
                        existing_state = str(existing.get("estado") or "")
                        if existing_state in {"Confirmado", "Faturado", "Anulado"}:
                            results.append(
                                {"id": job_id, "state": existing_state, "locked": True}
                            )
                            continue

                        cursor.execute(
                            "SELECT codigo FROM clientes WHERE nome=%s ORDER BY codigo LIMIT 1",
                            (client_name,),
                        )
                        client = cursor.fetchone() or {}
                        client_code = str(client.get("codigo") or "")

                        vat_rate = max(0.0, min(100.0, _number(raw.get("vatRate"))))
                        lines = []
                        subtotal = 0.0
                        for index, line in enumerate(raw.get("lines") or [], start=1):
                            if not isinstance(line, dict):
                                continue
                            description = str(line.get("description") or "").strip()
                            quantity = _number(line.get("quantity"))
                            unit_price = _number(line.get("unitPrice"))
                            if not description or quantity <= 0:
                                continue
                            kind = str(line.get("kind") or "service").strip().lower()
                            if kind not in {"service", "product", "material", "set"}:
                                kind = "service"
                            base = round(quantity * unit_price, 2)
                            tax = round(base * vat_rate / 100.0, 2)
                            subtotal += base
                            lines.append(
                                {
                                    "id": f"MOB-{index:03d}",
                                    "kind": kind,
                                    "ref": str(line.get("ref") or "")[:80],
                                    "description": description[:255],
                                    "qty": round(quantity, 3),
                                    "unit": str(line.get("unit") or "UN")[:20],
                                    "unit_price": round(unit_price, 4),
                                    "iva_perc": round(vat_rate, 2),
                                    "subtotal": base,
                                    "valor_iva": tax,
                                    "total": round(base + tax, 2),
                                }
                            )
                        subtotal = round(subtotal, 2)
                        vat_amount = round(subtotal * vat_rate / 100.0, 2)
                        total = round(subtotal + vat_amount, 2)

                        notes = [str(raw.get("notes") or "").strip()]
                        phone = str(raw.get("clientPhone") or "").strip()
                        if phone:
                            notes.append(f"Contacto: {phone}")
                        photo_count = len(raw.get("photoPaths") or [])
                        if photo_count:
                            notes.append(f"LuGEST Field: {photo_count} fotografia(s) no dispositivo.")
                        if str(raw.get("signaturePath") or "").strip():
                            notes.append("LuGEST Field: assinatura do cliente recolhida.")
                        mobile_status = str(raw.get("status") or "scheduled")
                        notes.append(f"Estado móvel: {mobile_status}")

                        scheduled = str(raw.get("scheduledAt") or "")[:10]
                        if len(scheduled) != 10:
                            scheduled = date.today().isoformat()
                        now = datetime.now()
                        cursor.execute(
                            """
                            INSERT INTO servicos_diretos (
                              numero, cliente_codigo, cliente_nome, data_servico,
                              data_vencimento, estado, local_servico, responsavel,
                              obs, subtotal, valor_iva, total, stock_consumido,
                              created_at, updated_at, linhas_json
                            ) VALUES (
                              %s, %s, %s, %s, DATE_ADD(%s, INTERVAL 30 DAY),
                              'Rascunho', %s, 'LuGEST Field', %s, %s, %s, %s,
                              0, %s, %s, %s
                            )
                            ON DUPLICATE KEY UPDATE
                              cliente_codigo=VALUES(cliente_codigo),
                              cliente_nome=VALUES(cliente_nome),
                              data_servico=VALUES(data_servico),
                              data_vencimento=VALUES(data_vencimento),
                              local_servico=VALUES(local_servico),
                              responsavel=VALUES(responsavel),
                              obs=VALUES(obs), subtotal=VALUES(subtotal),
                              valor_iva=VALUES(valor_iva), total=VALUES(total),
                              updated_at=VALUES(updated_at), linhas_json=VALUES(linhas_json)
                            """,
                            (
                                job_id,
                                client_code or None,
                                client_name,
                                scheduled,
                                scheduled,
                                str(raw.get("address") or "")[:255],
                                "\n".join(note for note in notes if note),
                                subtotal,
                                vat_amount,
                                total,
                                now,
                                now,
                                json.dumps(lines, ensure_ascii=False),
                            ),
                        )
                        results.append(
                            {
                                "id": job_id,
                                "state": "Rascunho",
                                "locked": False,
                                "client_matched": bool(client_code),
                            }
                        )
                connection.commit()
            except Exception:
                connection.rollback()
                raise
        return {"synced": len(results), "jobs": results}

    def _save_attachment(self, payload: dict[str, Any]) -> dict[str, Any]:
        job_id = validate_job_id(payload.get("jobId"))
        encoded = str(payload.get("dataBase64") or "")
        if not encoded:
            raise ValueError("O conteúdo do anexo está vazio.")
        try:
            data = base64.b64decode(encoded, validate=True)
        except (binascii.Error, ValueError) as exc:
            raise ValueError("O conteúdo do anexo é inválido.") from exc
        if not data or len(data) > 10 * 1024 * 1024:
            raise ValueError("O anexo está vazio ou excede 10 MB.")
        digest = str(payload.get("sha256") or "").strip().lower()
        if not hmac.compare_digest(hashlib.sha256(data).hexdigest(), digest):
            raise ValueError("A verificação de integridade do anexo falhou.")

        with _database_connection() as connection:
            with connection.cursor() as cursor:
                cursor.execute(
                    "SELECT numero FROM servicos_diretos WHERE numero=%s LIMIT 1",
                    (job_id,),
                )
                if cursor.fetchone() is None:
                    raise ValueError("Sincroniza primeiro o serviço antes de enviar anexos.")

        stored = ATTACHMENT_STORE.save(
            job_id=job_id,
            kind=str(payload.get("kind") or ""),
            original_name=str(payload.get("fileName") or "anexo"),
            declared_mime=str(payload.get("mimeType") or "").lower(),
            digest=digest,
            data=data,
            device_id=str(getattr(self, "device", {}).get("id", "")),
        )
        return {"attachment": stored}

    def _read_json(self, *, max_bytes: int = 4 * 1024 * 1024) -> dict[str, Any]:
        try:
            length = int(self.headers.get("Content-Length", "0") or 0)
        except ValueError as exc:
            raise ValueError("Content-Length inválido.") from exc
        if length <= 0 or length > max_bytes:
            raise ValueError("Pedido vazio ou demasiado grande.")
        try:
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise ValueError("JSON inválido.") from exc
        if not isinstance(payload, dict):
            raise ValueError("O corpo do pedido tem de ser um objeto JSON.")
        return payload

    @staticmethod
    def _query_text(query: dict[str, list[str]]) -> str:
        return str((query.get("q") or [""])[0]).strip()[:120]

    @staticmethod
    def _in_stock_only(query: dict[str, list[str]]) -> bool:
        return str((query.get("in_stock") or ["1"])[0]).lower() not in {"0", "false", "no"}

    def _send_json(self, status: HTTPStatus, payload: dict[str, Any]) -> None:
        body = json.dumps(payload, ensure_ascii=False, default=_json_default).encode("utf-8")
        self.send_response(int(status))
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("X-Frame-Options", "DENY")
        self.send_header("Referrer-Policy", "no-referrer")
        if str(self.headers.get("X-Forwarded-Proto", "") or "").lower() == "https":
            self.send_header("Strict-Transport-Security", "max-age=31536000")
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, fmt: str, *args: Any) -> None:
        print(f"[{self.log_date_time_string()}] {self.client_address[0]} {fmt % args}")


def run() -> None:
    _required_env("LUGEST_MOBILE_API_TOKEN")
    host = str(os.environ.get("LUGEST_MOBILE_API_BIND", "127.0.0.1") or "127.0.0.1")
    port = int(os.environ.get("LUGEST_MOBILE_API_PORT", "8765") or 8765)
    server = ThreadingHTTPServer((host, port), MobileGatewayHandler)
    print(f"LuGEST Mobile Gateway ativo em {host}:{port}")
    print("Prima Ctrl+C para terminar.")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    run()
