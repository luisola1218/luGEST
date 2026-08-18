from __future__ import annotations

import json
import argparse
import base64
import hashlib
import os
import secrets
import shutil
import sqlite3
import sys
import threading
import urllib.error
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from mobile_gateway.server import (
    ATTACHMENT_STORE,
    MobileGatewayHandler,
    ThreadingHTTPServer,
    _database_connection,
)


def request(url: str, token: str, payload: dict | None = None) -> tuple[int, dict]:
    body = None if payload is None else json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=body,
        headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
            "Content-Type": "application/json",
        },
        method="POST" if body is not None else "GET",
    )
    try:
        with urllib.request.urlopen(req, timeout=10) as response:
            return response.status, json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        return exc.code, json.loads(exc.read().decode("utf-8"))


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--write-smoke", action="store_true")
    args = parser.parse_args()
    token = secrets.token_hex(32)
    test_job_id = f"SV-VERIFY-{secrets.token_hex(4).upper()}"
    os.environ["LUGEST_MOBILE_API_TOKEN"] = token
    server = ThreadingHTTPServer(("127.0.0.1", 0), MobileGatewayHandler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    base = f"http://127.0.0.1:{server.server_address[1]}"
    try:
        status, _ = request(f"{base}/health", "invalid")
        assert status == 401, status

        status, health = request(f"{base}/health", token)
        assert status == 200 and health.get("ok") is True, health

        status, summary = request(f"{base}/v1/stock/summary", token)
        assert status == 200 and int(summary.get("products", -1)) >= 0, summary

        status, products = request(f"{base}/v1/stock/products?in_stock=1", token)
        assert status == 200 and isinstance(products.get("items"), list), products

        status, materials = request(f"{base}/v1/stock/materials?in_stock=1", token)
        assert status == 200 and isinstance(materials.get("items"), list), materials
        status, rejected = request(
            f"{base}/v1/service-jobs/sync", token, {"jobs": "invalid"}
        )
        assert status == 400 and rejected.get("error"), rejected
        if args.write_smoke:
            status, synced = request(
                f"{base}/v1/service-jobs/sync",
                token,
                {
                    "jobs": [
                        {
                            "id": test_job_id,
                            "title": "Teste automático da integração móvel",
                            "clientName": "Cliente de teste API",
                            "clientPhone": "",
                            "address": "Teste local",
                            "scheduledAt": "2026-08-18T10:00:00",
                            "status": "readyToInvoice",
                            "notes": "Registo temporário removido pelo próprio teste.",
                            "vatRate": 23,
                            "lines": [
                                {
                                    "description": "Serviço de teste",
                                    "quantity": 1,
                                    "unit": "SV",
                                    "unitPrice": 10,
                                }
                            ],
                            "photoPaths": [],
                        }
                    ]
                },
            )
            assert status == 200 and synced.get("synced") == 1, synced
            png = b"\x89PNG\r\n\x1a\n" + b"LuGEST verification image"
            digest = hashlib.sha256(png).hexdigest()
            status, attachment = request(
                f"{base}/v1/service-jobs/attachments",
                token,
                {
                    "jobId": test_job_id,
                    "kind": "signature",
                    "fileName": "assinatura.png",
                    "mimeType": "image/png",
                    "sha256": digest,
                    "dataBase64": base64.b64encode(png).decode("ascii"),
                },
            )
            assert status == 200 and attachment["attachment"]["sha256"] == digest, attachment
            status, duplicate = request(
                f"{base}/v1/service-jobs/attachments",
                token,
                {
                    "jobId": test_job_id,
                    "kind": "signature",
                    "fileName": "assinatura.png",
                    "mimeType": "image/png",
                    "sha256": digest,
                    "dataBase64": base64.b64encode(png).decode("ascii"),
                },
            )
            assert status == 200 and duplicate["attachment"]["stored"] is True, duplicate
        print(
            "mobile-gateway-ok "
            f"products={len(products['items'])} materials={len(materials['items'])}"
        )
    finally:
        if args.write_smoke and test_job_id.startswith("SV-VERIFY-"):
            with _database_connection() as connection:
                with connection.cursor() as cursor:
                    cursor.execute(
                        "DELETE FROM servicos_diretos WHERE numero=%s",
                        (test_job_id,),
                    )
                connection.commit()
            attachment_directory = (ATTACHMENT_STORE.files_root / test_job_id).resolve()
            files_root = ATTACHMENT_STORE.files_root.resolve()
            if attachment_directory.parent == files_root and attachment_directory.exists():
                shutil.rmtree(attachment_directory)
            if ATTACHMENT_STORE.database_path.exists():
                with sqlite3.connect(ATTACHMENT_STORE.database_path, timeout=10) as connection:
                    connection.execute(
                        "DELETE FROM attachments WHERE job_id=?", (test_job_id,)
                    )
                    connection.commit()
        server.shutdown()
        server.server_close()
        thread.join(timeout=5)


if __name__ == "__main__":
    main()
