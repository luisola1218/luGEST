from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app.config import settings


def _check_mysql_connectivity() -> tuple[bool, str]:
    try:
        import pymysql
    except Exception as exc:
        return False, f"PyMySQL indisponivel: {exc}"
    try:
        conn = pymysql.connect(
            host=settings.db_host,
            port=settings.db_port,
            user=settings.db_user,
            password=settings.db_pass,
            database=settings.db_name,
            charset="utf8mb4",
            connect_timeout=5,
            read_timeout=10,
            write_timeout=10,
            autocommit=True,
        )
        try:
            with conn.cursor() as cur:
                cur.execute("SELECT 1")
                cur.fetchone()
        finally:
            conn.close()
        return True, f"MySQL ok: {settings.db_host}:{settings.db_port}/{settings.db_name}"
    except Exception as exc:
        return False, f"MySQL falhou: {exc}"


def main() -> int:
    issues = settings.runtime_errors()
    if issues:
        print("preflight-config-error")
        for issue in issues:
            print(f"- {issue}")
        return 1

    ok, message = _check_mysql_connectivity()
    if not ok:
        print("preflight-mysql-error")
        print(f"- {message}")
        return 2

    print("preflight-ok")
    print(message)
    print(f"API bind: {settings.api_host}:{settings.api_port}")
    print(f"CORS origins: {len(settings.api_allowed_origins)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
