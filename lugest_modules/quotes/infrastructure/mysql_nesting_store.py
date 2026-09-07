"""SQL mirror for quote nesting studies, using an injected connection factory."""
from __future__ import annotations
import json
from typing import Any, Callable
from lugest_modules.quotes.application.nesting_studies import json_clone

class MysqlNestingStudyStore:
    def __init__(self, connect: Callable[[], Any] | None):
        self.connect = connect

    def ensure_table(self, conn: Any) -> None:
        with conn.cursor() as cur:
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS orc_nesting_studies (
                    quote_number VARCHAR(80) NOT NULL,
                    group_key VARCHAR(190) NOT NULL,
                    group_label VARCHAR(255) NULL,
                    study_json LONGTEXT NULL,
                    created_at DATETIME NULL,
                    updated_at DATETIME NULL,
                    PRIMARY KEY (quote_number, group_key)
                )
                """
            )

    def studies(self, numero: str) -> dict[str, Any]:
        numero_txt = str(numero or "").strip()
        if not numero_txt:
            return {}
        conn = None
        studies: dict[str, Any] = {}
        try:
            connect = self.connect
            if not callable(connect):
                return {}
            conn = connect()
            self.ensure_table(conn)
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT group_key, study_json
                    FROM orc_nesting_studies
                    WHERE quote_number=%s
                    ORDER BY updated_at DESC, group_key ASC
                    """,
                    (numero_txt,),
                )
                rows = list(cur.fetchall() or [])
            for row in rows:
                group_key = str((row.get("group_key") if isinstance(row, dict) else row[0]) or "").strip()
                raw = row.get("study_json") if isinstance(row, dict) else row[1]
                if not group_key:
                    continue
                if isinstance(raw, (bytes, bytearray)):
                    raw = raw.decode("utf-8", errors="ignore")
                try:
                    parsed = json.loads(str(raw or "{}"))
                except Exception:
                    parsed = {}
                if isinstance(parsed, dict):
                    studies[group_key] = parsed
        except Exception:
            studies = {}
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass
        return studies

    def save(self, numero: str, group_key: str, group_label: str, payload: dict[str, Any]) -> None:
        numero_txt = str(numero or "").strip()
        group_key_txt = str(group_key or "").strip()
        if not numero_txt or not group_key_txt:
            return
        conn = None
        try:
            connect = self.connect
            if not callable(connect):
                return
            conn = connect()
            self.ensure_table(conn)
            clean = json.dumps(json_clone(payload), ensure_ascii=False)
            with conn.cursor() as cur:
                cur.execute(
                    """
                    INSERT INTO orc_nesting_studies (
                        quote_number,
                        group_key,
                        group_label,
                        study_json,
                        created_at,
                        updated_at
                    )
                    VALUES (%s, %s, %s, %s, NOW(), NOW())
                    ON DUPLICATE KEY UPDATE
                        group_label=VALUES(group_label),
                        study_json=VALUES(study_json),
                        updated_at=VALUES(updated_at)
                    """,
                    (numero_txt, group_key_txt, str(group_label or "").strip(), clean),
                )
            conn.commit()
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass

    def delete(self, numero: str, group_key: str = "") -> None:
        numero_txt = str(numero or "").strip()
        group_key_txt = str(group_key or "").strip()
        if not numero_txt:
            return
        conn = None
        try:
            connect = self.connect
            if not callable(connect):
                return
            conn = connect()
            self.ensure_table(conn)
            with conn.cursor() as cur:
                if group_key_txt:
                    cur.execute(
                        "DELETE FROM orc_nesting_studies WHERE quote_number=%s AND group_key=%s",
                        (numero_txt, group_key_txt),
                    )
                else:
                    cur.execute("DELETE FROM orc_nesting_studies WHERE quote_number=%s", (numero_txt,))
            conn.commit()
        finally:
            try:
                if conn:
                    conn.close()
            except Exception:
                pass
