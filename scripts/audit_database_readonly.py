from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import main


CHECKS = {
    "billing_duplicate_quote_source": """
        SELECT COUNT(*) AS n FROM (
          SELECT orcamento_numero FROM faturacao_registos
          WHERE COALESCE(orcamento_numero, '') <> ''
          GROUP BY orcamento_numero HAVING COUNT(*) > 1
        ) x
    """,
    "billing_duplicate_order_source": """
        SELECT COUNT(*) AS n FROM (
          SELECT encomenda_numero FROM faturacao_registos
          WHERE COALESCE(encomenda_numero, '') <> ''
          GROUP BY encomenda_numero HAVING COUNT(*) > 1
        ) x
    """,
    "billing_duplicate_service_source": """
        SELECT COUNT(*) AS n FROM (
          SELECT servico_numero FROM faturacao_registos
          WHERE COALESCE(servico_numero, '') <> ''
          GROUP BY servico_numero HAVING COUNT(*) > 1
        ) x
    """,
    "invoice_duplicate_document_id": """
        SELECT COUNT(*) AS n FROM (
          SELECT documento_id FROM faturacao_faturas
          WHERE COALESCE(documento_id, '') <> ''
          GROUP BY documento_id HAVING COUNT(*) > 1
        ) x
    """,
    "invoice_duplicate_legal_number": """
        SELECT COUNT(*) AS n FROM (
          SELECT legal_invoice_no FROM faturacao_faturas
          WHERE COALESCE(legal_invoice_no, '') <> ''
          GROUP BY legal_invoice_no HAVING COUNT(*) > 1
        ) x
    """,
    "invoice_duplicate_series_sequence": """
        SELECT COUNT(*) AS n FROM (
          SELECT serie_id, seq_num FROM faturacao_faturas
          WHERE COALESCE(serie_id, '') <> '' AND COALESCE(seq_num, 0) > 0
          GROUP BY serie_id, seq_num HAVING COUNT(*) > 1
        ) x
    """,
    "payment_duplicate_id": """
        SELECT COUNT(*) AS n FROM (
          SELECT pagamento_id FROM faturacao_pagamentos
          WHERE COALESCE(pagamento_id, '') <> ''
          GROUP BY pagamento_id HAVING COUNT(*) > 1
        ) x
    """,
    "payment_orphan_invoice": """
        SELECT COUNT(*) AS n
        FROM faturacao_pagamentos p
        LEFT JOIN faturacao_faturas f
          ON f.documento_id = p.fatura_documento_id
         AND f.registo_numero = p.registo_numero
        WHERE COALESCE(p.fatura_documento_id, '') <> '' AND f.id IS NULL
    """,
    "service_orphan_billing_record": """
        SELECT COUNT(*) AS n
        FROM servicos_diretos s
        LEFT JOIN faturacao_registos r ON r.numero = s.faturacao_numero
        WHERE COALESCE(s.faturacao_numero, '') <> '' AND r.numero IS NULL
    """,
    "billing_orphan_service": """
        SELECT COUNT(*) AS n
        FROM faturacao_registos r
        LEFT JOIN servicos_diretos s ON s.numero = r.servico_numero
        WHERE COALESCE(r.servico_numero, '') <> '' AND s.numero IS NULL
    """,
    "negative_service_totals": """
        SELECT COUNT(*) AS n FROM servicos_diretos
        WHERE COALESCE(subtotal, 0) < 0 OR COALESCE(valor_iva, 0) < 0 OR COALESCE(total, 0) < 0
    """,
    "negative_invoice_or_payment_values": """
        SELECT
          (SELECT COUNT(*) FROM faturacao_faturas WHERE COALESCE(valor_total, 0) < 0) +
          (SELECT COUNT(*) FROM faturacao_pagamentos WHERE COALESCE(valor, 0) < 0) AS n
    """,
}


def _value(row) -> int:
    if isinstance(row, dict):
        for key, value in row.items():
            if str(key).lower() == "n":
                return int(value or 0)
    return int((row or [0])[0] or 0)


def audit() -> dict:
    connection = main._mysql_connect()
    try:
        with connection.cursor() as cursor:
            cursor.execute("SET TRANSACTION READ ONLY")
            cursor.execute("START TRANSACTION READ ONLY")
            results: dict[str, int] = {}
            for name, sql in CHECKS.items():
                cursor.execute(sql)
                results[name] = _value(cursor.fetchone())

            cursor.execute("SELECT linhas_json FROM servicos_diretos WHERE COALESCE(linhas_json, '') <> ''")
            invalid_json = 0
            for row in cursor.fetchall() or []:
                raw = row.get("linhas_json", "") if isinstance(row, dict) else row[0]
                try:
                    json.loads(str(raw or ""))
                except (TypeError, ValueError, json.JSONDecodeError):
                    invalid_json += 1
            results["invalid_service_lines_json"] = invalid_json

            cursor.execute(
                """
                SELECT COUNT(*) AS n
                FROM information_schema.tables
                WHERE table_schema = DATABASE()
                  AND table_type = 'BASE TABLE'
                  AND engine <> 'InnoDB'
                """
            )
            non_innodb = _value(cursor.fetchone())
            cursor.execute(
                """
                SELECT COUNT(*) AS n
                FROM information_schema.tables
                WHERE table_schema = DATABASE()
                  AND table_type = 'BASE TABLE'
                  AND table_collation NOT LIKE 'utf8mb4%'
                """
            )
            non_utf8mb4 = _value(cursor.fetchone())
            connection.rollback()
    finally:
        connection.close()

    return {
        "mode": "read-only",
        "integrity_checks": results,
        "schema_hardening": {
            "non_innodb_tables": non_innodb,
            "tables_not_utf8mb4": non_utf8mb4,
        },
        "integrity_issue_count": sum(results.values()),
    }


def main_entry() -> int:
    report = audit()
    print(json.dumps(report, ensure_ascii=False, indent=2))
    print("database-readonly-audit-ok")
    return 0


if __name__ == "__main__":
    raise SystemExit(main_entry())
