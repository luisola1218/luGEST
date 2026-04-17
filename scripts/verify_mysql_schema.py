from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import main


REQUIRED = {
    "app_config": {"ckey", "cvalue"},
    "clientes": {"codigo", "nome"},
    "conjuntos_modelo": {"codigo", "descricao", "ativo", "updated_at"},
    "conjuntos_modelo_itens": {"conjunto_codigo", "linha_ordem", "tipo_item", "qtd", "preco_unit"},
    "fornecedores": {"id", "nome"},
    "materiais": {"id", "material", "espessura", "quantidade", "reservado", "lote_fornecedor"},
    "encomendas": {"numero", "cliente_codigo", "estado", "data_entrega"},
    "encomenda_espessuras": {"encomenda_numero", "material", "espessura", "tempo_min", "estado"},
    "encomenda_montagem_itens": {"encomenda_numero", "linha_ordem", "tipo_item", "qtd_planeada", "qtd_consumida", "estado"},
    "pecas": {"id", "encomenda_numero", "ref_interna", "ref_externa", "of_codigo", "opp_codigo", "estado"},
    "orcamentos": {"numero", "cliente_codigo", "estado"},
    "orcamento_linhas": {
        "orcamento_numero",
        "ref_interna",
        "ref_externa",
        "material",
        "espessura",
        "tempo_peca_min",
        "tipo_item",
        "produto_codigo",
        "produto_unid",
        "conjunto_codigo",
        "conjunto_nome",
        "grupo_uuid",
        "qtd_base",
    },
    "orc_referencias_historico": {"ref_externa", "ref_interna", "material", "espessura"},
    "notas_encomenda": {"numero", "fornecedor_id", "estado", "total"},
    "notas_encomenda_linhas": {"ne_numero", "ref_material", "origem", "qtd", "preco", "total"},
    "notas_encomenda_entregas": {"ne_numero", "data_entrega", "guia", "fatura"},
    "notas_encomenda_linha_entregas": {"ne_numero", "linha_ordem", "qtd"},
    "notas_encomenda_documentos": {"ne_numero", "tipo", "titulo", "caminho", "guia", "fatura", "data_entrega", "data_documento"},
    "expedicoes": {"numero", "encomenda_numero", "estado"},
    "expedicao_linhas": {"expedicao_numero", "ref_interna", "qtd"},
    "faturacao_registos": {"numero", "orcamento_numero", "encomenda_numero", "cliente_codigo", "data_venda", "data_vencimento"},
    "faturacao_faturas": {
        "registo_numero",
        "documento_id",
        "numero_fatura",
        "data_emissao",
        "valor_total",
        "legal_invoice_no",
        "system_entry_date",
        "source_id",
        "source_billing",
        "hash",
        "hash_control",
        "communication_status",
    },
    "faturacao_pagamentos": {"registo_numero", "pagamento_id", "fatura_documento_id", "data_pagamento", "valor"},
    "plano": {"bloco_id", "encomenda_numero", "data_planeada", "inicio", "duracao_min"},
    "produtos": {"codigo", "descricao", "qty"},
    "produtos_mov": {"codigo", "tipo", "qtd", "data"},
    "stock_log": {"acao", "data"},
    "users": {"username", "password", "role"},
}


def _row_lookup(row, *keys):
    if isinstance(row, dict):
        lowered = {str(key or "").strip().lower(): value for key, value in row.items()}
        for key in keys:
            key_txt = str(key or "").strip()
            if not key_txt:
                continue
            if key_txt in row:
                return row[key_txt]
            found = lowered.get(key_txt.lower())
            if found is not None:
                return found
        return None
    if row:
        try:
            return row[0]
        except Exception:
            return None
    return None


def main_entry() -> int:
    conn = main._mysql_connect()
    try:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT table_name FROM information_schema.tables WHERE table_schema = DATABASE()"
            )
            existing_tables = {
                str(_row_lookup(row, "table_name", "TABLE_NAME") or "").strip()
                for row in (cur.fetchall() or [])
                if str(_row_lookup(row, "table_name", "TABLE_NAME") or "").strip()
            }
            missing_tables = sorted(table for table in REQUIRED if table not in existing_tables)
            if missing_tables:
                raise RuntimeError(f"Tabelas em falta: {', '.join(missing_tables)}")

            missing_columns: list[str] = []
            for table, expected_columns in REQUIRED.items():
                cur.execute(
                    """
                    SELECT column_name
                    FROM information_schema.columns
                    WHERE table_schema = DATABASE() AND table_name = %s
                    """,
                    (table,),
                )
                existing_columns = {
                    str(_row_lookup(row, "column_name", "COLUMN_NAME") or "").strip()
                    for row in (cur.fetchall() or [])
                    if str(_row_lookup(row, "column_name", "COLUMN_NAME") or "").strip()
                }
                missing = sorted(col for col in expected_columns if col not in existing_columns)
                if missing:
                    missing_columns.append(f"{table}: {', '.join(missing)}")

            if missing_columns:
                raise RuntimeError("Colunas em falta:\n" + "\n".join(missing_columns))

    finally:
        conn.close()

    print("mysql-schema-ok")
    return 0


if __name__ == "__main__":
    raise SystemExit(main_entry())
