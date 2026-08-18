from __future__ import annotations

import argparse
import concurrent.futures
import importlib.util
import json
import math
import re
import statistics
import sys
import time
import traceback
from datetime import datetime, timedelta
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import main


SAFE_DB_RE = re.compile(r"^lugest_stress_[a-z0-9_]{8,80}$")
REPORT_DIR = ROOT / "reports" / "stress"
SCHEMA_PATH = ROOT / "mysql" / "lugest.sql"
SEED_PATH = ROOT / "scripts" / "stress_seed_company.py"


def _load_module(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Nao foi possivel carregar {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


installer = _load_module(ROOT / "mysql" / "install_lugest_mysql.py", "lugest_mysql_installer")
seed = _load_module(SEED_PATH, "lugest_stress_seed")


def _safe_database_name(value: str, production_name: str) -> str:
    name = str(value or "").strip().lower()
    if not SAFE_DB_RE.fullmatch(name):
        raise ValueError(
            "A base descartavel tem de seguir o formato lugest_stress_<identificador>."
        )
    if "stress" not in name or name == str(production_name or "").strip().lower():
        raise ValueError("Protecao ativa: a base de teste nao pode coincidir com a base configurada.")
    return name


def _server_connection(database: str | None = None, *, dict_rows: bool = False):
    kwargs = {
        "host": main.MYSQL_HOST,
        "port": main.MYSQL_PORT,
        "user": main.MYSQL_USER,
        "password": main.MYSQL_PASSWORD,
        "charset": "utf8mb4",
        "autocommit": False,
        "connect_timeout": 10,
        "read_timeout": 60,
        "write_timeout": 60,
    }
    if database:
        kwargs["database"] = database
    if dict_rows:
        kwargs["cursorclass"] = main.DictCursor
    return main.pymysql.connect(**kwargs)


def _schema_snapshot(database: str) -> dict:
    result: dict[str, int] = {}
    conn = _server_connection(database, dict_rows=True)
    try:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT table_name FROM information_schema.tables "
                "WHERE table_schema=%s AND table_type='BASE TABLE' ORDER BY table_name",
                (database,),
            )
            tables = [str(row.get("table_name") or row.get("TABLE_NAME") or "") for row in cur.fetchall()]
            for table in tables:
                if not re.fullmatch(r"[A-Za-z0-9_]+", table):
                    continue
                cur.execute(f"SELECT COUNT(*) AS n FROM `{table}`")
                result[table] = int((cur.fetchone() or {}).get("n") or 0)
    finally:
        conn.close()
    return result


def _create_disposable_database(database: str) -> dict:
    conn = _server_connection(dict_rows=True)
    try:
        with conn.cursor() as cur:
            cur.execute(
                "SELECT COUNT(*) AS n FROM information_schema.schemata WHERE schema_name=%s",
                (database,),
            )
            if int((cur.fetchone() or {}).get("n") or 0):
                raise RuntimeError(
                    f"A base {database} ja existe. O ensaio recusa reutilizar ou apagar bases preexistentes."
                )
            cur.execute(f"CREATE DATABASE `{database}` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci")
        conn.commit()
    finally:
        conn.close()

    sql_text = SCHEMA_PATH.read_text(encoding="utf-8-sig")
    statements = installer.sanitize_base_schema(sql_text)
    conn = _server_connection(database, dict_rows=True)
    executed = 0
    try:
        with conn.cursor() as cur:
            for statement in statements:
                cur.execute(statement)
                executed += 1
            issues = installer.validate_schema(cur)
            if issues:
                raise RuntimeError("Esquema descartavel invalido: " + " | ".join(issues))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    return {"schema_statements": executed, "schema_validation": "ok"}


def _drop_disposable_database(database: str, production_name: str) -> bool:
    database = _safe_database_name(database, production_name)
    conn = _server_connection(dict_rows=True)
    try:
        with conn.cursor() as cur:
            cur.execute(f"DROP DATABASE IF EXISTS `{database}`")
        conn.commit()
        with conn.cursor() as cur:
            cur.execute(
                "SELECT COUNT(*) AS n FROM information_schema.schemata WHERE schema_name=%s",
                (database,),
            )
            return int((cur.fetchone() or {}).get("n") or 0) == 0
    finally:
        conn.close()


def _switch_main_database(database: str) -> None:
    main.MYSQL_DB_NAME = database
    main._MYSQL_SCHEMA_SYNCED = False
    main._LAST_SAVE_FINGERPRINT = None
    if hasattr(main, "_LAST_SAVE_TS"):
        main._LAST_SAVE_TS = 0.0


def _configure_seed(args, tag: str) -> None:
    seed.TAG = tag
    seed.TARGET_CLIENTES = args.clients
    seed.TARGET_ORCAMENTOS = args.orders
    seed.TARGET_ENCOMENDAS = args.orders
    seed.TARGET_ENCOMENDAS_FINALIZADAS = min(args.completed, args.orders)
    seed.TARGET_NE = args.purchase_notes
    seed.TARGET_PRODUTOS_NE = args.products
    seed.TARGET_MATERIAIS = args.materials


def _add_full_flow(data: dict, orders: list[dict], clients: list[dict], args, tag: str) -> None:
    now = datetime.now()
    expedition_count = min(args.expeditions, len(orders))
    transport_count = min(args.transports, expedition_count)
    invoice_count = min(args.invoices, len(orders))

    for index, order in enumerate(orders[:expedition_count], start=1):
        client = clients[(index - 1) % len(clients)]
        pieces = main.encomenda_pecas(order)
        piece = pieces[0] if pieces else {}
        qty = main.parse_float(piece.get("quantidade_pedida", 1), 1)
        number = f"GT-STRESS-{index:06d}"
        data.setdefault("expedicoes", []).append(
            {
                "numero": number,
                "tipo": "Guia de transporte",
                "encomenda": order.get("numero", ""),
                "cliente": client.get("codigo", ""),
                "cliente_nome": client.get("nome", ""),
                "destinatario": client.get("nome", ""),
                "dest_nif": client.get("nif", ""),
                "dest_morada": client.get("morada", ""),
                "local_carga": "Armazem central de teste",
                "local_descarga": client.get("morada", ""),
                "data_emissao": main.now_iso(),
                "data_transporte": (now + timedelta(days=index % 14)).isoformat(timespec="seconds"),
                "matricula": f"ST-{index % 100:02d}-TS",
                "transportador": "Frota de teste",
                "estado": "Emitida" if index % 4 else "Entregue",
                "observacoes": tag,
                "created_by": "stress-bot",
                "anulada": False,
                "linhas": [
                    {
                        "encomenda": order.get("numero", ""),
                        "peca_id": piece.get("id", ""),
                        "ref_interna": piece.get("ref_interna", ""),
                        "ref_externa": piece.get("ref_externa", ""),
                        "descricao": f"Expedicao sintetica {index}",
                        "qtd": qty,
                        "unid": "UN",
                        "peso": round(qty * 2.75, 3),
                        "manual": False,
                    }
                ],
            }
        )

    for index, order in enumerate(orders[:transport_count], start=1):
        client = clients[(index - 1) % len(clients)]
        expedition = data["expedicoes"][index - 1]
        number = f"TR-STRESS-{index:06d}"
        data.setdefault("transportes", []).append(
            {
                "numero": number,
                "tipo_responsavel": "Proprio" if index % 3 else "Transportadora",
                "estado": "Concluida" if index % 5 == 0 else "Planeada",
                "data_planeada": (now + timedelta(days=index % 21)).strftime("%Y-%m-%d"),
                "hora_saida": f"{7 + index % 9:02d}:00",
                "viatura": "Viatura de ensaio",
                "matricula": f"ST-{index % 100:02d}-TS",
                "motorista": f"Motorista {index % 30 + 1}",
                "telefone_motorista": "910000000",
                "origem": "Armazem central de teste",
                "transportadora_nome": "Transportadora sintetica",
                "referencia_transporte": tag,
                "custo_previsto": round(35 + index % 80, 2),
                "paletes_total_manual": float(1 + index % 4),
                "peso_total_manual_kg": round(80 + index % 700, 3),
                "volume_total_manual_m3": round(0.2 + (index % 15) / 10, 3),
                "observacoes": tag,
                "created_by": "stress-bot",
                "created_at": main.now_iso(),
                "updated_at": main.now_iso(),
                "paragens": [
                    {
                        "ordem": 1,
                        "encomenda_numero": order.get("numero", ""),
                        "expedicao_numero": expedition.get("numero", ""),
                        "cliente_codigo": client.get("codigo", ""),
                        "cliente_nome": client.get("nome", ""),
                        "zona_transporte": f"Zona {index % 12 + 1}",
                        "local_descarga": client.get("morada", ""),
                        "contacto": client.get("contacto", ""),
                        "data_planeada": (now + timedelta(days=index % 21)).isoformat(timespec="seconds"),
                        "paletes": float(1 + index % 4),
                        "peso_bruto_kg": round(80 + index % 700, 3),
                        "volume_m3": round(0.2 + (index % 15) / 10, 3),
                        "custo_transporte": round(35 + index % 80, 2),
                        "estado": "Entregue" if index % 5 == 0 else "Planeada",
                        "check_carga_ok": index % 5 == 0,
                        "check_docs_ok": index % 5 == 0,
                        "check_paletes_ok": index % 5 == 0,
                        "pod_estado": "Confirmado" if index % 5 == 0 else "Pendente",
                        "observacoes": tag,
                    }
                ],
            }
        )
        order["transporte_numero"] = number
        order["estado_transporte"] = "Entregue" if index % 5 == 0 else "Planeado"

    for index, order in enumerate(orders[:invoice_count], start=1):
        client = clients[(index - 1) % len(clients)]
        quote_number = order.get("numero_orcamento", "")
        quote = next((o for o in data.get("orcamentos", []) if o.get("numero") == quote_number), {})
        total = round(main.parse_float(quote.get("total", 0), 0), 2)
        subtotal = round(total / 1.23, 2) if total else 0.0
        invoice_id = f"FT-STRESS-{index:06d}"
        reg_number = f"FAT-STRESS-{index:06d}"
        paid = index % 4 == 0
        record = {
            "numero": reg_number,
            "origem": "Encomenda",
            "orcamento_numero": quote_number,
            "encomenda_numero": order.get("numero", ""),
            "cliente_codigo": client.get("codigo", ""),
            "cliente_nome": client.get("nome", ""),
            "data_venda": now.strftime("%Y-%m-%d"),
            "data_vencimento": (now + timedelta(days=30)).strftime("%Y-%m-%d"),
            "valor_venda_manual": total,
            "estado_pagamento_manual": "Pago" if paid else "Pendente",
            "obs": tag,
            "created_at": main.now_iso(),
            "updated_at": main.now_iso(),
            "faturas": [
                {
                    "id": invoice_id,
                    "doc_type": "FT",
                    "numero_fatura": invoice_id,
                    "serie": "STRESS",
                    "data_emissao": now.strftime("%Y-%m-%d"),
                    "data_vencimento": (now + timedelta(days=30)).strftime("%Y-%m-%d"),
                    "moeda": "EUR",
                    "iva_perc": 23.0,
                    "subtotal": subtotal,
                    "valor_iva": round(total - subtotal, 2),
                    "valor_total": total,
                    "estado": "Emitida",
                    "anulada": False,
                    "source_billing": "P",
                    "created_at": main.now_iso(),
                }
            ],
            "pagamentos": [],
        }
        if paid:
            record["pagamentos"].append(
                {
                    "id": f"PG-STRESS-{index:06d}",
                    "fatura_id": invoice_id,
                    "data_pagamento": now.strftime("%Y-%m-%d"),
                    "valor": total,
                    "metodo": "Transferencia",
                    "referencia": tag,
                    "created_at": main.now_iso(),
                }
            )
        data.setdefault("faturacao", []).append(record)


def _generate_dataset(args, tag: str) -> tuple[dict, dict]:
    _configure_seed(args, tag)
    started = time.perf_counter()
    data = main.load_data()
    clients = seed._ensure_clientes(data)
    products = seed._ensure_produtos(data)
    materials = seed._ensure_materiais(data)
    supplier = seed._ensure_fornecedor_stress(data)
    seed._create_orcamentos_encomendas(data, clients, materials)
    orders = [row for row in data.get("encomendas", []) if seed._enc_is_tagged(row)]
    orders.sort(key=lambda row: str(row.get("numero", "")))
    seed._plan_week(data, orders[: args.orders])
    seed._finalizar_e_baixar(data, orders, materials)
    seed._create_notas_encomenda(data, supplier, products)
    _add_full_flow(data, orders, clients, args, tag)
    counts = {
        "clientes": len(data.get("clientes", [])),
        "produtos": len(data.get("produtos", [])),
        "materiais": len(data.get("materiais", [])),
        "orcamentos": len(data.get("orcamentos", [])),
        "encomendas": len(data.get("encomendas", [])),
        "pecas": sum(len(main.encomenda_pecas(row)) for row in data.get("encomendas", [])),
        "plano": len(data.get("plano", [])),
        "notas_encomenda": len(data.get("notas_encomenda", [])),
        "expedicoes": len(data.get("expedicoes", [])),
        "transportes": len(data.get("transportes", [])),
        "faturacao": len(data.get("faturacao", [])),
    }
    return data, {"generation_sec": round(time.perf_counter() - started, 3), "generated_counts": counts}


def _percentile(values: list[float], percentile: float) -> float:
    if not values:
        return 0.0
    ordered = sorted(values)
    position = (len(ordered) - 1) * percentile
    lower = math.floor(position)
    upper = math.ceil(position)
    if lower == upper:
        return ordered[lower]
    return ordered[lower] + (ordered[upper] - ordered[lower]) * (position - lower)


QUERIES = {
    "orders_active": "SELECT COUNT(*) AS n FROM encomendas WHERE estado IN ('Preparacao','Em curso')",
    "orders_by_client": (
        "SELECT e.numero, c.nome, e.estado, e.data_entrega FROM encomendas e "
        "LEFT JOIN clientes c ON c.codigo=e.cliente_codigo ORDER BY e.data_entrega LIMIT 100"
    ),
    "weekly_plan": (
        "SELECT data_planeada, COUNT(*) AS blocos, SUM(duracao_min) AS minutos FROM plano "
        "GROUP BY data_planeada ORDER BY data_planeada LIMIT 14"
    ),
    "critical_stock": (
        "SELECT id, material, quantidade, reservado FROM materiais "
        "WHERE (quantidade-reservado)<=10 ORDER BY quantidade LIMIT 100"
    ),
    "billing_total": "SELECT COUNT(*) AS docs, COALESCE(SUM(valor_venda_manual),0) AS total FROM faturacao_registos",
    "purchase_pipeline": (
        "SELECT n.estado, COUNT(*) AS docs, COALESCE(SUM(n.total),0) AS total "
        "FROM notas_encomenda n GROUP BY n.estado"
    ),
}


def _one_query(database: str, name: str) -> float:
    conn = _server_connection(database, dict_rows=True)
    try:
        started = time.perf_counter()
        with conn.cursor() as cur:
            cur.execute(QUERIES[name])
            cur.fetchall()
        return (time.perf_counter() - started) * 1000.0
    finally:
        conn.close()


def _benchmark_queries(database: str, repetitions: int, workers: int) -> dict:
    sequential: dict[str, dict] = {}
    for name in QUERIES:
        samples = [_one_query(database, name) for _ in range(repetitions)]
        sequential[name] = {
            "avg_ms": round(statistics.fmean(samples), 3),
            "p95_ms": round(_percentile(samples, 0.95), 3),
            "max_ms": round(max(samples), 3),
        }

    tasks = [name for _ in range(repetitions) for name in QUERIES]
    samples: list[float] = []
    errors: list[str] = []
    started = time.perf_counter()
    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as pool:
        futures = [pool.submit(_one_query, database, name) for name in tasks]
        for future in concurrent.futures.as_completed(futures):
            try:
                samples.append(future.result())
            except Exception as exc:
                errors.append(str(exc))
    return {
        "sequential": sequential,
        "concurrent": {
            "workers": workers,
            "queries": len(tasks),
            "elapsed_sec": round(time.perf_counter() - started, 3),
            "avg_ms": round(statistics.fmean(samples), 3) if samples else None,
            "p95_ms": round(_percentile(samples, 0.95), 3) if samples else None,
            "p99_ms": round(_percentile(samples, 0.99), 3) if samples else None,
            "max_ms": round(max(samples), 3) if samples else None,
            "errors": errors[:10],
            "error_count": len(errors),
        },
    }


def _integrity_checks(database: str, expected: dict) -> dict:
    checks = {
        "orphan_orders_clients": (
            "SELECT COUNT(*) AS n FROM encomendas e LEFT JOIN clientes c ON c.codigo=e.cliente_codigo "
            "WHERE e.cliente_codigo IS NOT NULL AND e.cliente_codigo<>'' AND c.codigo IS NULL"
        ),
        "orphan_pieces_orders": (
            "SELECT COUNT(*) AS n FROM pecas p LEFT JOIN encomendas e ON e.numero=p.encomenda_numero "
            "WHERE e.numero IS NULL"
        ),
        "orphan_quote_lines": (
            "SELECT COUNT(*) AS n FROM orcamento_linhas l LEFT JOIN orcamentos o ON o.numero=l.orcamento_numero "
            "WHERE o.numero IS NULL"
        ),
        "orphan_plan_orders": (
            "SELECT COUNT(*) AS n FROM plano p LEFT JOIN encomendas e ON e.numero=p.encomenda_numero "
            "WHERE e.numero IS NULL"
        ),
        "orphan_expedition_lines": (
            "SELECT COUNT(*) AS n FROM expedicao_linhas l LEFT JOIN expedicoes e ON e.numero=l.expedicao_numero "
            "WHERE e.numero IS NULL"
        ),
        "orphan_transport_stops": (
            "SELECT COUNT(*) AS n FROM transportes_paragens p LEFT JOIN transportes t ON t.numero=p.transporte_numero "
            "WHERE t.numero IS NULL"
        ),
        "orphan_invoices": (
            "SELECT COUNT(*) AS n FROM faturacao_faturas f LEFT JOIN faturacao_registos r ON r.numero=f.registo_numero "
            "WHERE r.numero IS NULL"
        ),
        "negative_material_stock": "SELECT COUNT(*) AS n FROM materiais WHERE quantidade<0 OR reservado<0",
        "negative_product_stock": "SELECT COUNT(*) AS n FROM produtos WHERE qty<0",
    }
    conn = _server_connection(database, dict_rows=True)
    results: dict[str, int] = {}
    counts: dict[str, int] = {}
    try:
        with conn.cursor() as cur:
            for name, sql in checks.items():
                cur.execute(sql)
                results[name] = int((cur.fetchone() or {}).get("n") or 0)
            table_map = {
                "clientes": "clientes",
                "produtos": "produtos",
                "materiais": "materiais",
                "orcamentos": "orcamentos",
                "encomendas": "encomendas",
                "pecas": "pecas",
                "plano": "plano",
                "notas_encomenda": "notas_encomenda",
                "expedicoes": "expedicoes",
                "transportes": "transportes",
                "faturacao": "faturacao_registos",
            }
            for key, table in table_map.items():
                cur.execute(f"SELECT COUNT(*) AS n FROM `{table}`")
                counts[key] = int((cur.fetchone() or {}).get("n") or 0)
    finally:
        conn.close()
    mismatches = {
        key: {"expected": value, "actual": counts.get(key)}
        for key, value in expected.items()
        if key in counts and counts.get(key) != value
    }
    return {
        "sql_checks": results,
        "persisted_counts": counts,
        "count_mismatches": mismatches,
        "ok": all(value == 0 for value in results.values()) and not mismatches,
    }


def _write_report(report: dict, stamp: str) -> tuple[Path, Path]:
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    json_path = REPORT_DIR / f"stress_test_{stamp}.json"
    md_path = REPORT_DIR / f"stress_test_{stamp}.md"
    json_path.write_text(json.dumps(report, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    timings = report.get("timings", {})
    counts = report.get("integrity", {}).get("persisted_counts", {})
    concurrent = report.get("benchmarks", {}).get("concurrent", {})
    lines = [
        "# Ensaio de esforço isolado luGEST",
        "",
        f"- Estado: **{report.get('status', 'desconhecido')}**",
        f"- Executado em: {report.get('finished_at') or report.get('started_at')}",
        f"- Base descartável: `{report.get('database')}`",
        f"- Base operacional permaneceu inalterada: **{'sim' if report.get('production_unchanged') else 'não'}**",
        f"- Base descartável removida: **{'sim' if report.get('database_removed') else 'não'}**",
        "",
        "## Volume persistido",
        "",
    ]
    for key, value in counts.items():
        lines.append(f"- {key}: {value}")
    lines.extend(
        [
            "",
            "## Tempos principais",
            "",
            f"- Geração sintética: {timings.get('generation_sec', '-')} s",
            f"- Gravação relacional: {timings.get('save_sec', '-')} s",
            f"- Leitura integral (primeira): {timings.get('load_first_sec', '-')} s",
            f"- Leitura integral (média): {timings.get('load_avg_sec', '-')} s",
            "",
            "## Concorrência de leitura",
            "",
            f"- Workers: {concurrent.get('workers', '-')}",
            f"- Consultas: {concurrent.get('queries', '-')}",
            f"- p95: {concurrent.get('p95_ms', '-')} ms",
            f"- p99: {concurrent.get('p99_ms', '-')} ms",
            f"- Erros: {concurrent.get('error_count', '-')}",
            "",
            "## Integridade",
            "",
            f"- Resultado: **{'OK' if report.get('integrity', {}).get('ok') else 'REVER'}**",
        ]
    )
    for key, value in report.get("integrity", {}).get("sql_checks", {}).items():
        lines.append(f"- {key}: {value}")
    if report.get("error"):
        lines.extend(["", "## Erro", "", f"```text\n{report['error']}\n```"])
    md_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return json_path, md_path


def run(args) -> int:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    production_name = str(main.MYSQL_DB_NAME or "lugest").strip()
    database = _safe_database_name(args.database or f"lugest_stress_{stamp}", production_name)
    tag = f"STRESS_{stamp}"
    report: dict = {
        "status": "running",
        "started_at": datetime.now().isoformat(timespec="seconds"),
        "database": database,
        "production_database": production_name,
        "parameters": vars(args),
        "database_removed": False,
        "production_unchanged": False,
    }
    created = False
    production_before: dict = {}
    exit_code = 1
    try:
        print(f"[1/7] A fotografar a base operacional {production_name!r} (somente leitura)...", flush=True)
        production_before = _schema_snapshot(production_name)
        report["production_before"] = production_before

        print(f"[2/7] A criar a base descartavel {database!r}...", flush=True)
        report["schema"] = _create_disposable_database(database)
        created = True
        _switch_main_database(database)

        print(
            f"[3/7] A gerar {args.clients} clientes e {args.orders} encomendas com os fluxos associados...",
            flush=True,
        )
        data, generation = _generate_dataset(args, tag)
        report["timings"] = {"generation_sec": generation["generation_sec"]}
        report["generated_counts"] = generation["generated_counts"]

        print("[4/7] A gravar o snapshot relacional completo...", flush=True)
        started = time.perf_counter()
        main.save_data(data, force=True)
        report["timings"]["save_sec"] = round(time.perf_counter() - started, 3)

        print("[5/7] A medir aberturas integrais e consultas concorrentes...", flush=True)
        load_samples: list[float] = []
        for _ in range(args.load_repetitions):
            started = time.perf_counter()
            main.load_data()
            load_samples.append(time.perf_counter() - started)
        report["timings"]["load_first_sec"] = round(load_samples[0], 3)
        report["timings"]["load_avg_sec"] = round(statistics.fmean(load_samples), 3)
        report["timings"]["load_max_sec"] = round(max(load_samples), 3)
        report["benchmarks"] = _benchmark_queries(database, args.query_repetitions, args.workers)

        print("[6/7] A validar contagens, relacoes e stock...", flush=True)
        report["integrity"] = _integrity_checks(database, generation["generated_counts"])
        if not report["integrity"]["ok"]:
            raise RuntimeError("A validacao de integridade encontrou divergencias; consulte o relatorio.")
        report["status"] = "passed"
        exit_code = 0
    except Exception as exc:
        report["status"] = "failed"
        report["error"] = f"{type(exc).__name__}: {exc}"
        report["traceback"] = traceback.format_exc()
        print(report["error"], flush=True)
    finally:
        print("[7/7] A repor o contexto e a remover apenas a base descartavel criada por este ensaio...", flush=True)
        main.MYSQL_DB_NAME = production_name
        try:
            production_after = _schema_snapshot(production_name)
            report["production_after"] = production_after
            report["production_unchanged"] = production_before == production_after
        except Exception as exc:
            report["production_snapshot_error"] = str(exc)
        if created and not args.keep_database:
            try:
                report["database_removed"] = _drop_disposable_database(database, production_name)
            except Exception as exc:
                report["database_drop_error"] = str(exc)
        elif created:
            report["database_removed"] = False
            report["database_kept_by_request"] = True
        report["finished_at"] = datetime.now().isoformat(timespec="seconds")
        paths = _write_report(report, stamp)
        print(json.dumps({
            "status": report["status"],
            "production_unchanged": report["production_unchanged"],
            "database_removed": report["database_removed"],
            "report_json": str(paths[0]),
            "report_markdown": str(paths[1]),
        }, ensure_ascii=False, indent=2), flush=True)
    return exit_code


def parse_args(argv: list[str] | None = None):
    parser = argparse.ArgumentParser(description="Ensaio de esforco seguro numa base MySQL descartavel.")
    parser.add_argument("--clients", type=int, default=1000)
    parser.add_argument("--orders", type=int, default=4000)
    parser.add_argument("--completed", type=int, default=800)
    parser.add_argument("--purchase-notes", type=int, default=400)
    parser.add_argument("--products", type=int, default=500)
    parser.add_argument("--materials", type=int, default=240)
    parser.add_argument("--expeditions", type=int, default=800)
    parser.add_argument("--transports", type=int, default=400)
    parser.add_argument("--invoices", type=int, default=800)
    parser.add_argument("--workers", type=int, default=8)
    parser.add_argument("--query-repetitions", type=int, default=12)
    parser.add_argument("--load-repetitions", type=int, default=3)
    parser.add_argument("--database", default="")
    parser.add_argument("--keep-database", action="store_true")
    args = parser.parse_args(argv)
    positive = (
        "clients", "orders", "completed", "purchase_notes", "products", "materials",
        "expeditions", "transports", "invoices", "workers", "query_repetitions", "load_repetitions",
    )
    for name in positive:
        if getattr(args, name) < 1:
            parser.error(f"--{name.replace('_', '-')} tem de ser maior que zero")
    return args


if __name__ == "__main__":
    raise SystemExit(run(parse_args()))
