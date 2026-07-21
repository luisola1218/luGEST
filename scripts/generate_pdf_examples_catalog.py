from __future__ import annotations

import copy
import os
import shutil
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Callable

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from pypdf import PdfReader

from lugest_qt.services.legacy_backend import LegacyBackend


def _desktop_catalog_dir() -> Path:
    desktop = Path.home() / "Desktop"
    base = desktop / f"PDFs LUGEST PRODUCAO LIGHT - {date.today().isoformat()}"
    if not base.exists():
        return base
    stamp = datetime.now().strftime("%H%M%S")
    return desktop / f"{base.name} - {stamp}"


def _first_piece(backend: LegacyBackend, order: dict) -> dict:
    pieces = list(backend.desktop_main.encomenda_pecas(order) or [])
    return next((row for row in pieces if isinstance(row, dict)), {})


def main() -> int:
    backend = LegacyBackend()
    live_data = backend.ensure_data()
    backend._replace_data_cache(copy.deepcopy(live_data))
    backend._save = lambda *args, **kwargs: None

    target_dir = _desktop_catalog_dir()
    target_dir.mkdir(parents=True, exist_ok=False)
    results: list[tuple[str, str, str]] = []

    def produce(filename: str, render: Callable[[Path], object]) -> None:
        target = target_dir / filename
        try:
            result = render(target)
            result_path = Path(result) if isinstance(result, (str, Path)) else target
            if result_path.resolve() != target.resolve():
                shutil.copy2(result_path, target)
            if not target.exists() or target.stat().st_size <= 0:
                raise RuntimeError("ficheiro vazio ou inexistente")
            pages = len(PdfReader(str(target)).pages)
            results.append((filename, "OK", f"{pages} pag. | {target.stat().st_size} bytes"))
        except Exception as exc:
            target.unlink(missing_ok=True)
            results.append((filename, "INDISPONIVEL", f"{type(exc).__name__}: {exc}"))

    data = backend.ensure_data()
    conjuntos = [row for row in data.get("conjuntos", []) if isinstance(row, dict) and row.get("itens")]
    quotes = [row for row in data.get("orcamentos", []) if isinstance(row, dict) and row.get("linhas")]
    notes = [row for row in data.get("notas_encomenda", []) if isinstance(row, dict)]
    expeditions = [row for row in data.get("expedicoes", []) if isinstance(row, dict)]
    materials = [row for row in data.get("materiais", []) if isinstance(row, dict)]
    products = [row for row in data.get("produtos", []) if isinstance(row, dict)]
    orders = [row for row in data.get("encomendas", []) if isinstance(row, dict)]

    if conjuntos:
        conjunto = max(conjuntos, key=lambda row: len(list(row.get("itens", []) or [])))
        produce("01_Dossier_Tecnico_Conjunto.pdf", lambda path: backend.conjunto_sheet_pdf(conjunto.get("codigo", ""), path))
    if quotes:
        quote = max(quotes, key=lambda row: len(list(row.get("linhas", []) or [])))
        produce("02_Orcamento_Cliente.pdf", lambda path: backend.orc_render_pdf(quote.get("numero", ""), path))
        if not dict(quote.get("nesting_studies", {}) or {}):
            quote["nesting_studies"] = {
                "EXEMPLO_VISUAL": {
                    "group_key": "EXEMPLO_VISUAL",
                    "group_label": "Chapa S235JR 8 mm",
                    "updated_at": backend.desktop_main.now_iso(),
                    "result_data": {
                        "summary": {
                            "part_count_requested": 4,
                            "part_count_placed": 4,
                            "sheet_count": 1,
                            "utilization_net_pct": 72.5,
                            "material_net_cost_eur": 185.0,
                            "material_purchase_requirement_eur": 0,
                            "selected_sheet_profile": {"name": "Stock disponivel"},
                        },
                        "sheets": [
                            {
                                "index": 1,
                                "source_label": "Chapa stock 3000x1500",
                                "source_kind": "stock",
                                "sheet_width_mm": 3000,
                                "sheet_height_mm": 1500,
                                "part_count": 4,
                                "utilization_net_pct": 72.5,
                                "utilization_bbox_pct": 78.0,
                                "placements": [
                                    {"ref_externa": "P-001", "x_mm": 80, "y_mm": 80, "width_mm": 1350, "height_mm": 620},
                                    {"ref_externa": "P-002", "x_mm": 1510, "y_mm": 80, "width_mm": 1350, "height_mm": 620},
                                    {"ref_externa": "P-003", "x_mm": 80, "y_mm": 780, "width_mm": 900, "height_mm": 600},
                                    {"ref_externa": "P-004", "x_mm": 1060, "y_mm": 780, "width_mm": 1800, "height_mm": 600},
                                ],
                            }
                        ],
                        "sheet_candidates": [],
                        "unplaced": [],
                        "warnings": ["Estudo sintetico apenas para avaliacao visual do PDF."],
                    },
                    "quote_bridge": {
                        "part_count_requested": 4,
                        "part_count_placed": 4,
                        "sheet_count": 1,
                        "selected_profile_name": "Stock disponivel",
                    },
                    "cost_report": {
                        "part_rows": [],
                        "decision_lines": ["Exemplo visual sem impacto no orcamento."],
                        "totals": {},
                    },
                    "options": {"allow_purchase": False},
                }
            }
            quote["latest_nesting_group_key"] = "EXEMPLO_VISUAL"
        produce("03_Estudo_Nesting.pdf", lambda path: backend.orc_render_nesting_study_pdf(quote.get("numero", ""), path))
    if materials:
        material = materials[0]
        produce("04_Stock_Materia_Prima.pdf", lambda path: backend.material_render_stock_pdf(path))
        history_rows = backend.material_history_rows(limit=240)
        produce("05_Historico_Materia_Prima.pdf", lambda path: backend.material_render_history_pdf(history_rows, "Historico de materia-prima", path))
        produce("06_Etiqueta_Materia_Prima.pdf", lambda path: backend.material_identification_label_pdf(material.get("id", ""), path))
    if products:
        product = products[0]
        produce("07_Stock_Produtos.pdf", lambda path: backend.product_render_stock_pdf(path))
        produce("08_Ficha_Produto.pdf", lambda path: backend.product_sheet_pdf(product.get("codigo", ""), path))
        produce("09_Etiqueta_Produto.pdf", lambda path: backend.product_label_pdf(product.get("codigo", ""), path))
    if notes:
        note = max(notes, key=lambda row: len(list(row.get("lines", row.get("linhas", [])) or [])))
        produce("10_Nota_Encomenda.pdf", lambda path: backend.ne_render_pdf(note.get("numero", ""), quote=False, output_path=path))
        produce("11_Pedido_Cotacao.pdf", lambda path: backend.ne_render_pdf(note.get("numero", ""), quote=True, output_path=path))
    if expeditions:
        expedition = expeditions[0]
        produce("12_Guia_Expedicao.pdf", lambda path: backend.expedicao_render_pdf(expedition.get("numero", ""), path))

    order = max(orders, key=lambda row: len(list(backend.desktop_main.encomenda_pecas(row) or [])), default={})
    piece = _first_piece(backend, order) if order else {}
    order_number = str(order.get("numero", "") or "")
    piece_id = str(piece.get("id", "") or "")
    if order_number:
        produce("13_Ordem_Fabrico.pdf", lambda path: backend.order_fabrication_pdf(order_number, output_path=path))
        produce("14_Fluxo_Planeamento_Encomenda.pdf", lambda path: backend.planning_render_order_detail_pdf(order_number, output_path=str(path)))
        if piece_id:
            produce("15_Etiqueta_Operador_Unidade.pdf", lambda path: backend.operator_unit_labels_pdf(order_number, [piece_id], output_path=path))
            produce("16_Etiqueta_Operador_Palete.pdf", lambda path: backend.operator_pallet_labels_pdf(order_number, [piece_id], output_path=path))
        opp = str(piece.get("opp", "") or "")
        if opp:
            original_startfile = os.startfile
            try:
                os.startfile = lambda *args, **kwargs: None
                produce("17_Etiqueta_OPP.pdf", lambda path: backend.opp_open_pdf(opp))
            finally:
                os.startfile = original_startfile

    produce("18_Prazos_Corte_Laser.pdf", lambda path: backend.planning_render_laser_deadlines_pdf(str(path)))
    produce("19_Separacao_Material.pdf", lambda path: backend.material_assistant_render_separation_pdf(output_path=path))

    nc_rows = list(backend.quality_nc_rows("", "Todos") or [])
    if nc_rows:
        nc_id = str(nc_rows[0].get("id", "") or "")
        produce("20_Nao_Conformidade.pdf", lambda path: backend.quality_nc_pdf(nc_id))
        produce("21_Etiqueta_Fornecedor_Qualidade.pdf", lambda path: backend.quality_supplier_label_pdf(nc_id, output_path=path))
    produce("22_Dossier_Qualidade.pdf", lambda path: backend.quality_dossier_pdf())

    synthetic_trip = {
        "numero": "TR-EXEMPLO-001",
        "data_planeada": date.today().isoformat(),
        "hora_saida": "08:00",
        "tipo_responsavel": "Frota propria",
        "estado": "Planeada",
        "viatura": "00-AA-00",
        "motorista": "Motorista exemplo",
        "origem": "Instalacoes LUGEST",
        "transportadora_nome": "Transportadora exemplo",
        "referencia_transporte": "REF-EXEMPLO",
        "paragens": [
            {
                "ordem": 1,
                "encomenda_numero": order_number or "LUGEST-EXEMPLO",
                "cliente_nome": "Cliente industrial exemplo",
                "local_descarga": "Instalacoes do cliente",
                "data_planeada": date.today().isoformat(),
                "guia_numero": str((expeditions[0] if expeditions else {}).get("numero", "") or "GT-EXEMPLO"),
                "estado": "Planeada",
                "paletes": 2,
                "peso_bruto_kg": 850,
                "volume_m3": 3.2,
                "check_carga_ok": True,
                "check_docs_ok": True,
                "check_paletes_ok": True,
            }
        ],
    }
    data.setdefault("transportes", []).append(synthetic_trip)
    produce("23_Folha_Rota_Transporte.pdf", lambda path: backend.transport_route_sheet_render(synthetic_trip["numero"], path))

    invoice_document = {
        "titulo": "Fatura",
        "subtitulo": "Exemplo visual sem valor fiscal",
        "numero_fatura": "FT EXEMPLO/001",
        "legal_invoice_no": "FT EXEMPLO/001",
        "serie": "EXEMPLO",
        "data_emissao": date.today().isoformat(),
        "data_vencimento": date.today().isoformat(),
        "moeda": "EUR",
        "atcud": "DOCUMENTO-NAO-FISCAL",
        "qr_payload": "EXEMPLO VISUAL LUGEST",
        "issuer": {"nome": "Empresa exemplo", "nif": "000000000", "morada": "Morada da empresa"},
        "customer": {"nome": "Cliente industrial exemplo", "nif": "999999990", "morada": "Morada do cliente"},
        "references": {"encomenda": order_number or "-", "guia": "GT-EXEMPLO"},
        "lines": [
            {
                "reference": "PRD-EXEMPLO",
                "description": "Equipamento industrial configurado",
                "quantity": 1,
                "unit": "UN",
                "unit_price": 1000,
                "iva_perc": 23,
                "subtotal": 1000,
                "total": 1230,
            }
        ],
        "subtotal": 1000,
        "valor_iva": 230,
        "valor_total": 1230,
        "saldo": 1230,
        "tax_summary": [{"label": "Base normal", "rate_label": "23%", "base": 1000, "tax": 230}],
        "obs": "EXEMPLO VISUAL. Documento sem validade fiscal.",
    }
    produce("24_Fatura_Exemplo_Nao_Fiscal.pdf", lambda path: backend.billing_pdf_actions.render_invoice_pdf(backend, path, invoice_document))

    ok_count = sum(1 for _name, status, _detail in results if status == "OK")
    report_lines = [
        "CATALOGO PDF LUGEST - PRODUCAO LIGHT",
        f"Gerado em: {datetime.now().isoformat(timespec='seconds')}",
        f"Diretorio: {target_dir}",
        f"Sucesso: {ok_count}/{len(results)}",
        "",
        "Os ficheiros de transporte e faturacao identificados como exemplo foram gerados",
        "com dados sinteticos e nao possuem validade operacional ou fiscal.",
        "",
    ]
    report_lines.extend(f"[{status}] {name} - {detail}" for name, status, detail in results)
    (target_dir / "CATALOGO.txt").write_text("\n".join(report_lines) + "\n", encoding="utf-8")
    print(target_dir)
    for line in report_lines[7:]:
        print(line)
    return 0 if ok_count else 1


if __name__ == "__main__":
    raise SystemExit(main())
