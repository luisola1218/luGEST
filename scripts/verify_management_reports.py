from __future__ import annotations

import sys
import tempfile
from pathlib import Path


def main() -> int:
    repo_root = Path(__file__).resolve().parents[1]
    if str(repo_root) not in sys.path:
        sys.path.insert(0, str(repo_root))

    from lugest_qt.services.legacy_backend import LegacyBackend
    from lugest_qt.services.runtime_service import RuntimeService

    backend = LegacyBackend()
    backend.ensure_data()
    payload = dict(backend.finance_dashboard("Todos") or {})
    summary = dict(payload.get("executive_summary", {}) or {})

    stock_parts = float(summary.get("stock_materias", 0) or 0) + float(summary.get("stock_produtos", 0) or 0)
    if abs(stock_parts - float(summary.get("stock_total", 0) or 0)) > 0.01:
        raise RuntimeError("O total de stock não coincide com matéria-prima + produto acabado.")
    purchase_parts = float(summary.get("compras_materias", 0) or 0) + float(summary.get("compras_produtos", 0) or 0)
    if abs(purchase_parts - float(summary.get("compras_total", 0) or 0)) > 0.01:
        raise RuntimeError("O total de compras não coincide com as duas origens.")

    target_dir = Path(tempfile.gettempdir()) / "lugest_management_report_verify"
    target_dir.mkdir(parents=True, exist_ok=True)
    company_pdf = backend.dashboard_render_company_report_pdf("Todos", str(target_dir / "company.pdf"))
    pulse_payload = RuntimeService().dashboard(period="7 dias")
    pulse_pdf = backend.dashboard_render_pulse_report_pdf(
        pulse_payload,
        {"period": "7 dias", "year": "Todos", "origin": "Ambos", "view": "Todas"},
        str(target_dir / "pulse.pdf"),
    )
    for pdf_path in (Path(company_pdf), Path(pulse_pdf)):
        if not pdf_path.exists() or pdf_path.stat().st_size < 5_000:
            raise RuntimeError(f"PDF de gestão inválido: {pdf_path}")
        if pdf_path.read_bytes()[:4] != b"%PDF":
            raise RuntimeError(f"Assinatura PDF inválida: {pdf_path}")

    print(
        "management-reports-ok "
        f"stock={float(summary.get('stock_total', 0) or 0):.2f} "
        f"purchases={float(summary.get('compras_total', 0) or 0):.2f} "
        f"company_pdf={Path(company_pdf).stat().st_size} pulse_pdf={Path(pulse_pdf).stat().st_size}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
