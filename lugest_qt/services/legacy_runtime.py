from __future__ import annotations

from types import SimpleNamespace
from typing import Any


LEGACY_CONFIGURED_MODULES = (
    "app_misc_actions",
    "encomendas_actions",
    "materia_actions",
    "ne_expedicao_actions",
    "orc_actions",
    "operador_ordens_actions",
    "plan_actions",
    "produtos_actions",
)


def load_legacy_runtime() -> SimpleNamespace:
    import main as desktop_main
    from lugest_desktop.legacy import app_misc_actions
    from lugest_desktop.legacy import encomendas_actions
    from lugest_desktop.legacy import materia_actions
    from lugest_desktop.legacy import ne_expedicao_actions
    from lugest_desktop.legacy import operador_ordens_actions
    from lugest_desktop.legacy import orc_actions
    from lugest_desktop.legacy import plan_actions
    from lugest_desktop.legacy import produtos_actions
    from lugest_infra.pdf import billing_invoice as billing_pdf_actions

    try:
        from lugest_core.compliance import tax as tax_compliance
    except Exception:
        tax_compliance = SimpleNamespace()

    modules: dict[str, Any] = {
        "app_misc_actions": app_misc_actions,
        "encomendas_actions": encomendas_actions,
        "materia_actions": materia_actions,
        "ne_expedicao_actions": ne_expedicao_actions,
        "orc_actions": orc_actions,
        "operador_ordens_actions": operador_ordens_actions,
        "plan_actions": plan_actions,
        "produtos_actions": produtos_actions,
    }
    for module in modules.values():
        configure = getattr(module, "configure", None)
        if callable(configure):
            configure(desktop_main.__dict__)

    return SimpleNamespace(
        desktop_main=desktop_main,
        billing_pdf_actions=billing_pdf_actions,
        tax_compliance=tax_compliance,
        **modules,
    )


__all__ = ["LEGACY_CONFIGURED_MODULES", "load_legacy_runtime"]
