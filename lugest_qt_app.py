import sys


def _print_qt_setup_hint() -> None:
    try:
        sys.stderr.write(
            "Erro de arranque Qt: PySide6 nao esta instalado neste ambiente.\n"
            "Crie/atualize a .venv local com:\n"
            "  py -m venv .venv\n"
            "  .venv\\Scripts\\python.exe -m pip install --upgrade pip\n"
            "  .venv\\Scripts\\python.exe -m pip install -r requirements-qt.txt\n"
        )
    except Exception:
        pass


try:
    from lugest_qt.app import main
except ModuleNotFoundError as exc:
    missing_name = str(getattr(exc, "name", "") or "").strip()
    if missing_name == "PySide6" or missing_name.startswith("PySide6."):
        _print_qt_setup_hint()
        raise SystemExit(1)
    raise


if __name__ == "__main__":
    raise SystemExit(main())
