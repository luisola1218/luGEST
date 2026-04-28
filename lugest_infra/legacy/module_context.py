def configure_module(module_globals, main_globals):
    module_globals.update(main_globals)
    module_globals["_CONFIGURED"] = True


def ensure_module(module_globals, module_name):
    if module_globals.get("_CONFIGURED"):
        return
    try:
        import main as _main
        module_globals.update(_main.__dict__)
        module_globals["_CONFIGURED"] = True
    except Exception as ex:
        raise RuntimeError(f"{module_name} is not configured: {ex}")
