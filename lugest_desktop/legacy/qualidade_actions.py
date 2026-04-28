from module_context import configure_module, ensure_module

_CONFIGURED = False


def configure(main_globals):
    configure_module(globals(), main_globals)


def _ensure_configured():
    ensure_module(globals(), "qualidade_actions")

def refresh_qualidade(self):
    _ensure_configured()
    for i in self.tbl_qualidade.get_children():
        self.tbl_qualidade.delete(i)
    query = (self.q_filter.get().strip().lower() if hasattr(self, "q_filter") else "")
    for idx, q in enumerate(self.data["qualidade"]):
        values = (q["encomenda"], q["peca"], q["ok"], q["nok"], q["motivo"], q["data"])
        if query and not any(query in str(v).lower() for v in values):
            continue
        tag = "odd" if idx % 2 else "even"
        self.tbl_qualidade.insert("", END, values=values, tags=(tag,))
