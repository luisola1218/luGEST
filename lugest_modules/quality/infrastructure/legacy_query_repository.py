"""Read-only detached quality catalog projection from the runtime."""
from copy import deepcopy
from lugest_modules.quality.application.queries import QualityCatalogs

class LegacyQualityQueryRepository:
    fields = {'quality_nonconformities': 'nonconformities', 'quality_documents': 'documents', 'audit_log': 'audit', 'encomendas': 'orders', 'materiais': 'materials', 'produtos': 'products', 'fornecedores': 'suppliers', 'clientes': 'clients', 'plano': 'plan'}

    def __init__(self, get_data):
        self.get_data = get_data

    def load(self):
        data = self.get_data()
        return QualityCatalogs(**{name: deepcopy(data.get(key, []) or []) for key, name in self.fields.items()})
