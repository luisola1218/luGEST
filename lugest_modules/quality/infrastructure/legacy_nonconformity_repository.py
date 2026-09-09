"""Compatibility adapter for the nonconformity catalog."""
from .legacy_catalog_repository import LegacyQualityCatalogRepository


class LegacyNonconformityRepository(LegacyQualityCatalogRepository):
    def __init__(self, get_data, save_dataset, append_audit):
        super().__init__(get_data, save_dataset, append_audit, "quality_nonconformities")
