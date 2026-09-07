"""Composition only: binds the quote module to the existing runtime."""
from lugest_modules.quotes.application.nesting_studies import NestingStudyService
from lugest_modules.quotes.infrastructure.legacy_nesting_repository import LegacyNestingStudyRepository
from lugest_modules.quotes.infrastructure.mysql_nesting_store import MysqlNestingStudyStore


def nesting_sql_store(backend) -> MysqlNestingStudyStore:
    return MysqlNestingStudyStore(getattr(backend.desktop_main, '_mysql_connect', None))


def nesting_study_service(backend) -> NestingStudyService:
    repository = LegacyNestingStudyRepository(
        get_data=backend.ensure_data,
        save_dataset=backend._save,
        load_remote=backend._mysql_orc_nesting_studies,
        save_remote=backend._mysql_save_orc_nesting_study,
    )
    return NestingStudyService(repository, now=backend.desktop_main.now_iso)
