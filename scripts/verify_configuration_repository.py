"""Persistence failure matrix, without MySQL or workstation configuration."""
from copy import deepcopy
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from lugest_infra.config.repository import ConfigurationRepository, ConfigurationWriteError


class Store:
    def __init__(self, payload=None, fail=False):
        self.payload = payload or {}
        self.fail = fail

    def load(self, default=None):
        if self.fail:
            raise OSError("disk failure")
        return self.payload

    def save(self, payload):
        if self.fail:
            raise OSError("disk failure")
        self.payload = deepcopy(payload)


class Connection:
    def __init__(self, row=None, fail=False):
        self.row, self.fail = row, fail
        self.closed = self.committed = self.rolled_back = False
        self.statements = []

    def cursor(self): return self
    def __enter__(self): return self
    def __exit__(self, *args): return False
    def fetchone(self): return self.row
    def close(self): self.closed = True
    def commit(self): self.committed = True
    def rollback(self): self.rolled_back = True
    def execute(self, sql, params=()):
        self.statements.append((sql, params))
        if self.fail:
            raise OSError("database failure")


def main():
    local = Store({"rows": [{"value": "local"}]})
    db = Connection({"cvalue": b'{"source": "database"}'})
    result = ConfigurationRepository(local, lambda: db).load()
    assert result.payload == {"source": "database"} and db.closed
    assert all(sql.lstrip().startswith("SELECT") for sql, _ in db.statements)
    for row in ((b'{"source": "tuple"}',), {"cvalue": "invalid json"}, None):
        db = Connection(row)
        result = ConfigurationRepository(local, lambda: db).load()
        assert db.closed
        assert result.payload == ({"source": "tuple"} if isinstance(row, tuple) else local.payload)
    db = Connection(fail=True)
    result = ConfigurationRepository(local, lambda: db).load()
    assert "MySQL" in result.error and db.closed and result.payload == local.payload
    result.payload["rows"][0]["value"] = "mutated"
    assert local.payload["rows"][0]["value"] == "local"
    for disk_failure, sql_failure in ((False, False), (True, False), (False, True), (True, True)):
        store = Store(fail=disk_failure)
        db = Connection(fail=sql_failure)
        repository = ConfigurationRepository(store, lambda: db)
        try:
            result = repository.save({"nested": {"value": 3}})
            assert not (disk_failure and sql_failure), "Total failure must raise"
            assert result.payload == {"nested": {"value": 3}}
            assert bool(result.error) == (disk_failure or sql_failure)
        except ConfigurationWriteError:
            assert disk_failure and sql_failure
        assert db.closed
        assert db.committed == (not sql_failure)
        assert db.rolled_back == sql_failure
    # The bridge cache must not leak nested changes between settings screens.
    from lugest_qt.services.main_bridge import LegacyBackend
    backend = LegacyBackend.__new__(LegacyBackend)
    backend._qt_config_cache = {"nested": {"value": 1}}
    backend._load_qt_config()["nested"]["value"] = 2
    assert backend._load_qt_config()["nested"]["value"] == 1
    print("configuration-repository-ok read-only=yes fallback=yes failure-matrix=yes isolation=yes")


if __name__ == "__main__":
    main()
