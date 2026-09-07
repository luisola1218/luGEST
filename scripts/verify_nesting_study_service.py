"""Nesting validation, snapshot replacement, merge and persistence failures."""
from copy import deepcopy
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def main():
    from lugest_modules.quotes.application.nesting_studies import NestingStudyService
    from lugest_modules.quotes.infrastructure.legacy_nesting_repository import LegacyNestingStudyRepository
    from lugest_modules.quotes.infrastructure.mysql_nesting_store import MysqlNestingStudyStore
    state = [{'orcamentos': [{'numero': 'ORC-1', 'descricao': 'Preservar'}]}]
    stored, remote, writes = {}, {}, []
    failure = [False]

    def save_dataset(**kwargs):
        if failure[0]:
            raise OSError('Database unavailable')
        stored.clear()
        stored.update(deepcopy(state[0]))
        state[0] = deepcopy(state[0])  # The production backend replaces snapshots.
        writes.append(kwargs)

    def save_remote(number, key, label, study):
        remote[key] = deepcopy(study)

    repository = LegacyNestingStudyRepository(get_data=lambda: state[0], save_dataset=save_dataset,
                                             load_remote=lambda number: remote, save_remote=save_remote)
    service = NestingStudyService(repository, now=lambda: '2026-09-07T12:00:00')
    for number, payload in [('missing', {'group_key': 'A'}), ('ORC-1', {})]:
        try:
            service.save(number, payload)
        except ValueError:
            pass
        else:
            raise AssertionError('Invalid study accepted')
    assert not writes
    payload = {'group_key': ' A ', 'group_label': 'Chapa', 'quote_bridge': {'qty': 2}}
    saved = service.save(' ORC-1 ', payload)
    assert saved['quote_number'] == 'ORC-1' and saved['group_key'] == 'A'
    saved['quote_bridge']['qty'] = 999
    assert payload['quote_bridge']['qty'] == 2
    assert stored['orcamentos'][0]['nesting_studies']['A']['quote_bridge']['qty'] == 2
    assert remote['A']['quote_bridge']['qty'] == 2
    before = deepcopy(state[0])
    failure[0] = True
    try:
        service.save('ORC-1', {'group_key': 'B'})
    except OSError:
        pass
    else:
        raise AssertionError('Failed save accepted')
    assert state[0] == before and 'B' not in remote
    failure[0] = False
    remote['A']['updated_at'] = '2026-09-06T12:00:00'
    remote['A']['quote_bridge']['qty'] = 3
    assert service.studies('ORC-1')['A']['quote_bridge']['qty'] == 2
    remote['A']['updated_at'] = '2026-09-08T12:00:00'
    assert service.studies('ORC-1')['A']['quote_bridge']['qty'] == 3
    def failed_mirror(*args):
        raise OSError('Mirror unavailable')
    repository.save_remote = failed_mirror
    assert service.save('ORC-1', {'group_key': 'C'})['group_key'] == 'C'
    assert 'C' in stored['orcamentos'][0]['nesting_studies']
    assert MysqlNestingStudyStore(None).studies('ORC-1') == {}
    # SQL calls parameterize user data, close the connection on both paths.
    class Connection:
        def __init__(self):
            self.closed = False
            self.committed = False
            self.calls = []
        def cursor(self): return self
        def __enter__(self): return self
        def __exit__(self, *args): pass
        def execute(self, sql, args=None): self.calls.append((sql, args))
        def fetchall(self): return [{'group_key': 'A', 'study_json': b'{"qty": 2}'}]
        def commit(self): self.committed = True
        def close(self): self.closed = True
    connection = Connection()
    sql = MysqlNestingStudyStore(lambda: connection)
    assert sql.studies("ORC'1") == {'A': {'qty': 2}}
    assert connection.closed and connection.calls[-1][1] == ("ORC'1",)
    connection = Connection()
    sql.save('ORC-1', 'A', 'Chapa', {'qty': 2})
    assert connection.closed and connection.committed
    assert connection.calls[-1][1][:3] == ('ORC-1', 'A', 'Chapa')
    assert 'main' not in sys.modules
    print('nesting-study-service-ok validation=yes detached=yes failed-save-restored=yes remote-merge=yes sql=yes')


if __name__ == '__main__':
    main()
