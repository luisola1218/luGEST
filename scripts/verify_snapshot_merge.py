"""Regression cases for concurrent edits and unkeyed historical rows."""
from copy import deepcopy
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from lugest_core.snapshots import SnapshotMergePolicy


def main():
    policy = SnapshotMergePolicy()
    base = [{"codigo": "A", "qty": 1}, {"codigo": "B", "qty": 2}]
    current = [{"codigo": "A", "qty": 4}, {"codigo": "C", "qty": 3}]
    latest = deepcopy(base) + [{"codigo": "D", "qty": 8}]
    before = deepcopy((base, current, latest))
    merged = policy.merge_list_bucket("produtos", current, base, latest)
    assert {row["codigo"]: row["qty"] for row in merged} == {"A": 4, "C": 3, "D": 8}
    assert (base, current, latest) == before
    merged[0]["qty"] = 999
    assert (base, current, latest) == before, "Merge results must not alias input snapshots"
    # Older rows can have no ID; retain once, delete locally, preserve remote adds.
    unkeyed = {"description": "historical"}
    remote = {"description": "remote addition"}
    assert policy.merge_list_bucket("produtos", [unkeyed], [unkeyed], [unkeyed]) == [unkeyed]
    assert policy.merge_list_bucket("produtos", [], [unkeyed], [unkeyed]) == []
    assert policy.merge_list_bucket("produtos", [unkeyed], [unkeyed], []) == [], "Unchanged local rows must not resurrect remote deletions"
    assert policy.merge_list_bucket("produtos", [unkeyed], [unkeyed], [unkeyed, remote]) == [unkeyed, remote]
    assert policy.merge_list_bucket("produtos", [remote], [], [remote]) == [remote]
    assert policy.merge_list_bucket("produtos", [unkeyed], [], []) == [unkeyed]
    assert policy.merge_list_bucket("unknown", current, base, latest) is None
    assert policy.changed_buckets({"produtos": current, "__cache": 2}, {"produtos": base, "__cache": 1}) == ["produtos"]
    # Exercise the existing public backend entry path as well as the pure policy.
    from lugest_qt.services.main_bridge import LegacyBackend
    backend = LegacyBackend.__new__(LegacyBackend)
    assert backend._merge_list_bucket_by_identity("produtos", [], [unkeyed], [unkeyed]) == []
    print("snapshot-merge-ok concurrent-edits=yes no-aliasing=yes unkeyed-deletion=yes no-duplicates=yes")


if __name__ == "__main__":
    main()
