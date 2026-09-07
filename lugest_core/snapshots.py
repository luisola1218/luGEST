"""Pure snapshot comparison and three-way row merging, without storage or UI."""
from __future__ import annotations

import copy
import json
from typing import Any


class SnapshotMergePolicy:
    def clone_mapping(self, data: dict[str, Any] | None) -> dict[str, Any] | None:
        if not isinstance(data, dict):
            return None
        try:
            return copy.deepcopy(data)
        except Exception:
            try:
                return json.loads(json.dumps(data, ensure_ascii=False, default=str))
            except Exception:
                return dict(data)

    def signature(self, value: Any) -> str:
        try:
            return json.dumps(value, ensure_ascii=False, sort_keys=True, default=str, separators=(",", ":"))
        except Exception:
            return repr(value)

    def changed_buckets(self, current: dict[str, Any], base: dict[str, Any] | None) -> list[str]:
        if not isinstance(current, dict):
            return []
        if not isinstance(base, dict):
            return [key for key in current.keys() if not str(key or "").startswith("__")]
        changed: list[str] = []
        keys = {
            str(key)
            for key in set(list(current.keys()) + list(base.keys()))
            if not str(key or "").startswith("__")
        }
        for key in sorted(keys):
            if self.signature(current.get(key)) != self.signature(base.get(key)):
                changed.append(key)
        return changed

    def identity_field(self, bucket_name: str) -> str:
        return {
            "clientes": "codigo",
            "fornecedores": "id",
            "materiais": "id",
            "produtos": "codigo",
            "produtos_mov": "id",
            "conjuntos": "codigo",
            "conjuntos_modelo": "codigo",
            "orcamentos": "numero",
            "encomendas": "numero",
            "notas_encomenda": "numero",
            "expedicoes": "numero",
            "transportes": "numero",
            "faturacao_registos": "numero",
            "servicos_diretos": "numero",
            "quality_nonconformities": "id",
            "quality_documents": "id",
            "workcenter_catalog": "id",
            "operations_catalog": "name",
            "plano": "id",
            "plano_hist": "id",
        }.get(str(bucket_name or ""), "")

    def merge_list_bucket(self, bucket_name: str, current_value: Any, base_value: Any, latest_value: Any) -> list[Any] | None:
        identity = self.identity_field(bucket_name)
        if not identity or not isinstance(current_value, list) or not isinstance(base_value, list) or not isinstance(latest_value, list):
            return None

        def item_key(item: Any) -> str:
            if not isinstance(item, dict):
                return ""
            return str(item.get(identity, "") or "").strip().casefold()

        current_map: dict[str, Any] = {}
        base_map: dict[str, Any] = {}
        latest_map: dict[str, Any] = {}
        unkeyed_current = []
        unkeyed_base_signatures: set[str] = set()
        unkeyed_latest_signatures: set[str] = set()
        for collection, target, unkeyed in (
            (current_value, current_map, unkeyed_current),
            (base_value, base_map, None),
            (latest_value, latest_map, None),
        ):
            for item in collection:
                key = item_key(item)
                if key:
                    target[key] = item
                    continue
                signature = self.signature(item)
                if unkeyed is not None:
                    unkeyed.append((signature, item))
                elif collection is base_value:
                    unkeyed_base_signatures.add(signature)
                else:
                    unkeyed_latest_signatures.add(signature)

        merged = [copy.deepcopy(item) for item in latest_value if item_key(item)]
        merged_index = {item_key(item): idx for idx, item in enumerate(merged)}

        for key in base_map.keys():
            if key not in current_map and key in merged_index:
                idx = merged_index[key]
                merged[idx] = None

        for key, current_item in current_map.items():
            base_item = base_map.get(key)
            if key not in base_map or self.signature(current_item) != self.signature(base_item):
                cloned = copy.deepcopy(current_item)
                if key in merged_index and merged[merged_index[key]] is not None:
                    merged[merged_index[key]] = cloned
                else:
                    merged.append(cloned)

        keyed = [item for item in merged if item is not None]
        current_unkeyed_signatures = {signature for signature, _item in unkeyed_current}
        for item in latest_value:
            if item_key(item):
                continue
            signature = self.signature(item)
            if signature in unkeyed_base_signatures and signature not in current_unkeyed_signatures:
                continue
            keyed.append(copy.deepcopy(item))
        for signature, item in unkeyed_current:
            if signature not in unkeyed_latest_signatures and signature not in unkeyed_base_signatures:
                keyed.append(copy.deepcopy(item))
        return keyed
