"""Payment aggregate adapter; does not expose mutable records to callers."""
from copy import deepcopy


class LegacyPaymentRepository:
    def __init__(self, get_data, save_dataset):
        self.get_data = get_data
        self.save_dataset = save_dataset

    def get(self, number):
        return deepcopy(next((row for row in self.get_data().get("faturacao", [])
                              if str(row.get("numero", "") or "").strip() == number), None))

    def save(self, record, *, expected):
        number = str(record.get("numero", "") or "").strip()
        rows = self.get_data().get("faturacao", [])
        index = next((index for index, row in enumerate(rows)
                      if str(row.get("numero", "") or "").strip() == number), None)
        if index is None or rows[index] != expected:
            raise ValueError("O registo de faturacao foi alterado. Atualize e tente novamente.")
        rows[index] = deepcopy(record)
        try:
            self.save_dataset(force=True, blocking=True)
        except Exception:
            current = self.get_data().setdefault("faturacao", [])
            for index, row in enumerate(current):
                if str(row.get("numero", "") or "").strip() == number:
                    current[index] = deepcopy(expected)
                    break
            else:
                current.append(deepcopy(expected))
            raise
