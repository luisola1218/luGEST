"""Narrow client actions accepted by the presentation layer."""
from dataclasses import dataclass
from typing import Callable


@dataclass(frozen=True)
class ClientActions:
    rows: Callable[[str], list[dict]]
    next_code: Callable[[], str]
    save: Callable[[dict], dict]
    remove: Callable[[str], None]
