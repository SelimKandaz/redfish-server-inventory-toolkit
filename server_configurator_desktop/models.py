from __future__ import annotations

from dataclasses import dataclass


@dataclass
class TaskRow:
    ip: str
    user: str
    password: str
    enabled: bool = True
    status: str = "Pending"
