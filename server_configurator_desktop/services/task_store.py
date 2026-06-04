from __future__ import annotations

import csv
from pathlib import Path
from typing import Iterable

from ..models import TaskRow


class TaskStore:
    HEADERS = ["IP", "User", "Password"]

    def __init__(self, csv_path: Path):
        self.csv_path = Path(csv_path)

    def load(self) -> list[TaskRow]:
        if not self.csv_path.exists():
            return []
        rows: list[TaskRow] = []
        with self.csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
            reader = csv.DictReader(handle)
            for record in reader:
                ip = (record.get("IP") or "").strip()
                user = (record.get("User") or "").strip()
                password = (record.get("Password") or "").strip()
                if ip:
                    rows.append(TaskRow(ip=ip, user=user, password=password))
        return rows

    def save(self, rows: Iterable[TaskRow]) -> None:
        self.csv_path.parent.mkdir(parents=True, exist_ok=True)
        with self.csv_path.open("w", encoding="utf-8", newline="") as handle:
            writer = csv.DictWriter(handle, fieldnames=self.HEADERS)
            writer.writeheader()
            for row in rows:
                writer.writerow({
                    "IP": row.ip.strip(),
                    "User": row.user.strip(),
                    "Password": row.password,
                })
