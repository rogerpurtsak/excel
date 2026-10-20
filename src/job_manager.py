"""SQLite-backed job state persistence for recovery across restarts."""

from __future__ import annotations

import json
import sqlite3
import time
from enum import Enum
from pathlib import Path
from typing import Any

from .config import APP_SUPPORT_DIR

DB_PATH = APP_SUPPORT_DIR / "jobs.db"


class JobState(str, Enum):
    PENDING = "PENDING"
    RUNNING = "RUNNING"
    PAUSED = "PAUSED"
    COMPLETED = "COMPLETED"
    FAILED = "FAILED"
    CANCELLED = "CANCELLED"


_ACTIVE_STATES = (JobState.PENDING.value, JobState.RUNNING.value, JobState.PAUSED.value)


class JobManager:
    def __init__(self, db_path: Path = DB_PATH):
        self.db_path = db_path
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._conn = sqlite3.connect(str(self.db_path), check_same_thread=False)
        self._conn.row_factory = sqlite3.Row
        self._init_schema()

    def _init_schema(self) -> None:
        self._conn.execute(
            """
            CREATE TABLE IF NOT EXISTS jobs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                workbook_path TEXT NOT NULL,
                mode TEXT NOT NULL,
                state TEXT NOT NULL,
                completed_tasks INTEGER DEFAULT 0,
                total_tasks INTEGER DEFAULT 0,
                state_blob TEXT,
                created_at REAL NOT NULL,
                updated_at REAL NOT NULL
            )
            """
        )
        self._conn.commit()

    # ------------------------------------------------------------------ CRUD
    def create_job(self, workbook_path: str, mode: str) -> int:
        now = time.time()
        cur = self._conn.execute(
            "INSERT INTO jobs (workbook_path, mode, state, created_at, updated_at) "
            "VALUES (?, ?, ?, ?, ?)",
            (workbook_path, mode, JobState.PENDING.value, now, now),
        )
        self._conn.commit()
        return int(cur.lastrowid)

    def get_job(self, job_id: int) -> dict[str, Any] | None:
        row = self._conn.execute("SELECT * FROM jobs WHERE id = ?", (job_id,)).fetchone()
        return dict(row) if row else None

    def get_active_job(self, workbook_path: str) -> dict[str, Any] | None:
        placeholders = ",".join("?" for _ in _ACTIVE_STATES)
        row = self._conn.execute(
            f"SELECT * FROM jobs WHERE workbook_path = ? AND state IN ({placeholders}) "
            "ORDER BY created_at DESC LIMIT 1",
            (workbook_path, *_ACTIVE_STATES),
        ).fetchone()
        return dict(row) if row else None

    def _set_state(self, job_id: int, state: JobState) -> None:
        self._conn.execute(
            "UPDATE jobs SET state = ?, updated_at = ? WHERE id = ?",
            (state.value, time.time(), job_id),
        )
        self._conn.commit()

    def update_job_progress(self, job_id: int, completed_tasks: int, total_tasks: int) -> None:
        self._conn.execute(
            "UPDATE jobs SET completed_tasks = ?, total_tasks = ?, state = ?, updated_at = ? "
            "WHERE id = ?",
            (completed_tasks, total_tasks, JobState.RUNNING.value, time.time(), job_id),
        )
        self._conn.commit()

    def pause_job(self, job_id: int) -> None:
        self._set_state(job_id, JobState.PAUSED)

    def resume_job(self, job_id: int) -> None:
        self._set_state(job_id, JobState.RUNNING)

    def cancel_job(self, job_id: int) -> None:
        self._set_state(job_id, JobState.CANCELLED)

    def complete_job(self, job_id: int) -> None:
        self._set_state(job_id, JobState.COMPLETED)

    def fail_job(self, job_id: int) -> None:
        self._set_state(job_id, JobState.FAILED)

    def save_job_state(self, job_id: int, state_dict: dict[str, Any]) -> None:
        self._conn.execute(
            "UPDATE jobs SET state_blob = ?, updated_at = ? WHERE id = ?",
            (json.dumps(state_dict), time.time(), job_id),
        )
        self._conn.commit()

    def load_job_state(self, job_id: int) -> dict[str, Any] | None:
        row = self._conn.execute(
            "SELECT state_blob FROM jobs WHERE id = ?", (job_id,)
        ).fetchone()
        if row and row["state_blob"]:
            return json.loads(row["state_blob"])
        return None

    def list_jobs(self) -> list[dict[str, Any]]:
        rows = self._conn.execute("SELECT * FROM jobs ORDER BY created_at DESC").fetchall()
        return [dict(r) for r in rows]

    def close(self) -> None:
        self._conn.close()
