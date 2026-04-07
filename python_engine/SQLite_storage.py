import sqlite3
import json
import os
from datetime import datetime

DB_PATH = os.environ.get("GPT_EXCEL_DB", "gpt_excel.db")


def get_connection() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    """Create all tables if they don't exist."""
    conn = get_connection()
    cur = conn.cursor()

    cur.executescript("""
        CREATE TABLE IF NOT EXISTS uploaded_files (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            filename    TEXT NOT NULL,
            file_path   TEXT NOT NULL,
            uploaded_at TEXT NOT NULL,
            row_count   INTEGER,
            col_count   INTEGER,
            columns     TEXT
        );

        CREATE TABLE IF NOT EXISTS analysis_results (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            file_id     INTEGER REFERENCES uploaded_files(id),
            task        TEXT NOT NULL,
            result_json TEXT NOT NULL,
            created_at  TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS generated_files (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            file_id      INTEGER REFERENCES uploaded_files(id),
            output_type  TEXT NOT NULL,
            output_path  TEXT NOT NULL,
            created_at   TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS automation_logs (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            task       TEXT NOT NULL,
            status     TEXT NOT NULL,
            message    TEXT,
            ran_at     TEXT NOT NULL
        );
    """)
    conn.commit()
    conn.close()


# ── Uploaded Files ────────────────────────────────────────────────────────────

def save_file_record(filename: str, file_path: str,
                     row_count: int = 0, col_count: int = 0,
                     columns: list = None) -> int:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO uploaded_files (filename, file_path, uploaded_at, row_count, col_count, columns)
        VALUES (?, ?, ?, ?, ?, ?)
    """, (
        filename, file_path,
        datetime.now().isoformat(),
        row_count, col_count,
        json.dumps(columns or [])
    ))
    file_id = cur.lastrowid
    conn.commit()
    conn.close()
    return file_id


def get_all_files() -> list:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM uploaded_files ORDER BY uploaded_at DESC")
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows


def get_file_by_id(file_id: int) -> dict | None:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SELECT * FROM uploaded_files WHERE id = ?", (file_id,))
    row = cur.fetchone()
    conn.close()
    return dict(row) if row else None


# ── Analysis Results ──────────────────────────────────────────────────────────

def save_analysis(file_id: int, task: str, result: dict) -> int:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO analysis_results (file_id, task, result_json, created_at)
        VALUES (?, ?, ?, ?)
    """, (file_id, task, json.dumps(result), datetime.now().isoformat()))
    row_id = cur.lastrowid
    conn.commit()
    conn.close()
    return row_id


def get_analysis_by_file(file_id: int) -> list:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        "SELECT * FROM analysis_results WHERE file_id = ? ORDER BY created_at DESC",
        (file_id,)
    )
    rows = [dict(r) for r in cur.fetchall()]
    for r in rows:
        try:
            r["result_json"] = json.loads(r["result_json"])
        except Exception:
            pass
    conn.close()
    return rows


# ── Generated Files ───────────────────────────────────────────────────────────

def save_generated_file(file_id: int, output_type: str, output_path: str) -> int:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO generated_files (file_id, output_type, output_path, created_at)
        VALUES (?, ?, ?, ?)
    """, (file_id, output_type, output_path, datetime.now().isoformat()))
    row_id = cur.lastrowid
    conn.commit()
    conn.close()
    return row_id


def get_generated_files(file_id: int) -> list:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        "SELECT * FROM generated_files WHERE file_id = ? ORDER BY created_at DESC",
        (file_id,)
    )
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows


# ── Automation Logs ───────────────────────────────────────────────────────────

def log_automation(task: str, status: str, message: str = "") -> None:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO automation_logs (task, status, message, ran_at)
        VALUES (?, ?, ?, ?)
    """, (task, status, message, datetime.now().isoformat()))
    conn.commit()
    conn.close()


def get_automation_logs(limit: int = 50) -> list:
    conn = get_connection()
    cur = conn.cursor()
    cur.execute(
        "SELECT * FROM automation_logs ORDER BY ran_at DESC LIMIT ?",
        (limit,)
    )
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return rows


# ── Stats ─────────────────────────────────────────────────────────────────────

def get_db_stats() -> dict:
    conn = get_connection()
    cur = conn.cursor()
    stats = {}
    for table in ("uploaded_files", "analysis_results", "generated_files", "automation_logs"):
        cur.execute(f"SELECT COUNT(*) FROM {table}")
        stats[table] = cur.fetchone()[0]
    conn.close()
    return stats


# Auto-init on import
init_db()
