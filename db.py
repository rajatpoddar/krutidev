# coding=utf-8
"""
SQLite persistence layer for Krutidev Editor documents.
No ORM — plain sqlite3 for zero extra dependencies.
"""
import sqlite3
import json
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional, List

DB_PATH = Path(__file__).parent / 'data' / 'sheets.db'


def _conn():
    DB_PATH.parent.mkdir(exist_ok=True)
    con = sqlite3.connect(str(DB_PATH), check_same_thread=False)
    con.row_factory = sqlite3.Row
    con.execute('PRAGMA journal_mode=WAL')   # better concurrent read performance
    con.execute('PRAGMA synchronous=NORMAL') # safe + faster than FULL
    return con


def init_db():
    with _conn() as con:
        con.executescript("""
        CREATE TABLE IF NOT EXISTS documents (
            id          TEXT PRIMARY KEY,
            title       TEXT NOT NULL DEFAULT 'Untitled',
            data        TEXT NOT NULL DEFAULT '{}',
            sheet_names TEXT NOT NULL DEFAULT '[]',
            access      TEXT NOT NULL DEFAULT 'edit',
            password    TEXT,
            source_path TEXT,
            created_at  TEXT NOT NULL,
            updated_at  TEXT NOT NULL
        );
        """)
        # Migrate existing tables that lack source_path column
        try:
            con.execute("ALTER TABLE documents ADD COLUMN source_path TEXT")
        except Exception:
            pass  # Column already exists


# ── CRUD ──────────────────────────────────────────────────────────────────────

def create_doc(title: str, sheets_data: dict, sheet_names: list,
               access: str = 'edit', password: str = None,
               source_path: str = None) -> str:
    doc_id = str(uuid.uuid4())
    now = _now()
    with _conn() as con:
        con.execute(
            "INSERT INTO documents VALUES (?,?,?,?,?,?,?,?,?)",
            (doc_id, title, json.dumps(sheets_data, ensure_ascii=False),
             json.dumps(sheet_names), access, password, source_path, now, now)
        )
    return doc_id


def get_doc(doc_id: str) -> Optional[dict]:
    with _conn() as con:
        row = con.execute("SELECT * FROM documents WHERE id=?", (doc_id,)).fetchone()
    if not row:
        return None
    return _row_to_dict(row)


def update_doc(doc_id: str, sheets_data: dict, sheet_names: list,
               title: str = None, source_path: str = None):
    now = _now()
    with _conn() as con:
        if title and source_path:
            con.execute(
                "UPDATE documents SET data=?, sheet_names=?, title=?, source_path=?, updated_at=? WHERE id=?",
                (json.dumps(sheets_data, ensure_ascii=False),
                 json.dumps(sheet_names), title, source_path, now, doc_id)
            )
        elif title:
            con.execute(
                "UPDATE documents SET data=?, sheet_names=?, title=?, updated_at=? WHERE id=?",
                (json.dumps(sheets_data, ensure_ascii=False),
                 json.dumps(sheet_names), title, now, doc_id)
            )
        elif source_path:
            con.execute(
                "UPDATE documents SET data=?, sheet_names=?, source_path=?, updated_at=? WHERE id=?",
                (json.dumps(sheets_data, ensure_ascii=False),
                 json.dumps(sheet_names), source_path, now, doc_id)
            )
        else:
            con.execute(
                "UPDATE documents SET data=?, sheet_names=?, updated_at=? WHERE id=?",
                (json.dumps(sheets_data, ensure_ascii=False),
                 json.dumps(sheet_names), now, doc_id)
            )


def list_docs(limit: int = 50) -> List[dict]:
    with _conn() as con:
        rows = con.execute(
            "SELECT id, title, access, created_at, updated_at FROM documents "
            "ORDER BY updated_at DESC LIMIT ?", (limit,)
        ).fetchall()
    return [dict(r) for r in rows]


def delete_doc(doc_id: str):
    with _conn() as con:
        con.execute("DELETE FROM documents WHERE id=?", (doc_id,))


# ── Helpers ───────────────────────────────────────────────────────────────────

def _now() -> str:
    return datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S')


def _row_to_dict(row) -> dict:
    d = dict(row)
    data = json.loads(d['data'])
    # Normalise: if data was accidentally stored as a list, reset to empty dict
    d['data']        = data if isinstance(data, dict) else {}
    d['sheet_names'] = json.loads(d['sheet_names'])
    return d