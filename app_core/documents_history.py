# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import sqlite3
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

from .constants import app_dir


def _iso(dt: Optional[datetime]) -> str:
    if not dt:
        return ""
    try:
        return dt.isoformat(timespec="seconds")
    except Exception:
        return str(dt)


def _normalize_level(value: Any) -> str:
    raw = str(value or "").strip().lower()
    if raw in ("info", "informacao", "informação"):
        return "info"
    if raw in ("warn", "warning", "alerta", "atencao", "atenção"):
        return "warn"
    if raw in ("error", "erro", "err", "falha", "failed", "exception"):
        return "error"
    return "info"


@dataclass(frozen=True)
class DocumentRecord:
    boleto_grid: str
    movto_id: str
    customer_id: str
    customer_email: str
    customer_name: str
    documento: str
    generated_at: str
    status: str
    last_attempt_at: str
    sent_at: str
    error: str


@dataclass(frozen=True)
class RunRecord:
    id: int
    started_at: str
    finished_at: str
    status: str
    window_start: str
    window_end: str
    result: Dict[str, Any]


@dataclass(frozen=True)
class EventRecord:
    id: int
    created_at: str
    kind: str
    source: str
    level: str
    title: str
    message: str


class DocumentsHistory:
    def __init__(self, db_path: Optional[Path] = None):
        self.db_path = db_path or (app_dir() / "envio_documentos.sqlite3")
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(str(self.db_path))
        conn.execute("PRAGMA journal_mode=WAL;")
        conn.execute("PRAGMA synchronous=NORMAL;")
        return conn

    def _init_db(self) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS documents (
                    boleto_grid TEXT PRIMARY KEY,
                    movto_id TEXT,
                    customer_id TEXT,
                    customer_email TEXT,
                    customer_name TEXT,
                    documento TEXT,
                    generated_at TEXT,
                    status TEXT,
                    last_attempt_at TEXT,
                    sent_at TEXT,
                    error TEXT
                )
                """
            )
            try:
                conn.execute("ALTER TABLE documents ADD COLUMN customer_name TEXT")
            except Exception:
                pass
            conn.execute("CREATE INDEX IF NOT EXISTS idx_documents_status ON documents(status)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_documents_generated_at ON documents(generated_at)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_documents_movto_id ON documents(movto_id)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_documents_documento ON documents(documento)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_documents_customer_id ON documents(customer_id)")
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS runs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    started_at TEXT,
                    finished_at TEXT,
                    status TEXT,
                    window_start TEXT,
                    window_end TEXT,
                    result_json TEXT
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS events (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    created_at TEXT,
                    kind TEXT,
                    source TEXT,
                    level TEXT,
                    title TEXT,
                    message TEXT
                )
                """
            )
            conn.execute("CREATE INDEX IF NOT EXISTS idx_events_created_at ON events(created_at)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_events_kind ON events(kind)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_events_level ON events(level)")
            conn.commit()

    def start_run(self, window_start: datetime, window_end: datetime) -> int:
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                "INSERT INTO runs(started_at, finished_at, status, window_start, window_end, result_json) VALUES (?,?,?,?,?,?)",
                (_iso(datetime.now()), "", "running", _iso(window_start), _iso(window_end), ""),
            )
            conn.commit()
            return int(cur.lastrowid)

    def finish_run(self, run_id: int, status: str, result: Dict[str, Any]) -> None:
        with self._connect() as conn:
            conn.execute(
                "UPDATE runs SET finished_at=?, status=?, result_json=? WHERE id=?",
                (_iso(datetime.now()), str(status or ""), json.dumps(result or {}, ensure_ascii=False), int(run_id)),
            )
            conn.commit()

    def list_runs(self, limit: int = 200) -> List[RunRecord]:
        limit = max(1, min(5000, int(limit or 200)))
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT id, started_at, finished_at, status, window_start, window_end, result_json
                  FROM runs
                 ORDER BY COALESCE(NULLIF(finished_at,''), started_at) DESC, id DESC
                 LIMIT ?
                """,
                (limit,),
            ).fetchall()
        out: List[RunRecord] = []
        for r in rows:
            try:
                payload = json.loads(str(r[6] or "").strip() or "{}")
            except Exception:
                payload = {}
            out.append(
                RunRecord(
                    id=int(r[0] or 0),
                    started_at=str(r[1] or ""),
                    finished_at=str(r[2] or ""),
                    status=str(r[3] or ""),
                    window_start=str(r[4] or ""),
                    window_end=str(r[5] or ""),
                    result=payload if isinstance(payload, dict) else {},
                )
            )
        return out

    def add_event(
        self,
        *,
        kind: str,
        source: str,
        title: str,
        message: str,
        level: str = "info",
        created_at: Optional[datetime] = None,
    ) -> int:
        kind = str(kind or "").strip()
        source = str(source or "").strip()
        level = _normalize_level(level)
        title = str(title or "").strip()
        message = str(message or "").strip()
        created_at = created_at or datetime.now()
        with self._connect() as conn:
            cur = conn.cursor()
            cur.execute(
                "INSERT INTO events(created_at, kind, source, level, title, message) VALUES (?,?,?,?,?,?)",
                (_iso(created_at), kind, source, level, title, message[:4000]),
            )
            conn.commit()
            return int(cur.lastrowid or 0)

    def list_events(self, limit: int = 500) -> List[EventRecord]:
        limit = max(1, min(5000, int(limit or 500)))
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT id, created_at, kind, source, level, title, message
                  FROM events
                 ORDER BY created_at DESC, id DESC
                 LIMIT ?
                """,
                (limit,),
            ).fetchall()
        out: List[EventRecord] = []
        for r in rows:
            out.append(
                EventRecord(
                    id=int(r[0] or 0),
                    created_at=str(r[1] or ""),
                    kind=str(r[2] or ""),
                    source=str(r[3] or ""),
                    level=str(r[4] or ""),
                    title=str(r[5] or ""),
                    message=str(r[6] or ""),
                )
            )
        return out

    def clear_all(self) -> None:
        with self._connect() as conn:
            conn.execute("DELETE FROM events")
            conn.execute("DELETE FROM runs")
            conn.execute("DELETE FROM documents")
            conn.commit()

    def upsert_generated(self, row: Dict[str, Any], *, allow_duplicate: bool = False) -> str:
        boleto_grid = str(row.get("boleto_grid") or "").strip()
        if not boleto_grid:
            return "ignored"
        movto_id = str(row.get("movto_id") or "").strip()
        customer_id = str(row.get("customer_id") or "").strip()
        customer_email = str(row.get("customer_email") or "").strip()
        customer_name = str(row.get("cliente") or row.get("customer_name") or "").strip()
        documento = str(row.get("documento") or "").strip()
        generated_at = str(row.get("generated_at") or "").strip()
        with self._connect() as conn:
            existing = conn.execute(
                "SELECT status FROM documents WHERE boleto_grid=?",
                (boleto_grid,),
            ).fetchone()
            existing_status = str(existing[0] or "").strip().lower() if existing else ""
            if existing and existing_status == "sent" and not allow_duplicate:
                return "kept_sent"
            if not allow_duplicate and not existing and movto_id:
                already_sent = conn.execute(
                    "SELECT boleto_grid FROM documents WHERE movto_id=? AND status='sent' LIMIT 1",
                    (movto_id,),
                ).fetchone()
                if already_sent:
                    conn.execute(
                        """
                        INSERT INTO documents(
                            boleto_grid, movto_id, customer_id, customer_email, customer_name, documento, generated_at,
                            status, last_attempt_at, sent_at, error
                        ) VALUES (?,?,?,?,?,?,?,?,?,?,?)
                        """,
                        (
                            boleto_grid,
                            movto_id,
                            customer_id,
                            customer_email,
                            customer_name,
                            documento,
                            generated_at,
                            "skipped_duplicate",
                            _iso(datetime.now()),
                            "",
                            f"Já enviado anteriormente para este movto_id (grid={str(already_sent[0] or '').strip()}).",
                        ),
                    )
                    conn.commit()
                    return "skipped_duplicate"
            if existing:
                conn.execute(
                    """
                    UPDATE documents
                       SET movto_id=?, customer_id=?, customer_email=?, customer_name=?, documento=?, generated_at=?, status=CASE WHEN ? THEN 'pending' ELSE status END, error=CASE WHEN ? THEN '' ELSE error END
                     WHERE boleto_grid=?
                    """,
                    (movto_id, customer_id, customer_email, customer_name, documento, generated_at, bool(allow_duplicate), bool(allow_duplicate), boleto_grid),
                )
            else:
                conn.execute(
                    """
                    INSERT INTO documents(
                        boleto_grid, movto_id, customer_id, customer_email, customer_name, documento, generated_at,
                        status, last_attempt_at, sent_at, error
                    ) VALUES (?,?,?,?,?,?,?,?,?,?,?)
                    """,
                    (boleto_grid, movto_id, customer_id, customer_email, customer_name, documento, generated_at, "pending", "", "", ""),
                )
            conn.commit()
        return "pending"

    def list_pending(self, limit: int = 500) -> List[DocumentRecord]:
        limit = max(1, min(5000, int(limit or 500)))
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT boleto_grid, movto_id, customer_id, customer_email, customer_name, documento, generated_at, status, last_attempt_at, sent_at, error
                  FROM documents
                 WHERE status = 'pending'
                 ORDER BY generated_at, boleto_grid
                 LIMIT ?
                """,
                (limit,),
            ).fetchall()
        out: List[DocumentRecord] = []
        for r in rows:
            out.append(
                DocumentRecord(
                    boleto_grid=str(r[0] or ""),
                    movto_id=str(r[1] or ""),
                    customer_id=str(r[2] or ""),
                    customer_email=str(r[3] or ""),
                    customer_name=str(r[4] or ""),
                    documento=str(r[5] or ""),
                    generated_at=str(r[6] or ""),
                    status=str(r[7] or ""),
                    last_attempt_at=str(r[8] or ""),
                    sent_at=str(r[9] or ""),
                    error=str(r[10] or ""),
                )
            )
        return out

    def list_pending_by_grids(self, boleto_grids: Iterable[str]) -> List[DocumentRecord]:
        grids = [str(g or "").strip() for g in (boleto_grids or []) if str(g or "").strip()]
        if not grids:
            return []
        if len(grids) > 900:
            grids = grids[:900]
        placeholders = ",".join(["?"] * len(grids))
        sql = f"""
            SELECT boleto_grid, movto_id, customer_id, customer_email, customer_name, documento, generated_at, status, last_attempt_at, sent_at, error
              FROM documents
             WHERE status = 'pending'
               AND boleto_grid IN ({placeholders})
             ORDER BY generated_at, boleto_grid
        """
        with self._connect() as conn:
            rows = conn.execute(sql, tuple(grids)).fetchall()
        out: List[DocumentRecord] = []
        for r in rows:
            out.append(
                DocumentRecord(
                    boleto_grid=str(r[0] or ""),
                    movto_id=str(r[1] or ""),
                    customer_id=str(r[2] or ""),
                    customer_email=str(r[3] or ""),
                    customer_name=str(r[4] or ""),
                    documento=str(r[5] or ""),
                    generated_at=str(r[6] or ""),
                    status=str(r[7] or ""),
                    last_attempt_at=str(r[8] or ""),
                    sent_at=str(r[9] or ""),
                    error=str(r[10] or ""),
                )
            )
        return out

    def list_retryable(self, *, limit: int = 500, no_email_retry_hours: int = 24) -> List[DocumentRecord]:
        limit = max(1, min(5000, int(limit or 500)))
        no_email_retry_hours = max(1, min(24 * 30, int(no_email_retry_hours or 24)))
        cutoff = _iso(datetime.now() - timedelta(hours=no_email_retry_hours))
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT boleto_grid, movto_id, customer_id, customer_email, customer_name, documento, generated_at, status, last_attempt_at, sent_at, error
                  FROM documents
                 WHERE status IN ('pending','failed')
                    OR (status='no_email' AND (last_attempt_at='' OR last_attempt_at < ?))
                 ORDER BY generated_at, boleto_grid
                 LIMIT ?
                """,
                (cutoff, limit),
            ).fetchall()
        out: List[DocumentRecord] = []
        for r in rows:
            out.append(
                DocumentRecord(
                    boleto_grid=str(r[0] or ""),
                    movto_id=str(r[1] or ""),
                    customer_id=str(r[2] or ""),
                    customer_email=str(r[3] or ""),
                    customer_name=str(r[4] or ""),
                    documento=str(r[5] or ""),
                    generated_at=str(r[6] or ""),
                    status=str(r[7] or ""),
                    last_attempt_at=str(r[8] or ""),
                    sent_at=str(r[9] or ""),
                    error=str(r[10] or ""),
                )
            )
        return out

    def list_problems(self) -> List[DocumentRecord]:
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT boleto_grid, movto_id, customer_id, customer_email, customer_name, documento, generated_at, status, last_attempt_at, sent_at, error
                  FROM documents
                 WHERE status IN ('failed', 'no_email')
                 ORDER BY last_attempt_at DESC
                 LIMIT 1000
                """
            ).fetchall()
        out: List[DocumentRecord] = []
        for r in rows:
            out.append(
                DocumentRecord(
                    boleto_grid=str(r[0] or ""),
                    movto_id=str(r[1] or ""),
                    customer_id=str(r[2] or ""),
                    customer_email=str(r[3] or ""),
                    customer_name=str(r[4] or ""),
                    documento=str(r[5] or ""),
                    generated_at=str(r[6] or ""),
                    status=str(r[7] or ""),
                    last_attempt_at=str(r[8] or ""),
                    sent_at=str(r[9] or ""),
                    error=str(r[10] or ""),
                )
            )
        return out

    def list_sent(self, limit: int = 1000) -> List[DocumentRecord]:
        limit = max(1, min(5000, int(limit or 1000)))
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT boleto_grid, movto_id, customer_id, customer_email, customer_name, documento, generated_at, status, last_attempt_at, sent_at, error
                  FROM documents
                 WHERE status = 'sent'
                 ORDER BY sent_at DESC, boleto_grid
                 LIMIT ?
                """,
                (limit,),
            ).fetchall()
        out: List[DocumentRecord] = []
        for r in rows:
            out.append(
                DocumentRecord(
                    boleto_grid=str(r[0] or ""),
                    movto_id=str(r[1] or ""),
                    customer_id=str(r[2] or ""),
                    customer_email=str(r[3] or ""),
                    customer_name=str(r[4] or ""),
                    documento=str(r[5] or ""),
                    generated_at=str(r[6] or ""),
                    status=str(r[7] or ""),
                    last_attempt_at=str(r[8] or ""),
                    sent_at=str(r[9] or ""),
                    error=str(r[10] or ""),
                )
            )
        return out

    def list_sent_for_email_around(
        self,
        to_email: str,
        center_at: Any,
        *,
        window_minutes: int = 10,
        limit: int = 300,
    ) -> List[DocumentRecord]:
        to_email = str(to_email or "").strip()
        if not to_email:
            return []
        try:
            dt = datetime.fromisoformat(str(center_at or "").strip())
        except Exception:
            return []
        window_minutes = max(1, min(120, int(window_minutes or 10)))
        limit = max(1, min(5000, int(limit or 300)))
        start = _iso(dt - timedelta(minutes=window_minutes))
        end = _iso(dt + timedelta(minutes=window_minutes))
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT boleto_grid, movto_id, customer_id, customer_email, customer_name, documento, generated_at, status, last_attempt_at, sent_at, error
                  FROM documents
                 WHERE status = 'sent'
                   AND lower(customer_email) = lower(?)
                   AND sent_at >= ?
                   AND sent_at <= ?
                 ORDER BY sent_at DESC, boleto_grid
                 LIMIT ?
                """,
                (to_email, start, end, limit),
            ).fetchall()
        out: List[DocumentRecord] = []
        for r in rows:
            out.append(
                DocumentRecord(
                    boleto_grid=str(r[0] or ""),
                    movto_id=str(r[1] or ""),
                    customer_id=str(r[2] or ""),
                    customer_email=str(r[3] or ""),
                    customer_name=str(r[4] or ""),
                    documento=str(r[5] or ""),
                    generated_at=str(r[6] or ""),
                    status=str(r[7] or ""),
                    last_attempt_at=str(r[8] or ""),
                    sent_at=str(r[9] or ""),
                    error=str(r[10] or ""),
                )
            )
        return out

    def list_sent_by_grids(self, boleto_grids: Iterable[str]) -> Dict[str, DocumentRecord]:
        grids = [str(g or "").strip() for g in (boleto_grids or []) if str(g or "").strip()]
        if not grids:
            return {}
        if len(grids) > 900:
            grids = grids[:900]
        placeholders = ",".join(["?"] * len(grids))
        sql = f"""
            SELECT boleto_grid, movto_id, customer_id, customer_email, customer_name, documento, generated_at, status, last_attempt_at, sent_at, error
              FROM documents
             WHERE status = 'sent'
               AND boleto_grid IN ({placeholders})
        """
        with self._connect() as conn:
            rows = conn.execute(sql, tuple(grids)).fetchall()
        out: Dict[str, DocumentRecord] = {}
        for r in rows:
            rec = DocumentRecord(
                boleto_grid=str(r[0] or ""),
                movto_id=str(r[1] or ""),
                customer_id=str(r[2] or ""),
                customer_email=str(r[3] or ""),
                customer_name=str(r[4] or ""),
                documento=str(r[5] or ""),
                generated_at=str(r[6] or ""),
                status=str(r[7] or ""),
                last_attempt_at=str(r[8] or ""),
                sent_at=str(r[9] or ""),
                error=str(r[10] or ""),
            )
            if rec.boleto_grid:
                out[rec.boleto_grid] = rec
        return out

    def list_sent_by_movto_ids(self, movto_ids: Iterable[str]) -> Dict[str, DocumentRecord]:
        ids = [str(g or "").strip() for g in (movto_ids or []) if str(g or "").strip()]
        if not ids:
            return {}
        if len(ids) > 900:
            ids = ids[:900]
        placeholders = ",".join(["?"] * len(ids))
        sql = f"""
            SELECT boleto_grid, movto_id, customer_id, customer_email, customer_name, documento, generated_at, status, last_attempt_at, sent_at, error
              FROM documents
             WHERE status = 'sent'
               AND movto_id IN ({placeholders})
        """
        with self._connect() as conn:
            rows = conn.execute(sql, tuple(ids)).fetchall()
        out: Dict[str, DocumentRecord] = {}
        for r in rows:
            rec = DocumentRecord(
                boleto_grid=str(r[0] or ""),
                movto_id=str(r[1] or ""),
                customer_id=str(r[2] or ""),
                customer_email=str(r[3] or ""),
                customer_name=str(r[4] or ""),
                documento=str(r[5] or ""),
                generated_at=str(r[6] or ""),
                status=str(r[7] or ""),
                last_attempt_at=str(r[8] or ""),
                sent_at=str(r[9] or ""),
                error=str(r[10] or ""),
            )
            if rec.movto_id:
                out[rec.movto_id] = rec
        return out

    def mark_sent(self, boleto_grids: Iterable[str], to_email: str) -> None:
        now = _iso(datetime.now())
        to_email = str(to_email or "").strip()
        values = [(str(g or "").strip(),) for g in (boleto_grids or []) if str(g or "").strip()]
        if not values:
            return
        with self._connect() as conn:
            for (g,) in values:
                conn.execute(
                    """
                    UPDATE documents
                       SET status='sent', last_attempt_at=?, sent_at=?, customer_email=COALESCE(NULLIF(customer_email,''), ?), error=''
                     WHERE boleto_grid=?
                    """,
                    (now, now, to_email, g),
                )
            conn.commit()

    def reset_to_pending(self, boleto_grids: Iterable[str]) -> None:
        values = [(str(g or "").strip(),) for g in (boleto_grids or []) if str(g or "").strip()]
        if not values:
            return
        with self._connect() as conn:
            for (g,) in values:
                conn.execute(
                    "UPDATE documents SET status='pending', last_attempt_at='', sent_at='', error='' WHERE boleto_grid=?",
                    (g,),
                )
            conn.commit()

    def mark_no_email(self, boleto_grids: Iterable[str]) -> None:
        now = _iso(datetime.now())
        values = [(str(g or "").strip(),) for g in (boleto_grids or []) if str(g or "").strip()]
        if not values:
            return
        with self._connect() as conn:
            for (g,) in values:
                conn.execute(
                    "UPDATE documents SET status='no_email', last_attempt_at=?, error='' WHERE boleto_grid=?",
                    (now, g),
                )
            conn.commit()

    def mark_failed(self, boleto_grids: Iterable[str], error: str) -> None:
        now = _iso(datetime.now())
        error = str(error or "").strip()
        values = [(str(g or "").strip(),) for g in (boleto_grids or []) if str(g or "").strip()]
        if not values:
            return
        with self._connect() as conn:
            for (g,) in values:
                conn.execute(
                    "UPDATE documents SET status='failed', last_attempt_at=?, error=? WHERE boleto_grid=?",
                    (now, error[:2000], g),
                )
            conn.commit()

    def mark_closed(self, boleto_grids: Iterable[str]) -> None:
        now = _iso(datetime.now())
        values = [(str(g or "").strip(),) for g in (boleto_grids or []) if str(g or "").strip()]
        if not values:
            return
        with self._connect() as conn:
            for (g,) in values:
                conn.execute(
                    "UPDATE documents SET status='closed', last_attempt_at=?, error='' WHERE boleto_grid=?",
                    (now, g),
                )
            conn.commit()

    def vacuum(self) -> None:
        with self._connect() as conn:
            conn.execute("VACUUM")

