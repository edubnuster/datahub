import sqlite3
import tempfile
import unittest
from datetime import datetime, timedelta
from pathlib import Path
import shutil
import time

from app_core.documents_history import DocumentsHistory


class DocumentsHistoryTests(unittest.TestCase):
    def _new_history(self):
        tmp_dir = Path(tempfile.mkdtemp(prefix="auto_docs_test_"))

        def _cleanup():
            for _ in range(5):
                try:
                    shutil.rmtree(str(tmp_dir), ignore_errors=False)
                    return
                except PermissionError:
                    time.sleep(0.05)
                except Exception:
                    return

        self.addCleanup(_cleanup)
        db_path = tmp_dir / "envio_documentos.sqlite3"
        return DocumentsHistory(db_path=db_path), db_path

    def test_skip_duplicate_when_movto_already_sent(self):
        h, db_path = self._new_history()

        h.upsert_generated(
            {
                "boleto_grid": "1",
                "movto_id": "10",
                "customer_id": "20",
                "customer_email": "a@a.com",
                "documento": "DOC",
                "generated_at": "2026-01-01 10:00:00",
            }
        )
        h.mark_sent(["1"], to_email="a@a.com")

        h.upsert_generated(
            {
                "boleto_grid": "2",
                "movto_id": "10",
                "customer_id": "20",
                "customer_email": "a@a.com",
                "documento": "DOC",
                "generated_at": "2026-01-01 10:05:00",
            }
        )

        pending = h.list_pending(limit=50)
        self.assertEqual([p.boleto_grid for p in pending], [])

        with sqlite3.connect(str(db_path)) as conn:
            row = conn.execute("SELECT status FROM documents WHERE boleto_grid='2'").fetchone()
        self.assertEqual(str(row[0] or ""), "skipped_duplicate")

    def test_no_email_reappears_after_retry_window(self):
        h, db_path = self._new_history()

        h.upsert_generated(
            {
                "boleto_grid": "1",
                "movto_id": "10",
                "customer_id": "20",
                "customer_email": "",
                "documento": "DOC",
                "generated_at": "2026-01-01 10:00:00",
            }
        )
        h.mark_no_email(["1"])

        now = datetime.now()
        recent = h.list_retryable(limit=50, no_email_retry_hours=24)
        self.assertFalse(any(r.boleto_grid == "1" for r in recent))

        old_attempt = (now - timedelta(hours=48)).isoformat(timespec="seconds")
        with sqlite3.connect(str(db_path)) as conn:
            conn.execute("UPDATE documents SET last_attempt_at=? WHERE boleto_grid='1'", (old_attempt,))
            conn.commit()

        retryable = h.list_retryable(limit=50, no_email_retry_hours=24)
        self.assertTrue(any(r.boleto_grid == "1" for r in retryable))

    def test_clear_all_clears_history(self):
        h, _db_path = self._new_history()
        run_id = h.start_run(window_start=datetime.now() - timedelta(hours=1), window_end=datetime.now())
        h.finish_run(run_id, "ok", {"emails_sent": 1})
        h.add_event(kind="manual_send", source="faturas_receber", level="info", title="Envio", message="Teste")
        h.upsert_generated(
            {
                "boleto_grid": "1",
                "movto_id": "10",
                "customer_id": "20",
                "customer_email": "a@a.com",
                "documento": "DOC",
                "generated_at": "2026-01-01 10:00:00",
            }
        )
        h.mark_failed(["1"], error="Falha")
        self.assertTrue(len(h.list_events(limit=10)) > 0)
        self.assertTrue(len(h.list_runs(limit=10)) > 0)
        self.assertTrue(len(h.list_problems()) > 0)

        h.clear_all()

        self.assertEqual(h.list_events(limit=10), [])
        self.assertEqual(h.list_runs(limit=10), [])
        self.assertEqual(h.list_problems(), [])
        self.assertEqual(h.list_pending(limit=10), [])
        self.assertEqual(h.list_sent(limit=10), [])

if __name__ == "__main__":
    unittest.main()
