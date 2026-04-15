# -*- coding: utf-8 -*-
from __future__ import annotations

import gzip
import logging
import os
import shutil
from logging.handlers import TimedRotatingFileHandler
from pathlib import Path
from typing import Optional

from .constants import app_dir, log_dir as get_log_dir


_LOGGING_INITIALIZED = False


def _safe_move(src: Path, dest_dir: Path) -> None:
    dest_dir.mkdir(parents=True, exist_ok=True)
    dest = dest_dir / src.name
    if dest.exists():
        for i in range(1, 1000):
            alt = dest_dir / f"{src.name}.old{i}"
            if not alt.exists():
                dest = alt
                break
    shutil.move(str(src), str(dest))


def _migrate_known_logs(old_dir: Path, new_dir: Path) -> None:
    if not old_dir or not new_dir:
        return
    old_dir = Path(old_dir)
    new_dir = Path(new_dir)
    if old_dir.resolve() == new_dir.resolve():
        return
    bases = ("system.log", "docs_sent.log", "docs_generated.log", "audit.log")
    try:
        for p in old_dir.iterdir():
            if not p.is_file():
                continue
            name = p.name
            if any(name == b or name.startswith(b + ".") for b in bases):
                try:
                    _safe_move(p, new_dir)
                except Exception:
                    pass
    except Exception:
        return


def _gzip_rotator(source: str, dest: str) -> None:
    with open(source, "rb") as sf:
        with gzip.open(dest, "wb") as df:
            while True:
                chunk = sf.read(1024 * 1024)
                if not chunk:
                    break
                df.write(chunk)
    try:
        os.remove(source)
    except Exception:
        pass


def _gzip_namer(name: str) -> str:
    return name + ".gz"


def _build_timed_handler(
    filename: Path,
    when: str,
    interval: int,
    backup_count: int,
    level: int,
    fmt: logging.Formatter,
) -> TimedRotatingFileHandler:
    filename.parent.mkdir(parents=True, exist_ok=True)
    h = TimedRotatingFileHandler(
        str(filename),
        when=when,
        interval=interval,
        backupCount=backup_count,
        encoding="utf-8",
        delay=True,
        utc=False,
    )
    h.setLevel(level)
    h.setFormatter(fmt)
    h.namer = _gzip_namer
    h.rotator = _gzip_rotator
    return h


def init_logging(log_level: int = logging.INFO, log_dir: Optional[Path] = None) -> None:
    global _LOGGING_INITIALIZED
    if _LOGGING_INITIALIZED:
        return

    if log_dir is None:
        log_dir = get_log_dir()

    try:
        _migrate_known_logs(app_dir(), log_dir)
    except Exception:
        pass

    fmt = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(name)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    root = logging.getLogger()
    root.setLevel(log_level)

    docs_sent_logger = logging.getLogger("docs_sent")
    docs_sent_logger.setLevel(logging.INFO)
    docs_sent_logger.propagate = True

    docs_generated_logger = logging.getLogger("docs_generated")
    docs_generated_logger.setLevel(logging.INFO)
    docs_generated_logger.propagate = True

    system_handler = _build_timed_handler(
        filename=log_dir / "system.log",
        when="midnight",
        interval=1,
        backup_count=30,
        level=log_level,
        fmt=fmt,
    )
    root.addHandler(system_handler)

    audit_logger = logging.getLogger("audit")
    audit_logger.setLevel(logging.INFO)
    audit_logger.propagate = False
    audit_logger.addHandler(
        _build_timed_handler(
            filename=log_dir / "audit.log",
            when="midnight",
            interval=1,
            backup_count=90,
            level=logging.INFO,
            fmt=fmt,
        )
    )

    logging.getLogger("system").info("logging inicializado dir=%s", str(log_dir))
    try:
        audit_logger.info("logging inicializado")
    except Exception:
        pass

    _LOGGING_INITIALIZED = True


def get_system_logger() -> logging.Logger:
    return logging.getLogger("system")


def get_docs_generated_logger() -> logging.Logger:
    return logging.getLogger("docs_generated")


def get_docs_sent_logger() -> logging.Logger:
    return logging.getLogger("docs_sent")

