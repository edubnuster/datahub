# -*- coding: utf-8 -*-
from __future__ import annotations

import gzip
import logging
import os
from logging.handlers import TimedRotatingFileHandler
from pathlib import Path
from typing import Optional

from .constants import app_dir


_LOGGING_INITIALIZED = False


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
        log_dir = app_dir()

    fmt = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] %(name)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    root = logging.getLogger()
    root.setLevel(log_level)

    system_handler = _build_timed_handler(
        filename=log_dir / "system.log",
        when="W0",
        interval=1,
        backup_count=12,
        level=log_level,
        fmt=fmt,
    )
    root.addHandler(system_handler)

    docs_sent_logger = logging.getLogger("docs_sent")
    docs_sent_logger.setLevel(logging.INFO)
    docs_sent_logger.propagate = False
    docs_sent_logger.addHandler(
        _build_timed_handler(
            filename=log_dir / "docs_sent.log",
            when="W0",
            interval=1,
            backup_count=52,
            level=logging.INFO,
            fmt=fmt,
        )
    )

    docs_generated_logger = logging.getLogger("docs_generated")
    docs_generated_logger.setLevel(logging.INFO)
    docs_generated_logger.propagate = False
    docs_generated_logger.addHandler(
        _build_timed_handler(
            filename=log_dir / "docs_generated.log",
            when="midnight",
            interval=1,
            backup_count=7,
            level=logging.INFO,
            fmt=fmt,
        )
    )

    logging.getLogger("system").info("logging inicializado dir=%s", str(log_dir))
    try:
        docs_sent_logger.info("logging inicializado")
    except Exception:
        pass
    try:
        docs_generated_logger.info("logging inicializado")
    except Exception:
        pass

    _LOGGING_INITIALIZED = True


def get_system_logger() -> logging.Logger:
    return logging.getLogger("system")


def get_docs_generated_logger() -> logging.Logger:
    return logging.getLogger("docs_generated")


def get_docs_sent_logger() -> logging.Logger:
    return logging.getLogger("docs_sent")

