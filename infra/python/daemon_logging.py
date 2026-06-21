"""One-line structlog setup for IPC daemons (JSON to stderr)."""

from __future__ import annotations

import sys
from typing import Callable

_configured: str | None = None


def _add_daemon_key(service: str) -> Callable:
    def processor(_logger, _method_name, event_dict: dict) -> dict:
        event_dict["daemon"] = service
        return event_dict

    return processor


def configure_daemon_logging(service: str) -> None:
    """Idempotent: safe to call once per process with the daemon name."""
    global _configured
    if _configured is not None:
        return
    _configured = service

    import structlog

    structlog.configure(
        processors=[
            structlog.processors.add_log_level,
            structlog.processors.TimeStamper(fmt="iso", utc=False, key="ts"),
            _add_daemon_key(service),
            structlog.processors.JSONRenderer(),
        ],
        logger_factory=structlog.PrintLoggerFactory(file=sys.stderr),
        cache_logger_on_first_use=True,
    )


def get_logger():
    import structlog

    return structlog.get_logger()
