"""Общая настройка логирования для всех ETL-отчётов.

Консольный вывод идёт через RichHandler на том же Console, что и спиннер
(см. common/ui.py) — poэтому строки лога корректно всплывают над крутящимся
спиннером вместо того, чтобы ломать его отрисовку raw-текстом.
"""
import logging
from pathlib import Path

from rich.logging import RichHandler

from common.ui import console


def get_logger(name: str, log_dir: Path) -> logging.Logger:
    """Возвращает логгер с выводом в консоль (через rich) и в файл logs/<name>.log."""
    log_dir.mkdir(parents=True, exist_ok=True)

    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)

    if not logger.handlers:
        stream_handler = RichHandler(
            console=console,
            show_time=True,
            show_path=False,
            markup=False,
            rich_tracebacks=True,
        )
        stream_handler.setFormatter(logging.Formatter("%(message)s"))
        logger.addHandler(stream_handler)

        file_handler = logging.FileHandler(log_dir / f"{name}.log", encoding="utf-8")
        file_handler.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s"))
        logger.addHandler(file_handler)

    return logger
