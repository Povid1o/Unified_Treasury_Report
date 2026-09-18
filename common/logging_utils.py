"""Общая настройка логирования для всех ETL-отчётов.

Консольный вывод идёт через RichHandler на том же Console, что и спиннер
(см. common/ui.py) — poэтому строки лога корректно всплывают над крутящимся
спиннером вместо того, чтобы ломать его отрисовку raw-текстом.
"""
import logging
from pathlib import Path
from typing import Optional

from rich.logging import RichHandler

from common import settings
from common.ui import console


def get_logger(name: str, log_dir: Optional[Path] = None) -> logging.Logger:
    """Возвращает логгер с выводом в консоль (через rich) и в файл <logs>/<name>.log.

    log_dir не задан — берётся из настроек (пункт меню «Настройки» -> «Общее» ->
    «Папка логов»). Логгеры создаются при импорте модулей отчётов, поэтому смена
    этой настройки применяется со следующего запуска консоли.
    """
    log_dir = Path(log_dir) if log_dir is not None else Path(settings.get("logs_dir"))
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
