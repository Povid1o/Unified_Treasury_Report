"""Общая настройка логирования для всех ETL-отчётов."""
import logging
from pathlib import Path


def get_logger(name: str, log_dir: Path) -> logging.Logger:
    """Возвращает логгер с выводом в консоль и в файл logs/<name>.log."""
    log_dir.mkdir(parents=True, exist_ok=True)

    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)

    if not logger.handlers:
        fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")

        stream_handler = logging.StreamHandler()
        stream_handler.setFormatter(fmt)
        logger.addHandler(stream_handler)

        file_handler = logging.FileHandler(log_dir / f"{name}.log", encoding="utf-8")
        file_handler.setFormatter(fmt)
        logger.addHandler(file_handler)

    return logger
