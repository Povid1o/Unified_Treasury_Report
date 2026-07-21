"""Общий интерфейс отчёта для единой консоли запуска ETL (см. console.py).

run() и сбор интерактивных параметров разделены нарочно: run() — это
"загрузка" (сеть/файлы/pandas), её можно обернуть спиннером и прервать
по Ctrl+C; сбор параметров — это input()-диалог с пользователем, во время
которого спиннер не нужен и не должен показываться.
"""
import argparse
from abc import ABC, abstractmethod
from typing import Optional


class Report(ABC):
    """Каждый отчёт регистрируется в консоли как реализация этого интерфейса."""

    slug: str
    title: str
    description: str

    @abstractmethod
    def add_arguments(self, parser: argparse.ArgumentParser) -> None:
        """Регистрирует аргументы командной строки, специфичные для отчёта."""

    @abstractmethod
    def run(self, args: argparse.Namespace) -> None:
        """Выполняет отчёт (без input()) — эта часть оборачивается спиннером
        и может быть прервана пользователем в любой момент."""

    @abstractmethod
    def collect_interactive_args(self) -> Optional[argparse.Namespace]:
        """Запрашивает параметры отчёта через input() и возвращает Namespace,
        пригодный для run(). Возвращает None, если пользователь отменил ввод
        или запуск невозможен (например, во входном файле нет нужных данных)."""
