"""Общий интерфейс отчёта для единой консоли запуска ETL (см. console.py)."""
import argparse
from abc import ABC, abstractmethod


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
        """Запускает отчёт с уже разобранными аргументами (неинтерактивный режим)."""

    @abstractmethod
    def run_interactive(self) -> None:
        """Запускает отчёт, самостоятельно запрашивая параметры через input()."""
