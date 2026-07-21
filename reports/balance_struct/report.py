"""Обёртка отчёта «Структура баланса» для единой консоли (см. console.py)."""
import argparse
from datetime import datetime
from pathlib import Path
from typing import Optional

import config
from common import file_discovery, ui
from reports.base import Report
from reports.balance_struct import etl


class BalanceStructReport(Report):
    slug = "balance-struct"
    title = "Структура баланса"
    description = "Разбор Excel-файла «ПФ_ДД_ММ_ГГГГ» по иерархии активов/пассивов в плоский CSV"

    def add_arguments(self, parser: argparse.ArgumentParser) -> None:
        parser.add_argument(
            "--date", type=str, default=None,
            help="Дата исходного файла (YYYY-MM-DD). По умолчанию — самый свежий файл в папке источника.",
        )
        parser.add_argument("--input", type=str, default=None, help="Явный путь к файлу (в обход поиска по дате)")
        parser.add_argument("--output", type=str, default=None, help="Путь для сохранения CSV")

    def run(self, args: argparse.Namespace) -> None:
        input_path = Path(args.input) if args.input else self._resolve_by_date(args.date)
        df = etl.build_report(input_path)
        output_path = Path(args.output) if args.output else _default_output_path(input_path)
        etl.save_report(df, output_path)
        ui.success(f"Готово: {len(df)} строк сохранено в {output_path}")

    def collect_interactive_args(self) -> Optional[argparse.Namespace]:
        input_path = file_discovery.prompt_for_file(config.BALANCE_STRUCT_SOURCE)
        return argparse.Namespace(date=None, input=str(input_path), output=None)

    @staticmethod
    def _resolve_by_date(date_str: str) -> Path:
        if not date_str:
            _, path = file_discovery.latest_file(config.BALANCE_STRUCT_SOURCE)
            return path
        try:
            target_date = datetime.strptime(date_str, "%Y-%m-%d").date()
        except ValueError as exc:
            raise ValueError(f"Некорректный формат даты '{date_str}', ожидается YYYY-MM-DD") from exc
        return file_discovery.resolve_file_for_date(config.BALANCE_STRUCT_SOURCE, target_date)


def _default_output_path(input_path: Path) -> Path:
    return config.BALANCE_STRUCT_OUTPUT_DIR / f"Навигатор_{Path(input_path).stem}.csv"
