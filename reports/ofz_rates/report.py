"""Обёртка отчёта «Ставки ОФЗ» для единой консоли запуска (см. console.py)."""
import argparse
from datetime import date, datetime
from pathlib import Path
from typing import Optional

from common import ui
from reports.base import Report
from reports.ofz_rates import etl


def _parse_date(value: str) -> date:
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError(f"Некорректный формат даты '{value}', ожидается YYYY-MM-DD") from exc


class OfzRatesReport(Report):
    slug = "ofz-rates"
    title = "Ставки ОФЗ"
    description = "Кривая доходности ОФЗ из CBonds API в формате BI (date_, name_group, name_st, unit, type_val, fvalue)"

    def add_arguments(self, parser: argparse.ArgumentParser) -> None:
        parser.add_argument(
            "--date", type=str, default=None,
            help="Дата, на которую формируется отчёт (YYYY-MM-DD). По умолчанию — сегодня.",
        )
        parser.add_argument(
            "--lookback-days", type=int, default=etl.LOOKBACK_DAYS,
            help=f"Глубина истории в календарных днях назад от --date (по умолчанию {etl.LOOKBACK_DAYS}).",
        )
        parser.add_argument(
            "--output", type=str, default=None,
            help=f"Путь для сохранения CSV (по умолчанию {etl.OUTPUT_PATH}).",
        )

    def run(self, args: argparse.Namespace) -> None:
        as_of = _parse_date(args.date) if args.date else None
        output_path = Path(args.output) if args.output else etl.OUTPUT_PATH

        df = etl.build_report(as_of_date=as_of, lookback_days=args.lookback_days)
        etl.save_report(df, output_path)
        ui.success(f"Готово: {len(df)} строк сохранено в {output_path}")

    def collect_interactive_args(self) -> Optional[argparse.Namespace]:
        date_str = ui.ask("Дата отчёта YYYY-MM-DD (Enter = сегодня)")
        if date_str:
            _parse_date(date_str)  # валидируем сразу, чтобы не тратить время на спиннер впустую

        lookback_str = ui.ask("Глубина истории в днях", default=str(etl.LOOKBACK_DAYS))

        return argparse.Namespace(date=date_str or None, lookback_days=int(lookback_str), output=None)
