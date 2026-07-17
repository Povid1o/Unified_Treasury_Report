"""Обёртка отчёта «Ставки ОФЗ» для единой консоли запуска (см. console.py)."""
import argparse
from datetime import date, datetime
from pathlib import Path

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
        print(f"Готово: {len(df)} строк сохранено в {output_path}")

    def run_interactive(self) -> None:
        date_str = input("Дата отчёта (YYYY-MM-DD, Enter = сегодня): ").strip()
        as_of = _parse_date(date_str) if date_str else None

        lookback_str = input(f"Глубина истории в днях (Enter = {etl.LOOKBACK_DAYS}): ").strip()
        lookback_days = int(lookback_str) if lookback_str else etl.LOOKBACK_DAYS

        df = etl.build_report(as_of_date=as_of, lookback_days=lookback_days)
        output_path = etl.save_report(df)
        print(f"Готово: {len(df)} строк сохранено в {output_path}")
