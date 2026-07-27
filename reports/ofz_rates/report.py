"""Обёртка отчёта «Ставки ОФЗ» для единой консоли запуска (см. console.py).

Поддерживает три режима выбора периода: одна дата (+ глубина истории назад),
явный интервал дат, список конкретных (не обязательно смежных) дат.
"""
import argparse
from datetime import date, datetime
from pathlib import Path
from typing import List, Optional

from common import ui
from reports.base import Report
from reports.ofz_rates import etl


def _parse_date(value: str) -> date:
    try:
        return datetime.strptime(value.strip(), "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError(f"Некорректный формат даты '{value}', ожидается YYYY-MM-DD") from exc


def _parse_date_list(value: str) -> List[date]:
    return [_parse_date(v) for v in value.split(",") if v.strip()]


class OfzRatesReport(Report):
    slug = "ofz-rates"
    title = "Ставки ОФЗ"
    description = "Кривая доходности ОФЗ из CBonds API в формате BI (date_, name_group, name_st, unit, type_val, fvalue)"

    def add_arguments(self, parser: argparse.ArgumentParser) -> None:
        parser.add_argument(
            "--date", type=str, default=None,
            help="Дата, на которую формируется отчёт (YYYY-MM-DD, вместе с --lookback-days). По умолчанию — сегодня.",
        )
        parser.add_argument(
            "--lookback-days", type=int, default=etl.LOOKBACK_DAYS,
            help=f"Глубина истории в календарных днях назад от --date (по умолчанию {etl.LOOKBACK_DAYS}).",
        )
        parser.add_argument("--date-from", type=str, default=None, help="Начало интервала (YYYY-MM-DD)")
        parser.add_argument("--date-to", type=str, default=None, help="Конец интервала (YYYY-MM-DD)")
        parser.add_argument(
            "--dates", type=str, default=None,
            help="Список конкретных дат через запятую (YYYY-MM-DD,YYYY-MM-DD,...) вместо интервала",
        )
        parser.add_argument(
            "--output", type=str, default=None,
            help=f"Путь для сохранения CSV (по умолчанию {etl.OUTPUT_PATH}).",
        )

    def run(self, args: argparse.Namespace) -> None:
        output_path = Path(args.output) if args.output else etl.OUTPUT_PATH

        if getattr(args, "dates", None):
            df = etl.build_report(dates=_parse_date_list(args.dates))
        elif getattr(args, "date_from", None) or getattr(args, "date_to", None):
            if not (args.date_from and args.date_to):
                raise ValueError("Нужно указать оба значения: --date-from и --date-to")
            df = etl.build_report(date_from=_parse_date(args.date_from), date_to=_parse_date(args.date_to))
        else:
            as_of = _parse_date(args.date) if args.date else None
            df = etl.build_report(as_of_date=as_of, lookback_days=args.lookback_days)

        etl.save_report(df, output_path)
        ui.success(f"Готово: {len(df)} строк сохранено в {output_path}")

    def collect_interactive_args(self) -> Optional[argparse.Namespace]:
        mode = ui.ask(
            "Режим периода: 1) одна дата + история назад  2) интервал дат  3) список дат",
            default="1",
        )

        if mode == "2":
            date_from = _parse_date(ui.ask("Начало интервала (YYYY-MM-DD)"))
            date_to = _parse_date(ui.ask("Конец интервала (YYYY-MM-DD)"))
            return argparse.Namespace(
                date=None, lookback_days=etl.LOOKBACK_DAYS,
                date_from=date_from.isoformat(), date_to=date_to.isoformat(),
                dates=None, output=None,
            )

        if mode == "3":
            dates_str = ui.ask("Даты через запятую (YYYY-MM-DD,YYYY-MM-DD,...)")
            _parse_date_list(dates_str)  # валидируем сразу, чтобы не тратить время на спиннер впустую
            return argparse.Namespace(
                date=None, lookback_days=etl.LOOKBACK_DAYS,
                date_from=None, date_to=None, dates=dates_str, output=None,
            )

        date_str = ui.ask("Дата отчёта YYYY-MM-DD (Enter = сегодня)")
        if date_str:
            _parse_date(date_str)
        lookback_str = ui.ask("Глубина истории в днях", default=str(etl.LOOKBACK_DAYS))
        return argparse.Namespace(
            date=date_str or None, lookback_days=int(lookback_str),
            date_from=None, date_to=None, dates=None, output=None,
        )
