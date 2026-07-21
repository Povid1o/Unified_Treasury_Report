"""Обёртка отчёта «Трансфертные ставки» для единой консоли (см. console.py).

Отчёту нужны ДВА независимых исходных файла (короткие и длинные ставки),
каждый со своей датой в имени файла — поэтому источник файла запрашивается
дважды (см. config.TRANSFERT_SHORT_SOURCE / TRANSFERT_LONG_SOURCE).
"""
import argparse
from datetime import date as _date
from datetime import datetime
from pathlib import Path
from typing import Optional

import config
from common import file_discovery, ui
from reports.base import Report
from reports.transfert_stavka import etl


class TransfertStavkaReport(Report):
    slug = "transfert-stavka"
    title = "Трансфертные ставки"
    description = "Короткие (до 3М) + длинные (свыше 3М) трансфертные ставки из двух Excel-файлов в формат BI"

    def add_arguments(self, parser: argparse.ArgumentParser) -> None:
        parser.add_argument(
            "--short-date", type=str, default=None,
            help="Дата файла коротких ставок (YYYY-MM-DD). По умолчанию — самый свежий.",
        )
        parser.add_argument(
            "--long-date", type=str, default=None,
            help="Дата файла длинных ставок (YYYY-MM-DD). По умолчанию — самый свежий.",
        )
        parser.add_argument("--short-input", type=str, default=None, help="Явный путь к файлу коротких ставок")
        parser.add_argument("--long-input", type=str, default=None, help="Явный путь к файлу длинных ставок")
        parser.add_argument("--output", type=str, default=None, help="Путь для сохранения CSV")

    def run(self, args: argparse.Namespace) -> None:
        short_path = Path(args.short_input) if args.short_input else _resolve_by_date(
            config.TRANSFERT_SHORT_SOURCE, args.short_date
        )
        long_path = Path(args.long_input) if args.long_input else _resolve_by_date(
            config.TRANSFERT_LONG_SOURCE, args.long_date
        )

        df = etl.build_report(short_path, long_path)
        output_path = Path(args.output) if args.output else _default_output_path()
        etl.save_report(df, output_path)
        ui.success(f"Готово: {len(df)} строк сохранено в {output_path}")

    def collect_interactive_args(self) -> Optional[argparse.Namespace]:
        ui.console.print("[bold]Файл коротких ставок (до 3М):[/bold]")
        short_path = file_discovery.prompt_for_file(config.TRANSFERT_SHORT_SOURCE)
        ui.console.print("[bold]Файл длинных ставок (свыше 3М):[/bold]")
        long_path = file_discovery.prompt_for_file(config.TRANSFERT_LONG_SOURCE)

        return argparse.Namespace(
            short_input=str(short_path), long_input=str(long_path),
            short_date=None, long_date=None, output=None,
        )


def _resolve_by_date(source: file_discovery.SourceConfig, date_str: str) -> Path:
    if not date_str:
        _, path = file_discovery.latest_file(source)
        return path
    try:
        target_date = datetime.strptime(date_str, "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError(f"Некорректный формат даты '{date_str}', ожидается YYYY-MM-DD") from exc
    return file_discovery.resolve_file_for_date(source, target_date)


def _default_output_path() -> Path:
    return config.TRANSFERT_OUTPUT_DIR / f"result_transert_{_date.today().isoformat()}.csv"
