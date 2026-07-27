"""Обёртка отчёта «ЧПД» для единой консоли (см. console.py).

Поддерживает пакетный режим: несколько файлов за один запуск — каждый в
свой CSV (по умолчанию) либо все вместе в один (--combine).
"""
import argparse
from datetime import datetime
from pathlib import Path
from typing import List, Optional

import config
from common import batch, file_discovery, ui
from reports.base import Report
from reports.chpd import etl


class ChpdReport(Report):
    slug = "chpd"
    title = "ЧПД"
    description = "Разбор Excel-файла «ЧПД YYYY MM DD» (лист Table) в плоский CSV"

    def add_arguments(self, parser: argparse.ArgumentParser) -> None:
        parser.add_argument(
            "--date", type=str, default=None,
            help="Дата исходного файла (YYYY-MM-DD). По умолчанию — самый свежий файл в папке источника.",
        )
        parser.add_argument(
            "--dates", type=str, default=None,
            help="Несколько дат через запятую — пакетный запуск (см. --combine)",
        )
        parser.add_argument("--input", type=str, default=None, help="Явный путь к одному файлу")
        parser.add_argument(
            "--inputs", type=str, nargs="*", default=None,
            help="Явные пути к нескольким файлам — пакетный запуск (см. --combine)",
        )
        parser.add_argument("--output", type=str, default=None, help="Путь для сохранения CSV (один файл)")
        parser.add_argument(
            "--combine", action="store_true",
            help="Объединить все файлы в один CSV вместо отдельного файла на каждый",
        )

    def run(self, args: argparse.Namespace) -> None:
        input_paths = self._resolve_inputs(args)
        combined_output = Path(args.output) if args.output else _combined_output_path()

        if args.output and not args.combine and len(input_paths) > 1:
            ui.warning("--output задан без --combine при нескольких файлах — используются имена по умолчанию для каждого файла.")

        batch.run_batch(
            input_paths=input_paths,
            build_report=etl.build_report,
            save_report=etl.save_report,
            default_output_path=_default_output_path,
            combine=args.combine,
            combined_output_path=combined_output,
        )

    def collect_interactive_args(self) -> Optional[argparse.Namespace]:
        paths = file_discovery.prompt_for_multiple_files(config.CHPD_SOURCE)
        combine = False
        if len(paths) > 1:
            combine = ui.ask("Объединить все файлы в один CSV? (y/N)", default="N").strip().lower().startswith("y")
        return argparse.Namespace(
            date=None, dates=None,
            input=None, inputs=[str(p) for p in paths],
            output=None, combine=combine,
        )

    @staticmethod
    def _resolve_inputs(args: argparse.Namespace) -> List[Path]:
        if args.inputs:
            return [Path(p) for p in args.inputs]
        if args.input:
            return [Path(args.input)]
        if args.dates:
            dates = [_parse_date(d) for d in args.dates.split(",") if d.strip()]
            return [file_discovery.resolve_file_for_date(config.CHPD_SOURCE, d) for d in dates]
        if args.date:
            return [file_discovery.resolve_file_for_date(config.CHPD_SOURCE, _parse_date(args.date))]
        _, path = file_discovery.latest_file(config.CHPD_SOURCE)
        return [path]


def _parse_date(date_str: str):
    try:
        return datetime.strptime(date_str.strip(), "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError(f"Некорректный формат даты '{date_str}', ожидается YYYY-MM-DD") from exc


def _default_output_path(input_path: Path) -> Path:
    return config.CHPD_OUTPUT_DIR / f"{Path(input_path).stem}_converted.csv"


def _combined_output_path() -> Path:
    from datetime import date as _date
    return config.CHPD_OUTPUT_DIR / f"combined_{_date.today().isoformat()}_converted.csv"
