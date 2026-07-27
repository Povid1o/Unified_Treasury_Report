"""Обёртка отчёта «ОВП» для единой консоли запуска (см. console.py).

У файлов ОВП нет даты в имени, поэтому источник (config.OVP_SOURCE) сортирует
их по дате изменения, а не по regex из имени (см. common/file_discovery.py).
Если папка из конфига недоступна или в ней ничего не нашлось — консоль сама
предложит вписать путь(и) вручную.
"""
import argparse
from datetime import datetime
from pathlib import Path
from typing import List, Optional

import config
from common import batch, file_discovery, ui
from reports.base import Report
from reports.ovp import etl


class OvpReport(Report):
    slug = "ovp"
    title = "ОВП"
    description = "Преобразование Excel-отчёта «Открытая валютная позиция» (форма 634) в нормализованный CSV"

    def add_arguments(self, parser: argparse.ArgumentParser) -> None:
        parser.add_argument(
            "--date", type=str, default=None,
            help="Дата изменения исходного файла (YYYY-MM-DD). По умолчанию — самый свежий файл в папке источника.",
        )
        parser.add_argument(
            "--dates", type=str, default=None,
            help="Несколько дат изменения через запятую — пакетный запуск (см. --combine)",
        )
        parser.add_argument("--input", default=None, help="Путь к одному исходному Excel-файлу (.xlsx)")
        parser.add_argument(
            "--inputs", nargs="*", default=None,
            help="Пути к нескольким исходным Excel-файлам — пакетный запуск (см. --combine)",
        )
        parser.add_argument(
            "--output", default=None,
            help="Путь для сохранения CSV (по умолчанию — output/ovp/<имя файла>_converted.csv)",
        )
        parser.add_argument(
            "--full-history", nargs="*", default=[], metavar="CURRENCY",
            help="Валютные листы, для которых нужно выгрузить все найденные даты "
                 "(по умолчанию для всех листов берётся только последняя дата)",
        )
        parser.add_argument(
            "--currencies", nargs="*", default=None, metavar="CURRENCY",
            help="Ограничить обработку конкретными листами (по умолчанию — все, кроме СВОД)",
        )
        parser.add_argument(
            "--combine", action="store_true",
            help="Объединить все файлы в один CSV вместо отдельного файла на каждый",
        )

    def run(self, args: argparse.Namespace) -> None:
        input_paths = self._resolve_inputs(args)
        combined_output = Path(args.output) if args.output else _combined_output_path()

        if args.output and not args.combine and len(input_paths) > 1:
            ui.warning("--output задан без --combine при нескольких файлах — используются имена по умолчанию для каждого файла.")

        def build_one(path: Path):
            return etl.convert_ovp_report(
                input_path=path,
                full_history_currencies=args.full_history,
                currencies=args.currencies,
            )

        batch.run_batch(
            input_paths=input_paths,
            build_report=build_one,
            save_report=etl.save_report,
            default_output_path=_default_output_path,
            combine=args.combine,
            combined_output_path=combined_output,
        )

    def collect_interactive_args(self) -> Optional[argparse.Namespace]:
        paths = file_discovery.prompt_for_multiple_files(config.OVP_SOURCE)

        full_history_str = ui.ask(
            "Для каких валют выгрузить ВСЕ даты, через запятую (Enter = только последняя дата для всех)"
        )
        full_history = [c.strip() for c in full_history_str.split(",") if c.strip()] if full_history_str else []

        combine = False
        if len(paths) > 1:
            combine = ui.ask("Объединить все файлы в один CSV? (y/N)", default="N").strip().lower().startswith("y")

        return argparse.Namespace(
            date=None, dates=None,
            input=None, inputs=[str(p) for p in paths], output=None,
            full_history=full_history, currencies=None, combine=combine,
        )

    @staticmethod
    def _resolve_inputs(args: argparse.Namespace) -> List[Path]:
        if args.inputs:
            return [Path(p) for p in args.inputs]
        if args.input:
            return [Path(args.input)]
        if args.dates:
            dates = [_parse_date(d) for d in args.dates.split(",") if d.strip()]
            return [file_discovery.resolve_file_for_date(config.OVP_SOURCE, d) for d in dates]
        if args.date:
            return [file_discovery.resolve_file_for_date(config.OVP_SOURCE, _parse_date(args.date))]
        _, path = file_discovery.latest_file(config.OVP_SOURCE)
        return [path]


def _parse_date(date_str: str):
    try:
        return datetime.strptime(date_str.strip(), "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError(f"Некорректный формат даты '{date_str}', ожидается YYYY-MM-DD") from exc


def _default_output_path(input_path: Path) -> Path:
    return config.OVP_OUTPUT_DIR / f"{Path(input_path).stem}_converted.csv"


def _combined_output_path() -> Path:
    from datetime import date as _date
    return config.OVP_OUTPUT_DIR / f"combined_{_date.today().isoformat()}_converted.csv"
