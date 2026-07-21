"""Обёртка отчёта «ОВП» для единой консоли запуска (см. console.py)."""
import argparse
from pathlib import Path
from typing import Optional

from common import ui
from reports.base import Report
from reports.ovp import etl


def _default_output_path(input_path: Path) -> Path:
    return etl.OUTPUT_DIR / f"{Path(input_path).stem}_converted.csv"


class OvpReport(Report):
    slug = "ovp"
    title = "ОВП"
    description = "Преобразование Excel-отчёта «Открытая валютная позиция» (форма 634) в нормализованный CSV"

    def add_arguments(self, parser: argparse.ArgumentParser) -> None:
        parser.add_argument("--input", required=True, help="Путь к исходному Excel-файлу (.xlsx)")
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

    def run(self, args: argparse.Namespace) -> None:
        input_path = Path(args.input)
        output_path = Path(args.output) if args.output else _default_output_path(input_path)

        df = etl.convert_ovp_report(
            input_path=input_path,
            full_history_currencies=args.full_history,
            currencies=args.currencies,
        )
        etl.save_report(df, output_path)
        ui.success(f"Готово: {len(df)} строк сохранено в {output_path}")

    def collect_interactive_args(self) -> Optional[argparse.Namespace]:
        input_str = ui.ask("Путь к исходному Excel-файлу (.xlsx)").strip('"')
        input_path = Path(input_str)

        sheets = etl.list_currency_sheets(input_path)
        if not sheets:
            ui.warning("Валютные листы не найдены (все листы содержат 'СВОД' в названии).")
            return None
        ui.console.print(f"Найдены валютные листы: [bold]{', '.join(sheets)}[/bold]")

        full_history_str = ui.ask(
            "Для каких валют выгрузить ВСЕ даты, через запятую (Enter = только последняя дата для всех)"
        )
        full_history = [c.strip() for c in full_history_str.split(",") if c.strip()] if full_history_str else []

        default_output = _default_output_path(input_path)
        output_str = ui.ask("Путь для сохранения CSV", default=str(default_output))

        return argparse.Namespace(
            input=str(input_path),
            output=output_str,
            full_history=full_history,
            currencies=sheets,
        )
