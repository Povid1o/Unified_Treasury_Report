"""Обёртка «Отчёта по портфелям» для единой консоли (см. console.py).

Вход — одна выгрузка «Позиция за период» на T-1 (предыдущий рабочий день).
Лежит там же, где выгрузки «Динамики портфелей»: папки-даты, плоская папка
или загрузки — и настройки источника у отчётов общие. Файл из загрузок
кладётся в папку своей даты, как это сделала бы «Динамика портфелей».

Плюс два неявных входа: предыдущий выпуск самого отчёта (из него берётся
вчерашний Open QTY — самый свежий выпуск раньше даты позиций) и RGBI, RUONIA
и RWA на сегодня — их вводят руками.
"""
import argparse
import datetime as dt
from pathlib import Path
from typing import List, Optional

from rich import box
from rich.table import Table

import config
from common import ui
from reports.base import Report
from reports.portfolio_report import etl


class PortfolioReport(Report):
    slug = "portfolio-report"
    title = "Отчёт по портфелям"
    description = ("Выгрузка позиций на T-1 -> плоский CSV для BI: Open QTY и его изменение, "
                   "PL, стоимость, DV01, Yield по портфелям + RGBI, RUONIA, RWA")

    def add_arguments(self, parser: argparse.ArgumentParser) -> None:
        parser.add_argument(
            "--date", type=str, default=None,
            help="Дата позиций (YYYY-MM-DD). По умолчанию — T-1, предыдущий рабочий день; "
                 "если на него выгрузки нет, берётся самая свежая более ранняя.",
        )
        parser.add_argument("--input", type=str, default=None,
                            help="Явный путь к выгрузке «Позиция за период»")
        parser.add_argument(
            "--previous", type=str, default=None,
            help="Предыдущий выпуск (.csv), с которым сравнивается Open QTY. По умолчанию — "
                 "самый свежий выпуск в папке результатов с датой раньше даты позиций.",
        )
        parser.add_argument("--rgbi", type=str, default=None, help="RGBI на сегодня")
        parser.add_argument("--ruonia", type=str, default=None, help="RUONIA на сегодня, %%")
        parser.add_argument("--rwa", type=str, default=None, help="RWA на сегодня")
        parser.add_argument("--no-import", action="store_true",
                            help="Не заглядывать в папку загрузок")
        parser.add_argument("--output", type=str, default=None, help="Путь для сохранения .csv")

    def run(self, args: argparse.Namespace) -> None:
        source_path = _resolve_input(args)
        snapshot_date = etl.positions.read_business_date(source_path)
        previous = Path(args.previous) if args.previous else None
        if previous is None and snapshot_date is not None:
            previous = etl.find_previous_release(config.PORTFOLIO_REPORT_OUTPUT_DIR,
                                                 before=snapshot_date)

        market = etl.MarketInputs(
            rgbi=_number(args.rgbi, "RGBI"),
            ruonia=_number(args.ruonia, "RUONIA"),
            rwa=_number(args.rwa, "RWA"),
        )
        data = etl.build_data(source_path, previous_path=previous, market=market)
        output = Path(args.output) if args.output else etl.default_output_path(data.business_date)
        etl.save_report(data, output)

        ui.success(f"Готово: {len(data.frame)} портфелей на {data.business_date.isoformat()} "
                   f"-> {output}")
        if data.previous_path is None:
            ui.warning("Предыдущего выпуска нет — изменение Open QTY появится со следующего запуска.")
        missing = data.market.missing()
        if missing:
            ui.warning(f"Не введены: {', '.join(missing)} — ячейки в отчёте пустые.")

    def collect_interactive_args(self) -> Optional[argparse.Namespace]:
        no_import = not config.PORTFOLIO_DYNAMICS_IMPORT_FROM_DOWNLOADS
        sources = etl.find_sources(etl.downloads_dir(no_import))
        if not sources:
            ui.warning(
                "Не найдено ни одной выгрузки «Позиция за период» — ни в "
                f"{config.PORTFOLIO_DYNAMICS_DIR}, ни в загрузках. Папки задаются в "
                "«Настройки» -> «Динамика портфелей»."
            )
            answer = ui.ask("Путь к файлу выгрузки (Enter — отмена)")
            if not answer.strip():
                ui.cancelled("Отчёт не сформирован: нет выгрузки.")
                return None
            source_path = Path(answer.strip().strip('"'))
        else:
            item = _ask_for_source(sources)
            # Файл едет из загрузок уже здесь, до спиннера: пользователь должен
            # видеть, что куда переложили.
            source_path = etl.take(item)

        snapshot_date = etl.positions.read_business_date(source_path)
        previous = None
        if snapshot_date is not None:
            previous = etl.find_previous_release(config.PORTFOLIO_REPORT_OUTPUT_DIR,
                                                 before=snapshot_date)
        yesterday = etl.MarketInputs()
        if previous is not None:
            ui.console.print(f"[grey70]Open QTY сравнивается с выпуском "
                             f"[bold]{previous.name}[/bold][/grey70]")
            try:
                yesterday = etl.load_previous_release(previous).market
            except etl.PortfolioReportError as exc:
                ui.warning(str(exc))
        else:
            ui.warning("Предыдущего выпуска нет — изменение Open QTY появится со следующего запуска.")

        today = dt.date.today()
        ui.console.print(f"[bold]Показатели на {today:%d.%m.%Y}[/bold] "
                         "[grey50](Enter — оставить пустым)[/grey50]")
        return argparse.Namespace(
            date=None, input=str(source_path), no_import=no_import, output=None,
            previous=str(previous) if previous is not None else None,
            rgbi=_ask_number("RGBI", yesterday.rgbi),
            ruonia=_ask_number("RUONIA, %", yesterday.ruonia),
            rwa=_ask_number("RWA", yesterday.rwa),
        )


def _number(raw: Optional[str], name: str) -> Optional[float]:
    if raw is None or not str(raw).strip():
        return None
    value = etl.positions.parse_number(raw)
    if value is None:
        raise ValueError(f"{name}: не удалось разобрать число «{raw}».")
    return value


def _ask_number(label: str, previous: Optional[float]) -> Optional[str]:
    """Спрашивает число, пока не введут разбираемое или не оставят пустым."""
    hint = f" [grey50](вчера: {previous:,.2f})[/grey50]" if previous is not None else ""
    while True:
        answer = ui.ask(f"{label}{hint}").strip()
        if not answer:
            return None
        if etl.positions.parse_number(answer) is not None:
            return answer
        ui.warning(f"«{answer}» — не число. Например: 14.25 или 14,25.")


def _sources_table(sources: List[etl.SourceFile], default: dt.date) -> Table:
    table = Table(title="Отчёт по портфелям — на какие даты есть выгрузки",
                  title_style="bold cyan", box=box.SIMPLE_HEAVY, header_style="bold cyan")
    table.add_column("#", justify="right", style="bold yellow", no_wrap=True, width=3)
    table.add_column("Дата", style="bold white", no_wrap=True)
    table.add_column("Откуда", no_wrap=True)
    table.add_column("Файл", style="grey70", overflow="fold")
    for number, item in enumerate(sources, start=1):
        mark = "  ← T-1" if item.business_date == default else ""
        origin = (f"[bold yellow]{item.origin}[/bold yellow]"
                  if item.origin == etl.ORIGIN_DOWNLOADS else item.origin)
        table.add_row(str(number), item.business_date.isoformat() + mark, origin, item.path.name)
    return table


def _ask_for_source(sources: List[etl.SourceFile]) -> etl.SourceFile:
    t_minus_1 = etl.previous_business_day(dt.date.today())
    shown = sources[:10]
    ui.console.print(_sources_table(shown, t_minus_1))
    default = next((item for item in sources if item.business_date == t_minus_1), None)
    if default is None:
        default = shown[0]
        ui.warning(f"Выгрузки на T-1 ({t_minus_1.isoformat()}) нет — по умолчанию самая "
                   f"свежая, на {default.business_date.isoformat()}.")
    answer = ui.ask(f"Дата позиций: номер или YYYY-MM-DD "
                    f"(Enter — {default.business_date.isoformat()})").strip()
    if not answer:
        return default
    if answer.isdigit() and 1 <= int(answer) <= len(shown):
        return shown[int(answer) - 1]
    try:
        target = dt.datetime.strptime(answer, "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError(f"Не удалось разобрать '{answer}': ни номер из списка, "
                         "ни дата YYYY-MM-DD.") from exc
    return etl.pick_source(sources, target)


def _resolve_input(args: argparse.Namespace) -> Path:
    if args.input:
        return Path(args.input)
    sources = etl.find_sources(etl.downloads_dir(getattr(args, "no_import", False)))
    if args.date:
        try:
            target = dt.datetime.strptime(args.date.strip(), "%Y-%m-%d").date()
        except ValueError as exc:
            raise ValueError(f"Некорректный формат даты '{args.date}', ожидается YYYY-MM-DD") from exc
        return etl.take(etl.pick_source(sources, target))
    return etl.take(etl.pick_source(sources, etl.previous_business_day(dt.date.today()),
                                    strict=False))
