"""Обёртка «Отчёта по портфелям» для единой консоли (см. console.py).

Вход — одна выгрузка «Позиция за период» с начала года по T-1 (предыдущий
рабочий день). Лежит там же, где выгрузки «Динамики портфелей»: папки-даты,
плоская папка или загрузки — и настройки источника у отчётов общие. Файл из
загрузок кладётся в папку своей даты: «Динамике портфелей» он тоже годится.

Плюс неявные входы: комментарии из xlsx-витрины последнего выпуска и RGBI,
RUONIA и RWA на сегодня — их вводят руками, и каждый запуск дописывает их в
файл истории (market.py). «Дополнительные портфели» — из Excel-файла, их P&L
с начала года вводится тут же и копится в своей истории (manual.py).
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
from reports.portfolio_report import etl, manual, market, workbook


class PortfolioReport(Report):
    slug = "portfolio-report"
    title = "Отчёт по портфелям"
    description = ("Выгрузка позиций с начала года по T-1 -> плоский CSV для BI: Open QTY и "
                   "его изменение с начала года, "
                   "PL, стоимость, DV01, Yield, комментарии по портфелям + история "
                   "RGBI, RUONIA, RWA")

    def add_arguments(self, parser: argparse.ArgumentParser) -> None:
        parser.add_argument(
            "--date", type=str, default=None,
            help="Дата позиций (YYYY-MM-DD). По умолчанию — T-1, предыдущий рабочий день; "
                 "если на него выгрузки нет, берётся самая свежая более ранняя.",
        )
        parser.add_argument("--input", type=str, default=None,
                            help="Явный путь к выгрузке «Позиция за период» с начала года")
        parser.add_argument("--rgbi", type=str, default=None, help="RGBI на сегодня")
        parser.add_argument("--ruonia", type=str, default=None, help="RUONIA на сегодня, %%")
        parser.add_argument("--rwa", type=str, default=None, help="RWA на сегодня")
        parser.add_argument(
            "--manual-pl", type=str, default=None,
            help="P&L с начала года по «Дополнительным портфелям», млн RUB: "
                 "«КОД=число, КОД2=число». Не указан — берётся прошлое значение.",
        )
        parser.add_argument("--no-import", action="store_true",
                            help="Не заглядывать в папку загрузок")
        parser.add_argument("--output", type=str, default=None, help="Путь для сохранения .csv")
        parser.add_argument(
            "--diagnose", action="store_true",
            help="Ничего не считать и не записывать: показать, какие колонки нашлись в "
                 "выгрузке, сколько бумаг увидено и откуда берутся DV01 и Yield.",
        )

    def run(self, args: argparse.Namespace) -> None:
        if getattr(args, "diagnose", False):
            _diagnose(args)
            return
        source_path = _resolve_input(args)
        snapshot_date = etl.positions.read_business_date(source_path)

        entered = etl.MarketInputs(
            rgbi=_number(args.rgbi, "RGBI"),
            ruonia=_number(args.ruonia, "RUONIA"),
            rwa=_number(args.rwa, "RWA"),
        )
        output_dir = (Path(args.output).parent if args.output
                      else Path(config.PORTFOLIO_REPORT_OUTPUT_DIR))

        manual_records = manual.records()
        manual_codes = [r["code"] for r in manual_records]
        manual_entered = manual.parse_entered(getattr(args, "manual_pl", None))
        unknown = sorted(set(manual_entered) - set(manual_codes))
        if unknown:
            ui.warning(f"P&L введён для портфелей, которых нет в «Дополнительных портфелях»: "
                       f"{', '.join(unknown)} — пропущен.")
        pl_history = manual.read_history() if manual_codes else {}
        manual_pl = {code: value for code, value in manual_entered.items() if code in manual_codes}

        # Своя дата не считается уже выгруженной: файл за неё перезаписывается.
        history, already = market.load(output_dir, skip=snapshot_date)
        inputs, pending = entered, {}
        comments = {}
        if snapshot_date is not None:
            if manual_codes:
                manual_pl = manual.apply_inputs(pl_history, snapshot_date, manual_codes,
                                                manual_entered)
            inputs = market.apply_inputs(history, snapshot_date, entered)
            pending = market.pending(history, already, snapshot_date)
            comments_path = workbook.find_comments_release(output_dir, snapshot_date)
            if comments_path is not None:
                comments = workbook.read_comments(comments_path)
                ui.console.print(f"[grey70]Комментарии ({len(comments)}) — из "
                                 f"[bold]{comments_path.name}[/bold][/grey70]")

        data = etl.build_data(source_path, market=inputs,
                              comments=comments, market_history=pending,
                              manual_portfolios=manual_records, manual_pl=manual_pl)
        history_file = market.write_history(market.history_path(), history)
        if manual_codes and snapshot_date is not None:
            manual.write_history(pl_history)
        output = Path(args.output) if args.output else etl.default_output_path(data.business_date)
        etl.save_report(data, output)
        book = workbook.write_workbook(data, workbook.workbook_path(output))

        ui.success(f"Готово: {len(data.frame)} портфелей на {data.business_date.isoformat()} "
                   f"(изменение Open QTY — с {data.period_start:%d.%m.%Y}) -> {output}")
        ui.console.print(f"[grey70]Комментарии пишутся в жёлтой колонке {book.name}[/grey70]")
        ui.console.print(f"[grey70]История рынка: {history_file} — "
                         f"{market.dates_span(history)}[/grey70]")
        if data.market_history:
            ui.console.print(f"[grey70]В CSV добавлены показатели рынка из истории: "
                             f"{market.dates_span(data.market_history)}[/grey70]")
        if data.manual_codes:
            ui.console.print(f"[grey70]Дополнительные портфели ({len(data.manual_codes)}): "
                             f"{', '.join(data.manual_codes)} — из "
                             f"{manual.manual_store.file_path()}[/grey70]")
        missing = data.market.missing() + [
            f"P&L {code}" for code in data.manual_codes if code not in manual_pl]
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
        return ask_inputs(source_path, no_import)


def ask_inputs(source_path: Path, no_import: bool) -> argparse.Namespace:
    """Диалог показателей, которых нет в выгрузке: RGBI, RUONIA, RWA и P&L
    дополнительных портфелей. Общий с пунктом «Динамика + отчёт по портфелям»."""
    snapshot_date = etl.positions.read_business_date(source_path)

    # Подсказки при вводе — из истории: последнее значение до даты позиций и
    # уже записанное на неё (при повторном прогоне его не нужно вводить снова).
    yesterday, recorded = etl.MarketInputs(), etl.MarketInputs()
    if snapshot_date is not None:
        try:
            history, _ = market.load(skip=snapshot_date)
            yesterday = market.latest_before(history, snapshot_date)
            recorded = etl.MarketInputs(**history.get(snapshot_date, {}))
        except etl.PortfolioReportError as exc:
            ui.warning(str(exc))

    today = dt.date.today()
    target = (f" — запишутся на дату позиций {snapshot_date:%d.%m.%Y}"
              if snapshot_date is not None else "")
    ui.console.print(f"[bold]Показатели на {today:%d.%m.%Y}[/bold]{target} "
                     "[grey50](Enter — оставить записанное или пустым)[/grey50]")
    rgbi = _ask_number("RGBI", yesterday.rgbi, recorded.rgbi)
    ruonia = _ask_number("RUONIA, %", yesterday.ruonia, recorded.ruonia)
    rwa = _ask_number("RWA", yesterday.rwa, recorded.rwa)
    return argparse.Namespace(
        date=None, input=str(source_path), no_import=no_import, output=None,
        rgbi=rgbi, ruonia=ruonia, rwa=rwa,
        manual_pl=_ask_manual_pl(snapshot_date),
    )


def _number(raw: Optional[str], name: str) -> Optional[float]:
    if raw is None or not str(raw).strip():
        return None
    value = etl.positions.parse_number(raw)
    if value is None:
        raise ValueError(f"{name}: не удалось разобрать число «{raw}».")
    return value


def _ask_number(label: str, previous: Optional[float],
                recorded: Optional[float] = None) -> Optional[str]:
    """Спрашивает число, пока не введут разбираемое или не оставят пустым.

    Пусто — None: тогда в выпуск идёт значение, уже записанное в историю на
    эту дату (recorded), если оно есть.
    """
    hints = []
    if previous is not None:
        hints.append(f"прошлое: {previous:,.2f}")
    if recorded is not None:
        hints.append(f"уже записано: {recorded:,.2f}, Enter — оставить")
    hint = f" [grey50]({'; '.join(hints)})[/grey50]" if hints else ""
    while True:
        answer = ui.ask(f"{label}{hint}").strip()
        if not answer:
            return None
        if etl.positions.parse_number(answer) is not None:
            return answer
        ui.warning(f"«{answer}» — не число. Например: 14.25 или 14,25.")


def _ask_manual_pl(snapshot_date: Optional[dt.date]) -> Optional[str]:
    """P&L с начала года по каждому «Дополнительному портфелю», млн RUB.

    Enter — как вчера: run() возьмёт уже записанное на дату или последнее до неё.
    """
    records = manual.records()
    if not records:
        return None
    previous, recorded = {}, {}
    if snapshot_date is not None:
        try:
            history = manual.read_history()
            previous = manual.latest_before(history, snapshot_date)
            recorded = history.get(snapshot_date, {})
        except etl.PortfolioReportError as exc:
            ui.warning(str(exc))
    ui.console.print("[bold]P&L с начала года по дополнительным портфелям, млн RUB[/bold] "
                     "[grey50](Enter — как вчера)[/grey50]")
    answers = []
    for record in records:
        code = record["code"]
        label = code if record["name"] == code else f"{code} ({record['name']})"
        answer = _ask_number(label, previous.get(code), recorded.get(code))
        if answer is not None:
            answers.append(f"{code}={answer}")
    return "; ".join(answers) or None


def _sources_table(sources: List[etl.SourceFile], default: dt.date) -> Table:
    table = Table(title="Отчёт по портфелям — на какие даты есть выгрузки",
                  title_style="bold cyan", box=box.SIMPLE_HEAVY, header_style="bold cyan")
    table.add_column("#", justify="right", style="bold yellow", no_wrap=True, width=3)
    table.add_column("Дата", style="bold white", no_wrap=True)
    table.add_column("Период с", no_wrap=True)
    table.add_column("Откуда", no_wrap=True)
    table.add_column("Файл", style="grey70", overflow="fold")
    for number, item in enumerate(sources, start=1):
        mark = "  ← T-1" if item.business_date == default else ""
        origin = (f"[bold yellow]{item.origin}[/bold yellow]"
                  if item.origin == etl.ORIGIN_DOWNLOADS else item.origin)
        table.add_row(str(number), item.business_date.isoformat() + mark,
                      item.period_start.strftime("%d.%m.%Y") if item.period_start else "?",
                      origin, item.path.name)
    return table


def _ask_for_source(sources: List[etl.SourceFile]) -> etl.SourceFile:
    t_minus_1 = etl.previous_business_day(dt.date.today())
    # Годятся только выгрузки «01.01 - дата»: за один день изменение Open QTY
    # всегда ноль, с другой начальной даты — не с начала года. Остальные не
    # показываем, чтобы не выбрать по ошибке.
    usable = [item for item in sources if item.from_year_start]
    if not usable:
        raise etl.PortfolioReportError(
            "Нашлись только выгрузки «Позиция за период» не с начала года — отчёту нужна "
            f"выгрузка с начала года («01.01.{t_minus_1.year} - {t_minus_1:%d.%m.%Y}»).")
    skipped = len(sources) - len(usable)
    if skipped:
        ui.console.print(f"[grey50]Выгрузок не с начала года (за день и т.п.) не показано: "
                         f"{skipped}[/grey50]")
    shown = usable[:10]
    ui.console.print(_sources_table(shown, t_minus_1))
    default = next((item for item in usable if item.business_date == t_minus_1), None)
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


def _diagnose(args: argparse.Namespace) -> None:
    """--diagnose: файл берётся тем же выбором, что и при запуске, но не перекладывается."""
    if args.input:
        path = Path(args.input)
    else:
        sources = etl.find_sources(etl.downloads_dir(getattr(args, "no_import", False)))
        if args.date:
            target = dt.datetime.strptime(args.date.strip(), "%Y-%m-%d").date()
            path = etl.pick_source(sources, target).path
        else:
            path = etl.pick_source(sources, etl.previous_business_day(dt.date.today()),
                                   strict=False).path
    for line in etl.diagnose(path):
        ui.console.print(line, markup=False, highlight=False)
