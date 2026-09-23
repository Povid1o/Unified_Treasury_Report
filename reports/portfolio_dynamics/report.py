"""Обёртка отчёта «Динамика портфелей» для единой консоли (см. console.py).

Отчёту нужны ДВА исходных файла (срез T0 и срез T-7). Штатная раскладка —
папка-дата: внутри папки исходных файлов создаётся подпапка вида 2026-09-18,
в неё кладутся обе выгрузки, и отчёт без аргументов сам берёт самую свежую
папку, В КОТОРОЙ ЕСТЬ файлы (пустые, созданные заранее, пропускаются), сам
разбирается, какой из двух файлов T0, а какой T-7 (по дате самой выгрузки), и
кладёт результат в output. Папки на несколько дней вперёд создаются разом:
«Настройки» -> «Создать папки по датам».

Если в папке нужной даты срезов не хватает, отчёт ищет их в папке загрузок и
раскладывает сам (см. inbox.py) — носить файлы руками не нужно. Отключается
настройкой «Забирать выгрузки из загрузок».

Плоская раскладка (все выгрузки за все даты в одной папке) продолжает
работать: при выключенных папках-датах или когда подходящей папки нет, файл
запрашивается дважды — как у «Трансфертных ставок».

Плюс третий, неявный вход:
предыдущий выпуск самого отчёта, откуда переносятся история объёмов по типам,
лимиты и заметки (см. reports/portfolio_dynamics/etl.py). По умолчанию он
ищется сам — самый свежий xlsx в выходной папке.

Единственный отчёт, который отдаёт .xlsx вместо .csv: выход — это шаблон
обмена схемы v3.0 с формулами, витринами и проверками (CONTRACT.md).
"""
import argparse
from datetime import datetime
from pathlib import Path
from typing import Optional

from rich import box
from rich.table import Table
from rich.text import Text

import config
from common import file_discovery, ui
from reports.base import Report
from reports.portfolio_dynamics import etl, history, inbox, workbook


class PortfolioDynamicsReport(Report):
    slug = "portfolio-dynamics"
    title = "Динамика портфелей"
    description = "Два среза выгрузки позиций (T0 и T-7) -> xlsx схемы v3.0 с историей, лимитами и проверками"

    def add_arguments(self, parser: argparse.ArgumentParser) -> None:
        parser.add_argument(
            "--t0-date", type=str, default=None,
            help="Дата среза T0 (YYYY-MM-DD) — вторая дата в имени файла. По умолчанию — самый свежий.",
        )
        parser.add_argument(
            "--t7-date", type=str, default=None,
            help="Дата среза T-7 (YYYY-MM-DD) — вторая дата в имени файла. По умолчанию — самый свежий.",
        )
        parser.add_argument("--t0-input", type=str, default=None, help="Явный путь к файлу среза T0")
        parser.add_argument("--t7-input", type=str, default=None, help="Явный путь к файлу среза T-7")
        parser.add_argument(
            "--date", type=str, default=None,
            help="Отчётная дата (YYYY-MM-DD): взять папку этой даты, а недостающие срезы "
                 "найти в загрузках. По умолчанию — последняя доступная дата.",
        )
        parser.add_argument(
            "--no-import", action="store_true",
            help="Не трогать папку загрузок: работать только с тем, что уже лежит в папке отчёта.",
        )
        parser.add_argument(
            "--folder", type=str, default=None,
            help="Папка-дата с обеими выгрузками: дата YYYY-MM-DD, имя папки или полный путь. "
                 "По умолчанию — самая свежая папка, в которой есть файлы.",
        )
        parser.add_argument(
            "--previous", type=str, default=None,
            help="Предыдущий выпуск отчёта (.xlsx), из которого переносятся история, лимиты и "
                 "заметки. По умолчанию — самый свежий файл в выходной папке.",
        )
        parser.add_argument(
            "--history", type=str, default=None,
            help="Отчёт старого формата, из которого подтянуть накопленную историю "
                 "объёмов по типам (листы «Динамика AFS/HTM/TSS»). По умолчанию — путь "
                 "из настроек, если он задан.",
        )
        parser.add_argument(
            "--bootstrap", action="store_true",
            help="Первый выпуск: предыдущего файла нет, история заводится с одной даты, "
                 "лимиты — нулевые (заполняются руками).",
        )
        parser.add_argument("--output", type=str, default=None, help="Путь для сохранения .xlsx")
        parser.add_argument(
            "--diagnose", action="store_true",
            help="Ничего не считать: показать, какие папки и файлы отчёт видит и почему "
                 "не находит данные.",
        )

    def run(self, args: argparse.Namespace) -> None:
        if getattr(args, "diagnose", False):
            diagnose(args)
            return

        t0_path, t7_path = _resolve_slice_paths(args)
        if t0_path == t7_path:
            raise etl.PortfolioDynamicsError(
                f"Срезы T0 и T-7 указывают на один и тот же файл ({t0_path.name}). "
                "Укажите разные даты (--t0-date/--t7-date) или пути (--t0-input/--t7-input)."
            )

        bootstrap = bool(getattr(args, "bootstrap", False))
        previous_path = Path(args.previous) if args.previous else None
        if previous_path is None and not bootstrap:
            previous_path = etl.find_previous_release(config.PORTFOLIO_DYNAMICS_OUTPUT_DIR)

        limits_path = _resolve_limits_path(t0_path, args)
        data = etl.build_data(t0_path, t7_path, previous_path=previous_path,
                              bootstrap=bootstrap, limits_path=limits_path,
                              history_path=_resolve_history_path(args))

        output_path = Path(args.output) if args.output else _default_output_path(data.business_date)
        checks = workbook.evaluate_checks(data)
        workbook.save_workbook(data, output_path, checks=checks)

        failed = [cid for cid, status, _value in checks if status == "FAIL"]
        if limits_path is None:
            ui.warning(
                "Файл лимитов не найден — лимиты и границы зон взяты из предыдущего "
                "выпуска. Чтобы понять, почему он не нашёлся, запустите "
                "«python console.py portfolio-dynamics --diagnose»."
            )
        ui.success(
            f"Готово: {len(data.fact_portfolio_snapshot)} портфелей, "
            f"{len(data.fact_type_daily)} строк истории -> {output_path}"
        )
        if failed:
            ui.warning(
                f"Лист checks: FAIL в {len(failed)} проверках ({', '.join(failed)}) — "
                "загрузка в BI заблокирована, см. лист checks в файле."
            )

    def collect_interactive_args(self) -> Optional[argparse.Namespace]:
        source = config.PORTFOLIO_DYNAMICS_T0_SOURCE
        no_import = not config.PORTFOLIO_DYNAMICS_IMPORT_FROM_DOWNLOADS
        probe = argparse.Namespace(no_import=no_import, t0_input=None, t7_input=None,
                                   folder=None, date=None, t0_date=None, t7_date=None)

        # Файлы поедут из загрузок уже здесь, до спиннера: это диалог, а не
        # загрузка, и пользователь должен видеть, что куда переложили.
        target_date = _ask_for_date(probe)
        probe.date = target_date.isoformat() if target_date else None
        try:
            # Один и тот же путь разрешения, что и у CLI: спросить дату —
            # единственное, что интерактив добавляет. Иначе интерактив не видел
            # бы, например, плоской раскладки, которую CLI прекрасно находит.
            t0_path, t7_path = _resolve_slice_paths(probe)
        except etl.PortfolioDynamicsError as exc:
            # Молча сваливаться в «выберите файл» нельзя: человек не узнает, ни
            # где искали, ни почему не нашли, — сначала показываем разбор.
            ui.warning(str(exc))
            answer = ui.ask("Указать пути к двум файлам вручную? (y/N)", default="N")
            if not answer.strip().lower().startswith("y"):
                ui.cancelled("Отчёт не сформирован. Запустите "
                             "«python console.py portfolio-dynamics --diagnose», "
                             "чтобы увидеть, что именно видно отчёту.")
                return None
            ui.console.print("[bold]Срез на сегодня (T0):[/bold]")
            t0_path = file_discovery.prompt_for_file(source)
            ui.console.print("[bold]Срез на T-7:[/bold]")
            t7_path = file_discovery.prompt_for_file(config.PORTFOLIO_DYNAMICS_T7_SOURCE)
        else:
            ui.console.print(
                f"[grey70]T0 = {t0_path.name}, T-7 = {t7_path.name}[/grey70]"
            )

        previous = etl.find_previous_release(config.PORTFOLIO_DYNAMICS_OUTPUT_DIR)
        bootstrap = False
        history = None
        if previous is None:
            ui.warning(
                "Предыдущий выпуск отчёта не найден в "
                f"{config.PORTFOLIO_DYNAMICS_OUTPUT_DIR}. История объёмов по типам, лимиты и "
                "заметки берутся только оттуда — без него получится ПЕРВЫЙ выпуск: история с "
                "одной датой, нулевые лимиты, пустые заметки."
            )
            answer = ui.ask("Создать первый выпуск? (y/N) или путь к предыдущему файлу", default="N")
            if answer.strip().lower().startswith("y"):
                bootstrap = True
                history = _ask_for_history()
            elif answer.strip() and not answer.strip().lower().startswith("n"):
                previous = Path(answer.strip().strip('"'))
            else:
                ui.cancelled("Отчёт не сформирован: не указан предыдущий выпуск.")
                return None
        else:
            ui.console.print(f"[grey70]Предыдущий выпуск: [bold]{previous.name}[/bold][/grey70]")

        return argparse.Namespace(
            t0_input=str(t0_path), t7_input=str(t7_path), folder=None, diagnose=False,
            date=None, no_import=no_import, t0_date=None, t7_date=None,
            previous=str(previous) if previous is not None else None,
            bootstrap=bootstrap, history=history, output=None,
        )


def diagnose(args: argparse.Namespace) -> None:
    """Показывает, что отчёт видит: пути, папки-даты, содержимое загрузок.

    Нужно ровно для одного вопроса — «почему не находит данные». Отвечает на
    него по шагам: существуют ли папки, какие подпапки-даты есть и сколько в
    них файлов, какие файлы в загрузках подошли под шаблон, а какие нет и
    почему.
    """
    source = config.PORTFOLIO_DYNAMICS_T0_SOURCE
    data_dir = Path(source.directory)
    downloads = Path(config.DOWNLOADS_DIR)

    def mark(path: Path) -> str:
        return "[bold green]есть[/bold green]" if path.is_dir() else "[bold red]НЕ НАЙДЕНА[/bold red]"

    ui.console.print()
    ui.console.print("[bold]Пути[/bold]")
    ui.console.print(f"  исходная папка : {data_dir}  {mark(data_dir)}")
    ui.console.print(f"  папка загрузок : {downloads}  {mark(downloads)}")
    ui.console.print(f"  папка выгрузки : {config.PORTFOLIO_DYNAMICS_OUTPUT_DIR}")
    ui.console.print(f"  папки по датам : "
                     f"{'включены' if source.uses_date_folders else 'выключены'}"
                     f" (формат {config.DATE_FOLDER_FORMAT})")
    ui.console.print(f"  приёмка из загрузок: "
                     f"{'включена' if config.PORTFOLIO_DYNAMICS_IMPORT_FROM_DOWNLOADS else 'выключена'}"
                     f", режим: {'перенос' if config.PORTFOLIO_DYNAMICS_MOVE_FROM_DOWNLOADS else 'копирование'}")
    ui.console.print(Text.assemble(
        "  шаблон имени файла: ", (str(source.filename_regex), "grey70")))

    folders = file_discovery.find_date_folders(source)
    ui.console.print()
    ui.console.print(f"[bold]Папки-даты в исходной папке: {len(folders)}[/bold]")
    # Какие именно срезы лежат внутри каждой папки: дата ПАПКИ и даты файлов в
    # ней — разные вещи (в папке 18.09 лежит и срез за 11.09), а пару отчёт
    # собирает по датам файлов, где бы они ни лежали.
    slices_by_folder = {}
    for candidate in inbox.slices_in_other_date_folders(
            source, Path(""), config.PORTFOLIO_DYNAMICS_LIMITS_SOURCE):
        slices_by_folder.setdefault(candidate.path.parent, set()).add(
            candidate.business_date.isoformat())
    for folder in folders[:10]:
        count = len(folder.files)
        here = sorted(slices_by_folder.get(folder.path, ()), reverse=True)
        if count >= 2:
            state = "[bold green]готова[/bold green]"
        elif count:
            state = "[grey50]пара доберётся из других папок или загрузок[/grey50]"
        else:
            state = "[grey50]пусто[/grey50]"
        srezy = ("  срезы: %s" % ", ".join(here)) if here else ""
        ui.console.print(f"  {folder.path.name}  файлов: {count}{srezy}  {state}")
    if not folders:
        ui.console.print("  [grey50]нет — их создаёт «Настройки» → «Создать папки по датам», "
                         "либо приёмка из загрузок создаст нужную сама[/grey50]")

    ui.console.print()
    _print_loose_files(source, data_dir)

    ui.console.print()
    _print_downloads(source, downloads)

    ui.console.print()
    _print_limits(downloads, data_dir)

    ui.console.print()
    _print_history_source()

    rows = _available_dates(args)
    ui.console.print()
    if rows:
        ui.console.print(_dates_table(rows))
        ui.console.print(f"[bold green]Запуск без аргументов возьмёт дату "
                         f"{rows[0].date.isoformat()}.[/bold green]")
    else:
        ui.console.print("[bold red]Ни одной даты, на которую можно построить отчёт.[/bold red]")
        ui.console.print("[grey70]Отчёту нужны ДВА файла: на отчётную дату и на более раннюю "
                         "(T-7). Одного файла недостаточно.[/grey70]")


def _print_history_source() -> None:
    """Откуда возьмётся накопленная история объёмов по типам."""
    configured = config.PORTFOLIO_DYNAMICS_HISTORY_FILE
    previous = etl.find_previous_release(config.PORTFOLIO_DYNAMICS_OUTPUT_DIR)
    ui.console.print("[bold]История объёмов по типам (лист fact_type_daily)[/bold]")
    if previous is not None:
        ui.console.print(f"  [bold green]✓[/bold green] переносится из предыдущего "
                         f"выпуска: {previous.name}")
    else:
        ui.console.print("  [yellow]![/yellow] предыдущего выпуска нет — история начнётся "
                         "с одной даты")
    if configured and str(configured).strip():
        found = history.resolve_path(configured, _history_search_dirs())
        if found is None:
            ui.console.print(f"  [bold red]✗[/bold red] импорт из отчёта старого формата: "
                             f"{configured} — файл не найден, отчёт соберётся без него")
        else:
            ui.console.print(f"  [bold green]✓[/bold green] импорт из отчёта старого "
                             f"формата: {found}")
    else:
        ui.console.print("  [grey50]импорт из отчёта старого формата не настроен "
                         "(«Файл с историей», либо аргумент --history)[/grey50]")
        # Раз не настроен — сразу показываем, что можно было бы взять: иначе о
        # самой возможности узнать неоткуда.
        found = history.find_candidates(_history_search_dirs(), limit=5)
        for path, sheets in found:
            ui.console.print(Text.assemble(
                ("  подходит: ", "grey70"), (path.name, "bold"),
                ("  [%s]" % ", ".join(sheets), "grey50"),
                ("\n            %s" % path.parent, "grey50"),
            ))
        if found:
            ui.console.print("  [grey50]задаётся в «Настройки» → «и — история из "
                             "старого отчёта» или аргументом --history[/grey50]")


def _print_limits(downloads: Path, data_dir: Path) -> None:
    """Файлы лимитов — третий вход отчёта, и их тоже надо видеть.

    Ищутся по своему шаблону имени (дата через подчёркивания) и берутся строго
    на отчётную дату: файл лимитов на другую дату не подойдёт.
    """
    limits_source = config.PORTFOLIO_DYNAMICS_LIMITS_SOURCE
    ui.console.print("[bold]Файлы лимитов[/bold]")
    ui.console.print(Text.assemble(
        "  шаблон имени: ", (str(limits_source.filename_regex), "grey70")))

    found = []
    if downloads.is_dir():
        found += [("загрузки", c) for c in inbox.scan_limits(limits_source, downloads)]
    if data_dir.is_dir():
        for folder in file_discovery.find_date_folders(limits_source):
            files = [f for f in folder.files if etl.is_limits_file(f)]
            found += [("папка " + folder.path.name, c)
                      for c in inbox._dated(files, limits_source)]

    if not found:
        ui.console.print("  [grey50]не найдено ни одного — лимиты будут перенесены "
                         "из предыдущего выпуска[/grey50]")
    for where, candidate in sorted(found, key=lambda item: item[1].business_date, reverse=True)[:10]:
        ui.console.print(Text.assemble(
            ("  ✓ ", "bold green"), candidate.path.name,
            (f"   на дату {candidate.business_date.isoformat()}, {where}", "grey70")))

    allocations = etl.parse_nested_limits()
    parents = etl.parse_type_parents()
    nesting = ", ".join("{} в {}".format(child, parent)
                        for child, parent in parents.items())
    allocated = ", ".join("{} = {:,.0f}".format(name, value)
                          for name, value in allocations.items())
    ui.console.print("  [grey50]вложенность типов: %s[/grey50]"
                     % (nesting or "не задана"))
    ui.console.print("  [grey50]выделено вложенным: %s[/grey50]"
                     % (allocated or "не задано (объём вложенного типа "
                                     "складывается с объемлющим)"))


def _print_loose_files(source: file_discovery.SourceConfig, data_dir: Path) -> None:
    """Файлы, лежащие прямо в папке исходных файлов (не в подпапке-дате).

    Самый частый повод удивиться: «файлы же на месте» — а имена не подходят под
    шаблон, и отчёт их не видит. Вердикт по каждому отвечает на это сразу.
    """
    if not data_dir.is_dir():
        return
    loose = sorted(f for f in data_dir.glob(source.glob_pattern)
                   if f.is_file() and not f.name.startswith("~$"))
    if not loose:
        return

    matched = {c.path for c in _match_by_name(source, loose)}
    ui.console.print(f"[bold]Файлы прямо в папке исходных файлов: {len(loose)}, "
                     f"подошли под шаблон: {len(matched)}[/bold]")
    for path in loose[:15]:
        if path in matched:
            ui.console.print(Text.assemble(("  ✓ ", "bold green"), path.name))
        else:
            ui.console.print(Text.assemble(
                ("  ✗ ", "grey50"), (f"{path.name} — имя не подходит под шаблон", "grey50")))
    if len(loose) > 15:
        ui.console.print(f"  [grey50]… и ещё {len(loose) - 15}[/grey50]")


def _match_by_name(source: file_discovery.SourceConfig, paths):
    """Какие из файлов подходят под шаблон имени источника."""
    import re
    if not source.filename_regex:
        return [inbox.Candidate(path=p, business_date=None) for p in paths]
    pattern = re.compile(source.filename_regex)
    found = []
    for path in paths:
        match = pattern.search(path.name)
        if not match:
            continue
        try:
            import datetime as _dt
            parsed = _dt.datetime.strptime(match.group(1), source.date_format).date()
        except ValueError:
            continue
        found.append(inbox.Candidate(path=path, business_date=parsed))
    return found


def _print_downloads(source: file_discovery.SourceConfig, downloads: Path) -> None:
    """Список .xlsx в загрузках с вердиктом по каждому — подошёл под шаблон или нет."""
    if not downloads.is_dir():
        ui.console.print(f"[bold]Загрузки[/bold]: папка не найдена ({downloads})")
        return

    everything = sorted(f for f in downloads.glob("*.xlsx")
                        if f.is_file() and not f.name.startswith("~$"))
    by_name = {c.path: c.business_date
               for c in inbox.scan_slices(source, downloads, deep=False)}
    matched = {c.path: c.business_date for c in inbox.scan_slices(source, downloads)}
    # Файлы, опознанные по содержимому, показываем отдельно: иначе непонятно,
    # почему отчёт берёт файл с «неправильным» именем.
    by_content = {path: value for path, value in matched.items() if path not in by_name}
    ui.console.print(f"[bold]Файлы .xlsx в загрузках: {len(everything)}, "
                     f"подошли под шаблон: {len(by_name)}, "
                     f"опознаны по содержимому: {len(by_content)}[/bold]")
    if not everything:
        ui.console.print("  [grey50]пусто (ищется только верхний уровень папки, не подпапки)[/grey50]")
    # Сначала подошедшие: в папке на тысячу файлов иначе их не видно.
    shown = sorted(matched, key=lambda f: matched[f], reverse=True)
    shown += [f for f in everything if f not in matched][:15]
    for path in shown:
        if path in by_name:
            ui.console.print(Text.assemble(
                ("  ✓ ", "bold green"), path.name,
                (f"   дата среза: {matched[path].isoformat()}", "grey70")))
        elif path in by_content:
            ui.console.print(Text.assemble(
                ("  ✓ ", "bold yellow"), path.name,
                (f"   дата среза: {matched[path].isoformat()}"
                 " — по содержимому, имя под шаблон не подходит", "grey70")))
        else:
            ui.console.print(Text.assemble(
                ("  ✗ ", "grey50"), (f"{path.name} — ни имя, ни содержимое не подошли", "grey50")))
    hidden = len(everything) - len([f for f in shown if f in everything])
    if hidden > 0:
        ui.console.print(f"  [grey50]… и ещё {hidden}[/grey50]")


def _dates_table(rows) -> Table:
    table = Table(
        title="Динамика портфелей — на какие даты есть данные",
        title_style="bold cyan", box=box.SIMPLE_HEAVY, show_header=True,
        header_style="bold cyan",
        caption="«загрузки» — файлы будут взяты из папки загрузок и разложены "
                "по папкам-датам автоматически.",
        caption_style="grey50",
    )
    table.add_column("#", justify="right", style="bold yellow", no_wrap=True, width=3)
    table.add_column("Дата", style="bold white", no_wrap=True)
    table.add_column("Откуда", no_wrap=True)
    table.add_column("Файлы", style="grey70", overflow="fold", ratio=2)
    for i, row in enumerate(rows, start=1):
        origin = row.origin if row.origin == "папка" else f"[bold yellow]{row.origin}[/bold yellow]"
        table.add_row(str(i), row.date.isoformat(), origin, ", ".join(row.files))
    return table


def _history_search_dirs() -> list:
    """Где искать отчёт старого формата: загрузки, папка отчёта, папка выгрузки."""
    return [Path(config.DOWNLOADS_DIR), Path(config.PORTFOLIO_DYNAMICS_DIR),
            Path(config.PORTFOLIO_DYNAMICS_OUTPUT_DIR)]


def _pick_history_file() -> Optional[str]:
    """Показывает найденные файлы старого формата и даёт выбрать номером.

    Искать по имени бесполезно: у старого отчёта оно произвольное. Зато листы
    «Динамика <ТИП>» видны в оглавлении книги мгновенно, поэтому кандидаты
    ищутся по содержимому — человеку остаётся выбрать номер, а не вспоминать
    и набирать путь.
    """
    candidates = history.find_candidates(_history_search_dirs())
    if candidates:
        ui.console.print("[bold]Похожие файлы (листы «Динамика <ТИП>»):[/bold]")
        for number, (path, sheets) in enumerate(candidates, start=1):
            ui.console.print(Text.assemble(
                ("  %d) " % number, "bold"), (path.name, ""),
                ("  [%s]" % ", ".join(sheets), "grey50"),
                ("\n     %s" % path.parent, "grey50"),
            ))
        answer = ui.ask("Номер файла, путь к своему файлу или Enter — не подтягивать")
    else:
        ui.console.print("[grey50]Автоматически ничего похожего не нашлось "
                         "(искали в загрузках, папке отчёта и папке выгрузки).[/grey50]")
        answer = ui.ask("Путь к файлу с историей (Enter — не подтягивать)")

    token = answer.strip().strip('"')
    if not token:
        return None
    if token.isdigit() and 1 <= int(token) <= len(candidates):
        return str(candidates[int(token) - 1][0])
    return token


def _ask_for_history() -> Optional[str]:
    """Предлагает подтянуть накопленную историю из отчёта старого формата.

    Спрашивается только на ПЕРВОМ выпуске: дальше история накапливается сама и
    живёт в предыдущем файле отчёта.
    """
    configured = config.PORTFOLIO_DYNAMICS_HISTORY_FILE
    if configured and str(configured).strip():
        ui.console.print(f"[grey70]История будет подтянута из {configured} "
                         "(настройка «Файл с историей»).[/grey70]")
        return None  # путь возьмётся из настройки

    ui.console.print(
        "[grey70]Лист истории объёмов по типам (fact_type_daily) у первого выпуска "
        "пуст: она накапливается по одной дате за запуск. Если история уже ведётся "
        "в отчёте старого формата (листы «Динамика AFS», «Динамика HTM», "
        "«Динамика TSS»), её можно подтянуть оттуда.[/grey70]"
    )
    return _pick_history_file()


def _ask_for_date(args: argparse.Namespace):
    """Спрашивает отчётную дату. None — подходящих дат не нашлось совсем."""
    rows = _available_dates(args)
    if not rows:
        return None

    ui.console.print(_dates_table(rows))
    latest = rows[0]
    choice = ui.ask(
        f"Отчётная дата: номер из списка или YYYY-MM-DD "
        f"(Enter — последняя, {latest.date.isoformat()})"
    )
    if not choice:
        return latest.date

    token = choice.strip()
    if token.isdigit() and 1 <= int(token) <= len(rows):
        return rows[int(token) - 1].date
    try:
        return datetime.strptime(token, "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError(
            f"Не удалось разобрать '{token}': ни номер из списка, ни дата YYYY-MM-DD."
        ) from exc


def _resolve_slice_paths(args: argparse.Namespace) -> tuple:
    """Определяет пути к срезам T0 и T-7 — в порядке убывания явности.

    1. --t0-input/--t7-input — оба явно;
    2. --folder — папка-дата (дата, имя или путь), внутри разбираемся сами;
    3. --t0-date/--t7-date — по датам в плоской раскладке;
    4. по умолчанию — самая свежая папка-дата с файлами, а если папок нет,
       два самых свежих файла в плоской папке.
    """
    if args.t0_input and args.t7_input:
        return Path(args.t0_input), Path(args.t7_input)
    if args.t0_input or args.t7_input:
        raise etl.PortfolioDynamicsError(
            "--t0-input и --t7-input задаются только вместе: отчёту нужны оба среза. "
            "Чтобы взять оба файла из одной папки-даты, используйте --folder."
        )

    source = config.PORTFOLIO_DYNAMICS_T0_SOURCE

    if args.folder:
        folder = file_discovery.resolve_date_folder(source, args.folder)
        if len(folder.files) < 2:
            # Папка названа явно, но неполна — недостающее ищем в загрузках.
            return _from_date(folder.date, args)
        return _from_folder(folder)

    if getattr(args, "date", None):
        try:
            target = datetime.strptime(args.date.strip(), "%Y-%m-%d").date()
        except ValueError as exc:
            raise ValueError(f"Некорректный формат даты '{args.date}', ожидается YYYY-MM-DD") from exc
        return _from_date(target, args)

    if args.t0_date or args.t7_date:
        if not (args.t0_date and args.t7_date):
            raise etl.PortfolioDynamicsError(
                "--t0-date и --t7-date задаются только вместе. Если оба среза лежат "
                "в одной папке-дате, укажите её через --folder."
            )
        return (
            _resolve_by_date(source, args.t0_date),
            _resolve_by_date(config.PORTFOLIO_DYNAMICS_T7_SOURCE, args.t7_date),
        )

    # Последняя доступная дата — по папкам И по загрузкам сразу. Проверять
    # сначала готовые папки нельзя: свежая пара, лежащая в загрузках, тогда
    # проиграла бы старой укомплектованной папке, и «по умолчанию последняя»
    # молча пересобирало бы позавчерашний отчёт.
    dates = _available_dates(args)
    if dates:
        return _from_date(dates[0].date, args)

    folder = file_discovery.latest_folder_with_data(source, min_files=2)
    if folder is not None:
        return _from_folder(folder)

    return _two_latest_flat(source)


def _resolve_history_path(args: argparse.Namespace) -> Optional[Path]:
    """Файл старого формата для разового импорта истории: аргумент или настройка.

    Путь, заданный аргументом (или введённый только что в диалоге), обязан
    найтись — человек явно попросил этот файл. Путь из НАСТРОЙКИ — нет: импорт
    истории разовый и необязательный, и устаревшая настройка не должна валить
    весь отчёт, когда оба среза и предыдущий выпуск на месте.
    """
    explicit = getattr(args, "history", None)
    raw = explicit or config.PORTFOLIO_DYNAMICS_HISTORY_FILE
    if not raw or not str(raw).strip():
        return None
    path = history.resolve_path(raw, _history_search_dirs())
    if path is None:
        if explicit:
            raise etl.PortfolioDynamicsError(f"Файл с историей не найден: {raw}.")
        etl.logger.warning(
            "Файл с историей не найден: %s — история из отчёта старого формата не "
            "подтягивается, отчёт собирается без неё. Поправьте путь в «Настройки» → "
            "«история из старого отчёта» или очистите его.", raw,
        )
        return None
    if str(path) != str(raw).strip().strip('"'):
        etl.logger.info("Файл с историей: %s", path)
    return path


def _resolve_limits_path(t0_path: Path, args: argparse.Namespace) -> Optional[Path]:
    """Выгрузка «Состояние лимитов» на дату среза T0 — рядом с ним или из загрузок.

    Ищется на дату САМОГО среза, а не на имя папки: папку могли назвать иначе,
    а лимиты обязаны соответствовать отчётной дате.
    """
    target_date = etl.read_business_date(t0_path)
    if target_date is None:
        return None
    found = inbox.resolve_limits_file(
        config.PORTFOLIO_DYNAMICS_LIMITS_SOURCE, _downloads_dir(args),
        Path(t0_path).parent, target_date,
        move=config.PORTFOLIO_DYNAMICS_MOVE_FROM_DOWNLOADS,
    )
    if found is None:
        etl.logger.warning(
            "Файл лимитов на %s не найден (шаблон имени: %s) — лимиты и границы зон "
            "переносятся из предыдущего выпуска. Положите «Состояние лимитов на дату …» "
            "в загрузки или в папку %s, чтобы они обновились.",
            target_date.isoformat(), config.PORTFOLIO_DYNAMICS_LIMITS_REGEX,
            Path(t0_path).parent,
        )
    else:
        etl.logger.info("Файл лимитов: %s", found.name)
    return found


def _downloads_dir(args: argparse.Namespace):
    """Папка загрузок или None, если приёмка выключена настройкой или --no-import."""
    if getattr(args, "no_import", False):
        return None
    if not config.PORTFOLIO_DYNAMICS_IMPORT_FROM_DOWNLOADS:
        return None
    return config.DOWNLOADS_DIR


def _available_dates(args: argparse.Namespace):
    return inbox.available_dates(
        config.PORTFOLIO_DYNAMICS_T0_SOURCE, _downloads_dir(args),
        limits_source=config.PORTFOLIO_DYNAMICS_LIMITS_SOURCE,
    )


def _from_date(target_date, args: argparse.Namespace) -> tuple:
    """Срезы на дату: что уже в папке + недостающее из загрузок."""
    source = config.PORTFOLIO_DYNAMICS_T0_SOURCE
    plan = inbox.plan_import(
        source, _downloads_dir(args), target_date,
        archive_own_date=config.PORTFOLIO_DYNAMICS_ARCHIVE_OWN_DATE,
        limits_source=config.PORTFOLIO_DYNAMICS_LIMITS_SOURCE,
    )
    if not plan.complete:
        raise etl.PortfolioDynamicsError(plan.problem)

    files = inbox.apply_import(plan, move=config.PORTFOLIO_DYNAMICS_MOVE_FROM_DOWNLOADS)
    t0_path, t7_path = etl.split_slice_files(files, plan.folder.name)
    etl.logger.info(
        "Дата %s (%s): T0 = %s, T-7 = %s",
        target_date.isoformat(), plan.folder.name, t0_path.name, t7_path.name,
    )
    return t0_path, t7_path


def _from_folder(folder: file_discovery.DateFolder) -> tuple:
    """Разбирает содержимое папки-даты на T0 и T-7 и сверяет дату с именем папки."""
    t0_path, t7_path = etl.split_slice_files(folder.files, folder.path.name)
    etl.logger.info(
        "Папка-дата %s: T0 = %s, T-7 = %s", folder.path.name, t0_path.name, t7_path.name
    )

    # Имя папки — это дата, на которую её завели; дата отчёта берётся из самой
    # выгрузки. Расхождение означает, что файлы положили не в ту папку, — молча
    # это пропускать нельзя, отчёт уйдёт с чужой датой.
    t0_date = etl.read_business_date(t0_path)
    if t0_date is not None and t0_date != folder.date:
        etl.logger.warning(
            "Папка называется %s, а срез T0 (%s) — на %s. Отчёт будет построен на дату "
            "из выгрузки (%s); если это ошибка, переложите файлы в папку нужной даты.",
            folder.path.name, t0_path.name, t0_date.isoformat(), t0_date.isoformat(),
        )
    return t0_path, t7_path


def _two_latest_flat(source: file_discovery.SourceConfig) -> tuple:
    """Плоская раскладка: два самых свежих файла по дате из имени.

    Более свежий — T0, предыдущий — T-7. Это разумное умолчание и для тех, кто
    ещё не перешёл на папки-даты: раньше без аргументов оба источника отдавали
    ОДИН и тот же самый свежий файл, и запуск падал на проверке «срезы совпадают».
    """
    try:
        found = file_discovery.find_dated_files(source)
    except file_discovery.SourceFileError:
        found = []
    if len(found) < 2:
        raise etl.PortfolioDynamicsError(_nothing_found_message(source))
    return found[0][1], found[1][1]


def _nothing_found_message(source: file_discovery.SourceConfig) -> str:
    """Почему не нашлось ни одного среза — с проверкой обоих мест и что делать.

    Самая частая причина при первом запуске — пути остались значениями по
    умолчанию (сетевая папка Jupiter), поэтому текст показывает сами пути и
    отвечает, существуют ли они: так опечатка и непримонтированный диск видны
    сразу, а не после третьего запуска.
    """
    data_dir = Path(source.directory)
    downloads = Path(config.DOWNLOADS_DIR)
    import_on = config.PORTFOLIO_DYNAMICS_IMPORT_FROM_DOWNLOADS

    def mark(path: Path) -> str:
        return "есть" if path.is_dir() else "НЕ НАЙДЕНА"

    lines = [
        "Не нашлось двух срезов (T0 и T-7). Где искали:",
        f"  исходная папка: {data_dir} — {mark(data_dir)}",
    ]
    if import_on:
        matched = len(inbox.scan_slices(source, downloads))
        lines.append(
            f"  загрузки:       {downloads} — {mark(downloads)}, "
            f"подходящих выгрузок: {matched}"
        )
    else:
        lines.append("  загрузки:       приёмка выключена настройкой "
                     "«Забирать выгрузки из загрузок»")
    lines += [
        f"  шаблон имени:   {source.filename_regex}",
        "",
        "Что делать:",
        "  python console.py portfolio-dynamics --diagnose   # показать, что именно видно",
        '  python console.py settings --set portfolio_dynamics_dir="<путь>" '
        'downloads_dir="<путь>"',
        "  либо указать файлы напрямую: --t0-input <файл> --t7-input <файл>",
    ]
    return "\n".join(lines)


def _resolve_by_date(source: file_discovery.SourceConfig, date_str: Optional[str]) -> Path:
    if not date_str:
        _, path = file_discovery.latest_file(source)
        return path
    try:
        target_date = datetime.strptime(date_str, "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError(f"Некорректный формат даты '{date_str}', ожидается YYYY-MM-DD") from exc
    return file_discovery.resolve_file_for_date(source, target_date)


def _default_output_path(business_date) -> Path:
    return (config.PORTFOLIO_DYNAMICS_OUTPUT_DIR
            / f"{etl.OUTPUT_FILENAME_PREFIX}{business_date.isoformat()}.xlsx")
