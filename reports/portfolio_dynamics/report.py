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

import config
from common import file_discovery, ui
from reports.base import Report
from reports.portfolio_dynamics import etl, inbox, workbook


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
            "--bootstrap", action="store_true",
            help="Первый выпуск: предыдущего файла нет, история заводится с одной даты, "
                 "лимиты — нулевые (заполняются руками).",
        )
        parser.add_argument("--output", type=str, default=None, help="Путь для сохранения .xlsx")

    def run(self, args: argparse.Namespace) -> None:
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

        data = etl.build_data(t0_path, t7_path, previous_path=previous_path, bootstrap=bootstrap)

        output_path = Path(args.output) if args.output else _default_output_path(data.business_date)
        checks = workbook.evaluate_checks(data)
        workbook.save_workbook(data, output_path, checks=checks)

        failed = [cid for cid, status, _value in checks if status == "FAIL"]
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
        probe = argparse.Namespace(no_import=no_import)

        target_date = _ask_for_date(probe)
        if target_date is not None:
            # Файлы поедут из загрузок уже здесь, до спиннера: это диалог,
            # а не загрузка, и пользователь должен видеть, что куда переложили.
            t0_path, t7_path = _from_date(target_date, probe)
            ui.console.print(
                f"[grey70]Дата [bold]{target_date.isoformat()}[/bold]: "
                f"T0 = {t0_path.name}, T-7 = {t7_path.name}[/grey70]"
            )
        else:
            # Ни папок с данными, ни подходящих файлов в загрузках — плоская
            # раскладка, выбираем два файла руками, как раньше.
            ui.console.print("[bold]Срез на сегодня (T0):[/bold]")
            t0_path = file_discovery.prompt_for_file(source)
            ui.console.print("[bold]Срез на T-7:[/bold]")
            t7_path = file_discovery.prompt_for_file(config.PORTFOLIO_DYNAMICS_T7_SOURCE)

        previous = etl.find_previous_release(config.PORTFOLIO_DYNAMICS_OUTPUT_DIR)
        bootstrap = False
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
            elif answer.strip() and not answer.strip().lower().startswith("n"):
                previous = Path(answer.strip().strip('"'))
            else:
                ui.cancelled("Отчёт не сформирован: не указан предыдущий выпуск.")
                return None
        else:
            ui.console.print(f"[grey70]Предыдущий выпуск: [bold]{previous.name}[/bold][/grey70]")

        return argparse.Namespace(
            t0_input=str(t0_path), t7_input=str(t7_path), folder=None,
            date=None, no_import=no_import, t0_date=None, t7_date=None,
            previous=str(previous) if previous is not None else None,
            bootstrap=bootstrap, output=None,
        )


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


def _downloads_dir(args: argparse.Namespace):
    """Папка загрузок или None, если приёмка выключена настройкой или --no-import."""
    if getattr(args, "no_import", False):
        return None
    if not config.PORTFOLIO_DYNAMICS_IMPORT_FROM_DOWNLOADS:
        return None
    return config.DOWNLOADS_DIR


def _available_dates(args: argparse.Namespace):
    return inbox.available_dates(
        config.PORTFOLIO_DYNAMICS_T0_SOURCE, _downloads_dir(args)
    )


def _from_date(target_date, args: argparse.Namespace) -> tuple:
    """Срезы на дату: что уже в папке + недостающее из загрузок."""
    source = config.PORTFOLIO_DYNAMICS_T0_SOURCE
    plan = inbox.plan_import(
        source, _downloads_dir(args), target_date,
        archive_own_date=config.PORTFOLIO_DYNAMICS_ARCHIVE_OWN_DATE,
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
    found = file_discovery.find_dated_files(source)
    if len(found) < 2:
        raise etl.PortfolioDynamicsError(
            f"[{source.label}] В {source.directory} найдено файлов: {len(found)}, "
            "а отчёту нужны два среза (T0 и T-7). Положите обе выгрузки в папку-дату "
            "(«Настройки» -> «Создать папки по датам») или укажите файлы через "
            "--t0-input/--t7-input."
        )
    return found[0][1], found[1][1]


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
