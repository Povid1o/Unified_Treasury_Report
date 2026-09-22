"""Пункт меню «Настройки»: правка путей и констант отчётов из консоли.

Отделено от common/settings.py по той же причине, по какой в отчётах разделены
etl.py и report.py: там — хранилище и проверка значений без всякого ввода-вывода,
здесь — диалог с пользователем. Модуль импортирует config (чтобы пересобрать
его константы после сохранения), поэтому config его импортировать не может и
не должен.

Ctrl+C на любом шаге отменяет текущий шаг и возвращает на предыдущий экран —
как и везде в консоли; наружу KeyboardInterrupt пробрасывается только с самого
верхнего экрана, где его ловит console.py.
"""
from pathlib import Path
from typing import List, Optional

from rich import box
from rich.table import Table
from rich.text import Text

import config
from common import file_discovery, settings, ui

MENU_TITLE = "Настройки"
MENU_DESCRIPTION = "Пути к папкам, шаблоны имён файлов и константы отчётов — без правки config.py"


MANUAL_PORTFOLIOS_KEY = "portfolio_dynamics_manual_portfolios"


def _value_repr(key: str) -> str:
    value = settings.get(key)
    if settings.SETTINGS_BY_KEY[key].kind == "portfolios":
        if not value:
            return "не заданы"
        return f"{len(value)} шт.: " + ", ".join(p["code"] for p in value)
    return str(value)


def _groups_table() -> Table:
    table = Table(
        title=f"Настройки — {settings.settings_path()}",
        title_style="bold cyan", box=box.ROUNDED, show_header=True,
        header_style="bold cyan", border_style="grey50",
    )
    table.add_column("#", justify="right", style="bold yellow", no_wrap=True)
    table.add_column("Раздел", style="bold white", no_wrap=True)
    table.add_column("Настроек", justify="right", style="grey70")
    table.add_column("Изменено", justify="right", style="grey70")

    for i, group in enumerate(settings.GROUPS, start=1):
        items = settings.SETTINGS_BY_GROUP[group.key]
        changed = sum(1 for s in items if settings.is_overridden(s.key))
        table.add_row(
            str(i), group.title, str(len(items)),
            f"[bold yellow]{changed}[/bold yellow]" if changed else "—",
        )
    return table


def _settings_table(group: settings.Group) -> Table:
    # Три колонки, а не четыре: ключ показывается второй строкой в той же
    # ячейке, что и название. На узком терминале (80 знаков) лишняя колонка
    # съедает всю ширину у значения, и путь превращается в столбик по букве.
    table = Table(
        title=group.title, title_style="bold cyan", box=box.SIMPLE_HEAVY,
        show_header=True, header_style="bold cyan",
        caption="Жёлтым — изменённые значения; остальные взяты по умолчанию.",
        caption_style="grey50",
    )
    table.add_column("#", justify="right", style="bold yellow", no_wrap=True, width=3)
    table.add_column("Настройка", style="bold white", max_width=28)
    table.add_column("Значение", style="grey70", overflow="fold", ratio=2)

    for i, setting in enumerate(settings.SETTINGS_BY_GROUP[group.key], start=1):
        value = _value_repr(setting.key)
        if settings.is_overridden(setting.key):
            value = f"[bold yellow]{value}[/bold yellow]"
        table.add_row(str(i), f"{setting.label}\n[grey50]{setting.key}[/grey50]", value)
    return table


def _paths_table() -> Table:
    table = Table(
        title="Проверка путей", title_style="bold cyan", box=box.SIMPLE_HEAVY,
        show_header=True, header_style="bold cyan",
        caption="Недоступная папка — обычно не примонтированный сетевой диск "
                "или опечатка в пути.",
        caption_style="grey50",
    )
    table.add_column("Что", style="bold white", max_width=26)
    table.add_column("Путь", style="grey70", overflow="fold", ratio=2)
    table.add_column("Состояние", overflow="fold", max_width=26)

    titles = {g.key: g.title for g in settings.GROUPS}
    for status in settings.check_paths():
        if status.is_problem:
            state = f"[bold red]✗[/bold red] {status.detail}"
        elif not status.exists or status.files == 0:
            state = f"[yellow]![/yellow] {status.detail}"
        else:
            state = f"[bold green]✓[/bold green] {status.detail}"
        table.add_row(
            f"{titles[status.group]}\n[grey50]{status.label}[/grey50]",
            str(status.path), state,
        )
    return table


def _edit_setting(setting: settings.Setting) -> bool:
    """Диалог правки одной настройки. True — значение изменилось."""
    current = settings.get(setting.key)
    default = settings.default_of(setting.key)
    overridden = settings.is_overridden(setting.key)

    ui.console.print()
    ui.console.print(f"[bold white]{setting.label}[/bold white]  [grey50]({setting.key})[/grey50]")
    ui.console.print(f"[grey70]{setting.help}[/grey70]")
    ui.console.print(f"[grey50]Тип: {settings.KIND_HINTS[setting.kind]}[/grey50]")
    ui.console.print(f"Текущее значение: [bold]{current}[/bold]"
                     + ("  [yellow](изменено)[/yellow]" if overridden else "  [grey50](по умолчанию)[/grey50]"))
    if overridden:
        ui.console.print(f"[grey50]По умолчанию: {default}[/grey50]")

    raw = ui.ask("Новое значение (Enter — оставить, «-» — вернуть значение по умолчанию)")
    if not raw:
        return False
    if raw.strip() == "-":
        if not overridden:
            ui.console.print("[grey70]Значение и так по умолчанию — ничего не меняем.[/grey70]")
            return False
        settings.reset(setting.key)
        ui.success(f"{setting.key}: возвращено значение по умолчанию ({settings.get(setting.key)})")
        return True

    try:
        value = settings.set_value(setting.key, raw)
    except settings.SettingsError as exc:
        ui.error(str(exc))
        return False

    ui.success(f"{setting.key} = {value}")
    _warn_if_missing(setting, value)
    return True


def _warn_if_missing(setting: settings.Setting, value) -> None:
    """Предупреждает о несуществующем пути, но НЕ отвергает его.

    Сетевой диск может быть не примонтирован прямо сейчас, а путь при этом
    правильный — запрещать ввод в такой ситуации было бы вредно.
    """
    if setting.kind not in ("dir", "file"):
        return
    path = Path(value)
    if path.exists():
        return
    what = "Папка" if setting.kind == "dir" else "Файл"
    ui.warning(f"{what} не существует: {path}. Значение сохранено — проверьте, "
               "примонтирован ли диск и нет ли опечатки.")


def _edit_group(group: settings.Group) -> bool:
    """Экран одного раздела. True — что-то изменилось."""
    items = settings.SETTINGS_BY_GROUP[group.key]
    changed = False
    while True:
        ui.console.print()
        ui.console.print(_settings_table(group))
        try:
            choice = ui.ask("Номер настройки для правки (0 — назад)", default="0")
        except KeyboardInterrupt:
            ui.console.print()
            return changed

        if choice in ("0", ""):
            return changed
        try:
            setting = items[int(choice) - 1]
        except (ValueError, IndexError):
            ui.warning("Некорректный выбор, попробуйте снова.")
            continue

        try:
            changed |= _edit_setting(setting)
        except KeyboardInterrupt:
            ui.console.print()
            ui.cancelled("Правка настройки отменена.")


# ── Дополнительные портфели ──────────────────────────────────────────────────
def _portfolios_table(records) -> Table:
    table = Table(
        title="Дополнительные портфели", title_style="bold cyan", box=box.SIMPLE_HEAVY,
        show_header=True, header_style="bold cyan",
        caption="Портфели, которых нет в выгрузке позиций; объём ведётся вручную.",
        caption_style="grey50",
    )
    table.add_column("#", justify="right", style="bold yellow", no_wrap=True, width=3)
    table.add_column("Код", style="bold white", no_wrap=True)
    table.add_column("Название", style="grey70", overflow="fold")
    table.add_column("Тип", no_wrap=True)
    table.add_column("Объём, млн", justify="right", no_wrap=True)
    table.add_column("Дюрация", justify="right", no_wrap=True)
    for i, record in enumerate(records, start=1):
        duration = ("—" if record["duration"] is None else f"{record['duration']:.2f}")
        table.add_row(str(i), record["code"], record["name"], record["type"],
                      f"{record['volume']:,.0f}", duration)
    return table


def _ask_portfolio(existing: Optional[dict] = None) -> Optional[dict]:
    """Диалог одной записи. None — пользователь отказался."""
    known = ", ".join(settings.KNOWN_TYPES_HINT)
    current = existing or {}

    code = ui.ask("Код портфеля (латиницей, например OFZ_EXTRA)",
                  default=current.get("code", "")).strip().upper()
    if not code:
        ui.cancelled("Код не указан — запись не добавлена.")
        return None

    name = ui.ask("Название (Enter — совпадает с кодом)",
                  default=current.get("name", "")).strip() or code
    portfolio_type = ui.ask(f"Тип портфеля ({known})",
                            default=current.get("type", "")).strip().upper()
    if not portfolio_type:
        ui.cancelled("Тип не указан — запись не добавлена.")
        return None

    volume = ui.ask("Объём, млн RUB",
                    default=("" if not current else str(current.get("volume", "")))).strip()
    if not volume:
        ui.cancelled("Объём не указан — запись не добавлена.")
        return None

    ui.console.print("[grey70]Дюрацию можно не указывать, но тогда проверка CHK_14 "
                     "(«пустая текущая дюрация в срезе») покажет FAIL.[/grey70]")
    duration_default = "" if not current or current.get("duration") is None else str(current["duration"])
    duration = ui.ask("Текущая дюрация, лет (Enter — не указывать)",
                      default=duration_default).strip()

    return {"code": code, "name": name, "type": portfolio_type,
            "volume": volume, "duration": duration or None}


def _save_portfolios(records) -> bool:
    try:
        settings.set_value(MANUAL_PORTFOLIOS_KEY, records)
    except settings.SettingsError as exc:
        ui.error(str(exc))
        return False
    return True


def _manual_portfolios_screen() -> bool:
    """Экран добавления/правки/удаления. True — что-то изменилось."""
    changed = False
    while True:
        records = list(settings.get(MANUAL_PORTFOLIOS_KEY))
        ui.console.print()
        if records:
            ui.console.print(_portfolios_table(records))
        else:
            ui.console.print("[grey70]Дополнительных портфелей нет.[/grey70]")
        ui.console.print("[grey70]д — добавить, номер — изменить, у<номер> — удалить, "
                         "0 — назад[/grey70]")
        choice = ui.ask("Действие", default="0").strip().lower()

        if choice in ("0", ""):
            return changed

        if choice in ("д", "d"):
            record = _ask_portfolio()
            if record is None:
                continue
            if _save_portfolios(records + [record]):
                ui.success(f"Портфель {record['code']} добавлен.")
                changed = True
            continue

        if choice.startswith(("у", "u")):
            index = choice[1:].strip()
            if not index.isdigit() or not 1 <= int(index) <= len(records):
                ui.warning("Укажите номер удаляемой записи, например «у1».")
                continue
            removed = records.pop(int(index) - 1)
            if _save_portfolios(records):
                ui.success(f"Портфель {removed['code']} удалён.")
                changed = True
            continue

        if choice.isdigit() and 1 <= int(choice) <= len(records):
            position = int(choice) - 1
            record = _ask_portfolio(records[position])
            if record is None:
                continue
            records[position] = record
            if _save_portfolios(records):
                ui.success(f"Портфель {record['code']} изменён.")
                changed = True
            continue

        ui.warning("Некорректный выбор, попробуйте снова.")


# ── История из отчёта старого формата ────────────────────────────────────────
HISTORY_FILE_KEY = "portfolio_dynamics_history_file"


def _history_candidates():
    """Книги с листами «Динамика <ТИП>» в загрузках и папках отчёта."""
    from reports.portfolio_dynamics import history
    directories = [settings.get("downloads_dir"),
                   settings.get("portfolio_dynamics_dir"),
                   settings.get("portfolio_dynamics_output_dir")]
    return history.find_candidates([Path(d) for d in directories if d])


def _history_screen() -> bool:
    """Выбор файла старого формата, из которого подтянуть историю по типам.

    Отдельный пункт меню, а не просто настройка-путь: иначе возможность негде
    найти — она теряется среди двух десятков настроек отчёта. Файл здесь же и
    разбирается, чтобы сразу было видно, что из него прочиталось.
    """
    current = settings.get(HISTORY_FILE_KEY)
    ui.console.print()
    ui.console.print(
        "[bold]История объёмов по типам (лист fact_type_daily)[/bold]\n"
        "[grey70]Накапливается по одной дате за запуск. Если история уже ведётся в "
        "отчёте старого формата — листы «Динамика AFS», «Динамика HTM», «Динамика "
        "TSS» с колонками «Дата» и «Текущий объём», — её можно подтянуть оттуда. "
        "Импорт только дополняет: даты, накопленные своими запусками, не "
        "перезаписываются.[/grey70]"
    )
    if current and str(current).strip():
        ui.console.print(f"Сейчас задан файл: [bold]{current}[/bold]")
        ui.console.print("[grey70]Enter — оставить, «-» — больше не подтягивать, "
                         "либо укажите другой путь.[/grey70]")
    # Искать по имени бесполезно: у старого отчёта оно произвольное. Зато листы
    # «Динамика <ТИП>» видны в оглавлении книги мгновенно — поэтому кандидаты
    # ищутся по СОДЕРЖИМОМУ в загрузках и папках отчёта, и человеку остаётся
    # выбрать номер, а не вспоминать и набирать путь целиком.
    candidates = _history_candidates()
    if candidates:
        ui.console.print("[bold]Похожие файлы (листы «Динамика <ТИП>»):[/bold]")
        for number, (found, sheets) in enumerate(candidates, start=1):
            ui.console.print(Text.assemble(
                ("  %d) " % number, "bold"), (found.name, ""),
                ("  [%s]" % ", ".join(sheets), "grey50"),
                ("\n     %s" % found.parent, "grey50"),
            ))
        raw = ui.ask("Номер файла из списка или путь к своему")
    else:
        ui.console.print("[grey50]Автоматически ничего похожего не нашлось — "
                         "искали в папке загрузок, папке исходных файлов отчёта и "
                         "папке выгрузки. Укажите путь вручную.[/grey50]")
        raw = ui.ask("Путь к файлу с историей")

    if not raw.strip():
        return False
    if raw.strip() == "-":
        settings.reset(HISTORY_FILE_KEY)
        ui.success("История из старого отчёта больше не подтягивается.")
        return True

    token = raw.strip()
    if token.isdigit() and 1 <= int(token) <= len(candidates):
        path = candidates[int(token) - 1][0]
    else:
        path = Path(token.strip('"')).expanduser()
    if not path.exists():
        ui.error(f"Файл не найден: {path}")
        return False

    # Разбираем сразу: пусть человек увидит, что именно прочиталось, а не
    # узнает о несовпадении формата через сутки при очередном запуске.
    try:
        from reports.portfolio_dynamics import history
        frame = history.parse_history_file(path)
    except Exception as exc:
        ui.error(str(exc))
        return False

    types = ", ".join(sorted(frame["portfolio_type"].unique()))
    ui.success(
        f"Прочитано строк: {len(frame)}; типы: {types}; период "
        f"{frame['business_date'].min().isoformat()} .. "
        f"{frame['business_date'].max().isoformat()}"
    )
    try:
        settings.set_value(HISTORY_FILE_KEY, str(path))
    except settings.SettingsError as exc:
        ui.error(str(exc))
        return False
    ui.console.print("[grey70]История подтянется при следующем запуске отчёта "
                     "«Динамика портфелей».[/grey70]")
    return True


# ── Создание папок по датам ──────────────────────────────────────────────────
def _source_configs(report: settings.ReportSource) -> List[file_discovery.SourceConfig]:
    """Источники отчёта как SourceConfig — их собирает config по ключам настроек.

    Сопоставление «ключ настройки -> константа config» держится здесь: settings
    про config ничего не знает (он его импортирует), а таблица отчётов нужна
    только этому экрану.
    """
    by_key = {
        "ovp_dir": [config.OVP_SOURCE],
        "balance_struct_dir": [config.BALANCE_STRUCT_SOURCE],
        "chpd_dir": [config.CHPD_SOURCE],
        "nim_dir": [config.NIM_SOURCE],
        "transfert_short_dir": [config.TRANSFERT_SHORT_SOURCE],
        "transfert_long_dir": [config.TRANSFERT_LONG_SOURCE],
        "portfolio_dynamics_dir": [config.PORTFOLIO_DYNAMICS_T0_SOURCE],
    }
    sources: List[file_discovery.SourceConfig] = []
    for key in report.dir_keys:
        sources.extend(by_key.get(key, []))
    return sources


def _reports_table() -> Table:
    table = Table(
        title="Создать папки по датам", title_style="bold cyan", box=box.SIMPLE_HEAVY,
        show_header=True, header_style="bold cyan",
        caption="Папки-даты включаются настройкой отчёта; у остальных отчётов "
                "раскладка плоская.",
        caption_style="grey50",
    )
    table.add_column("#", justify="right", style="bold yellow", no_wrap=True, width=3)
    table.add_column("Отчёт", style="bold white", max_width=24)
    table.add_column("Папка исходных файлов", style="grey70", overflow="fold", ratio=2)
    table.add_column("Папки-даты", no_wrap=True)

    for i, report in enumerate(settings.REPORT_SOURCES, start=1):
        dirs = "\n".join(str(settings.get(key)) for key in report.dir_keys)
        enabled = any(s.uses_date_folders for s in _source_configs(report))
        table.add_row(
            str(i), report.title, dirs,
            "[bold green]вкл[/bold green]" if enabled else "[grey50]выкл[/grey50]",
        )
    return table


def _make_date_folders(report: settings.ReportSource, days: int, include_weekends: bool) -> None:
    """Создаёт подпапки-даты у всех источников отчёта и печатает итог."""
    sources = _source_configs(report)
    if not sources:
        raise settings.SettingsError(f"У отчёта «{report.title}» нет папок с исходными файлами.")

    for source in sources:
        if not source.uses_date_folders:
            ui.warning(
                f"[{source.label}] Папки по датам выключены — включите их в настройках "
                f"отчёта «{report.title}», иначе отчёт не будет искать файлы в подпапках."
            )
            continue
        created, existed = file_discovery.create_date_folders(source, days, include_weekends)
        ui.success(
            f"[{source.label}] {source.directory}: создано папок {len(created)}, "
            f"уже было {len(existed)}"
        )
        if created:
            ui.console.print(
                "[grey70]   " + ", ".join(f.name for f in created[:12])
                + (" …" if len(created) > 12 else "") + "[/grey70]"
            )


def make_folders_cli(slug: str, days: int, include_weekends: bool) -> None:
    """Неинтерактивный режим: python console.py settings --make-folders <отчёт>."""
    report = settings.REPORT_SOURCES_BY_SLUG.get(slug)
    if report is None:
        raise settings.SettingsError(
            f"Неизвестный отчёт {slug!r}. Доступны: "
            + ", ".join(r.slug for r in settings.REPORT_SOURCES)
        )
    _make_date_folders(report, days, include_weekends)


def _make_folders_screen() -> None:
    ui.console.print()
    ui.console.print(_reports_table())
    choice = ui.ask("Отчёт (номер), 0 — назад", default="0")
    if choice in ("0", ""):
        return
    try:
        report = settings.REPORT_SOURCES[int(choice) - 1]
    except (ValueError, IndexError):
        ui.warning("Некорректный выбор.")
        return

    days_raw = ui.ask("На сколько дней вперёд создать папки", default="14")
    try:
        days = int(days_raw)
    except ValueError:
        ui.error(f"Ожидается целое число дней, получено {days_raw!r}.")
        return

    weekends = ui.ask("Создавать папки и на выходные? (y/N)", default="N")
    include_weekends = weekends.strip().lower().startswith("y")

    try:
        _make_date_folders(report, days, include_weekends)
    except (settings.SettingsError, file_discovery.SourceFileError) as exc:
        ui.error(str(exc))


def _reset_all() -> bool:
    keys = settings.overridden_keys()
    if not keys:
        ui.console.print("[grey70]Изменённых настроек нет — сбрасывать нечего.[/grey70]")
        return False
    ui.warning(f"Будут сброшены все изменённые настройки ({len(keys)}): {', '.join(keys)}")
    answer = ui.ask("Точно сбросить всё к значениям по умолчанию? (y/N)", default="N")
    if not answer.strip().lower().startswith("y"):
        ui.cancelled("Сброс отменён.")
        return False
    settings.reset_all()
    ui.success("Все настройки возвращены к значениям по умолчанию.")
    return True


def run_interactive() -> None:
    """Главный экран настроек. Вызывается из меню console.py."""
    changed = False
    try:
        while True:
            ui.console.print()
            ui.console.print(_groups_table())
            ui.console.print(
                "[grey70]п — проверить пути, д — создать папки по датам, "
                "р — дополнительные портфели, и — история из старого отчёта, "
                "с — сбросить всё к значениям по умолчанию, 0 — выйти[/grey70]"
            )
            choice = ui.ask("Раздел (номер) или действие", default="0").strip().lower()

            if choice in ("0", ""):
                return
            if choice in ("п", "p"):
                ui.console.print()
                ui.console.print(_paths_table())
                continue
            if choice in ("и", "i"):
                try:
                    changed |= _history_screen()
                except KeyboardInterrupt:
                    ui.console.print()
                    ui.cancelled("Выбор файла истории отменён.")
                continue
            if choice in ("р", "r"):
                try:
                    changed |= _manual_portfolios_screen()
                except KeyboardInterrupt:
                    ui.console.print()
                    ui.cancelled("Правка дополнительных портфелей отменена.")
                continue
            if choice in ("д", "d"):
                try:
                    _make_folders_screen()
                except KeyboardInterrupt:
                    ui.console.print()
                    ui.cancelled("Создание папок отменено.")
                continue
            if choice in ("с", "c", "s"):
                changed |= _reset_all()
                continue

            try:
                group = settings.GROUPS[int(choice) - 1]
            except (ValueError, IndexError):
                ui.warning("Некорректный выбор, попробуйте снова.")
                continue

            changed |= _edit_group(group)
    finally:
        if changed:
            # Отчёты читают config.X в момент запуска, поэтому без пересборки
            # в уже запущенной консоли остались бы старые пути до перезапуска.
            config.reload()
            ui.console.print(
                f"[grey70]Настройки сохранены в {settings.settings_path()} "
                "и применены к текущему сеансу.[/grey70]"
            )


# ── Неинтерактивный режим: python console.py settings ... ────────────────────
def print_all(group_key: Optional[str] = None) -> None:
    """Печатает настройки для --list: по таблице на раздел.

    Одной плоской таблицей это не печатается: колонка раздела съедает ширину,
    и на терминале в 80 знаков значение сжимается до столбика по букве.
    """
    known = {g.key for g in settings.GROUPS}
    if group_key and group_key not in known:
        raise settings.SettingsError(
            f"Неизвестный раздел {group_key!r}. Доступны: {', '.join(sorted(known))}"
        )

    ui.console.print(f"[grey50]Файл настроек: {settings.settings_path()}[/grey50]")
    for group in settings.GROUPS:
        if group_key and group.key != group_key:
            continue
        table = Table(
            title=group.title, title_style="bold cyan", box=box.SIMPLE_HEAVY,
            show_header=True, header_style="bold cyan",
        )
        table.add_column("Ключ", style="bold white", max_width=34)
        table.add_column("Значение", style="grey70", overflow="fold", ratio=2)
        table.add_column("Источник", no_wrap=True)
        for setting in settings.SETTINGS_BY_GROUP[group.key]:
            overridden = settings.is_overridden(setting.key)
            table.add_row(
                setting.key, _value_repr(setting.key),
                "[bold yellow]изменено[/bold yellow]" if overridden else "[grey50]по умолчанию[/grey50]",
            )
        ui.console.print(table)


def print_paths() -> int:
    """Печатает проверку путей. Возвращает число недоступных — для кода возврата."""
    ui.console.print(_paths_table())
    broken = [s for s in settings.check_paths() if s.is_problem]
    if broken:
        ui.warning(f"Недоступно путей: {len(broken)} ({', '.join(s.key for s in broken)})")
    else:
        ui.success("Все обязательные пути на месте (выходные папки создадутся при первом запуске).")
    return len(broken)


def apply_assignments(assignments: List[str]) -> None:
    """Применяет пары ключ=значение из --set."""
    for item in assignments:
        if "=" not in item:
            raise settings.SettingsError(
                f"Ожидается ключ=значение, получено {item!r}. "
                'Пример: --set ovp_dir="/Volumes/Jupiter/data/OVP"'
            )
        key, _, raw = item.partition("=")
        value = settings.set_value(key.strip(), raw)
        ui.success(f"{key.strip()} = {value}")
    config.reload()
