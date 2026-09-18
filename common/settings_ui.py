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

import config
from common import settings, ui

MENU_TITLE = "Настройки"
MENU_DESCRIPTION = "Пути к папкам, шаблоны имён файлов и константы отчётов — без правки config.py"


def _value_repr(key: str) -> str:
    value = settings.get(key)
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
                "[grey70]п — проверить пути, с — сбросить всё к значениям по умолчанию, "
                "0 — выйти в главное меню[/grey70]"
            )
            choice = ui.ask("Раздел (номер) или действие", default="0").strip().lower()

            if choice in ("0", ""):
                return
            if choice in ("п", "p"):
                ui.console.print()
                ui.console.print(_paths_table())
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
