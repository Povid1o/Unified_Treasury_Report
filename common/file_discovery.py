"""Поиск входных файлов отчётов по дате в имени файла (см. config.py).

Общий механизм для всех "файловых" отчётов: указываем папку и регулярное
выражение с одной группой-датой, дальше можно получить самый свежий файл,
файл на конкретную дату или устроить пользователю интерактивный выбор.

Для отчётов, у файлов которых нет даты в имени (например, ОВП), можно
оставить filename_regex/date_format пустыми — тогда файлы сортируются по
дате изменения (mtime) вместо даты, разобранной из имени.

Если папка из конфига недоступна (не примонтирован сетевой диск, опечатка в
пути и т.п.) или в ней не нашлось ни одного подходящего файла — интерактивные
prompt_for_file/prompt_for_multiple_files не падают с ошибкой, а предлагают
вписать путь(и) к файлу вручную (см. _prompt_manual_path/_paths).
"""
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import List, Optional, Tuple

from rich import box
from rich.table import Table

from common import ui


@dataclass(frozen=True)
class SourceConfig:
    """Описывает, где и как искать входные файлы одного источника отчёта.

    filename_regex/date_format заданы — дата берётся из имени файла.
    filename_regex не задан (None) — дата берётся из времени изменения файла
    (mtime), а glob_pattern используется для отбора подходящих файлов.
    """

    directory: Path
    filename_regex: Optional[str] = None  # регулярное выражение с ОДНОЙ группой — датой в имени файла
    date_format: Optional[str] = None  # strptime-формат для содержимого этой группы
    glob_pattern: str = "*.xlsx"  # используется, когда filename_regex не задан
    label: str = "файл"  # для сообщений/промптов, если у отчёта несколько источников

    @property
    def uses_mtime(self) -> bool:
        return not self.filename_regex


class SourceFileError(RuntimeError):
    """Ошибка поиска входного файла отчёта."""


def find_dated_files(source: SourceConfig) -> List[Tuple[date, Path]]:
    """Сканирует source.directory и возвращает список (дата, путь),
    отсортированный по дате по убыванию (сначала самые свежие).

    Дата берётся из имени файла (filename_regex/date_format) либо, если они
    не заданы, из времени последнего изменения файла (mtime).
    """
    if not source.directory.exists():
        raise SourceFileError(f"[{source.label}] Папка с исходными файлами не найдена: {source.directory}")

    results: List[Tuple[date, Path]] = []

    if source.uses_mtime:
        for f in source.directory.glob(source.glob_pattern):
            if not f.is_file():
                continue
            results.append((datetime.fromtimestamp(f.stat().st_mtime).date(), f))
    else:
        pattern = re.compile(source.filename_regex)
        for f in source.directory.iterdir():
            if not f.is_file():
                continue
            m = pattern.search(f.name)
            if not m:
                continue
            try:
                parsed = datetime.strptime(m.group(1), source.date_format).date()
            except ValueError:
                continue
            results.append((parsed, f))

    results.sort(key=lambda item: item[0], reverse=True)
    return results


def latest_file(source: SourceConfig) -> Tuple[date, Path]:
    """Возвращает (дата, путь) самого свежего файла источника."""
    found = find_dated_files(source)
    if not found:
        raise SourceFileError(_not_found_message(source))
    return found[0]


def resolve_file_for_date(source: SourceConfig, target_date: date) -> Path:
    """Неинтерактивный поиск файла на конкретную дату (для CLI-режима --date)."""
    for found_date, path in find_dated_files(source):
        if found_date == target_date:
            return path
    raise SourceFileError(
        f"[{source.label}] Файл на дату {target_date.isoformat()} не найден в {source.directory}"
    )


def _not_found_message(source: SourceConfig) -> str:
    if source.uses_mtime:
        return f"[{source.label}] В папке {source.directory} не найдено файлов по шаблону {source.glob_pattern!r}."
    return f"[{source.label}] В папке {source.directory} не найдено файлов, подходящих под шаблон {source.filename_regex!r}."


def _file_table(source: SourceConfig, top: List[Tuple[date, Path]], total: int) -> Table:
    date_column = "Дата изменения" if source.uses_mtime else "Дата"
    table = Table(
        title=f"{source.label} — {source.directory}",
        title_style="bold cyan",
        box=box.SIMPLE_HEAVY,
        show_header=True,
        header_style="bold cyan",
        caption=f"Всего найдено файлов: {total}",
        caption_style="grey50",
    )
    table.add_column("#", justify="right", style="bold yellow")
    table.add_column(date_column, style="bold white")
    table.add_column("Файл", style="grey70")
    for i, (d, p) in enumerate(top, start=1):
        table.add_row(str(i), d.isoformat(), p.name)
    return table


def _resolve_token(
    token: str, source: SourceConfig, top: List[Tuple[date, Path]], found: List[Tuple[date, Path]]
) -> Path:
    """Разбирает один пользовательский токен в путь к файлу: номер из
    показанного списка (top), дата YYYY-MM-DD (ищется по ВСЕМ найденным
    файлам, не только top — так доступны файлы старше показанных N), либо
    точное/частичное совпадение по имени файла (для файлов, до которых не
    достаёт регулярка даты или которые за пределами выборки)."""
    token = token.strip()

    if token.isdigit() and 1 <= int(token) <= len(top):
        return top[int(token) - 1][1]

    try:
        target_date = datetime.strptime(token, "%Y-%m-%d").date()
    except ValueError:
        target_date = None
    if target_date is not None:
        for d, p in found:
            if d == target_date:
                return p
        raise SourceFileError(f"[{source.label}] Файл с датой {target_date.isoformat()} не найден в {source.directory}")

    # Похоже на имя файла — ищем точное совпадение, затем частичное, среди
    # ВСЕХ файлов папки (не только тех, что подошли под date-регулярку).
    candidates = [f for f in source.directory.iterdir() if f.is_file()]
    exact = [f for f in candidates if f.name.lower() == token.lower()]
    if exact:
        return exact[0]
    partial = [f for f in candidates if token.lower() in f.name.lower()]
    if len(partial) == 1:
        return partial[0]
    if len(partial) > 1:
        names = ", ".join(f.name for f in partial[:10])
        raise SourceFileError(f"[{source.label}] Имени '{token}' соответствует несколько файлов: {names}")

    raise SourceFileError(
        f"[{source.label}] Не удалось разобрать '{token}': ни номер из списка, "
        f"ни дата YYYY-MM-DD, ни имя файла в {source.directory}"
    )


def _prompt_manual_path(source: SourceConfig, reason: str) -> Path:
    ui.warning(f"[{source.label}] {reason}")
    raw = ui.ask(f"[{source.label}] Введите путь к файлу вручную")
    if not raw:
        raise SourceFileError(f"[{source.label}] Путь к файлу не указан.")
    return Path(raw.strip().strip('"'))


def _prompt_manual_paths(source: SourceConfig, reason: str) -> List[Path]:
    ui.warning(f"[{source.label}] {reason}")
    raw = ui.ask(f"[{source.label}] Введите путь(и) к файлу через запятую вручную")
    paths = [Path(p.strip().strip('"')) for p in raw.split(",") if p.strip()]
    if not paths:
        raise SourceFileError(f"[{source.label}] Путь к файлу не указан.")
    return paths


def prompt_for_file(source: SourceConfig, n_recent: int = 5) -> Path:
    """Интерактивный выбор ОДНОГО файла: показывает N последних дат (по
    умолчанию — самая свежая), позволяет выбрать по номеру, дате или имени
    файла (см. _resolve_token). Если папка недоступна или в ней ничего не
    нашлось — предлагает вписать путь вручную (см. _prompt_manual_path).
    """
    try:
        found = find_dated_files(source)
    except SourceFileError as exc:
        return _prompt_manual_path(source, str(exc))
    if not found:
        return _prompt_manual_path(source, _not_found_message(source))

    latest_date, latest_path = found[0]
    top = found[:n_recent]

    ui.console.print(_file_table(source, top, total=len(found)))

    choice = ui.ask(
        f"[{source.label}] Номер из списка, дата YYYY-MM-DD, имя файла, "
        f"или Enter для последней ({latest_date.isoformat()})"
    )
    if not choice:
        return latest_path
    return _resolve_token(choice, source, top, found)


def prompt_for_multiple_files(source: SourceConfig, n_recent: int = 10) -> List[Path]:
    """Интерактивный выбор НЕСКОЛЬКИХ файлов за один раз: показывает N
    последних дат (не больше — чтобы не засорять консоль), пользователь
    вводит через запятую любую смесь номеров/дат/имён файлов. Пустой ввод —
    только самый свежий файл. Если папка недоступна или в ней ничего не
    нашлось — предлагает вписать путь(и) вручную (см. _prompt_manual_paths).
    """
    try:
        found = find_dated_files(source)
    except SourceFileError as exc:
        return _prompt_manual_paths(source, str(exc))
    if not found:
        return _prompt_manual_paths(source, _not_found_message(source))

    latest_date, latest_path = found[0]
    top = found[:n_recent]

    ui.console.print(_file_table(source, top, total=len(found)))

    raw = ui.ask(
        f"[{source.label}] Номера/даты/имена файлов через запятую (можно смешивать), "
        f"или Enter для последней ({latest_date.isoformat()})"
    )
    if not raw:
        return [latest_path]

    selected: List[Path] = []
    seen = set()
    for token in raw.split(","):
        if not token.strip():
            continue
        path = _resolve_token(token, source, top, found)
        if path not in seen:
            seen.add(path)
            selected.append(path)
    return selected
