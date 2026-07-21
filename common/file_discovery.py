"""Поиск входных файлов отчётов по дате, зашитой в имени файла (см. config.py).

Общий механизм для всех "файловых" отчётов: указываем папку и регулярное
выражение с одной группой-датой, дальше можно получить самый свежий файл,
файл на конкретную дату или устроить пользователю интерактивный выбор.
"""
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import List, Tuple

from rich import box
from rich.table import Table

from common import ui


@dataclass(frozen=True)
class SourceConfig:
    """Описывает, где и как искать входные файлы одного источника отчёта."""

    directory: Path
    filename_regex: str  # регулярное выражение с ОДНОЙ группой — датой в имени файла
    date_format: str  # strptime-формат для содержимого этой группы
    label: str = "файл"  # для сообщений/промптов, если у отчёта несколько источников


class SourceFileError(RuntimeError):
    """Ошибка поиска входного файла отчёта."""


def find_dated_files(source: SourceConfig) -> List[Tuple[date, Path]]:
    """Сканирует source.directory, извлекает дату из имени каждого файла по
    filename_regex и возвращает список (дата, путь), отсортированный по дате
    по убыванию (сначала самые свежие).
    """
    if not source.directory.exists():
        raise SourceFileError(f"[{source.label}] Папка с исходными файлами не найдена: {source.directory}")

    pattern = re.compile(source.filename_regex)
    results: List[Tuple[date, Path]] = []
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
        raise SourceFileError(
            f"[{source.label}] В папке {source.directory} не найдено файлов, "
            f"подходящих под шаблон {source.filename_regex!r}."
        )
    return found[0]


def resolve_file_for_date(source: SourceConfig, target_date: date) -> Path:
    """Неинтерактивный поиск файла на конкретную дату (для CLI-режима --date)."""
    for found_date, path in find_dated_files(source):
        if found_date == target_date:
            return path
    raise SourceFileError(
        f"[{source.label}] Файл на дату {target_date.isoformat()} не найден в {source.directory}"
    )


def _file_table(source: SourceConfig, top: List[Tuple[date, Path]], total: int) -> Table:
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
    table.add_column("Дата", style="bold white")
    table.add_column("Файл", style="grey70")
    for i, (d, p) in enumerate(top, start=1):
        table.add_row(str(i), d.isoformat(), p.name)
    return table


def prompt_for_file(source: SourceConfig, n_recent: int = 5) -> Path:
    """Интерактивный выбор файла: показывает N последних дат (по умолчанию —
    самая свежая), позволяет выбрать по номеру из списка или вписать
    произвольную дату в формате YYYY-MM-DD.
    """
    found = find_dated_files(source)
    if not found:
        raise SourceFileError(
            f"[{source.label}] В папке {source.directory} не найдено файлов, "
            f"подходящих под шаблон {source.filename_regex!r}."
        )

    latest_date, latest_path = found[0]
    top = found[:n_recent]

    ui.console.print(_file_table(source, top, total=len(found)))

    choice = ui.ask(
        f"[{source.label}] Номер из списка, дата YYYY-MM-DD, "
        f"или Enter для последней ({latest_date.isoformat()})"
    )

    if not choice:
        return latest_path

    if choice.isdigit() and 1 <= int(choice) <= len(top):
        return top[int(choice) - 1][1]

    try:
        target_date = datetime.strptime(choice, "%Y-%m-%d").date()
    except ValueError as exc:
        raise SourceFileError(
            f"[{source.label}] Не удалось разобрать ввод '{choice}': "
            f"ни номер варианта из списка, ни дата в формате YYYY-MM-DD"
        ) from exc

    for d, p in found:
        if d == target_date:
            return p
    raise SourceFileError(f"[{source.label}] Файл с датой {target_date.isoformat()} не найден в {source.directory}")
