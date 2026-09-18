"""Поиск входных файлов отчётов по дате в имени файла (см. config.py).

Общий механизм для всех "файловых" отчётов: указываем папку и регулярное
выражение с одной группой-датой, дальше можно получить самый свежий файл,
файл на конкретную дату или устроить пользователю интерактивный выбор.

Для отчётов, у файлов которых нет даты в имени (например, ОВП), можно
оставить filename_regex/date_format пустыми — тогда файлы сортируются по
дате изменения (mtime) вместо даты, разобранной из имени.

ПАПКИ-ДАТЫ. Если у источника задан date_folder_format, то внутри его папки
ищутся ещё и подпапки, имя которых разбирается как дата (по умолчанию
2026-09-18). Файлы внутри такой подпапки получают её дату независимо от
собственных имён. Смысл: складывать выгрузки за день в одну папку, а не
копить в общей папке «зоопарк» файлов за все даты сразу. Плоская раскладка
при этом продолжает работать — обе сканируются одновременно, поэтому переход
на папки можно делать постепенно.

Если папка из конфига недоступна (не примонтирован сетевой диск, опечатка в
пути и т.п.) или в ней не нашлось ни одного подходящего файла — интерактивные
prompt_for_file/prompt_for_multiple_files не падают с ошибкой, а предлагают
вписать путь(и) к файлу вручную (см. _prompt_manual_path/_paths).
"""
import re
from dataclasses import dataclass
from datetime import date, datetime, timedelta
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
    date_folder_format задан — дополнительно сканируются подпапки-даты
    (см. модульную docstring); файлы внутри них датируются именем папки.
    """

    directory: Path
    filename_regex: Optional[str] = None  # регулярное выражение с ОДНОЙ группой — датой в имени файла
    date_format: Optional[str] = None  # strptime-формат для содержимого этой группы
    glob_pattern: str = "*.xlsx"  # используется, когда filename_regex не задан
    label: str = "файл"  # для сообщений/промптов, если у отчёта несколько источников
    date_folder_format: Optional[str] = None  # strptime-формат ИМЕНИ подпапки-даты

    @property
    def uses_mtime(self) -> bool:
        return not self.filename_regex

    @property
    def uses_date_folders(self) -> bool:
        return bool(self.date_folder_format)


@dataclass(frozen=True)
class DateFolder:
    """Подпапка-дата и лежащие в ней файлы источника."""

    date: date
    path: Path
    files: List[Path]

    @property
    def has_data(self) -> bool:
        return bool(self.files)


# Временные файлы Excel («~$Отчёт.xlsx») появляются рядом с открытым файлом и
# данными не являются — иначе папка «с данными» находилась бы по мусору.
def _is_real_file(path: Path) -> bool:
    return path.is_file() and not path.name.startswith("~$")


class SourceFileError(RuntimeError):
    """Ошибка поиска входного файла отчёта."""


def find_date_folders(source: SourceConfig) -> List[DateFolder]:
    """Подпапки-даты источника с их файлами, свежие сначала.

    Пустые папки тоже возвращаются (их заранее создаёт «Создать папки по датам»),
    поэтому вызывающий код сам решает, что считать папкой «с данными».
    """
    if not source.uses_date_folders:
        return []
    if not source.directory.exists():
        raise SourceFileError(f"[{source.label}] Папка с исходными файлами не найдена: {source.directory}")

    folders: List[DateFolder] = []
    for entry in source.directory.iterdir():
        if not entry.is_dir():
            continue
        try:
            folder_date = datetime.strptime(entry.name, source.date_folder_format).date()
        except ValueError:
            continue  # обычная папка, не дата — не наша
        files = sorted(f for f in entry.glob(source.glob_pattern) if _is_real_file(f))
        folders.append(DateFolder(date=folder_date, path=entry, files=files))

    folders.sort(key=lambda f: f.date, reverse=True)
    return folders


def latest_folder_with_data(source: SourceConfig, min_files: int = 1) -> Optional[DateFolder]:
    """Самая свежая подпапка-дата, в которой лежит не меньше min_files файлов.

    Пустые папки пропускаются: они заранее созданы «на будущее», и брать их
    как последнюю дату значило бы каждый раз спотыкаться о завтрашнюю папку.
    """
    for folder in find_date_folders(source):
        if len(folder.files) >= min_files:
            return folder
    return None


def find_dated_files(source: SourceConfig) -> List[Tuple[date, Path]]:
    """Сканирует source.directory и возвращает список (дата, путь),
    отсортированный по дате по убыванию (сначала самые свежие).

    Дата берётся из имени файла (filename_regex/date_format) либо, если они
    не заданы, из времени последнего изменения файла (mtime). Файлы внутри
    подпапок-дат датируются именем подпапки — их собственные имена при этом
    не важны.
    """
    if not source.directory.exists():
        raise SourceFileError(f"[{source.label}] Папка с исходными файлами не найдена: {source.directory}")

    results: List[Tuple[date, Path]] = []

    if source.uses_mtime:
        for f in source.directory.glob(source.glob_pattern):
            if not _is_real_file(f):
                continue
            results.append((datetime.fromtimestamp(f.stat().st_mtime).date(), f))
    else:
        pattern = re.compile(source.filename_regex)
        for f in source.directory.iterdir():
            if not _is_real_file(f):
                continue
            m = pattern.search(f.name)
            if not m:
                continue
            try:
                parsed = datetime.strptime(m.group(1), source.date_format).date()
            except ValueError:
                continue
            results.append((parsed, f))

    for folder in find_date_folders(source):
        results.extend((folder.date, f) for f in folder.files)

    results.sort(key=lambda item: item[0], reverse=True)
    return results


def create_date_folders(
    source: SourceConfig, days: int, include_weekends: bool = False,
    start: Optional[date] = None,
) -> Tuple[List[Path], List[Path]]:
    """Заранее создаёт подпапки-даты на days дней вперёд, начиная с сегодня.

    Возвращает (созданные, уже существовавшие). По умолчанию выходные
    пропускаются: отчёты казначейства строятся по рабочим дням, и папки на
    субботу с воскресеньем только мешают выбирать последнюю дату.
    """
    if not source.uses_date_folders:
        raise SourceFileError(
            f"[{source.label}] У источника не задан формат имени подпапки-даты — "
            "включите папки-даты в настройках этого отчёта."
        )
    if days < 1:
        raise SourceFileError(f"[{source.label}] Количество дней должно быть больше нуля, получено {days}.")

    try:
        source.directory.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise SourceFileError(
            f"[{source.label}] Не удалось создать папку {source.directory}: {exc}. "
            "Проверьте, примонтирован ли диск и есть ли права на запись."
        ) from exc

    created: List[Path] = []
    existed: List[Path] = []
    current = start or date.today()
    for _ in range(days):
        if include_weekends or current.weekday() < 5:
            folder = source.directory / current.strftime(source.date_folder_format)
            if folder.exists():
                existed.append(folder)
            else:
                folder.mkdir(parents=True)
                created.append(folder)
        current += timedelta(days=1)
    return created, existed


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


def _folder_table(source: SourceConfig, folders: List[DateFolder]) -> Table:
    table = Table(
        title=f"{source.label} — папки по датам в {source.directory}",
        title_style="bold cyan", box=box.SIMPLE_HEAVY, show_header=True,
        header_style="bold cyan",
        caption="Пустые папки созданы заранее — положите в них выгрузки.",
        caption_style="grey50",
    )
    table.add_column("#", justify="right", style="bold yellow")
    table.add_column("Дата", style="bold white")
    table.add_column("Файлов", justify="right")
    table.add_column("Файлы", style="grey70", overflow="fold")
    for i, folder in enumerate(folders, start=1):
        count = str(len(folder.files)) if folder.files else "[grey50]—[/grey50]"
        names = ", ".join(f.name for f in folder.files) if folder.files else "[grey50]пусто[/grey50]"
        table.add_row(str(i), folder.date.isoformat(), count, names)
    return table


def prompt_for_date_folder(
    source: SourceConfig, min_files: int = 1, n_recent: int = 7
) -> Optional[DateFolder]:
    """Интерактивный выбор подпапки-даты. None — подходящих папок нет.

    None означает «работаем по-старому»: вызывающий отчёт откатывается на выбор
    отдельных файлов, чтобы плоская раскладка продолжала работать.
    """
    if not source.uses_date_folders:
        return None
    try:
        folders = find_date_folders(source)
    except SourceFileError:
        return None
    if not folders:
        return None

    with_data = [f for f in folders if len(f.files) >= min_files]
    if not with_data:
        ui.warning(
            f"[{source.label}] Папки по датам есть, но ни в одной нет нужных файлов "
            f"(нужно минимум {min_files}). Положите выгрузки в папку нужной даты."
        )
        return None

    ui.console.print(_folder_table(source, folders[:n_recent]))
    latest = with_data[0]
    choice = ui.ask(
        f"[{source.label}] Номер папки или дата YYYY-MM-DD "
        f"(Enter — последняя с данными, {latest.date.isoformat()})"
    )
    if not choice:
        return latest

    token = choice.strip()
    shown = folders[:n_recent]
    if token.isdigit() and 1 <= int(token) <= len(shown):
        return shown[int(token) - 1]
    try:
        wanted = datetime.strptime(token, "%Y-%m-%d").date()
    except ValueError as exc:
        raise SourceFileError(
            f"[{source.label}] Не удалось разобрать '{token}': ни номер из списка, ни дата YYYY-MM-DD."
        ) from exc
    for folder in folders:
        if folder.date == wanted:
            return folder
    raise SourceFileError(
        f"[{source.label}] Папка на дату {wanted.isoformat()} не найдена в {source.directory}"
    )


def resolve_date_folder(source: SourceConfig, token: str) -> DateFolder:
    """Разбирает значение --folder: путь к папке либо дата YYYY-MM-DD."""
    candidate = Path(token).expanduser()
    if not candidate.is_dir() and source.uses_date_folders:
        try:
            wanted = datetime.strptime(token.strip(), "%Y-%m-%d").date()
        except ValueError:
            wanted = None
        if wanted is not None:
            for folder in find_date_folders(source):
                if folder.date == wanted:
                    return folder
            raise SourceFileError(
                f"[{source.label}] Папка на дату {wanted.isoformat()} не найдена в {source.directory}"
            )
        candidate = source.directory / token

    if not candidate.is_dir():
        raise SourceFileError(f"[{source.label}] Папка не найдена: {candidate}")

    files = sorted(f for f in candidate.glob(source.glob_pattern) if _is_real_file(f))
    try:
        folder_date = datetime.strptime(candidate.name, source.date_folder_format or "%Y-%m-%d").date()
    except ValueError:
        folder_date = datetime.fromtimestamp(candidate.stat().st_mtime).date()
    return DateFolder(date=folder_date, path=candidate, files=files)


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
