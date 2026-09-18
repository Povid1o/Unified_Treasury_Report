"""Приёмка выгрузок из папки загрузок в папку-дату отчёта.

Зачем. Выгрузку скачивают из системы в «Загрузки», а отчёт ждёт её в
data/<Отчёт>/<дата>/ — и между этими двумя точками файл до сих пор носили
руками. Здесь этот шаг автоматизирован: перед запуском ETL отчёт смотрит,
есть ли в папке нужной даты оба среза, и если чего-то не хватает — ищет
недостающее в загрузках и раскладывает по местам.

Что считается «нашим» файлом в загрузках: имя, подходящее под тот же шаблон,
по которому отчёт ищет файлы в своей папке (portfolio_dynamics_regex —
«Позиция за период [дд.мм.гггг] - [дд.мм.гггг] - SECURITIES.xlsx»). По
содержимому книги ничего не открывается: перебирать и читать все .xlsx в
чужой папке долго и бесцеремонно, а имя выгрузки и так однозначно.

Пара срезов на дату D собирается так: T0 — файл с датой D, T-7 — самый
свежий файл с датой строго раньше D. Ни переименовывать, ни раскладывать в
правильном порядке ничего не нужно.

Каждый срез кладётся ДВАЖДЫ: в папку отчётной даты (где он нужен как часть
пары) и в папку своей собственной даты. Иначе файл за 11.09, попавший в
папку 2026-09-18 как T-7, оказывался бы единственным экземпляром и отчёт на
11.09 потом было бы не собрать — пришлось бы выкачивать выгрузку заново.
Второй экземпляр это снимает: папка 2026-09-11 получает свой файл сразу.
"""
import datetime as dt
import re
import shutil
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional, Tuple

BASE_DIR = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(BASE_DIR))

from common.file_discovery import SourceConfig  # noqa: E402
from reports.portfolio_dynamics.etl import (  # noqa: E402
    PortfolioDynamicsError, logger, read_business_date,
)


@dataclass(frozen=True)
class Candidate:
    """Файл-кандидат (в загрузках или уже в папке отчёта) и дата его среза."""

    path: Path
    business_date: dt.date


@dataclass
class ImportPlan:
    """Что нужно сделать, чтобы в папке нужной даты оказались оба среза."""

    target_date: dt.date
    folder: Path
    existing: List[Path] = field(default_factory=list)          # уже на месте
    to_import: List[Tuple[Path, Path]] = field(default_factory=list)  # (откуда, куда)
    archive: List[Tuple[Path, Path]] = field(default_factory=list)  # копия в папку своей даты
    problem: Optional[str] = None  # чего не хватает, человекочитаемо

    @property
    def complete(self) -> bool:
        return self.problem is None

    @property
    def needs_import(self) -> bool:
        return bool(self.to_import)

    def describe(self) -> str:
        if not self.complete:
            return self.problem or ""
        if not self.to_import:
            return f"оба среза уже лежат в {self.folder.name}"
        moved = ", ".join(src.name for src, _dst in self.to_import)
        return f"из загрузок будет взято: {moved}"


@dataclass(frozen=True)
class AvailableDate:
    """Строка для выбора даты: что на эту дату есть и откуда оно возьмётся."""

    date: dt.date
    origin: str        # «папка», «загрузки», «папка + загрузки»
    files: List[str]
    complete: bool


def _is_real_file(path: Path) -> bool:
    return path.is_file() and not path.name.startswith("~$")


def folder_for_date(source: SourceConfig, target_date: dt.date) -> Path:
    """Папка, в которой должны лежать срезы на дату.

    С включёнными папками-датами это подпапка вида 2026-09-18, с выключенными —
    сама папка исходных файлов (плоская раскладка): приёмка из загрузок полезна
    и там, просто без раскладки по дням.
    """
    if not source.uses_date_folders:
        return Path(source.directory)
    return Path(source.directory) / target_date.strftime(source.date_folder_format)


def _dated(paths: List[Path], source: SourceConfig) -> List[Candidate]:
    """Файлы с их датами среза; файлы без распознаваемой даты отбрасываются."""
    pattern = re.compile(source.filename_regex) if source.filename_regex else None
    result: List[Candidate] = []
    for path in paths:
        parsed: Optional[dt.date] = None
        if pattern is not None:
            match = pattern.search(path.name)
            if match:
                try:
                    parsed = dt.datetime.strptime(match.group(1), source.date_format).date()
                except ValueError:
                    parsed = None
        if parsed is None:
            parsed = read_business_date(path)  # дороже: открывает книгу
        if parsed is not None:
            result.append(Candidate(path=path, business_date=parsed))
    result.sort(key=lambda c: c.business_date, reverse=True)
    return result


def scan_downloads(source: SourceConfig, downloads_dir: Path) -> List[Candidate]:
    """Выгрузки отчёта, лежащие в папке загрузок, свежие сначала.

    Недоступная папка загрузок — не ошибка: значит, приёмке просто неоткуда
    брать файлы, и отчёт работает как раньше, по своей папке.
    """
    downloads_dir = Path(downloads_dir)
    if not downloads_dir.is_dir():
        logger.info("Папка загрузок недоступна (%s) — приёмка из загрузок пропущена.", downloads_dir)
        return []

    pattern = re.compile(source.filename_regex) if source.filename_regex else None
    files = [
        f for f in downloads_dir.glob(source.glob_pattern)
        if _is_real_file(f) and (pattern is None or pattern.search(f.name))
    ]
    return _dated(files, source)


def _where_we_looked(folder: Path, downloads_dir: Optional[Path]) -> str:
    """Человекочитаемое «где искали» — чтобы опечатка в пути была видна сразу."""
    if downloads_dir is None:
        return f"в {folder} (приёмка из загрузок выключена)"
    return f"ни в {folder}, ни в загрузках ({downloads_dir})"


def _files_in_folder(source: SourceConfig, folder: Path) -> List[Path]:
    if not folder.is_dir():
        return []
    return sorted(f for f in folder.glob(source.glob_pattern) if _is_real_file(f))


def plan_import(source: SourceConfig, downloads_dir: Optional[Path],
                target_date: dt.date, archive_own_date: bool = True) -> ImportPlan:
    """Собирает план: что уже на месте, что взять из загрузок, чего не хватает.

    Ничего не перекладывает — только считает. Так план можно показать
    пользователю до того, как файлы поедут, и так же им пользуется экран
    выбора даты.
    """
    folder = folder_for_date(source, target_date)
    plan = ImportPlan(target_date=target_date, folder=folder)

    in_folder = _dated(_files_in_folder(source, folder), source)
    if source.uses_date_folders and len(_files_in_folder(source, folder)) >= 2:
        # Папка-дата уже укомплектована — в загрузки можно не заглядывать.
        plan.existing = _files_in_folder(source, folder)
        if archive_own_date:
            plan.archive = _plan_archive(source, folder, in_folder)
        return plan

    downloads = scan_downloads(source, downloads_dir) if downloads_dir else []
    taken_names = {c.path.name for c in in_folder}
    pool = in_folder + [c for c in downloads if c.path.name not in taken_names]
    where = _where_we_looked(folder, downloads_dir)

    t0 = next((c for c in pool if c.business_date == target_date), None)
    if t0 is None:
        plan.problem = f"Срез на {target_date.isoformat()} не найден {where}."
        return plan

    t7 = next((c for c in pool if c.business_date < target_date), None)
    if t7 is None:
        plan.problem = (
            f"Для среза на {target_date.isoformat()} не нашлось более раннего среза (T-7) "
            f"{where}. Отчёту нужны ДВА файла: на отчётную дату и на неделю раньше."
        )
        return plan

    for candidate in (t0, t7):
        if candidate.path.parent == folder:
            plan.existing.append(candidate.path)
        else:
            plan.to_import.append((candidate.path, folder / candidate.path.name))
    if archive_own_date:
        plan.archive = _plan_archive(source, folder, [t0, t7])
    return plan


def _plan_archive(source: SourceConfig, folder: Path,
                  candidates: List[Candidate]) -> List[Tuple[Path, Path]]:
    """Куда продублировать каждый срез — в папку его СОБСТВЕННОЙ даты.

    У T0 своя дата и есть отчётная, поэтому запись получается только для T-7.
    Источником копии берётся итоговый путь в папке отчётной даты: оригинал в
    загрузках к моменту копирования может быть уже перенесён.
    """
    if not source.uses_date_folders:
        return []
    entries: List[Tuple[Path, Path]] = []
    for candidate in candidates:
        own_folder = folder_for_date(source, candidate.business_date)
        if own_folder == folder:
            continue
        final = folder / candidate.path.name
        destination = own_folder / candidate.path.name
        if destination.exists():
            continue
        entries.append((final, destination))
    return entries


def apply_import(plan: ImportPlan, move: bool = True) -> List[Path]:
    """Перекладывает файлы по плану. Возвращает пути, оказавшиеся в папке.

    move=True переносит (файл исчезает из загрузок), False — копирует. Файл,
    который в папке уже есть под тем же именем, не трогается: повторный запуск
    за ту же дату ничего не ломает.
    """
    if not plan.complete:
        raise PortfolioDynamicsError(plan.problem)
    if not plan.to_import:
        _apply_archive(plan)
        return list(plan.existing)

    try:
        plan.folder.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise PortfolioDynamicsError(
            f"Не удалось создать папку {plan.folder}: {exc}. "
            "Проверьте, примонтирован ли диск и есть ли права на запись."
        ) from exc

    landed = list(plan.existing)
    verb = "Перенесено" if move else "Скопировано"
    for src, dst in plan.to_import:
        if dst.exists():
            logger.info("%s уже лежит в %s — пропущено.", dst.name, plan.folder.name)
            landed.append(dst)
            continue
        try:
            if move:
                shutil.move(str(src), str(dst))
            else:
                shutil.copy2(str(src), str(dst))
        except (OSError, shutil.Error) as exc:
            raise PortfolioDynamicsError(
                f"Не удалось перенести {src.name} из загрузок в {plan.folder}: {exc}. "
                "Проверьте права на запись и не открыт ли файл в Excel."
            ) from exc
        logger.info("%s из загрузок: %s -> %s", verb, src.name, plan.folder)
        landed.append(dst)

    _apply_archive(plan)
    return landed


def _apply_archive(plan: ImportPlan) -> List[Path]:
    """Кладёт вторые экземпляры срезов в папки их собственных дат.

    Всегда копия, даже когда основной режим — перенос: файл должен остаться и
    в папке отчётной даты, где он нужен как часть пары. Сбой копирования не
    валит запуск — отчёт на текущую дату от этого не страдает, а в лог уйдёт
    предупреждение.
    """
    copied: List[Path] = []
    for source_path, destination in plan.archive:
        if destination.exists() or not source_path.exists():
            continue
        try:
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(str(source_path), str(destination))
        except (OSError, shutil.Error) as exc:
            logger.warning(
                "Не удалось продублировать %s в папку его даты (%s): %s. "
                "На текущий отчёт это не влияет, но собрать отчёт на %s позже "
                "получится только после повторной выгрузки.",
                source_path.name, destination.parent, exc, destination.parent.name,
            )
            continue
        logger.info("Срез %s продублирован в папку своей даты: %s",
                    source_path.name, destination.parent)
        copied.append(destination)
    return copied


def available_dates(source: SourceConfig, downloads_dir: Optional[Path],
                    limit: int = 10) -> List[AvailableDate]:
    """Даты, на которые отчёт можно построить: из папок, из загрузок или из обоих.

    Считается тем же plan_import, что и сам запуск, — чтобы показанное в списке
    и то, что произойдёт при выборе, не могли разойтись.
    """
    candidate_dates = set()

    if source.uses_date_folders and Path(source.directory).is_dir():
        for entry in Path(source.directory).iterdir():
            if not entry.is_dir():
                continue
            try:
                candidate_dates.add(dt.datetime.strptime(entry.name, source.date_folder_format).date())
            except ValueError:
                continue

    downloads = scan_downloads(source, downloads_dir) if downloads_dir else []
    candidate_dates.update(c.business_date for c in downloads)

    rows: List[AvailableDate] = []
    for target_date in sorted(candidate_dates, reverse=True)[:limit]:
        plan = plan_import(source, downloads_dir, target_date)
        if not plan.complete:
            continue
        from_downloads = bool(plan.to_import)
        from_folder = bool(plan.existing)
        origin = ("папка + загрузки" if from_downloads and from_folder
                  else "загрузки" if from_downloads else "папка")
        files = [p.name for p in plan.existing] + [src.name for src, _dst in plan.to_import]
        rows.append(AvailableDate(date=target_date, origin=origin, files=files, complete=True))
    return rows


def resolve_pair(source: SourceConfig, downloads_dir: Optional[Path],
                 target_date: dt.date, move: bool = True,
                 archive_own_date: bool = True) -> List[Path]:
    """План + перекладывание одним вызовом: вернуть файлы срезов на дату."""
    plan = plan_import(source, downloads_dir, target_date, archive_own_date=archive_own_date)
    if not plan.complete:
        raise PortfolioDynamicsError(plan.problem)
    return apply_import(plan, move=move)
