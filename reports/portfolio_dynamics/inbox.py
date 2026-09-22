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

from common.file_discovery import SourceConfig, find_date_folders  # noqa: E402
from reports.portfolio_dynamics.etl import (  # noqa: E402
    PROBE_SUFFIXES, PortfolioDynamicsError, logger, probe_business_date,
    read_business_date,
)
from reports.portfolio_dynamics.limits import probe_limits_date  # noqa: E402


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
    to_import: List[Tuple[Path, Path]] = field(default_factory=list)  # (откуда, куда) — из загрузок
    reuse: List[Tuple[Path, Path]] = field(default_factory=list)  # (откуда, куда) — из другой папки-даты
    archive: List[Tuple[Path, Path]] = field(default_factory=list)  # копия в папку своей даты
    limits: Optional[Path] = None  # файл лимитов на эту дату, если нашёлся
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
        parts = []
        if self.to_import:
            parts.append("из загрузок: " + ", ".join(src.name for src, _dst in self.to_import))
        if self.reuse:
            parts.append("из папок-дат: " + ", ".join(
                "%s (%s)" % (src.name, src.parent.name) for src, _dst in self.reuse))
        if not parts:
            return f"оба среза уже лежат в {self.folder.name}"
        return "будет взято — " + "; ".join(parts)


@dataclass(frozen=True)
class AvailableDate:
    """Строка для выбора даты: что на эту дату есть и откуда оно возьмётся."""

    date: dt.date
    origin: str        # «папка», «загрузки», «папки-даты» или их сочетание
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


def _parse_with_format(text: str, date_format: str) -> Optional[dt.date]:
    """Дата из строки, не придираясь к разделителю.

    Выгрузка называет файл то через «_», то через «.» — привязываться к одному
    написанию значит регулярно не находить файл на ровном месте.
    """
    variants = {text}
    for separator in ("_", ".", "-"):
        variants.add(re.sub(r"[._-]", separator, text))
    formats = {date_format}
    for separator in ("_", ".", "-"):
        formats.add(re.sub(r"[._-]", separator, date_format))
    for variant in variants:
        for fmt in formats:
            try:
                return dt.datetime.strptime(variant, fmt).date()
            except ValueError:
                continue
    return None


def _dated(paths: List[Path], source: SourceConfig) -> List[Candidate]:
    """Файлы с их датами среза; файлы без распознаваемой даты отбрасываются."""
    pattern = re.compile(source.filename_regex) if source.filename_regex else None
    result: List[Candidate] = []
    for path in paths:
        parsed: Optional[dt.date] = None
        if pattern is not None:
            match = pattern.search(path.name)
            if match:
                parsed = _parse_with_format(match.group(1), source.date_format)
        if parsed is None:
            parsed = read_business_date(path)  # дороже: открывает книгу
        if parsed is not None:
            result.append(Candidate(path=path, business_date=parsed))
    result.sort(key=lambda c: c.business_date, reverse=True)
    return result


# Сколько чужих книг в загрузках готовы открыть, разыскивая выгрузку по
# содержимому. Папка загрузок у людей на тысячи файлов, поэтому перебор идёт
# от самых свежих и обрывается: выгрузка, скачанная позавчера, в эту сотню
# попадёт, а прошлогодний курсовик — нет.
DEEP_PROBE_LIMIT = 80

# Что уже открывали: (путь, время правки, размер, чем проверяли) -> дата или None.
# За один запуск папка загрузок перебирается многократно (список доступных дат
# проверяет каждую дату отдельным планом), а открытие книги — единственное, что
# в этом переборе стоит времени. Ключ включает время и размер, поэтому
# подменённый файл перечитывается.
_PROBE_CACHE: dict = {}


def forget_probe_cache() -> None:
    """Сбросить память о проверенных книгах (нужно тестам и повторным запускам)."""
    _PROBE_CACHE.clear()


def _probe_cached(path: Path, probe) -> Tuple[Optional[dt.date], bool]:
    """(дата, впервые ли проверили) — чтобы не писать в лог одно и то же по кругу."""
    try:
        stat = path.stat()
        key = (str(path), stat.st_mtime, stat.st_size, getattr(probe, "__name__", str(probe)))
    except OSError:
        return None, False
    if key in _PROBE_CACHE:
        return _PROBE_CACHE[key], False
    parsed = probe(path)
    _PROBE_CACHE[key] = parsed
    return parsed, True


def scan_slices(source: SourceConfig, downloads_dir: Path,
                deep: bool = True) -> List[Candidate]:
    """Срезы позиций в загрузках."""
    return scan_downloads(source, downloads_dir, probe=probe_business_date, deep=deep)


def scan_limits(source: SourceConfig, downloads_dir: Path,
                deep: bool = True) -> List[Candidate]:
    """Выгрузки лимитов в загрузках."""
    return scan_downloads(source, downloads_dir, probe=probe_limits_date, deep=deep)


def scan_downloads(source: SourceConfig, downloads_dir: Path,
                   probe=None, deep: bool = True) -> List[Candidate]:
    """Выгрузки отчёта, лежащие в папке загрузок, свежие сначала.

    Ищет ДВУМЯ способами. Сначала по имени файла — это бесплатно. Потом, если
    probe задан, по СОДЕРЖИМОМУ остальных книг: выгрузку легко переименовать
    (браузер дописывает «(1)», человек правит имя руками, система однажды
    меняет формулировку), и привязка к одному шаблону имени означала бы, что
    файл, лежащий прямо в загрузках, отчёт «не видит» без всяких объяснений.
    Содержимое же подделать нечем: probe требует и шапку нужной выгрузки, и
    дату в ней.

    Недоступная папка загрузок — не ошибка: значит, приёмке просто неоткуда
    брать файлы, и отчёт работает как раньше, по своей папке.
    """
    downloads_dir = Path(downloads_dir)
    if not downloads_dir.is_dir():
        logger.info("Папка загрузок недоступна (%s) — приёмка из загрузок пропущена.", downloads_dir)
        return []

    pattern = re.compile(source.filename_regex) if source.filename_regex else None
    by_name = [
        f for f in downloads_dir.glob(source.glob_pattern)
        if _is_real_file(f) and (pattern is None or pattern.search(f.name))
    ]
    found = _dated(by_name, source)
    if not deep or probe is None or pattern is None:
        return found

    found += _probe_the_rest(downloads_dir, {f.name for f in by_name}, source, probe)
    found.sort(key=lambda c: c.business_date, reverse=True)
    return found


def _probe_the_rest(downloads_dir: Path, already: set,
                    source: SourceConfig, probe) -> List[Candidate]:
    """Книги, не подошедшие по имени, — вдруг это всё-таки наша выгрузка."""
    try:
        others = [f for f in downloads_dir.iterdir()
                  if _is_real_file(f) and f.suffix.lower() in PROBE_SUFFIXES
                  and f.name not in already]
    except OSError as exc:
        logger.info("Папку загрузок не удалось прочитать (%s): %s", downloads_dir, exc)
        return []

    def freshness(path: Path) -> float:
        try:
            return path.stat().st_mtime
        except OSError:
            return 0.0

    others.sort(key=freshness, reverse=True)
    if len(others) > DEEP_PROBE_LIMIT:
        logger.debug("В загрузках %d книг, по содержимому проверены %d самых свежих.",
                     len(others), DEEP_PROBE_LIMIT)

    result: List[Candidate] = []
    for path in others[:DEEP_PROBE_LIMIT]:
        parsed, first_time = _probe_cached(path, probe)
        if parsed is None:
            continue
        if first_time:
            logger.info(
                "Загрузки: имя «%s» под шаблон не подходит, но по содержимому это "
                "выгрузка на %s — файл взят. Если такое имя теперь постоянное, "
                "поправьте шаблон в настройках.",
                path.name, parsed.isoformat(),
            )
        result.append(Candidate(path=path, business_date=parsed))
    return result


def _where_we_looked(folder: Path, downloads_dir: Optional[Path],
                     other_folders: int = 0) -> str:
    """Человекочитаемое «где искали» — чтобы опечатка в пути была видна сразу."""
    places = [str(folder)]
    if other_folders:
        places.append("%d других папках-датах" % other_folders)
    if downloads_dir is not None:
        places.append("загрузках (%s)" % downloads_dir)
    tail = " (приёмка из загрузок выключена)" if downloads_dir is None else ""
    if len(places) == 1:
        return "в %s%s" % (places[0], tail)
    return "ни в " + ", ни в ".join(places) + tail


def slices_in_other_date_folders(source: SourceConfig, folder: Path,
                                 limits_source: Optional[SourceConfig]) -> List[Candidate]:
    """Срезы, уже разложенные по ДРУГИМ папкам-датам этого же источника.

    Без этого приёмка противоречила сама себе: она раскладывает каждый срез по
    папке его собственной даты и вычищает загрузки, а пару потом ищет только в
    папке отчётной даты и в загрузках. Через неделю-другую работы срез T-7
    оказывался единственным экземпляром в СВОЕЙ папке — и отчёт объявлял, что
    файла нет, стоя ровно в той папке, куда сам его и положил.
    """
    if not source.uses_date_folders:
        return []
    result: List[Candidate] = []
    for entry in find_date_folders(source):
        if entry.path == folder:
            continue
        files = [f for f in entry.files if not _is_limits(limits_source, f)]
        result += _dated(files, source)
    result.sort(key=lambda c: c.business_date, reverse=True)
    return result


def _files_in_folder(source: SourceConfig, folder: Path) -> List[Path]:
    if not folder.is_dir():
        return []
    return sorted(f for f in folder.glob(source.glob_pattern) if _is_real_file(f))


def _plan_limits(limits_source: Optional[SourceConfig], downloads_dir: Optional[Path],
                 folder: Path, target_date: dt.date) -> Tuple[Optional[Path], List[Tuple[Path, Path]]]:
    """Файл лимитов на отчётную дату: уже в папке или его надо взять из загрузок.

    Лимиты приходят ТРЕТЬИМ файлом на ту же дату и едут в ту же папку. Их
    отсутствие — не ошибка: тогда лимиты переносятся из предыдущего выпуска.
    """
    if limits_source is None:
        return None, []

    # Строго по шаблону имени, без запасного чтения книги: у выгрузки позиций
    # дата в шапке листа тоже разбирается, и нестрогий поиск принял бы срез за
    # файл лимитов.
    in_folder = [f for f in _files_in_folder(limits_source, folder)
                 if _is_limits(limits_source, f)]
    candidates = [(c, False) for c in _dated(in_folder, limits_source)]
    if downloads_dir:
        here = {c.path.name for c, _ in candidates}
        candidates += [(c, True) for c in scan_limits(limits_source, downloads_dir)
                       if c.path.name not in here]

    # Лимиты из будущего брать нельзя; на прошлую дату — можно, они меняются
    # редко, и отчёт без лимитов вовсе бесполезнее отчёта с чуть устаревшими.
    usable = [(c, from_downloads) for c, from_downloads in candidates
              if c.business_date <= target_date]
    if not usable:
        return None, []

    exact = [item for item in usable if item[0].business_date == target_date]
    chosen, from_downloads = (exact or sorted(
        usable, key=lambda item: item[0].business_date, reverse=True))[0]

    if chosen.business_date != target_date:
        logger.warning(
            "Файла лимитов на %s нет — взят ближайший более ранний, на %s (%s). "
            "Если лимиты с тех пор менялись, выгрузите их на отчётную дату.",
            target_date.isoformat(), chosen.business_date.isoformat(), chosen.path.name,
        )

    if not from_downloads:
        return chosen.path, []
    destination = folder / chosen.path.name
    return destination, [(chosen.path, destination)]


def plan_import(source: SourceConfig, downloads_dir: Optional[Path],
                target_date: dt.date, archive_own_date: bool = True,
                limits_source: Optional[SourceConfig] = None) -> ImportPlan:
    """Собирает план: что уже на месте, что взять из загрузок, чего не хватает.

    Ничего не перекладывает — только считает. Так план можно показать
    пользователю до того, как файлы поедут, и так же им пользуется экран
    выбора даты.
    """
    folder = folder_for_date(source, target_date)
    plan = ImportPlan(target_date=target_date, folder=folder)

    slice_files = [f for f in _files_in_folder(source, folder) if not _is_limits(limits_source, f)]
    in_folder = _dated(slice_files, source)
    if source.uses_date_folders and len(slice_files) >= 2:
        # Папка-дата уже укомплектована срезами — за ними в загрузки не идём,
        # но файл лимитов всё равно может там лежать и быть нужен.
        plan.existing = slice_files
        if archive_own_date:
            plan.archive = _plan_archive(source, folder, in_folder)
        plan.limits, limits_import = _plan_limits(limits_source, downloads_dir, folder, target_date)
        plan.to_import.extend(limits_import)
        return plan

    # Три источника, в порядке убывания «своего»: папка отчётной даты, другие
    # папки-даты (туда приёмка сама разложила прежние срезы) и загрузки.
    archived = slices_in_other_date_folders(source, folder, limits_source)
    downloads = [c for c in (scan_slices(source, downloads_dir) if downloads_dir else [])
                 if not _is_limits(limits_source, c.path)]
    archived_paths = {c.path for c in archived}

    pool: List[Candidate] = list(in_folder)
    taken_names = {c.path.name for c in in_folder}
    already_filed = []
    for candidate in archived + downloads:
        if candidate.path.name in taken_names:
            # Тот же файл, просто лежит ещё где-то. Второй экземпляр не нужен,
            # но и удалять его нельзя — приёмка ничего не удаляет.
            if candidate in downloads:
                already_filed.append(candidate.path.name)
            continue
        taken_names.add(candidate.path.name)
        pool.append(candidate)
    if already_filed:
        logger.info(
            "В загрузках лежат дубли уже разложенных срезов (%s) — отчёт их не "
            "берёт и не удаляет. Их можно убрать руками.", ", ".join(already_filed),
        )
    # Сортировка стабильная, поэтому при равной дате остаётся тот экземпляр,
    # что ближе: сначала из папки отчётной даты, потом из архива, потом из загрузок.
    pool.sort(key=lambda c: c.business_date, reverse=True)

    where = _where_we_looked(folder, downloads_dir, len({c.path.parent for c in archived}))

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
        elif candidate.path in archived_paths:
            # Из чужой папки-даты файл КОПИРУЕТСЯ, а не переносится: он там не
            # лишний, а единственный экземпляр отчёта на свою дату.
            plan.reuse.append((candidate.path, folder / candidate.path.name))
        else:
            plan.to_import.append((candidate.path, folder / candidate.path.name))
    if archive_own_date:
        plan.archive = _plan_archive(source, folder, [t0, t7])
    plan.limits, limits_import = _plan_limits(limits_source, downloads_dir, folder, target_date)
    plan.to_import.extend(limits_import)
    return plan


def _is_limits(limits_source: Optional[SourceConfig], path: Path) -> bool:
    """Файл подходит под шаблон имени выгрузки лимитов?"""
    if limits_source is None or not limits_source.filename_regex:
        return False
    return re.search(limits_source.filename_regex, Path(path).name) is not None


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
    if not plan.to_import and not plan.reuse:
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

    landed += _apply_reuse(plan)
    _apply_archive(plan)
    return landed


def _apply_reuse(plan: ImportPlan) -> List[Path]:
    """Копирует срезы, взятые из ДРУГИХ папок-дат, в папку отчётной даты.

    Всегда копия: в своей папке файл — единственный экземпляр отчёта на ту дату,
    и переносить его оттуда значило бы чинить сегодняшний отчёт, ломая прошлый.
    Сбой копирования не валит запуск: отчёт соберётся и по исходному пути, —
    но в лог уйдёт предупреждение.
    """
    landed: List[Path] = []
    for source_path, destination in plan.reuse:
        if destination.exists():
            landed.append(destination)
            continue
        try:
            shutil.copy2(str(source_path), str(destination))
        except (OSError, shutil.Error) as exc:
            logger.warning(
                "Не удалось скопировать %s из папки %s в %s: %s. Отчёт собран по "
                "исходному пути.", source_path.name, source_path.parent.name,
                plan.folder.name, exc,
            )
            landed.append(source_path)
            continue
        logger.info("Срез %s взят из папки %s и скопирован в %s.",
                    source_path.name, source_path.parent.name, plan.folder.name)
        landed.append(destination)
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


def resolve_limits_file(limits_source: Optional[SourceConfig], downloads_dir: Optional[Path],
                        folder: Path, target_date: dt.date, move: bool = True) -> Optional[Path]:
    """Файл лимитов на дату: найти в папке, иначе забрать из загрузок. None — нет.

    Отсутствие файла лимитов не ошибка: тогда лимиты переносятся из предыдущего
    выпуска, как было до появления этой выгрузки.
    """
    path, to_import = _plan_limits(limits_source, downloads_dir, folder, target_date)
    if path is None:
        return None
    for src, dst in to_import:
        if dst.exists():
            continue
        try:
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(src), str(dst)) if move else shutil.copy2(str(src), str(dst))
        except (OSError, shutil.Error) as exc:
            logger.warning("Не удалось взять файл лимитов %s из загрузок: %s", src.name, exc)
            return dst if dst.exists() else None
        logger.info("%s из загрузок: %s -> %s",
                    "Перенесён файл лимитов" if move else "Скопирован файл лимитов",
                    src.name, folder)
    return path if path.exists() else None


def available_dates(source: SourceConfig, downloads_dir: Optional[Path],
                    limit: int = 10,
                    limits_source: Optional[SourceConfig] = None) -> List[AvailableDate]:
    """Даты, на которые отчёт можно построить: из папок, из загрузок или из обоих.

    Считается тем же plan_import, что и сам запуск, — чтобы показанное в списке
    и то, что произойдёт при выборе, не могли разойтись.
    """
    candidate_dates = set()

    if source.uses_date_folders and Path(source.directory).is_dir():
        for entry in find_date_folders(source):
            candidate_dates.add(entry.date)
            # Дата ПАПКИ и даты лежащих в ней срезов — разные вещи: в папке
            # 2026-09-18 лежит и срез за 11.09. Отчёт на 11.09 тоже можно
            # собрать, поэтому в кандидаты идут и даты самих файлов.
            files = [f for f in entry.files if not _is_limits(limits_source, f)]
            candidate_dates.update(c.business_date for c in _dated(files, source))

    downloads = scan_slices(source, downloads_dir) if downloads_dir else []
    candidate_dates.update(c.business_date for c in downloads)

    rows: List[AvailableDate] = []
    for target_date in sorted(candidate_dates, reverse=True)[:limit]:
        # limits_source передаётся обязательно: без него файл лимитов, лежащий
        # в папке-дате, считается срезом, папка выглядит укомплектованной, и
        # дата попадает в список как готовая — а при запуске оказывается, что
        # второго среза нет. Список и запуск обязаны видеть одно и то же.
        plan = plan_import(source, downloads_dir, target_date, limits_source=limits_source)
        if not plan.complete:
            continue
        origin = " + ".join(
            [word for word, present in (("папка", bool(plan.existing)),
                                        ("папки-даты", bool(plan.reuse)),
                                        ("загрузки", bool(plan.to_import)))
             if present]) or "папка"
        files = ([p.name for p in plan.existing]
                 + [src.name for src, _dst in plan.reuse]
                 + [src.name for src, _dst in plan.to_import])
        rows.append(AvailableDate(date=target_date, origin=origin, files=files, complete=True))
    return rows


def resolve_pair(source: SourceConfig, downloads_dir: Optional[Path],
                 target_date: dt.date, move: bool = True,
                 archive_own_date: bool = True,
                 limits_source: Optional[SourceConfig] = None) -> List[Path]:
    """План + перекладывание одним вызовом: вернуть файлы срезов на дату."""
    plan = plan_import(source, downloads_dir, target_date, archive_own_date=archive_own_date,
                       limits_source=limits_source)
    if not plan.complete:
        raise PortfolioDynamicsError(plan.problem)
    return apply_import(plan, move=move)
