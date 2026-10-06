"""История рыночных показателей «Отчёта по портфелям»: RUONIA, RGBI, RWA.

Показатели вводятся руками при каждом запуске, и каждый запуск дописывает их
в файл истории — полный бэкап всего, что когда-либо вводили. Из той же
истории в CSV для BI попадают даты, которых там ещё нет: среднее с начала
года и прочие расчёты по истории BI делает сам по своей базе.

Файлов два, и это нарочно.

- Рабочий (настройка portfolio_report_market_history, по умолчанию
  market_history.csv рядом с settings.json) — сам бэкап. Как settings.json и
  portfolio_types.json, у каждой установки он свой: в git и в архив
  export_project.sh не попадает. Иначе распаковка свежего архива поверх
  рабочей копии затёрла бы всё, что накопилось на рабочем компьютере.
- Начальный (seed/market_history_seed.csv или .xlsx в папке проекта) —
  история, собранная руками один раз. Едет на рабочий компьютер вместе с
  проектом. Каждый запуск ДОПОЛНЯЕТ из него рабочий файл: только даты и
  показатели, которых в рабочем файле нет. Уже записанное он не перебивает.

Кроме начального файла, рабочий дополняется значениями из уже выпущенных
CSV в папке результатов: то, что вводили до появления файла истории, тоже
попадает в бэкап.

Формат обоих файлов — таблица с колонками «date» (или «Дата»), RUONIA, RGBI,
RWA; пустая ячейка — нет значения. Дата — та, с которой значение идёт в BI,
то есть дата позиций выпуска (T-1), как у портфелей. Даты — 2026-09-29 или
29.09.2026; числа — через точку или запятую. CSV из русского Excel (через «;»,
в кодировке Windows-1251) читается как есть. RWA — в рублях, как его вводят.
"""
import datetime as dt
import io
import os
import sys
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd

BASE_DIR = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(BASE_DIR))

import config  # noqa: E402
from common.logging_utils import get_logger  # noqa: E402
from reports.portfolio_report import etl  # noqa: E402

logger = get_logger("portfolio_report")

# {дата: {ключ показателя (rgbi/ruonia/rwa): значение}}
History = Dict[dt.date, Dict[str, float]]

KEYS = [key for key, _title, _unit in etl.MARKET_METRICS]
TITLES = {key: title for key, title, _unit in etl.MARKET_METRICS}
_BY_TITLE = {title.casefold(): key for key, title in TITLES.items()}
_DATE_HEADERS = {"date", "date_", "дата"}
FILE_COLUMNS = ["date"] + [TITLES[key] for key in KEYS]

SEED_DIR = BASE_DIR / "seed"
SEED_NAMES = ("market_history_seed.csv", "market_history_seed.xlsx")


def history_path() -> Path:
    return Path(config.PORTFOLIO_REPORT_MARKET_HISTORY)


def seed_path() -> Optional[Path]:
    """Начальный файл истории или None, если его в проекте нет."""
    for name in SEED_NAMES:
        path = SEED_DIR / name
        if path.is_file():
            return path
    return None


# ════════════════════════════════════════════════════════════════════════════
# Чтение и запись файла истории
# ════════════════════════════════════════════════════════════════════════════
def _parse_date(value) -> Optional[dt.date]:
    if value is None or (not isinstance(value, str) and pd.isna(value)):
        return None
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%Y-%m-%d %H:%M:%S", "%d.%m.%Y %H:%M:%S", "%d.%m.%y"):
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def _read_table(path: Path) -> pd.DataFrame:
    if path.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
        return pd.read_excel(path, dtype=object)
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "cp1251"):
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    else:
        raise etl.PortfolioReportError(f"{path.name}: не удалось определить кодировку файла.")
    header = text.splitlines()[0] if text.strip() else ""
    # Русский Excel сохраняет CSV через «;» — запятая там десятичная.
    sep = ";" if ";" in header else "\t" if "\t" in header else ","
    return pd.read_csv(io.StringIO(text), sep=sep, dtype=str, keep_default_na=False)


def read_history(path: Path) -> History:
    """Файл истории -> {дата: {показатель: значение}}. Нет файла — пустая история."""
    path = Path(path)
    if not path.is_file():
        return {}
    try:
        table = _read_table(path)
    except etl.PortfolioReportError:
        raise
    except Exception as exc:  # битый файл не должен молча обнулять историю
        raise etl.PortfolioReportError(f"Не удалось прочитать историю {path}: {exc}") from exc

    columns: Dict[str, str] = {}
    date_column = None
    for column in table.columns:
        name = str(column).strip().casefold()
        if name in _DATE_HEADERS:
            date_column = column
        elif name in _BY_TITLE:
            columns[_BY_TITLE[name]] = column
    if date_column is None:
        raise etl.PortfolioReportError(
            f"{path.name}: нет колонки с датой («date» или «Дата»). Ожидаются колонки "
            f"{', '.join(FILE_COLUMNS)}.")
    if not columns:
        raise etl.PortfolioReportError(
            f"{path.name}: нет ни одной из колонок {', '.join(TITLES.values())}.")

    history: History = {}
    for index, row in table.iterrows():
        raw_date = row[date_column]
        day = _parse_date(raw_date)
        if day is None:
            if str(raw_date).strip() and not pd.isna(raw_date):
                logger.warning("%s: строка %d — дата «%s» не разобрана, пропущена.",
                               path.name, index + 2, raw_date)
            continue
        for key, column in columns.items():
            value = etl.positions.parse_number(row[column])
            if value is not None:
                history.setdefault(day, {})[key] = value
    return history


def write_history(path: Path, history: History) -> Path:
    """Пишет историю целиком, через временный файл: оборванная запись не
    должна оставить бэкап полупустым."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    rows = [[day.isoformat()] + [history[day].get(key) for key in KEYS]
            for day in sorted(history)]
    frame = pd.DataFrame(rows, columns=FILE_COLUMNS)
    tmp = path.with_name(path.name + ".tmp")
    try:
        frame.to_csv(tmp, index=False, encoding="utf-8-sig")
        os.replace(str(tmp), str(path))
    except PermissionError as exc:
        raise etl.PortfolioReportError(
            f"Нет доступа для записи в {path} (файл открыт в Excel?): {exc}. "
            "Закройте файл и повторите запуск.") from exc
    return path


def fill_gaps(history: History, extra: History) -> int:
    """Дописывает в history показатели из extra, которых в ней нет. Сколько дописано."""
    added = 0
    for day, values in extra.items():
        own = history.setdefault(day, {})
        for key, value in values.items():
            if key not in own:
                own[key] = value
                added += 1
    for day in [d for d, values in history.items() if not values]:
        del history[day]
    return added


# ════════════════════════════════════════════════════════════════════════════
# Выпуски в папке результатов
# ════════════════════════════════════════════════════════════════════════════
def _release_files(output_dir: Path, skip: Optional[dt.date] = None) -> List[Path]:
    output_dir = Path(output_dir)
    if not output_dir.is_dir():
        return []
    files = []
    for path in output_dir.glob(f"{etl.OUTPUT_FILENAME_PREFIX}*.csv"):
        when = etl.release_date(path)
        if path.is_file() and when is not None and when != skip:
            files.append(path)
    return sorted(files)


def released(output_dir: Path, skip: Optional[dt.date] = None) -> History:
    """Рыночные показатели, уже записанные в CSV папки результатов (в базовых единицах).

    skip — дата выпуска, который сейчас перезаписывается: его содержимое
    собирается заново и уже выгруженным не считается.
    """
    found: History = {}
    for path in _release_files(output_dir, skip):
        try:
            flat = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
        except (OSError, ValueError, UnicodeDecodeError) as exc:
            logger.warning("Выпуск %s не прочитан (%s) — его рыночные показатели не учтены.",
                           path.name, exc)
            continue
        if not {"date_", "axis_1", "axis_3", "value"} <= set(flat.columns):
            continue
        has_unit = "axis_4" in flat.columns
        for row in flat.itertuples(index=False):
            key = _BY_TITLE.get(str(row.axis_3).strip().casefold())
            day = _parse_date(row.date_)
            value = etl.positions.parse_number(row.value)
            if row.axis_1 != etl.MARKET_GROUP or key is None or day is None or value is None:
                continue
            value = etl._to_base_unit(value, key, row.axis_4 if has_unit else None, path)
            found.setdefault(day, {})[key] = value
    return found


def pending(history: History, already: History, until: dt.date) -> History:
    """Показатели истории по until включительно, которых ещё нет ни в одном выпуске."""
    result: History = {}
    for day, values in history.items():
        if day > until:
            continue
        sent = already.get(day, {})
        for key, value in values.items():
            if key not in sent:
                result.setdefault(day, {})[key] = value
    return result


# ════════════════════════════════════════════════════════════════════════════
# Запуск
# ════════════════════════════════════════════════════════════════════════════
def load(output_dir: Optional[Path] = None, skip: Optional[dt.date] = None
         ) -> Tuple[History, History]:
    """(вся история, уже выгруженное в CSV папки результатов).

    История — рабочий файл, дополненный уже выпущенными CSV и начальным файлом.
    """
    output_dir = Path(output_dir or config.PORTFOLIO_REPORT_OUTPUT_DIR)
    history = read_history(history_path())
    already = released(output_dir, skip)
    from_releases = fill_gaps(history, already)
    seed = seed_path()
    from_seed = fill_gaps(history, read_history(seed)) if seed is not None else 0
    if from_releases or from_seed:
        logger.info("История рынка дополнена: из выпусков %d знач., из начального файла %d знач.",
                    from_releases, from_seed)
    return history, already


def apply_inputs(history: History, day: dt.date, market: "etl.MarketInputs") -> "etl.MarketInputs":
    """Введённое на day -> в историю; не введённое берётся из истории на day.

    Возвращает показатели выпуска: введённое руками важнее записанного раньше,
    пустой ввод запись не стирает (повторный прогон без ввода ничего не теряет).
    """
    own = history.setdefault(day, {})
    for key in KEYS:
        value = getattr(market, key)
        if value is None:
            continue
        before = own.get(key)
        if before is not None and before != value:
            logger.warning("%s на %s: было %s, записано %s.", TITLES[key], day.isoformat(),
                           f"{before:,.4f}", f"{value:,.4f}")
        own[key] = value
    if not own:
        del history[day]
        return etl.MarketInputs()
    return etl.MarketInputs(**{key: own.get(key) for key in KEYS})


def latest_before(history: History, day: dt.date) -> "etl.MarketInputs":
    """Последнее известное до day значение каждого показателя (для подсказки при вводе)."""
    result = {}
    for key in KEYS:
        dates = [d for d, values in history.items() if d < day and key in values]
        if dates:
            result[key] = history[max(dates)][key]
    return etl.MarketInputs(**result)


def dates_span(history: Iterable[dt.date]) -> str:
    dates = sorted(history)
    if not dates:
        return "нет дат"
    if len(dates) == 1:
        return dates[0].isoformat()
    return f"{dates[0].isoformat()} … {dates[-1].isoformat()} ({len(dates)} дат)"

