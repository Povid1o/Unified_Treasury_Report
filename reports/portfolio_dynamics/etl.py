"""ETL отчёта «Динамика портфелей»: два среза выгрузки позиций (T0 и T-7)
плюс предыдущий выпуск отчёта -> четыре таблицы схемы v3.0.

Запись .xlsx живёт отдельно, в workbook.py: здесь только разбор входных
файлов и слияние с предыдущим выпуском.

Почему предыдущий выпуск — это ВХОД. Выгрузка даёт ровно два среза, а лист
fact_type_daily требует ежедневной истории объёмов по типам с 01.01.2026.
Взять её неоткуда, кроме как накапливать: каждый запуск дописывает в историю
одну дату (T0) и переносит всё остальное — историю прошлых дат, лимиты
(их заводят руками), справочник портфелей и заметки — из предыдущего файла.

Разбор иерархии. В выгрузке нет колонки с портфелем: иерархия задана
значениями внутри колонки «Тип актива». Строка «Позиция: <CODE>» открывает
портфель, все последующие строки — его бумаги, вплоть до следующей такой
строки. Строки до первого маркера — шапка выгрузки.
"""
import datetime as dt
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

BASE_DIR = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(BASE_DIR))

import config  # noqa: E402
from common import excel_io  # noqa: E402
from common.logging_utils import get_logger  # noqa: E402

logger = get_logger("portfolio_dynamics")


class PortfolioDynamicsError(RuntimeError):
    """Ошибка чтения или обработки входных файлов «Динамики портфелей»."""


# ── Колонки выходных таблиц (схема v3.0, см. CONTRACT.md) ────────────────────
DIM_COLUMNS = [
    "portfolio_code", "portfolio_name", "portfolio_type",
    "include_in_total", "is_limit_controlled", "sort_order",
]
LIMIT_COLUMNS = [
    "portfolio_type", "limit_amount", "green_max_util", "yellow_max_util",
    "red_max_util", "valid_from", "updated_by",
]
TYPE_DAILY_COLUMNS = ["business_date", "portfolio_type", "volume_amount"]
SNAPSHOT_COLUMNS = [
    "business_date", "portfolio_code", "volume_t0", "volume_t7",
    "duration_current_yrs", "duration_target_yrs", "note_text",
]

# ── Колонки входной выгрузки ─────────────────────────────────────────────────
# Сравнение идёт по excel_io.normalize_label (без пробелов, casefold, ё->е):
# в выгрузках у заголовков висячие пробелы, неразрывные пробелы и переносы.
COL_ASSET_TYPE = "Тип актива"
COL_VALUE = "Чистая стоимость позиции (кон.)"
COL_DURATION_START = "Duration (нач.)"
COL_DURATION_END = "Дюрация"

# Маппинг дюраций на схему намеренно «перекрёстный»: «Duration (нач.)» ->
# duration_current_yrs, «Дюрация» (конечная) -> duration_target_yrs. Имена не
# совпадают по смыслу с «нач.»/«кон.» — это осознанное решение при согласовании
# схемы v3.0, колонки не переименовывать и маппинг соблюдать буквально.

POSITION_MARKER = excel_io.normalize_label("Позиция:")
TOTAL_MARKERS = tuple(excel_io.normalize_label(x) for x in ("Итого", "Всего", "Grand Total", "Total"))

# Вторая дата из "Позиция за период ... - ... - SECURITIES" — и в имени файла,
# и в первой строке самого листа. Квадратные скобки необязательны: выгрузка
# именует ФАЙЛ без них и с произвольным числом пробелов
# ("Позиция за период   18.09.2026  -  18.09.2026   - SECURITIES.xlsx"),
# а в ШАПКЕ ЛИСТА пишет со скобками. Разбирать надо оба вида.
PERIOD_DATES_PATTERN = re.compile(
    r"\[?(\d{2}\.\d{2}\.\d{4})\]?\s*-\s*\[?(\d{2}\.\d{2}\.\d{4})\]?"
)

# "30,000,000" — запятая как разделитель тысяч; "3,14" — как десятичная.
_THOUSANDS_GROUPED = re.compile(r"^-?\d{1,3}(,\d{3})+$")
_SPACES = re.compile(r"[\s   ]+")
_MISSING_TOKENS = {"", "-", "—", "–", "...", "…", "н/д", "нд", "n/a", "na", "#н/д"}


# ════════════════════════════════════════════════════════════════════════════
# Разбор одного среза
# ════════════════════════════════════════════════════════════════════════════
@dataclass
class SliceStats:
    """Счётчики разбора — уходят в лог, чтобы молчаливая потеря строк была видна."""

    sheet_name: str = ""
    header_row: int = 0
    columns: Dict[str, int] = field(default_factory=dict)
    portfolios: int = 0
    securities: int = 0
    rows_without_portfolio: int = 0
    total_rows_skipped: int = 0
    subtotal_mismatches: List[str] = field(default_factory=list)


@dataclass
class PortfolioSlice:
    """Один срез выгрузки, свёрнутый до портфелей."""

    path: Path
    business_date: Optional[dt.date]
    frame: pd.DataFrame  # portfolio_code, volume, duration_current_yrs, duration_target_yrs
    stats: SliceStats


def parse_number(value) -> Optional[float]:
    """Число из ячейки выгрузки: и настоящее число, и текст с разделителями.

    В выгрузках встречается всё сразу: "1 234 567,89" (пробелы + запятая как
    десятичная), "30,000,000" (запятая как разделитель тысяч), неразрывные
    пробелы, прочерки и "..." вместо пропуска. Пустая ячейка и любой из
    прочерков -> None ("нет данных"), а не 0: ноль исказил бы и суммы, и веса.
    """
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return None if pd.isna(value) else float(value)

    text = _SPACES.sub("", str(value)).replace("−", "-")
    if text.casefold() in _MISSING_TOKENS:
        return None
    text = text.replace("₽", "").replace("руб.", "").replace("%", "")

    if "," in text and "." in text:
        # Смешанный формат: точка — десятичная, запятая — разделитель тысяч.
        text = text.replace(",", "")
    elif _THOUSANDS_GROUPED.match(text):
        text = text.replace(",", "")
    else:
        text = text.replace(",", ".")

    try:
        return float(text)
    except ValueError:
        return None


def _is_position_marker(normalized: str) -> bool:
    return normalized.startswith(POSITION_MARKER)


def _extract_portfolio_code(raw_value) -> str:
    """"Позиция: AFS_TR_RUR" -> "AFS_TR_RUR" (и "Позиция:AFS_TR_RUR" тоже)."""
    text = str(raw_value)
    _, _, tail = text.partition(":")
    return _SPACES.sub("", tail).upper()


def _find_sheet_and_header(path: Path) -> Tuple[str, pd.DataFrame, int, Dict[str, int]]:
    """Ищет лист по НАЛИЧИЮ колонки «Тип актива», а не по индексу.

    Выгрузки кладут перед нужным листом служебные (кэш сводной, параметры
    подключения), и «первый лист по умолчанию» читается пустым.
    """
    wanted = {
        COL_ASSET_TYPE: excel_io.normalize_label(COL_ASSET_TYPE),
        COL_VALUE: excel_io.normalize_label(COL_VALUE),
        COL_DURATION_START: excel_io.normalize_label(COL_DURATION_START),
        COL_DURATION_END: excel_io.normalize_label(COL_DURATION_END),
    }
    seen_sheets: List[str] = []

    for sheet_name, matrix in excel_io.iter_sheet_matrices(path):
        seen_sheets.append(sheet_name)
        if matrix.empty:
            continue
        for row_idx in range(len(matrix)):
            row = matrix.iloc[row_idx]
            normalized = [excel_io.normalize_label(v) for v in row]
            if wanted[COL_ASSET_TYPE] not in normalized:
                continue
            columns = {
                label: normalized.index(norm)
                for label, norm in wanted.items()
                if norm in normalized
            }
            return sheet_name, matrix, row_idx, columns

    explanation = excel_io.describe_uncached_formulas(path)
    message = (
        f"В файле {path.name} не найден лист с колонкой «{COL_ASSET_TYPE}». "
        f"Листы в файле: {seen_sheets}."
    )
    if explanation:
        message = f"{message}\n{explanation}"
    raise PortfolioDynamicsError(message)


def _business_date_from_name(path: Path) -> Optional[dt.date]:
    match = PERIOD_DATES_PATTERN.search(path.name)
    if not match:
        return None
    return dt.datetime.strptime(match.group(2), config.PORTFOLIO_DYNAMICS_DATE_FORMAT).date()


def _business_date_from_matrix(matrix: pd.DataFrame, header_row: int) -> Optional[dt.date]:
    """Та же строка «Позиция за период [..] - [..] - SECURITIES», но из шапки листа.

    Резерв на случай, если файл переименовали руками и дата из имени пропала.
    """
    for row_idx in range(min(header_row + 1, len(matrix))):
        for value in matrix.iloc[row_idx]:
            if value is None:
                continue
            match = PERIOD_DATES_PATTERN.search(str(value))
            if match:
                return dt.datetime.strptime(
                    match.group(2), config.PORTFOLIO_DYNAMICS_DATE_FORMAT
                ).date()
    return None


def read_business_date(path: Path) -> Optional[dt.date]:
    """Дата среза одного файла выгрузки: из имени файла, иначе из шапки листа.

    Разбор имени бесплатный, чтение книги — нет, поэтому лист читается только
    когда имя ничего не дало (файл переименовали руками).
    """
    path = Path(path)
    from_name = _business_date_from_name(path)
    if from_name is not None:
        return from_name
    try:
        _sheet, matrix, header_row, _cols = _find_sheet_and_header(path)
    except (PortfolioDynamicsError, excel_io.ExcelSourceError):
        return None
    return _business_date_from_matrix(matrix, header_row)


def split_slice_files(files: List[Path], folder_label: str = "") -> Tuple[Path, Path]:
    """Из файлов одной папки-даты выбирает, какой срез T0, а какой T-7.

    Определяется по дате самой выгрузки (вторая дата периода), а не по имени
    файла как строке: поздний срез — T0, ранний — T-7. Так пользователю не нужно
    ничего переименовывать и раскладывать в правильном порядке — достаточно
    положить в папку два файла.
    """
    where = f" в {folder_label}" if folder_label else ""
    real_files = [f for f in files if f.is_file() and not f.name.startswith("~$")]
    if len(real_files) < 2:
        raise PortfolioDynamicsError(
            f"Для отчёта нужны ДВА файла (срез T0 и срез T-7), а{where} найдено "
            f"{len(real_files)}: {', '.join(f.name for f in real_files) or 'ничего'}."
        )

    dated = [(read_business_date(f), f) for f in real_files]
    undated = [f.name for d, f in dated if d is None]
    if undated:
        raise PortfolioDynamicsError(
            f"Не удалось определить дату среза у файлов{where}: {', '.join(undated)}. "
            "В имени файла (или в шапке листа) должна быть строка вида "
            "«Позиция за период [дд.мм.гггг] - [дд.мм.гггг]»."
        )

    dated.sort(key=lambda item: item[0], reverse=True)
    if len(dated) > 2:
        ignored = ", ".join(f.name for _d, f in dated[2:])
        logger.warning(
            "В папке%s больше двух файлов — взяты два самых свежих среза (%s и %s), "
            "остальные проигнорированы: %s",
            where, dated[0][1].name, dated[1][1].name, ignored,
        )

    (t0_date, t0_path), (t7_date, t7_path) = dated[0], dated[1]
    if t0_date == t7_date:
        raise PortfolioDynamicsError(
            f"Оба файла{where} — срезы на одну и ту же дату {t0_date.isoformat()} "
            f"({t0_path.name}, {t7_path.name}). Нужны два РАЗНЫХ среза: T0 и T-7."
        )
    return t0_path, t7_path


def _weighted_duration(pairs: List[Tuple[float, float]]) -> Optional[float]:
    """Дюрация портфеля — средневзвешенная по стоимости, а не сумма и не среднее.

    pairs — (дюрация, вес) только по бумагам с заполненной дюрацией: бумага без
    дюрации исключается из весов, но её объём остаётся в сумме портфеля.
    Нулевая сумма весов -> None (делить на ноль нечем и нечего).
    """
    weight_sum = sum(weight for _duration, weight in pairs)
    if not weight_sum:
        return None
    return sum(duration * weight for duration, weight in pairs) / weight_sum


def parse_slice(path: Path, label: str, value_scale: Optional[float] = None) -> PortfolioSlice:
    """Разбирает один срез выгрузки и сворачивает его до портфелей."""
    path = Path(path)
    if not path.exists():
        raise PortfolioDynamicsError(f"[{label}] Файл не найден: {path}")

    scale = config.PORTFOLIO_DYNAMICS_VALUE_SCALE if value_scale is None else value_scale
    sheet_name, matrix, header_row, columns = _find_sheet_and_header(path)

    stats = SliceStats(sheet_name=sheet_name, header_row=header_row + 1, columns=dict(columns))
    for missing in (COL_VALUE, COL_DURATION_START, COL_DURATION_END):
        if missing not in columns:
            logger.warning(
                "[%s] %s: колонка «%s» не найдена на листе %r — соответствующие "
                "значения останутся пустыми.", label, path.name, missing, sheet_name,
            )
    if COL_VALUE not in columns:
        raise PortfolioDynamicsError(
            f"[{label}] В файле {path.name} (лист {sheet_name!r}) нет колонки «{COL_VALUE}» — "
            "объёмы портфелей брать неоткуда."
        )

    type_col = columns[COL_ASSET_TYPE]
    value_col = columns[COL_VALUE]

    order: List[str] = []
    securities: Dict[str, List[Tuple[Optional[float], Optional[float], Optional[float]]]] = {}
    stated: Dict[str, Dict[str, Optional[float]]] = {}
    current: Optional[str] = None

    for row_idx in range(header_row + 1, len(matrix)):
        row = matrix.iloc[row_idx]
        raw_type = row.iloc[type_col]
        normalized = excel_io.normalize_label(raw_type)
        if not normalized:
            continue

        if _is_position_marker(normalized):
            code = _extract_portfolio_code(raw_type)
            if not code:
                logger.warning(
                    "[%s] %s: строка %d — маркер портфеля без кода (%r), пропущена.",
                    label, path.name, row_idx + 1, raw_type,
                )
                continue
            if code not in securities:
                securities[code] = []
                order.append(code)
            current = code
            # Подытоги из самой строки маркера: если выгрузка их даёт, они
            # авторитетнее посчитанных по бумагам (см. сверку ниже).
            stated[code] = {
                "volume": parse_number(row.iloc[value_col]),
                "duration_current_yrs": (
                    parse_number(row.iloc[columns[COL_DURATION_START]])
                    if COL_DURATION_START in columns else None
                ),
                "duration_target_yrs": (
                    parse_number(row.iloc[columns[COL_DURATION_END]])
                    if COL_DURATION_END in columns else None
                ),
            }
            continue

        if normalized.startswith(TOTAL_MARKERS):
            stats.total_rows_skipped += 1
            continue

        if current is None:
            # Шапка выгрузки до первого маркера — в данные не попадает.
            stats.rows_without_portfolio += 1
            continue

        securities[current].append((
            parse_number(row.iloc[value_col]),
            parse_number(row.iloc[columns[COL_DURATION_START]]) if COL_DURATION_START in columns else None,
            parse_number(row.iloc[columns[COL_DURATION_END]]) if COL_DURATION_END in columns else None,
        ))
        stats.securities += 1

    stats.portfolios = len(order)
    if not order:
        raise PortfolioDynamicsError(
            f"[{label}] В файле {path.name} (лист {sheet_name!r}) не найдено ни одной строки "
            f"«{COL_ASSET_TYPE}» вида «Позиция: <КОД>» — портфели определить не по чему."
        )

    records = []
    for code in order:
        rows = securities[code]
        computed_volume = sum(value for value, _dc, _dt in rows if value is not None)
        computed = {
            "volume": computed_volume,
            "duration_current_yrs": _weighted_duration(
                [(dc, value) for value, dc, _dt in rows if dc is not None and value is not None]
            ),
            "duration_target_yrs": _weighted_duration(
                [(dtg, value) for value, _dc, dtg in rows if dtg is not None and value is not None]
            ),
        }
        chosen = {
            key: (stated[code][key] if stated[code][key] is not None else computed[key])
            for key in computed
        }

        declared = stated[code]["volume"]
        if declared is not None and _relative_diff(computed_volume, declared) > config.PORTFOLIO_DYNAMICS_TOLERANCE:
            message = (
                f"{code}: подытог в строке «Позиция: ...» = {declared:,.2f}, "
                f"сумма по бумагам = {computed_volume:,.2f}"
            )
            stats.subtotal_mismatches.append(message)
            logger.warning("[%s] %s: расхождение подытога и суммы по бумагам — %s", label, path.name, message)

        records.append({
            "portfolio_code": code,
            "volume": None if chosen["volume"] is None else chosen["volume"] / scale,
            "duration_current_yrs": chosen["duration_current_yrs"],
            "duration_target_yrs": chosen["duration_target_yrs"],
        })

    business_date = _business_date_from_name(path) or _business_date_from_matrix(matrix, header_row)

    logger.info(
        "[%s] %s: лист %r, шапка в строке %d, портфелей %d, бумаг %d, "
        "строк вне иерархии %d, строк «Итого» отброшено %d, дата среза %s",
        label, path.name, sheet_name, header_row + 1, stats.portfolios, stats.securities,
        stats.rows_without_portfolio, stats.total_rows_skipped,
        business_date.isoformat() if business_date else "не определена",
    )

    return PortfolioSlice(
        path=path,
        business_date=business_date,
        frame=pd.DataFrame(records, columns=["portfolio_code", "volume", "duration_current_yrs", "duration_target_yrs"]),
        stats=stats,
    )


def _relative_diff(computed: Optional[float], declared: Optional[float]) -> float:
    if declared is None or computed is None:
        return 0.0
    if declared == 0:
        return 0.0 if computed == 0 else float("inf")
    return abs(computed / declared - 1)


# ════════════════════════════════════════════════════════════════════════════
# Предыдущий выпуск отчёта
# ════════════════════════════════════════════════════════════════════════════
@dataclass
class PreviousRelease:
    """Четыре таблицы, перенос которых и делает отчёт инкрементальным."""

    path: Optional[Path]
    dim_portfolio: pd.DataFrame
    fact_limit: pd.DataFrame
    fact_type_daily: pd.DataFrame
    notes: pd.DataFrame  # portfolio_code, note_text

    @classmethod
    def empty(cls) -> "PreviousRelease":
        return cls(
            path=None,
            dim_portfolio=pd.DataFrame(columns=DIM_COLUMNS),
            fact_limit=pd.DataFrame(columns=LIMIT_COLUMNS),
            fact_type_daily=pd.DataFrame(columns=TYPE_DAILY_COLUMNS),
            notes=pd.DataFrame(columns=["portfolio_code", "note_text"]),
        )


OUTPUT_FILENAME_PREFIX = "dinamika_portfeley_"


def find_previous_release(output_dir: Path) -> Optional[Path]:
    """Самый свежий выпуск отчёта в выходной папке (по дате изменения).

    Файл за сегодняшний T0 тоже считается предыдущим выпуском — именно на этом
    держится идемпотентность: повторный прогон читает его и ЗАМЕНЯЕТ строки
    истории за свою дату, а не дописывает вторые такие же.
    """
    output_dir = Path(output_dir)
    if not output_dir.exists():
        return None
    candidates = [
        f for f in output_dir.glob(f"{OUTPUT_FILENAME_PREFIX}*.xlsx")
        if f.is_file() and not f.name.startswith("~$")
    ]
    if not candidates:
        return None
    return max(candidates, key=lambda f: f.stat().st_mtime)


def _read_sheet(path: Path, sheet_name: str, columns: List[str]) -> pd.DataFrame:
    try:
        frame = pd.read_excel(path, sheet_name=sheet_name)
    except ValueError as exc:  # листа нет в книге
        raise PortfolioDynamicsError(
            f"В предыдущем выпуске {path.name} нет листа {sheet_name!r}: {exc}. "
            "Похоже, это не файл «Динамики портфелей» — укажите другой через --previous "
            "или запустите с --bootstrap."
        ) from exc
    except Exception as exc:
        raise PortfolioDynamicsError(
            f"Не удалось прочитать лист {sheet_name!r} из предыдущего выпуска {path}: {exc}"
        ) from exc

    missing = [c for c in columns if c not in frame.columns]
    if missing:
        raise PortfolioDynamicsError(
            f"В предыдущем выпуске {path.name}, лист {sheet_name!r}: нет колонок {missing}. "
            "Строка 1 листа должна содержать технические заголовки snake_case."
        )
    return frame[columns].copy()


def load_previous_release(path: Path) -> PreviousRelease:
    """Читает из предыдущего выпуска ровно то, что скрипт не умеет получить из выгрузки."""
    path = Path(path)
    if not path.exists():
        raise PortfolioDynamicsError(f"Предыдущий выпуск не найден: {path}")

    dim = _read_sheet(path, "dim_portfolio", DIM_COLUMNS)
    limit = _read_sheet(path, "fact_limit", LIMIT_COLUMNS)
    history = _read_sheet(path, "fact_type_daily", TYPE_DAILY_COLUMNS)
    snapshot = _read_sheet(path, "fact_portfolio_snapshot", SNAPSHOT_COLUMNS)

    dim = dim[dim["portfolio_code"].notna()].copy()
    limit = limit[limit["portfolio_type"].notna()].copy()
    history = history[history["business_date"].notna() & history["portfolio_type"].notna()].copy()
    history["business_date"] = pd.to_datetime(history["business_date"], errors="coerce").dt.date
    history = history[history["business_date"].notna()]

    notes = snapshot[["portfolio_code", "note_text"]]
    notes = notes[notes["portfolio_code"].notna()].copy()

    return PreviousRelease(path=path, dim_portfolio=dim, fact_limit=limit,
                           fact_type_daily=history, notes=notes)


# ════════════════════════════════════════════════════════════════════════════
# Сборка четырёх таблиц
# ════════════════════════════════════════════════════════════════════════════
@dataclass
class PortfolioDynamicsData:
    """То, что etl отдаёт в workbook: четыре таблицы схемы v3.0 и параметры витрин."""

    business_date: dt.date
    lookback_days: int
    dim_portfolio: pd.DataFrame
    fact_limit: pd.DataFrame
    fact_type_daily: pd.DataFrame
    fact_portfolio_snapshot: pd.DataFrame
    new_portfolio_codes: List[str] = field(default_factory=list)
    missing_portfolio_codes: List[str] = field(default_factory=list)
    history_rows_carried: int = 0
    history_rows_added: int = 0
    history_rows_replaced: int = 0
    notes_restored: int = 0


def guess_type(portfolio_code: str) -> str:
    """Тип по префиксу кода до первого "_": AFS_TR_RUR -> AFS, HTM_ALCO -> HTM."""
    return str(portfolio_code).split("_", 1)[0].upper()


def _next_sort_order(dim: pd.DataFrame) -> int:
    values = pd.to_numeric(dim["sort_order"], errors="coerce").dropna()
    return int(values.max()) + 10 if len(values) else 10


def _build_dim(t0: PortfolioSlice, previous: PreviousRelease, bootstrap: bool
               ) -> Tuple[pd.DataFrame, List[str]]:
    """Справочник портфелей: предыдущий как есть + новые коды из T0.

    Тип нового кода угадывается по префиксу, и только если такой тип уже заведён
    в fact_limit — иначе остаётся пустым. В любом случае привязку подтверждает
    человек, поэтому список новых кодов уходит в лог предупреждением.
    """
    dim = previous.dim_portfolio.copy()
    if dim.empty:
        dim = pd.DataFrame(columns=DIM_COLUMNS)
    known_types = set(previous.fact_limit["portfolio_type"].astype(str)) if not previous.fact_limit.empty else set()
    if bootstrap:
        known_types = {guess_type(code) for code in t0.frame["portfolio_code"]}

    known_codes = set(dim["portfolio_code"].astype(str))
    new_codes = [c for c in t0.frame["portfolio_code"] if c not in known_codes]

    sort_order = _next_sort_order(dim)
    additions = []
    for code in new_codes:
        guessed = guess_type(code)
        additions.append({
            "portfolio_code": code,
            "portfolio_name": code,
            "portfolio_type": guessed if guessed in known_types else None,
            "include_in_total": True,
            "is_limit_controlled": True,
            "sort_order": sort_order,
        })
        sort_order += 10

    if additions:
        dim = pd.concat([dim, pd.DataFrame(additions, columns=DIM_COLUMNS)], ignore_index=True)

    dim["sort_order"] = pd.to_numeric(dim["sort_order"], errors="coerce")
    dim = dim.sort_values(["sort_order", "portfolio_code"], na_position="last").reset_index(drop=True)
    return dim[DIM_COLUMNS], new_codes


def _build_limits(dim: pd.DataFrame, previous: PreviousRelease, business_date: dt.date,
                  bootstrap: bool) -> pd.DataFrame:
    """Лимиты переносятся из предыдущего выпуска ЦЕЛИКОМ: скрипт их не трогает.

    Исключение — первый запуск: там заводятся строки типов, встреченных в T0, с
    нулевыми лимитами, чтобы файл читался и витрины считались. Заполнить их
    руками обязан человек (об этом пишется предупреждение в лог).
    """
    if not bootstrap or not previous.fact_limit.empty:
        return previous.fact_limit[LIMIT_COLUMNS].copy()

    types = sorted({guess_type(code) for code in dim["portfolio_code"]})
    return pd.DataFrame(
        [{
            "portfolio_type": t, "limit_amount": 0, "green_max_util": 0,
            "yellow_max_util": 0, "red_max_util": 0,
            "valid_from": business_date, "updated_by": None,
        } for t in types],
        columns=LIMIT_COLUMNS,
    )


def _build_snapshot(t0: PortfolioSlice, t7: PortfolioSlice, business_date: dt.date,
                    previous: PreviousRelease) -> Tuple[pd.DataFrame, int]:
    """Срез пересобирается целиком; note_text возвращается на место по portfolio_code.

    Это единственная ручная колонка на машинном листе — без merge полная
    перезапись листа стёрла бы заметки казначейства (CONTRACT.md, п. 2).
    """
    snapshot = t0.frame.rename(columns={"volume": "volume_t0"}).copy()
    volumes_t7 = t7.frame[["portfolio_code", "volume"]].rename(columns={"volume": "volume_t7"})
    snapshot = snapshot.merge(volumes_t7, on="portfolio_code", how="left")
    snapshot.insert(0, "business_date", business_date)

    notes = previous.notes.copy()
    if not notes.empty:
        notes = notes.drop_duplicates(subset=["portfolio_code"], keep="last")
        notes["note_text"] = notes["note_text"].where(notes["note_text"].notna(), None)
        snapshot = snapshot.merge(notes, on="portfolio_code", how="left")
    else:
        snapshot["note_text"] = None

    restored = int(snapshot["note_text"].notna().sum())
    return snapshot[SNAPSHOT_COLUMNS], restored


def _build_type_daily(snapshot: pd.DataFrame, dim: pd.DataFrame, limits: pd.DataFrame,
                      previous: PreviousRelease, business_date: dt.date
                      ) -> Tuple[pd.DataFrame, int, int, int]:
    """История по типам: прошлые даты как есть + свежая свёртка T0 по типам.

    Идемпотентность: строки за business_date из предыдущего выпуска ЗАМЕНЯЮТСЯ,
    а не дополняются. История за прошлые даты не переписывается никогда — она
    существует только в этом файле, восстановить её из выгрузок невозможно.
    """
    history = previous.fact_type_daily.copy()
    replaced = 0
    if not history.empty:
        same_date = history["business_date"] == business_date
        replaced = int(same_date.sum())
        history = history[~same_date]
    carried = len(history)

    types_frame = snapshot.merge(
        dim[["portfolio_code", "portfolio_type"]], on="portfolio_code", how="left"
    )
    sums = (
        types_frame[types_frame["portfolio_type"].notna()]
        .groupby("portfolio_type")["volume_t0"].sum()
    )

    # Одна строка на КАЖДЫЙ тип из fact_limit (в том числе на тип без портфелей):
    # так CHK_04 «строк на отчётную дату = число типов» остаётся выполнимой.
    known_types = [str(t) for t in limits["portfolio_type"]] if not limits.empty else []
    for extra in sums.index:
        if str(extra) not in known_types:
            known_types.append(str(extra))

    fresh = pd.DataFrame(
        [{
            "business_date": business_date,
            "portfolio_type": t,
            "volume_amount": float(sums.get(t, 0.0)),
        } for t in known_types],
        columns=TYPE_DAILY_COLUMNS,
    )

    combined = pd.concat([history, fresh], ignore_index=True) if carried else fresh
    combined = combined.sort_values(["business_date", "portfolio_type"]).reset_index(drop=True)
    return combined[TYPE_DAILY_COLUMNS], carried, len(fresh), replaced


def build_data(t0_path: Path, t7_path: Path, previous_path: Optional[Path] = None,
               bootstrap: bool = False) -> PortfolioDynamicsData:
    """Полный цикл ETL: два среза + предыдущий выпуск -> четыре таблицы схемы v3.0."""
    t0 = parse_slice(t0_path, "T0")
    t7 = parse_slice(t7_path, "T-7")

    if t0.business_date is None:
        raise PortfolioDynamicsError(
            f"Не удалось определить дату среза T0: ни в имени файла {Path(t0_path).name}, "
            "ни в шапке листа нет строки вида «Позиция за период [дд.мм.гггг] - [дд.мм.гггг]». "
            "Переименуйте файл по образцу выгрузки."
        )
    business_date = t0.business_date

    if t7.business_date is not None and t7.business_date >= business_date:
        logger.warning(
            "Дата среза T-7 (%s) не раньше даты T0 (%s) — проверьте, что файлы не перепутаны.",
            t7.business_date.isoformat(), business_date.isoformat(),
        )
    lookback = (
        (business_date - t7.business_date).days
        if t7.business_date is not None and t7.business_date < business_date
        else config.PORTFOLIO_DYNAMICS_DEFAULT_LOOKBACK
    )

    if bootstrap:
        previous = PreviousRelease.empty()
    elif previous_path is not None:
        previous = load_previous_release(previous_path)
        logger.info("Предыдущий выпуск: %s", previous.path)
    else:
        raise PortfolioDynamicsError(
            "Предыдущий выпуск отчёта не найден. История объёмов по типам, лимиты и "
            "заметки существуют только в нём и из выгрузок не восстанавливаются. "
            "Укажите файл через --previous либо запустите первый выпуск с --bootstrap."
        )

    dim, new_codes = _build_dim(t0, previous, bootstrap)
    limits = _build_limits(dim, previous, business_date, bootstrap)
    snapshot, notes_restored = _build_snapshot(t0, t7, business_date, previous)
    history, carried, added, replaced = _build_type_daily(snapshot, dim, limits, previous, business_date)

    missing = sorted(set(dim["portfolio_code"].astype(str)) - set(snapshot["portfolio_code"].astype(str)))

    if bootstrap:
        logger.warning(
            "Первый выпуск (--bootstrap): лист fact_limit заведён с НУЛЕВЫМИ лимитами по "
            "%d типам (%s). Заполните лимиты и границы зон руками — до этого светофор на "
            "view_by_type и проверки CHK_19/CHK_23/CHK_24 работать не будут.",
            len(limits), ", ".join(str(t) for t in limits["portfolio_type"]),
        )
    if new_codes:
        logger.warning(
            "Новых портфелей в T0: %d (%s) — дописаны в dim_portfolio. Проверьте и "
            "подтвердите привязку к типу руками.",
            len(new_codes), ", ".join(new_codes),
        )
    if missing:
        logger.warning(
            "Портфелей в dim_portfolio, которых нет в срезе T0: %d (%s) — CHK_08 покажет FAIL, "
            "пока они не будут удалены из справочника или не вернутся в выгрузку.",
            len(missing), ", ".join(missing),
        )
    empty_types = dim["portfolio_type"].isna().sum() + (dim["portfolio_type"].astype(str) == "").sum()
    if empty_types:
        logger.warning(
            "Портфелей без типа в dim_portfolio: %d — их объём не попадёт в fact_type_daily.",
            int(empty_types),
        )

    logger.info(
        "История fact_type_daily: перенесено %d строк, дописано %d за %s (заменено %d), "
        "заметок сохранено %d, портфелей в срезе %d, сдвиг сравнения %d дн.",
        carried, added, business_date.isoformat(), replaced, notes_restored, len(snapshot), lookback,
    )

    return PortfolioDynamicsData(
        business_date=business_date,
        lookback_days=lookback,
        dim_portfolio=dim,
        fact_limit=limits,
        fact_type_daily=history,
        fact_portfolio_snapshot=snapshot,
        new_portfolio_codes=new_codes,
        missing_portfolio_codes=missing,
        history_rows_carried=carried,
        history_rows_added=added,
        history_rows_replaced=replaced,
        notes_restored=notes_restored,
    )
