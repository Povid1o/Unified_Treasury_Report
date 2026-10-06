"""ETL «Отчёта по портфелям»: выгрузка «Позиция за период» с начала года по T-1 -> свод по портфелям.

На выходе — плоский CSV для загрузки в BI, в той же раскладке, что у ЧПД и
NIM: одна строка — одно значение, смысл значения задают оси (см. OUT_COLUMNS).

Выгрузка та же, что у «Динамики портфелей», и разбирается по тем же правилам
(см. reports/portfolio_dynamics/etl.py): лист ищется по колонке «Тип актива»,
строка «Позиция: <CODE>» открывает портфель, строки под ней — его бумаги.

Период выгрузки — С НАЧАЛА ГОДА по T-1 («Позиция за период 01.01.2026 -
29.09.2026»): колонки «(нач.)» в ней — на начало периода, «(кон.)» — на T-1.
Изменение Open QTY — это «(кон.)» минус «(нач.)», то есть изменение с начала
года, из самого файла. Выгрузка за один день (обе даты совпадают) для отчёта
не годится: изменение в ней всегда ноль, поэтому такие файлы не берутся.

Что берётся по портфелю:
- Open QTY на конец и на начало периода, Total Full PL with Funding, чистая
  стоимость — из строки «Позиция: …»: выгрузка считает их по портфелю сама.
  Если в строке пусто — сумма по бумагам;
- DV01 — ВСЕГДА сумма «DV01 (кон.)» по бумагам портфеля (строки под
  «Позиция: …» до следующей такой строки). DV01 бумаги в выгрузке посчитан на
  всю позицию, поэтому у портфеля он складывается, а не усредняется. Что
  написано в строке портфеля, не используется;
- Yield — ВСЕГДА средневзвешенная «Yield (кон.)» по бумагам портфеля, вес —
  «Чистая стоимость позиции (кон.)» бумаги. Строка портфеля не используется.

Два входа, которых нет в выгрузке, — как в «Динамике портфелей»:
- комментарии к портфелям пишутся в xlsx-витрине рядом с CSV и переносятся
  из выпуска в выпуск (reports/portfolio_report/workbook.py);
- RUONIA, RGBI и RWA вводятся руками и копятся в файле истории; в CSV
  попадают и даты из истории, которых в папке результатов ещё нет
  (reports/portfolio_report/market.py).
"""
import datetime as dt
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

BASE_DIR = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(BASE_DIR))

import config  # noqa: E402
from common import excel_io, rounding  # noqa: E402
from common.file_discovery import find_date_folders  # noqa: E402
from common.logging_utils import get_logger  # noqa: E402
from reports.portfolio_dynamics import etl as positions  # noqa: E402
from reports.portfolio_dynamics import inbox  # noqa: E402

logger = get_logger("portfolio_report")


class PortfolioReportError(RuntimeError):
    """Ошибка чтения или обработки входных файлов «Отчёта по портфелям»."""


# ── Колонки входной выгрузки ─────────────────────────────────────────────────
# Сравнение по excel_io.normalize_label, как и в «Динамике портфелей».
COL_QTY = "Open QTY (кон.)"
COL_QTY_START = "Open QTY (нач.)"
COL_PL = "Total Full PL with Funding"
COL_DV01 = "DV01 (кон.)"
COL_YIELD = "Yield (кон.)"
COL_VALUE = positions.COL_VALUE  # «Чистая стоимость позиции (кон.)»

INPUT_COLUMNS = (COL_QTY, COL_QTY_START, COL_PL, COL_DV01, COL_YIELD, COL_VALUE)

# Свод по портфелям внутри расчёта (суммы в рублях, как в выгрузке).
DATA_COLUMNS = [
    "portfolio_code", "portfolio_type", "open_qty", "open_qty_start", "open_qty_change",
    "net_value", "total_pl", "dv01", "yield",
]

# ── Плоский выход для BI ─────────────────────────────────────────────────────
# Раскладка как у ЧПД/NIM: одна строка — одно значение.
#   date_  — дата позиций (T-1), одна на весь файл: RGBI/RUONIA/RWA, введённые
#            в день формирования, относятся к этому же выпуску;
#   axis_0 — отчёт («Портфели»);
#   axis_1 — тип портфеля (AFS, HTM, TSS…) или «Рынок» для RGBI/RUONIA/RWA;
#   axis_2 — код портфеля (у рыночных показателей пусто);
#   axis_3 — показатель (METRICS ниже, либо RGBI/RUONIA/RWA);
#   axis_4 — единица измерения;
#   text_value — текст (комментарий к портфелю), у числовых строк пусто.
# date_ у рыночных показателей — своя у каждой строки: кроме даты выпуска, в
# файл попадают даты из истории, которых в CSV папки ещё нет.
# Пустые значения не пишутся: строка без значения в BI — только шум.
# Единицу и знаки каждого показателя можно поменять настройкой «Округление»
# (common/rounding.py, ключ показателя = колонка свода или ключ MarketInputs);
# axis_4 тогда меняется вместе со значением.
OUT_COLUMNS = ["id", "date_", "axis_0", "axis_1", "axis_2", "axis_3", "axis_4",
               "value", "text_value", "nversionid"]
AXIS_0 = "Портфели"
MARKET_GROUP = "Рынок"
# (колонка свода, название показателя в axis_3, единица в axis_4)
METRICS = [
    ("open_qty", "Open QTY", "шт"),
    ("open_qty_start", "Open QTY на начало года", "шт"),
    ("open_qty_change", "Изменение Open QTY", "шт"),
    ("net_value", "Чистая стоимость", "руб"),
    ("total_pl", "Total Full PL with Funding", "руб"),
    ("dv01", "DV01", "руб"),
    ("yield", "Yield", "%"),
]
MARKET_METRICS = [("rgbi", "RGBI", "пункты"), ("ruonia", "RUONIA", "%"), ("rwa", "RWA", "руб")]
COMMENT_METRIC = "Комментарий"

OUTPUT_FILENAME_PREFIX = "otchet_po_portfelyam_"
_OUTPUT_DATE = re.compile(re.escape(OUTPUT_FILENAME_PREFIX) + r"(\d{4}-\d{2}-\d{2})")


@dataclass
class MarketInputs:
    """Показатели на сегодня, которых нет в выгрузке: их вводят руками."""

    rgbi: Optional[float] = None
    ruonia: Optional[float] = None
    rwa: Optional[float] = None

    def missing(self) -> List[str]:
        return [name for name, value in (("RGBI", self.rgbi), ("RUONIA", self.ruonia),
                                         ("RWA", self.rwa)) if value is None]


@dataclass
class PositionsSnapshot:
    """Выгрузка на одну дату, свёрнутая до портфелей (суммы в рублях)."""

    path: Path
    business_date: Optional[dt.date]  # конец периода выгрузки (T-1)
    period_start: Optional[dt.date]   # начало периода (начало года)
    frame: pd.DataFrame  # portfolio_code, open_qty, open_qty_start, net_value, total_pl, dv01, yield
    securities: int = 0
    missing_columns: List[str] = field(default_factory=list)
    diagnostics: Optional["ParseDiagnostics"] = None


@dataclass
class ParseDiagnostics:
    """Что разбор увидел в файле — для --diagnose и предупреждений в логе.

    DV01 и Yield (в отличие от Open QTY, PL и стоимости) обычно считаются по
    строкам БУМАГ, поэтому пустые DV01/Yield почти всегда значат одно из двух:
    колонка не найдена по названию или строки бумаг не распознаны.
    """

    sheet_name: str
    header_row: int                                  # номер строки в Excel, с 1
    headers: List[str]                               # заголовки как в файле
    columns: Dict[str, int]                          # колонка выгрузки -> индекс
    # Строки бумаг с пустым «Тип актива» (номер строки Excel). Бумагами они
    # считаются: пустая ячейка приходит из Excel как NaN, а не как пустая строка.
    untyped_rows: List[int] = field(default_factory=list)
    # Непустые ячейки, которые не разобрались как число: колонка -> примеры.
    unparsed: Dict[str, List[str]] = field(default_factory=dict)
    # Портфель -> {"securities": N, "dv01": откуда, "yield": откуда}.
    portfolios: Dict[str, Dict[str, object]] = field(default_factory=dict)


@dataclass
class PortfolioReportData:
    business_date: dt.date          # дата позиций (T-1)
    period_start: dt.date           # с какой даты считается изменение Open QTY
    report_date: dt.date            # дата отчёта и введённых руками показателей
    source_path: Path
    frame: pd.DataFrame             # DATA_COLUMNS, суммы в рублях
    market: MarketInputs
    comments: Dict[str, str] = field(default_factory=dict)  # код (верхний регистр) -> текст
    # Комментарии портфелей, которых в этой выгрузке нет: в CSV их не пишут
    # (у портфеля нет чисел на дату), но витрина хранит их, чтобы портфель,
    # выпавший на день, не потерял комментарий — как в «Динамике портфелей».
    absent_comments: Dict[str, str] = field(default_factory=dict)
    # Рыночные показатели за ДРУГИЕ даты, которых ещё нет в CSV папки результатов
    # (см. market.pending): {дата: {rgbi/ruonia/rwa: значение}}.
    market_history: Dict[dt.date, Dict[str, float]] = field(default_factory=dict)


# ════════════════════════════════════════════════════════════════════════════
# Даты
# ════════════════════════════════════════════════════════════════════════════
def previous_business_day(day: dt.date) -> dt.date:
    """T-1: предыдущий рабочий день. Праздники не учитываются — только выходные."""
    day -= dt.timedelta(days=1)
    while day.weekday() >= 5:
        day -= dt.timedelta(days=1)
    return day


# ════════════════════════════════════════════════════════════════════════════
# Разбор выгрузки
# ════════════════════════════════════════════════════════════════════════════
def _sum(values) -> Optional[float]:
    """Сумма заполненных значений; None, если не заполнено ни одно."""
    present = [v for v in values if v is not None]
    return sum(present) if present else None


def parse_positions(path: Path) -> PositionsSnapshot:
    """Разбирает выгрузку «Позиция за период» и сворачивает её до портфелей."""
    path = Path(path)
    if not path.exists():
        raise PortfolioReportError(f"Файл не найден: {path}")

    sheet_name, matrix, header_row, _ = positions._find_sheet_and_header(path)
    header = [excel_io.normalize_label(v) for v in matrix.iloc[header_row]]
    type_col = header.index(excel_io.normalize_label(positions.COL_ASSET_TYPE))
    columns: Dict[str, int] = {}
    for label in INPUT_COLUMNS:
        norm = excel_io.normalize_label(label)
        if norm in header:
            columns[label] = header.index(norm)

    missing = [label for label in INPUT_COLUMNS if label not in columns]
    for label in missing:
        similar = _similar_headers(label, matrix.iloc[header_row])
        logger.warning("%s: колонка «%s» не найдена на листе %r — значения останутся пустыми.%s",
                       path.name, label, sheet_name,
                       f" Похожие заголовки в файле: {', '.join(repr(h) for h in similar)}."
                       if similar else "")
    if COL_VALUE not in columns:
        raise PortfolioReportError(
            f"В файле {path.name} (лист {sheet_name!r}) нет колонки «{COL_VALUE}» — "
            "без неё не посчитать ни стоимость, ни средневзвешенную доходность."
        )

    diag = ParseDiagnostics(
        sheet_name=sheet_name, header_row=header_row + 1,
        headers=["" if v is None or (not isinstance(v, str) and pd.isna(v)) else str(v).strip()
                 for v in matrix.iloc[header_row]],
        columns=dict(columns),
    )

    def cell(row, label) -> Optional[float]:
        if label not in columns:
            return None
        raw = row.iloc[columns[label]]
        value = positions.parse_number(raw)
        if value is None and raw is not None and not (not isinstance(raw, str) and pd.isna(raw)) \
                and str(raw).strip() and str(raw).strip().casefold() not in positions._MISSING_TOKENS:
            examples = diag.unparsed.setdefault(label, [])
            if len(examples) < 5:
                examples.append(str(raw))
        return value

    order: List[str] = []
    stated: Dict[str, Dict[str, Optional[float]]] = {}
    rows_by_code: Dict[str, List[Dict[str, Optional[float]]]] = {}
    current: Optional[str] = None
    securities = 0

    for row_idx in range(header_row + 1, len(matrix)):
        row = matrix.iloc[row_idx]
        raw_type = row.iloc[type_col]
        normalized = excel_io.normalize_label(raw_type)
        if not normalized:
            continue
        values = {label: cell(row, label) for label in INPUT_COLUMNS}

        if positions._is_position_marker(normalized):
            code = positions._extract_portfolio_code(raw_type)
            if not code:
                logger.warning("%s: строка %d — маркер портфеля без кода (%r), пропущена.",
                               path.name, row_idx + 1, raw_type)
                continue
            if code not in rows_by_code:
                rows_by_code[code] = []
                order.append(code)
            current = code
            stated[code] = values
            continue

        if normalized.startswith(positions.TOTAL_MARKERS) or current is None:
            continue  # «Итого» и шапка выгрузки до первого портфеля
        rows_by_code[current].append(values)
        securities += 1
        if raw_type is None or (not isinstance(raw_type, str) and pd.isna(raw_type)):
            diag.untyped_rows.append(row_idx + 1)

    if not order:
        raise PortfolioReportError(
            f"В файле {path.name} (лист {sheet_name!r}) нет ни одной строки "
            f"«{positions.COL_ASSET_TYPE}» вида «Позиция: <КОД>» — портфели определить не по чему."
        )

    for label, examples in diag.unparsed.items():
        logger.warning("%s: в колонке «%s» есть значения, которые не разобрались как число "
                       "(например: %s) — они считаются пустыми.", path.name, label,
                       "; ".join(repr(e) for e in examples))

    records = []
    for code in order:
        rows = rows_by_code[code]
        own = stated[code]

        def stated_or_sum(label):
            if own[label] is not None:
                return own[label]
            return _sum(r[label] for r in rows)

        computed_value = _sum(r[COL_VALUE] for r in rows)
        if (own[COL_VALUE] is not None and computed_value is not None
                and positions._relative_diff(computed_value, own[COL_VALUE])
                > config.PORTFOLIO_DYNAMICS_TOLERANCE):
            logger.warning(
                "%s: %s — чистая стоимость в строке «Позиция: …» = %s, сумма по бумагам = %s.",
                path.name, code, f"{own[COL_VALUE]:,.2f}", f"{computed_value:,.2f}",
            )

        yield_pairs = [(r[COL_YIELD], r[COL_VALUE]) for r in rows
                       if r[COL_YIELD] is not None and r[COL_VALUE] is not None]
        yield_ = positions._weighted_duration(yield_pairs)
        dv01 = _sum(r[COL_DV01] for r in rows)
        dv01_rows = sum(1 for r in rows if r[COL_DV01] is not None)
        diag.portfolios[code] = {
            "securities": len(rows),
            "dv01": f"сумма по {dv01_rows} бумагам" if dv01 is not None else "пусто",
            "yield": (f"средневзвешенная по {len(yield_pairs)} бумагам" if yield_ is not None
                      else "пусто"),
        }
        records.append({
            "portfolio_code": code,
            "open_qty": stated_or_sum(COL_QTY),
            "open_qty_start": stated_or_sum(COL_QTY_START),
            "net_value": stated_or_sum(COL_VALUE),
            "total_pl": stated_or_sum(COL_PL),
            "dv01": dv01,
            "yield": yield_,
        })

    for column, title in (("dv01", "DV01"), ("yield", "Yield")):
        if records and all(r[column] is None for r in records):
            logger.warning("%s: %s не посчитан ни по одному портфелю. Что отчёт увидел в файле, "
                           "покажет: python console.py portfolio-report --diagnose --input "
                           "\"<путь к файлу>\"", path.name, title)

    period = (positions._period_from_name(path)
              or positions._period_from_matrix(matrix, header_row))
    period_start, business_date = period if period else (None, None)
    logger.info("%s: лист %r, портфелей %d, бумаг %d, период %s",
                path.name, sheet_name, len(order), securities,
                f"{period_start.isoformat()} - {business_date.isoformat()}" if period
                else "не определён")
    return PositionsSnapshot(
        path=path, business_date=business_date, period_start=period_start,
        frame=pd.DataFrame(records, columns=["portfolio_code", "open_qty", "open_qty_start",
                                             "net_value", "total_pl", "dv01", "yield"]),
        securities=securities, missing_columns=missing, diagnostics=diag,
    )


_HINTS = {COL_DV01: ("dv01",), COL_YIELD: ("yield", "доход"), COL_QTY: ("qty", "колич"),
          COL_QTY_START: ("qty", "колич"), COL_PL: ("pl",), COL_VALUE: ("стоимост",)}


def _similar_headers(label: str, header_row) -> List[str]:
    """Заголовки файла, похожие на искомую колонку (для подсказки, что не нашлось)."""
    hints = _HINTS.get(label, ())
    result = []
    for value in header_row:
        if value is None or (not isinstance(value, str) and pd.isna(value)):
            continue
        if any(k in excel_io.normalize_label(value) for k in hints):
            result.append(str(value).strip())
    return result


def diagnose(path: Path) -> List[str]:
    """Отчёт для --diagnose: какие колонки нашлись, сколько бумаг увидено, откуда DV01/Yield."""
    snapshot = parse_positions(path)
    diag = snapshot.diagnostics
    from openpyxl.utils import get_column_letter

    lines = [f"Файл: {Path(path).name}",
             f"Лист: {diag.sheet_name!r}, строка заголовков: {diag.header_row}",
             "Период: " + (f"{snapshot.period_start:%d.%m.%Y} - {snapshot.business_date:%d.%m.%Y}"
                           if snapshot.business_date else "не определён"),
             "", "Колонки:"]
    for label in INPUT_COLUMNS:
        if label in diag.columns:
            index = diag.columns[label]
            lines.append(f"  [найдена]    {label} — колонка {get_column_letter(index + 1)}")
            continue
        similar = _similar_headers(label, diag.headers)
        lines.append(f"  [НЕ НАЙДЕНА] {label}"
                     + (f" — похожие заголовки в файле: {', '.join(repr(h) for h in similar)}"
                        if similar else " — похожих заголовков нет"))
    for label, examples in diag.unparsed.items():
        lines.append(f"  [не числа]   {label}: {'; '.join(repr(e) for e in examples)}")

    lines += ["", f"Портфелей: {len(diag.portfolios)}, строк бумаг: {snapshot.securities}"]
    if diag.untyped_rows:
        lines.append(f"  из них с пустым «{positions.COL_ASSET_TYPE}»: {len(diag.untyped_rows)} "
                     "(считаются бумагами)")
    lines.append("")
    width = max((len(code) for code in diag.portfolios), default=10)
    lines.append(f"  {'Портфель'.ljust(width)}  бумаг  DV01 — откуда / Yield — откуда")
    for code, info in diag.portfolios.items():
        lines.append(f"  {code.ljust(width)}  {info['securities']:>5}  "
                     f"{info['dv01']} / {info['yield']}")
    return lines


# ════════════════════════════════════════════════════════════════════════════
# Поиск выгрузки на дату
# ════════════════════════════════════════════════════════════════════════════
# Где искать выгрузку, в порядке убывания «своего»: своя папка исходных файлов
# (папки-даты и плоская раскладка), папка «Динамики портфелей» (та же
# выгрузка могла попасть туда её загрузчиком) и загрузки. Найденное не в своей
# папке кладётся туда копией (см. take) — так у отчёта копится свой архив.
ORIGIN_OWN_FOLDER = "папка своей даты"
ORIGIN_OTHER_FOLDER = "другая папка-дата"
ORIGIN_FLAT = "папка"
ORIGIN_DYNAMICS = "папка «Динамики»"
ORIGIN_DOWNLOADS = "загрузки"
_ORIGIN_PRIORITY = [ORIGIN_OWN_FOLDER, ORIGIN_OTHER_FOLDER, ORIGIN_FLAT, ORIGIN_DYNAMICS,
                    ORIGIN_DOWNLOADS]
_OWN_ORIGINS = (ORIGIN_OWN_FOLDER, ORIGIN_OTHER_FOLDER, ORIGIN_FLAT)


@dataclass(frozen=True)
class SourceFile:
    path: Path
    business_date: dt.date             # конец периода выгрузки
    origin: str
    period_start: Optional[dt.date] = None

    @property
    def from_year_start(self) -> bool:
        """Выгрузка «01.01 - дата» того же года — только такая годится отчёту."""
        return is_year_start_period(self.period_start, self.business_date)


# Выгрузка за один день не годится (изменение Open QTY в ней всегда ноль),
# выгрузка с другой начальной даты — тоже: изменение посчиталось бы не с
# начала года, и на дашборде это было бы не видно.
is_year_start_period = positions.is_year_start_period


def _period_text(start: Optional[dt.date], end: dt.date) -> str:
    return f"{start:%d.%m.%Y} - {end:%d.%m.%Y}" if start else f"? - {end:%d.%m.%Y}"


def _source():
    return config.PORTFOLIO_REPORT_SOURCE


def _dynamics_source():
    return config.PORTFOLIO_DYNAMICS_T0_SOURCE


def downloads_dir(no_import: bool = False) -> Optional[Path]:
    """Папка загрузок или None, если приёмка выключена настройкой или --no-import."""
    if no_import or not config.PORTFOLIO_DYNAMICS_IMPORT_FROM_DOWNLOADS:
        return None
    return Path(config.DOWNLOADS_DIR)


def find_sources(downloads: Optional[Path]) -> List[SourceFile]:
    """По одной выгрузке на каждую дату, свежие сначала.

    Если на дату есть несколько файлов, выгрузка с начала года важнее
    выгрузки за один день (та для отчёта не годится, см. pick_source). При
    равенстве берётся лежащий в своей папке (папка своей даты, другая
    папка-дата, плоская), затем в папке «Динамики», и только потом — в
    загрузках: разложенное по местам важнее того, что ещё не принято.
    """
    found: List[SourceFile] = []
    found += _scan_folder(_source())
    found += [SourceFile(item.path, item.business_date, ORIGIN_DYNAMICS)
              for item in _scan_folder(_dynamics_source())]
    if downloads is not None:
        found += [SourceFile(c.path, c.business_date, ORIGIN_DOWNLOADS)
                  for c in inbox.scan_slices(_source(), downloads)]

    def rank(item: SourceFile):
        return (not item.from_year_start, _ORIGIN_PRIORITY.index(item.origin))

    best: Dict[dt.date, SourceFile] = {}
    for item in found:
        period = positions.read_period(item.path)
        item = SourceFile(item.path, item.business_date, item.origin,
                          period[0] if period else None)
        current = best.get(item.business_date)
        if current is None or rank(item) < rank(current):
            best[item.business_date] = item
    return [best[d] for d in sorted(best, reverse=True)]


def _scan_folder(source) -> List[SourceFile]:
    """Выгрузки в папке источника: папки-даты и плоская раскладка."""
    found: List[SourceFile] = []
    for folder in find_date_folders(source):
        files = [f for f in folder.files if not positions.is_limits_file(f)]
        for candidate in inbox._dated(files, source):
            origin = (ORIGIN_OWN_FOLDER if candidate.business_date == folder.date
                      else ORIGIN_OTHER_FOLDER)
            found.append(SourceFile(candidate.path, candidate.business_date, origin))

    directory = Path(source.directory)
    if directory.is_dir() and source.filename_regex:
        pattern = re.compile(source.filename_regex)
        flat = [f for f in directory.glob(source.glob_pattern)
                if f.is_file() and not f.name.startswith("~$") and pattern.search(f.name)
                and not positions.is_limits_file(f)]
        found += [SourceFile(c.path, c.business_date, ORIGIN_FLAT)
                  for c in inbox._dated(flat, source)]
    return found


def take(item: SourceFile, move: Optional[bool] = None) -> Path:
    """Путь к выгрузке в своей папке; найденная в другом месте сначала кладётся туда.

    Из загрузок выгрузка раскладывается сразу по ОБЕИМ папкам — своей и
    «Динамики» (в «Динамику» — только если в её папке-дате ещё нет среза на ту
    же дату), см. inbox.file_position_export. Из папки «Динамики» — копией в
    свою: оттуда ничего не уносится. Перенос или копия из загрузок — по
    настройке «Переносить, а не копировать».
    """
    if item.origin in _OWN_ORIGINS:
        return item.path
    return inbox.file_position_export(item.path, item.business_date,
                                      from_downloads=item.origin == ORIGIN_DOWNLOADS, move=move)


def pick_source(sources: List[SourceFile], target: dt.date,
                strict: bool = True) -> SourceFile:
    """Выгрузка с начала года на дату target.

    strict=False — дата не задана явно (T-1 по умолчанию): если на неё файла
    нет (праздник, выгрузку не сделали), берётся самая свежая более ранняя, с
    предупреждением. Выгрузка за один день не берётся никогда: изменение Open
    QTY в ней всегда ноль, и отчёт молча показал бы, что ничего не менялось. Не
    берётся и выгрузка с другой начальной даты (не 01.01).
    """
    for item in sources:
        if item.business_date == target and item.from_year_start:
            return item
    unusable = [item for item in sources
                if item.business_date == target and not item.from_year_start]
    earlier = [item for item in sources if item.business_date < target and item.from_year_start]
    if not strict and earlier:
        logger.warning("Выгрузки с начала года на %s нет — взята самая свежая более ранняя, "
                       "на %s.", target.isoformat(), earlier[0].business_date.isoformat())
        return earlier[0]
    if unusable:
        item = unusable[0]
        raise PortfolioReportError(
            f"На {target.isoformat()} есть только выгрузка за период "
            f"{_period_text(item.period_start, item.business_date)} ({item.path.name}). "
            f"Отчёту нужна выгрузка с начала года: «Позиция за период 01.01.{target.year} - "
            f"{target:%d.%m.%Y}» — из неё считается изменение Open QTY."
        )
    available = ", ".join(item.business_date.isoformat() for item in sources[:10]
                          if item.from_year_start) or "ни одной"
    raise PortfolioReportError(
        f"Не найдена выгрузка «Позиция за период» с начала года на {target.isoformat()}. "
        f"Есть на даты: {available}. Искали в {Path(_source().directory)}, "
        f"{Path(_dynamics_source().directory)} (с папками-датами)"
        + (" и в загрузках." if downloads_dir() is not None else "; приёмка из загрузок выключена.")
    )


# ════════════════════════════════════════════════════════════════════════════
# Прошлые выпуски (их читает market.released)
# ════════════════════════════════════════════════════════════════════════════
def release_date(path: Path) -> Optional[dt.date]:
    match = _OUTPUT_DATE.search(Path(path).name)
    if not match:
        return None
    try:
        return dt.date.fromisoformat(match.group(1))
    except ValueError:
        return None


def _to_base_unit(value: float, key: str, unit: Optional[str], path: Path) -> float:
    """Значение прошлого выпуска -> исходная единица показателя (руб, %…).

    Выпуск мог быть записан с другой настройкой «Округление» (RWA в млрд):
    без пересчёта в бэкап истории ушло бы значение в чужой единице.
    """
    if unit is None:
        return value
    shift = rounding.scale_of_label("portfolio_report", key, unit)
    if shift is None:
        logger.warning("%s: у показателя %s неизвестная единица «%s» — значение взято как есть.",
                       Path(path).name, key, unit)
        return value
    return rounding.Rule(shift=shift).scale(value) if shift else value


# ════════════════════════════════════════════════════════════════════════════
# Сборка
# ════════════════════════════════════════════════════════════════════════════
def _type_of(code: str) -> str:
    return positions.guess_type(code, positions.KNOWN_PORTFOLIO_TYPES)


def build_data(source_path: Path,
               market: Optional[MarketInputs] = None,
               report_date: Optional[dt.date] = None,
               comments: Optional[Dict[str, str]] = None,
               market_history: Optional[Dict[dt.date, Dict[str, float]]] = None
               ) -> PortfolioReportData:
    snapshot = parse_positions(source_path)
    name = Path(source_path).name
    if snapshot.business_date is None:
        raise PortfolioReportError(
            f"Не удалось определить дату позиций в {name}: в имени файла или в шапке "
            "листа должна быть строка «Позиция за период [дд.мм.гггг] - [дд.мм.гггг]»."
        )
    start, end = snapshot.period_start, snapshot.business_date
    if not is_year_start_period(start, end):
        kind = "выгрузка за один день" if start == end else "выгрузка не с начала года"
        raise PortfolioReportError(
            f"{name} — {kind} (период {_period_text(start, end)}). Отчёту нужна выгрузка "
            f"с начала года: «Позиция за период 01.01.{end.year} - {end:%d.%m.%Y}» — "
            "изменение Open QTY считается от 01.01."
        )
    frame = snapshot.frame.copy()
    frame["portfolio_type"] = frame["portfolio_code"].map(_type_of)
    frame["open_qty_change"] = [
        None if qty is None or pd.isna(qty) or first is None or pd.isna(first) else qty - first
        for qty, first in zip(frame["open_qty"], frame["open_qty_start"])
    ]

    market = market or MarketInputs()
    for name in market.missing():
        logger.warning("%s не введён — в отчёте ячейка останется пустой.", name)

    comments = {str(code).strip().upper(): text for code, text in (comments or {}).items()
                if str(text or "").strip()}
    dropped = sorted(set(comments) - set(frame["portfolio_code"].str.upper()))
    if dropped:
        logger.warning("Портфелей %s нет в выгрузке — их комментарии сохранены в xlsx-витрине, "
                       "но в CSV не попадут.", ", ".join(dropped))
    history = {day: values for day, values in (market_history or {}).items()
               if day != snapshot.business_date and values}

    return PortfolioReportData(
        business_date=snapshot.business_date,
        report_date=report_date or dt.date.today(),
        source_path=Path(source_path),
        frame=frame[DATA_COLUMNS].reset_index(drop=True),
        market=market,
        period_start=start,
        comments={code: text for code, text in comments.items() if code not in dropped},
        absent_comments={code: comments[code] for code in dropped},
        market_history=history,
    )


def to_flat(data: PortfolioReportData) -> pd.DataFrame:
    """Свод -> плоская таблица для BI (раскладка — см. OUT_COLUMNS)."""
    date_ = data.business_date.isoformat()
    rows = []

    rules = {key: rounding.rule("portfolio_report", key)
             for key, _metric, _unit in METRICS + MARKET_METRICS}

    def add(group, code, key, metric, unit, value, day=date_):
        if value is None or pd.isna(value):
            return
        rule = rules[key]
        rows.append({"id": len(rows), "date_": day, "axis_0": AXIS_0, "axis_1": group,
                     "axis_2": code, "axis_3": metric, "axis_4": rule.label(unit),
                     "value": float(value) if rule.is_default else rule.apply(float(value)),
                     "text_value": "", "nversionid": ""})

    for record in data.frame.to_dict("records"):
        for column, metric, unit in METRICS:
            add(record["portfolio_type"], record["portfolio_code"], column, metric, unit,
                record[column])
        note = data.comments.get(str(record["portfolio_code"]).upper())
        if note:
            rows.append({"id": len(rows), "date_": date_, "axis_0": AXIS_0,
                         "axis_1": record["portfolio_type"], "axis_2": record["portfolio_code"],
                         "axis_3": COMMENT_METRIC, "axis_4": "", "value": None,
                         "text_value": note, "nversionid": ""})
    for day in sorted(data.market_history):
        for key, metric, unit in MARKET_METRICS:
            add(MARKET_GROUP, "", key, metric, unit, data.market_history[day].get(key),
                day=day.isoformat())
    for key, metric, unit in MARKET_METRICS:
        add(MARKET_GROUP, "", key, metric, unit, getattr(data.market, key))
    if all(rule.is_default for rule in rules.values()):
        return pd.DataFrame(rows, columns=OUT_COLUMNS)
    return rounding.frame_with_values(rows, OUT_COLUMNS)


def save_report(data: PortfolioReportData, output_path: Path) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    flat = to_flat(data)
    try:
        flat.to_csv(output_path, index=False, encoding="utf-8-sig")
    except PermissionError as exc:
        raise PortfolioReportError(
            f"Нет доступа для записи в {output_path} (файл открыт в Excel?): {exc}. "
            "Закройте файл и повторите запуск."
        ) from exc
    logger.info("Отчёт сохранён: %s (портфелей %d, строк %d, рынок из истории за %d дат)",
                output_path, len(data.frame), len(flat), len(data.market_history))
    return output_path


def default_output_path(business_date: dt.date) -> Path:
    return Path(config.PORTFOLIO_REPORT_OUTPUT_DIR) / (
        f"{OUTPUT_FILENAME_PREFIX}{business_date.isoformat()}.csv")
