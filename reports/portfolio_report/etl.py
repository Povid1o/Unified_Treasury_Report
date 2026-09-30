"""ETL «Отчёта по портфелям»: выгрузка «Позиция за период» на T-1 -> свод по портфелям.

На выходе — плоский CSV для загрузки в BI, в той же раскладке, что у ЧПД и
NIM: одна строка — одно значение, смысл значения задают оси (см. OUT_COLUMNS).

Выгрузка та же, что у «Динамики портфелей», и разбирается по тем же правилам
(см. reports/portfolio_dynamics/etl.py): лист ищется по колонке «Тип актива»,
строка «Позиция: <CODE>» открывает портфель, строки под ней — его бумаги.

Что берётся по портфелю:
- Open QTY, Total Full PL with Funding, чистая стоимость — из строки
  «Позиция: …»: выгрузка считает их по портфелю сама. Если в строке пусто —
  сумма по бумагам;
- DV01 — сумма по бумагам: DV01 бумаги в выгрузке посчитан на всю позицию,
  поэтому у портфеля он складывается, а не усредняется;
- Yield — средневзвешенная по чистой стоимости бумаг: по портфелю выгрузка
  её не считает.

Почему предыдущий выпуск — это ВХОД. Open QTY сравнивается со значением на
предыдущую дату, а в выгрузке на T-1 его нет. Каждый выпуск содержит Open QTY
по всем портфелям, и следующий запуск берёт вчерашние значения оттуда.
"""
import datetime as dt
import re
import shutil
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

BASE_DIR = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(BASE_DIR))

import config  # noqa: E402
from common import excel_io  # noqa: E402
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
COL_PL = "Total Full PL with Funding"
COL_DV01 = "DV01 (кон.)"
COL_YIELD = "Yield (кон.)"
COL_VALUE = positions.COL_VALUE  # «Чистая стоимость позиции (кон.)»

INPUT_COLUMNS = (COL_QTY, COL_PL, COL_DV01, COL_YIELD, COL_VALUE)

# Свод по портфелям внутри расчёта (суммы в рублях, как в выгрузке).
DATA_COLUMNS = [
    "portfolio_code", "portfolio_type", "open_qty", "open_qty_prev", "open_qty_change",
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
#   axis_4 — единица измерения.
# Пустые значения не пишутся: строка без значения в BI — только шум.
OUT_COLUMNS = ["id", "date_", "axis_0", "axis_1", "axis_2", "axis_3", "axis_4",
               "value", "nversionid"]
AXIS_0 = "Портфели"
MARKET_GROUP = "Рынок"
# (колонка свода, название показателя в axis_3, единица в axis_4)
METRICS = [
    ("open_qty", "Open QTY", "шт"),
    ("open_qty_change", "Изменение Open QTY", "шт"),
    ("net_value", "Чистая стоимость", "руб"),
    ("total_pl", "Total Full PL with Funding", "руб"),
    ("dv01", "DV01", "руб"),
    ("yield", "Yield", "%"),
]
MARKET_METRICS = [("rgbi", "RGBI", "пункты"), ("ruonia", "RUONIA", "%"), ("rwa", "RWA", "руб")]
OPEN_QTY_METRIC = "Open QTY"

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
    business_date: Optional[dt.date]
    frame: pd.DataFrame  # portfolio_code, open_qty, net_value, total_pl, dv01, yield
    securities: int = 0
    missing_columns: List[str] = field(default_factory=list)


@dataclass
class PreviousRelease:
    path: Path
    business_date: Optional[dt.date]
    open_qty: Dict[str, float]  # портфель -> Open QTY
    market: MarketInputs


@dataclass
class PortfolioReportData:
    business_date: dt.date          # дата позиций (T-1)
    report_date: dt.date            # дата отчёта и введённых руками показателей
    source_path: Path
    frame: pd.DataFrame             # DATA_COLUMNS, суммы в рублях
    market: MarketInputs
    previous_date: Optional[dt.date] = None
    previous_path: Optional[Path] = None


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
        logger.warning("%s: колонка «%s» не найдена на листе %r — значения останутся пустыми.",
                       path.name, label, sheet_name)
    if COL_VALUE not in columns:
        raise PortfolioReportError(
            f"В файле {path.name} (лист {sheet_name!r}) нет колонки «{COL_VALUE}» — "
            "без неё не посчитать ни стоимость, ни средневзвешенную доходность."
        )

    def cell(row, label) -> Optional[float]:
        if label not in columns:
            return None
        return positions.parse_number(row.iloc[columns[label]])

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

    if not order:
        raise PortfolioReportError(
            f"В файле {path.name} (лист {sheet_name!r}) нет ни одной строки "
            f"«{positions.COL_ASSET_TYPE}» вида «Позиция: <КОД>» — портфели определить не по чему."
        )

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

        dv01 = _sum(r[COL_DV01] for r in rows)
        yield_ = positions._weighted_duration(
            [(r[COL_YIELD], r[COL_VALUE]) for r in rows
             if r[COL_YIELD] is not None and r[COL_VALUE] is not None]
        )
        records.append({
            "portfolio_code": code,
            "open_qty": stated_or_sum(COL_QTY),
            "net_value": stated_or_sum(COL_VALUE),
            "total_pl": stated_or_sum(COL_PL),
            # Бумаг с DV01/Yield нет — тогда то, что выгрузка написала по портфелю.
            "dv01": dv01 if dv01 is not None else own[COL_DV01],
            "yield": yield_ if yield_ is not None else own[COL_YIELD],
        })

    business_date = (positions._business_date_from_name(path)
                     or positions._business_date_from_matrix(matrix, header_row))
    logger.info("%s: лист %r, портфелей %d, бумаг %d, дата позиций %s",
                path.name, sheet_name, len(order), securities,
                business_date.isoformat() if business_date else "не определена")
    return PositionsSnapshot(
        path=path, business_date=business_date,
        frame=pd.DataFrame(records, columns=["portfolio_code", "open_qty", "net_value",
                                             "total_pl", "dv01", "yield"]),
        securities=securities, missing_columns=missing,
    )


# ════════════════════════════════════════════════════════════════════════════
# Поиск выгрузки на дату
# ════════════════════════════════════════════════════════════════════════════
# Выгрузки лежат там же, где у «Динамики портфелей»: папки-даты, плоская папка
# и загрузки. Настройки источника общие — отдельно их заводить незачем.
ORIGIN_OWN_FOLDER = "папка своей даты"
ORIGIN_OTHER_FOLDER = "другая папка-дата"
ORIGIN_FLAT = "папка"
ORIGIN_DOWNLOADS = "загрузки"
_ORIGIN_PRIORITY = [ORIGIN_OWN_FOLDER, ORIGIN_OTHER_FOLDER, ORIGIN_FLAT, ORIGIN_DOWNLOADS]


@dataclass(frozen=True)
class SourceFile:
    path: Path
    business_date: dt.date
    origin: str


def _source():
    return config.PORTFOLIO_DYNAMICS_T0_SOURCE


def downloads_dir(no_import: bool = False) -> Optional[Path]:
    """Папка загрузок или None, если приёмка выключена настройкой или --no-import."""
    if no_import or not config.PORTFOLIO_DYNAMICS_IMPORT_FROM_DOWNLOADS:
        return None
    return Path(config.DOWNLOADS_DIR)


def find_sources(downloads: Optional[Path]) -> List[SourceFile]:
    """По одной выгрузке на каждую дату, свежие сначала.

    Если на дату есть несколько файлов, берётся лежащий в папке своей даты,
    затем в другой папке-дате, затем в плоской папке, и только потом — в
    загрузках: разложенное по местам важнее того, что ещё не принято.
    """
    source = _source()
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

    if downloads is not None:
        found += [SourceFile(c.path, c.business_date, ORIGIN_DOWNLOADS)
                  for c in inbox.scan_slices(source, downloads)]

    best: Dict[dt.date, SourceFile] = {}
    for item in found:
        current = best.get(item.business_date)
        if current is None or (_ORIGIN_PRIORITY.index(item.origin)
                               < _ORIGIN_PRIORITY.index(current.origin)):
            best[item.business_date] = item
    return [best[d] for d in sorted(best, reverse=True)]


def take(item: SourceFile, move: Optional[bool] = None) -> Path:
    """Путь к выгрузке; из загрузок она сначала кладётся в папку своей даты.

    Туда же её положила бы «Динамика портфелей», так что файл потом пригодится
    и ей. Ничего не удаляется: перенос или копия — по той же настройке.
    """
    if item.origin != ORIGIN_DOWNLOADS:
        return item.path
    if move is None:
        move = config.PORTFOLIO_DYNAMICS_MOVE_FROM_DOWNLOADS
    folder = inbox.folder_for_date(_source(), item.business_date)
    destination = folder / item.path.name
    if destination.exists():
        logger.info("%s уже лежит в %s — взят оттуда.", destination.name, folder)
        return destination
    try:
        folder.mkdir(parents=True, exist_ok=True)
        if move:
            shutil.move(str(item.path), str(destination))
        else:
            shutil.copy2(str(item.path), str(destination))
    except (OSError, shutil.Error) as exc:
        logger.warning("Не удалось переложить %s из загрузок в %s: %s. Файл прочитан "
                       "прямо из загрузок.", item.path.name, folder, exc)
        return item.path
    logger.info("%s из загрузок: %s -> %s", "Перенесено" if move else "Скопировано",
                item.path.name, folder)
    return destination


def pick_source(sources: List[SourceFile], target: dt.date,
                strict: bool = True) -> SourceFile:
    """Выгрузка на дату target.

    strict=False — дата не задана явно (T-1 по умолчанию): если на неё файла
    нет (праздник, выгрузку не сделали), берётся самая свежая более ранняя, с
    предупреждением.
    """
    for item in sources:
        if item.business_date == target:
            return item
    earlier = [item for item in sources if item.business_date < target]
    if not strict and earlier:
        logger.warning("Выгрузки на %s нет — взята самая свежая более ранняя, на %s.",
                       target.isoformat(), earlier[0].business_date.isoformat())
        return earlier[0]
    available = ", ".join(item.business_date.isoformat() for item in sources[:10]) or "ни одной"
    raise PortfolioReportError(
        f"Не найдена выгрузка «Позиция за период» на {target.isoformat()}. "
        f"Есть на даты: {available}. Искали в {Path(_source().directory)} (с папками-датами)"
        + (" и в загрузках." if downloads_dir() is not None else "; приёмка из загрузок выключена.")
    )


# ════════════════════════════════════════════════════════════════════════════
# Предыдущий выпуск
# ════════════════════════════════════════════════════════════════════════════
def release_date(path: Path) -> Optional[dt.date]:
    match = _OUTPUT_DATE.search(Path(path).name)
    if not match:
        return None
    try:
        return dt.date.fromisoformat(match.group(1))
    except ValueError:
        return None


def find_previous_release(output_dir: Path, before: dt.date) -> Optional[Path]:
    """Самый свежий выпуск с датой позиций СТРОГО раньше before.

    Строго: повторный прогон за ту же дату должен сравниваться со вчерашним
    выпуском, а не с самим собой (иначе изменение Open QTY всегда было бы 0).
    """
    output_dir = Path(output_dir)
    if not output_dir.is_dir():
        return None
    dated = []
    for path in output_dir.glob(f"{OUTPUT_FILENAME_PREFIX}*.csv"):
        if not path.is_file():
            continue
        when = release_date(path)
        if when is not None and when < before:
            dated.append((when, path))
    if not dated:
        return None
    return max(dated, key=lambda item: item[0])[1]


def _optional_float(value) -> Optional[float]:
    return positions.parse_number(value)


def load_previous_release(path: Path) -> PreviousRelease:
    """Вчерашний Open QTY по портфелям и введённые тогда RGBI/RUONIA/RWA."""
    path = Path(path)
    try:
        flat = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    except (OSError, ValueError, UnicodeDecodeError) as exc:
        raise PortfolioReportError(
            f"Не удалось прочитать предыдущий выпуск {path.name}: {exc}. Укажите другой "
            "файл через --previous или удалите битый выпуск из папки результатов."
        ) from exc
    missing = [c for c in ("date_", "axis_1", "axis_2", "axis_3", "value") if c not in flat.columns]
    if missing:
        raise PortfolioReportError(
            f"Предыдущий выпуск {path.name} не в формате отчёта — нет колонок {missing}."
        )

    open_qty: Dict[str, float] = {}
    market: Dict[str, Optional[float]] = {}
    market_names = {title: key for key, title, _unit in MARKET_METRICS}
    for row in flat.itertuples(index=False):
        value = _optional_float(row.value)
        if value is None:
            continue
        if row.axis_3 == OPEN_QTY_METRIC and row.axis_2:
            open_qty[row.axis_2.strip().upper()] = value
        elif row.axis_1 == MARKET_GROUP and row.axis_3 in market_names:
            market[market_names[row.axis_3]] = value

    business_date = release_date(path)
    dates = {d for d in flat["date_"] if d}
    if len(dates) == 1:
        try:
            business_date = dt.date.fromisoformat(dates.pop())
        except ValueError:
            pass
    return PreviousRelease(path=path, business_date=business_date, open_qty=open_qty,
                           market=MarketInputs(**market))


# ════════════════════════════════════════════════════════════════════════════
# Сборка
# ════════════════════════════════════════════════════════════════════════════
def _type_of(code: str) -> str:
    return positions.guess_type(code, positions.KNOWN_PORTFOLIO_TYPES)


def build_data(source_path: Path, previous_path: Optional[Path] = None,
               market: Optional[MarketInputs] = None,
               report_date: Optional[dt.date] = None) -> PortfolioReportData:
    snapshot = parse_positions(source_path)
    if snapshot.business_date is None:
        raise PortfolioReportError(
            f"Не удалось определить дату позиций в {Path(source_path).name}: в имени файла "
            "или в шапке листа должна быть строка «Позиция за период [дд.мм.гггг] - [дд.мм.гггг]»."
        )
    frame = snapshot.frame.copy()
    frame["portfolio_type"] = frame["portfolio_code"].map(_type_of)

    previous = load_previous_release(previous_path) if previous_path else None
    if previous is not None and previous.business_date is not None \
            and previous.business_date >= snapshot.business_date:
        logger.warning("Предыдущий выпуск %s — на %s, не раньше позиций (%s): сравнение "
                       "Open QTY покажет не то, что нужно.", previous.path.name,
                       previous.business_date.isoformat(), snapshot.business_date.isoformat())

    if previous is None:
        logger.warning("Предыдущего выпуска нет — Open QTY сравнивать не с чем, колонки "
                       "«пред.» и «изменение» останутся пустыми.")
        frame["open_qty_prev"] = None
        frame["open_qty_change"] = None
    else:
        prev_qty = pd.Series(previous.open_qty, dtype=float)
        frame["open_qty_prev"] = frame["portfolio_code"].map(prev_qty.to_dict())
        # Портфеля не было вчера — количество выросло с нуля; нет сегодня — упало до нуля.
        gone = [code for code in prev_qty.index if code not in set(frame["portfolio_code"])]
        if gone:
            logger.warning("Портфели из предыдущего выпуска, которых нет в выгрузке: %s — "
                           "показаны с нулевым Open QTY.", ", ".join(gone))
            extra = pd.DataFrame({
                "portfolio_code": gone,
                "portfolio_type": [_type_of(code) for code in gone],
                "open_qty": [0.0] * len(gone),
                "open_qty_prev": [prev_qty[code] for code in gone],
            })
            frame = pd.concat([frame, extra], ignore_index=True)
        new = frame.loc[frame["open_qty_prev"].isna(), "portfolio_code"].tolist()
        if new:
            logger.info("Новые портфели (в предыдущем выпуске их не было): %s", ", ".join(new))
        frame["open_qty_change"] = [
            None if qty is None or pd.isna(qty) else
            qty - (0.0 if prev is None or pd.isna(prev) else prev)
            for qty, prev in zip(frame["open_qty"], frame["open_qty_prev"])
        ]

    market = market or MarketInputs()
    for name in market.missing():
        logger.warning("%s не введён — в отчёте ячейка останется пустой.", name)

    return PortfolioReportData(
        business_date=snapshot.business_date,
        report_date=report_date or dt.date.today(),
        source_path=Path(source_path),
        frame=frame[DATA_COLUMNS].reset_index(drop=True),
        market=market,
        previous_date=previous.business_date if previous else None,
        previous_path=previous.path if previous else None,
    )


def to_flat(data: PortfolioReportData) -> pd.DataFrame:
    """Свод -> плоская таблица для BI (раскладка — см. OUT_COLUMNS)."""
    date_ = data.business_date.isoformat()
    rows = []

    def add(group, code, metric, unit, value):
        if value is None or pd.isna(value):
            return
        rows.append({"id": len(rows), "date_": date_, "axis_0": AXIS_0, "axis_1": group,
                     "axis_2": code, "axis_3": metric, "axis_4": unit,
                     "value": float(value), "nversionid": ""})

    for record in data.frame.to_dict("records"):
        for column, metric, unit in METRICS:
            add(record["portfolio_type"], record["portfolio_code"], metric, unit, record[column])
    for key, metric, unit in MARKET_METRICS:
        add(MARKET_GROUP, "", metric, unit, getattr(data.market, key))
    return pd.DataFrame(rows, columns=OUT_COLUMNS)


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
    logger.info("Отчёт сохранён: %s (портфелей %d, строк %d)",
                output_path, len(data.frame), len(flat))
    return output_path


def default_output_path(business_date: dt.date) -> Path:
    return Path(config.PORTFOLIO_REPORT_OUTPUT_DIR) / (
        f"{OUTPUT_FILENAME_PREFIX}{business_date.isoformat()}.csv")
