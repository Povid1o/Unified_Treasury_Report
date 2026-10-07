"""«Дополнительные портфели» в «Отчёте по портфелям» и история их P&L.

Список портфелей — общий с «Динамикой портфелей»: Excel-файл из настроек
(common/manual_portfolios.py). Из него берутся тип, балансовая стоимость
(«Чистая стоимость», в файле — млн RUB) и Open QTY. Open QTY не меняется,
поэтому на начало года он тот же, а изменение с начала года — ноль.

P&L с начала года меняется каждый день, и в выгрузке его нет — его вводят при
запуске. Каждый запуск дописывает его в файл истории (по умолчанию
manual_pl_history.csv рядом с settings.json): колонки date, code, pl_mln_rub.
Не введено — берётся последнее значение до даты позиций («как вчера»), и оно
тоже записывается на эту дату.

DV01 и Yield у дополнительных портфелей не заполняются — строк с ними в CSV
нет, и в средневзвешенный Yield по типу такие портфели не входят.
"""
import datetime as dt
import os
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd

BASE_DIR = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(BASE_DIR))

import config  # noqa: E402
from common import manual_portfolios as manual_store, settings  # noqa: E402
from common.logging_utils import get_logger  # noqa: E402
from reports.portfolio_report import etl  # noqa: E402

logger = get_logger("portfolio_report")

# {дата: {код портфеля: P&L с начала года, млн RUB}}
PlHistory = Dict[dt.date, Dict[str, float]]

FILE_COLUMNS = ["date", "code", "pl_mln_rub"]
_ITEM_SPLIT = re.compile(r"[;,]\s*(?=[^\s=,;]+\s*=)")


def records() -> List[dict]:
    """Дополнительные портфели из Excel-файла (с синхронизацией настроек)."""
    try:
        result = manual_store.sync()
    except settings.SettingsError as exc:
        raise etl.PortfolioReportError(str(exc)) from exc
    if result.created:
        logger.info("Создан файл дополнительных портфелей: %s", result.path)
    if result.updated:
        logger.info("Дополнительные портфели взяты из %s: %d шт.", result.path, len(result.records))
    return [dict(r, type=etl.positions.canonical_type(r["type"])) for r in result.records]


def history_path() -> Path:
    return Path(config.PORTFOLIO_REPORT_MANUAL_PL_HISTORY)


def read_history(path: Optional[Path] = None) -> PlHistory:
    path = Path(path or history_path())
    if not path.is_file():
        return {}
    try:
        table = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    except Exception as exc:  # битый файл не должен молча обнулять историю
        raise etl.PortfolioReportError(f"Не удалось прочитать историю P&L {path}: {exc}") from exc
    missing = [c for c in FILE_COLUMNS if c not in table.columns]
    if missing:
        raise etl.PortfolioReportError(
            f"{path.name}: нет колонок {', '.join(missing)}. Ожидаются {', '.join(FILE_COLUMNS)}.")
    history: PlHistory = {}
    for row in table.itertuples(index=False):
        try:
            day = dt.date.fromisoformat(str(row.date).strip())
        except ValueError:
            continue
        code = str(row.code).strip().upper()
        value = etl.positions.parse_number(row.pl_mln_rub)
        if code and value is not None:
            history.setdefault(day, {})[code] = value
    return history


def write_history(history: PlHistory, path: Optional[Path] = None) -> Path:
    path = Path(path or history_path())
    path.parent.mkdir(parents=True, exist_ok=True)
    rows = [[day.isoformat(), code, history[day][code]]
            for day in sorted(history) for code in sorted(history[day])]
    tmp = path.with_name(path.name + ".tmp")
    try:
        pd.DataFrame(rows, columns=FILE_COLUMNS).to_csv(tmp, index=False, encoding="utf-8-sig")
        os.replace(str(tmp), str(path))
    except PermissionError as exc:
        raise etl.PortfolioReportError(
            f"Нет доступа для записи в {path} (файл открыт в Excel?): {exc}. "
            "Закройте файл и повторите запуск.") from exc
    return path


def latest_before(history: PlHistory, day: dt.date) -> Dict[str, float]:
    """Последнее записанное до day значение P&L каждого портфеля."""
    result: Dict[str, float] = {}
    for when in sorted(d for d in history if d < day):
        result.update(history[when])
    return result


def apply_inputs(history: PlHistory, day: dt.date, codes: List[str],
                 entered: Dict[str, Optional[float]]) -> Dict[str, float]:
    """P&L выпуска по каждому коду; записывает его в историю на day.

    Введённое важнее записанного; не введено — то, что уже записано на day
    (повторный прогон), иначе последнее до day («как вчера»). Нет ничего —
    портфеля в результате нет, и строки P&L в CSV у него не будет.
    """
    previous = latest_before(history, day)
    own = history.setdefault(day, {})
    result: Dict[str, float] = {}
    for code in codes:
        value = entered.get(code)
        if value is not None:
            before = own.get(code)
            if before is not None and before != value:
                logger.warning("P&L %s на %s: было %s, записано %s (млн RUB).", code,
                               day.isoformat(), f"{before:,.4f}", f"{value:,.4f}")
            own[code] = value
        elif code not in own and code in previous:
            own[code] = previous[code]
            logger.info("P&L %s на %s не введён — взят прошлый: %s млн RUB.",
                        code, day.isoformat(), f"{previous[code]:,.4f}")
        if code in own:
            result[code] = own[code]
        else:
            logger.warning("P&L %s не введён и раньше не записывался — в отчёте его нет.", code)
    if not own:
        del history[day]
    return result


def parse_entered(raw: Optional[str]) -> Dict[str, float]:
    """«OFZ_EXTRA=12.5, X=3» (млн RUB) -> {код: значение} — для --manual-pl."""
    result: Dict[str, float] = {}
    # Делит только перед «КОД=»: запятая внутри числа (12,5) остаётся числу.
    for item in _ITEM_SPLIT.split(str(raw or "")):
        code, sep, value = item.partition("=")
        if not item.strip():
            continue
        number = etl.positions.parse_number(value) if sep else None
        if not code.strip() or number is None:
            raise ValueError(f"--manual-pl: не разобрано «{item.strip()}», ожидается КОД=число.")
        result[code.strip().upper()] = number
    return result
