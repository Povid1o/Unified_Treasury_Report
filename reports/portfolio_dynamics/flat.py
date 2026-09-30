# -*- coding: utf-8 -*-
"""Второй выход «Динамики портфелей»: плоский CSV для BI в длинном формате.

xlsx схемы v3.0 остаётся как был — этот файл пишется РЯДОМ с ним из тех же
данных и ничего в книге не меняет. Раскладка та же, что у «Отчёта по
портфелям» (reports/portfolio_report/etl.py): одна строка — одно значение,
смысл задают оси, плюс колонка text_value для текстовых значений.

    date_      — дата значения;
    axis_0     — отчёт («Динамика портфелей»);
    axis_1     — тип портфеля (HTM, HTM_KUAP, AFS, TSS);
    axis_2     — код портфеля; у показателей уровня типа пусто;
    axis_3     — показатель;
    axis_4     — единица измерения;
    value      — число (у текстовых показателей пусто);
    text_value — текст: наименование портфеля и комментарий.

Что попадает в ежедневный файл — всё с листа view_monitor на отчётную дату
и строки fact_type_daily ЗА ЭТУ ЖЕ дату. Историю BI накапливает сам,
дописывая файл за файлом; прошлые даты выгружаются один раз отдельным
файлом (history_to_flat), чтобы при ежедневной дозаписи ничего не задвоилось.
Для этого хватает одного готового xlsx: история объёмов по типам целиком
лежит на его листе fact_type_daily. Истории по отдельным портфелям нет нигде —
xlsx хранит портфели только на свою дату.

Значения берутся из среза напрямую, а не из view_monitor_values(): та
повторяет поведение Excel, где пустой объём T-7 становится нулём, и в CSV
это дало бы ложный прирост на весь объём T0. Пустых значений здесь нет
вовсе — строка без значения не пишется.

Лимит типа пишется одной строкой на тип, а не в каждой строке портфеля, как
на view_monitor: иначе сумма в BI умножила бы лимит на число портфелей.
Объём HTM в fact_type_daily уже включает HTM_KUAP — складывать «Объём типа»
по всем типам нельзя.
"""
import datetime as dt
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd

BASE_DIR = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(BASE_DIR))

from common.logging_utils import get_logger  # noqa: E402
from reports.portfolio_dynamics.etl import (  # noqa: E402
    MLN_IN_UNIT, OUTPUT_FILENAME_PREFIX, REPORT_UNIT, PortfolioDynamicsData,
    PortfolioDynamicsError,
)
from reports.portfolio_dynamics.workbook import (  # noqa: E402
    _cell_value, _portfolio_notes, order_for_views,
)

logger = get_logger("portfolio_dynamics")

OUT_COLUMNS = ["id", "date_", "axis_0", "axis_1", "axis_2", "axis_3", "axis_4",
               "value", "text_value", "nversionid"]
AXIS_0 = "Динамика портфелей"

# Показатели уровня портфеля (строки view_monitor).
M_NAME = "Наименование портфеля"
M_T0 = "Объём T0"
M_T7 = "Объём T-7"
M_DELTA = "Изменение объёма"
M_DELTA_PCT = "Изменение объёма, %"
M_DUR = "Дюрация текущая"
M_DUR_TARGET = "Дюрация-КУАП"
M_DUR_GAP = "Изменение дюрации"
M_NOTE = "Комментарий"
# Показатели уровня типа.
M_TYPE_VOLUME = "Объём типа"
M_TYPE_LIMIT = "Лимит типа"

U_AMOUNT = REPORT_UNIT
U_PCT = "%"
U_YEARS = "лет"

# Разность сумм в млрд даёт хвосты вида 0.5999999999999979 — в базу они
# уехали бы как есть. 9 знаков в млрд — точность до рубля.
_AMOUNT_DIGITS = 9
_PCT_DIGITS = 10

HISTORY_FILENAME_PREFIX = OUTPUT_FILENAME_PREFIX + "istoriya_"


def _to_unit(value_mln: Optional[float]) -> Optional[float]:
    """млн RUB (единица ETL) -> REPORT_UNIT, как в xlsx."""
    if value_mln is None:
        return None
    return round(value_mln / MLN_IN_UNIT[REPORT_UNIT], _AMOUNT_DIGITS)


def _number(value: Any) -> Optional[float]:
    value = _cell_value(value)
    if value is None or isinstance(value, (str, bool, dt.date)):
        return None
    return float(value)


def _date(value: Any) -> Optional[dt.date]:
    value = _cell_value(value)
    return value if isinstance(value, dt.date) else None


class _Rows:
    def __init__(self) -> None:
        self.rows: List[Dict[str, Any]] = []

    def add(self, day: dt.date, portfolio_type: Optional[str], code: Optional[str],
            metric: str, unit: str, value: Optional[float] = None,
            text: Optional[str] = None) -> None:
        if value is None and not text:
            return
        self.rows.append({
            "id": len(self.rows), "date_": day.isoformat(), "axis_0": AXIS_0,
            "axis_1": portfolio_type or "", "axis_2": code or "", "axis_3": metric,
            "axis_4": unit, "value": value, "text_value": text or "", "nversionid": "",
        })

    def frame(self) -> pd.DataFrame:
        return pd.DataFrame(self.rows, columns=OUT_COLUMNS)


def _type_volumes(history: pd.DataFrame, rows: _Rows, keep) -> None:
    """Строки fact_type_daily (суммы в млн RUB), для дат которых keep(date) истинно."""
    if history is None or history.empty:
        return
    records = []
    for record in history.to_dict("records"):
        day = _date(record["business_date"])
        if day is not None and keep(day):
            records.append((day, str(_cell_value(record["portfolio_type"]) or ""),
                            _number(record["volume_amount"])))
    for day, portfolio_type, volume in sorted(records, key=lambda r: (r[0], r[1])):
        rows.add(day, portfolio_type, None, M_TYPE_VOLUME, U_AMOUNT, _to_unit(volume))


def to_flat(data: PortfolioDynamicsData) -> pd.DataFrame:
    """Ежедневный файл: view_monitor + объёмы типов за отчётную дату."""
    day = data.business_date
    rows = _Rows()
    notes = _portfolio_notes(data)

    snapshot: Dict[str, Dict[str, Any]] = {}
    snap = data.fact_portfolio_snapshot
    if snap is not None and not snap.empty:
        for record in snap.to_dict("records"):
            code = _cell_value(record["portfolio_code"])
            if code is not None:
                snapshot.setdefault(str(code).upper(), record)

    dim = order_for_views(data.dim_portfolio)
    portfolios = [] if dim is None or dim.empty else dim.to_dict("records")
    listed = {str(_cell_value(p["portfolio_code"])).upper() for p in portfolios}
    # Портфель среза без строки в справочнике (CHK_11) не должен пропасть молча.
    portfolios += [{"portfolio_code": record["portfolio_code"], "portfolio_name": None,
                    "portfolio_type": None}
                   for key, record in snapshot.items() if key not in listed]

    for portfolio in portfolios:
        code = _cell_value(portfolio["portfolio_code"])
        if code is None:
            continue
        code = str(code)
        portfolio_type = _cell_value(portfolio["portfolio_type"])
        portfolio_type = str(portfolio_type) if portfolio_type is not None else None

        def add(metric, unit, value=None, text=None):
            rows.add(day, portfolio_type, code, metric, unit, value, text)

        name = _cell_value(portfolio.get("portfolio_name"))
        add(M_NAME, "", text=str(name) if name is not None else None)

        record = snapshot.get(code.upper())
        if record is not None:
            t0 = _number(record["volume_t0"])
            t7 = _number(record["volume_t7"])
            dur = _number(record["duration_current_yrs"])
            dur_target = _number(record["duration_target_yrs"])
            add(M_T0, U_AMOUNT, _to_unit(t0))
            add(M_T7, U_AMOUNT, _to_unit(t7))
            if t0 is not None and t7 is not None:
                add(M_DELTA, U_AMOUNT, _to_unit(t0 - t7))
                if t7 != 0:
                    add(M_DELTA_PCT, U_PCT, round((t0 / t7 - 1) * 100, _PCT_DIGITS))
            add(M_DUR, U_YEARS, dur)
            add(M_DUR_TARGET, U_YEARS, dur_target)
            if dur is not None and dur_target is not None:
                add(M_DUR_GAP, U_YEARS, dur - dur_target)

        add(M_NOTE, "", text=notes.get(code))

    limits = data.fact_limit
    if limits is not None and not limits.empty:
        for record in limits.to_dict("records"):
            portfolio_type = _cell_value(record["portfolio_type"])
            if portfolio_type is not None:
                rows.add(day, str(portfolio_type), None, M_TYPE_LIMIT, U_AMOUNT,
                         _to_unit(_number(record["limit_amount"])))

    _type_volumes(data.fact_type_daily, rows, lambda d: d == day)
    return rows.frame()


def history_to_flat(history: pd.DataFrame, until: Optional[dt.date] = None) -> pd.DataFrame:
    """Разовая выгрузка истории: объёмы типов за все даты ПО until включительно.

    history — fact_type_daily в млн RUB (как в ETL и в load_previous_release);
    until = None — вся история.
    """
    rows = _Rows()
    _type_volumes(history, rows, lambda d: until is None or d <= until)
    return rows.frame()


def _write(frame: pd.DataFrame, output_path: Path) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        frame.to_csv(output_path, index=False, encoding="utf-8-sig")
    except PermissionError as exc:
        raise PortfolioDynamicsError(
            f"Нет доступа для записи в {output_path} (файл открыт в Excel?): {exc}. "
            "Закройте файл и повторите запуск."
        ) from exc
    return output_path


def save_flat(data: PortfolioDynamicsData, output_path: Path) -> Path:
    frame = to_flat(data)
    _write(frame, output_path)
    logger.info("Плоский CSV сохранён: %s (строк %d)", output_path, len(frame))
    return Path(output_path)


def save_history(history: pd.DataFrame, until: Optional[dt.date],
                 output_dir: Path) -> Optional[Path]:
    """Пишет историю по until включительно. None — таких дат в истории нет."""
    frame = history_to_flat(history, until)
    if frame.empty:
        logger.warning("История объёмов по типам: дат%s нет — файл истории не записан.",
                       f" по {until.isoformat()} включительно" if until else "")
        return None
    first, last = frame["date_"].min(), frame["date_"].max()
    path = _write(frame, Path(output_dir) / f"{HISTORY_FILENAME_PREFIX}{first}_{last}.csv")
    logger.info("История объёмов по типам сохранена: %s (дат %d, строк %d)",
                path, frame["date_"].nunique(), len(frame))
    return path
