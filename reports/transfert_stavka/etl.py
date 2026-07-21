"""ETL отчёта «Трансфертные ставки»: два Excel-файла (короткие + длинные
сроки) -> объединённый CSV в формате BI (date_, name_group, name_st,
scurrency, stype, unit, type_val, fvalue).

Бизнес-логика перенесена дословно из
EXPORT_FOLDER/TransfertStavka_Full.ipynb (ячейки "Короткие ставки" /
"Длинные ставки" / "Объединение") — вынесены только захардкоженные пути
и запуск через notebook-ячейки.
"""
import re
import sys
from pathlib import Path
from typing import Dict

import pandas as pd

BASE_DIR = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(BASE_DIR))

from common.logging_utils import get_logger  # noqa: E402

logger = get_logger("transfert_stavka", BASE_DIR / "logs")

NAME_GROUP = "ЕТС"
UNIT = "%"
TYPE_VAL = "Значение"
STYPE = "FIX"

OUT_COLUMNS = ["date_", "name_group", "name_st", "scurrency", "stype", "unit", "type_val", "fvalue"]

# ── Короткие ставки (до 3 месяцев) ───────────────────────────────────────────
SHORT_DATE_PATTERN = re.compile(r"\d{2}\.\d{2}\.\d{4}")
SHORT_MAPPING: Dict[int, str] = {1: "ON", 31: "1M", 90: "3M"}

# ── Длинные ставки (свыше 3 месяцев) ─────────────────────────────────────────
LONG_DATE_PATTERN = re.compile(r"\d{2} \d{2} \d{4}")
LONG_MAPPING: Dict[int, str] = {12: "1Y", 24: "2Y", 36: "3Y", 60: "5Y", 84: "7Y", 120: "10Y", 150: "15Y"}
LONG_MAPPING_FALLBACK: Dict[int, str] = {366: "1Y", 732: "2Y", 1098: "3Y", 1830: "5Y", 2562: "7Y", 3660: "10Y"}


class TransfertStavkaError(RuntimeError):
    """Ошибка чтения или обработки Excel-файлов трансфертных ставок."""


def _extract_date(file_path: Path, pattern: re.Pattern, date_format: str, label: str) -> str:
    match = pattern.search(str(file_path))
    if not match:
        raise TransfertStavkaError(f"[{label}] В имени файла {file_path.name} не найдена дата (шаблон {pattern.pattern!r})")
    return pd.to_datetime(match.group(), format=date_format).strftime("%Y-%m-%d")


# ── Короткие ставки ───────────────────────────────────────────────────────────
def _build_short_currency(
    file_path: Path, sheet_name: str, ncols: int, value_col: str, currency: str, date_str: str
) -> pd.DataFrame:
    try:
        raw_df = pd.read_excel(file_path, sheet_name=sheet_name, header=1)
    except Exception as exc:
        raise TransfertStavkaError(f"Не удалось прочитать лист {sheet_name!r} файла {file_path}: {exc}") from exc

    df = raw_df.iloc[:, 1:ncols]
    if sheet_name != "RUR":
        df = df[pd.to_numeric(df["от, дней"], errors="coerce").notna()]
    df = df.set_index("от, дней")

    filtered = df.loc[df["до, дней"].isin(SHORT_MAPPING)].copy()
    if value_col not in filtered.columns:
        raise TransfertStavkaError(
            f"На листе {sheet_name!r} не найдена колонка значения {value_col!r}. "
            f"Доступные колонки: {list(filtered.columns)}"
        )

    return pd.DataFrame({
        "date_": date_str,
        "name_group": NAME_GROUP,
        "name_st": filtered["до, дней"].map(SHORT_MAPPING),
        "scurrency": currency,
        "stype": STYPE,
        "unit": UNIT,
        "type_val": TYPE_VAL,
        "fvalue": filtered[value_col],
    }).reset_index(drop=True)


def build_short_report(file_path: Path) -> pd.DataFrame:
    """Короткие трансфертные ставки (ON, 1M, 3M) по RUB/USD/EUR/CNY."""
    file_path = Path(file_path)
    if not file_path.exists():
        raise TransfertStavkaError(f"Файл не найден: {file_path}")

    date_str = _extract_date(file_path, SHORT_DATE_PATTERN, "%d.%m.%Y", "ТС до 3М")

    parts = [
        _build_short_currency(file_path, "RUR", 8, "Трансфертная кривая размещения (КРЕДИТЫ), рубли ", "RUB", date_str),
        _build_short_currency(file_path, "USD", 4, "Трансфертная кривая размещения, USD", "USD", date_str),
        _build_short_currency(file_path, "EUR", 4, "Трансфертная кривая размещения, евро", "EUR", date_str),
        _build_short_currency(file_path, "CNY", 5, "Трансфертная кривая размещения, CNY", "CNY", date_str),
    ]
    combined = pd.concat(parts, ignore_index=True)
    result = combined[(combined["fvalue"] != 99) & (combined["fvalue"] != 0) & (combined["fvalue"] != "-")]
    logger.info("Короткие ставки (%s): %d строк", date_str, len(result))
    return result.reset_index(drop=True)


# ── Длинные ставки ────────────────────────────────────────────────────────────
def _build_long_currency(file_path: Path, sheet_name: str, currency: str, date_str: str) -> pd.DataFrame:
    try:
        raw_df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)
    except Exception as exc:
        raise TransfertStavkaError(f"Не удалось прочитать лист {sheet_name!r} файла {file_path}: {exc}") from exc

    df = raw_df[pd.to_numeric(raw_df["от, дней"], errors="coerce").notna()].copy()
    df["Месяц"] = pd.to_numeric(df["Месяц"], errors="coerce")

    filtered_main = df.loc[df["Месяц"].isin(LONG_MAPPING)].copy()
    filtered_fallback = df.loc[df["от, дней"].isin(LONG_MAPPING_FALLBACK)].copy()

    if sheet_name == "RUR":
        value_col = "Трансфертная кривая размещения (месяц), рубли "
    else:
        value_col = df.columns[3]

    main_df = pd.DataFrame({
        "date_": date_str, "name_group": NAME_GROUP,
        "name_st": filtered_main["Месяц"].map(LONG_MAPPING),
        "scurrency": currency, "stype": STYPE, "unit": UNIT, "type_val": TYPE_VAL,
        "fvalue": filtered_main[value_col],
    }).reset_index(drop=True)

    fallback_df = pd.DataFrame({
        "date_": date_str, "name_group": NAME_GROUP,
        "name_st": filtered_fallback["от, дней"].map(LONG_MAPPING_FALLBACK),
        "scurrency": currency, "stype": STYPE, "unit": UNIT, "type_val": TYPE_VAL,
        "fvalue": filtered_fallback[value_col],
    }).reset_index(drop=True)

    # Основной маппинг (по колонке "Месяц") имеет приоритет; резервный
    # (по "от, дней" в календарных днях) закрывает случаи, где "Месяц" не
    # заполнен. drop_duplicates(keep="first") оставляет основной результат.
    return pd.concat([main_df, fallback_df], ignore_index=True).drop_duplicates(subset=["name_st"], keep="first")


def build_long_report(file_path: Path) -> pd.DataFrame:
    """Длинные трансфертные ставки (1Y..15Y) по RUB/USD/EUR/CNY."""
    file_path = Path(file_path)
    if not file_path.exists():
        raise TransfertStavkaError(f"Файл не найден: {file_path}")

    date_str = _extract_date(file_path, LONG_DATE_PATTERN, "%d %m %Y", "ТС свыше 3М")

    parts = [
        _build_long_currency(file_path, "RUR", "RUB", date_str),
        _build_long_currency(file_path, "USD", "USD", date_str),
        _build_long_currency(file_path, "EUR", "EUR", date_str),
        _build_long_currency(file_path, "CNY", "CNY", date_str),
    ]
    combined = pd.concat(parts, ignore_index=True)
    result = combined[combined["fvalue"] != 99].reset_index(drop=True)
    logger.info("Длинные ставки (%s): %d строк", date_str, len(result))
    return result


# ── Объединение ────────────────────────────────────────────────────────────────
def build_report(short_path: Path, long_path: Path) -> pd.DataFrame:
    """Полный цикл: короткие + длинные ставки -> единый DataFrame в формате BI."""
    short_df = build_short_report(short_path)
    long_df = build_long_report(long_path)

    result = pd.concat([long_df, short_df], ignore_index=True)[OUT_COLUMNS]
    if result.empty:
        raise TransfertStavkaError("Итоговый DataFrame пуст — не найдено ни одной строки данных.")

    logger.info("Итоговый DataFrame: %d строк", len(result))
    return result


def save_report(df: pd.DataFrame, output_path: Path) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        df.to_csv(output_path, index=False, encoding="utf-8-sig")
    except PermissionError as exc:
        raise TransfertStavkaError(
            f"Нет доступа для записи в {output_path} (файл открыт в другой программе?): {exc}"
        ) from exc
    logger.info("Отчёт сохранён: %s (%d строк)", output_path, len(df))
    return output_path
