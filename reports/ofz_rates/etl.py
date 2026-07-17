"""ETL отчёта «Ставки ОФЗ»: CBonds API -> DataFrame в формате BI."""
import re
import sys
from datetime import date, timedelta
from pathlib import Path
from typing import List, Optional

import pandas as pd
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(BASE_DIR))

# .env лежит внутри перенесённой библиотеки (CBonds_API/.env). Загружаем его явно,
# не полагаясь на автопоиск python-dotenv внутри cbonds_api_test.py, чтобы
# скрипт работал независимо от того, откуда его запускают.
load_dotenv(BASE_DIR / "CBonds_API" / ".env")

from CBonds_API.cbonds_api_test import AVAILABLE_INDICES, CBondsAPI  # noqa: E402
from common.logging_utils import get_logger  # noqa: E402

logger = get_logger("ofz_rates", BASE_DIR / "logs")

# ── Конфигурация отчёта ──────────────────────────────────────────────────────
OUTPUT_PATH = BASE_DIR / "output" / "ofz_rates" / "ofz_report.csv"

# Глубина архива для get_index_value_new у демо-подписки CBonds ограничена
# 100 календарными днями (см. CBonds_API/README.md, раздел "Ограничения").
# Берём немного меньше, с запасом под задержку публикации данных.
LOOKBACK_DAYS = 90

# Сроки кривой доходности ОФЗ, которые нужно выгрузить для группы
# "Доходность ОФЗ" (пример из ТЗ по структуре BI).
OFZ_YIELD_TENORS: List[str] = ["1Y", "2Y", "3Y", "4Y", "5Y", "6Y", "9Y", "10Y"]

# Итоговые колонки итогового DataFrame — строго по спецификации BI.
BI_COLUMNS = ["date_", "name_group", "name_st", "unit", "type_val", "fvalue"]
DATE_PATTERN = re.compile(r"^\d{4}-\d{2}-\d{2}$")

# Показатели, которые не удалось получить через доступные эндпоинты CBonds API
# (см. итоговый отчёт: проверенные эндпоинты и причина отсутствия данных).
KNOWN_GAPS = [
    {
        "name_group": "Размещение ОФЗ",
        "name_st": "Спрос",
        "reason": (
            "В доступных эндпоинтах аккаунта (get_emissions, get_flow_new, get_offert, "
            "get_emission_default, get_emission_guarantors, get_floater_coupons, "
            "get_tradings_new, get_index_types, get_index_value_new) нет поля со спросом "
            "на аукционе ОФЗ. get_emissions (210 полей) содержит только announced_volume_new "
            "и placed_volume_new — поля 'спрос'/'demand' отсутствуют."
        ),
    },
    {
        "name_group": "Размещение ОФЗ",
        "name_st": "Разместили",
        "reason": (
            "get_emissions.placed_volume_new — это накопленный объём выпуска на текущий "
            "момент (по конкретному ISIN), а не результат конкретного еженедельного "
            "аукциона на дату. Привязка к дате первичного размещения (date_of_start_placing) "
            "даёт неверные значения для до-размещаемых траншей (доразмещения на последующих "
            "аукционах не отражаются отдельной датой). Достоверного эндпоинта с результатами "
            "аукционов по датам в доступе аккаунта нет."
        ),
    },
    {
        "name_group": "Операции",
        "name_st": "Объём торгов",
        "reason": (
            "Агрегированного показателя оборота рынка ОФЗ нет ни среди 39 индексов, "
            "доступных через get_index_types/get_index_value_new для этого аккаунта, ни в "
            "виде отдельного эндпоинта. get_tradings_new отдаёт оборот (поле overturn) только "
            "по одной облигации за раз (по ISIN); построение рыночного агрегата потребовало бы "
            "перебора всех выпусков ОФЗ и не опирается на существующий метод библиотеки."
        ),
    },
]


class OfzDataError(RuntimeError):
    """Ошибка получения или валидации данных отчёта «Ставки ОФЗ»."""


# ── Получение данных ─────────────────────────────────────────────────────────
def build_api_client() -> CBondsAPI:
    """Создаёт клиент CBondsAPI, используя переменные окружения из .env."""
    try:
        return CBondsAPI()
    except ValueError as exc:
        raise OfzDataError(
            f"Не удалось инициализировать CBondsAPI: {exc}. "
            f"Проверьте файл {BASE_DIR / 'CBonds_API' / '.env'}."
        ) from exc


def fetch_yield_curve_group(
    api: CBondsAPI, date_from: str, date_to: str
) -> pd.DataFrame:
    """Группа "Доходность ОФЗ": кривая доходности по каждому сроку из OFZ_YIELD_TENORS.

    Используется существующий метод CBondsAPI.get_index_values (эндпоинт
    get_index_value_new), т.к. он поддерживает диапазон дат — в отличие от
    более узкого get_yield_curve, у которого нет параметров date_from/date_to.
    """
    rows = []
    for tenor in OFZ_YIELD_TENORS:
        index_key = f"RUB_Yield_Curve_{tenor}"

        if index_key not in AVAILABLE_INDICES:
            logger.warning(
                "Пропускаю срок %s: индекс %s отсутствует в AVAILABLE_INDICES "
                "и в списке индексов, доступных аккаунту (get_index_types). "
                "Проверенный эндпоинт: get_index_value_new.",
                tenor,
                index_key,
            )
            continue

        try:
            items = api.get_index_values(
                index_key, date_from=date_from, date_to=date_to, limit=200
            )
        except RuntimeError as exc:
            logger.error("Ошибка запроса кривой доходности ОФЗ (%s): %s", tenor, exc)
            continue

        if not items:
            logger.warning("Нет данных по сроку %s за период %s..%s", tenor, date_from, date_to)
            continue

        for item in items:
            rows.append(
                {
                    "date_": item["date"],
                    "name_group": "Доходность ОФЗ",
                    "name_st": tenor,
                    "unit": "%",
                    "type_val": "Значение",
                    "fvalue": round(float(item["value"]), 2),
                }
            )

        logger.info("Доходность ОФЗ %s: получено %d значений", tenor, len(items))

    return pd.DataFrame(rows, columns=BI_COLUMNS)


def log_known_gaps() -> None:
    """Подробно логирует показатели, которые не удалось получить (см. KNOWN_GAPS)."""
    for gap in KNOWN_GAPS:
        logger.warning(
            "Показатель не получен: группа='%s', показатель='%s'. Причина: %s",
            gap["name_group"],
            gap["name_st"],
            gap["reason"],
        )


# ── Преобразование и валидация ───────────────────────────────────────────────
def validate_bi_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Проверяет и приводит DataFrame к спецификации BI.

    Проверяется: набор и порядок колонок, отсутствие пропусков, формат дат
    YYYY-MM-DD, единственное допустимое значение type_val, числовой тип fvalue.
    Обнаруженные несоответствия исправляются автоматически там, где это
    возможно без потери данных; иначе выбрасывается OfzDataError.
    """
    if list(df.columns) != BI_COLUMNS:
        missing = set(BI_COLUMNS) - set(df.columns)
        extra = set(df.columns) - set(BI_COLUMNS)
        if missing:
            raise OfzDataError(f"В DataFrame отсутствуют колонки: {missing}")
        if extra:
            logger.warning("Удаляю лишние колонки, не входящие в спецификацию BI: %s", extra)
        df = df[BI_COLUMNS]

    if df.isna().any().any():
        na_counts = df.isna().sum()
        raise OfzDataError(f"В DataFrame есть пропуски:\n{na_counts[na_counts > 0]}")

    for col in ["date_", "name_group", "name_st", "unit", "type_val"]:
        df[col] = df[col].astype(str)

    bad_dates = ~df["date_"].str.match(DATE_PATTERN)
    if bad_dates.any():
        raise OfzDataError(
            f"Найдены даты не в формате YYYY-MM-DD: {df.loc[bad_dates, 'date_'].unique().tolist()}"
        )

    invalid_type_val = df["type_val"] != "Значение"
    if invalid_type_val.any():
        raise OfzDataError(
            f"Найдены недопустимые значения type_val: {df.loc[invalid_type_val, 'type_val'].unique().tolist()}"
        )

    df["fvalue"] = pd.to_numeric(df["fvalue"], errors="coerce")
    if df["fvalue"].isna().any():
        raise OfzDataError("Колонка fvalue содержит значения, которые не удалось привести к float")
    df["fvalue"] = df["fvalue"].astype(float)

    before = len(df)
    df = df.drop_duplicates(subset=["date_", "name_group", "name_st"], keep="last")
    if len(df) != before:
        logger.warning("Удалено %d дублирующихся строк (date_, name_group, name_st)", before - len(df))

    df = df.sort_values(["name_group", "name_st", "date_"]).reset_index(drop=True)
    return df


# ── Оркестрация ETL ──────────────────────────────────────────────────────────
def build_report(
    as_of_date: Optional[date] = None,
    lookback_days: int = LOOKBACK_DAYS,
) -> pd.DataFrame:
    """Полный цикл: подключение к API -> получение данных -> формат BI.

    :param as_of_date: дата, на которую формируется отчёт ("любая дата" из
        консоли запуска). По умолчанию — сегодня.
    :param lookback_days: глубина истории в календарных днях назад от as_of_date.
    """
    as_of_date = as_of_date or date.today()
    date_from = (as_of_date - timedelta(days=lookback_days)).isoformat()
    date_to = as_of_date.isoformat()

    logger.info("Формирование отчёта «Ставки ОФЗ» за период %s..%s", date_from, date_to)

    api = build_api_client()

    df = fetch_yield_curve_group(api, date_from, date_to)
    log_known_gaps()

    if df.empty:
        raise OfzDataError("Итоговый DataFrame пуст — ни одного показателя не удалось получить")

    df = validate_bi_dataframe(df)
    logger.info("Итоговый DataFrame: %d строк, %d показателей", len(df), df["name_st"].nunique())
    return df


def save_report(df: pd.DataFrame, output_path: Path = OUTPUT_PATH) -> Path:
    """Сохраняет готовый к загрузке в BI DataFrame в CSV."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(output_path, index=False, encoding="utf-8-sig")
    logger.info("Отчёт сохранён: %s", output_path)
    return output_path
