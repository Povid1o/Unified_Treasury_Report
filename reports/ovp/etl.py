"""ETL отчёта «ОВП»: преобразование Excel-отчёта ОВП (форма 634) в нормализованный CSV.

Бизнес-логика перенесена из OVPCombine/app_gui.py (Tkinter GUI) без изменений
правил разбора иерархии и дат — вынесена только обработка данных, без UI.
"""
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

BASE_DIR = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(BASE_DIR))

from common.logging_utils import get_logger  # noqa: E402

logger = get_logger("ovp", BASE_DIR / "logs")

# Словарь иерархии категорий/подкатегорий отчёта ОВП (форма 634 и управленческие формы).
HIERARCHY: Dict[str, List[str]] = {
    "АКТИВЫ(1)": [
        "Денежные средства и их эквиваленты(2)", "ФОР(1214)", "МБК свыше 90 дней(974)",
        "Наличные средства(3)", "Портфель ценных бумаг(10413)", "РЕПО до 30 дней(974)",
        "Денежные требования по операциям обратного РЕПО(969)", "Ссудная и приравненная к ней задолженность(77)",
        "Ценные бумаги, имеющиеся в наличии для продажи(1216)", "Активы, предназначенные для продажи(2009)",
        "Другие активы(1235)", "неразобранные АКТИВЫ(371)",
    ],
    "ПАССИВЫ(220)": [
        "Депозиты и прочие привлеченные ресурсы финансовых организаций(221)", "Привлечение с финасовых рынков(10464)",
        "Клиентское привлечение и РЕПО(10410)", "Межбанковское привлечение(221)", "Выпущенные ценные бумаги(1364)",
        "Денежные обязательства по операциям прямого РЕПО (967)", "Обязательства по сделкам обратного РЕПО и займа ценных бумаг(1211)",
        "Прочие пассивы(311)", "неразобранные Пассивы(372)",
    ],
    "ВНЕБАЛАНС(431)": [
        "Срочные сделки", "Сделки Спот", "Резервы по внебалансовым обязательствам(437)",
    ],
    "ОВП (634 форма)": [
        "ОВП БФР", "Экономическая ОВП (с учетом резервов по МСФО)",
    ],
}

OUT_COLUMNS = [
    "id", "ord", "category_name", "subcategory",
    "curr_balance", "reserve_msfo", "currency", "report_date",
]

DATE_STR_PATTERN = re.compile(r"^\d{2}\.\d{2}\.\d{4}")


class OvpDataError(RuntimeError):
    """Ошибка чтения или обработки Excel-отчёта ОВП."""


def clean_name(name) -> str:
    """Убирает маркер иерархии "•" и лишние пробелы из названия строки."""
    if not isinstance(name, str):
        return ""
    return name.replace("•", "").strip()


def _clean_num(value) -> int:
    """Приводит значение баланса/резерва к int, подставляя 0 для NaN/мусора."""
    if pd.isna(value) or value == "":
        return 0
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return 0


def list_currency_sheets(input_path: Path) -> List[str]:
    """Список валютных листов файла (сводные листы с "СВОД" в имени исключены)."""
    input_path = Path(input_path)
    if not input_path.exists():
        raise OvpDataError(f"Файл не найден: {input_path}")
    try:
        xls = pd.ExcelFile(input_path)
    except Exception as exc:
        raise OvpDataError(f"Не удалось прочитать Excel-файл {input_path}: {exc}") from exc
    return [s for s in xls.sheet_names if "СВОД" not in s.upper()]


def _find_date_row(df_raw: pd.DataFrame) -> Tuple[Optional[int], List[Tuple[str, int]]]:
    """Ищет в первых 20 строках листа строку с заголовками дат.

    Возвращает (индекс_строки, [(дата в формате YYYY-MM-DD, индекс_колонки), ...]).
    Дата может быть строкой "DD.MM.YYYY" или datetime-объектом, который Excel
    распознал сам.
    """
    for idx, row in df_raw.head(20).iterrows():
        row_dates = []
        for col_idx, val in enumerate(row):
            if isinstance(val, str) and DATE_STR_PATTERN.match(val):
                try:
                    dt = pd.to_datetime(val, dayfirst=True)
                    row_dates.append((dt.strftime("%Y-%m-%d"), col_idx))
                except (ValueError, TypeError):
                    pass
            elif hasattr(val, "strftime"):
                row_dates.append((val.strftime("%Y-%m-%d"), col_idx))
        if row_dates:
            return idx, row_dates
    return None, []


def convert_ovp_report(
    input_path: Path,
    full_history_currencies: Optional[List[str]] = None,
    currencies: Optional[List[str]] = None,
) -> pd.DataFrame:
    """Преобразует Excel-отчёт ОВП в плоский DataFrame.

    :param input_path: путь к исходному .xlsx.
    :param full_history_currencies: валютные листы, для которых нужно выгрузить
        все найденные даты. Остальные листы выгружаются только за последнюю дату.
    :param currencies: ограничить обработку конкретными листами (по умолчанию —
        все листы, кроме содержащих "СВОД" в названии).
    """
    input_path = Path(input_path)
    available = list_currency_sheets(input_path)

    target_currencies = currencies if currencies is not None else available
    unknown = set(target_currencies) - set(available)
    if unknown:
        raise OvpDataError(f"Листы не найдены в файле: {sorted(unknown)}. Доступны: {available}")

    full_history_set = set(full_history_currencies or [])
    xls = pd.ExcelFile(input_path)

    all_rows = []
    global_id = 1

    for currency in target_currencies:
        try:
            df_raw = pd.read_excel(xls, sheet_name=currency, header=None)
        except Exception as exc:
            logger.warning("Лист '%s' пропущен: не удалось прочитать (%s)", currency, exc)
            continue

        date_row_idx, dates_found = _find_date_row(df_raw)
        if date_row_idx is None:
            logger.warning("Лист '%s' пропущен: не найдена строка с датами", currency)
            continue

        dates_target = dates_found if currency in full_history_set else [dates_found[-1]]
        data_start_row = date_row_idx + 2

        for r_date, col_idx in dates_target:
            ord_counter = 1
            current_category = None
            col_bal, col_res = col_idx, col_idx + 1
            ncols = len(df_raw.columns)

            for r_idx in range(data_start_row, len(df_raw)):
                row_data = df_raw.iloc[r_idx]
                clean = clean_name(row_data[0])
                if not clean:
                    continue

                if clean in HIERARCHY:
                    current_category = clean
                    target_subcategory = ""
                elif current_category and clean in HIERARCHY[current_category]:
                    target_subcategory = clean
                else:
                    continue

                val_bal = row_data[col_bal] if col_bal < ncols else 0
                val_res = row_data[col_res] if col_res < ncols else 0

                all_rows.append({
                    "id": global_id,
                    "ord": ord_counter,
                    "category_name": current_category,
                    "subcategory": target_subcategory,
                    "curr_balance": _clean_num(val_bal),
                    "reserve_msfo": _clean_num(val_res),
                    "currency": currency,
                    "report_date": r_date,
                })
                global_id += 1
                ord_counter += 1

        logger.info("Лист '%s': обработано дат — %d", currency, len(dates_target))

    if not all_rows:
        raise OvpDataError("Не найдено подходящих данных ни на одном листе.")

    return pd.DataFrame(all_rows, columns=OUT_COLUMNS)


def save_report(df: pd.DataFrame, output_path: Path) -> Path:
    """Сохраняет результат в CSV с разделителем ';' и кодировкой utf-8-sig (для Excel)."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        df.to_csv(output_path, sep=";", index=False, encoding="utf-8-sig")
    except PermissionError as exc:
        raise OvpDataError(
            f"Нет доступа для записи в {output_path} (файл открыт в другой программе?): {exc}"
        ) from exc
    logger.info("Отчёт ОВП сохранён: %s (%d строк)", output_path, len(df))
    return output_path
