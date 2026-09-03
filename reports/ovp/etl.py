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

from common import excel_io  # noqa: E402
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
TITLE_DATE_PATTERN = re.compile(r"(?<!\d)(\d{2}\.\d{2}\.\d{4})(?!\d)")

# В новом формате все валюты находятся на одном листе. В выгрузке название
# евро исторически записывается как EURO, поэтому возможный заголовок EUR
# также приводим к этому значению, чтобы не менять контракт выходного CSV.
CONSOLIDATED_CURRENCIES: Dict[str, str] = {
    "cny": "CNY",
    "usd": "USD",
    "eur": "EURO",
    "euro": "EURO",
}
BALANCE_HEADER = excel_io.normalize_label("Валютный Баланс")
RESERVE_HEADER = excel_io.normalize_label("Резервы МСФО")
HEADER_SCAN_ROWS = 12

# Нормализованный вид -> эталонное имя из HIERARCHY. Считается один раз при
# импорте. Нужен, чтобы статья находилась независимо от пробелов и регистра в
# файле (см. excel_io.normalize_label), а в выгрузку шло эталонное написание,
# а не то, как её записал конкретный файл.
_NORM_CATEGORIES: Dict[str, str] = {
    excel_io.normalize_label(cat): cat for cat in HIERARCHY
}
_NORM_SUBCATEGORIES: Dict[str, Dict[str, str]] = {
    cat: {excel_io.normalize_label(sub): sub for sub in subs}
    for cat, subs in HIERARCHY.items()
}


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


def list_sheets(input_path: Path) -> List[str]:
    """Возвращает все листы книги в порядке их расположения."""
    input_path = Path(input_path)
    if not input_path.exists():
        raise OvpDataError(f"Файл не найден: {input_path}")
    try:
        xls = pd.ExcelFile(input_path)
    except Exception as exc:
        raise OvpDataError(f"Не удалось прочитать Excel-файл {input_path}: {exc}") from exc
    try:
        return list(xls.sheet_names)
    finally:
        xls.close()


def list_currency_sheets(input_path: Path) -> List[str]:
    """Список валютных листов старого формата (служебные СВОД исключены)."""
    return [sheet for sheet in list_sheets(input_path) if "СВОД" not in sheet.upper()]


def _resolve_sheet_selection(available: List[str], sheet_name: Optional[str]) -> str:
    """Автовыбор единственного листа или проверка явного выбора пользователя."""
    if not available:
        raise OvpDataError("В книге нет листов.")
    if sheet_name is None:
        if len(available) == 1:
            return available[0]
        raise OvpDataError(
            "В книге несколько листов — автоматический выбор отключён. "
            f"Выберите один лист: {available}. При запуске из командной строки "
            "используйте параметр --sheet."
        )
    if sheet_name not in available:
        raise OvpDataError(
            f"Лист {sheet_name!r} не найден. Доступные листы: {available}"
        )
    return sheet_name


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


def _find_title_date(df_raw: pd.DataFrame) -> Optional[str]:
    """Возвращает дату YYYY-MM-DD из заголовка нового однолистового отчёта."""
    for _idx, row in df_raw.head(HEADER_SCAN_ROWS).iterrows():
        for value in row:
            if not isinstance(value, str):
                continue
            match = TITLE_DATE_PATTERN.search(value)
            if not match:
                continue
            try:
                return pd.to_datetime(match.group(1), dayfirst=True).strftime("%Y-%m-%d")
            except (TypeError, ValueError):
                continue
    return None


def _canonical_currency(value) -> Optional[str]:
    """Приводит заголовок валюты нового отчёта к значению выходного CSV."""
    normalized = excel_io.normalize_label(value)
    return CONSOLIDATED_CURRENCIES.get(normalized)


def _find_consolidated_layout(df_raw: pd.DataFrame) -> Optional[List[Tuple[str, int, int]]]:
    """Ищет валютные блоки нового формата.

    Возвращает список ``(валюта, колонка баланса, колонка резервов)`` в том
    порядке, в котором блоки расположены на листе. Колонки ИТОГО намеренно не
    возвращаются. Заголовки валют могут быть объединёнными ячейками: pandas
    видит значение только в их левой верхней ячейке, чего здесь достаточно.
    """
    ncols = len(df_raw.columns)
    for row_idx, row in df_raw.head(HEADER_SCAN_ROWS).iterrows():
        currency_starts: List[Tuple[str, int]] = []
        seen = set()
        for col_idx, value in enumerate(row):
            currency = _canonical_currency(value)
            if currency and currency not in seen:
                currency_starts.append((currency, col_idx))
                seen.add(currency)

        if not currency_starts:
            continue

        layout: List[Tuple[str, int, int]] = []
        for pos, (currency, start_col) in enumerate(currency_starts):
            end_col = currency_starts[pos + 1][1] if pos + 1 < len(currency_starts) else ncols
            balance_col = None
            reserve_col = None

            # В текущем шаблоне названия показателей находятся на следующей
            # строке, но небольшой запас делает чтение устойчивым к вставке
            # служебной строки над ними.
            for header_row in range(row_idx + 1, min(row_idx + 4, len(df_raw))):
                for col_idx in range(start_col, end_col):
                    normalized = excel_io.normalize_label(df_raw.iat[header_row, col_idx])
                    if normalized == BALANCE_HEADER:
                        balance_col = col_idx
                    elif normalized == RESERVE_HEADER:
                        reserve_col = col_idx
                if balance_col is not None and reserve_col is not None:
                    break

            if balance_col is not None and reserve_col is not None:
                layout.append((currency, balance_col, reserve_col))

        found_currencies = {currency for currency, _balance, _reserve in layout}
        if found_currencies == {"CNY", "USD", "EURO"}:
            return layout

    return None


def _match_hierarchy_label(clean: str, current_category: Optional[str]) -> Tuple[Optional[str], Optional[str]]:
    """Сопоставляет подпись строки с эталонной категорией/подкатегорией."""
    norm = excel_io.normalize_label(clean)
    if norm in _NORM_CATEGORIES:
        return _NORM_CATEGORIES[norm], ""

    if current_category:
        subcategory = _NORM_SUBCATEGORIES[current_category].get(norm)
        if subcategory:
            return current_category, subcategory

        # В новом файле к подписи строки добавлена фактическая дата формы:
        # "ОВП БФР за 01.09.2026". Для прежнего выходного формата суффикс
        # отбрасывается; report_date всегда берётся из заголовка файла.
        if current_category == "ОВП (634 форма)":
            ovp_bfr_norm = excel_io.normalize_label("ОВП БФР")
            dated_ovp_bfr = rf"{re.escape(ovp_bfr_norm)}за\d{{2}}\.\d{{2}}\.\d{{4}}(?:г\.?)?"
            if re.fullmatch(dated_ovp_bfr, norm):
                return current_category, "ОВП БФР"

    return current_category, None


def _append_currency_rows(
    all_rows: List[dict],
    df_raw: pd.DataFrame,
    currency: str,
    report_date: str,
    col_bal: int,
    col_res: int,
    global_id: int,
    data_start_row: int = 0,
) -> int:
    """Добавляет распознанные строки одной валюты и возвращает следующий id."""
    ord_counter = 1
    current_category = None
    ncols = len(df_raw.columns)

    for r_idx in range(data_start_row, len(df_raw)):
        row_data = df_raw.iloc[r_idx]
        clean = clean_name(row_data.iloc[0])
        if not clean:
            continue

        matched_category, target_subcategory = _match_hierarchy_label(clean, current_category)
        if target_subcategory is None:
            continue
        current_category = matched_category

        val_bal = row_data.iloc[col_bal] if col_bal < ncols else 0
        val_res = row_data.iloc[col_res] if col_res < ncols else 0
        all_rows.append({
            "id": global_id,
            "ord": ord_counter,
            "category_name": current_category,
            "subcategory": target_subcategory,
            "curr_balance": _clean_num(val_bal),
            "reserve_msfo": _clean_num(val_res),
            "currency": currency,
            "report_date": report_date,
        })
        global_id += 1
        ord_counter += 1

    return global_id


def _convert_consolidated_sheet(
    df_raw: pd.DataFrame,
    layout: List[Tuple[str, int, int]],
    currencies: Optional[List[str]],
) -> pd.DataFrame:
    """Преобразует новый однолистовый макет в прежнюю плоскую структуру."""
    report_date = _find_title_date(df_raw)
    if report_date is None:
        raise OvpDataError("В заголовке сводного листа не найдена дата в формате ДД.ММ.ГГГГ.")

    available = [currency for currency, _bal, _res in layout]
    if currencies is None:
        selected = set(available)
    else:
        requested = []
        for value in currencies:
            canonical = _canonical_currency(value)
            if canonical is None:
                canonical = str(value).strip().upper()
            requested.append(canonical)
        unknown = set(requested) - set(available)
        if unknown:
            raise OvpDataError(
                f"Валюты не найдены в сводном листе: {sorted(unknown)}. Доступны: {available}"
            )
        selected = set(requested)

    all_rows: List[dict] = []
    global_id = 1
    for currency, col_bal, col_res in layout:
        if currency not in selected:
            continue
        global_id = _append_currency_rows(
            all_rows=all_rows,
            df_raw=df_raw,
            currency=currency,
            report_date=report_date,
            col_bal=col_bal,
            col_res=col_res,
            global_id=global_id,
        )

    if not all_rows:
        raise OvpDataError("На сводном листе не найдено строк иерархии ОВП.")
    return pd.DataFrame(all_rows, columns=OUT_COLUMNS)


def convert_ovp_report(
    input_path: Path,
    full_history_currencies: Optional[List[str]] = None,
    currencies: Optional[List[str]] = None,
    sheet_name: Optional[str] = None,
) -> pd.DataFrame:
    """Преобразует Excel-отчёт ОВП в плоский DataFrame.

    :param input_path: путь к исходному .xlsx.
    Новый формат (валютные блоки на одном листе) определяется автоматически.
    Для него дата берётся из заголовка, а колонки ИТОГО игнорируются. Если
    новый макет не найден, применяется прежняя логика валютных листов.

    :param full_history_currencies: валютные листы старого формата, для которых
        нужно выгрузить все найденные даты. В новом формате всегда одна дата.
    :param currencies: ограничить обработку конкретными валютами.
    :param sheet_name: лист для обработки. Если в книге один лист, он выбирается
        автоматически; если листов несколько, параметр обязателен.
    """
    input_path = Path(input_path)
    if not input_path.exists():
        raise OvpDataError(f"Файл не найден: {input_path}")

    try:
        xls = pd.ExcelFile(input_path)
    except Exception as exc:
        raise OvpDataError(f"Не удалось прочитать Excel-файл {input_path}: {exc}") from exc

    sheet_matrices: List[Tuple[str, pd.DataFrame]] = []
    try:
        selected_sheet = _resolve_sheet_selection(list(xls.sheet_names), sheet_name)
        try:
            sheet_matrices.append(
                (
                    selected_sheet,
                    pd.read_excel(xls, sheet_name=selected_sheet, header=None),
                )
            )
        except Exception as exc:
            raise OvpDataError(
                f"Не удалось прочитать лист {selected_sheet!r}: {exc}"
            ) from exc
    finally:
        xls.close()

    if not sheet_matrices:
        raise OvpDataError("В книге нет доступных для чтения листов.")

    return _convert_sheet_matrices(
        sheet_matrices=sheet_matrices,
        full_history_currencies=full_history_currencies,
        currencies=currencies,
    )


def _convert_sheet_matrices(
    sheet_matrices: List[Tuple[str, pd.DataFrame]],
    full_history_currencies: Optional[List[str]] = None,
    currencies: Optional[List[str]] = None,
) -> pd.DataFrame:
    """Преобразует уже прочитанные листы; отдельно вынесено для тестирования."""
    for sheet_name, df_raw in sheet_matrices:
        layout = _find_consolidated_layout(df_raw)
        if layout is None:
            continue
        if full_history_currencies:
            logger.info(
                "Лист '%s': новый однодатный формат, параметр полного исторического периода не применяется.",
                sheet_name,
            )
        result = _convert_consolidated_sheet(df_raw, layout, currencies)
        logger.info(
            "Лист '%s': обработан новый сводный формат (%s).",
            sheet_name,
            ", ".join(result["currency"].drop_duplicates()),
        )
        return result

    available = [
        sheet_name for sheet_name, _df in sheet_matrices
        if "СВОД" not in sheet_name.upper()
    ]
    target_currencies = currencies if currencies is not None else available
    unknown = set(target_currencies) - set(available)
    if unknown:
        raise OvpDataError(
            f"Листы не найдены в файле: {sorted(unknown)}. Доступны: {available}"
        )

    frames_by_name = dict(sheet_matrices)
    full_history_set = set(full_history_currencies or [])

    all_rows = []
    global_id = 1

    for currency in target_currencies:
        df_raw = frames_by_name[currency]

        date_row_idx, dates_found = _find_date_row(df_raw)
        if date_row_idx is None:
            logger.warning("Лист '%s' пропущен: не найдена строка с датами", currency)
            continue

        dates_target = dates_found if currency in full_history_set else [dates_found[-1]]
        data_start_row = date_row_idx + 2

        for r_date, col_idx in dates_target:
            col_bal, col_res = col_idx, col_idx + 1
            global_id = _append_currency_rows(
                all_rows=all_rows,
                df_raw=df_raw,
                currency=currency,
                report_date=r_date,
                col_bal=col_bal,
                col_res=col_res,
                global_id=global_id,
                data_start_row=data_start_row,
            )

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
