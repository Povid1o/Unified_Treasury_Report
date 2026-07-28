"""ETL отчёта «NIM»: Excel «NIM_YYYY_MM» -> плоский CSV.

Бизнес-логика перенесена дословно из EXPORT_FOLDER/NIM.ipynb, ячейка
"Версия кода с МАРКЕРАМИ" (более совершенная версия — ищет данные по
текстовым маркерам "NIM - ОСНОВА"/"NIM - RUR"/"NIM - ВАЛЮТА" в колонке A,
а не по захардкоженным номерам строк, как в "Основной версии"). Вынесены
только input() и print() заменены на аргумент функции и логирование.
"""
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

BASE_DIR = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(BASE_DIR))

from common import excel_io  # noqa: E402
from common.logging_utils import get_logger  # noqa: E402

logger = get_logger("nim", BASE_DIR / "logs")

MARKER_OSNOVA = "NIM - ОСНОВА"
MARKER_RUR = "NIM - RUR"
MARKER_VALUTA = "NIM - ВАЛЮТА"

# Нормализованные формы маркеров считаются один раз при импорте: сравнение
# идёт по ним, чтобы различия в пробелах/регистре не ломали поиск.
_NORM_OSNOVA = excel_io.normalize_label(MARKER_OSNOVA)
_NORM_RUR = excel_io.normalize_label(MARKER_RUR)
_NORM_VALUTA = excel_io.normalize_label(MARKER_VALUTA)

# Смещения строк данных блока ОСНОВА относительно строки маркера (+1..+4).
# Все 4 строки блока подписаны в колонке A одинаково ("NIM" встречается
# трижды с разным смыслом a2) — по тексту их не различить, поэтому здесь,
# в отличие от блоков RUR/ВАЛЮТА (см. _extract_block_rows), приходится
# полагаться на фиксированную позицию.
MAPPING_OSNOVA: Dict[int, Dict[str, str]] = {
    1: {"a1": "NIM", "a2": "ALL", "a3": ""},
    2: {"a1": "% результат", "a2": "ALL", "a3": ""},
    3: {"a1": "NIM", "a2": "RUR", "a3": ""},
    4: {"a1": "NIM", "a2": "CUR", "a3": ""},
}

# Метки строк данных блоков RUR/ВАЛЮТА — уникальны в колонке A, поэтому
# ищутся по содержимому (см. _extract_block_rows), а не по смещению.
RUR_VALUTA_LABELS: List[str] = ["% Активы", "% Пассивы"]

OUT_COLUMNS = ["id", "date_", "axis_0", "axis_1", "axis_2", "axis_3", "value", "axis_4", "nversionid", "axis_5"]


class NimDataError(RuntimeError):
    """Ошибка чтения или обработки Excel-отчёта NIM."""


def _find_markers(df_full: pd.DataFrame) -> Tuple[Optional[int], Optional[int], Optional[int]]:
    """Ищет строки-маркеры в первой колонке. Возвращает (idx_osnova, idx_rur, idx_valuta).

    Сравнение идёт по нормализованному виду (см. excel_io.normalize_label):
    в выгрузках маркер может быть записан с другими пробелами или регистром.
    """
    idx_osnova, idx_rur, idx_valuta = None, None, None
    for i, val in enumerate(df_full.iloc[:, 0]):
        cell_val = excel_io.normalize_label(val)
        if cell_val == _NORM_OSNOVA:
            idx_osnova = i
        elif cell_val == _NORM_RUR:
            idx_rur = i
        elif cell_val == _NORM_VALUTA:
            idx_valuta = i
    return idx_osnova, idx_rur, idx_valuta


def _find_label_row(df_full: pd.DataFrame, start_idx: int, end_idx: int, label_text: str) -> Optional[int]:
    """Ищет строку по метке в колонке A, сравнивая нормализованные значения —
    иначе "% Активы" в константе и "%Активы" в файле считались бы разными."""
    target = excel_io.normalize_label(label_text)
    for i in range(start_idx, end_idx):
        if excel_io.normalize_label(df_full.iloc[i, 0]) == target:
            return i
    return None


def _extract_block_rows(
    df_full: pd.DataFrame, marker_idx: int, window_end: int, expected_labels: List[str], label: str
) -> Optional[List[int]]:
    """Строки данных блока RUR/ВАЛЮТА, найденные по точному тексту метки в
    колонке A (те же подписи, что и в RUR_VALUTA_LABELS — "% Активы",
    "% Пассивы"), а не по фиксированному смещению от маркера.

    Раньше строки данных брались как range(marker_idx, marker_idx + N) —
    строго по позиции. Выгрузки из OLAP-куба иногда вставляют между
    маркером и данными служебные строки (фильтр, подытог, лишний
    разделитель) в любом месте блока — тогда позиционное смещение съезжает,
    и данные тихо подставляются не под те axis-метки (без единой ошибки:
    файл читается, просто значения оказываются не там). Поиск строк по
    содержимому устойчив к таким вставкам — лишние строки просто
    пропускаются, где бы они ни оказались, пока сама метка не переименована.

    window_end ограничивает поиск строкой следующего маркера (или концом
    листа), чтобы не залезть в чужой блок. Возвращает None, если хотя бы
    одна метка не найдена в пределах окна — блок в этом случае пропускается
    целиком (см. warning).
    """
    rows = []
    search_from = marker_idx + 1
    for expected in expected_labels:
        row_idx = _find_label_row(df_full, search_from, window_end, expected)
        if row_idx is None:
            logger.warning(
                "Блок '%s' (строка %d): не найдена строка с меткой '%s' до строки %d — блок пропущен.",
                label, marker_idx, expected, window_end,
            )
            return None
        rows.append(row_idx)
        search_from = row_idx + 1
    return rows


def _select_data_sheet(file_path: Path) -> Tuple[str, pd.DataFrame]:
    """Ищет среди ВСЕХ листов книги тот, где в первой колонке есть маркеры NIM.

    Раньше лист не задавался вообще, и чтение шло с первого по умолчанию
    (pd.read_excel(sheet_name=0)). В выгрузках из OLAP-кубов перед нужным
    листом обычно лежат служебные (кэш сводной таблицы, параметры подключения,
    титульный), поэтому первый лист оказывался пустым — отчёт молча выходил
    пустым. Ищем лист по тем же маркерам, по которым и так ищутся строки,
    чтобы не заводить ещё одну настройку и переживать переименование листов.
    """
    sheets_seen: List[str] = []
    for sheet_name, df in excel_io.iter_sheet_matrices(file_path):
        sheets_seen.append(sheet_name)
        if df.empty or len(df.columns) == 0:
            continue
        if any(idx is not None for idx in _find_markers(df)):
            return sheet_name, df

    raise NimDataError(
        f"Ни на одном листе файла {Path(file_path).name} не найдены маркеры "
        f"'{MARKER_OSNOVA}', '{MARKER_RUR}', '{MARKER_VALUTA}'. "
        f"Проверенные листы: {sheets_seen}"
    )


def _transform_date(val) -> str:
    """202501 -> 2025-01-01."""
    if pd.isna(val) or val == "Пустая_дата":
        return ""
    try:
        s = str(int(float(val)))
        if len(s) == 6:
            return f"{s[:4]}-{s[4:6]}-01"
        return s
    except (ValueError, TypeError):
        return str(val)


def build_report(file_path: Path) -> pd.DataFrame:
    """Полный цикл: поиск маркеров -> сборка блока строк -> перевод в длинный формат."""
    file_path = Path(file_path)
    if not file_path.exists():
        raise NimDataError(f"Файл не найден: {file_path}")

    try:
        # Лист выбирается по маркерам среди всех листов книги (см.
        # _select_data_sheet), а не берётся первый по умолчанию. Объединённые
        # ячейки при чтении разворачиваются — иначе маркеры в них не видны.
        sheet_name, df_full = _select_data_sheet(file_path)
    except NimDataError:
        raise
    except excel_io.ExcelSourceError as exc:
        raise NimDataError(str(exc)) from exc
    except Exception as exc:
        raise NimDataError(f"Не удалось прочитать файл {file_path}: {exc}") from exc

    logger.info("Данные найдены на листе %r", sheet_name)

    explanation = excel_io.describe_uncached_formulas(file_path, sheet_name)
    if explanation:
        logger.warning("Лист %r: часть значений не сохранена в файле. %s", sheet_name, explanation)

    # Ограничиваем колонками A:N. Раньше это делал usecols="A:N" в read_excel,
    # но он падал с ParserError, если в файле колонок меньше 14 — а у выгрузок
    # из кубов их часто меньше.
    df_full = df_full.iloc[:, :14]

    idx_osnova, idx_rur, idx_valuta = _find_markers(df_full)

    # Даты-заголовки берутся из строки самого раннего найденного маркера.
    header_marker_idx = min(i for i in (idx_osnova, idx_rur, idx_valuta) if i is not None)
    date_cols = [(col_pos, _transform_date(df_full.iloc[header_marker_idx, col_pos]))
                 for col_pos in range(1, len(df_full.columns))]

    output_data: List[dict] = []
    row_id_counter = 1

    def _append_value(row_idx: int, col_pos: int, formatted_date: str, a1: str, a2: str, a3: str) -> None:
        nonlocal row_id_counter
        val = df_full.iloc[row_idx, col_pos]
        try:
            val = float(val)
            if a1 != "% результат":  # умножаем на 100 всё, кроме "% результат"
                val = val * 100
            val = round(val, 2)
        except (ValueError, TypeError):
            pass  # текст или NaN — не трогаем

        output_data.append({
            "id": row_id_counter, "date_": formatted_date, "axis_0": "NIM",
            "axis_1": a1, "axis_2": a2, "axis_3": a3, "value": val,
            "axis_4": "", "nversionid": "", "axis_5": "",
        })
        row_id_counter += 1

    # --- Блок ОСНОВА: строки данных неразличимы по тексту (см. MAPPING_OSNOVA),
    # поэтому берём их по фиксированному смещению от маркера, но хотя бы
    # проверяем, что блок не залезает в следующий найденный маркер — если
    # залезает, значит данных меньше ожидаемого и доверять смещению нельзя.
    if idx_osnova is not None:
        next_marker = min((i for i in (idx_rur, idx_valuta) if i is not None), default=None)
        if next_marker is not None and next_marker <= idx_osnova + 4:
            raise NimDataError(
                f"Блок ОСНОВА (строка {idx_osnova}): между маркером '{MARKER_OSNOVA}' и следующим "
                f"маркером (строка {next_marker}) меньше 4 строк данных — структура файла "
                "не соответствует ожидаемой (похоже на нестандартную выгрузку из OLAP-куба)."
            )
        for offset, axis_info in MAPPING_OSNOVA.items():
            row_idx = idx_osnova + offset
            for col_pos, formatted_date in date_cols:
                _append_value(row_idx, col_pos, formatted_date, axis_info["a1"], axis_info["a2"], axis_info["a3"])
    else:
        logger.warning("Маркер '%s' не найден — блок ОСНОВА (NIM/%% результат) будет пропущен.", MARKER_OSNOVA)

    # --- Блоки RUR / ВАЛЮТА: строки данных ищутся по тексту метки в колонке A
    # (см. _extract_block_rows) — устойчиво к лишним строкам из OLAP-куба.
    for marker_idx, currency, marker_name, next_marker in (
        (idx_rur, "RUR", MARKER_RUR, idx_valuta),
        (idx_valuta, "CUR", MARKER_VALUTA, None),
    ):
        if marker_idx is None:
            logger.warning("Маркер '%s' не найден — блок %s (%% Активы/Пассивы) будет пропущен.", marker_name, currency)
            continue
        window_end = next_marker if next_marker is not None else len(df_full)
        found_rows = _extract_block_rows(df_full, marker_idx, window_end, RUR_VALUTA_LABELS, marker_name)
        if found_rows is None:
            continue
        for label_text, row_idx in zip(RUR_VALUTA_LABELS, found_rows):
            for col_pos, formatted_date in date_cols:
                _append_value(row_idx, col_pos, formatted_date, label_text, currency, "")

    result_df = pd.DataFrame(output_data, columns=OUT_COLUMNS)
    result_df["value"] = pd.to_numeric(result_df["value"], errors="coerce")

    if result_df.empty:
        raise NimDataError("Итоговый DataFrame пуст — не найдено ни одной строки данных.")

    # Строки нашлись, но все значения пустые — типично для выгрузок из куба,
    # где числа остались формулами без сохранённого результата. Молча отдать
    # такой отчёт нельзя: на выходе получится CSV из пустых value.
    if result_df["value"].notna().sum() == 0:
        raise NimDataError(
            explanation
            or f"На листе {sheet_name!r} файла {file_path.name} найдена структура "
               "отчёта (маркеры), но все значения пусты."
        )

    logger.info("Обработано полезных строк: %d", len(result_df))
    return result_df


def save_report(df: pd.DataFrame, output_path: Path) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        df.to_csv(output_path, index=False, sep=",", encoding="utf-8-sig")
    except PermissionError as exc:
        raise NimDataError(
            f"Нет доступа для записи в {output_path} (файл открыт в другой программе?): {exc}"
        ) from exc
    logger.info("Отчёт сохранён: %s (%d строк)", output_path, len(df))
    return output_path
