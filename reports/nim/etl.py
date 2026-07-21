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

from common.logging_utils import get_logger  # noqa: E402

logger = get_logger("nim", BASE_DIR / "logs")

MARKER_OSNOVA = "NIM - ОСНОВА"
MARKER_RUR = "NIM - RUR"
MARKER_VALUTA = "NIM - ВАЛЮТА"

# Маппинг относительных смещений строк от маркера к осям (a1, a2, a3).
# Смещения считаются от строки маркера внутри объединённого блока rows_to_keep
# (см. build_report): 0-3 — блок ОСНОВА, 6-7 — блок RUR, 10-11 — блок ВАЛЮТА.
MAPPING: Dict[int, Dict[str, str]] = {
    # --- Блок 1: NIM - ОСНОВА (строки 0-3) ---
    0: {"a1": "NIM", "a2": "ALL", "a3": ""},
    1: {"a1": "% результат", "a2": "ALL", "a3": ""},
    2: {"a1": "NIM", "a2": "RUR", "a3": ""},
    3: {"a1": "NIM", "a2": "CUR", "a3": ""},
    # --- Блок 2: NIM - RUR (строки 6-7) ---
    6: {"a1": "% Активы", "a2": "RUR", "a3": ""},
    7: {"a1": "% Пассивы", "a2": "RUR", "a3": ""},
    # --- Блок 3: NIM - ВАЛЮТА (строки 10-11) ---
    10: {"a1": "% Активы", "a2": "CUR", "a3": ""},
    11: {"a1": "% Пассивы", "a2": "CUR", "a3": ""},
}

OUT_COLUMNS = ["id", "date_", "axis_0", "axis_1", "axis_2", "axis_3", "value", "axis_4", "nversionid", "axis_5"]


class NimDataError(RuntimeError):
    """Ошибка чтения или обработки Excel-отчёта NIM."""


def _find_markers(df_full: pd.DataFrame) -> Tuple[Optional[int], Optional[int], Optional[int]]:
    """Ищет строки-маркеры в первой колонке. Возвращает (idx_osnova, idx_rur, idx_valuta)."""
    idx_osnova, idx_rur, idx_valuta = None, None, None
    for i, val in enumerate(df_full.iloc[:, 0]):
        cell_val = str(val).strip()
        if cell_val == MARKER_OSNOVA:
            idx_osnova = i
        elif cell_val == MARKER_RUR:
            idx_rur = i
        elif cell_val == MARKER_VALUTA:
            idx_valuta = i
    return idx_osnova, idx_rur, idx_valuta


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
        df_full = pd.read_excel(file_path, usecols="A:N", header=None)
    except Exception as exc:
        raise NimDataError(f"Не удалось прочитать файл {file_path}: {exc}") from exc

    idx_osnova, idx_rur, idx_valuta = _find_markers(df_full)
    if idx_osnova is None and idx_rur is None and idx_valuta is None:
        raise NimDataError(
            f"В файле не найден ни один из маркеров: "
            f"'{MARKER_OSNOVA}', '{MARKER_RUR}', '{MARKER_VALUTA}'"
        )

    rows_to_keep: List[int] = []
    if idx_osnova is not None:
        rows_to_keep.extend(range(idx_osnova, idx_osnova + 5))  # маркер (в нём даты) + 4 строки данных
    else:
        logger.warning("Маркер '%s' не найден — блок ОСНОВА (NIM/% результат) будет пропущен.", MARKER_OSNOVA)
    if idx_rur is not None:
        rows_to_keep.extend(range(idx_rur, idx_rur + 4))  # маркер + пустая + 2 строки данных
    else:
        logger.warning("Маркер '%s' не найден — блок RUR (% Активы/Пассивы) будет пропущен.", MARKER_RUR)
    if idx_valuta is not None:
        rows_to_keep.extend(range(idx_valuta, idx_valuta + 5))  # маркер + пустая + 2 строки данных + пустая
    else:
        logger.warning("Маркер '%s' не найден — блок ВАЛЮТА (% Активы/Пассивы) будет пропущен.", MARKER_VALUTA)

    rows_to_keep = sorted(set(rows_to_keep))
    df = df_full.iloc[rows_to_keep].reset_index(drop=True)

    # Заголовки (даты) берутся из нулевой строки — строки с маркером ОСНОВА.
    raw_names = df.iloc[0, 1:].tolist()
    new_names = ["Тип данных"]
    for val in raw_names:
        if pd.isna(val):
            new_names.append("Пустая_дата")
        else:
            try:
                new_names.append(str(int(float(val))))
            except (ValueError, TypeError):
                new_names.append(str(val))
    df.columns = new_names
    df = df.drop(df.index[0]).reset_index(drop=True)

    output_data = []
    row_id_counter = 1
    date_cols = [c for c in df.columns if c != "Тип данных"]

    for col in date_cols:
        formatted_date = _transform_date(col)

        for idx, axis_info in MAPPING.items():
            if idx not in df.index:
                continue
            val = df.loc[idx, col]

            try:
                val = float(val)
                if axis_info["a1"] != "% результат":  # умножаем на 100 всё, кроме "% результат"
                    val = val * 100
                val = round(val, 2)
            except (ValueError, TypeError):
                pass  # текст или NaN — не трогаем

            output_data.append({
                "id": row_id_counter,
                "date_": formatted_date,
                "axis_0": "NIM",
                "axis_1": axis_info["a1"],
                "axis_2": axis_info["a2"],
                "axis_3": axis_info["a3"],
                "value": val,
                "axis_4": "",
                "nversionid": "",
                "axis_5": "",
            })
            row_id_counter += 1

    result_df = pd.DataFrame(output_data, columns=OUT_COLUMNS)
    result_df["value"] = pd.to_numeric(result_df["value"], errors="coerce")

    if result_df.empty:
        raise NimDataError("Итоговый DataFrame пуст — не найдено ни одной строки данных.")

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
