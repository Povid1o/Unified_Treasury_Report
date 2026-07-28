"""ETL отчёта «Структура баланса» (BalanceStruct): Excel «ПФ» -> плоский CSV.

Бизнес-логика перенесена дословно из EXPORT_FOLDER/BalanceStruct.ipynb
("Финальная рабочая версия с поиском и удалением дубликатов") — вынесены
только запуск через input()/захардкоженные пути и print() заменены на
логирование.
"""
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

BASE_DIR = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(BASE_DIR))

from common import excel_io  # noqa: E402
from common.logging_utils import get_logger  # noqa: E402

logger = get_logger("balance_struct", BASE_DIR / "logs")

# Лист в исходном файле называется именно так (опечатка "Прензентация" —
# часть реального имени листа в файлах "ПФ_*.xlsx", а не ошибка переноса).
SHEET_NAME = "Прензентация"

# Строки, которые нужно выбросить перед разбором иерархии.
DROP_ROW_NAMES = [
    "Итого",
    "МОБ портфель ценных бумаг",
    "МОБ кредитный портфель",
    "Резервы по кредитному портфелю МОБ",
]

# Индексы колонок (1 = C, 2 = D в Excel), для которых nversionid = 0.
COLS_NVERSION_ZERO = [1, 2]

PRECISION = 0

OUT_COLUMNS = ["date_", "axis_0", "axis_1", "axis_2", "axis_3", "value", "axis_4", "nversionid"]

# Иерархия категорий/подкатегорий «Плана фондирования».
NESTED_BALANCE: Dict[str, Any] = {
    "Активы": {
        "level": "axis 1",
        "children": {
            "Ликвидные активы": "axis 2",
            "ФОР": "axis 2",
            "МБК, Обратное РЕПО": "axis 2",
            "МБК ВГО размещение": "axis 2",
            "Портфель ценных бумаг": {
                "level": "axis 2",
                "children": {
                    "Казначейство": "axis 3",
                    "БФР": "axis 3",
                    "Прочие": "axis 3",
                },
            },
            "Кредитный портфель (гросс)": {
                "level": "axis 2",
                "children": {
                    "КБ": "axis 3",
                    "Спец проекты": "axis 3",
                    "ОПК": "axis 3",
                    "МБ": "axis 3",
                    "РБ": "axis 3",
                    "ДЧК": "axis 3",
                    "ДБП": "axis 3",
                    "ДМБ": "axis 3",
                    "Прочие клиенты": "axis 3",
                    "Новые территории": "axis 3",
                },
            },
            "Резервы по кредитному портфелю": "axis 2",
            "Неработающие и прочие активы": "axis 2",
        },
    },
    "Пассивы": {
        "level": "axis 1",
        "children": {
            "МБК, Лоро, Прямое РЕПО": "axis 2",
            "МБК ВГО привлечение": "axis 2",
            "Клиентское привлечение (в т.ч. ЦФА*)": {
                "level": "axis 2",
                "children": {
                    "ДВС": {
                        "level": "axis 3",
                        "children": {
                            "КБ": "axis 4",
                            "Спец проекты": "axis 4",
                            "ОПК": "axis 4",
                            "МБ": "axis 4",
                            "РБ": "axis 4",
                            "ДЧК": "axis 4",
                            "ДБП": "axis 4",
                            "ДМБ": "axis 3",
                            "Прочие клиенты": "axis 4",
                            "Новые территории": "axis 4",
                        },
                    },
                    "Срочное": {
                        "level": "axis 3",
                        "children": {
                            "КБ": "axis 4",
                            "Спец проекты": "axis 4",
                            "ОПК": "axis 4",
                            "МБ": "axis 4",
                            "РБ": "axis 4",
                            "ДЧК": "axis 4",
                            "ДБП": "axis 4",
                            "ДМБ": "axis 4",
                            "Прочие клиенты, в т.ч. депозиты ФК": "axis 4",
                            "Новые территории": "axis 4",
                        },
                    },
                    "Цифровые финансовые активы": "axis 3",
                },
            },
            "Средства ГОЗ": "axis 2",
            "Средства Минпромторга": "axis 2",
            "Бюджетные средства": "axis 2",
            "Привлечение на финансовых рынках": "axis 2",
            "Прочие обязательства": "axis 2",
            "Капитал": "axis 2",
            "Внебаланс": "axis 2",
            "НКЛ": "axis 2",
            "НЧСФ": "axis 2",
        },
    },
}


class BalanceStructError(RuntimeError):
    """Ошибка чтения или обработки Excel-отчёта BalanceStruct."""


# ── 0. Сброс дублей ──────────────────────────────────────────────────────────
def drop_exact_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    """Удаляет строки, полностью дублирующие друг друга (дата, все оси,
    значение и nversionid совпадают)."""
    original_count = len(df)
    df_clean = df.drop_duplicates(keep="first").reset_index(drop=True)
    deleted_count = original_count - len(df_clean)
    if deleted_count > 0:
        logger.info("Удалено полных дублей: %d строк.", deleted_count)
    else:
        logger.info("Полных дублей не найдено.")
    return df_clean


# ── 1. Округление ────────────────────────────────────────────────────────────
def smart_round(value: Any, precision: int = 3) -> Any:
    try:
        num = float(value)
        if pd.isna(num):
            return np.nan
        return int(round(num)) if precision <= 0 else round(num, precision)
    except (ValueError, TypeError):
        return value


# ── 2. Построение карты путей (с поддержкой дубликатов имён) ────────────────
def build_name_to_path(nested: Dict) -> Dict[str, List[Tuple]]:
    result: Dict[str, List[Tuple]] = {}

    def _recurse(node_name: str, node_data: Any, path: List[str]) -> None:
        # Ключ нормализован (см. excel_io.normalize_label): в выгрузках статья
        # может отличаться пробелами/регистром от эталона в NESTED_BALANCE.
        clean_name = excel_io.normalize_label(node_name)
        axis = [None] * 4
        for i in range(min(len(path), 4)):
            axis[i] = path[i]

        result.setdefault(clean_name, []).append(tuple(axis))

        if isinstance(node_data, dict) and "children" in node_data:
            for child_name, child_node in node_data["children"].items():
                _recurse(child_name, child_node, path + [child_name])

    for root_name, root_node in nested.items():
        _recurse(root_name, root_node, [root_name])
    return result


# ── 2.5. Очистка DataFrame (удаление строк/колонок) ──────────────────────────
def apply_cleaning(
    df: pd.DataFrame,
    drop_col_idxs: Optional[List[int]] = None,
    drop_col_names: Optional[List[str]] = None,
    drop_row_idxs: Optional[List[int]] = None,
    drop_row_names: Optional[List[str]] = None,
) -> pd.DataFrame:
    """Удаляет строки и колонки по индексам или именам."""
    df = df.copy()

    if drop_col_idxs:
        cols_to_drop = [df.columns[i] for i in drop_col_idxs if i < len(df.columns)]
        df = df.drop(columns=cols_to_drop)
        logger.info("Удалены колонки по индексу: %s", drop_col_idxs)

    if drop_col_names:
        existing = [c for c in drop_col_names if c in df.columns]
        df = df.drop(columns=existing)
        logger.info("Удалены колонки по имени: %s", existing)

    if drop_row_idxs:
        existing_idxs = [i for i in drop_row_idxs if i in df.index]
        df = df.drop(index=existing_idxs)
        logger.info("Удалены строки по индексу: %s", existing_idxs)

    if drop_row_names:
        first_col = df.columns[0]
        # Сравниваем нормализованно — иначе "Итого" с висячим пробелом или
        # неразрывным пробелом не отсеется и попадёт в отчёт лишней строкой.
        targets = {excel_io.normalize_label(n) for n in drop_row_names}
        mask = df[first_col].map(lambda v: excel_io.normalize_label(v) in targets)
        removed = df.loc[mask, first_col].tolist()
        df = df[~mask]
        logger.info("Удалены строки по имени статьи: %s", removed)

    return df.reset_index(drop=True)


# ── 3. Чтение и очистка (с динамическим якорем) ──────────────────────────────
def read_raw_file(file_path: Path, sheet_name: Any = 0) -> pd.DataFrame:
    """Читает лист как голую матрицу, находит ячейку-якорь "структура баланса
    ... млрд" и обрезает всё, что выше и левее неё, затем собирает заголовки.
    """
    try:
        # read_matrix вместо pd.read_excel: разворачивает объединённые ячейки
        # (иначе якорь и названия статей из выгрузок кубов не находятся) и
        # объясняет ситуацию с формулами без кэша. См. common/excel_io.
        raw = excel_io.read_matrix(file_path, sheet_name=sheet_name, logger=logger)
    except excel_io.ExcelSourceError as exc:
        raise BalanceStructError(str(exc)) from exc
    except Exception as exc:
        raise BalanceStructError(f"Не удалось прочитать файл {file_path} (лист {sheet_name!r}): {exc}") from exc

    anchor_row = None
    anchor_col = None
    for r in range(len(raw)):
        for c in range(len(raw.columns)):
            cell_val = str(raw.iloc[r, c]).lower().replace("\n", " ")
            if "структура баланса" in cell_val and "млрд" in cell_val:
                anchor_row, anchor_col = r, c
                break
        if anchor_row is not None:
            break

    if anchor_row is None:
        raise BalanceStructError("Якорь 'Структура баланса...' не найден в файле!")

    logger.info("Якорь найден на строке %d, колонке %d. Обрезаю таблицу...", anchor_row, anchor_col)

    df = raw.iloc[anchor_row:, anchor_col:].copy().reset_index(drop=True)

    new_cols = []
    for i in range(len(df.columns)):
        row_idx = 0 if i == 0 or i == 4 else 1
        new_cols.append(df.iloc[row_idx, i])
    df.columns = new_cols
    df = df.drop([0, 1]).reset_index(drop=True)

    if len(df.columns) > 5:
        df = df.drop(df.columns[5], axis=1)

    return df.dropna(axis=1, how="all")


# ── 4. Парсинг с контекстом и nversionid (C, D = 0) ──────────────────────────
def parse_raw_rows(
    df: pd.DataFrame,
    name_map: Dict[str, List[Tuple]],
    cols_nversion_zero: Optional[List[int]] = None,
) -> pd.DataFrame:
    """cols_nversion_zero: индексы колонок (начиная с 1), которым нужно
    присвоить nversionid = 0. Остальным — 1."""
    if cols_nversion_zero is None:
        cols_nversion_zero = []

    new_data_list = []
    active_context: List[Optional[str]] = [None, None, None, None]

    for i in range(len(df)):
        statya = str(df.iloc[i, 0]).strip()

        candidates = name_map.get(excel_io.normalize_label(statya), [])
        selected_axis: Tuple = (None, None, None, None)

        if not candidates:
            pass
        elif len(candidates) == 1:
            selected_axis = candidates[0]
        else:
            best_candidate = None
            max_matches = -1
            for cand in candidates:
                matches = 0
                if cand[0] == active_context[0]:
                    matches += 1
                if cand[1] == active_context[1]:
                    matches += 1
                if cand[2] == active_context[2]:
                    matches += 1
                if matches > max_matches:
                    max_matches = matches
                    best_candidate = cand
            selected_axis = best_candidate if best_candidate else candidates[0]

        if selected_axis[0]:
            level_idx = -1
            if selected_axis[3]:
                level_idx = 3
            elif selected_axis[2]:
                level_idx = 2
            elif selected_axis[1]:
                level_idx = 1
            elif selected_axis[0]:
                level_idx = 0

            for lvl in range(4):
                active_context[lvl] = selected_axis[lvl] if lvl <= level_idx else None

        for j in range(1, len(df.columns)):
            col_header = df.columns[j]
            value = df.iloc[i, j]
            n_val = 0 if j in cols_nversion_zero else 1

            new_data_list.append(
                {
                    "cat": statya,
                    "date_": str(col_header),
                    "value": value,
                    "nversionid": n_val,
                    "axis_1": selected_axis[0],
                    "axis_2": selected_axis[1],
                    "axis_3": selected_axis[2],
                    "axis_4": selected_axis[3],
                }
            )

    return pd.DataFrame(new_data_list)


# ── 5. Обогащение ────────────────────────────────────────────────────────────
def enrich_df(parsed_df: pd.DataFrame, precision: int = 3) -> pd.DataFrame:
    parsed_df = parsed_df.copy()
    parsed_df["date_"] = pd.to_datetime(parsed_df["date_"], dayfirst=True, errors="coerce")
    parsed_df["value"] = pd.to_numeric(parsed_df["value"], errors="coerce")

    mask_valid = (
        parsed_df["value"].notna()
        & parsed_df["cat"].str.strip().ne("")
        & parsed_df["date_"].notna()
    )
    df = parsed_df.loc[mask_valid].copy()

    unmapped_mask = df["axis_1"].isna()
    if unmapped_mask.any():
        unmapped_cats = df.loc[unmapped_mask, "cat"].unique()
        logger.warning(
            "Не найдено в nested_structure: %d статей, отброшено %d строк. Примеры: %s",
            len(unmapped_cats),
            unmapped_mask.sum(),
            list(unmapped_cats[:10]),
        )
        df = df[~unmapped_mask].copy()

    df["date_"] = df["date_"].dt.strftime("%Y-%m-%d")
    df["value"] = df["value"].apply(lambda x: smart_round(x, precision))
    df["axis_0"] = "План фондирования"

    return df[["date_", "axis_0", "axis_1", "axis_2", "axis_3", "value", "axis_4", "nversionid"]]


# ── 6. Проверка на дубликаты ─────────────────────────────────────────────────
def check_duplicates(df: pd.DataFrame) -> None:
    subset_cols = ["date_", "axis_1", "axis_2", "axis_3", "axis_4", "nversionid"]
    duplicates = df[df.duplicated(subset=subset_cols, keep=False)]

    if not duplicates.empty:
        logger.warning("Найдено дублей (по %s): %d", subset_cols, len(duplicates))
    else:
        logger.info("Дубликатов не найдено.")


# ── 7. Конвейер ───────────────────────────────────────────────────────────────
def build_report(file_path: Path) -> pd.DataFrame:
    """Полный цикл: чтение Excel -> разбор иерархии -> плоский DataFrame."""
    file_path = Path(file_path)
    if not file_path.exists():
        raise BalanceStructError(f"Файл не найден: {file_path}")

    name_map = build_name_to_path(NESTED_BALANCE)
    raw_df = read_raw_file(file_path, sheet_name=SHEET_NAME)

    cleaned_df = apply_cleaning(raw_df, drop_row_names=DROP_ROW_NAMES)
    parsed = parse_raw_rows(cleaned_df, name_map=name_map, cols_nversion_zero=COLS_NVERSION_ZERO)

    # Ни одной строки со значением — типично для выгрузок из куба, где числа
    # остались формулами без сохранённого результата (см. common/excel_io).
    # Без этой проверки enrich_df падал с невнятным KeyError: 'date_'.
    if parsed.empty or parsed["value"].isna().all():
        explanation = excel_io.describe_uncached_formulas(file_path, SHEET_NAME)
        raise BalanceStructError(
            explanation
            or f"В файле {file_path.name} найден якорь таблицы, но все значения пусты."
        )

    final_df = enrich_df(parsed, PRECISION)

    final_df = drop_exact_duplicates(final_df)
    check_duplicates(final_df)

    if final_df.empty:
        raise BalanceStructError("Итоговый DataFrame пуст — не найдено ни одной строки данных.")

    return final_df.reset_index(drop=True)[OUT_COLUMNS]


def save_report(df: pd.DataFrame, output_path: Path) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        df.to_csv(output_path, index=False, encoding="utf-8-sig")
    except PermissionError as exc:
        raise BalanceStructError(
            f"Нет доступа для записи в {output_path} (файл открыт в другой программе?): {exc}"
        ) from exc
    logger.info("Отчёт сохранён: %s (%d строк)", output_path, len(df))
    return output_path
