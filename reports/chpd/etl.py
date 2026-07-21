"""ETL отчёта «ЧПД»: Excel «ЧПД YYYY MM DD» -> плоский CSV.

Бизнес-логика перенесена дословно из EXPORT_FOLDER/CHPD.ipynb — вынесены
только запуск и print() заменены на логирование.
"""
import re
import sys
from pathlib import Path
from typing import List, Optional, Tuple

import pandas as pd

BASE_DIR = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(BASE_DIR))

from common.logging_utils import get_logger  # noqa: E402

logger = get_logger("chpd", BASE_DIR / "logs")

SHEET_NAME = "Table"

# ─────────────────────────────────────────────────────────────────────────────
# Иерархия строк: (leaf_name, axis_1, axis_2, axis_3, axis_4).
# Порядок строго соответствует строкам листа Table: Excel rows 4-17, 19-33
# (row 18 — пустой разделитель, dropna его уберёт).
# ─────────────────────────────────────────────────────────────────────────────
INDEX_PATHS: List[Tuple[str, Optional[str], Optional[str], Optional[str], Optional[str]]] = [
    # ── АКТИВЫ ────────────────────────────────────────────────────────────────
    ("АКТИВЫ",                     "Активы",  None,                          None,                    None               ),  # row 4
    ("Ликвидные активы",           "Активы",  "Ликвидные активы",            None,                    None               ),  # row 5
    ("Портфель ценных бумаг",      "Активы",  "Ликвидные активы",            "Портфель ценных бумаг", None               ),  # row 6
    ("Депозит в ЦБ",               "Активы",  "Ликвидные активы",            "Депозит в ЦБ",          None               ),  # row 7
    ("МБК, Обратное РЕПО",         "Активы",  "Ликвидные активы",            "МБК, Обратное РЕПО",    None               ),  # row 8
    ("Корр счета и касса",         "Активы",  "Ликвидные активы",            "Корр счета и касса",    None               ),  # row 9
    ("Кредитный портфель (гросс)", "Активы",  "Кредитный портфель (гросс)", None,                    None               ),  # row 10
    ("Клиенты ОПК*",               "Активы",  "Кредитный портфель (гросс)", "Клиенты ОПК*",          None               ),  # row 11
    ("Юридические лица",           "Активы",  "Кредитный портфель (гросс)", "Юридические лица",      None               ),  # row 12
    ("Физические лица",            "Активы",  "Кредитный портфель (гросс)", "Физические лица",       None               ),  # row 13
    ("Проблемные активы",          "Активы",  "Кредитный портфель (гросс)", "Проблемные активы",     None               ),  # row 14
    ("Резервы по КП",              "Активы",  "Резервы по КП",              None,                    None               ),  # row 15
    ("Неработающие активы",        "Активы",  "Неработающие активы",        None,                    None               ),  # row 16
    ("SWAP",                       "Активы",  "SWAP",                       None,                    None               ),  # row 17
    # ── ПАССИВЫ ───────────────────────────────────────────────────────────────
    ("ПАССИВЫ",                    "Пассивы", None,                          None,                    None               ),  # row 19
    ("МБК, Лоро, Прямое РЕПО",    "Пассивы", "МБК, Лоро, Прямое РЕПО",    None,                    None               ),  # row 20
    ("Клиентское привлечение",     "Пассивы", "Клиентское привлечение",     None,                    None               ),  # row 21
    ("Срочное**",                  "Пассивы", "Клиентское привлечение",     "Срочное",               None               ),  # row 22
    ("Средства ГОЗ",               "Пассивы", "Клиентское привлечение",     "Срочное",               "Средства ГОЗ"     ),  # row 23
    ("Юридические лица",           "Пассивы", "Клиентское привлечение",     "Срочное",               "Юридические лица" ),  # row 24
    ("Физические лица",            "Пассивы", "Клиентское привлечение",     "Срочное",               "Физические лица"  ),  # row 25
    ("ДВС**",                      "Пассивы", "Клиентское привлечение",     "ДВС",                   None               ),  # row 26
    ("Средства ГОЗ, МПТ",          "Пассивы", "Клиентское привлечение",     "ДВС",                   "Средства ГОЗ, МПТ"),  # row 27
    ("Юридические лица",           "Пассивы", "Клиентское привлечение",     "ДВС",                   "Юридические лица" ),  # row 28
    ("Физические лица",            "Пассивы", "Клиентское привлечение",     "ДВС",                   "Физические лица"  ),  # row 29
    ("Привлечение на ФР",          "Пассивы", "Привлечение на ФР",          None,                    None               ),  # row 30
    ("Собственные средства",       "Пассивы", "Собственные средства",       None,                    None               ),  # row 31
    ("Прочие обязательства",       "Пассивы", "Прочие обязательства",       None,                    None               ),  # row 32
    ("SWAP",                       "Пассивы", "SWAP",                       None,                    None               ),  # row 33
]

# Суффиксы для 12 столбцов B:M.
COLUMN_SUFFIXES = [
    "RUB_Volume", "RUB_%",          # B, C  — дата 1 | RUB
    "RUB_Volume", "RUB_%",          # D, E  — дата 2 | RUB
    "RUB",        "%",              # F, G  — Изменение за нед. | RUB
    "Summary_Volume", "Summary_%",  # H, I  — дата 1 | ИТОГО  ← brief берёт эти
    "Summary_Volume", "Summary_%",  # J, K  — дата 2 | ИТОГО  ← и эти
    "Summary_RUB",    "Summary_%",  # L, M  — Изменение за нед. | ИТОГО
]

OUT_COLUMNS = ["id", "date_", "axis_0", "axis_1", "axis_2", "axis_3", "axis_4", "value", "nversionid", "axis_5"]


class ChpdDataError(RuntimeError):
    """Ошибка чтения или обработки Excel-отчёта ЧПД."""


def build_long(df: pd.DataFrame, brief: bool = False) -> pd.DataFrame:
    """Переводит DataFrame из "широкого" вида в "длинный".

    Если brief=True, обрабатываются только столбцы *_Summary_Volume/*_Summary_%.
    """
    if len(INDEX_PATHS) < len(df):
        raise ChpdDataError(
            f"INDEX_PATHS содержит {len(INDEX_PATHS)} записей, но DataFrame имеет "
            f"{len(df)} строк. Добавьте недостающие элементы в INDEX_PATHS."
        )

    if brief:
        vol_cols = [c for c in df.columns if c.endswith("_Summary_Volume")]
        col_pairs = [(v, v.replace("_Summary_Volume", "_Summary_%")) for v in vol_cols]
    else:
        col_pairs = [(df.columns[i], df.columns[i + 1]) for i in range(0, len(df.columns), 2)]

    rows: List[dict] = []

    for idx, _leaf in enumerate(df.index):
        path = INDEX_PATHS[idx]
        leaf_name, ax1, *rest = path
        ax2 = rest[0] if len(rest) > 0 else None
        ax3 = rest[1] if len(rest) > 1 else None
        ax4_val = rest[2] if len(rest) > 2 else None

        # df.loc[leaf, col] возвращает Series при дублях в индексе ("Юридические
        # лица", "Физические лица", "SWAP" повторяются в Активах и Пассивах).
        # df.iloc[idx] всегда даёт нужную строку по позиции.
        row_data = df.iloc[idx]

        for vol_col, perc_col in col_pairs:
            vol_value = row_data[vol_col]
            if pd.isna(vol_value):
                vol_value = None
            if isinstance(vol_value, str) and vol_value.strip() == "-":
                vol_value = 0

            perc_value = row_data[perc_col]
            if pd.isna(perc_value):
                perc_value = None
            if isinstance(perc_value, str) and perc_value.strip() == "-":
                perc_value = 0

            base_date_str = vol_col.split("_")[0]
            # Заголовки RUB- и Summary-блоков (B/H, D/J) часто содержат одну и
            # ту же дату — pandas переименовывает повторный заголовок в
            # "20.05.2026.1" (mangle_dupe_cols), иначе pd.to_datetime не
            # распарсит эту дату. Суффикс — артефакт pandas, не часть даты.
            base_date_str = re.sub(r"\.\d+$", "", base_date_str)
            try:
                # dayfirst=True: даты в отчёте — формата ДД.ММ.ГГГГ (как и везде
                # в проекте); без этого pandas путает день/месяц для дат <=12.
                date_formatted = pd.to_datetime(base_date_str, dayfirst=True).strftime("%Y-%m-%d")
            except Exception:
                date_formatted = base_date_str

            if vol_value is not None and isinstance(vol_value, (int, float)):
                vol_value = int(round(vol_value))
            if perc_value is not None and isinstance(perc_value, (int, float)):
                perc_value = int(round(perc_value * 100))

            rows.append({
                "id": len(rows), "date_": date_formatted, "axis_0": "ЧПД",
                "axis_1": ax1, "axis_2": ax2, "axis_3": ax3, "axis_4": ax4_val,
                "value": vol_value, "nversionid": "", "axis_5": "млрд руб",
            })

            if perc_value is not None:
                rows.append({
                    "id": len(rows), "date_": date_formatted, "axis_0": "ЧПД",
                    "axis_1": ax1, "axis_2": ax2, "axis_3": ax3, "axis_4": ax4_val,
                    "value": perc_value, "nversionid": "", "axis_5": "%",
                })

    return pd.DataFrame(rows, columns=OUT_COLUMNS)


def build_report(file_path: Path) -> pd.DataFrame:
    """Полный цикл: чтение листа Table -> восстановление дат из merged-ячеек
    -> присвоение суффиксов -> перевод в длинный формат (только ИТОГО-колонки).
    """
    file_path = Path(file_path)
    if not file_path.exists():
        raise ChpdDataError(f"Файл не найден: {file_path}")

    try:
        data = pd.read_excel(
            file_path,
            sheet_name=SHEET_NAME,
            nrows=32,       # читает Excel rows 2-33 (row 1 = header)
            usecols="A:M",
            index_col=0,    # столбец A сразу становится индексом при чтении
        )
    except Exception as exc:
        raise ChpdDataError(f"Не удалось прочитать файл {file_path} (лист {SHEET_NAME!r}): {exc}") from exc

    # Восстанавливаем даты merged-ячеек (forward fill по Unnamed-столбцам B:M).
    cols = data.columns.to_frame()
    cols[0] = cols[0].mask(cols[0].astype(str).str.contains("Unnamed")).ffill()
    data.columns = pd.MultiIndex.from_frame(cols)
    df = data

    if all(isinstance(col, tuple) and len(col) == 1 for col in df.columns):
        df.columns = [col[0] for col in df.columns]

    if len(df.columns) != len(COLUMN_SUFFIXES):
        raise ChpdDataError(
            f"Ожидалось {len(COLUMN_SUFFIXES)} колонок данных (B:M), получено {len(df.columns)}. "
            "Проверьте структуру листа Table."
        )
    df.columns = [f"{str(col)}_{suf}" for col, suf in zip(df.columns, COLUMN_SUFFIXES)]

    # Пропускаем технические строки-подзаголовки (Excel rows 2-3).
    df = df.iloc[2:]
    # Удаляем полностью пустые строки (разделитель между Активами и Пассивами).
    df_clean = df.dropna(how="all")

    index_paths_trimmed = INDEX_PATHS[: len(df_clean)]
    if len(index_paths_trimmed) != len(df_clean):
        logger.warning(
            "INDEX_PATHS (%d) != строк в df_clean (%d). Проверьте INDEX_PATHS или структуру листа.",
            len(index_paths_trimmed), len(df_clean),
        )

    result = build_long(df_clean, brief=True)
    if result.empty:
        raise ChpdDataError("Итоговый DataFrame пуст — не найдено ни одной строки данных.")

    logger.info("Обработано строк входа: %d, строк на выходе: %d", len(df_clean), len(result))
    return result


def save_report(df: pd.DataFrame, output_path: Path) -> Path:
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        df.to_csv(output_path, index=False, encoding="utf-8")
    except PermissionError as exc:
        raise ChpdDataError(
            f"Нет доступа для записи в {output_path} (файл открыт в другой программе?): {exc}"
        ) from exc
    logger.info("Отчёт сохранён: %s (%d строк)", output_path, len(df))
    return output_path
