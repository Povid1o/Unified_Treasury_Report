"""Лимиты по типам портфелей из отдельной выгрузки «Состояние лимитов».

Зачем. Раньше лист fact_limit заполнялся руками и переносился из выпуска в
выпуск как есть. Теперь лимиты приходят файлом на отчётную дату — отдельной
выгрузкой вида «Состояние лимитов на дату 21_09_2026 - Результат.xlsx».

Структура выгрузки: на 6-й строке заголовки колонок (нужны «Тип лимита» и
«Лимит сверху»), под ними ПУСТАЯ строка, дальше таблица. Ищем шапку по
наличию колонки, а не по номеру строки: номер поедет от любой правки формы.

Два правила, которых нет ни в одной другой выгрузке:

1. Торговый портфель называется в этом файле «Облигации» — соответствие имён
   задаётся таблицей псевдонимов (настройка, а не константа: формулировки в
   выгрузках меняются чаще, чем код).
2. AFS встречается ДВАЖДЫ. Берётся строка с БОЛЬШИМ лимитом, меньшая
   игнорируется — это оговорено казначейством.

Границы зон лимитом не приходят: они считаются процентами от него по
таблице ZONE_PERCENTS.
"""
import datetime as dt
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd

BASE_DIR = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(BASE_DIR))

import config  # noqa: E402
from common import excel_io  # noqa: E402
from reports.portfolio_dynamics.etl import (  # noqa: E402
    KNOWN_PORTFOLIO_TYPES, LIMIT_COLUMNS, PROBE_SUFFIXES, PortfolioDynamicsError,
    canonical_type, logger, parse_number, parse_nested_limits, parse_type_parents,
)

COL_LIMIT_TYPE = "Тип лимита"
COL_LIMIT_AMOUNT = "Лимит сверху"
# Необязательная колонка: сколько лимита ещё не выбрано по данным самой системы
# лимитов. Нужна для сверки с объёмом, посчитанным по выгрузке позиций.
COL_LIMIT_REMAINING = "Остаток лимита сверху"

# Границы зон — доли от установленного лимита (limit_amount). Значения
# ЗАФИКСИРОВАНЫ казначейством и намеренно НЕ вынесены в настройки: это не
# параметр запуска, а согласованная политика, одна и та же для всех запусков и
# всех пользователей. Менять — только правкой этой таблицы, вместе с
# согласованием; тест test_zone_percents_match_the_agreed_table сверяет её с
# исходными цифрами.
#
#   Торговый (TSS)  красная 95,10 %  жёлтая 85,09 %  зелёная 78,08 %
#   AFS             красная 96,77 %  жёлтая 93,55 %  зелёная 90,32 %
#   HTM             красная 96,43 %  жёлтая 92,86 %  зелёная 89,29 %
#   HTM_KUAP        красная 95,00 %  жёлтая 85,00 %  зелёная 78,00 %
#
# red_max_util здесь НИЖЕ лимита, в отличие от эталонного шаблона, где он был
# равен лимиту.
#            (зелёная, жёлтая, красная)
ZONE_PERCENTS: Dict[str, Tuple[float, float, float]] = {
    "TSS":      (0.7808, 0.8509, 0.9510),  # торговый портфель, в файле «Облигации»
    "AFS":      (0.9032, 0.9355, 0.9677),
    "HTM":      (0.8929, 0.9286, 0.9643),
    "HTM_KUAP": (0.7800, 0.8500, 0.9500),
}
# Для типа, которого нет в таблице выше: те же доли, что были в эталонном
# шаблоне (70/90/100 % лимита). Попадание сюда логируется — значит, завели
# новый тип и проценты для него надо согласовать.
DEFAULT_ZONE_PERCENTS = (0.70, 0.90, 1.00)


def parse_aliases(raw: str) -> Dict[str, str]:
    """«Облигации=TTS, Прочее=HTM» -> {нормализованное имя: тип портфеля}."""
    aliases: Dict[str, str] = {}
    for item in str(raw or "").split(","):
        name, _, portfolio_type = item.partition("=")
        if name.strip() and portfolio_type.strip():
            aliases[excel_io.normalize_label(name)] = portfolio_type.strip().upper()
    return aliases


def resolve_type(raw_name, aliases: Dict[str, str]) -> str:
    """Имя строки из файла лимитов -> тип портфеля.

    В выгрузке типы названы «Облигации AFS», «Облигации HTM» и просто
    «Облигации» (это торговый). Точного псевдонима на каждую формулировку не
    напасёшься, поэтому: сначала точное соответствие из настроек, затем поиск
    ИЗВЕСТНОГО ТИПА внутри названия (самого длинного — чтобы «Облигации
    HTM_KUAP» не свелось к HTM), и только потом имя как есть.
    """
    normalized = excel_io.normalize_label(raw_name)
    if normalized in aliases:
        return canonical_type(aliases[normalized])

    upper = str(raw_name).upper()
    inside = [t for t in KNOWN_PORTFOLIO_TYPES if t in upper]
    if inside:
        return max(inside, key=len)

    return canonical_type(raw_name)


# Дата в шапке файла лимитов: «на дату 21.09.2026» / «21_09_2026».
_DATE_IN_TEXT = re.compile(r"(\d{2})[._-](\d{2})[._-](\d{4})")
_PROBE_ROWS = 20


def probe_limits_date(path: Path) -> Optional[dt.date]:
    """Дата выгрузки лимитов по СОДЕРЖИМОМУ книги. None — файл не тот.

    Нужна, когда файл переименовали и имя под шаблон больше не подходит.
    Требуется И колонка «Тип лимита», И дата в верхних строках: без первого
    признака за выгрузку лимитов принялся бы любой файл с датой в шапке, без
    второго — дату брать было бы неоткуда, а угадывать её по времени файла
    нельзя: молча устаревший лимит хуже отсутствующего.
    """
    path = Path(path)
    if path.suffix.lower() not in PROBE_SUFFIXES:
        return None
    wanted = excel_io.normalize_label(COL_LIMIT_TYPE)
    try:
        from openpyxl import load_workbook
        workbook = load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return None
    try:
        for worksheet in workbook.worksheets:
            found_date = None
            has_column = False
            for row in worksheet.iter_rows(max_row=_PROBE_ROWS, values_only=True):
                for value in row:
                    if value is None:
                        continue
                    if isinstance(value, dt.datetime):
                        found_date = found_date or value.date()
                        continue
                    if isinstance(value, dt.date):
                        found_date = found_date or value
                        continue
                    text = str(value)
                    if not has_column and excel_io.normalize_label(text) == wanted:
                        has_column = True
                    if found_date is None:
                        match = _DATE_IN_TEXT.search(text)
                        if match:
                            try:
                                found_date = dt.date(int(match.group(3)), int(match.group(2)),
                                                     int(match.group(1)))
                            except ValueError:
                                pass
            if has_column and found_date is not None:
                return found_date
    except Exception:
        return None
    finally:
        try:
            workbook.close()
        except Exception:
            pass
    return None


def _find_header(path: Path) -> Tuple[str, pd.DataFrame, int, Dict[str, int]]:
    """Лист и строка с колонкой «Тип лимита»; рядом — индексы нужных колонок."""
    wanted = {
        COL_LIMIT_TYPE: excel_io.normalize_label(COL_LIMIT_TYPE),
        COL_LIMIT_AMOUNT: excel_io.normalize_label(COL_LIMIT_AMOUNT),
        COL_LIMIT_REMAINING: excel_io.normalize_label(COL_LIMIT_REMAINING),
    }
    seen: List[str] = []
    for sheet_name, matrix in excel_io.iter_sheet_matrices(path):
        seen.append(sheet_name)
        if matrix.empty:
            continue
        for row_idx in range(len(matrix)):
            normalized = [excel_io.normalize_label(v) for v in matrix.iloc[row_idx]]
            if wanted[COL_LIMIT_TYPE] not in normalized:
                continue
            columns = {label: normalized.index(norm)
                       for label, norm in wanted.items() if norm in normalized}
            if COL_LIMIT_AMOUNT not in columns:
                raise PortfolioDynamicsError(
                    f"В файле лимитов {path.name} (лист {sheet_name!r}) есть колонка "
                    f"«{COL_LIMIT_TYPE}», но нет «{COL_LIMIT_AMOUNT}» — брать лимит неоткуда."
                )
            return sheet_name, matrix, row_idx, columns

    raise PortfolioDynamicsError(
        f"В файле лимитов {path.name} не найден лист с колонкой «{COL_LIMIT_TYPE}». "
        f"Листы в файле: {seen}."
    )


def parse_limits_file(path: Path, scale: Optional[float] = None,
                      aliases: Optional[Dict[str, str]] = None) -> pd.DataFrame:
    """Разбирает выгрузку лимитов -> DataFrame[portfolio_type, limit_amount].

    Дубли типа схлопываются по МАКСИМУМУ: в выгрузке AFS приходит двумя
    строками, и казначейство ориентируется на большую из них.
    """
    path = Path(path)
    if not path.exists():
        raise PortfolioDynamicsError(f"Файл лимитов не найден: {path}")

    scale = config.PORTFOLIO_DYNAMICS_LIMIT_SCALE if scale is None else scale
    aliases = parse_aliases(config.PORTFOLIO_DYNAMICS_LIMIT_ALIASES) if aliases is None else aliases

    sheet_name, matrix, header_row, columns = _find_header(path)
    type_col = columns[COL_LIMIT_TYPE]
    amount_col = columns[COL_LIMIT_AMOUNT]
    remaining_col = columns.get(COL_LIMIT_REMAINING)
    if remaining_col is None:
        logger.info("В файле лимитов нет колонки «%s» — сверка использования пропущена.",
                    COL_LIMIT_REMAINING)

    rows: List[Tuple[str, str, float, Optional[float]]] = []
    skipped_without_amount = 0
    for row_idx in range(header_row + 1, len(matrix)):
        raw_name = matrix.iloc[row_idx].iloc[type_col]
        normalized = excel_io.normalize_label(raw_name)
        if not normalized:
            continue  # та самая пустая строка под шапкой, и любые другие
        amount = parse_number(matrix.iloc[row_idx].iloc[amount_col])
        if amount is None:
            skipped_without_amount += 1
            continue
        portfolio_type = resolve_type(raw_name, aliases)
        remaining = (parse_number(matrix.iloc[row_idx].iloc[remaining_col])
                     if remaining_col is not None else None)
        rows.append((portfolio_type, str(raw_name).strip(), amount / scale,
                     None if remaining is None else remaining / scale))

    if not rows:
        raise PortfolioDynamicsError(
            f"В файле лимитов {path.name} (лист {sheet_name!r}, шапка в строке "
            f"{header_row + 1}) не нашлось ни одной строки с заполненным "
            f"«{COL_LIMIT_AMOUNT}»."
        )

    frame = pd.DataFrame(
        rows, columns=["portfolio_type", "source_name", "limit_amount", "remaining_amount"])
    duplicated = frame[frame.duplicated("portfolio_type", keep=False)]
    best = (frame.sort_values("limit_amount", ascending=False)
                 .drop_duplicates("portfolio_type", keep="first")
                 .sort_values("portfolio_type")
                 .reset_index(drop=True))

    for portfolio_type, group in duplicated.groupby("portfolio_type"):
        chosen = group["limit_amount"].max()
        dropped = sorted(v for v in group["limit_amount"] if v != chosen)
        logger.info(
            "Лимиты: тип %s встречается в файле %d раз — взят больший лимит %s, "
            "меньшие отброшены (%s).",
            portfolio_type, len(group), f"{chosen:,.0f}",
            ", ".join(f"{v:,.0f}" for v in dropped) or "—",
        )

    logger.info(
        "Файл лимитов %s: лист %r, шапка в строке %d, строк с лимитом %d, "
        "типов после схлопывания дублей %d%s",
        path.name, sheet_name, header_row + 1, len(frame), len(best),
        f", строк без суммы пропущено {skipped_without_amount}" if skipped_without_amount else "",
    )
    return best[["portfolio_type", "limit_amount", "remaining_amount"]]


def _split_nested(from_file: Dict[str, float], known_types: set) -> Dict[str, float]:
    """Делит совокупный лимит между объемлющим типом и вложенными.

    В выгрузке лимит на HTM — общий: он покрывает и HTM, и HTM_KUAP. Сколько из
    него отведено КУАП, выгрузка не знает, это решение казначейства — оно
    задаётся настройкой «Сколько отдано вложенным типам» (HTM_KUAP=100).
    Тогда КУАП получает лимит 100, а объемлющий — остаток: 900 - 100 = 800.

    Подлимит не задан — ничего не делим: объём вложенного типа в этом случае
    складывается с объемлющим (см. etl.aggregated_parents) и сравнивается с
    общим лимитом целиком.
    """
    allocations = parse_nested_limits()
    if not allocations:
        return from_file

    parents = parse_type_parents()
    result = dict(from_file)
    for child, allocated in allocations.items():
        parent = parents.get(child)
        if parent is None:
            logger.warning(
                "Подлимит задан для %s, но этот тип не объявлен вложенным ни в какой "
                "другой (настройка «Вложенность типов»). Лимит %s взят как есть, из "
                "общего лимита он не вычитается.", child, f"{allocated:,.0f}",
            )
            result[child] = allocated
            continue

        total = result.get(parent)
        if total is None:
            logger.warning(
                "Подлимит %s=%s задан, но лимита объемлющего типа %s в файле нет — "
                "вычитать не из чего. Подлимит установлен, остаток не рассчитан.",
                child, f"{allocated:,.0f}", parent,
            )
            result[child] = allocated
            continue

        if allocated > total:
            raise PortfolioDynamicsError(
                f"Подлимит {child} = {allocated:,.0f} больше совокупного лимита "
                f"{parent} = {total:,.0f}. Поправьте настройку «Сколько отдано "
                f"вложенным типам» либо проверьте файл лимитов."
            )

        if child in from_file:
            logger.info(
                "Подлимит %s взят из настроек (%s), строка этого типа в файле лимитов "
                "(%s) не используется.", child, f"{allocated:,.0f}",
                f"{from_file[child]:,.0f}",
            )
        result[child] = allocated
        result[parent] = total - allocated
        logger.info(
            "Совокупный лимит %s = %s разделён: %s -> %s, остальной %s -> %s.",
            parent, f"{total:,.0f}", child, f"{allocated:,.0f}",
            parent, f"{total - allocated:,.0f}",
        )
    return result


def zones_for(portfolio_type: str, limit_amount: float) -> Tuple[float, float, float]:
    """Границы зон как доли от лимита (зелёная, жёлтая, красная)."""
    percents = ZONE_PERCENTS.get(canonical_type(portfolio_type))
    if percents is None:
        percents = DEFAULT_ZONE_PERCENTS
        logger.warning(
            "Для типа %s не задано процентов зон — взяты значения по умолчанию "
            "(%.0f/%.0f/%.0f %% лимита). Согласуйте проценты и добавьте тип в "
            "ZONE_PERCENTS (reports/portfolio_dynamics/limits.py).",
            portfolio_type, *[p * 100 for p in DEFAULT_ZONE_PERCENTS],
        )
    return tuple(round(limit_amount * p, 2) for p in percents)


def _remaining_for(portfolio_type: str, limit_amount: float,
                   limits_as_in_file: Dict[str, float],
                   remaining_by_type: Dict[str, Optional[float]],
                   split_types: List[str]) -> Optional[float]:
    """Остаток лимита для строки fact_limit. None — записывать нечего.

    Остаток пишется, только когда записанный лимит совпал с тем, что стоит в
    файле. Если лимит поделён на подлимиты («Сколько отдано вложенным типам»),
    остаток из файла относится к СОВОКУПНОМУ лимиту, и рядом с долей от него
    он означал бы неправду — а неправда в колонке хуже пустой колонки.
    """
    remaining = remaining_by_type.get(portfolio_type)
    if remaining is None or pd.isna(remaining):
        return None
    in_file = limits_as_in_file.get(portfolio_type)
    if in_file is None or round(in_file, 2) != round(limit_amount, 2):
        split_types.append(portfolio_type)
        return None
    return round(float(remaining), 2)


def build_fact_limit(parsed: pd.DataFrame, known_types: List[str],
                     previous: pd.DataFrame, business_date, source_name: str) -> pd.DataFrame:
    """Собирает лист fact_limit: лимиты из файла + границы зон по процентам.

    known_types — типы, которые вообще есть в отчёте (из dim_portfolio и из
    предыдущего выпуска). Строки файла, не сводящиеся ни к одному из них,
    в fact_limit не попадают: выгрузка лимитов ведётся не только по портфельным
    типам, и тянуть в реестр всё подряд нельзя.

    Тип, которого в файле нет, сохраняет лимит из предыдущего выпуска — иначе
    один неполный файл обнулил бы уже согласованные лимиты.
    """
    # Типы, которые вообще существуют в этом отчёте: из справочника и из
    # предыдущего выпуска.
    known_upper = {canonical_type(t) for t in known_types if pd.notna(t) and str(t).strip()}
    # Согласованные типы узнаются в файле, даже если портфелей такого типа
    # сегодня нет, — но САМИ ПО СЕБЕ в реестр не добавляются: тип без лимита,
    # без портфелей и без строки в файле — это пустая строка, из-за которой
    # CHK_19 («лимиты <= 0») падал бы на каждом запуске.
    recognised = known_upper | set(ZONE_PERCENTS)

    from_file = {canonical_type(row.portfolio_type): float(row.limit_amount)
                 for row in parsed.itertuples()}
    # Остаток приходит той же строкой файла и относится к ТОМУ ЖЕ лимиту, что
    # в ней указан. Запоминаем и лимит до дележа на подлимиты: после него
    # остаток к записанной сумме уже не относится (см. ниже).
    remaining_by_type = {}
    if "remaining_amount" in parsed.columns:
        remaining_by_type = {canonical_type(row.portfolio_type): row.remaining_amount
                             for row in parsed.itertuples()}
    limits_as_in_file = dict(from_file)
    from_file = _split_nested(from_file, recognised)
    matched = {t: v for t, v in from_file.items() if t in recognised}
    ignored = sorted(set(from_file) - set(matched))
    if ignored:
        logger.info(
            "Лимиты: строк файла не отнесены ни к одному типу портфелей и пропущены "
            "(%d): %s. Известные типы: %s. Если какая-то из них всё же нужна — "
            "добавьте соответствие в настройку «Соответствия типов в файле лимитов».",
            len(ignored), ", ".join(ignored), ", ".join(sorted(recognised)),
        )

    previous_by_type = {}
    if previous is not None and not previous.empty:
        previous_by_type = {canonical_type(row.portfolio_type): row
                            for row in previous.itertuples()}

    # В реестр идут типы, у которых есть хоть что-то: лимит из файла, портфели
    # в справочнике или строка в предыдущем выпуске.
    types_in_report = sorted(set(matched) | known_upper | set(previous_by_type))

    records = []
    split_types = []
    for portfolio_type in types_in_report:
        if portfolio_type in matched:
            limit_amount = matched[portfolio_type]
            green, yellow, red = zones_for(portfolio_type, limit_amount)
            records.append({
                "portfolio_type": portfolio_type, "limit_amount": round(limit_amount, 2),
                "green_max_util": green, "yellow_max_util": yellow, "red_max_util": red,
                "valid_from": business_date, "updated_by": source_name,
                "limit_remaining": _remaining_for(
                    portfolio_type, limit_amount, limits_as_in_file,
                    remaining_by_type, split_types),
            })
            continue

        kept = previous_by_type.get(portfolio_type)
        if kept is not None:
            records.append({column: getattr(kept, column) for column in LIMIT_COLUMNS})
            continue

        records.append({
            "portfolio_type": portfolio_type, "limit_amount": 0,
            "green_max_util": 0, "yellow_max_util": 0, "red_max_util": 0,
            "valid_from": business_date, "updated_by": None, "limit_remaining": None,
        })

    if split_types:
        logger.info(
            "Лимиты: «%s» не записан для типов %s — их лимит поделён на подлимиты, "
            "а остаток в файле указан к совокупному лимиту и к записанной сумме "
            "больше не относится.", COL_LIMIT_REMAINING, ", ".join(split_types),
        )

    missing = sorted(t for t in types_in_report if t not in matched)
    if missing:
        logger.warning(
            "Лимиты: в файле нет строк для типов %s — для них сохранены значения из "
            "предыдущего выпуска (или нули, если их там не было).", ", ".join(missing),
        )

    frame = pd.DataFrame(records, columns=LIMIT_COLUMNS)
    _log_changes(frame, previous_by_type)
    return frame


def check_utilisation(parsed: pd.DataFrame, volumes: Dict[str, float]) -> List[str]:
    """Сверяет использование лимита по двум независимым источникам.

    Файл лимитов знает, сколько лимита ещё не выбрано («Остаток лимита сверху»),
    то есть косвенно — сколько уже занято. Отчёт считает занятое сам, из
    выгрузки позиций. Расхождение означает, что либо типы сопоставлены не так,
    либо источники разошлись, — и это надо увидеть, а не узнать от казначейства.

    Возвращает список расхождений (для лога и для тестов).
    """
    if parsed is None or parsed.empty or "remaining_amount" not in parsed.columns:
        return []

    problems: List[str] = []
    for row in parsed.itertuples():
        remaining = getattr(row, "remaining_amount", None)
        if remaining is None or pd.isna(remaining):
            continue
        portfolio_type = canonical_type(row.portfolio_type)
        ours = volumes.get(portfolio_type)
        if ours is None:
            continue
        implied = float(row.limit_amount) - float(remaining)
        logger.info(
            "Лимиты: %s — лимит %s, остаток по файлу %s, значит занято %s; "
            "по выгрузке позиций занято %s.",
            portfolio_type, f"{row.limit_amount:,.0f}", f"{remaining:,.0f}",
            f"{implied:,.0f}", f"{ours:,.0f}",
        )
        reference = max(abs(implied), abs(ours))
        if reference and abs(implied - ours) / reference > config.PORTFOLIO_DYNAMICS_TOLERANCE:
            problems.append(
                f"{portfolio_type}: по файлу лимитов занято {implied:,.0f}, "
                f"по выгрузке позиций {ours:,.0f}"
            )

    if problems:
        logger.warning(
            "Использование лимита по файлу и по выгрузке позиций расходится (%d): %s. "
            "Обычно это значит, что строка файла отнесена не к тому типу, либо "
            "выгрузки сделаны на разные моменты.",
            len(problems), "; ".join(problems),
        )
    return problems


def _log_changes(frame: pd.DataFrame, previous_by_type: dict) -> None:
    """Пишет в лог, у каких типов лимит изменился относительно прошлого выпуска.

    Лист fact_limit человек считает своим, поэтому молча переписывать его
    нельзя: изменение лимита должно быть видно в логе запуска.
    """
    for row in frame.itertuples():
        old = previous_by_type.get(str(row.portfolio_type).upper())
        if old is None:
            logger.info("Лимиты: %s — установлен %s", row.portfolio_type,
                        f"{row.limit_amount:,.0f}")
            continue
        if float(old.limit_amount or 0) != float(row.limit_amount or 0):
            logger.warning(
                "Лимиты: %s изменён — было %s, стало %s (из файла лимитов).",
                row.portfolio_type, f"{float(old.limit_amount or 0):,.0f}",
                f"{float(row.limit_amount or 0):,.0f}",
            )
