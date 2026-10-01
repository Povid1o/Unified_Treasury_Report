"""Единица и число знаков выходных значений — настройка «Округление» у каждого отчёта.

Зачем. Округление и единицы были зашиты в код каждого отчёта по-своему (ЧПД —
до целых млрд, NIM — 2 знака, «Динамика» — млрд RUB…), и любая просьба
«покажите в миллионах» означала правку кода. Здесь у каждого отчёта — список
его показателей (CATALOG), и по каждому в настройках можно задать единицу и
число знаков после запятой.

Главное правило: пустая настройка — ровно прежний вывод, байт в байт. Поэтому
отчёт спрашивает правило показателя (rule) и, если оно по умолчанию
(Rule.is_default), идёт СВОИМ прежним путём, не трогая значение вовсе. Только
изменённый показатель проходит через Rule.apply.

Три вида показателей:

* money    — деньги в известной единице (руб, млн, млрд). Единица выбирается
             абсолютно: «млн руб», «тыс. руб»… Подпись единицы в выгрузке,
             если она там есть, меняется вместе со значением.
* relative — величина, единица которой отчёту неизвестна (ОВП — в единицах
             исходного файла) или не денежная (Open QTY, шт). Задаётся
             делителем: ÷10 … ÷1 000 000 000.
* none     — проценты, ставки, дюрации, пункты: делить их бессмысленно,
             настраивается только число знаков.

Знаки «как сейчас» означают прежнее округление отчёта (его применяет сам
отчёт уже к пересчитанному значению): ЧПД в миллионах так и останется «до
целых», только целых миллионов.

Модуль без импорта config на уровне модуля: его импортирует settings.py
(проверка значений), а config импортирует settings.
"""
import math
from dataclasses import dataclass
from typing import Any, Callable, Dict, List, Optional, Tuple

MONEY = "money"
RELATIVE = "relative"
NONE = "none"

MAX_DECIMALS = 10

# Абсолютные единицы денег: показатель степени десяти -> слово.
MONEY_SCALES: List[int] = [0, 3, 6, 9]
_MONEY_WORDS = {0: "", 3: "тыс.", 6: "млн", 9: "млрд"}
# Делители для relative: показатель степени -> подпись в диалоге и префикс единицы.
RELATIVE_SCALES: List[int] = [1, 2, 3, 6, 9]
_DIVISORS = {1: "÷10", 2: "÷100", 3: "÷1 000", 6: "÷1 000 000", 9: "÷1 000 000 000"}
_RELATIVE_TITLES = {1: "÷10", 2: "÷100", 3: "÷1 000 (тысячи)", 6: "÷1 000 000 (миллионы)",
                    9: "÷1 000 000 000 (миллиарды)"}
_RELATIVE_PREFIX = {1: "десятки", 2: "сотни", 3: "тыс.", 6: "млн", 9: "млрд"}


@dataclass(frozen=True)
class Category:
    """Показатель отчёта, которому можно задать единицу и знаки."""

    key: str
    label: str
    now: str                    # как округляется сейчас — для экрана настроек
    kind: str = NONE
    base: int = 0               # money: в какой единице показатель сейчас (степень десяти)
    unit: Optional[str] = None  # подпись единицы в выгрузке; None — подписи нет
    currency: str = "руб"       # money: как пишется валюта в подписи («руб» или «RUB»)
    note: str = ""

    def money_unit(self, scale: int) -> str:
        word = _MONEY_WORDS[scale]
        return f"{word} {self.currency}".strip()

    def scale_title(self, scale: Optional[int]) -> str:
        """Как показать выбранную единицу на экране настроек."""
        if self.kind == MONEY:
            return self.money_unit(self.base if scale is None else scale)
        if self.kind == RELATIVE:
            if not scale:
                return "как в источнике" if self.unit is None else self.unit
            if self.unit:
                return f"{_RELATIVE_PREFIX[scale]} {self.unit} ({_DIVISORS[scale]})"
            return _RELATIVE_TITLES[scale]
        return self.unit or "—"

    def scales(self) -> List[int]:
        if self.kind == MONEY:
            return list(MONEY_SCALES)
        if self.kind == RELATIVE:
            return [0] + list(RELATIVE_SCALES)
        return []

    def default_scale(self) -> int:
        return self.base if self.kind == MONEY else 0


def _c(key, label, now, kind=NONE, **kw) -> Category:
    return Category(key=key, label=label, now=now, kind=kind, **kw)


_DYNAMICS_CSV_MONEY = dict(kind=MONEY, base=9, unit="млрд RUB", currency="RUB")

# Ключ верхнего уровня — ключ раздела настроек (settings.GROUPS).
CATALOG: Dict[str, List[Category]] = {
    "ofz": [
        _c("yield", "Доходность ОФЗ", "2 знака", unit="%"),
    ],
    "ovp": [
        _c("curr_balance", "Остаток (curr_balance)", "без округления", RELATIVE,
           note="в единицах исходного файла ОВП"),
        _c("reserve_msfo", "Резерв МСФО (reserve_msfo)", "без округления", RELATIVE,
           note="в единицах исходного файла ОВП"),
    ],
    "balance_struct": [
        _c("value", "Значения плана фондирования", "до целых", MONEY, base=9,
           note="подписи единицы в выгрузке нет — исходная таблица в млрд"),
    ],
    "chpd": [
        _c("volume", "Объёмы", "до целых", MONEY, base=9, unit="млрд руб"),
        _c("share", "Доли, %", "до целых", unit="%"),
    ],
    "nim": [
        _c("nim", "NIM", "2 знака"),
        _c("result", "% результат", "2 знака"),
        _c("assets", "% Активы", "2 знака"),
        _c("liabilities", "% Пассивы", "2 знака"),
    ],
    "transfert": [
        _c("rate", "Трансфертная ставка", "без округления", unit="%"),
    ],
    "portfolio_dynamics": [
        _c("flat_t0", "CSV: Объём T0", "9 знаков (до рубля)", **_DYNAMICS_CSV_MONEY),
        _c("flat_t7", "CSV: Объём T-7", "9 знаков (до рубля)", **_DYNAMICS_CSV_MONEY),
        _c("flat_delta", "CSV: Изменение объёма", "9 знаков (до рубля)", **_DYNAMICS_CSV_MONEY),
        _c("flat_delta_pct", "CSV: Изменение объёма, %", "10 знаков", unit="%"),
        _c("flat_duration", "CSV: Дюрация текущая", "без округления", unit="лет"),
        _c("flat_duration_target", "CSV: Дюрация-КУАП", "без округления", unit="лет"),
        _c("flat_duration_gap", "CSV: Изменение дюрации", "без округления", unit="лет"),
        _c("flat_type_volume", "CSV: Объём типа (и разовая история)", "9 знаков (до рубля)",
           **_DYNAMICS_CSV_MONEY),
        _c("flat_type_limit", "CSV: Лимит типа", "9 знаков (до рубля)", **_DYNAMICS_CSV_MONEY),
        _c("xlsx_amounts", "xlsx: все суммы", "2 знака в отображении", MONEY, base=9,
           unit="млрд RUB", currency="RUB",
           note="единица одна на всю книгу (формулы сравнивают лимиты с объёмами); "
                "знаки — формат ячеек и значения view_monitor_raw, на машинных листах "
                "числа хранятся полностью"),
        _c("xlsx_percent", "xlsx: проценты (Δ%, утилизация)", "1 знак в отображении",
           note="формат ячеек и значения view_monitor_raw"),
        _c("xlsx_duration", "xlsx: дюрации", "2 знака в отображении",
           note="формат ячеек и значения view_monitor_raw"),
    ],
    "portfolio_report": [
        _c("open_qty", "Open QTY", "без округления", RELATIVE, unit="шт"),
        _c("open_qty_change", "Изменение Open QTY", "без округления", RELATIVE, unit="шт"),
        _c("net_value", "Чистая стоимость", "без округления", MONEY, base=0, unit="руб"),
        _c("total_pl", "Total Full PL with Funding", "без округления", MONEY, base=0, unit="руб"),
        _c("dv01", "DV01", "без округления", MONEY, base=0, unit="руб"),
        _c("yield", "Yield", "без округления", unit="%"),
        _c("rgbi", "RGBI", "без округления", unit="пункты"),
        _c("ruonia", "RUONIA", "без округления", unit="%"),
        _c("rwa", "RWA", "без округления", MONEY, base=0, unit="руб"),
    ],
}

CATEGORIES: Dict[str, Dict[str, Category]] = {
    report: {c.key: c for c in items} for report, items in CATALOG.items()
}


class RoundingError(ValueError):
    """Некорректное значение настройки округления."""


# ════════════════════════════════════════════════════════════════════════════
# Правило
# ════════════════════════════════════════════════════════════════════════════
@dataclass(frozen=True)
class Rule:
    """Что сделать со значением показателя: умножить на 10**shift и округлить."""

    shift: int = 0
    decimals: Optional[int] = None  # None — как сейчас (прежнее округление отчёта)
    unit: Optional[str] = None      # новая подпись единицы; None — прежняя

    @property
    def is_default(self) -> bool:
        return self.shift == 0 and self.decimals is None

    def scale(self, value: float) -> float:
        # Делим на целое 10**k, а не умножаем на 10**-k: 0.001 непредставимо
        # точно, и 412 345 678 * 0.001 дал бы хвост в последнем знаке.
        if self.shift >= 0:
            return value * (10 ** self.shift)
        return value / (10 ** -self.shift)

    def apply(self, value: Any, current: Optional[Callable[[float], Any]] = None) -> Any:
        """Значение по правилу. current — прежнее округление отчёта (для «как сейчас»).

        Не число (текст, пусто, NaN) возвращается как есть: решать, что с ним
        делать, — дело отчёта, а не настройки единиц.
        """
        number = _as_number(value)
        if number is None:
            return value
        if self.is_default:
            return current(number) if current else value
        scaled = self.scale(number)
        if self.decimals is None:
            return current(scaled) if current else scaled
        if self.decimals == 0:
            return int(round(scaled))
        return round(scaled, self.decimals)

    def label(self, current_label: Optional[str]) -> Optional[str]:
        return self.unit if self.unit is not None else current_label


DEFAULT_RULE = Rule()


def frame_with_values(rows: List[dict], columns: List[str], column: str = "value"):
    """DataFrame, в котором column хранит числа как есть — целые остаются целыми.

    pandas сводит колонку, где рядом целые и дробные, к float, и в CSV целое
    уходит как «3.0»: «0 знаков» у одного показателя превращались бы в «.0»,
    стоит другому показателю той же колонки получить дробные знаки. Отчёты
    зовут это только при изменённой настройке — по умолчанию DataFrame
    строится прежним способом.
    """
    import pandas as pd
    frame = pd.DataFrame(rows, columns=columns)
    frame[column] = pd.Series([row.get(column) for row in rows], index=frame.index, dtype=object)
    return frame


def _as_number(value: Any) -> Optional[float]:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)) or hasattr(value, "item"):
        try:
            number = float(value)
        except (TypeError, ValueError):
            return None
        return None if math.isnan(number) else number
    return None


def make_rule(category: Category, entry: Optional[dict]) -> Rule:
    """Правило из записи настройки ({"scale": …, "decimals": …}) для показателя."""
    if not entry:
        return DEFAULT_RULE
    decimals = entry.get("decimals")
    scale = entry.get("scale")
    if category.kind == MONEY and scale is not None and scale != category.base:
        unit = category.money_unit(scale) if category.unit is not None else None
        return Rule(shift=category.base - scale, decimals=decimals, unit=unit)
    if category.kind == RELATIVE and scale:
        unit = f"{_RELATIVE_PREFIX[scale]} {category.unit}" if category.unit else None
        return Rule(shift=-scale, decimals=decimals, unit=unit)
    return Rule(decimals=decimals)


def rule(report: str, key: str) -> Rule:
    """Действующее правило показателя key отчёта report (по текущим настройкам)."""
    import config  # здесь, а не наверху: config импортирует settings, settings — этот модуль
    category = CATEGORIES[report][key]
    entry = (getattr(config, "ROUNDING", None) or {}).get(report, {}).get(key)
    return make_rule(category, entry)


def scale_of_label(report: str, key: str, label: Optional[str]) -> Optional[int]:
    """Степень десяти, в которой записано значение с подписью label.

    Нужна там, где отчёт читает свой предыдущий выпуск (Open QTY в «Отчёте по
    портфелям»): вчерашнее значение могло быть записано в другой единице, и
    сравнивать его нужно, вернув в базовую. None — подпись не узнана.
    Возвращает сдвиг к базовой единице: value * 10**shift.
    """
    category = CATEGORIES[report][key]
    text = (label or "").strip()
    if category.kind == MONEY:
        for scale in MONEY_SCALES:
            if category.money_unit(scale) == text:
                return scale - category.base
        return None
    if category.kind == RELATIVE and category.unit:
        if text == category.unit:
            return 0
        for scale, prefix in _RELATIVE_PREFIX.items():
            if text == f"{prefix} {category.unit}":
                return scale
        return None
    return 0 if text == (category.unit or "") else None


# ════════════════════════════════════════════════════════════════════════════
# Значение настройки
# ════════════════════════════════════════════════════════════════════════════
def parse(report: str, raw: Any, key: str = "rounding") -> Dict[str, dict]:
    """Проверяет и нормализует значение настройки отчёта report.

    Записи «как сейчас» выбрасываются, чтобы в settings.json не копились
    пустые хвосты. Неизвестные показатели игнорируются молча: показатель могли
    переименовать в новой версии, и старый settings.json не должен ломать
    запуск — так же поступает settings.py с неизвестными ключами.
    """
    if raw in (None, "", {}):
        return {}
    if isinstance(raw, str):
        import json
        try:
            raw = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise RoundingError(
                f"[{key}] Ожидается объект вида {{\"показатель\": {{\"scale\": 6, "
                f"\"decimals\": 2}}}}; правьте его через «Настройки» -> раздел отчёта -> "
                f"«Округление». Разбор не удался: {exc}"
            ) from exc
    if not isinstance(raw, dict):
        raise RoundingError(f"[{key}] Ожидается объект, получено {type(raw).__name__}.")

    categories = CATEGORIES.get(report, {})
    result: Dict[str, dict] = {}
    for name, entry in raw.items():
        category = categories.get(name)
        if category is None:
            continue
        if not isinstance(entry, dict):
            raise RoundingError(f"[{key}] {category.label}: ожидается объект, получено {entry!r}.")
        extra = set(entry) - {"scale", "decimals"}
        if extra:
            raise RoundingError(f"[{key}] {category.label}: неизвестные поля {sorted(extra)}.")
        clean: dict = {}

        scale = entry.get("scale")
        if scale is not None:
            if isinstance(scale, bool) or not isinstance(scale, int):
                raise RoundingError(f"[{key}] {category.label}: единица задаётся целым числом, "
                                    f"получено {scale!r}.")
            allowed = category.scales()
            if scale not in allowed:
                what = ("единица этого показателя не меняется" if not allowed
                        else f"допустимо: {allowed}")
                raise RoundingError(f"[{key}] {category.label}: недопустимая единица {scale} — {what}.")
            if scale != category.default_scale():
                clean["scale"] = scale

        decimals = entry.get("decimals")
        if decimals is not None:
            if isinstance(decimals, bool) or not isinstance(decimals, int) \
                    or not 0 <= decimals <= MAX_DECIMALS:
                raise RoundingError(f"[{key}] {category.label}: знаков после запятой — целое "
                                    f"от 0 до {MAX_DECIMALS}, получено {decimals!r}.")
            clean["decimals"] = decimals

        if clean:
            result[name] = clean
    return result


def describe(report: str, value: Dict[str, dict]) -> str:
    """Короткая строка для таблицы настроек: «как сейчас» или что изменено."""
    if not value:
        return "как сейчас"
    parts = []
    for category in CATALOG.get(report, []):
        entry = value.get(category.key)
        if not entry:
            continue
        bits = []
        if "scale" in entry:
            bits.append(category.scale_title(entry["scale"]))
        if "decimals" in entry:
            bits.append(f"{entry['decimals']} зн.")
        parts.append(f"{category.label}: {', '.join(bits)}")
    return "; ".join(parts)


def current_state(category: Category, entry: Optional[dict]) -> Tuple[str, str]:
    """(единица, знаки) показателя для экрана настроек."""
    entry = entry or {}
    unit = category.scale_title(entry.get("scale")) if category.kind != NONE else (category.unit or "—")
    decimals = (f"{entry['decimals']}" if "decimals" in entry else f"как сейчас ({category.now})")
    return unit, decimals
