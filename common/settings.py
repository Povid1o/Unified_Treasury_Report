"""Пользовательские настройки: пути, шаблоны имён файлов, константы отчётов.

Зачем. Раньше единственным способом указать, где лежат исходные файлы, была
правка config.py — то есть правка кода в рабочей копии, которую затирает
следующее обновление проекта. Здесь введён слой поверх: реестр настроек со
значениями по умолчанию (те же, что были в config.py) и файл settings.json
рядом с проектом, куда попадают ТОЛЬКО изменённые пользователем значения.
config.py собирается из этого слоя, поэтому весь остальной код продолжает
читать привычные config.OVP_SOURCE и ничего не знает про настройки.

Чего здесь нет намеренно: логина и пароля CBonds. Они лежат в CBonds_API/.env,
который закрыт .gitignore; settings.json — обычный json рядом с кодом, и
класть в него секреты означало бы ухудшить их хранение. Настройкой сделан
только ПУТЬ к .env (cbonds_env_path).

Модуль без rich и без импорта config: его импортирует сам config.py, а
интерактивный редактор живёт отдельно, в common/settings_ui.py.
"""
import datetime as dt
import json
import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Union

BASE_DIR = Path(__file__).resolve().parents[1]

# Путь к файлу настроек можно переопределить переменной окружения — это нужно
# тестам (чтобы не трогать настройки разработчика) и сценарию "одна копия кода,
# разные наборы путей у разных пользователей".
SETTINGS_FILE_ENV = "TREASURY_SETTINGS_FILE"


def settings_path() -> Path:
    override = os.environ.get(SETTINGS_FILE_ENV)
    return Path(override) if override else BASE_DIR / "settings.json"


class SettingsError(ValueError):
    """Некорректное значение настройки или испорченный файл настроек."""


# ── Виды значений ────────────────────────────────────────────────────────────
# dir/file      — путь; хранится строкой, отдаётся как Path
# text          — произвольная строка (например, маска *.xlsx)
# regex1        — регулярка с РОВНО одной группой (дата в имени файла)
# date_format   — strptime-формат содержимого этой группы
# int/float     — число больше нуля
# bool          — да/нет
KIND_HINTS = {
    "dir": "путь к папке",
    "file": "путь к файлу",
    "text": "строка",
    "regex1": "регулярное выражение с одной группой — датой в имени файла",
    "date_format": "формат даты (strptime), например %d.%m.%Y",
    "int": "целое число > 0",
    "float": "число > 0",
    "bool": "да / нет",
    "pairs": "пары «ключ=значение» через запятую; пусто — ничего не задано",
}

_TRUE_WORDS = {"да", "д", "yes", "y", "true", "1", "вкл", "on"}
_FALSE_WORDS = {"нет", "н", "no", "n", "false", "0", "выкл", "off"}

Default = Union[Any, Callable[[Callable[[str], Any]], Any]]


@dataclass(frozen=True)
class Setting:
    key: str
    label: str
    kind: str
    group: str
    help: str
    default: Default


@dataclass(frozen=True)
class Group:
    key: str
    title: str


GROUPS: List[Group] = [
    Group("common", "Общее"),
    Group("ofz", "Ставки ОФЗ"),
    Group("ovp", "ОВП"),
    Group("balance_struct", "Структура баланса"),
    Group("chpd", "ЧПД"),
    Group("nim", "NIM"),
    Group("transfert", "Трансфертные ставки"),
    Group("portfolio_dynamics", "Динамика портфелей"),
]


def _under_root(*parts: str) -> Callable[[Callable[[str], Any]], Path]:
    """Значение по умолчанию, производное от корня архива.

    Пока пользователь не переопределил конкретную папку, она следует за
    jupiter_root: сменил корень — переехали все папки сразу.
    """
    return lambda get: Path(get("jupiter_root")).joinpath(*parts)


SETTINGS: List[Setting] = [
    # ── Общее ────────────────────────────────────────────────────────────────
    Setting(
        key="jupiter_root", label="Корень сетевого архива", kind="dir", group="common",
        default=r"O:\Exchequer\Sotrudniki\Башлыков\Навигатор\Jupiter",
        help="Общая папка Jupiter. От неё по умолчанию отсчитываются все папки "
             "data/<Отчёт> и output/<Отчёт>: смена корня двигает их все разом. "
             "Папку отдельного отчёта можно задать явно — тогда она от корня "
             "больше не зависит.",
    ),
    Setting(
        key="downloads_dir", label="Папка загрузок", kind="dir", group="common",
        default=Path.home() / "Downloads",
        help="Куда браузер складывает скачанные выгрузки. Отчёты, умеющие приёмку, "
             "ищут здесь недостающие файлы и раскладывают их по своим папкам-датам. "
             "На русской Windows папка может называться «Загрузки» — поправьте путь.",
    ),
    Setting(
        key="date_folder_format", label="Формат имени папки-даты", kind="date_format", group="common",
        default="%Y-%m-%d",
        help="Как называются подпапки с выгрузками за день (по умолчанию 2026-09-18). "
             "Применяется к отчётам, у которых включены папки-даты. Меняйте только "
             "до того, как папки созданы: уже созданные под старый формат перестанут "
             "находиться.",
    ),
    Setting(
        key="logs_dir", label="Папка логов", kind="dir", group="common",
        default=BASE_DIR / "logs",
        help="Куда писать logs/<отчёт>.log. Логгеры создаются при старте, "
             "поэтому новое значение применится со следующего запуска консоли.",
    ),

    # ── Ставки ОФЗ ───────────────────────────────────────────────────────────
    Setting(
        key="cbonds_env_path", label="Файл .env с доступом к CBonds", kind="file", group="ofz",
        default=BASE_DIR / "CBonds_API" / ".env",
        help="Файл с CBONDS_LOGIN и CBONDS_PASSWORD. Сами логин и пароль здесь "
             "не хранятся и в settings.json не попадают — только путь к файлу.",
    ),
    Setting(
        key="ofz_output_path", label="Файл результата", kind="file", group="ofz",
        default=BASE_DIR / "output" / "ofz_rates" / "ofz_report.csv",
        help="Куда сохранять CSV со ставками ОФЗ по умолчанию (перебивается --output).",
    ),
    Setting(
        key="ofz_lookback_days", label="Глубина истории, дней", kind="int", group="ofz",
        default=90,
        help="Сколько календарных дней назад запрашивать от отчётной даты. "
             "У демо-подписки CBonds архив ограничен 100 днями, поэтому 90 — с запасом.",
    ),

    # ── ОВП ──────────────────────────────────────────────────────────────────
    Setting(
        key="ovp_dir", label="Папка исходных файлов", kind="dir", group="ovp",
        default=_under_root("data", "OVP"),
        help="У файлов ОВП нет даты в имени, поэтому они сортируются по дате "
             "изменения (mtime), а не по дате из имени.",
    ),
    Setting(
        key="ovp_glob", label="Маска файлов", kind="text", group="ovp",
        default="*.xlsx",
        help="Какие файлы в папке считать файлами ОВП.",
    ),
    Setting(
        key="ovp_output_dir", label="Папка результатов", kind="dir", group="ovp",
        default=_under_root("output", "OVP"), help="Куда складывать CSV.",
    ),

    # ── Структура баланса ────────────────────────────────────────────────────
    Setting(
        key="balance_struct_dir", label="Папка исходных файлов", kind="dir", group="balance_struct",
        default=_under_root("data", "Balance_Struct"), help="Файлы вида «ПФ_18_06_2026.xlsx».",
    ),
    Setting(
        key="balance_struct_regex", label="Шаблон имени файла", kind="regex1", group="balance_struct",
        default=r"^ПФ_(\d{2}_\d{2}_\d{4})\.xlsx$",
        help="Группа в скобках — дата в имени файла.",
    ),
    Setting(
        key="balance_struct_date_format", label="Формат даты в имени", kind="date_format",
        group="balance_struct", default="%d_%m_%Y", help="Как разбирать содержимое группы из шаблона.",
    ),
    Setting(
        key="balance_struct_output_dir", label="Папка результатов", kind="dir",
        group="balance_struct", default=_under_root("output", "BalanceStruct"), help="Куда складывать CSV.",
    ),

    # ── ЧПД ──────────────────────────────────────────────────────────────────
    Setting(
        key="chpd_dir", label="Папка исходных файлов", kind="dir", group="chpd",
        default=_under_root("data", "CHPD"), help="Файлы вида «ЧПД 2025 12 31.xlsx».",
    ),
    Setting(
        key="chpd_regex", label="Шаблон имени файла", kind="regex1", group="chpd",
        default=r"^ЧПД (\d{4} \d{2} \d{2})\.xlsx$", help="Группа в скобках — дата в имени файла.",
    ),
    Setting(
        key="chpd_date_format", label="Формат даты в имени", kind="date_format", group="chpd",
        default="%Y %m %d", help="Как разбирать содержимое группы из шаблона.",
    ),
    Setting(
        key="chpd_output_dir", label="Папка результатов", kind="dir", group="chpd",
        default=_under_root("output", "CHPD"), help="Куда складывать CSV.",
    ),

    # ── NIM ──────────────────────────────────────────────────────────────────
    Setting(
        key="nim_dir", label="Папка исходных файлов", kind="dir", group="nim",
        default=_under_root("data", "NIM"), help="Файлы вида «NIM_2025_09.xlsx».",
    ),
    Setting(
        key="nim_regex", label="Шаблон имени файла", kind="regex1", group="nim",
        default=r"^NIM_(\d{4}_\d{2})\.xlsx$", help="Группа в скобках — дата в имени файла.",
    ),
    Setting(
        key="nim_date_format", label="Формат даты в имени", kind="date_format", group="nim",
        default="%Y_%m", help="Как разбирать содержимое группы из шаблона.",
    ),
    Setting(
        key="nim_output_dir", label="Папка результатов", kind="dir", group="nim",
        default=_under_root("output", "NIM"), help="Куда складывать CSV.",
    ),

    # ── Трансфертные ставки: два независимых источника ───────────────────────
    Setting(
        key="transfert_short_dir", label="Папка коротких ставок (до 3М)", kind="dir",
        group="transfert", default=_under_root("data", "Transferta", "Short"),
        help="Файлы вида «ТС до 3-х месяцев с 17.03.2026 +РусФар+ФОР.xlsx».",
    ),
    Setting(
        key="transfert_short_regex", label="Шаблон имени (короткие)", kind="regex1",
        group="transfert", default=r"с (\d{2}\.\d{2}\.\d{4})", help="Дата через точки.",
    ),
    Setting(
        key="transfert_short_date_format", label="Формат даты (короткие)", kind="date_format",
        group="transfert", default="%d.%m.%Y", help="Как разбирать дату из имени файла.",
    ),
    Setting(
        key="transfert_long_dir", label="Папка длинных ставок (свыше 3М)", kind="dir",
        group="transfert", default=_under_root("data", "Transferta", "Long"),
        help="Файлы вида «ТС свыше 3-х месяцев с 30 04 2026.xlsx».",
    ),
    Setting(
        key="transfert_long_regex", label="Шаблон имени (длинные)", kind="regex1",
        group="transfert", default=r"с (\d{2} \d{2} \d{4})", help="Дата через пробелы.",
    ),
    Setting(
        key="transfert_long_date_format", label="Формат даты (длинные)", kind="date_format",
        group="transfert", default="%d %m %Y", help="Как разбирать дату из имени файла.",
    ),
    Setting(
        key="transfert_output_dir", label="Папка результатов", kind="dir", group="transfert",
        default=_under_root("output", "Transferta"), help="Куда складывать CSV.",
    ),

    # ── Динамика портфелей ───────────────────────────────────────────────────
    Setting(
        key="portfolio_dynamics_dir", label="Папка исходных файлов", kind="dir",
        group="portfolio_dynamics", default=_under_root("data", "PortfolioDynamics"),
        help="Оба среза (T0 и T-7) лежат здесь: это одна и та же выгрузка на разные даты.",
    ),
    Setting(
        key="portfolio_dynamics_use_date_folders", label="Складывать выгрузки в папки по датам",
        kind="bool", group="portfolio_dynamics", default=True,
        help="Да — оба среза за день лежат в подпапке вида 2026-09-18 внутри папки "
             "исходных файлов, и отчёт по умолчанию берёт самую свежую папку, В КОТОРОЙ "
             "ЕСТЬ файлы. Пустые папки, созданные заранее, пропускаются. Нет — старая "
             "плоская раскладка: все выгрузки за все даты лежат в одной папке. Обе "
             "раскладки сканируются одновременно, так что переходить можно постепенно.",
    ),
    Setting(
        key="portfolio_dynamics_import_from_downloads", label="Забирать выгрузки из загрузок",
        kind="bool", group="portfolio_dynamics", default=True,
        help="Да — если в папке нужной даты не хватает срезов, отчёт сам поищет их в папке "
             "загрузок (по тому же шаблону имени) и положит куда надо. Нет — файлы туда "
             "кладёт человек.",
    ),
    Setting(
        key="portfolio_dynamics_move_from_downloads", label="Переносить, а не копировать",
        kind="bool", group="portfolio_dynamics", default=True,
        help="Да — файл переносится (из загрузок исчезает, чтобы они не копили мусор). "
             "Нет — копируется, оригинал остаётся в загрузках. Ничего не удаляется ни в "
             "том, ни в другом случае.",
    ),
    Setting(
        key="portfolio_dynamics_archive_own_date", label="Дублировать срез в папку его даты",
        kind="bool", group="portfolio_dynamics", default=True,
        help="Да — каждый срез кладётся и в папку отчётной даты, и в папку своей "
             "собственной. Иначе файл за 11.09, попавший в папку 2026-09-18 как T-7, "
             "останется в единственном экземпляре, и отчёт на 11.09 потом будет не "
             "собрать без повторной выгрузки. Стоит места ровно в один лишний файл "
             "на запуск.",
    ),
    Setting(
        key="portfolio_dynamics_regex", label="Шаблон имени файла", kind="regex1",
        group="portfolio_dynamics",
        default=r"\[?\d{2}\.\d{2}\.\d{4}\]?\s*-\s*\[?(\d{2}\.\d{2}\.\d{4})\]?",
        help="В имени выгрузки ДВЕ даты; группа в скобках захватывает вторую — она и "
             "есть дата среза. Шаблон терпим к оформлению: подходит и то, как выгрузка "
             "называет ФАЙЛ («Позиция за период   18.09.2026  -  18.09.2026   - "
             "SECURITIES.xlsx» — без квадратных скобок, пробелов по несколько), и то, "
             "как та же строка выглядит в ШАПКЕ ЛИСТА («[01.01.2026] - [18.09.2026]»).",
    ),
    Setting(
        key="portfolio_dynamics_date_format", label="Формат даты в имени", kind="date_format",
        group="portfolio_dynamics", default="%d.%m.%Y", help="Как разбирать дату среза.",
    ),
    Setting(
        key="portfolio_dynamics_type_parents", label="Вложенность типов", kind="pairs",
        group="portfolio_dynamics", default="HTM_KUAP=HTM",
        help="Через запятую, вида «вложенный тип=объемлющий». Лимит в выгрузке "
             "СОВОКУПНЫЙ: лимит HTM ограничивает HTM и HTM_KUAP вместе. Поэтому объём "
             "объемлющего типа считается вместе с вложенными, а у портфелей вложенного "
             "типа снимается флаг «Входит в итог», чтобы ИТОГО по банку не задвоилось. "
             "Свой лимит у вложенного типа при этом остаётся и работает как подлимит.",
    ),
    Setting(
        key="portfolio_dynamics_history_file", label="Файл с историей (старый формат)",
        kind="file", group="portfolio_dynamics", default="",
        help="Отчёт старого формата, из которого один раз подтягивается накопленная "
             "история объёмов по типам: листы «Динамика AFS», «Динамика HTM», "
             "«Динамика TSS» с колонками «Дата» и «Текущий объём». Используется, когда "
             "предыдущего выпуска ещё нет. Импорт только ДОПОЛНЯЕТ: даты, уже "
             "накопленные своими запусками, не перезаписываются. Пусто — не подтягивать.",
    ),
    Setting(
        key="portfolio_dynamics_history_aliases", label="Соответствия типов в файле истории",
        kind="pairs", group="portfolio_dynamics", default="TSS=TTS",
        help="Через запятую, вида «имя листа=тип портфеля». В старом отчёте торговый "
             "портфель назван TSS, а в схеме v3.0 он TTS.",
    ),
    Setting(
        key="portfolio_dynamics_history_scale", label="Делитель истории", kind="float",
        group="portfolio_dynamics", default=1_000_000.0,
        help="Объёмы в старом отчёте указаны в рублях, а схема требует млн RUB. "
             "Если импорт разойдётся с сегодняшними объёмами на порядки, отчёт "
             "предупредит об этом в логе.",
    ),
    Setting(
        key="portfolio_dynamics_nested_limits", label="Сколько отдано вложенным типам",
        kind="pairs", group="portfolio_dynamics", default="",
        help="Через запятую, вида «тип=сумма в млн RUB». Лимит в выгрузке совокупный: "
             "если на HTM указано 900, а здесь задано HTM_KUAP=100, то КУАП получает "
             "лимит 100, а на весь остальной HTM остаётся 800. Пусто — подлимит не "
             "выделен: тогда объём вложенного типа складывается с объемлющим и "
             "сравнивается с общим лимитом (так безопаснее, превышение не потеряется).",
    ),
    Setting(
        key="portfolio_dynamics_limits_regex", label="Шаблон имени файла лимитов",
        kind="regex1", group="portfolio_dynamics",
        default=r"Состояние лимитов на дату (\d{2}_\d{2}_\d{4})",
        help="Отдельная выгрузка с лимитами по типам портфелей — «Состояние лимитов "
             "на дату 21_09_2026 - Результат.xlsx». Группа в скобках захватывает дату; "
             "она должна совпадать с ОТЧЁТНОЙ датой (T0).",
    ),
    Setting(
        key="portfolio_dynamics_limits_date_format", label="Формат даты в имени файла лимитов",
        kind="date_format", group="portfolio_dynamics", default="%d_%m_%Y",
        help="Как разбирать дату из имени файла лимитов (в нём дата через подчёркивания, "
             "а не через точки, как у выгрузки позиций).",
    ),
    Setting(
        key="portfolio_dynamics_limit_scale", label="Делитель лимитов", kind="float",
        group="portfolio_dynamics", default=1_000_000.0,
        help="Колонка «Лимит сверху» приходит в рублях, а схема требует млн RUB. "
             "Отдельная настройка от делителя объёмов — на случай, если единицы в двух "
             "выгрузках разойдутся.",
    ),
    Setting(
        key="portfolio_dynamics_limit_aliases", label="Соответствия типов в файле лимитов",
        kind="pairs", group="portfolio_dynamics", default="Облигации=TTS",
        help="Через запятую, вида «имя в файле лимитов=тип портфеля». Торговый портфель "
             "выгрузка лимитов называет «Облигации», а в отчёте он TTS. Строки файла, не "
             "сводящиеся ни к одному известному типу, в fact_limit не попадают и "
             "перечисляются в логе.",
    ),
    Setting(
        key="portfolio_dynamics_output_dir", label="Папка результатов", kind="dir",
        group="portfolio_dynamics", default=_under_root("output", "PortfolioDynamics"),
        help="Куда складывать xlsx. Отсюда же берётся предыдущий выпуск отчёта — "
             "источник истории, лимитов и заметок.",
    ),
    Setting(
        key="portfolio_dynamics_value_scale", label="Делитель объёмов", kind="float",
        group="portfolio_dynamics", default=1_000_000.0,
        help="Схема v3.0 требует млн RUB, а выгрузка отдаёт рубли — делим на это число. "
             "Если выгрузка начнёт отдавать тысячи, поставьте 1000; если сразу млн — 1.",
    ),
    Setting(
        key="portfolio_dynamics_tolerance", label="Порог расхождения", kind="float",
        group="portfolio_dynamics", default=0.005,
        help="Долей единицы. Больше порога — предупреждение о расхождении подытога "
             "и суммы по бумагам и FAIL в CHK_16. 0.005 = 0.5%, как в CONTRACT.md.",
    ),
    Setting(
        key="portfolio_dynamics_default_lookback", label="Сдвиг сравнения по умолчанию, дней",
        kind="int", group="portfolio_dynamics", default=7,
        help="Обычно сдвиг считается как разница дат T0 и T-7 из имён файлов. "
             "Это значение используется, только если даты определить не удалось.",
    ),
]

SETTINGS_BY_KEY: Dict[str, Setting] = {s.key: s for s in SETTINGS}
SETTINGS_BY_GROUP: Dict[str, List[Setting]] = {
    g.key: [s for s in SETTINGS if s.group == g.key] for g in GROUPS
}


# ── Разбор и проверка значений ───────────────────────────────────────────────
def parse_value(setting: Setting, raw: Any) -> Any:
    """Приводит введённое значение к нужному типу и проверяет его.

    Проверяем то, что иначе всплывёт много позже и невнятно: регулярку без
    группы (файлы просто «не находятся»), формат даты, из которого дата не
    разбирается, отрицательный делитель объёмов.
    """
    if raw is None:
        raise SettingsError(f"[{setting.key}] Пустое значение недопустимо.")

    text = str(raw).strip().strip('"').strip("'") if not isinstance(raw, (int, float)) else raw

    if setting.kind in ("dir", "file"):
        if not str(text):
            # Пустой путь допустим там, где значение по умолчанию тоже пустое:
            # это означает «не задано», а не ошибку ввода.
            if str(setting.default) == "":
                return ""
            raise SettingsError(f"[{setting.key}] Путь не может быть пустым.")
        return Path(str(text)).expanduser()

    if setting.kind == "text":
        if not str(text):
            raise SettingsError(f"[{setting.key}] Значение не может быть пустым.")
        return str(text)

    if setting.kind == "regex1":
        pattern = str(text)
        try:
            compiled = re.compile(pattern)
        except re.error as exc:
            raise SettingsError(f"[{setting.key}] Некорректное регулярное выражение: {exc}") from exc
        if compiled.groups != 1:
            raise SettingsError(
                f"[{setting.key}] В шаблоне должна быть РОВНО одна группа в скобках — "
                f"дата в имени файла, а их {compiled.groups}. "
                r"Пример: ^ЧПД (\d{4} \d{2} \d{2})\.xlsx$"
            )
        return pattern

    if setting.kind == "pairs":
        text = str(text).strip()
        if not text:
            return ""  # пусто — законное значение: соответствий не задано
        for item in text.split(","):
            key, sep, value = item.partition("=")
            if not sep or not key.strip() or not value.strip():
                raise SettingsError(
                    f"[{setting.key}] Ожидаются пары «ключ=значение» через запятую, "
                    f"не разобрано: {item.strip()!r}. Пример: HTM_KUAP=HTM"
                )
        return text

    if setting.kind == "bool":
        if isinstance(raw, bool):
            return raw
        word = str(text).strip().casefold()
        if word in _TRUE_WORDS:
            return True
        if word in _FALSE_WORDS:
            return False
        raise SettingsError(f"[{setting.key}] Ожидается «да» или «нет», получено {raw!r}.")

    if setting.kind == "date_format":
        fmt = str(text)
        if "/" in fmt or "\\" in fmt:
            raise SettingsError(
                f"[{setting.key}] В формате даты не должно быть разделителей пути "
                "(/ и \\): он используется как имя файла или папки."
            )
        if "%" not in fmt:
            raise SettingsError(
                f"[{setting.key}] Это не формат даты: в нём нет ни одной директивы %. "
                "Пример: %d.%m.%Y для 17.03.2026."
            )
        try:
            sample = dt.date(2026, 3, 17).strftime(fmt)
            dt.datetime.strptime(sample, fmt)
        except (ValueError, TypeError) as exc:
            raise SettingsError(f"[{setting.key}] Формат даты не разбирается обратно: {exc}") from exc
        return fmt

    if setting.kind in ("int", "float"):
        try:
            value = int(text) if setting.kind == "int" else float(str(text).replace(",", "."))
        except (TypeError, ValueError) as exc:
            raise SettingsError(
                f"[{setting.key}] Ожидается {KIND_HINTS[setting.kind]}, получено {raw!r}."
            ) from exc
        if value <= 0:
            raise SettingsError(f"[{setting.key}] Значение должно быть больше нуля, получено {value}.")
        return value

    raise SettingsError(f"[{setting.key}] Неизвестный вид значения {setting.kind!r}.")


def _serialize(value: Any) -> Any:
    return str(value) if isinstance(value, Path) else value


# ── Хранилище ────────────────────────────────────────────────────────────────
_overrides: Optional[Dict[str, Any]] = None


def _load_overrides() -> Dict[str, Any]:
    """Читает settings.json. Неизвестные ключи игнорирует, битый файл — ошибка.

    Игнорировать неизвестные ключи важно при откате версии проекта: файл,
    записанный более новой версией, не должен ломать запуск старой.
    """
    global _overrides
    if _overrides is not None:
        return _overrides

    path = settings_path()
    if not path.exists():
        _overrides = {}
        return _overrides

    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as exc:
        raise SettingsError(
            f"Не удалось прочитать файл настроек {path}: {exc}. "
            "Поправьте его вручную или удалите — тогда вернутся значения по умолчанию."
        ) from exc
    if not isinstance(raw, dict):
        raise SettingsError(f"Файл настроек {path} должен содержать объект JSON, а не {type(raw).__name__}.")

    _overrides = {k: v for k, v in raw.items() if k in SETTINGS_BY_KEY}
    return _overrides


def reload() -> None:
    """Сбрасывает кэш — следующее чтение возьмёт значения из файла заново."""
    global _overrides
    _overrides = None


def is_overridden(key: str) -> bool:
    return key in _load_overrides()


def overridden_keys() -> List[str]:
    return [s.key for s in SETTINGS if s.key in _load_overrides()]


def default_of(key: str) -> Any:
    """Значение по умолчанию — с учётом того, что часть из них производна от корня."""
    setting = SETTINGS_BY_KEY[key]
    default = setting.default
    return default(get) if callable(default) else default


def get(key: str) -> Any:
    """Текущее значение настройки: переопределённое пользователем либо по умолчанию."""
    if key not in SETTINGS_BY_KEY:
        raise SettingsError(f"Неизвестная настройка {key!r}.")
    overrides = _load_overrides()
    if key in overrides:
        return parse_value(SETTINGS_BY_KEY[key], overrides[key])
    value = default_of(key)
    setting = SETTINGS_BY_KEY[key]
    if setting.kind in ("dir", "file") and not isinstance(value, Path):
        # Пустая строка означает «путь не задан»; Path("") дал бы текущую
        # папку («.»), и отчёт полез бы читать её как файл.
        return Path(value) if str(value) else ""
    return value


def values() -> Dict[str, Any]:
    return {s.key: get(s.key) for s in SETTINGS}


def _write(overrides: Dict[str, Any]) -> Path:
    path = settings_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    ordered = {s.key: overrides[s.key] for s in SETTINGS if s.key in overrides}
    try:
        path.write_text(json.dumps(ordered, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    except PermissionError as exc:
        raise SettingsError(
            f"Нет доступа для записи в {path} (файл открыт в другой программе?): {exc}"
        ) from exc
    return path


def set_value(key: str, raw: Any) -> Any:
    """Переопределяет настройку и сразу сохраняет файл. Возвращает разобранное значение."""
    if key not in SETTINGS_BY_KEY:
        raise SettingsError(f"Неизвестная настройка {key!r}. Список: python console.py settings --list")
    value = parse_value(SETTINGS_BY_KEY[key], raw)
    overrides = dict(_load_overrides())
    overrides[key] = _serialize(value)
    _write(overrides)
    reload()
    return value


def reset(key: str) -> Any:
    """Убирает переопределение — настройка возвращается к значению по умолчанию."""
    if key not in SETTINGS_BY_KEY:
        raise SettingsError(f"Неизвестная настройка {key!r}.")
    overrides = dict(_load_overrides())
    overrides.pop(key, None)
    _write(overrides)
    reload()
    return get(key)


def reset_all() -> None:
    _write({})
    reload()


# ── Источники отчётов: для «создать папки по датам» ──────────────────────────
@dataclass(frozen=True)
class ReportSource:
    """Отчёт и ключи настроек с его папками исходных файлов.

    Таблица живёт здесь, а не собирается из reports/*: settings импортирует
    config, и обратная зависимость замкнула бы импорты в кольцо.
    """

    slug: str
    title: str
    dir_keys: tuple
    date_folders_key: Optional[str] = None  # настройка «папки по датам», если она есть


REPORT_SOURCES: List[ReportSource] = [
    ReportSource("ovp", "ОВП", ("ovp_dir",)),
    ReportSource("balance-struct", "Структура баланса", ("balance_struct_dir",)),
    ReportSource("chpd", "ЧПД", ("chpd_dir",)),
    ReportSource("nim", "NIM", ("nim_dir",)),
    ReportSource("transfert-stavka", "Трансфертные ставки",
                 ("transfert_short_dir", "transfert_long_dir")),
    ReportSource("portfolio-dynamics", "Динамика портфелей", ("portfolio_dynamics_dir",),
                 date_folders_key="portfolio_dynamics_use_date_folders"),
]
REPORT_SOURCES_BY_SLUG: Dict[str, ReportSource] = {r.slug: r for r in REPORT_SOURCES}


# ── Диагностика путей ────────────────────────────────────────────────────────
@dataclass(frozen=True)
class PathStatus:
    key: str
    label: str
    group: str
    path: Path
    kind: str
    exists: bool
    files: Optional[int]  # сколько файлов нашлось по маске источника (для папок-источников)
    detail: str
    is_problem: bool  # отсутствие пути мешает работе (в отличие от «создастся сама»)


# Для каждой папки-источника — маска, по которой отчёт ищет в ней файлы:
# так «папка есть, а файлов нет» видно сразу, а не при первом запуске отчёта.
_SOURCE_GLOBS = {
    "ovp_dir": "ovp_glob",
    "balance_struct_dir": None,
    "chpd_dir": None,
    "nim_dir": None,
    "transfert_short_dir": None,
    "transfert_long_dir": None,
    "portfolio_dynamics_dir": None,
}

# Пути, которые создаются сами при первом сохранении: отсутствие такой папки —
# нормальное состояние чистой установки, а не поломка. Иначе проверка путей
# кричала бы красным на каждой новой машине и перестала бы что-то значить.
_CREATED_ON_WRITE = {
    "logs_dir", "ofz_output_path", "ovp_output_dir", "balance_struct_output_dir",
    "chpd_output_dir", "nim_output_dir", "transfert_output_dir",
    "portfolio_dynamics_output_dir",
}


def check_paths() -> List[PathStatus]:
    """Проверяет доступность всех настроенных путей — не дожидаясь запуска отчёта.

    Типовая причина «отчёт не находит файлы» — не примонтированный сетевой диск
    или опечатка в пути. Проверка отвечает на это сразу и списком.
    """
    statuses: List[PathStatus] = []
    for setting in SETTINGS:
        if setting.kind not in ("dir", "file"):
            continue
        path = Path(get(setting.key))
        exists = path.exists()
        created_on_write = setting.key in _CREATED_ON_WRITE
        files: Optional[int] = None
        is_problem = False

        if exists and setting.kind == "file":
            detail = "файл на месте"
        elif exists and setting.key in _SOURCE_GLOBS:
            glob_key = _SOURCE_GLOBS[setting.key]
            pattern = str(get(glob_key)) if glob_key else "*.xlsx"
            files = sum(1 for f in path.glob(pattern) if f.is_file())
            detail = f"файлов по маске {pattern}: {files}"
        elif exists:
            detail = "папка на месте"
        elif created_on_write:
            detail = "будет создана при первом запуске"
        elif setting.kind == "file":
            detail = "файл не найден"
            is_problem = True
        else:
            detail = "папка недоступна (не примонтирован диск или опечатка в пути)"
            is_problem = True

        statuses.append(PathStatus(
            key=setting.key, label=setting.label, group=setting.group, path=path,
            kind=setting.kind, exists=exists, files=files, detail=detail,
            is_problem=is_problem,
        ))
    return statuses
