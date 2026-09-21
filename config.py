"""Пути к исходным данным и результатам файловых отчётов.

Значения больше не зашиты в этот файл: он собирается из слоя настроек
(common/settings.py), где лежат значения по умолчанию, а переопределения
пользователя хранятся в settings.json рядом с проектом.

Менять пути правкой этого файла БОЛЬШЕ НЕ НУЖНО — и не нужно было бы делать,
даже если очень хочется: следующее обновление проекта затрёт правку. Вместо
этого:

    python console.py                  # пункт меню «Настройки»
    python console.py settings --list
    python console.py settings --set ovp_dir="/Volumes/Jupiter/data/OVP"
    python console.py settings --check # проверить, что все пути доступны

ВАЖНО про NIM, TransfertStavka и ОВП: в исходных ноутбуках/GUI не было
зафиксированной общей сетевой папки для входных файлов — NIM запрашивал путь
через input() при каждом запуске, TransfertStavka использовал личную папку
Downloads автора ноутбука, а ОВП выбирался вручную через диалог выбора файла.
Значения по умолчанию для них — ЭКСТРАПОЛЯЦИЯ по аналогии с BalanceStruct/CHPD
(Jupiter\\data\\<Отчёт> и Jupiter\\output\\<Отчёт>). Проверьте их через
«Настройки» -> «Проверить пути» перед использованием в проде.

Если папка/файлы по указанному пути не найдутся — консоль не упадёт с ошибкой,
а предложит вписать путь к файлу вручную.

Код отчётов читает отсюда обычные константы (config.OVP_SOURCE и т.п.) и про
слой настроек ничего не знает. После изменения настроек в том же процессе
нужно вызвать config.reload() — это делает редактор настроек.
"""
from pathlib import Path

from common import settings as _settings
from common.file_discovery import SourceConfig

# Имена, которые экспортирует модуль. Объявлены заранее, чтобы было видно
# состав конфига, даже когда значения проставляет _apply_settings().
JUPITER_ROOT: Path
LOGS_DIR: Path
DATE_FOLDER_FORMAT: str
DOWNLOADS_DIR: Path

CBONDS_ENV_PATH: Path
OFZ_OUTPUT_PATH: Path
OFZ_LOOKBACK_DAYS: int

OVP_SOURCE: SourceConfig
OVP_OUTPUT_DIR: Path

BALANCE_STRUCT_SOURCE: SourceConfig
BALANCE_STRUCT_OUTPUT_DIR: Path

CHPD_SOURCE: SourceConfig
CHPD_OUTPUT_DIR: Path

NIM_SOURCE: SourceConfig
NIM_OUTPUT_DIR: Path

TRANSFERT_SHORT_SOURCE: SourceConfig
TRANSFERT_LONG_SOURCE: SourceConfig
TRANSFERT_OUTPUT_DIR: Path

PORTFOLIO_DYNAMICS_FILENAME_REGEX: str
PORTFOLIO_DYNAMICS_DATE_FORMAT: str
PORTFOLIO_DYNAMICS_DIR: Path
PORTFOLIO_DYNAMICS_T0_SOURCE: SourceConfig
PORTFOLIO_DYNAMICS_T7_SOURCE: SourceConfig
PORTFOLIO_DYNAMICS_OUTPUT_DIR: Path
PORTFOLIO_DYNAMICS_VALUE_SCALE: float
PORTFOLIO_DYNAMICS_TOLERANCE: float
PORTFOLIO_DYNAMICS_DEFAULT_LOOKBACK: int
PORTFOLIO_DYNAMICS_TYPE_PARENTS: str
PORTFOLIO_DYNAMICS_NESTED_LIMITS: str
PORTFOLIO_DYNAMICS_MANUAL_PORTFOLIOS: list
PORTFOLIO_DYNAMICS_TYPE_RULES: str
PORTFOLIO_DYNAMICS_PORTFOLIO_TYPES: str
PORTFOLIO_DYNAMICS_HISTORY_FILE: str
PORTFOLIO_DYNAMICS_HISTORY_ALIASES: str
PORTFOLIO_DYNAMICS_HISTORY_SCALE: float
PORTFOLIO_DYNAMICS_LIMITS_REGEX: str
PORTFOLIO_DYNAMICS_LIMITS_SOURCE: SourceConfig
PORTFOLIO_DYNAMICS_LIMIT_SCALE: float
PORTFOLIO_DYNAMICS_LIMIT_ALIASES: str
PORTFOLIO_DYNAMICS_IMPORT_FROM_DOWNLOADS: bool
PORTFOLIO_DYNAMICS_MOVE_FROM_DOWNLOADS: bool
PORTFOLIO_DYNAMICS_ARCHIVE_OWN_DATE: bool


def _apply_settings() -> None:
    """Пересобирает константы модуля из текущих значений настроек."""
    v = _settings.values()
    g = globals()

    g["JUPITER_ROOT"] = v["jupiter_root"]
    g["LOGS_DIR"] = v["logs_dir"]
    # Имя подпапки с выгрузками за один день (см. common/file_discovery.py,
    # раздел «ПАПКИ-ДАТЫ»). Передаётся только тем источникам, у которых папки
    # по датам включены настройкой.
    g["DATE_FOLDER_FORMAT"] = v["date_folder_format"]
    # Папка, куда браузер складывает скачанные выгрузки: отчёты с приёмкой ищут
    # в ней недостающие файлы (см. reports/portfolio_dynamics/inbox.py).
    g["DOWNLOADS_DIR"] = v["downloads_dir"]

    # ── Ставки ОФЗ (CBonds API) ──────────────────────────────────────────────
    # Логин и пароль лежат в .env и настройкой не являются — см. common/settings.py.
    g["CBONDS_ENV_PATH"] = v["cbonds_env_path"]
    g["OFZ_OUTPUT_PATH"] = v["ofz_output_path"]
    g["OFZ_LOOKBACK_DAYS"] = v["ofz_lookback_days"]

    # ── ОВП ──────────────────────────────────────────────────────────────────
    # У файлов ОВП нет даты в имени (в отличие от остальных отчётов), поэтому
    # filename_regex/date_format не заданы — источник сортирует файлы по дате
    # изменения (mtime) вместо даты, разобранной из имени.
    g["OVP_SOURCE"] = SourceConfig(
        directory=v["ovp_dir"], glob_pattern=v["ovp_glob"], label="ОВП",
    )
    g["OVP_OUTPUT_DIR"] = v["ovp_output_dir"]

    # ── BalanceStruct («Структура баланса») ──────────────────────────────────
    g["BALANCE_STRUCT_SOURCE"] = SourceConfig(
        directory=v["balance_struct_dir"],
        filename_regex=v["balance_struct_regex"],
        date_format=v["balance_struct_date_format"],
        label="BalanceStruct (ПФ)",
    )
    g["BALANCE_STRUCT_OUTPUT_DIR"] = v["balance_struct_output_dir"]

    # ── ЧПД ──────────────────────────────────────────────────────────────────
    g["CHPD_SOURCE"] = SourceConfig(
        directory=v["chpd_dir"], filename_regex=v["chpd_regex"],
        date_format=v["chpd_date_format"], label="ЧПД",
    )
    g["CHPD_OUTPUT_DIR"] = v["chpd_output_dir"]

    # ── NIM ──────────────────────────────────────────────────────────────────
    g["NIM_SOURCE"] = SourceConfig(
        directory=v["nim_dir"], filename_regex=v["nim_regex"],
        date_format=v["nim_date_format"], label="NIM",
    )
    g["NIM_OUTPUT_DIR"] = v["nim_output_dir"]

    # ── Трансфертные ставки: ДВА независимых источника (короткие / длинные) ──
    g["TRANSFERT_SHORT_SOURCE"] = SourceConfig(
        directory=v["transfert_short_dir"], filename_regex=v["transfert_short_regex"],
        date_format=v["transfert_short_date_format"], label="ТС до 3М",
    )
    g["TRANSFERT_LONG_SOURCE"] = SourceConfig(
        directory=v["transfert_long_dir"], filename_regex=v["transfert_long_regex"],
        date_format=v["transfert_long_date_format"], label="ТС свыше 3М",
    )
    g["TRANSFERT_OUTPUT_DIR"] = v["transfert_output_dir"]

    # ── Динамика портфелей: ДВА среза одной и той же выгрузки (T0 и T-7) ─────
    # Имена файлов вида "Позиция за период [01.01.2026] - [01.09.2026] - SECURITIES.xlsx".
    # В имени ДВЕ даты; отчётной считается ВТОРАЯ (конец периода выгрузки) — её и
    # захватывает единственная группа регулярки. Оба среза лежат в одной папке и
    # различаются только датой, поэтому источник один и тот же, а label разный —
    # он попадает в заголовки интерактивного выбора файла.
    g["PORTFOLIO_DYNAMICS_FILENAME_REGEX"] = v["portfolio_dynamics_regex"]
    g["PORTFOLIO_DYNAMICS_DATE_FORMAT"] = v["portfolio_dynamics_date_format"]
    g["PORTFOLIO_DYNAMICS_DIR"] = v["portfolio_dynamics_dir"]
    # Оба среза за день лежат в подпапке-дате: отчёт берёт самую свежую папку,
    # в которой есть файлы, и сам определяет по ним, какой срез T0, а какой T-7.
    folder_format = v["date_folder_format"] if v["portfolio_dynamics_use_date_folders"] else None
    g["PORTFOLIO_DYNAMICS_T0_SOURCE"] = SourceConfig(
        directory=v["portfolio_dynamics_dir"], filename_regex=v["portfolio_dynamics_regex"],
        date_format=v["portfolio_dynamics_date_format"], label="Динамика портфелей T0",
        date_folder_format=folder_format,
    )
    g["PORTFOLIO_DYNAMICS_T7_SOURCE"] = SourceConfig(
        directory=v["portfolio_dynamics_dir"], filename_regex=v["portfolio_dynamics_regex"],
        date_format=v["portfolio_dynamics_date_format"], label="Динамика портфелей T-7",
        date_folder_format=folder_format,
    )
    g["PORTFOLIO_DYNAMICS_OUTPUT_DIR"] = v["portfolio_dynamics_output_dir"]

    # Лимиты приходят ТРЕТЬИМ файлом на отчётную дату — «Состояние лимитов на
    # дату 21_09_2026 - Результат.xlsx». Он лежит там же, где выгрузки позиций,
    # и датируется так же, но имя и формат даты у него свои.
    # Лимит в выгрузке совокупный: HTM ограничивает HTM вместе с HTM_KUAP.
    g["PORTFOLIO_DYNAMICS_TYPE_PARENTS"] = v["portfolio_dynamics_type_parents"]
    # Сколько из совокупного лимита выделено вложенному типу (млн RUB).
    g["PORTFOLIO_DYNAMICS_NESTED_LIMITS"] = v["portfolio_dynamics_nested_limits"]
    # Портфели, которых нет в выгрузке, но объём по ним ведётся вручную.
    g["PORTFOLIO_DYNAMICS_MANUAL_PORTFOLIOS"] = v["portfolio_dynamics_manual_portfolios"]
    # Разметка портфелей по типам: правила по подстроке в коде и точечные исключения.
    g["PORTFOLIO_DYNAMICS_TYPE_RULES"] = v["portfolio_dynamics_type_rules"]
    g["PORTFOLIO_DYNAMICS_PORTFOLIO_TYPES"] = v["portfolio_dynamics_portfolio_types"]
    # Разовый импорт накопленной истории из отчёта старого формата.
    g["PORTFOLIO_DYNAMICS_HISTORY_FILE"] = v["portfolio_dynamics_history_file"]
    g["PORTFOLIO_DYNAMICS_HISTORY_ALIASES"] = v["portfolio_dynamics_history_aliases"]
    g["PORTFOLIO_DYNAMICS_HISTORY_SCALE"] = v["portfolio_dynamics_history_scale"]
    g["PORTFOLIO_DYNAMICS_LIMITS_REGEX"] = v["portfolio_dynamics_limits_regex"]
    g["PORTFOLIO_DYNAMICS_LIMITS_SOURCE"] = SourceConfig(
        directory=v["portfolio_dynamics_dir"], filename_regex=v["portfolio_dynamics_limits_regex"],
        date_format=v["portfolio_dynamics_limits_date_format"], label="Лимиты портфелей",
        date_folder_format=folder_format,
    )
    g["PORTFOLIO_DYNAMICS_LIMIT_SCALE"] = v["portfolio_dynamics_limit_scale"]
    g["PORTFOLIO_DYNAMICS_LIMIT_ALIASES"] = v["portfolio_dynamics_limit_aliases"]

    # Схема v3.0 требует млн RUB, а выгрузка отдаёт объёмы в рублях: делим на это
    # число. Если формат выгрузки изменится (тысячи, уже млн) — правится в
    # настройках, в парсере масштаб не хардкодится.
    g["PORTFOLIO_DYNAMICS_VALUE_SCALE"] = v["portfolio_dynamics_value_scale"]
    # Порог сверки подытога в строке "Позиция: ..." с суммой по бумагам и сверки
    # грейнов на витрине (CHK_16). 0.5% — из CONTRACT.md.
    g["PORTFOLIO_DYNAMICS_TOLERANCE"] = v["portfolio_dynamics_tolerance"]
    # Сдвиг сравнения по умолчанию (календарных дней) — используется, если даты
    # срезов T0/T-7 определить не удалось. Обычно считается как разница их дат.
    g["PORTFOLIO_DYNAMICS_DEFAULT_LOOKBACK"] = v["portfolio_dynamics_default_lookback"]
    # Приёмка из загрузок: искать ли там недостающие срезы и переносить их или копировать.
    g["PORTFOLIO_DYNAMICS_IMPORT_FROM_DOWNLOADS"] = v["portfolio_dynamics_import_from_downloads"]
    g["PORTFOLIO_DYNAMICS_MOVE_FROM_DOWNLOADS"] = v["portfolio_dynamics_move_from_downloads"]
    # Второй экземпляр среза в папке его собственной даты — чтобы файл, взятый
    # как T-7, потом можно было использовать как T0 своей даты.
    g["PORTFOLIO_DYNAMICS_ARCHIVE_OWN_DATE"] = v["portfolio_dynamics_archive_own_date"]


def reload() -> None:
    """Перечитывает settings.json и пересобирает константы этого модуля.

    Нужно после правки настроек в том же процессе: отчёты читают config.X в
    момент запуска, поэтому иначе в уже запущенной консоли остались бы старые
    пути до перезапуска.
    """
    _settings.reload()
    _apply_settings()


_apply_settings()
