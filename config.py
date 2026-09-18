"""Единый конфиг путей к исходным данным и результатам файловых отчётов.

Пути ниже указывают на общий сетевой архив ("O:\\...\\Jupiter"), как в
исходных ноутбуках из EXPORT_FOLDER. Если сетевой диск на вашей машине
подключён под другой буквой/путём — поправьте константы ниже, ETL-код
их не хардкодит и берёт значения только отсюда.

ВАЖНО про NIM, TransfertStavka и ОВП: в исходных ноутбуках/GUI не было
зафиксированной общей сетевой папки для входных файлов — NIM запрашивал
путь через input() при каждом запуске, TransfertStavka использовал личную
папку Downloads автора ноутбука (C:\\Users\\BashlykovNV\\Downloads), а ОВП
выбирался вручную через диалог выбора файла в GUI. Пути NIM_SOURCE,
TRANSFERT_*_SOURCE и OVP_SOURCE ниже — это ЭКСТРАПОЛЯЦИЯ по аналогии с
BalanceStruct/CHPD (общий архив Jupiter\\data\\<Отчёт> и
Jupiter\\output\\<Отчёт>). Обязательно проверьте и поправьте их перед
использованием в проде.

Если папка/файлы по указанному пути не найдутся — консоль не упадёт с
ошибкой, а предложит вписать путь к файлу вручную (тот же сценарий, что
раньше был единственным для ОВП).
"""
from pathlib import Path

from common.file_discovery import SourceConfig

JUPITER_ROOT = Path(r"O:\Exchequer\Sotrudniki\Башлыков\Навигатор\Jupiter")

# ── ОВП (путь — экстраполяция, см. предупреждение в начале файла) ─────────
# У файлов ОВП нет даты в имени (в отличие от остальных отчётов), поэтому
# filename_regex/date_format не заданы — источник сортирует файлы *.xlsx по
# дате изменения (mtime) вместо даты, разобранной из имени.
OVP_SOURCE = SourceConfig(
    directory=JUPITER_ROOT / "data" / "OVP",
    glob_pattern="*.xlsx",
    label="ОВП",
)
OVP_OUTPUT_DIR = JUPITER_ROOT / "output" / "OVP"

# ── BalanceStruct ("Структура баланса") ────────────────────────────────────
# Имена файлов вида "ПФ_18_06_2026.xlsx" — путь и формат взяты дословно из
# BalanceStruct.ipynb.
BALANCE_STRUCT_SOURCE = SourceConfig(
    directory=JUPITER_ROOT / "data" / "Balance_Struct",
    filename_regex=r"^ПФ_(\d{2}_\d{2}_\d{4})\.xlsx$",
    date_format="%d_%m_%Y",
    label="BalanceStruct (ПФ)",
)
BALANCE_STRUCT_OUTPUT_DIR = JUPITER_ROOT / "output" / "BalanceStruct"

# ── ЧПД ──────────────────────────────────────────────────────────────────
# Имена файлов вида "ЧПД 2025 12 31.xlsx" — путь и формат взяты дословно из
# CHPD.ipynb.
CHPD_SOURCE = SourceConfig(
    directory=JUPITER_ROOT / "data" / "CHPD",
    filename_regex=r"^ЧПД (\d{4} \d{2} \d{2})\.xlsx$",
    date_format="%Y %m %d",
    label="ЧПД",
)
CHPD_OUTPUT_DIR = JUPITER_ROOT / "output" / "CHPD"

# ── NIM (путь — экстраполяция, см. предупреждение в начале файла) ─────────
# Имена файлов вида "NIM_2025_09.xlsx" (по примеру из NIM.ipynb).
NIM_SOURCE = SourceConfig(
    directory=JUPITER_ROOT / "data" / "NIM",
    filename_regex=r"^NIM_(\d{4}_\d{2})\.xlsx$",
    date_format="%Y_%m",
    label="NIM",
)
NIM_OUTPUT_DIR = JUPITER_ROOT / "output" / "NIM"

# ── Трансфертные ставки: ДВА независимых источника (короткие / длинные) ───
# (путь — экстраполяция, см. предупреждение в начале файла)
# Короткие: "ТС до 3-х месяцев с 17.03.2026 +РусФар+ФОР.xlsx" (дата через точки).
TRANSFERT_SHORT_SOURCE = SourceConfig(
    directory=JUPITER_ROOT / "data" / "Transferta" / "Short",
    filename_regex=r"с (\d{2}\.\d{2}\.\d{4})",
    date_format="%d.%m.%Y",
    label="ТС до 3М",
)
# Длинные: "ТС свыше 3-х месяцев с 30 04 2026.xlsx" (дата через пробелы).
TRANSFERT_LONG_SOURCE = SourceConfig(
    directory=JUPITER_ROOT / "data" / "Transferta" / "Long",
    filename_regex=r"с (\d{2} \d{2} \d{4})",
    date_format="%d %m %Y",
    label="ТС свыше 3М",
)
TRANSFERT_OUTPUT_DIR = JUPITER_ROOT / "output" / "Transferta"

# ── Динамика портфелей: ДВА среза одной и той же выгрузки (T0 и T-7) ──────
# (путь — экстраполяция, см. предупреждение в начале файла)
# Имена файлов вида "Позиция за период [01.01.2026] - [01.09.2026] - SECURITIES.xlsx".
# В имени ДВЕ даты; отчётной считается ВТОРАЯ (конец периода выгрузки) — её и
# захватывает единственная группа регулярки. Оба среза лежат в одной папке и
# различаются только датой, поэтому источник один и тот же, а label разный —
# он попадает в заголовки интерактивного выбора файла.
PORTFOLIO_DYNAMICS_FILENAME_REGEX = r"\[\d{2}\.\d{2}\.\d{4}\]\s*-\s*\[(\d{2}\.\d{2}\.\d{4})\]"
PORTFOLIO_DYNAMICS_DATE_FORMAT = "%d.%m.%Y"
PORTFOLIO_DYNAMICS_DIR = JUPITER_ROOT / "data" / "PortfolioDynamics"

PORTFOLIO_DYNAMICS_T0_SOURCE = SourceConfig(
    directory=PORTFOLIO_DYNAMICS_DIR,
    filename_regex=PORTFOLIO_DYNAMICS_FILENAME_REGEX,
    date_format=PORTFOLIO_DYNAMICS_DATE_FORMAT,
    label="Динамика портфелей T0",
)
PORTFOLIO_DYNAMICS_T7_SOURCE = SourceConfig(
    directory=PORTFOLIO_DYNAMICS_DIR,
    filename_regex=PORTFOLIO_DYNAMICS_FILENAME_REGEX,
    date_format=PORTFOLIO_DYNAMICS_DATE_FORMAT,
    label="Динамика портфелей T-7",
)
PORTFOLIO_DYNAMICS_OUTPUT_DIR = JUPITER_ROOT / "output" / "PortfolioDynamics"

# Схема v3.0 требует млн RUB, а выгрузка отдаёт объёмы в рублях: делим на это
# число. Если формат выгрузки изменится (тысячи, уже млн) — правится здесь, в
# парсере масштаб не хардкодится.
PORTFOLIO_DYNAMICS_VALUE_SCALE = 1_000_000

# Порог сверки подытога в строке "Позиция: ..." с суммой по бумагам и сверки
# грейнов на витрине (CHK_16). 0.5% — из CONTRACT.md.
PORTFOLIO_DYNAMICS_TOLERANCE = 0.005

# Сдвиг сравнения по умолчанию (календарных дней) — используется, если даты
# срезов T0/T-7 определить не удалось. Обычно считается как разница их дат.
PORTFOLIO_DYNAMICS_DEFAULT_LOOKBACK = 7
