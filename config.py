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
