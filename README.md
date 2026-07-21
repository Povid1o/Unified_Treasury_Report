# Unified_Treasury_Report

Единая консоль ETL-отчётов казначейства для BI.

## Установка

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

Учётные данные CBonds лежат в `CBonds_API/.env` (`CBONDS_LOGIN`, `CBONDS_PASSWORD`).

Пути к папкам с исходными Excel-файлами и результатами для файловых отчётов
(BalanceStruct, ЧПД, NIM, Трансфертные ставки) настраиваются в [config.py](config.py).

## Запуск

```bash
python console.py                              # интерактивное меню: выбрать отчёт -> дату -> запуск
python console.py ofz-rates --date 2026-07-01   # «Ставки ОФЗ» на конкретную дату (CBonds API)
python console.py ovp --input report.xlsx       # «ОВП»: конвертация Excel в CSV
python console.py balance-struct --date 2026-06-18   # «Структура баланса» — файл на дату, либо --input <путь>
python console.py chpd --date 2025-12-31
python console.py nim --date 2025-09-01
python console.py transfert-stavka --short-date 2026-03-17 --long-date 2026-04-30
```

Без аргумента `--date`/`--input` файловые отчёты (BalanceStruct, ЧПД, NIM,
Трансфертные ставки) сами находят самый свежий файл в папке источника
(см. `config.py`). В интерактивном режиме (без аргументов командной строки)
консоль дополнительно предложит 5 последних найденных дат на выбор — либо
можно ввести нужную дату вручную.

## Структура проекта

- `CBonds_API/` — перенесённая библиотека-клиент CBonds JSON API (общая для всех отчётов).
- `common/` — общие утилиты, переиспользуемые разными отчётами:
  - `logging_utils.py` — единое логирование (консоль + `logs/<report>.log`).
  - `file_discovery.py` — поиск входного файла по дате, зашитой в имени файла (топ-5 дат / ручной ввод даты / точный путь).
- `config.py` — пути к папкам с исходными данными и результатами для файловых отчётов (`SourceConfig` на каждый источник).
- `reports/<report_slug>/` — код конкретного отчёта: `etl.py` (получение и преобразование данных) + `report.py` (обёртка для консоли: CLI-аргументы и интерактивный режим).
  - `reports/ofz_rates/` — «Ставки ОФЗ» (кривая доходности ОФЗ из CBonds API).
  - `reports/ovp/` — «ОВП» (конвертация Excel-отчёта «Открытая валютная позиция» в CSV).
  - `reports/balance_struct/` — «Структура баланса» (разбор Excel «ПФ_ДД_ММ_ГГГГ» по иерархии активов/пассивов).
  - `reports/chpd/` — «ЧПД» (разбор Excel «ЧПД YYYY MM DD», лист Table).
  - `reports/nim/` — «NIM» (разбор Excel «NIM_YYYY_MM» по текстовым маркерам).
  - `reports/transfert_stavka/` — «Трансфертные ставки» (два файла: короткие + длинные сроки).
- `console.py` — точка входа: реестр отчётов + меню/CLI.
- `output/<report_slug>/` — результаты запусков (для файловых отчётов путь берётся из `config.py`).
- `logs/` — логи запусков (по одному файлу на отчёт).

## Добавление нового отчёта

1. Создать `reports/<slug>/etl.py` с получением и преобразованием данных.
2. Если у отчёта есть входные файлы с датой в имени — добавить `SourceConfig` в `config.py` и использовать `common.file_discovery` для поиска файла.
3. Создать `reports/<slug>/report.py` с классом, реализующим `reports.base.Report`.
4. Зарегистрировать отчёт в списке `REPORTS` в `console.py`.
