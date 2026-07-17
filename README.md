# Unified_Treasury_Report

Единая консоль ETL-отчётов казначейства для BI.

## Установка

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

Учётные данные CBonds лежат в `CBonds_API/.env` (`CBONDS_LOGIN`, `CBONDS_PASSWORD`).

## Запуск

```bash
python console.py                              # интерактивное меню: выбрать отчёт -> задать параметры
python console.py ofz-rates --date 2026-07-01   # «Ставки ОФЗ» на конкретную дату, напрямую
python console.py ovp --input report.xlsx       # «ОВП»: конвертация Excel в CSV
```

## Структура проекта

- `CBonds_API/` — перенесённая библиотека-клиент CBonds JSON API (общая для всех отчётов).
- `common/` — общие утилиты (логирование и т.п.), переиспользуемые разными отчётами.
- `reports/<report_slug>/` — код конкретного отчёта: `etl.py` (получение и преобразование данных) + `report.py` (обёртка для консоли: CLI-аргументы и интерактивный режим).
  - `reports/ofz_rates/` — «Ставки ОФЗ» (кривая доходности ОФЗ из CBonds API).
  - `reports/ovp/` — «ОВП» (конвертация Excel-отчёта «Открытая валютная позиция» в CSV).
- `console.py` — точка входа: реестр отчётов + меню/CLI.
- `output/<report_slug>/` — результаты запусков.
- `logs/` — логи запусков (по одному файлу на отчёт).

## Добавление нового отчёта

1. Создать `reports/<slug>/etl.py` с получением и преобразованием данных.
2. Создать `reports/<slug>/report.py` с классом, реализующим `reports.base.Report`.
3. Зарегистрировать отчёт в списке `REPORTS` в `console.py`.
