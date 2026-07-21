#!/usr/bin/env python3
"""Единая консоль запуска ETL-отчётов казначейства.

Использование:
    python console.py                                   # интерактивное меню
    python console.py ofz-rates --date 2026-07-01        # запуск отчёта напрямую
    python console.py ovp --input report.xlsx

В интерактивном режиме Ctrl+C в любой момент (при выборе отчёта, при вводе
параметров, во время самой загрузки) отменяет только текущую операцию и
возвращает в меню — программа не завершается.

Чтобы добавить новый отчёт: реализовать reports.base.Report в новой папке
reports/<report_slug>/report.py и зарегистрировать его в списке REPORTS ниже.
"""
import argparse
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))

from common import ui  # noqa: E402
from reports.ofz_rates.report import OfzRatesReport  # noqa: E402
from reports.ovp.report import OvpReport  # noqa: E402
from reports.balance_struct.report import BalanceStructReport  # noqa: E402
from reports.chpd.report import ChpdReport  # noqa: E402
from reports.nim.report import NimReport  # noqa: E402
from reports.transfert_stavka.report import TransfertStavkaReport  # noqa: E402

REPORTS = [
    OfzRatesReport(),
    OvpReport(),
    BalanceStructReport(),
    ChpdReport(),
    NimReport(),
    TransfertStavkaReport(),
]
REPORTS_BY_SLUG = {report.slug: report for report in REPORTS}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="console.py",
        description="Единая консоль запуска ETL-отчётов казначейства",
    )
    sub = parser.add_subparsers(dest="report", required=False)
    for report in REPORTS:
        p = sub.add_parser(report.slug, help=report.title, description=report.description)
        report.add_arguments(p)
    return parser


def execute(report, args: argparse.Namespace) -> None:
    """Выполняет отчёт под спиннером; Ctrl+C отменяет только эту операцию."""
    try:
        with ui.spinner(f"Формируем отчёт «{report.title}»..."):
            report.run(args)
    except KeyboardInterrupt:
        ui.cancelled(f"Формирование отчёта «{report.title}» отменено.")
    except Exception as exc:
        ui.error(str(exc))


def run_interactive_menu() -> None:
    ui.print_header()
    while True:
        ui.print_menu(REPORTS)
        try:
            choice = ui.ask("Выберите отчёт (номер)", default="0")
        except KeyboardInterrupt:
            ui.console.print()
            ui.cancelled("Выход.")
            return
        except EOFError:
            ui.console.print()
            ui.console.print("До встречи!")
            return

        if choice in ("0", ""):
            ui.console.print("До встречи!")
            return

        try:
            report = REPORTS[int(choice) - 1]
        except (ValueError, IndexError):
            ui.warning("Некорректный выбор, попробуйте снова.")
            continue

        try:
            report_args = report.collect_interactive_args()
        except KeyboardInterrupt:
            ui.console.print()
            ui.cancelled(f"Настройка отчёта «{report.title}» отменена.")
            continue
        except EOFError:
            # Ввод больше недоступен (например, Ctrl+D) — возвращаться в меню
            # бессмысленно, оно снова упрётся в закрытый stdin.
            ui.console.print()
            ui.console.print("До встречи!")
            return
        except Exception as exc:
            ui.error(str(exc))
            continue

        if report_args is None:
            continue  # отчёт сам сообщил пользователю причину отмены

        execute(report, report_args)
        ui.console.print()


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    if args.report is None:
        run_interactive_menu()
        return

    report = REPORTS_BY_SLUG[args.report]
    try:
        report.run(args)
    except KeyboardInterrupt:
        ui.cancelled(f"Формирование отчёта «{report.title}» отменено.")
        sys.exit(130)
    except Exception as exc:
        ui.error(str(exc))
        sys.exit(1)


if __name__ == "__main__":
    main()
