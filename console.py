#!/usr/bin/env python3
"""Единая консоль запуска ETL-отчётов казначейства.

Использование:
    python console.py                                   # интерактивное меню
    python console.py ofz-rates --date 2026-07-01        # запуск отчёта напрямую
    python console.py ovp --input report.xlsx

Чтобы добавить новый отчёт: реализовать reports.base.Report в новой папке
reports/<report_slug>/report.py и зарегистрировать его в списке REPORTS ниже.
"""
import argparse
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))

from reports.ofz_rates.report import OfzRatesReport  # noqa: E402
from reports.ovp.report import OvpReport  # noqa: E402

REPORTS = [
    OfzRatesReport(),
    OvpReport(),
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


def print_menu() -> None:
    print("\nДоступные отчёты:")
    for i, report in enumerate(REPORTS, start=1):
        print(f"  {i}. {report.title} — {report.description}")
    print("  0. Выход")


def run_interactive_menu() -> None:
    while True:
        print_menu()
        choice = input("Выберите отчёт (номер): ").strip()
        if choice in ("0", ""):
            print("Выход.")
            return
        try:
            report = REPORTS[int(choice) - 1]
        except (ValueError, IndexError):
            print("Некорректный выбор, попробуйте снова.")
            continue

        try:
            report.run_interactive()
        except Exception as exc:  # ошибки конкретного отчёта не должны валить консоль
            print(f"Ошибка выполнения отчёта «{report.title}»: {exc}")


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    if args.report is None:
        run_interactive_menu()
        return

    report = REPORTS_BY_SLUG[args.report]
    try:
        report.run(args)
    except Exception as exc:
        print(f"Ошибка выполнения отчёта «{report.title}»: {exc}")
        sys.exit(1)


if __name__ == "__main__":
    main()
