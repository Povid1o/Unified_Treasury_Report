#!/usr/bin/env python3
"""Единая консоль запуска ETL-отчётов казначейства.

Использование:
    python console.py                                   # интерактивное меню
    python console.py ofz-rates --date 2026-07-01        # запуск отчёта напрямую
    python console.py ovp --input report.xlsx

В интерактивном режиме Ctrl+C в любой момент (при выборе отчёта, при вводе
параметров, во время самой загрузки) отменяет только текущую операцию и
возвращает в меню — программа не завершается.

Пути к папкам с исходными файлами и результатами правятся не в config.py, а
пунктом меню «Настройки» (или подкомандой settings — см. build_parser).

Чтобы добавить новый отчёт: реализовать reports.base.Report в новой папке
reports/<report_slug>/report.py и зарегистрировать его в списке REPORTS ниже.
"""
import argparse
import sys
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))

from common import settings, settings_ui, ui  # noqa: E402
from reports.ofz_rates.report import OfzRatesReport  # noqa: E402
from reports.ovp.report import OvpReport  # noqa: E402
from reports.balance_struct.report import BalanceStructReport  # noqa: E402
from reports.chpd.report import ChpdReport  # noqa: E402
from reports.nim.report import NimReport  # noqa: E402
from reports.transfert_stavka.report import TransfertStavkaReport  # noqa: E402
from reports.portfolio_dynamics.report import PortfolioDynamicsReport  # noqa: E402

REPORTS = [
    OfzRatesReport(),
    OvpReport(),
    BalanceStructReport(),
    ChpdReport(),
    NimReport(),
    TransfertStavkaReport(),
    PortfolioDynamicsReport(),
]
REPORTS_BY_SLUG = {report.slug: report for report in REPORTS}

# «Настройки» — не отчёт: это диалог, а не загрузка, поэтому он не реализует
# reports.base.Report, не попадает в REPORTS и не оборачивается спиннером.
# В меню идёт последним номером, перед «Выходом».
SETTINGS_SLUG = "settings"
SETTINGS_CHOICE = str(len(REPORTS) + 1)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="console.py",
        description="Единая консоль запуска ETL-отчётов казначейства",
    )
    sub = parser.add_subparsers(dest="report", required=False)
    for report in REPORTS:
        p = sub.add_parser(report.slug, help=report.title, description=report.description)
        report.add_arguments(p)

    p = sub.add_parser(
        SETTINGS_SLUG, help=settings_ui.MENU_TITLE,
        description=settings_ui.MENU_DESCRIPTION + ". Без аргументов — интерактивный редактор.",
    )
    p.add_argument("--list", action="store_true", help="Показать все настройки и их значения")
    p.add_argument("--group", type=str, default=None,
                   help="Ограничить --list одним разделом (" + ", ".join(g.key for g in settings.GROUPS) + ")")
    p.add_argument("--get", type=str, default=None, metavar="KEY",
                   help="Напечатать значение одной настройки (только его — удобно для скриптов)")
    p.add_argument("--set", nargs="+", default=None, metavar="KEY=VALUE",
                   help='Задать значения, например: --set ovp_dir="/Volumes/Jupiter/data/OVP"')
    p.add_argument("--reset", nargs="+", default=None, metavar="KEY",
                   help="Вернуть настройки к значениям по умолчанию")
    p.add_argument("--reset-all", action="store_true", help="Сбросить ВСЕ настройки к значениям по умолчанию")
    p.add_argument("--make-folders", type=str, default=None, metavar="SLUG",
                   help="Создать папки по датам для отчёта (" +
                        ", ".join(r.slug for r in settings.REPORT_SOURCES) + ")")
    p.add_argument("--days", type=int, default=14,
                   help="Сколько дней вперёд создавать папки (по умолчанию 14, только рабочие)")
    p.add_argument("--include-weekends", action="store_true",
                   help="Создавать папки и на выходные тоже")
    p.add_argument("--check", action="store_true",
                   help="Проверить, что все настроенные папки и файлы доступны")
    p.add_argument("--path", action="store_true", help="Напечатать путь к файлу настроек")
    return parser


def run_settings_cli(args: argparse.Namespace) -> int:
    """Неинтерактивный режим настроек. Возвращает код возврата процесса."""
    if args.path:
        print(settings.settings_path())
        return 0
    if args.get:
        print(settings.get(args.get))
        return 0

    did_something = False
    if args.reset_all:
        settings.reset_all()
        ui.success("Все настройки возвращены к значениям по умолчанию.")
        did_something = True
    if args.reset:
        for key in args.reset:
            ui.success(f"{key} = {settings.reset(key)} (по умолчанию)")
        did_something = True
    if args.set:
        settings_ui.apply_assignments(args.set)
        did_something = True
    if args.make_folders:
        settings_ui.make_folders_cli(args.make_folders, args.days, args.include_weekends)
        did_something = True
    if args.check:
        return 1 if settings_ui.print_paths() else 0
    if args.list or not did_something:
        settings_ui.print_all(args.group)
    return 0


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
        ui.print_menu(
            REPORTS,
            extra_rows=[(SETTINGS_CHOICE, settings_ui.MENU_TITLE, settings_ui.MENU_DESCRIPTION)],
        )
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

        if choice == SETTINGS_CHOICE:
            try:
                settings_ui.run_interactive()
            except KeyboardInterrupt:
                ui.console.print()
                ui.cancelled("Настройки закрыты.")
            except Exception as exc:
                ui.error(str(exc))
            ui.console.print()
            continue

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

    if args.report == SETTINGS_SLUG:
        try:
            sys.exit(run_settings_cli(args))
        except settings.SettingsError as exc:
            ui.error(str(exc))
            sys.exit(1)

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
