"""Пункт меню «Динамика + отчёт по портфелям»: оба отчёта за один запуск.

Каждый день нужны оба отчёта, и строятся они по одной и той же выгрузке
«Позиция за период». Здесь диалог проходится один раз:

1. дата и срезы T0/T-7 — как у «Динамики портфелей» (её диалог как есть);
2. выгрузка с начала года для «Отчёта по портфелям» — на ту же дату, что
   T0, без отдельного выбора;
3. RGBI, RUONIA, RWA и P&L дополнительных портфелей — как у «Отчёта по
   портфелям».

Потом оба отчёта строятся подряд. Ошибка в одном не мешает другому: второй
всё равно формируется, а ошибки показываются в конце.

Сами отчёты ничего про этот пункт не знают: здесь только их диалоги и run()
по очереди.
"""
import argparse
from pathlib import Path
from typing import List, Optional

import config
from common import ui
from reports.base import Report
from reports.portfolio_dynamics import etl as dynamics_etl
from reports.portfolio_dynamics.report import PortfolioDynamicsReport, _resolve_slice_paths
from reports.portfolio_report import etl as report_etl
from reports.portfolio_report.report import PortfolioReport, ask_inputs


def _defaults(report: Report) -> argparse.Namespace:
    """Аргументы отчёта по умолчанию — как при запуске без флагов."""
    parser = argparse.ArgumentParser(add_help=False)
    report.add_arguments(parser)
    return parser.parse_args([])


class PortfoliosDailyReport(Report):
    slug = "portfolios-daily"
    title = "Динамика + отчёт по портфелям"
    description = ("Оба отчёта по портфелям за один запуск: одна дата, один ввод RGBI, "
                   "RUONIA, RWA и P&L дополнительных портфелей")

    def __init__(self) -> None:
        self.dynamics = PortfolioDynamicsReport()
        self.portfolio = PortfolioReport()

    def add_arguments(self, parser: argparse.ArgumentParser) -> None:
        parser.add_argument("--date", type=str, default=None,
                            help="Отчётная дата (YYYY-MM-DD). По умолчанию — последняя "
                                 "доступная дата «Динамики портфелей».")
        parser.add_argument("--rgbi", type=str, default=None, help="RGBI на сегодня")
        parser.add_argument("--ruonia", type=str, default=None, help="RUONIA на сегодня, %%")
        parser.add_argument("--rwa", type=str, default=None, help="RWA на сегодня")
        parser.add_argument("--manual-pl", type=str, default=None,
                            help="P&L с начала года по «Дополнительным портфелям», млн RUB: "
                                 "«КОД=число, КОД2=число». Не указан — берётся прошлое значение.")
        parser.add_argument("--no-import", action="store_true",
                            help="Не заглядывать в папку загрузок")

    def run(self, args: argparse.Namespace) -> None:
        dynamics_args = getattr(args, "dynamics", None)
        portfolio_args = getattr(args, "portfolio", None)
        if dynamics_args is None:  # командная строка: собрать аргументы обоих отчётов
            dynamics_args, portfolio_args = self._cli_args(args)

        errors: List[str] = []
        for report, report_args in ((self.dynamics, dynamics_args),
                                    (self.portfolio, portfolio_args)):
            if report_args is None:
                continue
            ui.console.print(f"[bold cyan]── {report.title} ──[/bold cyan]")
            try:
                report.run(report_args)
            except KeyboardInterrupt:
                raise
            except Exception as exc:  # второй отчёт строится, даже если первый упал
                ui.error(f"{report.title}: {exc}")
                errors.append(report.title)
        if errors:
            raise RuntimeError(f"Не сформированы: {', '.join(errors)} — причина выше.")

    def _cli_args(self, args: argparse.Namespace):
        dynamics_args = _defaults(self.dynamics)
        dynamics_args.date = args.date
        dynamics_args.no_import = args.no_import
        # Срезы определяются здесь, а не внутри run(): по T0 выбирается дата
        # выгрузки для «Отчёта по портфелям».
        t0_path, t7_path = _resolve_slice_paths(dynamics_args)
        dynamics_args.t0_input, dynamics_args.t7_input = str(t0_path), str(t7_path)

        portfolio_args = _defaults(self.portfolio)
        day = dynamics_etl.read_business_date(t0_path)
        portfolio_args.date = day.isoformat() if day else args.date
        portfolio_args.no_import = args.no_import
        for key in ("rgbi", "ruonia", "rwa", "manual_pl"):
            setattr(portfolio_args, key, getattr(args, key))
        return dynamics_args, portfolio_args

    def collect_interactive_args(self) -> Optional[argparse.Namespace]:
        dynamics_args = self.dynamics.collect_interactive_args()
        if dynamics_args is None:
            return None

        day = dynamics_etl.read_business_date(Path(dynamics_args.t0_input))
        no_import = not config.PORTFOLIO_DYNAMICS_IMPORT_FROM_DOWNLOADS
        try:
            source = year_start_source(day, no_import)
        except report_etl.PortfolioReportError as exc:
            ui.warning(f"{exc} «Отчёт по портфелям» не сформируется.")
            answer = ui.ask("Сформировать только «Динамику портфелей»? (Y/n)", default="Y")
            if answer.strip().lower().startswith("n"):
                ui.cancelled("Отчёты не сформированы.")
                return None
            return argparse.Namespace(dynamics=dynamics_args, portfolio=None)

        ui.console.print(f"[grey70]Отчёт по портфелям: {source.name}[/grey70]")
        portfolio_args = ask_inputs(source, no_import)
        return argparse.Namespace(dynamics=dynamics_args, portfolio=portfolio_args)


def year_start_source(day, no_import: bool) -> Path:
    """Выгрузка «01.01 – дата» на day для «Отчёта по портфелям».

    Ровно на day, без подмены более ранней датой: оба отчёта должны быть за
    одну дату. Нет — PortfolioReportError с объяснением, где искали.
    """
    if day is None:
        raise report_etl.PortfolioReportError("Не удалось определить дату среза T0.")
    sources = report_etl.find_sources(report_etl.downloads_dir(no_import))
    # Файл едет из загрузок уже здесь, до спиннера — как в самих отчётах.
    return report_etl.take(report_etl.pick_source(sources, day))
