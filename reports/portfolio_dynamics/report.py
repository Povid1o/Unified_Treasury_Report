"""Обёртка отчёта «Динамика портфелей» для единой консоли (см. console.py).

Отчёту нужны ДВА исходных файла (срез T0 и срез T-7) — как «Трансфертным
ставкам», поэтому файл запрашивается дважды. Плюс третий, неявный вход:
предыдущий выпуск самого отчёта, откуда переносятся история объёмов по типам,
лимиты и заметки (см. reports/portfolio_dynamics/etl.py). По умолчанию он
ищется сам — самый свежий xlsx в выходной папке.

Единственный отчёт, который отдаёт .xlsx вместо .csv: выход — это шаблон
обмена схемы v3.0 с формулами, витринами и проверками (CONTRACT.md).
"""
import argparse
from datetime import datetime
from pathlib import Path
from typing import Optional

import config
from common import file_discovery, ui
from reports.base import Report
from reports.portfolio_dynamics import etl, workbook


class PortfolioDynamicsReport(Report):
    slug = "portfolio-dynamics"
    title = "Динамика портфелей"
    description = "Два среза выгрузки позиций (T0 и T-7) -> xlsx схемы v3.0 с историей, лимитами и проверками"

    def add_arguments(self, parser: argparse.ArgumentParser) -> None:
        parser.add_argument(
            "--t0-date", type=str, default=None,
            help="Дата среза T0 (YYYY-MM-DD) — вторая дата в имени файла. По умолчанию — самый свежий.",
        )
        parser.add_argument(
            "--t7-date", type=str, default=None,
            help="Дата среза T-7 (YYYY-MM-DD) — вторая дата в имени файла. По умолчанию — самый свежий.",
        )
        parser.add_argument("--t0-input", type=str, default=None, help="Явный путь к файлу среза T0")
        parser.add_argument("--t7-input", type=str, default=None, help="Явный путь к файлу среза T-7")
        parser.add_argument(
            "--previous", type=str, default=None,
            help="Предыдущий выпуск отчёта (.xlsx), из которого переносятся история, лимиты и "
                 "заметки. По умолчанию — самый свежий файл в выходной папке.",
        )
        parser.add_argument(
            "--bootstrap", action="store_true",
            help="Первый выпуск: предыдущего файла нет, история заводится с одной даты, "
                 "лимиты — нулевые (заполняются руками).",
        )
        parser.add_argument("--output", type=str, default=None, help="Путь для сохранения .xlsx")

    def run(self, args: argparse.Namespace) -> None:
        t0_path = Path(args.t0_input) if args.t0_input else _resolve_by_date(
            config.PORTFOLIO_DYNAMICS_T0_SOURCE, args.t0_date
        )
        t7_path = Path(args.t7_input) if args.t7_input else _resolve_by_date(
            config.PORTFOLIO_DYNAMICS_T7_SOURCE, args.t7_date
        )
        if t0_path == t7_path:
            raise etl.PortfolioDynamicsError(
                f"Срезы T0 и T-7 указывают на один и тот же файл ({t0_path.name}). "
                "Укажите разные даты (--t0-date/--t7-date) или пути (--t0-input/--t7-input)."
            )

        bootstrap = bool(getattr(args, "bootstrap", False))
        previous_path = Path(args.previous) if args.previous else None
        if previous_path is None and not bootstrap:
            previous_path = etl.find_previous_release(config.PORTFOLIO_DYNAMICS_OUTPUT_DIR)

        data = etl.build_data(t0_path, t7_path, previous_path=previous_path, bootstrap=bootstrap)

        output_path = Path(args.output) if args.output else _default_output_path(data.business_date)
        checks = workbook.evaluate_checks(data)
        workbook.save_workbook(data, output_path, checks=checks)

        failed = [cid for cid, status, _value in checks if status == "FAIL"]
        ui.success(
            f"Готово: {len(data.fact_portfolio_snapshot)} портфелей, "
            f"{len(data.fact_type_daily)} строк истории -> {output_path}"
        )
        if failed:
            ui.warning(
                f"Лист checks: FAIL в {len(failed)} проверках ({', '.join(failed)}) — "
                "загрузка в BI заблокирована, см. лист checks в файле."
            )

    def collect_interactive_args(self) -> Optional[argparse.Namespace]:
        ui.console.print("[bold]Срез на сегодня (T0):[/bold]")
        t0_path = file_discovery.prompt_for_file(config.PORTFOLIO_DYNAMICS_T0_SOURCE)
        ui.console.print("[bold]Срез на T-7:[/bold]")
        t7_path = file_discovery.prompt_for_file(config.PORTFOLIO_DYNAMICS_T7_SOURCE)

        previous = etl.find_previous_release(config.PORTFOLIO_DYNAMICS_OUTPUT_DIR)
        bootstrap = False
        if previous is None:
            ui.warning(
                "Предыдущий выпуск отчёта не найден в "
                f"{config.PORTFOLIO_DYNAMICS_OUTPUT_DIR}. История объёмов по типам, лимиты и "
                "заметки берутся только оттуда — без него получится ПЕРВЫЙ выпуск: история с "
                "одной датой, нулевые лимиты, пустые заметки."
            )
            answer = ui.ask("Создать первый выпуск? (y/N) или путь к предыдущему файлу", default="N")
            if answer.strip().lower().startswith("y"):
                bootstrap = True
            elif answer.strip() and not answer.strip().lower().startswith("n"):
                previous = Path(answer.strip().strip('"'))
            else:
                ui.cancelled("Отчёт не сформирован: не указан предыдущий выпуск.")
                return None
        else:
            ui.console.print(f"[grey70]Предыдущий выпуск: [bold]{previous.name}[/bold][/grey70]")

        return argparse.Namespace(
            t0_input=str(t0_path), t7_input=str(t7_path),
            t0_date=None, t7_date=None,
            previous=str(previous) if previous is not None else None,
            bootstrap=bootstrap, output=None,
        )


def _resolve_by_date(source: file_discovery.SourceConfig, date_str: Optional[str]) -> Path:
    if not date_str:
        _, path = file_discovery.latest_file(source)
        return path
    try:
        target_date = datetime.strptime(date_str, "%Y-%m-%d").date()
    except ValueError as exc:
        raise ValueError(f"Некорректный формат даты '{date_str}', ожидается YYYY-MM-DD") from exc
    return file_discovery.resolve_file_for_date(source, target_date)


def _default_output_path(business_date) -> Path:
    return (config.PORTFOLIO_DYNAMICS_OUTPUT_DIR
            / f"{etl.OUTPUT_FILENAME_PREFIX}{business_date.isoformat()}.xlsx")
