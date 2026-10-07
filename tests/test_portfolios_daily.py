"""Тесты пункта «Динамика + отчёт по портфелям».

Проверяется: один запуск даёт оба отчёта за одну дату, дополнительный
портфель с комментариями из файла попадает в оба; выгрузка с начала года для
«Отчёта по портфелям» берётся ровно на дату T0; ошибка одного отчёта не мешает
второму.
"""
import argparse
import datetime as dt
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE_DIR))
sys.path.insert(0, str(BASE_DIR / "tests"))

import config  # noqa: E402
from common import settings  # noqa: E402
from reports.portfolio_report import etl as report_etl, market  # noqa: E402
from reports.portfolios_daily.report import (  # noqa: E402
    PortfoliosDailyReport, _defaults, year_start_source,
)
# Модулями, а не именами: импортированный TestCase unittest прогнал бы ещё раз.
import test_portfolio_dynamics_etl as dyn  # noqa: E402
import test_portfolio_report as rep  # noqa: E402

EXTRA = {"code": "OFZ_EXTRA", "type": "HTM", "volume": 150, "qty": 1_000, "duration": 4.2,
         "comment_report": "Для портфелей", "comment_dynamics": "Для динамики"}


class PortfoliosDailyTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self._saved_env = os.environ.get(settings.SETTINGS_FILE_ENV)
        settings_file = self.tmp / "settings.json"
        settings_file.write_text(json.dumps({
            "portfolio_dynamics_dir": str(self.tmp / "data"),
            "portfolio_report_dir": str(self.tmp / "report_data"),
            "portfolio_dynamics_output_dir": str(self.tmp / "out_dynamics"),
            "portfolio_report_output_dir": str(self.tmp / "out_report"),
            "downloads_dir": str(self.tmp / "downloads"),
            "portfolio_dynamics_types_file": str(self.tmp / "portfolio_types.json"),
        }), encoding="utf-8")
        self._saved_seed_dir = market.SEED_DIR
        market.SEED_DIR = self.tmp / "seed"
        os.environ[settings.SETTINGS_FILE_ENV] = str(settings_file)
        config.reload()
        settings.set_value("portfolio_dynamics_manual_portfolios", [EXTRA])
        config.reload()

        self.t0 = dyn.write_export(self.tmp / "slices" / dyn.export_name("21.09.2026"),
                                   dyn.T0_ROWS, period_end="21.09.2026")
        self.t7 = dyn.write_export(self.tmp / "slices" / dyn.export_name("14.09.2026"),
                                   dyn.T0_ROWS, period_end="14.09.2026")
        self.ytd = rep.write_export(
            self.tmp / "report_data" / "2026-09-21" / rep.export_name("21.09.2026"),
            rep.DAY1, "21.09.2026")

    def tearDown(self):
        market.SEED_DIR = self._saved_seed_dir
        if self._saved_env is None:
            os.environ.pop(settings.SETTINGS_FILE_ENV, None)
        else:
            os.environ[settings.SETTINGS_FILE_ENV] = self._saved_env
        config.reload()
        self._tmp.cleanup()

    def args(self, t0=None, portfolio_input=None):
        report = PortfoliosDailyReport()
        dynamics = _defaults(report.dynamics)
        dynamics.t0_input, dynamics.t7_input = str(t0 or self.t0), str(self.t7)
        dynamics.bootstrap = True
        portfolio = _defaults(report.portfolio)
        portfolio.input = str(portfolio_input or self.ytd)
        portfolio.no_import = True
        portfolio.rgbi, portfolio.manual_pl = "115,5", "OFZ_EXTRA=12,5"
        return report, argparse.Namespace(dynamics=dynamics, portfolio=portfolio)

    def test_one_run_builds_both_reports(self):
        report, args = self.args()
        report.run(args)

        flat = pd.read_csv(report_etl.default_output_path(dt.date(2026, 9, 21)), dtype=str,
                           keep_default_na=False, encoding="utf-8-sig")
        extra = flat[flat["axis_2"] == "OFZ_EXTRA"].set_index("axis_3")
        self.assertEqual(float(extra.loc["Total Full PL with Funding", "value"]), 12_500_000)
        self.assertEqual(extra.loc["Комментарий", "text_value"], "Для портфелей")

        dynamics = list((self.tmp / "out_dynamics").glob("*2026-09-21*.csv"))
        self.assertEqual(len(dynamics), 1, "плоский CSV «Динамики» за ту же дату")
        dyn_flat = pd.read_csv(dynamics[0], dtype=str, keep_default_na=False,
                               encoding="utf-8-sig")
        self.assertIn("Для динамики", set(dyn_flat["text_value"]))
        self.assertNotIn("Для портфелей", set(dyn_flat["text_value"]))

    def test_failure_of_one_report_does_not_stop_the_other(self):
        report, args = self.args(t0=self.tmp / "нет такого файла.xlsx")
        with self.assertRaisesRegex(RuntimeError, "Динамика портфелей"):
            report.run(args)
        self.assertTrue(report_etl.default_output_path(dt.date(2026, 9, 21)).is_file())

    def test_year_start_export_is_taken_for_the_t0_date(self):
        self.assertEqual(year_start_source(dt.date(2026, 9, 21), no_import=True), self.ytd)
        # На другую дату выгрузки нет — более ранняя не подставляется: оба
        # отчёта должны быть за одну дату.
        with self.assertRaises(report_etl.PortfolioReportError):
            year_start_source(dt.date(2026, 9, 22), no_import=True)


if __name__ == "__main__":
    unittest.main()
