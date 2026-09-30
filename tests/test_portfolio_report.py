"""Тесты «Отчёта по портфелям».

Проверяется то, что ломается молча: DV01 складывается, а не усредняется;
Yield — средневзвешенная по стоимости, а не простая средняя; значения из
строки «Позиция: …» важнее сумм по бумагам; Open QTY сравнивается со
ВЧЕРАШНИМ выпуском, а не с самим собой при повторном прогоне.
"""
import datetime as dt
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import Workbook

BASE_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE_DIR))

import config  # noqa: E402
from common import settings  # noqa: E402
from reports.portfolio_report import etl  # noqa: E402

# Колонки выгрузки: нужные пять вперемешку с лишними, как в реальном файле.
EXPORT_HEADER = [
    "Тип актива ", "ISIN ", "Open QTY (нач.)", "Open QTY (кон.)",
    "Total Full PL with Funding", "DV01 (кон.)", "Yield (кон.)",
    "Чистая стоимость позиции (кон.)", "Duration (нач.)",
]
FIELDS = ["Open QTY (кон.)", "Total Full PL with Funding", "DV01 (кон.)",
          "Yield (кон.)", "Чистая стоимость позиции (кон.)"]


def write_export(path: Path, rows, period_end: str) -> Path:
    """rows — (Тип актива, qty, pl, dv01, yield, value); None — пустая ячейка."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Financial Position"
    ws.cell(row=1, column=1,
            value=f"Позиция за период [{period_end}] - [{period_end}] - SECURITIES")
    for j, title in enumerate(EXPORT_HEADER, start=1):
        ws.cell(row=5, column=j, value=title)
    for i, (asset_type, *values) in enumerate(rows, start=6):
        ws.cell(row=i, column=1, value=asset_type)
        for name, value in zip(FIELDS, values):
            if value is not None:
                ws.cell(row=i, column=EXPORT_HEADER.index(name) + 1, value=value)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def export_name(period_end: str) -> str:
    return f"Позиция за период  {period_end}  -  {period_end}   - SECURITIES.xlsx"


# Портфель AFS: QTY/PL/стоимость написаны в строке портфеля, DV01 и Yield — нет.
# HTM: в строке портфеля пусто — всё считается по бумагам.
DAY1 = [
    ("Позиция: AFS_TR_RUR", 150, 5_000_000, None, None, 400_000_000),
    ("Bond", 100, 3_000_000, 20_000, 10.0, 300_000_000),
    ("Bond", 50, 2_000_000, 5_000, 14.0, 100_000_000),
    ("Позиция: HTM_ALCO", None, None, None, None, None),
    ("Bond", 200, 1_000_000, 30_000, 12.0, 200_000_000),
    ("Итого", 350, 6_000_000, 55_000, None, 600_000_000),
]
DAY2 = [
    ("Позиция: AFS_TR_RUR", 170, 6_000_000, None, None, 450_000_000),
    ("Bond", 120, 4_000_000, 22_000, 10.0, 350_000_000),
    ("Bond", 50, 2_000_000, 5_000, 14.0, 100_000_000),
    ("Позиция: OFZ_PD", 10, 100_000, None, None, 10_000_000),
    ("Bond", 10, 100_000, 1_000, 15.0, 10_000_000),
]


class PortfolioReportTest(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self._saved_env = os.environ.get(settings.SETTINGS_FILE_ENV)
        settings_file = self.tmp / "settings.json"
        settings_file.write_text(json.dumps({
            "portfolio_dynamics_dir": str(self.tmp / "data"),
            "portfolio_report_output_dir": str(self.tmp / "out"),
            "downloads_dir": str(self.tmp / "downloads"),
            "portfolio_dynamics_types_file": str(self.tmp / "portfolio_types.json"),
        }), encoding="utf-8")
        os.environ[settings.SETTINGS_FILE_ENV] = str(settings_file)
        config.reload()
        self.day1 = write_export(self.tmp / "data" / "2026-09-28" / export_name("28.09.2026"),
                                 DAY1, "28.09.2026")
        self.day2 = write_export(self.tmp / "downloads" / export_name("29.09.2026"),
                                 DAY2, "29.09.2026")

    def tearDown(self):
        if self._saved_env is None:
            os.environ.pop(settings.SETTINGS_FILE_ENV, None)
        else:
            os.environ[settings.SETTINGS_FILE_ENV] = self._saved_env
        config.reload()
        self._tmp.cleanup()

    def row(self, frame: pd.DataFrame, code: str) -> dict:
        return frame.set_index("portfolio_code").loc[code].to_dict()

    def test_portfolio_values(self):
        frame = etl.parse_positions(self.day1).frame
        afs = self.row(frame, "AFS_TR_RUR")
        # Из строки портфеля, а не суммой по бумагам.
        self.assertEqual(afs["open_qty"], 150)
        self.assertEqual(afs["net_value"], 400_000_000)
        # DV01 — сумма, Yield — средневзвешенная по стоимости: (10*300 + 14*100) / 400.
        self.assertEqual(afs["dv01"], 25_000)
        self.assertAlmostEqual(afs["yield"], 11.0)
        htm = self.row(frame, "HTM_ALCO")
        self.assertEqual(htm["open_qty"], 200)
        self.assertEqual(htm["total_pl"], 1_000_000)
        self.assertEqual(len(frame), 2)  # строка «Итого» не портфель

    def test_sources_and_import(self):
        sources = etl.find_sources(Path(config.DOWNLOADS_DIR))
        self.assertEqual([s.business_date for s in sources],
                         [dt.date(2026, 9, 29), dt.date(2026, 9, 28)])
        self.assertEqual(sources[0].origin, etl.ORIGIN_DOWNLOADS)
        path = etl.take(sources[0], move=False)
        self.assertEqual(path.parent.name, "2026-09-29")
        self.assertTrue(path.exists() and self.day2.exists())
        # Праздник/нет выгрузки на T-1 — по умолчанию берётся ближайшая более ранняя.
        self.assertEqual(etl.pick_source(sources, dt.date(2026, 9, 30), strict=False)
                         .business_date, dt.date(2026, 9, 29))
        with self.assertRaises(etl.PortfolioReportError):
            etl.pick_source(sources, dt.date(2026, 9, 30))

    def test_previous_business_day(self):
        self.assertEqual(etl.previous_business_day(dt.date(2026, 9, 28)),  # пн
                         dt.date(2026, 9, 25))
        self.assertEqual(etl.previous_business_day(dt.date(2026, 9, 30)),
                         dt.date(2026, 9, 29))

    def test_two_runs_compare_open_qty(self):
        out = Path(config.PORTFOLIO_REPORT_OUTPUT_DIR)
        first = etl.build_data(self.day1, market=etl.MarketInputs(rgbi=115.2))
        self.assertTrue(first.frame["open_qty_change"].isna().all())
        etl.save_report(first, etl.default_output_path(first.business_date))

        market = etl.MarketInputs(rgbi=116.0, ruonia=16.5, rwa=1.5e12)
        for _ in range(2):  # повторный прогон за ту же дату сравнивает со вчерашним
            previous = etl.find_previous_release(out, before=dt.date(2026, 9, 29))
            self.assertEqual(etl.release_date(previous), dt.date(2026, 9, 28))
            second = etl.build_data(self.day2, previous_path=previous, market=market,
                                    report_date=dt.date(2026, 9, 30))
            path = etl.save_report(second, etl.default_output_path(second.business_date))

        frame = second.frame
        self.assertEqual(self.row(frame, "AFS_TR_RUR")["open_qty_change"], 20)
        self.assertEqual(self.row(frame, "OFZ_PD")["open_qty_change"], 10)  # новый
        gone = self.row(frame, "HTM_ALCO")  # был вчера, сегодня нет
        self.assertEqual((gone["open_qty"], gone["open_qty_change"]), (0, -200))

        reloaded = etl.load_previous_release(path)
        self.assertEqual(reloaded.business_date, dt.date(2026, 9, 29))
        self.assertEqual(reloaded.open_qty, {"AFS_TR_RUR": 170, "OFZ_PD": 10, "HTM_ALCO": 0})
        self.assertEqual((reloaded.market.rgbi, reloaded.market.ruonia, reloaded.market.rwa),
                         (116.0, 16.5, 1.5e12))

        flat = pd.read_csv(path, encoding="utf-8-sig", keep_default_na=False)
        self.assertEqual(list(flat.columns), etl.OUT_COLUMNS)
        self.assertEqual(set(flat["date_"]), {"2026-09-29"})
        self.assertEqual(flat["id"].tolist(), list(range(len(flat))))

        def value(group, code, metric):
            hit = flat[(flat["axis_1"] == group) & (flat["axis_2"] == code)
                       & (flat["axis_3"] == metric)]
            self.assertEqual(len(hit), 1, (group, code, metric))
            return float(hit["value"].iloc[0])

        self.assertEqual(value("AFS", "AFS_TR_RUR", "Изменение Open QTY"), 20)
        self.assertEqual(value("AFS", "AFS_TR_RUR", "Чистая стоимость"), 450_000_000)
        self.assertEqual(value("AFS", "AFS_TR_RUR", "DV01"), 27_000)
        self.assertEqual(value("Рынок", "", "RUONIA"), 16.5)
        # У пропавшего портфеля только Open QTY и изменение — пустые значения не пишутся.
        self.assertEqual(sorted(flat.loc[flat["axis_2"] == "HTM_ALCO", "axis_3"]),
                         ["Open QTY", "Изменение Open QTY"])
        units = dict(zip(flat["axis_3"], flat["axis_4"]))
        self.assertEqual((units["Open QTY"], units["Yield"], units["RGBI"]), ("шт", "%", "пункты"))

if __name__ == "__main__":
    unittest.main()
