"""Тесты «Отчёта по портфелям».

Проверяется то, что ломается молча: DV01 — сумма по бумагам, а не среднее и
не значение из строки портфеля; Yield — средневзвешенная по стоимости бумаг, а
не простая средняя и не строка портфеля; Open QTY, PL и стоимость из строки
«Позиция: …» важнее сумм по бумагам; изменение Open QTY — с начала
года, из самого файла, а выгрузка за один день (где оно всегда ноль) не
принимается.
"""
import datetime as dt
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

import argparse

import pandas as pd
from openpyxl import Workbook, load_workbook

BASE_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE_DIR))

import config  # noqa: E402
from common import settings  # noqa: E402
from reports.portfolio_report import etl, market, workbook  # noqa: E402
from reports.portfolio_report.report import PortfolioReport  # noqa: E402

# Колонки выгрузки: нужные пять вперемешку с лишними, как в реальном файле.
EXPORT_HEADER = [
    "Тип актива ", "ISIN ", "Open QTY (нач.)", "Open QTY (кон.)",
    "Total Full PL with Funding", "DV01 (кон.)", "Yield (кон.)",
    "Чистая стоимость позиции (кон.)", "Duration (нач.)",
]
FIELDS = ["Open QTY (кон.)", "Total Full PL with Funding", "DV01 (кон.)",
          "Yield (кон.)", "Чистая стоимость позиции (кон.)", "Open QTY (нач.)"]


def write_export(path: Path, rows, period_end: str, period_start: str = "01.01.2026") -> Path:
    """rows — (Тип актива, qty, pl, dv01, yield, value[, qty на начало]); None — пусто."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Financial Position"
    ws.cell(row=1, column=1,
            value=f"Позиция за период [{period_start}] - [{period_end}] - SECURITIES")
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


def export_name(period_end: str, period_start: str = "01.01.2026") -> str:
    return f"Позиция за период  {period_start}  -  {period_end}   - SECURITIES.xlsx"


# Портфель AFS: QTY (на конец и на начало года)/PL/стоимость написаны в строке
# портфеля, DV01 и Yield — нет. HTM: в строке портфеля пусто — всё по бумагам.
DAY1 = [
    ("Позиция: AFS_TR_RUR", 150, 5_000_000, None, None, 400_000_000, 100),
    ("Bond", 100, 3_000_000, 20_000, 10.0, 300_000_000, 60),
    ("Bond", 50, 2_000_000, 5_000, 14.0, 100_000_000, 30),
    ("Позиция: HTM_ALCO", None, None, None, None, None, None),
    ("Bond", 200, 1_000_000, 30_000, 12.0, 200_000_000, 250),
    ("Итого", 350, 6_000_000, 55_000, None, 600_000_000, 350),
]
# OFZ_PD куплен в этом году: на начало года — ноль.
DAY2 = [
    ("Позиция: AFS_TR_RUR", 170, 6_000_000, None, None, 450_000_000, 100),
    ("Bond", 120, 4_000_000, 22_000, 10.0, 350_000_000, 70),
    ("Bond", 50, 2_000_000, 5_000, 14.0, 100_000_000, 30),
    ("Позиция: OFZ_PD", 10, 100_000, None, None, 10_000_000, None),
    ("Bond", 10, 100_000, 1_000, 15.0, 10_000_000, 0),
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
            "portfolio_report_market_history": str(self.tmp / "market_history.csv"),
        }), encoding="utf-8")
        self._saved_seed_dir = market.SEED_DIR
        market.SEED_DIR = self.tmp / "seed"
        os.environ[settings.SETTINGS_FILE_ENV] = str(settings_file)
        config.reload()
        self.day1 = write_export(self.tmp / "data" / "2026-09-28" / export_name("28.09.2026"),
                                 DAY1, "28.09.2026")
        self.day2 = write_export(self.tmp / "downloads" / export_name("29.09.2026"),
                                 DAY2, "29.09.2026")

    def tearDown(self):
        market.SEED_DIR = self._saved_seed_dir
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
        # Из строки портфеля, а не суммой по бумагам (60 + 30 = 90).
        self.assertEqual(afs["open_qty"], 150)
        self.assertEqual(afs["open_qty_start"], 100)
        self.assertEqual(afs["net_value"], 400_000_000)
        # DV01 — сумма, Yield — средневзвешенная по стоимости: (10*300 + 14*100) / 400.
        self.assertEqual(afs["dv01"], 25_000)
        self.assertAlmostEqual(afs["yield"], 11.0)
        # Строка портфеля заполнена и DV01, и Yield — оба всё равно считаются по
        # бумагам: DV01 = 20 000 + 5 000, Yield = (10*300 + 14*100) / 400.
        stated = write_export(self.tmp / "stated.xlsx", [
            ("Позиция: AFS_TR_RUR", 150, 5_000_000, 31_000, 99.0, 400_000_000),
            ("Bond", 100, 3_000_000, 20_000, 10.0, 300_000_000),
            ("Bond", 50, 2_000_000, 5_000, 14.0, 100_000_000),
        ], "28.09.2026")
        afs = self.row(etl.parse_positions(stated).frame, "AFS_TR_RUR")
        self.assertEqual(afs["dv01"], 25_000)
        self.assertAlmostEqual(afs["yield"], 11.0)
        htm = self.row(frame, "HTM_ALCO")
        self.assertEqual((htm["open_qty"], htm["open_qty_start"]), (200, 250))
        self.assertEqual(htm["total_pl"], 1_000_000)
        self.assertEqual(len(frame), 2)  # строка «Итого» не портфель

    def test_sources_and_import(self):
        # Выгрузки за один день («Динамика» по-старому): на 29.09 рядом с выгрузкой
        # с начала года, на 30.09 — единственная.
        for day in ("29.09.2026", "30.09.2026"):
            write_export(self.tmp / "data" / f"2026-09-{day[:2]}" / export_name(day, day),
                         DAY2, day, period_start=day)
        sources = etl.find_sources(Path(config.DOWNLOADS_DIR))
        self.assertEqual([s.business_date for s in sources],
                         [dt.date(2026, 9, 30), dt.date(2026, 9, 29), dt.date(2026, 9, 28)])
        # Выгрузка с начала года из загрузок важнее выгрузки за день в своей папке.
        self.assertEqual(sources[1].origin, etl.ORIGIN_DOWNLOADS)
        self.assertEqual(sources[1].period_start, dt.date(2026, 1, 1))
        self.assertFalse(sources[0].from_year_start)
        with self.assertRaisesRegex(etl.PortfolioReportError,
                                    "только выгрузка за период 30.09.2026 - 30.09.2026"):
            etl.pick_source(sources, dt.date(2026, 9, 30))
        sources = sources[1:]
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

    def test_change_is_since_start_of_year(self):
        market = etl.MarketInputs(rgbi=116.0, ruonia=16.5, rwa=1.5e12)
        for _ in range(2):  # повторный прогон даёт то же самое: прошлые выпуски не нужны
            data = etl.build_data(self.day2, market=market, report_date=dt.date(2026, 9, 30))
            path = etl.save_report(data, etl.default_output_path(data.business_date))
        self.assertEqual(data.period_start, dt.date(2026, 1, 1))

        frame = data.frame
        self.assertEqual(self.row(frame, "AFS_TR_RUR")["open_qty_change"], 70)  # 170 - 100
        self.assertEqual(self.row(frame, "OFZ_PD")["open_qty_change"], 10)      # куплен в году
        self.assertNotIn("HTM_ALCO", set(frame["portfolio_code"]))  # нет в выгрузке — нет в отчёте

        flat = pd.read_csv(path, encoding="utf-8-sig", keep_default_na=False)
        self.assertEqual(list(flat.columns), etl.OUT_COLUMNS)
        self.assertEqual(set(flat["date_"]), {"2026-09-29"})
        self.assertEqual(flat["id"].tolist(), list(range(len(flat))))

        def value(group, code, metric):
            hit = flat[(flat["axis_1"] == group) & (flat["axis_2"] == code)
                       & (flat["axis_3"] == metric)]
            self.assertEqual(len(hit), 1, (group, code, metric))
            return float(hit["value"].iloc[0])

        self.assertEqual(value("AFS", "AFS_TR_RUR", "Open QTY на начало года"), 100)
        self.assertEqual(value("AFS", "AFS_TR_RUR", "Изменение Open QTY"), 70)
        self.assertEqual(value("AFS", "AFS_TR_RUR", "Чистая стоимость"), 450_000_000)
        self.assertEqual(value("AFS", "AFS_TR_RUR", "DV01"), 27_000)
        self.assertEqual(value("Рынок", "", "RUONIA"), 16.5)
        units = dict(zip(flat["axis_3"], flat["axis_4"]))
        self.assertEqual((units["Open QTY"], units["Yield"], units["RGBI"]), ("шт", "%", "пункты"))

    def test_single_day_export_is_rejected(self):
        single = write_export(self.tmp / export_name("29.09.2026", "29.09.2026"), DAY2,
                              "29.09.2026", period_start="29.09.2026")
        with self.assertRaisesRegex(etl.PortfolioReportError, "за один день"):
            etl.build_data(single)

    def test_only_exports_from_january_first_are_accepted(self):
        """«01.07 - 29.09» — не с начала года: изменение было бы с июля, такой файл не берётся."""
        mid_year = write_export(self.tmp / "data" / "2026-09-30" / export_name("30.09.2026", "01.07.2026"),
                                DAY2, "30.09.2026", period_start="01.07.2026")
        with self.assertRaisesRegex(etl.PortfolioReportError, "не с начала года"):
            etl.build_data(mid_year)
        last_year = write_export(self.tmp / export_name("29.09.2026", "01.01.2025"), DAY2,
                                 "29.09.2026", period_start="01.01.2025")
        with self.assertRaisesRegex(etl.PortfolioReportError, "не с начала года"):
            etl.build_data(last_year)

        sources = etl.find_sources(None)
        by_date = {s.business_date: s for s in sources}
        self.assertFalse(by_date[dt.date(2026, 9, 30)].from_year_start)
        self.assertTrue(by_date[dt.date(2026, 9, 28)].from_year_start)
        with self.assertRaisesRegex(etl.PortfolioReportError, "01.07.2026 - 30.09.2026"):
            etl.pick_source(sources, dt.date(2026, 9, 30))
        # По умолчанию (T-1 без явной даты) берётся ближайшая подходящая — 28.09.
        self.assertEqual(etl.pick_source(sources, dt.date(2026, 9, 30), strict=False).business_date,
                         dt.date(2026, 9, 28))

    def test_diagnose_shows_why_dv01_and_yield_are_empty(self):
        """Колонки называются иначе, чем ждёт отчёт, — DV01/Yield пустые, и видно почему."""
        global EXPORT_HEADER
        saved = list(EXPORT_HEADER)
        try:
            EXPORT_HEADER[EXPORT_HEADER.index("DV01 (кон.)")] = "DV01"
            EXPORT_HEADER[EXPORT_HEADER.index("Yield (кон.)")] = "Yield"
            renamed = list(EXPORT_HEADER)
            FIELDS[FIELDS.index("DV01 (кон.)")] = "DV01"
            FIELDS[FIELDS.index("Yield (кон.)")] = "Yield"
            path = write_export(self.tmp / export_name("29.09.2026"), DAY2, "29.09.2026")
        finally:
            EXPORT_HEADER[:] = saved
            FIELDS[FIELDS.index("DV01")] = "DV01 (кон.)"
            FIELDS[FIELDS.index("Yield")] = "Yield (кон.)"
        self.assertIn("DV01", renamed)

        with self.assertLogs("portfolio_report", level="WARNING") as captured:
            lines = etl.diagnose(path)
        text = "\n".join(lines)
        self.assertIn("[НЕ НАЙДЕНА] DV01 (кон.) — похожие заголовки в файле: 'DV01'", text)
        self.assertIn("[НЕ НАЙДЕНА] Yield (кон.) — похожие заголовки в файле: 'Yield'", text)
        self.assertIn("AFS_TR_RUR      2  пусто / пусто", text)
        logged = "\n".join(captured.output)
        self.assertIn("Похожие заголовки в файле: 'DV01'", logged)
        self.assertIn("DV01 не посчитан ни по одному портфелю", logged)

    def test_diagnose_reports_values_that_are_not_numbers(self):
        path = write_export(self.tmp / export_name("29.09.2026"), [
            ("Позиция: AFS_TR_RUR", 170, 6_000_000, None, None, 450_000_000, 100),
            ("Bond", 120, 4_000_000, "(22 000)", 10.0, 350_000_000, 70),
            (None, 50, 2_000_000, 5_000, 14.0, 100_000_000, 30),
        ], "29.09.2026")
        text = "\n".join(etl.diagnose(path))
        self.assertIn("[не числа]   DV01 (кон.): '(22 000)'", text)
        self.assertIn("из них с пустым «Тип актива»: 1", text)
        self.assertIn("сумма по 1 бумагам / средневзвешенная по 2 бумагам", text)

    # ── Запуск целиком: история рынка и комментарии ─────────────────────────
    def run_report(self, path: Path, **inputs) -> pd.DataFrame:
        args = argparse.Namespace(date=None, input=str(path), previous=None, no_import=True,
                                  output=None, rgbi=inputs.get("rgbi"),
                                  ruonia=inputs.get("ruonia"), rwa=inputs.get("rwa"))
        PortfolioReport().run(args)
        day = etl.positions.read_business_date(path)
        return pd.read_csv(etl.default_output_path(day), dtype=str, keep_default_na=False,
                           encoding="utf-8-sig")

    @staticmethod
    def market_rows(flat: pd.DataFrame) -> dict:
        rows = flat[flat["axis_1"] == etl.MARKET_GROUP]
        return {(r.date_, r.axis_3): float(r.value) for r in rows.itertuples()}

    def write_seed(self, lines):
        """Начальный файл как из русского Excel: «;», запятая, cp1251, дд.мм.гггг."""
        seed = self.tmp / "seed" / "market_history_seed.csv"
        seed.parent.mkdir(parents=True, exist_ok=True)
        seed.write_bytes(("Дата;RUONIA;RGBI;RWA\n" + "\n".join(lines) + "\n").encode("cp1251"))

    def test_market_history_backup_and_csv_without_duplicates(self):
        self.write_seed(["25.09.2026;16,1;115,0;", "26.09.2026;16,2;;",
                         "28.09.2026;16,3;115,2;"])
        flat = self.run_report(self.day1, rgbi="115,5")
        # Первый выпуск: вся история по свою дату. Введённое важнее начального файла,
        # не введённое (RUONIA) берётся из него.
        self.assertEqual(self.market_rows(flat), {
            ("2026-09-25", "RUONIA"): 16.1, ("2026-09-25", "RGBI"): 115.0,
            ("2026-09-26", "RUONIA"): 16.2,
            ("2026-09-28", "RUONIA"): 16.3, ("2026-09-28", "RGBI"): 115.5,
        })
        backup = market.read_history(market.history_path())
        self.assertEqual(backup[dt.date(2026, 9, 28)], {"ruonia": 16.3, "rgbi": 115.5})

        # Начальный файл поправили задним числом и дописали пропущенный день:
        # записанное в бэкапе не перебивается, пропуск дополняется.
        self.write_seed(["25.09.2026;99;115,0;", "26.09.2026;16,2;;",
                         "27.09.2026;16,25;;", "28.09.2026;16,3;115,2;"])
        for _ in range(2):  # повторный прогон за ту же дату собирает то же самое
            flat = self.run_report(self.day2, ruonia="16.5", rwa="1 500 000 000 000")
            self.assertEqual(self.market_rows(flat), {
                ("2026-09-27", "RUONIA"): 16.25,
                ("2026-09-29", "RUONIA"): 16.5, ("2026-09-29", "RWA"): 1.5e12,
            })
        backup = market.read_history(market.history_path())
        self.assertEqual(backup[dt.date(2026, 9, 25)]["ruonia"], 16.1)
        self.assertEqual(backup[dt.date(2026, 9, 29)], {"ruonia": 16.5, "rwa": 1.5e12})

        # Повторный прогон без ввода не теряет записанное: берёт его из истории.
        flat = self.run_report(self.day2)
        self.assertEqual(self.market_rows(flat)[("2026-09-29", "RUONIA")], 16.5)

    def test_backup_picks_up_values_from_older_releases(self):
        # Выпуск, сделанный до появления файла истории, — его значения тоже в бэкап.
        old = etl.build_data(self.day1, market=etl.MarketInputs(ruonia=16.0, rwa=2e12))
        etl.save_report(old, etl.default_output_path(old.business_date))
        flat = self.run_report(self.day2, ruonia="16.5")
        self.assertEqual(self.market_rows(flat), {("2026-09-29", "RUONIA"): 16.5})
        backup = market.read_history(market.history_path())
        self.assertEqual(backup[dt.date(2026, 9, 28)], {"ruonia": 16.0, "rwa": 2e12})

    def set_comment(self, day: dt.date, code: str, text: str) -> None:
        path = workbook.workbook_path(etl.default_output_path(day))
        wb = load_workbook(path)
        ws = wb[workbook.SHEET]
        header = [c.value for c in ws[workbook.HEADER_ROW]]
        col = header.index(workbook.COMMENT_HEADER) + 1
        for row in range(workbook.HEADER_ROW + 1, ws.max_row + 1):
            if ws.cell(row=row, column=1).value == code:
                ws.cell(row=row, column=col, value=text)
        wb.save(path)

    def test_comments_carry_over_like_in_dynamics(self):
        day1, day2 = dt.date(2026, 9, 28), dt.date(2026, 9, 29)
        self.run_report(self.day1)
        self.set_comment(day1, "AFS_TR_RUR", "Докупаем ОФЗ")
        self.set_comment(day1, "HTM_ALCO", "Закрыт")

        def comments(flat):
            rows = flat[flat["axis_3"] == etl.COMMENT_METRIC]
            self.assertTrue((rows["value"] == "").all())
            return dict(zip(rows["axis_2"], rows["text_value"]))

        # Перезапуск за ту же дату — правка попадает в CSV этого выпуска.
        self.assertEqual(comments(self.run_report(self.day1)),
                         {"AFS_TR_RUR": "Докупаем ОФЗ", "HTM_ALCO": "Закрыт"})
        # Следующий день: комментарии переносятся. HTM_ALCO в выгрузке нет — в CSV его
        # комментария нет, но витрина его хранит и вернёт, когда портфель появится.
        flat = self.run_report(self.day2)
        self.assertEqual(comments(flat), {"AFS_TR_RUR": "Докупаем ОФЗ"})
        self.assertEqual(workbook.read_comments(
            workbook.workbook_path(etl.default_output_path(day2))),
            {"AFS_TR_RUR": "Докупаем ОФЗ", "HTM_ALCO": "Закрыт"})
        # Комментарий не мешает числам портфеля.
        change = flat[(flat["axis_2"] == "AFS_TR_RUR") & (flat["axis_3"] == "Изменение Open QTY")]
        self.assertEqual(float(change["value"].iloc[0]), 70)

if __name__ == "__main__":
    unittest.main()
