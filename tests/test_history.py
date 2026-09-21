"""Тесты импорта накопленной истории из отчёта СТАРОГО формата.

Старый формат: по листу на тип («Динамика AFS», «Динамика HTM», «Динамика
TSS»), в каждом колонки «Дата», «Текущий объём», «Изменение». Проверяется то,
что при переносе легко испортить незаметно: имя торгового типа (TSS против
TTS), единицы, и главное — что импорт только дополняет и никогда не
перезаписывает то, что отчёт накопил сам.
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
sys.path.insert(0, str(BASE_DIR / "tests"))

import config  # noqa: E402
from common import settings  # noqa: E402
from reports.portfolio_dynamics import etl, history  # noqa: E402
from test_portfolio_dynamics_etl import MLN, write_export  # noqa: E402

RUB = 1_000_000  # рублей в одном млн: старый отчёт вёлся в рублях


def write_history_file(path: Path, sheets: dict, header_row: int = 3,
                       date_as_text: bool = True) -> Path:
    """Отчёт старого формата. sheets — {имя типа: [(дата, объём в рублях), ...]}."""
    wb = Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(f"Динамика {name}")
        ws["A1"] = f"Динамика портфеля {name}"
        for j, title in enumerate(["Дата", "Текущий объём", "Изменение"], start=1):
            ws.cell(row=header_row, column=j, value=title)
        previous = None
        for i, (day, volume) in enumerate(rows, start=header_row + 1):
            ws.cell(row=i, column=1,
                    value=day.strftime("%d.%m.%Y") if date_as_text else day)
            ws.cell(row=i, column=2, value=volume)
            if volume is not None:
                ws.cell(row=i, column=3, value=0 if previous is None else volume - previous)
                previous = volume
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


DAYS = [dt.date(2026, 9, 1), dt.date(2026, 9, 2), dt.date(2026, 9, 3)]
DEFAULT_SHEETS = {
    "AFS": [(d, (300 + i) * RUB) for i, d in enumerate(DAYS)],
    "HTM": [(d, (600 + i) * RUB) for i, d in enumerate(DAYS)],
    "TSS": [(d, (200 + i) * RUB) for i, d in enumerate(DAYS)],  # торговый назван TSS
}


class HistoryTestCase(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self._saved_env = os.environ.get(settings.SETTINGS_FILE_ENV)
        settings_file = self.tmp / "settings.json"
        settings_file.write_text(json.dumps({
            "portfolio_dynamics_dir": str(self.tmp / "data"),
            "portfolio_dynamics_output_dir": str(self.tmp / "out"),
            "downloads_dir": str(self.tmp / "downloads"),
        }), encoding="utf-8")
        os.environ[settings.SETTINGS_FILE_ENV] = str(settings_file)
        config.reload()

    def tearDown(self):
        if self._saved_env is None:
            os.environ.pop(settings.SETTINGS_FILE_ENV, None)
        else:
            os.environ[settings.SETTINGS_FILE_ENV] = self._saved_env
        config.reload()
        self._tmp.cleanup()

    def parse(self, sheets=None, **kwargs) -> pd.DataFrame:
        path = write_history_file(self.tmp / "старый отчёт.xlsx",
                                  DEFAULT_SHEETS if sheets is None else sheets, **kwargs)
        return history.parse_history_file(path)


class ParseTests(HistoryTestCase):
    def test_one_sheet_per_type_is_read(self):
        frame = self.parse()
        self.assertEqual(len(frame), 9)
        self.assertEqual(sorted(frame["portfolio_type"].unique()), ["AFS", "HTM", "TTS"])

    def test_trading_type_is_renamed_from_the_old_name(self):
        """В старом отчёте торговый портфель — TSS, в схеме v3.0 — TTS."""
        frame = self.parse()
        self.assertIn("TTS", set(frame["portfolio_type"]))
        self.assertNotIn("TSS", set(frame["portfolio_type"]))

    def test_roubles_are_converted_to_millions(self):
        frame = self.parse().set_index(["business_date", "portfolio_type"])
        self.assertAlmostEqual(frame.loc[(DAYS[0], "AFS"), "volume_amount"], 300)

    def test_dates_are_read_both_as_text_and_as_real_dates(self):
        for as_text in (True, False):
            with self.subTest(date_as_text=as_text):
                frame = self.parse(date_as_text=as_text)
                self.assertEqual(sorted(set(frame["business_date"])), DAYS)

    def test_header_is_found_by_content_not_by_row_number(self):
        frame = self.parse(header_row=7)
        self.assertEqual(len(frame), 9)

    def test_change_column_is_ignored(self):
        """«Изменение» отчёт считает сам — в данные оно не попадает."""
        frame = self.parse()
        self.assertEqual(list(frame.columns), etl.TYPE_DAILY_COLUMNS)

    def test_sheets_that_are_not_dynamics_are_skipped(self):
        path = write_history_file(self.tmp / "старый отчёт.xlsx", DEFAULT_SHEETS)
        from openpyxl import load_workbook
        wb = load_workbook(path)
        extra = wb.create_sheet("Служебный лист")
        extra["A3"] = "Дата"
        extra["B3"] = "Текущий объём"
        extra["A4"] = "01.09.2026"
        extra["B4"] = 999 * RUB
        wb.save(path)

        frame = history.parse_history_file(path)
        self.assertEqual(sorted(frame["portfolio_type"].unique()), ["AFS", "HTM", "TTS"])

    def test_sheet_without_required_columns_is_reported(self):
        wb = Workbook()
        wb.remove(wb.active)
        ws = wb.create_sheet("Динамика AFS")
        ws["A1"] = "Совсем другие колонки"
        good = wb.create_sheet("Динамика HTM")
        for j, title in enumerate(["Дата", "Текущий объём"], start=1):
            good.cell(row=1, column=j, value=title)
        good["A2"] = "01.09.2026"
        good["B2"] = 600 * RUB
        path = self.tmp / "кривой.xlsx"
        wb.save(path)

        with self.assertLogs("portfolio_dynamics", level="WARNING") as captured:
            frame = history.parse_history_file(path)

        self.assertEqual(list(frame["portfolio_type"]), ["HTM"])
        self.assertTrue(any("Динамика AFS" in line for line in captured.output))

    def test_file_without_any_dynamics_sheets_says_what_it_saw(self):
        wb = Workbook()
        wb.active.title = "Просто лист"
        path = self.tmp / "чужой.xlsx"
        wb.save(path)
        with self.assertRaisesRegex(etl.PortfolioDynamicsError, "Динамика"):
            history.parse_history_file(path)

    def test_duplicate_dates_keep_the_last_row(self):
        sheets = {"AFS": [(DAYS[0], 300 * RUB), (DAYS[0], 350 * RUB)]}
        with self.assertLogs("portfolio_dynamics", level="WARNING"):
            frame = self.parse(sheets)
        self.assertEqual(len(frame), 1)
        self.assertAlmostEqual(frame.iloc[0]["volume_amount"], 350)

    def test_rows_without_a_date_or_volume_are_skipped(self):
        sheets = {"AFS": [(DAYS[0], 300 * RUB), (DAYS[1], None)]}
        frame = self.parse(sheets)
        self.assertEqual(len(frame), 1)

    def test_aliases_are_configurable(self):
        settings.set_value("portfolio_dynamics_history_aliases", "TSS=HFT")
        config.reload()
        frame = self.parse()
        self.assertIn("HFT", set(frame["portfolio_type"]))


class MergeTests(HistoryTestCase):
    def frame(self, rows):
        return pd.DataFrame(rows, columns=etl.TYPE_DAILY_COLUMNS)

    def test_import_only_fills_gaps(self):
        existing = self.frame([[DAYS[0], "AFS", 999]])
        imported = self.frame([[DAYS[0], "AFS", 111], [DAYS[1], "AFS", 222]])

        merged, added, skipped = history.merge_into(existing, imported)

        self.assertEqual((added, skipped), (1, 1))
        by_date = dict(zip(merged["business_date"], merged["volume_amount"]))
        self.assertEqual(by_date[DAYS[0]], 999, "накопленное своим запуском не перезаписывается")
        self.assertEqual(by_date[DAYS[1]], 222)

    def test_empty_existing_takes_everything(self):
        imported = self.frame([[DAYS[0], "AFS", 111]])
        merged, added, skipped = history.merge_into(pd.DataFrame(columns=etl.TYPE_DAILY_COLUMNS),
                                                    imported)
        self.assertEqual((len(merged), added, skipped), (1, 1, 0))

    def test_nothing_to_import_changes_nothing(self):
        existing = self.frame([[DAYS[0], "AFS", 999]])
        merged, added, skipped = history.merge_into(existing, pd.DataFrame())
        self.assertEqual((added, skipped), (0, 0))
        self.assertEqual(len(merged), 1)

    def test_scale_mismatch_is_warned_about(self):
        """Ошибка в единицах в миллион раз не видна глазом, но видна при сверке."""
        imported = self.frame([[DAYS[0], "AFS", 300_000_000]])
        with self.assertLogs("portfolio_dynamics", level="WARNING") as captured:
            history.warn_if_scale_looks_wrong(imported, {"AFS": 300})
        self.assertTrue(any("единицы" in line for line in captured.output))

    def test_matching_scale_is_silent(self):
        imported = self.frame([[DAYS[0], "AFS", 305]])
        with self.assertRaises(AssertionError):
            with self.assertLogs("portfolio_dynamics", level="WARNING"):
                history.warn_if_scale_looks_wrong(imported, {"AFS": 300})


class EndToEndTests(HistoryTestCase):
    def setUp(self):
        super().setUp()
        from test_portfolio_dynamics_etl import export_name
        portfolios = [("AFS_A", 300), ("HTM_G", 600), ("TTS_OFZ", 200)]

        def rows(scale):
            out = []
            for code, volume in portfolios:
                out += [(f"Позиция: {code}", None, None, None),
                        ("Bond", volume * scale * MLN, 3.0, 2.5)]
            return out

        self.t0 = write_export(self.tmp / export_name("21.09.2026"), rows(1.0),
                               period_end="21.09.2026")
        self.t7 = write_export(self.tmp / export_name("14.09.2026"), rows(0.97),
                               period_end="14.09.2026")
        self.history_file = write_history_file(self.tmp / "старый отчёт.xlsx", DEFAULT_SHEETS)

    def build(self, **kwargs):
        return etl.build_data(self.t0, self.t7, bootstrap=True, **kwargs)

    def test_history_lands_in_fact_type_daily(self):
        data = self.build(history_path=self.history_file)

        self.assertEqual(data.history_rows_imported, 9)
        self.assertEqual(len(data.fact_type_daily), 12)  # 9 импортированных + 3 за сегодня
        self.assertFalse(data.fact_type_daily.duplicated(
            ["business_date", "portfolio_type"]).any())

    def test_todays_row_is_still_computed_from_the_export(self):
        """Импорт не должен подменять дату, которую отчёт считает сам."""
        data = self.build(history_path=self.history_file)
        today = data.fact_type_daily[
            data.fact_type_daily["business_date"] == data.business_date]
        self.assertEqual(dict(zip(today["portfolio_type"], today["volume_amount"])),
                         {"AFS": 300, "HTM": 600, "TTS": 200})

    def test_without_history_nothing_changes(self):
        data = self.build()
        self.assertEqual(data.history_rows_imported, 0)
        self.assertEqual(len(data.fact_type_daily), 3)

    def test_second_import_adds_nothing(self):
        first = self.build(history_path=self.history_file)
        again, added, skipped = history.merge_into(
            first.fact_type_daily, history.parse_history_file(self.history_file))
        self.assertEqual((added, skipped), (0, 9))
        self.assertEqual(len(again), len(first.fact_type_daily))


if __name__ == "__main__":
    unittest.main()
