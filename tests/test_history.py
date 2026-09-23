"""Тесты импорта накопленной истории из отчёта СТАРОГО формата.

Старый формат: по листу на тип («Динамика AFS», «Динамика HTM», «Динамика
TSS»), в каждом колонки «Дата», «Текущий объём», «Изменение». Проверяется то,
что при переносе легко испортить незаметно: имя торгового типа (TSS против
TSS), единицы, и главное — что импорт только дополняет и никогда не
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
    "TSS": [(d, (200 + i) * RUB) for i, d in enumerate(DAYS)],  # торговый портфель
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
        self.assertEqual(sorted(frame["portfolio_type"].unique()), ["AFS", "HTM", "TSS"])

    def test_old_spelling_of_the_trading_type_is_accepted(self):
        """Лист мог называться и «Динамика TTS» — прежнее написание понимается."""
        sheets = {"TTS": [(DAYS[0], 200 * RUB)]}
        frame = self.parse(sheets)
        self.assertEqual(list(frame["portfolio_type"]), ["TSS"])

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
        self.assertEqual(sorted(frame["portfolio_type"].unique()), ["AFS", "HTM", "TSS"])

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
        portfolios = [("AFS_A", 300), ("HTM_G", 600), ("TSS_OFZ", 200)]

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
                         {"AFS": 300, "HTM": 600, "TSS": 200})

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


class FindCandidatesTests(HistoryTestCase):
    """Поиск отчёта старого формата по СОДЕРЖИМОМУ.

    Имя у него произвольное («Динамика портфелей (2).xlsx», «КУАП итог.xlsx»),
    и требовать от человека вспомнить и набрать путь — ровно та причина, по
    которой возможность импорта было не найти. Листы «Динамика <ТИП>» видны в
    оглавлении книги, поэтому искать можно без всяких догадок об имени.
    """

    def setUp(self):
        super().setUp()
        self.downloads = self.tmp / "downloads"
        self.downloads.mkdir(exist_ok=True)

    def test_old_report_is_found_whatever_it_is_called(self):
        write_history_file(self.downloads / "какой-то файл (3).xlsx", DEFAULT_SHEETS)

        found = history.find_candidates([self.downloads])

        self.assertEqual([p.name for p, _sheets in found], ["какой-то файл (3).xlsx"])

    def test_the_sheets_are_reported(self):
        """Человек должен видеть, ЧТО в файле, а не только его имя."""
        write_history_file(self.downloads / "архив.xlsx", DEFAULT_SHEETS)

        _path, sheets = history.find_candidates([self.downloads])[0]

        self.assertEqual(sorted(sheets),
                         ["Динамика AFS", "Динамика HTM", "Динамика TSS"])

    def test_a_workbook_without_such_sheets_is_not_offered(self):
        write_export(self.downloads / "позиции.xlsx",
                     [("Позиция: AFS_A", None, None, None), ("Bond", 1 * MLN, 3.0, 2.5)],
                     period_end="18.09.2026")

        self.assertEqual(history.find_candidates([self.downloads]), [])

    def test_a_broken_file_is_skipped_quietly(self):
        (self.downloads / "битый.xlsx").write_bytes(b"not a workbook")
        write_history_file(self.downloads / "архив.xlsx", DEFAULT_SHEETS)

        found = history.find_candidates([self.downloads])

        self.assertEqual([p.name for p, _s in found], ["архив.xlsx"])

    def test_missing_folder_is_not_an_error(self):
        self.assertEqual(history.find_candidates([self.tmp / "нет такой"]), [])

    def test_several_folders_are_searched_without_repeats(self):
        write_history_file(self.downloads / "архив.xlsx", DEFAULT_SHEETS)

        found = history.find_candidates([self.downloads, self.downloads])

        self.assertEqual(len(found), 1)

    def test_what_is_found_can_actually_be_imported(self):
        """Находить файл, который потом не читается, — хуже, чем не находить."""
        write_history_file(self.downloads / "архив.xlsx", DEFAULT_SHEETS)

        path, _sheets = history.find_candidates([self.downloads])[0]
        frame = history.parse_history_file(path)

        self.assertEqual(sorted(frame["portfolio_type"].unique()), ["AFS", "HTM", "TSS"])


class ResolvePathTests(HistoryTestCase):
    """Путь из настройки часто набран так, как его показывает Проводник.

    Проводник Windows скрывает расширения, и в настройку попадает
    «…\\Лимиты портфелей казначейства (1)» без «.xlsx». Из-за этого отчёт
    падал целиком, хотя файл лежал ровно там, куда указывал путь.
    """

    def setUp(self):
        super().setUp()
        self.downloads = self.tmp / "downloads"
        self.downloads.mkdir(exist_ok=True)
        self.file = write_history_file(
            self.downloads / "Лимиты портфелей казначейства (1).xlsx", DEFAULT_SHEETS)

    def test_path_without_extension_is_completed(self):
        raw = str(self.downloads / "Лимиты портфелей казначейства (1)")

        self.assertEqual(history.resolve_path(raw), self.file)

    def test_extension_is_appended_not_replaced(self):
        """with_suffix превратил бы «отчёт 22.09» в «отчёт 22.xlsx»."""
        dotted = write_history_file(self.downloads / "отчёт 22.09.xlsx", DEFAULT_SHEETS)

        self.assertEqual(history.resolve_path(str(self.downloads / "отчёт 22.09")), dotted)

    def test_quoted_path_is_accepted(self):
        self.assertEqual(history.resolve_path('"%s"' % self.file), self.file)

    def test_moved_file_is_found_by_name_in_search_folders(self):
        moved_to = self.tmp / "data"
        moved_to.mkdir()
        moved = moved_to / self.file.name
        self.file.rename(moved)

        self.assertEqual(history.resolve_path(str(self.file), [self.downloads, moved_to]), moved)

    def test_missing_file_gives_none(self):
        self.assertIsNone(history.resolve_path(str(self.downloads / "нет такого")))
        self.assertIsNone(history.resolve_path(""))

    def test_parse_accepts_path_without_extension(self):
        frame = history.parse_history_file(self.downloads / "Лимиты портфелей казначейства (1)")

        self.assertEqual(sorted(frame["portfolio_type"].unique()), ["AFS", "HTM", "TSS"])


class ImportOnceTests(HistoryTestCase):
    """Импорт разовый: накопилась история в выпуске — файл больше не читается.

    Иначе путь из настройки читался бы на каждом запуске, а удалённый из
    загрузок файл давал бы предупреждение каждый день.
    """

    def previous_with(self, dates: int) -> etl.PreviousRelease:
        previous = etl.PreviousRelease.empty()
        previous.fact_type_daily = pd.DataFrame(
            [{"business_date": dt.date(2026, 8, 1) + dt.timedelta(days=i),
              "portfolio_type": "AFS", "volume_amount": 300.0} for i in range(dates)],
            columns=etl.TYPE_DAILY_COLUMNS,
        )
        return previous

    def setUp(self):
        super().setUp()
        (self.tmp / "downloads").mkdir(exist_ok=True)
        self.file = write_history_file(self.tmp / "downloads" / "Лимиты (1).xlsx", DEFAULT_SHEETS)

    def test_short_history_is_topped_up_from_file(self):
        previous = self.previous_with(etl.HISTORY_ACCUMULATED_DATES - 1)

        self.assertEqual(etl._import_history(previous, str(self.file)), 9)

    def test_accumulated_history_is_carried_without_reading_file(self):
        previous = self.previous_with(etl.HISTORY_ACCUMULATED_DATES)
        before = previous.fact_type_daily.copy()

        self.assertEqual(etl._import_history(previous, str(self.file)), 0)
        pd.testing.assert_frame_equal(previous.fact_type_daily, before)

    def test_missing_file_is_silent_once_history_is_accumulated(self):
        previous = self.previous_with(etl.HISTORY_ACCUMULATED_DATES)

        # assertNoLogs появился только в 3.10, а на рабочих машинах 3.9.
        with self.assertLogs(etl.logger, level="INFO") as captured:
            etl._import_history(previous, str(self.tmp / "удалён.xlsx"))
        self.assertFalse([r for r in captured.records if r.levelname == "WARNING"])

    def test_explicit_path_is_imported_anyway(self):
        previous = self.previous_with(etl.HISTORY_ACCUMULATED_DATES)

        self.assertEqual(etl._import_history(previous, str(self.file), force=True), 9)

    def test_path_without_extension_works(self):
        previous = self.previous_with(0)

        self.assertEqual(etl._import_history(previous, str(self.tmp / "downloads" / "Лимиты (1)")), 9)

    def test_broken_setting_is_skipped_with_warning(self):
        previous = self.previous_with(0)

        with self.assertLogs(etl.logger, level="WARNING"):
            self.assertEqual(etl._import_history(previous, str(self.tmp / "нет такого")), 0)

    def test_explicit_argument_must_exist(self):
        with self.assertRaises(etl.PortfolioDynamicsError):
            etl._import_history(self.previous_with(0), str(self.tmp / "нет такого"), force=True)

    def test_report_marks_cli_argument_as_explicit(self):
        import argparse
        from reports.portfolio_dynamics import report
        settings.set_value("portfolio_dynamics_history_file", str(self.file))
        config.reload()

        self.assertEqual(report._history_request(argparse.Namespace(history=None)),
                         (str(self.file), False))
        self.assertEqual(report._history_request(argparse.Namespace(history="x.xlsx")),
                         ("x.xlsx", True))
