"""Тесты раскладки «папка на каждую дату»: поиск, создание и разбор содержимого.

Проверяется то, ради чего раскладка и вводилась: отчёт без аргументов должен
брать последнюю папку, В КОТОРОЙ ЕСТЬ файлы (а не завтрашнюю пустую, созданную
заранее), и сам понимать, какой из двух файлов T0, а какой T-7 — независимо от
того, как они названы и в каком порядке лежат.
"""
import argparse
import datetime as dt
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE_DIR))
sys.path.insert(0, str(BASE_DIR / "tests"))

import config  # noqa: E402
from common import file_discovery, settings  # noqa: E402
from reports.portfolio_dynamics import etl  # noqa: E402
from reports.portfolio_dynamics import report as pd_report  # noqa: E402
from test_portfolio_dynamics_etl import MLN, export_name, write_export  # noqa: E402


def slice_rows(afs: float, htm: float):
    return [
        ("Позиция: AFS_TR_RUR", None, None, None),
        ("Bond", afs * MLN, 3.0, 2.5),
        ("Позиция: HTM_ALCO", None, None, None),
        ("Bond", htm * MLN, 6.0, 5.5),
    ]


class DateFolderTestCase(unittest.TestCase):
    """Источник «Динамики портфелей» переставлен во временную папку."""

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.data = self.tmp / "data"
        self.data.mkdir()

        self._saved_env = os.environ.get(settings.SETTINGS_FILE_ENV)
        settings_file = self.tmp / "settings.json"
        settings_file.write_text(json.dumps({
            "portfolio_dynamics_dir": str(self.data),
            "portfolio_dynamics_output_dir": str(self.tmp / "out"),
        }), encoding="utf-8")
        os.environ[settings.SETTINGS_FILE_ENV] = str(settings_file)
        config.reload()
        self.source = config.PORTFOLIO_DYNAMICS_T0_SOURCE

    def tearDown(self):
        if self._saved_env is None:
            os.environ.pop(settings.SETTINGS_FILE_ENV, None)
        else:
            os.environ[settings.SETTINGS_FILE_ENV] = self._saved_env
        config.reload()
        self._tmp.cleanup()

    def make_folder(self, name: str, exports=()) -> Path:
        """Папка-дата с выгрузками; exports — список дат вида «18.09.2026»."""
        folder = self.data / name
        folder.mkdir(parents=True, exist_ok=True)
        for i, period_end in enumerate(exports):
            write_export(folder / export_name(period_end), slice_rows(30 + i, 60 + i),
                         period_end=period_end)
        return folder

    def no_args(self) -> argparse.Namespace:
        return argparse.Namespace(t0_input=None, t7_input=None, folder=None,
                                  t0_date=None, t7_date=None)


class DiscoveryTests(DateFolderTestCase):
    def test_empty_folders_created_in_advance_are_skipped(self):
        """Папки на будущие даты созданы заранее и пусты — последней считается не они."""
        self.make_folder("2026-09-18", ["11.09.2026", "18.09.2026"])
        self.make_folder("2026-09-21")
        self.make_folder("2026-09-22")

        folder = file_discovery.latest_folder_with_data(self.source, min_files=2)

        self.assertIsNotNone(folder)
        self.assertEqual(folder.date, dt.date(2026, 9, 18))
        self.assertEqual(len(folder.files), 2)

    def test_folder_with_one_file_is_not_enough(self):
        self.make_folder("2026-09-18", ["18.09.2026"])
        self.assertIsNone(file_discovery.latest_folder_with_data(self.source, min_files=2))
        self.assertIsNotNone(file_discovery.latest_folder_with_data(self.source, min_files=1))

    def test_newer_folder_with_data_wins(self):
        self.make_folder("2026-09-18", ["11.09.2026", "18.09.2026"])
        self.make_folder("2026-09-25", ["18.09.2026", "25.09.2026"])

        folder = file_discovery.latest_folder_with_data(self.source, min_files=2)
        self.assertEqual(folder.date, dt.date(2026, 9, 25))

    def test_excel_temp_files_are_not_data(self):
        """«~$Отчёт.xlsx» появляется рядом с открытым файлом и данными не является."""
        folder = self.make_folder("2026-09-18", ["18.09.2026"])
        (folder / "~$временный.xlsx").write_bytes(b"")

        self.assertIsNone(file_discovery.latest_folder_with_data(self.source, min_files=2))

    def test_folders_that_are_not_dates_are_ignored(self):
        self.make_folder("архив", ["18.09.2026", "11.09.2026"])
        self.make_folder("2026-09-18", ["11.09.2026", "18.09.2026"])

        folders = file_discovery.find_date_folders(self.source)
        self.assertEqual([f.path.name for f in folders], ["2026-09-18"])

    def test_files_in_folders_are_dated_by_the_folder(self):
        """Дата берётся из имени папки — как назван файл внутри, уже не важно."""
        self.make_folder("2026-09-18", ["11.09.2026", "18.09.2026"])

        dates = {d for d, _p in file_discovery.find_dated_files(self.source)}
        self.assertEqual(dates, {dt.date(2026, 9, 18)})

    def test_flat_layout_still_works_alongside_folders(self):
        """Переходить на папки можно постепенно: обе раскладки видны одновременно."""
        write_export(self.data / export_name("04.09.2026"), slice_rows(20, 40),
                     period_end="04.09.2026")
        self.make_folder("2026-09-18", ["11.09.2026", "18.09.2026"])

        dates = sorted({d for d, _p in file_discovery.find_dated_files(self.source)})
        self.assertEqual(dates, [dt.date(2026, 9, 4), dt.date(2026, 9, 18)])

    def test_folders_are_invisible_when_the_setting_is_off(self):
        self.make_folder("2026-09-18", ["11.09.2026", "18.09.2026"])
        settings.set_value("portfolio_dynamics_use_date_folders", "нет")
        config.reload()

        self.assertFalse(config.PORTFOLIO_DYNAMICS_T0_SOURCE.uses_date_folders)
        self.assertEqual(file_discovery.find_date_folders(config.PORTFOLIO_DYNAMICS_T0_SOURCE), [])
        self.assertEqual(file_discovery.find_dated_files(config.PORTFOLIO_DYNAMICS_T0_SOURCE), [])


class CreateFoldersTests(DateFolderTestCase):
    def test_weekends_are_skipped_by_default(self):
        # 2026-09-18 — пятница, 19 и 20 — выходные.
        created, existed = file_discovery.create_date_folders(
            self.source, days=7, start=dt.date(2026, 9, 18)
        )

        self.assertEqual([f.name for f in created],
                         ["2026-09-18", "2026-09-21", "2026-09-22", "2026-09-23", "2026-09-24"])
        self.assertEqual(existed, [])

    def test_weekends_can_be_included(self):
        created, _existed = file_discovery.create_date_folders(
            self.source, days=3, include_weekends=True, start=dt.date(2026, 9, 18)
        )
        self.assertEqual([f.name for f in created], ["2026-09-18", "2026-09-19", "2026-09-20"])

    def test_running_twice_creates_nothing_new(self):
        file_discovery.create_date_folders(self.source, days=5, start=dt.date(2026, 9, 18))
        created, existed = file_discovery.create_date_folders(
            self.source, days=5, start=dt.date(2026, 9, 18)
        )

        self.assertEqual(created, [])
        # 5 календарных дней с пятницы 18.09 — это 3 рабочих: 18, 21, 22.
        self.assertEqual([f.name for f in existed], ["2026-09-18", "2026-09-21", "2026-09-22"])

    def test_existing_folder_with_data_is_not_touched(self):
        folder = self.make_folder("2026-09-18", ["11.09.2026", "18.09.2026"])
        file_discovery.create_date_folders(self.source, days=5, start=dt.date(2026, 9, 18))

        self.assertEqual(len(sorted(folder.glob("*.xlsx"))), 2)

    def test_refuses_when_date_folders_are_off(self):
        settings.set_value("portfolio_dynamics_use_date_folders", "нет")
        config.reload()

        with self.assertRaisesRegex(file_discovery.SourceFileError, "папки-даты"):
            file_discovery.create_date_folders(config.PORTFOLIO_DYNAMICS_T0_SOURCE, days=5)

    def test_zero_days_is_rejected(self):
        with self.assertRaisesRegex(file_discovery.SourceFileError, "больше нуля"):
            file_discovery.create_date_folders(self.source, days=0)


class SplitSliceFilesTests(DateFolderTestCase):
    def test_t0_is_the_later_export_regardless_of_file_order(self):
        """Порядок и имена файлов не важны — важна дата самой выгрузки."""
        folder = self.make_folder("2026-09-18", ["18.09.2026", "11.09.2026"])
        files = sorted(folder.glob("*.xlsx"))  # по алфавиту T-7 идёт первым

        t0, t7 = etl.split_slice_files(files, folder.name)

        self.assertIn("18.09.2026", t0.name)
        self.assertIn("11.09.2026", t7.name)

    def test_one_file_is_not_enough(self):
        folder = self.make_folder("2026-09-18", ["18.09.2026"])
        with self.assertRaisesRegex(etl.PortfolioDynamicsError, "нужны ДВА файла"):
            etl.split_slice_files(sorted(folder.glob("*.xlsx")), folder.name)

    def test_two_slices_on_the_same_date_are_rejected(self):
        folder = self.make_folder("2026-09-18")
        write_export(folder / "первый.xlsx", slice_rows(30, 60), period_end="18.09.2026")
        write_export(folder / "второй.xlsx", slice_rows(31, 61), period_end="18.09.2026")

        with self.assertRaisesRegex(etl.PortfolioDynamicsError, "одну и ту же дату"):
            etl.split_slice_files(sorted(folder.glob("*.xlsx")), folder.name)

    def test_extra_files_are_ignored_with_a_warning(self):
        folder = self.make_folder("2026-09-18", ["04.09.2026", "11.09.2026", "18.09.2026"])

        with self.assertLogs("portfolio_dynamics", level="WARNING") as captured:
            t0, t7 = etl.split_slice_files(sorted(folder.glob("*.xlsx")), folder.name)

        self.assertIn("18.09.2026", t0.name)
        self.assertIn("11.09.2026", t7.name)
        self.assertTrue(any("04.09.2026" in line for line in captured.output))

    def test_date_is_read_from_the_sheet_when_the_name_has_none(self):
        """Файл переименовали руками — дата всё равно находится, в шапке листа."""
        folder = self.make_folder("2026-09-18")
        write_export(folder / "срез сегодня.xlsx", slice_rows(30, 60), period_end="18.09.2026")
        write_export(folder / "срез неделю назад.xlsx", slice_rows(28, 55), period_end="11.09.2026")

        t0, t7 = etl.split_slice_files(sorted(folder.glob("*.xlsx")), folder.name)

        self.assertEqual(t0.name, "срез сегодня.xlsx")
        self.assertEqual(t7.name, "срез неделю назад.xlsx")


class ReportResolutionTests(DateFolderTestCase):
    def test_no_arguments_picks_the_latest_folder_with_data(self):
        self.make_folder("2026-09-11", ["04.09.2026", "11.09.2026"])
        self.make_folder("2026-09-18", ["11.09.2026", "18.09.2026"])
        self.make_folder("2026-09-25")  # создана заранее, пустая

        t0, t7 = pd_report._resolve_slice_paths(self.no_args())

        self.assertEqual(t0.parent.name, "2026-09-18")
        self.assertIn("18.09.2026", t0.name)
        self.assertIn("11.09.2026", t7.name)

    def test_folder_argument_accepts_a_date(self):
        self.make_folder("2026-09-11", ["04.09.2026", "11.09.2026"])
        self.make_folder("2026-09-18", ["11.09.2026", "18.09.2026"])

        args = self.no_args()
        args.folder = "2026-09-11"
        t0, _t7 = pd_report._resolve_slice_paths(args)

        self.assertEqual(t0.parent.name, "2026-09-11")

    def test_folder_argument_accepts_a_path(self):
        folder = self.make_folder("2026-09-18", ["11.09.2026", "18.09.2026"])

        args = self.no_args()
        args.folder = str(folder)
        t0, _t7 = pd_report._resolve_slice_paths(args)

        self.assertEqual(t0.parent, folder)

    def test_flat_layout_falls_back_to_two_latest_files(self):
        """Без папок-дат берутся два самых свежих файла: новее — T0, предыдущий — T-7."""
        for period_end in ("04.09.2026", "11.09.2026", "18.09.2026"):
            write_export(self.data / export_name(period_end), slice_rows(30, 60),
                         period_end=period_end)

        t0, t7 = pd_report._resolve_slice_paths(self.no_args())

        self.assertIn("18.09.2026", t0.name)
        self.assertIn("11.09.2026", t7.name)

    def test_single_file_says_what_to_do(self):
        write_export(self.data / export_name("18.09.2026"), slice_rows(30, 60),
                     period_end="18.09.2026")

        with self.assertRaisesRegex(etl.PortfolioDynamicsError, "Создать папки по датам"):
            pd_report._resolve_slice_paths(self.no_args())

    def test_one_explicit_input_without_the_other_is_rejected(self):
        args = self.no_args()
        args.t0_input = "/tmp/a.xlsx"

        with self.assertRaisesRegex(etl.PortfolioDynamicsError, "только вместе"):
            pd_report._resolve_slice_paths(args)


if __name__ == "__main__":
    unittest.main()
