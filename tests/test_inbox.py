"""Тесты приёмки выгрузок из папки загрузок (reports/portfolio_dynamics/inbox.py).

Проверяется то, что при переносе файлов ломается дороже всего: чтобы забирались
ровно нужные два файла и ровно на нужную дату, чтобы посторонние файлы в чужой
папке никто не трогал, чтобы повторный запуск ничего не портил и чтобы при
выключенной приёмке загрузки не открывались вовсе.
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
from common import settings  # noqa: E402
from reports.portfolio_dynamics import etl, inbox  # noqa: E402
from reports.portfolio_dynamics import report as pd_report  # noqa: E402
from test_date_folders import slice_rows  # noqa: E402
from test_portfolio_dynamics_etl import export_name, write_export  # noqa: E402


class InboxTestCase(unittest.TestCase):
    """Папка отчёта и папка загрузок — во временном каталоге."""

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.data = self.tmp / "data"
        self.downloads = self.tmp / "downloads"
        self.data.mkdir()
        self.downloads.mkdir()

        self._saved_env = os.environ.get(settings.SETTINGS_FILE_ENV)
        settings_file = self.tmp / "settings.json"
        settings_file.write_text(json.dumps({
            "portfolio_dynamics_dir": str(self.data),
            "portfolio_dynamics_output_dir": str(self.tmp / "out"),
            "downloads_dir": str(self.downloads),
        }), encoding="utf-8")
        os.environ[settings.SETTINGS_FILE_ENV] = str(settings_file)
        config.reload()
        inbox.forget_probe_cache()
        self.source = config.PORTFOLIO_DYNAMICS_T0_SOURCE

    def tearDown(self):
        if self._saved_env is None:
            os.environ.pop(settings.SETTINGS_FILE_ENV, None)
        else:
            os.environ[settings.SETTINGS_FILE_ENV] = self._saved_env
        config.reload()
        self._tmp.cleanup()

    def download(self, period_end: str, volume: float = 30) -> Path:
        return write_export(self.downloads / export_name(period_end),
                            slice_rows(volume, volume * 2), period_end=period_end)

    def in_folder(self, folder_name: str, period_end: str, volume: float = 30) -> Path:
        return write_export(self.data / folder_name / export_name(period_end),
                            slice_rows(volume, volume * 2), period_end=period_end)

    def names(self, folder_name: str):
        folder = self.data / folder_name
        return sorted(f.name for f in folder.glob("*.xlsx")) if folder.is_dir() else []

    def downloads_names(self):
        return sorted(f.name for f in self.downloads.iterdir())

    def args(self, **kwargs) -> argparse.Namespace:
        base = dict(t0_input=None, t7_input=None, folder=None, date=None,
                    t0_date=None, t7_date=None, no_import=False)
        base.update(kwargs)
        return argparse.Namespace(**base)


class ScanTests(InboxTestCase):
    def test_only_files_matching_the_export_pattern_are_ours(self):
        self.download("18.09.2026")
        (self.downloads / "отчёт из другой системы.xlsx").write_bytes(b"not ours")
        (self.downloads / "~$Позиция за период [01.01.2026] - [18.09.2026].xlsx").write_bytes(b"")

        found = inbox.scan_slices(self.source, self.downloads)

        self.assertEqual(len(found), 1)
        self.assertEqual(found[0].business_date, dt.date(2026, 9, 18))

    def test_missing_downloads_folder_is_not_an_error(self):
        """Папку загрузок могли не примонтировать — это не повод падать."""
        self.assertEqual(inbox.scan_slices(self.source, self.tmp / "нет такой"), [])


class ContentProbeTests(InboxTestCase):
    """Выгрузку, которую ПЕРЕИМЕНОВАЛИ, отчёт всё равно обязан находить.

    Имя — не свойство данных: браузер дописывает «(1)», человек правит имя
    руками, выгружающая система однажды меняет формулировку. Пока поиск шёл
    только по шаблону имени, любой такой случай означал «файл лежит в
    загрузках, а отчёт его не видит» — без единого слова о причине.
    """

    def renamed(self, period_end: str, name: str) -> Path:
        return write_export(self.downloads / name, slice_rows(30, 60),
                            period_end=period_end)

    def test_renamed_export_is_found_by_its_contents(self):
        self.renamed("18.09.2026", "выгрузка (1).xlsx")

        found = inbox.scan_slices(self.source, self.downloads)

        self.assertEqual([c.business_date for c in found], [dt.date(2026, 9, 18)])

    def test_the_find_is_announced(self):
        """Молча брать файл с «неправильным» именем нельзя: человек должен
        понимать, почему отчёт собрался именно из него."""
        self.renamed("18.09.2026", "выгрузка (1).xlsx")

        with self.assertLogs("portfolio_dynamics", level="INFO") as captured:
            inbox.scan_slices(self.source, self.downloads)

        self.assertTrue(any("выгрузка (1).xlsx" in line and "содержимому" in line
                            for line in captured.output))

    def test_a_foreign_workbook_is_not_taken(self):
        """Признак — шапка выгрузки, а не просто дата: диапазон дат есть в любом отчёте."""
        from openpyxl import Workbook
        workbook = Workbook()
        workbook.active["A1"] = "Отчёт за период [01.01.2026] - [18.09.2026]"
        workbook.active["A3"] = "Совсем другие колонки"
        workbook.save(self.downloads / "чужой отчёт.xlsx")

        self.assertEqual(inbox.scan_slices(self.source, self.downloads), [])

    def test_a_broken_file_is_not_an_error(self):
        (self.downloads / "битый.xlsx").write_bytes(b"not a workbook at all")
        self.renamed("18.09.2026", "выгрузка (1).xlsx")

        found = inbox.scan_slices(self.source, self.downloads)

        self.assertEqual([c.path.name for c in found], ["выгрузка (1).xlsx"])

    def test_a_renamed_pair_gives_a_working_run(self):
        """Главное: переименованные файлы доходят до отчёта, а не просто находятся."""
        self.renamed("18.09.2026", "позиции свежие.xlsx")
        self.renamed("11.09.2026", "позиции прошлые.xlsx")

        t0, t7 = pd_report._resolve_slice_paths(self.args())

        self.assertEqual(etl.read_business_date(t0), dt.date(2026, 9, 18))
        self.assertEqual(etl.read_business_date(t7), dt.date(2026, 9, 11))

    def test_a_position_export_is_never_mistaken_for_a_limits_file(self):
        """У двух выгрузок разные шапки, и поиск по содержимому обязан их различать."""
        self.renamed("18.09.2026", "выгрузка (1).xlsx")

        found = inbox.scan_limits(config.PORTFOLIO_DYNAMICS_LIMITS_SOURCE, self.downloads)

        self.assertEqual(found, [])

    def test_name_search_alone_can_be_asked_for(self):
        """deep=False нужен диагностике, чтобы показать, что дало имя, а что содержимое."""
        self.renamed("18.09.2026", "выгрузка (1).xlsx")

        self.assertEqual(inbox.scan_slices(self.source, self.downloads, deep=False), [])


class RealFilenameTests(InboxTestCase):
    """Имена ровно в том виде, в каком их отдаёт выгрузка: без квадратных
    скобок, с несколькими пробелами, одна и та же дата дважды."""

    def real_download(self, day: str, volume: float = 30) -> Path:
        name = f"Позиция за период   {day}  -  {day}   - SECURITIES.xlsx"
        return write_export(self.downloads / name, slice_rows(volume, volume * 2),
                            period_end=day)

    def test_export_filenames_are_recognised(self):
        self.real_download("18.09.2026")
        self.real_download("11.09.2026")

        found = {c.business_date for c in inbox.scan_slices(self.source, self.downloads)}
        self.assertEqual(found, {dt.date(2026, 9, 18), dt.date(2026, 9, 11)})

    def test_full_run_from_downloads(self):
        self.real_download("18.09.2026")
        self.real_download("11.09.2026")

        t0, t7 = pd_report._resolve_slice_paths(self.args())

        self.assertEqual(t0.parent.name, "2026-09-18")
        self.assertIn("18.09.2026", t0.name)
        self.assertIn("11.09.2026", t7.name)
        self.assertEqual(self.downloads_names(), [], "файлы должны уехать из загрузок")


class LimitsImportTests(InboxTestCase):
    """Файл лимитов — третий файл на отчётную дату, едет в ту же папку."""

    def limits_download(self, day_underscored: str) -> Path:
        from test_limits import DEFAULT_ROWS, write_limits_file
        return write_limits_file(
            self.downloads / f"Состояние лимитов на дату {day_underscored} - Результат.xlsx",
            DEFAULT_ROWS)

    def test_limits_file_is_imported_together_with_the_slices(self):
        self.download("21.09.2026")
        self.download("14.09.2026")
        self.limits_download("21_09_2026")

        plan = inbox.plan_import(self.source, self.downloads, dt.date(2026, 9, 21),
                                 limits_source=config.PORTFOLIO_DYNAMICS_LIMITS_SOURCE)
        inbox.apply_import(plan, move=True)

        self.assertIsNotNone(plan.limits)
        self.assertEqual(len(self.names("2026-09-21")), 3)
        self.assertEqual(self.downloads_names(), [])

    def test_earlier_limits_file_is_used_with_a_warning(self):
        """Лимиты меняются редко: отчёт с чуть устаревшими полезнее отчёта без них."""
        self.download("21.09.2026")
        self.download("14.09.2026")
        self.limits_download("14_09_2026")

        with self.assertLogs("portfolio_dynamics", level="WARNING") as captured:
            plan = inbox.plan_import(self.source, self.downloads, dt.date(2026, 9, 21),
                                     limits_source=config.PORTFOLIO_DYNAMICS_LIMITS_SOURCE)

        self.assertIsNotNone(plan.limits)
        self.assertIn("14_09_2026", plan.limits.name)
        self.assertTrue(any("более ранний" in line for line in captured.output))

    def test_exact_date_wins_over_an_earlier_one(self):
        self.download("21.09.2026")
        self.download("14.09.2026")
        self.limits_download("14_09_2026")
        self.limits_download("21_09_2026")

        plan = inbox.plan_import(self.source, self.downloads, dt.date(2026, 9, 21),
                                 limits_source=config.PORTFOLIO_DYNAMICS_LIMITS_SOURCE)

        self.assertIn("21_09_2026", plan.limits.name)

    def test_limits_from_the_future_are_never_taken(self):
        """Лимиты, выгруженные позже отчётной даты, к ней отношения не имеют."""
        self.download("21.09.2026")
        self.download("14.09.2026")
        self.limits_download("28_09_2026")

        plan = inbox.plan_import(self.source, self.downloads, dt.date(2026, 9, 21),
                                 limits_source=config.PORTFOLIO_DYNAMICS_LIMITS_SOURCE)

        self.assertIsNone(plan.limits)
        self.assertTrue(plan.complete, "без лимитов отчёт всё равно должен собираться")

    def test_filename_variants_are_recognised(self):
        """Регистр, лишние пробелы и разделитель даты не должны мешать."""
        from test_limits import DEFAULT_ROWS, write_limits_file
        for name in ("состояние лимитов на дату 21.09.2026 - Результат.xlsx",
                     "СОСТОЯНИЕ ЛИМИТОВ НА  ДАТУ  21-09-2026.xlsx"):
            with self.subTest(name=name):
                for old in self.downloads.glob("Состояние*"):
                    old.unlink()
                for old in self.downloads.glob("состояние*"):
                    old.unlink()
                for old in self.downloads.glob("СОСТОЯНИЕ*"):
                    old.unlink()
                write_limits_file(self.downloads / name, DEFAULT_ROWS)
                found = inbox.scan_limits(
                    config.PORTFOLIO_DYNAMICS_LIMITS_SOURCE, self.downloads)
                self.assertEqual([c.business_date for c in found], [dt.date(2026, 9, 21)])

    def test_slices_are_still_chosen_correctly_with_limits_around(self):
        self.download("21.09.2026")
        self.download("14.09.2026")
        self.limits_download("21_09_2026")

        t0, t7 = pd_report._resolve_slice_paths(self.args())

        self.assertIn("21.09.2026", t0.name)
        self.assertIn("14.09.2026", t7.name)

    def test_missing_limits_file_is_not_an_error(self):
        self.download("21.09.2026")
        self.download("14.09.2026")

        plan = inbox.plan_import(self.source, self.downloads, dt.date(2026, 9, 21),
                                 limits_source=config.PORTFOLIO_DYNAMICS_LIMITS_SOURCE)

        self.assertTrue(plan.complete)
        self.assertIsNone(plan.limits)

    def test_limits_file_in_the_folder_does_not_pass_for_a_slice(self):
        """Список дат считает тем же планом, что и запуск, — значит, и файл
        лимитов он обязан исключать так же. Иначе папка «срез + лимиты»
        выглядит укомплектованной, дата предлагается как готовая, а при
        запуске оказывается, что второго среза нет."""
        from test_limits import DEFAULT_ROWS, write_limits_file
        folder = self.data / "2026-09-21"
        folder.mkdir()
        self.in_folder("2026-09-21", "21.09.2026")
        write_limits_file(folder / "Состояние лимитов на дату 21_09_2026 - Результат.xlsx",
                          DEFAULT_ROWS)

        rows = inbox.available_dates(
            self.source, None, limits_source=config.PORTFOLIO_DYNAMICS_LIMITS_SOURCE)

        self.assertEqual(rows, [], "одного среза мало, сколько бы файлов ни лежало рядом")

    def test_the_offered_date_really_builds(self):
        """Список и запуск обязаны видеть одно и то же."""
        from test_limits import DEFAULT_ROWS, write_limits_file
        self.download("21.09.2026")
        self.download("14.09.2026")
        write_limits_file(
            self.downloads / "Состояние лимитов на дату 21_09_2026 - Результат.xlsx",
            DEFAULT_ROWS)

        rows = pd_report._available_dates(self.args())

        self.assertEqual([r.date for r in rows], [dt.date(2026, 9, 21)])
        t0, _t7 = pd_report._resolve_slice_paths(self.args())
        self.assertIn("21.09.2026", t0.name)


class PlanTests(InboxTestCase):
    def test_pair_is_t0_plus_the_newest_earlier_slice(self):
        for period_end in ("04.09.2026", "11.09.2026", "18.09.2026"):
            self.download(period_end)

        plan = inbox.plan_import(self.source, self.downloads, dt.date(2026, 9, 18))

        self.assertTrue(plan.complete)
        self.assertEqual(plan.folder.name, "2026-09-18")
        taken = sorted(src.name for src, _dst in plan.to_import)
        self.assertEqual(len(taken), 2)
        self.assertIn("[18.09.2026]", taken[1])
        self.assertIn("[11.09.2026]", taken[0])  # 04.09 не нужен

    def test_complete_folder_needs_no_import(self):
        self.in_folder("2026-09-18", "11.09.2026")
        self.in_folder("2026-09-18", "18.09.2026")
        self.download("18.09.2026")

        plan = inbox.plan_import(self.source, self.downloads, dt.date(2026, 9, 18))

        self.assertTrue(plan.complete)
        self.assertEqual(plan.to_import, [])
        self.assertEqual(len(plan.existing), 2)

    def test_half_filled_folder_is_topped_up_from_downloads(self):
        self.in_folder("2026-09-18", "18.09.2026")
        self.download("11.09.2026")

        plan = inbox.plan_import(self.source, self.downloads, dt.date(2026, 9, 18))

        self.assertTrue(plan.complete)
        self.assertEqual([src.name for src, _dst in plan.to_import],
                         [export_name("11.09.2026")])
        self.assertEqual(len(plan.existing), 1)

    def test_missing_t0_says_where_it_looked(self):
        self.download("11.09.2026")
        plan = inbox.plan_import(self.source, self.downloads, dt.date(2026, 9, 18))

        self.assertFalse(plan.complete)
        self.assertIn("2026-09-18", plan.problem)
        self.assertIn(str(self.downloads), plan.problem)

    def test_missing_earlier_slice_is_reported_separately(self):
        self.download("18.09.2026")
        plan = inbox.plan_import(self.source, self.downloads, dt.date(2026, 9, 18))

        self.assertFalse(plan.complete)
        self.assertIn("T-7", plan.problem)

    def test_import_disabled_does_not_mention_downloads_as_searched(self):
        self.download("18.09.2026")
        self.download("11.09.2026")

        plan = inbox.plan_import(self.source, None, dt.date(2026, 9, 18))

        self.assertFalse(plan.complete)
        self.assertIn("выключена", plan.problem)


class ApplyTests(InboxTestCase):
    def test_move_takes_files_out_of_downloads(self):
        self.download("18.09.2026")
        self.download("11.09.2026")
        self.download("04.09.2026")  # лишний, должен остаться

        plan = inbox.plan_import(self.source, self.downloads, dt.date(2026, 9, 18))
        inbox.apply_import(plan, move=True)

        self.assertEqual(len(self.names("2026-09-18")), 2)
        self.assertEqual(self.downloads_names(), [export_name("04.09.2026")])

    def test_copy_leaves_the_originals(self):
        self.download("18.09.2026")
        self.download("11.09.2026")

        plan = inbox.plan_import(self.source, self.downloads, dt.date(2026, 9, 18))
        inbox.apply_import(plan, move=False)

        self.assertEqual(len(self.names("2026-09-18")), 2)
        self.assertEqual(len(self.downloads_names()), 2)

    def test_unrelated_files_in_downloads_are_never_touched(self):
        self.download("18.09.2026")
        self.download("11.09.2026")
        (self.downloads / "личное.xlsx").write_bytes(b"mine")
        (self.downloads / "квартальный.pdf").write_bytes(b"mine")

        plan = inbox.plan_import(self.source, self.downloads, dt.date(2026, 9, 18))
        inbox.apply_import(plan, move=True)

        self.assertEqual(self.downloads_names(), ["квартальный.pdf", "личное.xlsx"])

    def test_second_run_is_a_no_op(self):
        self.download("18.09.2026")
        self.download("11.09.2026")
        inbox.resolve_pair(self.source, self.downloads, dt.date(2026, 9, 18))

        # Файлы уже перенесены, загрузки пусты — повтор должен просто их найти.
        files = inbox.resolve_pair(self.source, self.downloads, dt.date(2026, 9, 18))

        self.assertEqual(len(files), 2)
        self.assertEqual(len(self.names("2026-09-18")), 2)

    def test_file_already_in_the_folder_is_not_overwritten(self):
        original = self.in_folder("2026-09-18", "18.09.2026", volume=99)
        stamp = original.read_bytes()
        self.download("18.09.2026", volume=1)  # то же имя, другое содержимое
        self.download("11.09.2026")

        plan = inbox.plan_import(self.source, self.downloads, dt.date(2026, 9, 18))
        inbox.apply_import(plan, move=True)

        self.assertEqual(original.read_bytes(), stamp)


class ArchiveOwnDateTests(InboxTestCase):
    """Каждый срез кладётся и в папку отчётной даты, и в папку своей собственной."""

    def test_t7_lands_in_its_own_date_folder_too(self):
        self.download("18.09.2026")
        self.download("11.09.2026")

        inbox.resolve_pair(self.source, self.downloads, dt.date(2026, 9, 18))

        self.assertEqual(len(self.names("2026-09-18")), 2)
        self.assertEqual(self.names("2026-09-11"), [export_name("11.09.2026")])

    def test_report_for_the_earlier_date_can_be_built_afterwards(self):
        """Ради этого дубль и делается: 11.09 остаётся доступным как T0 своей даты."""
        for period_end in ("04.09.2026", "11.09.2026", "18.09.2026"):
            self.download(period_end)

        pd_report._resolve_slice_paths(self.args(date="2026-09-18"))
        t0, t7 = pd_report._resolve_slice_paths(self.args(date="2026-09-11"))

        self.assertEqual(t0.parent.name, "2026-09-11")
        self.assertIn("[11.09.2026]", t0.name)
        self.assertIn("[04.09.2026]", t7.name)

    def test_t0_is_not_duplicated_anywhere(self):
        """Своя дата у T0 и есть отчётная — второй экземпляр ему не нужен."""
        self.download("18.09.2026")
        self.download("11.09.2026")

        inbox.resolve_pair(self.source, self.downloads, dt.date(2026, 9, 18))

        folders = sorted(p.name for p in self.data.iterdir() if p.is_dir())
        self.assertEqual(folders, ["2026-09-11", "2026-09-18"])

    def test_existing_file_in_the_own_folder_is_not_overwritten(self):
        original = self.in_folder("2026-09-11", "11.09.2026", volume=99)
        stamp = original.read_bytes()
        self.download("18.09.2026")
        self.download("11.09.2026", volume=1)

        inbox.resolve_pair(self.source, self.downloads, dt.date(2026, 9, 18))

        self.assertEqual(original.read_bytes(), stamp)

    def test_can_be_switched_off(self):
        self.download("18.09.2026")
        self.download("11.09.2026")

        inbox.resolve_pair(self.source, self.downloads, dt.date(2026, 9, 18),
                           archive_own_date=False)

        self.assertEqual(self.names("2026-09-11"), [])

    def test_setting_off_is_honoured_by_the_report(self):
        self.download("18.09.2026")
        self.download("11.09.2026")
        settings.set_value("portfolio_dynamics_archive_own_date", "нет")
        config.reload()

        pd_report._resolve_slice_paths(self.args())

        self.assertEqual(self.names("2026-09-11"), [])

    def test_already_complete_folder_still_gets_its_archive_copy(self):
        """Файлы разложили руками — дубль всё равно появится при ближайшем запуске."""
        self.in_folder("2026-09-18", "18.09.2026")
        self.in_folder("2026-09-18", "11.09.2026")

        pd_report._resolve_slice_paths(self.args(date="2026-09-18"))

        self.assertEqual(self.names("2026-09-11"), [export_name("11.09.2026")])


class ColdStartTests(InboxTestCase):
    """Первый запуск: папки исходных файлов ещё нет, всё лежит в загрузках."""

    def test_works_when_the_source_folder_does_not_exist_yet(self):
        import shutil
        shutil.rmtree(self.data)
        self.download("18.09.2026")
        self.download("11.09.2026")

        t0, t7 = pd_report._resolve_slice_paths(self.args())

        self.assertTrue(self.data.is_dir(), "папка исходных файлов должна создаться сама")
        self.assertEqual(t0.parent.name, "2026-09-18")
        self.assertIn("[11.09.2026]", t7.name)

    def test_missing_source_folder_is_not_an_error_for_folder_scan(self):
        """Раньше отсутствие папки роняло запуск раньше, чем доходило до загрузок."""
        import shutil
        from common import file_discovery
        shutil.rmtree(self.data)

        self.assertEqual(file_discovery.find_date_folders(self.source), [])
        self.assertIsNone(file_discovery.latest_folder_with_data(self.source, min_files=2))

    def test_error_message_names_both_places_and_the_way_out(self):
        message = pd_report._nothing_found_message(self.source)

        self.assertIn(str(self.data), message)
        self.assertIn(str(self.downloads), message)
        self.assertIn("--diagnose", message)
        self.assertIn("settings --set", message)

    def test_files_present_but_named_wrong_are_reported(self):
        """Самый частый сюрприз: «файлы же на месте» — а имена не по шаблону."""
        from common.file_discovery import _not_found_message
        (self.data / "Позиция_18_09_2026.xlsx").write_bytes(b"x")
        (self.data / "Позиция_11_09_2026.xlsx").write_bytes(b"x")

        message = _not_found_message(self.source)

        self.assertIn("Файлы в папке есть (2)", message)
        self.assertIn("Позиция_18_09_2026.xlsx", message)

    def test_empty_folder_says_so_explicitly(self):
        from common.file_discovery import _not_found_message
        self.assertIn("Файлов в папке нет вовсе", _not_found_message(self.source))

    def test_error_message_reports_how_many_files_matched(self):
        self.download("18.09.2026")  # один подходящий, пары не хватает
        (self.downloads / "посторонний.xlsx").write_bytes(b"x")

        message = pd_report._nothing_found_message(self.source)
        self.assertIn("подходящих выгрузок: 1", message)


class AvailableDatesTests(InboxTestCase):
    def test_dates_from_folders_and_downloads_are_merged(self):
        self.in_folder("2026-09-11", "04.09.2026")
        self.in_folder("2026-09-11", "11.09.2026")
        self.download("18.09.2026")
        self.download("11.09.2026")

        rows = inbox.available_dates(self.source, self.downloads)
        by_date = {r.date: r for r in rows}

        self.assertEqual(rows[0].date, dt.date(2026, 9, 18))
        # T0 придёт из загрузок, а пару к нему даёт срез, уже разложенный по
        # папкам-датам: он ближе, чем дубль того же файла в загрузках.
        self.assertEqual(by_date[dt.date(2026, 9, 18)].origin, "папки-даты + загрузки")
        self.assertEqual(by_date[dt.date(2026, 9, 11)].origin, "папка")

    def test_dates_without_a_pair_are_not_offered(self):
        self.download("18.09.2026")  # без более раннего среза пары нет
        self.assertEqual(inbox.available_dates(self.source, self.downloads), [])

    def test_a_date_whose_pair_lies_in_another_folder_is_offered(self):
        """Ровно тот случай, из-за которого свежая выгрузка «пропадала»: T-7
        уже разложен по своей папке-дате, в загрузках его больше нет."""
        self.in_folder("2026-09-11", "11.09.2026")
        self.download("18.09.2026")

        rows = inbox.available_dates(self.source, self.downloads)

        self.assertEqual([r.date for r in rows],
                         [dt.date(2026, 9, 18), dt.date(2026, 9, 11)][:len(rows)])
        self.assertIn(dt.date(2026, 9, 18), [r.date for r in rows])


class SlicesAlreadyFiledTests(InboxTestCase):
    """Пара собирается и тогда, когда T-7 уже разложен по своей папке-дате.

    Приёмка сама раскладывает каждый срез по папке его даты и вычищает
    загрузки. Если после этого пару искать только в папке отчётной даты и в
    загрузках, то через неделю работы свежая выгрузка перестаёт собираться:
    T-7 лежит в СВОЕЙ папке, и отчёт объявляет, что файла нет, — стоя ровно в
    той папке, куда сам его и положил.
    """

    def test_t7_is_taken_from_its_own_date_folder(self):
        self.in_folder("2026-09-11", "11.09.2026")
        self.download("18.09.2026")

        t0, t7 = pd_report._resolve_slice_paths(self.args())

        self.assertEqual(etl.read_business_date(t0), dt.date(2026, 9, 18))
        self.assertEqual(etl.read_business_date(t7), dt.date(2026, 9, 11))

    def test_the_source_folder_keeps_its_copy(self):
        """Из чужой папки файл копируется: там он единственный экземпляр
        отчёта на свою дату, и «перенести» значило бы сломать прошлый."""
        self.in_folder("2026-09-11", "11.09.2026")
        self.download("18.09.2026")

        pd_report._resolve_slice_paths(self.args())

        self.assertEqual(len(self.names("2026-09-11")), 1, "срез остался у себя")
        self.assertEqual(len(self.names("2026-09-18")), 2, "и появился в папке отчёта")

    def test_it_is_written_down_in_the_log(self):
        self.in_folder("2026-09-11", "11.09.2026")
        self.download("18.09.2026")

        with self.assertLogs("portfolio_dynamics", level="INFO") as captured:
            pd_report._resolve_slice_paths(self.args())

        self.assertTrue(any("2026-09-11" in line and "взят из папки" in line
                            for line in captured.output))

    def test_a_second_week_in_a_row(self):
        """Сценарий целиком: неделю назад собрали отчёт, теперь скачали одну
        свежую выгрузку — и она обязана собраться, а не пересобрать прошлую."""
        self.download("04.09.2026")
        self.download("11.09.2026")
        pd_report._resolve_slice_paths(self.args())      # неделя 1
        self.assertEqual(self.downloads_names(), [], "загрузки вычищены")

        self.download("18.09.2026")                      # неделя 2
        t0, t7 = pd_report._resolve_slice_paths(self.args())

        self.assertEqual(etl.read_business_date(t0), dt.date(2026, 9, 18))
        self.assertEqual(etl.read_business_date(t7), dt.date(2026, 9, 11))

    def test_the_older_pair_can_still_be_rebuilt(self):
        """Пересборка прошлой даты не должна ломаться от того, что появилась новая."""
        self.download("04.09.2026")
        self.download("11.09.2026")
        pd_report._resolve_slice_paths(self.args())
        self.download("18.09.2026")
        pd_report._resolve_slice_paths(self.args())

        t0, t7 = pd_report._resolve_slice_paths(self.args(date="2026-09-11"))

        self.assertEqual(etl.read_business_date(t0), dt.date(2026, 9, 11))
        self.assertEqual(etl.read_business_date(t7), dt.date(2026, 9, 4))

    def test_the_nearest_earlier_slice_wins_across_folders(self):
        """T-7 — САМЫЙ СВЕЖИЙ из более ранних, где бы он ни лежал."""
        self.in_folder("2026-08-28", "28.08.2026")
        self.in_folder("2026-09-11", "11.09.2026")
        self.download("18.09.2026")

        _t0, t7 = pd_report._resolve_slice_paths(self.args())

        self.assertEqual(etl.read_business_date(t7), dt.date(2026, 9, 11))

    def test_flat_layout_is_unaffected(self):
        """Без папок-дат искать по соседним папкам нечего и незачем."""
        settings.set_value("portfolio_dynamics_use_date_folders", False)
        config.reload()
        source = config.PORTFOLIO_DYNAMICS_T0_SOURCE
        self.download("11.09.2026")
        self.download("18.09.2026")

        plan = inbox.plan_import(source, self.downloads, dt.date(2026, 9, 18))

        self.assertTrue(plan.complete)
        self.assertEqual(plan.reuse, [])


class ReportIntegrationTests(InboxTestCase):
    def test_no_arguments_builds_the_latest_date_from_downloads(self):
        for period_end in ("04.09.2026", "11.09.2026", "18.09.2026"):
            self.download(period_end)

        t0, t7 = pd_report._resolve_slice_paths(self.args())

        self.assertEqual(t0.parent.name, "2026-09-18")
        self.assertIn("[18.09.2026]", t0.name)
        self.assertIn("[11.09.2026]", t7.name)

    def test_explicit_date_builds_that_date(self):
        for period_end in ("04.09.2026", "11.09.2026", "18.09.2026"):
            self.download(period_end)

        t0, t7 = pd_report._resolve_slice_paths(self.args(date="2026-09-11"))

        self.assertEqual(t0.parent.name, "2026-09-11")
        self.assertIn("[11.09.2026]", t0.name)
        self.assertIn("[04.09.2026]", t7.name)

    def test_no_import_flag_keeps_downloads_untouched(self):
        self.download("18.09.2026")
        self.download("11.09.2026")

        with self.assertRaises(etl.PortfolioDynamicsError):
            pd_report._resolve_slice_paths(self.args(no_import=True))
        self.assertEqual(len(self.downloads_names()), 2)

    def test_setting_off_disables_the_import(self):
        self.download("18.09.2026")
        self.download("11.09.2026")
        settings.set_value("portfolio_dynamics_import_from_downloads", "нет")
        config.reload()

        with self.assertRaises(etl.PortfolioDynamicsError):
            pd_report._resolve_slice_paths(self.args())
        self.assertEqual(len(self.downloads_names()), 2)

    def test_copy_mode_is_honoured_by_the_report(self):
        self.download("18.09.2026")
        self.download("11.09.2026")
        settings.set_value("portfolio_dynamics_move_from_downloads", "нет")
        config.reload()

        pd_report._resolve_slice_paths(self.args())

        self.assertEqual(len(self.downloads_names()), 2)
        self.assertEqual(len(self.names("2026-09-18")), 2)

    def test_existing_folder_wins_over_downloads_for_the_same_date(self):
        """Данные на месте — за ними в загрузки отчёт не ходит."""
        self.in_folder("2026-09-18", "11.09.2026")
        self.in_folder("2026-09-18", "18.09.2026")
        self.download("25.09.2026")
        self.download("18.09.2026")

        t0, _t7 = pd_report._resolve_slice_paths(self.args(date="2026-09-18"))

        self.assertEqual(t0.parent.name, "2026-09-18")
        self.assertEqual(len(self.downloads_names()), 2)

    def test_fresh_pair_in_downloads_beats_an_older_ready_folder(self):
        """«По умолчанию последняя» — по папкам И загрузкам сразу, а не только по папкам."""
        self.in_folder("2026-09-11", "04.09.2026")
        self.in_folder("2026-09-11", "11.09.2026")
        self.download("18.09.2026")
        self.download("11.09.2026")

        t0, _t7 = pd_report._resolve_slice_paths(self.args())

        self.assertEqual(t0.parent.name, "2026-09-18")
        self.assertIn("[18.09.2026]", t0.name)


if __name__ == "__main__":
    unittest.main()
