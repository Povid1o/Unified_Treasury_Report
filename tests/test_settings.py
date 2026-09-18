"""Тесты слоя настроек (common/settings.py) и его склейки с config.py.

Проверяется то, из-за чего настройки перестают быть полезными: значения по
умолчанию, которые разъехались с прежним config.py; производные от корня пути,
которые перестали следовать за корнем; принятое некорректное значение, из-за
которого отчёт потом «просто не находит файлы»; и несохранение/неприменение
изменений в текущем сеансе.
"""
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BASE_DIR))

import config  # noqa: E402
from common import settings  # noqa: E402


class SettingsTestCase(unittest.TestCase):
    """Каждый тест работает со своим файлом настроек во временной папке."""

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.settings_file = Path(self._tmp.name) / "settings.json"
        self._saved_env = os.environ.get(settings.SETTINGS_FILE_ENV)
        os.environ[settings.SETTINGS_FILE_ENV] = str(self.settings_file)
        config.reload()

    def tearDown(self):
        if self._saved_env is None:
            os.environ.pop(settings.SETTINGS_FILE_ENV, None)
        else:
            os.environ[settings.SETTINGS_FILE_ENV] = self._saved_env
        config.reload()  # возвращаем процессу настройки разработчика
        self._tmp.cleanup()

    def stored(self) -> dict:
        return json.loads(self.settings_file.read_text(encoding="utf-8"))


class DefaultsTests(SettingsTestCase):
    def test_defaults_match_the_paths_config_used_before(self):
        """Значения по умолчанию — ровно те, что были захардкожены в config.py."""
        root = Path(r"O:\Exchequer\Sotrudniki\Башлыков\Навигатор\Jupiter")
        self.assertEqual(config.JUPITER_ROOT, root)
        self.assertEqual(config.OVP_SOURCE.directory, root / "data" / "OVP")
        self.assertEqual(config.CHPD_SOURCE.filename_regex, r"^ЧПД (\d{4} \d{2} \d{2})\.xlsx$")
        self.assertEqual(config.CHPD_SOURCE.date_format, "%Y %m %d")
        self.assertEqual(config.TRANSFERT_LONG_SOURCE.directory, root / "data" / "Transferta" / "Long")
        self.assertEqual(config.PORTFOLIO_DYNAMICS_VALUE_SCALE, 1_000_000)
        self.assertEqual(config.PORTFOLIO_DYNAMICS_TOLERANCE, 0.005)
        self.assertEqual(config.OFZ_LOOKBACK_DAYS, 90)

    def test_no_settings_file_means_no_file_is_created(self):
        """Пока пользователь ничего не менял, settings.json не появляется."""
        settings.values()
        self.assertFalse(self.settings_file.exists())

    def test_every_setting_has_a_group_and_resolves(self):
        known_groups = {g.key for g in settings.GROUPS}
        for setting in settings.SETTINGS:
            self.assertIn(setting.group, known_groups, setting.key)
            self.assertIsNotNone(settings.get(setting.key), setting.key)


class OverrideTests(SettingsTestCase):
    def test_derived_paths_follow_the_root(self):
        """Сменил корень — переехали все папки, которые не заданы явно."""
        settings.set_value("jupiter_root", "/Volumes/Jupiter")
        config.reload()

        self.assertEqual(config.OVP_SOURCE.directory, Path("/Volumes/Jupiter/data/OVP"))
        self.assertEqual(config.CHPD_OUTPUT_DIR, Path("/Volumes/Jupiter/output/CHPD"))
        self.assertEqual(
            config.PORTFOLIO_DYNAMICS_T0_SOURCE.directory,
            Path("/Volumes/Jupiter/data/PortfolioDynamics"),
        )

    def test_explicit_path_stops_following_the_root(self):
        settings.set_value("ovp_dir", "/data/ovp-special")
        settings.set_value("jupiter_root", "/Volumes/Jupiter")
        config.reload()

        self.assertEqual(config.OVP_SOURCE.directory, Path("/data/ovp-special"))
        self.assertEqual(config.CHPD_SOURCE.directory, Path("/Volumes/Jupiter/data/CHPD"))

    def test_only_changed_values_are_stored(self):
        settings.set_value("nim_dir", "/data/nim")
        self.assertEqual(list(self.stored()), ["nim_dir"])

    def test_reset_returns_the_default_and_drops_the_key(self):
        settings.set_value("nim_dir", "/data/nim")
        self.assertTrue(settings.is_overridden("nim_dir"))

        settings.reset("nim_dir")
        config.reload()

        self.assertFalse(settings.is_overridden("nim_dir"))
        self.assertEqual(self.stored(), {})
        self.assertEqual(config.NIM_SOURCE.directory, config.JUPITER_ROOT / "data" / "NIM")

    def test_reset_all_clears_every_override(self):
        settings.set_value("nim_dir", "/data/nim")
        settings.set_value("ofz_lookback_days", 30)
        settings.reset_all()
        config.reload()

        self.assertEqual(settings.overridden_keys(), [])
        self.assertEqual(config.OFZ_LOOKBACK_DAYS, 90)

    def test_changes_survive_a_reload_from_disk(self):
        settings.set_value("portfolio_dynamics_value_scale", 1000)
        settings.reload()
        config.reload()

        self.assertEqual(config.PORTFOLIO_DYNAMICS_VALUE_SCALE, 1000.0)
        self.assertEqual(self.stored()["portfolio_dynamics_value_scale"], 1000.0)

    def test_unknown_keys_in_the_file_are_ignored(self):
        """Файл, записанный более новой версией проекта, не должен ломать старую."""
        self.settings_file.write_text(
            json.dumps({"nim_dir": "/data/nim", "setting_from_the_future": 42}),
            encoding="utf-8",
        )
        settings.reload()
        config.reload()

        self.assertEqual(config.NIM_SOURCE.directory, Path("/data/nim"))
        self.assertEqual(settings.overridden_keys(), ["nim_dir"])

    def test_broken_file_says_what_to_do(self):
        self.settings_file.write_text("{не json", encoding="utf-8")
        settings.reload()
        with self.assertRaisesRegex(settings.SettingsError, "удалите"):
            settings.get("nim_dir")


class ValidationTests(SettingsTestCase):
    def test_regex_without_exactly_one_group_is_rejected(self):
        """Регулярка без группы находит файл, но не дату — отчёт молча теряет файлы."""
        for bad in (r"^ЧПД .*\.xlsx$", r"^(ЧПД) (\d{4})\.xlsx$"):
            with self.subTest(bad=bad):
                with self.assertRaisesRegex(settings.SettingsError, "одна группа"):
                    settings.set_value("chpd_regex", bad)

    def test_invalid_regex_is_rejected(self):
        with self.assertRaisesRegex(settings.SettingsError, "регулярное выражение"):
            settings.set_value("chpd_regex", r"^ЧПД (\d{4}\.xlsx$")

    def test_date_format_must_round_trip(self):
        with self.assertRaisesRegex(settings.SettingsError, "формат даты"):
            settings.set_value("chpd_date_format", "просто текст")
        self.assertEqual(settings.set_value("chpd_date_format", "%d-%m-%Y"), "%d-%m-%Y")

    def test_numbers_must_be_positive(self):
        with self.assertRaisesRegex(settings.SettingsError, "больше нуля"):
            settings.set_value("ofz_lookback_days", 0)
        with self.assertRaisesRegex(settings.SettingsError, "больше нуля"):
            settings.set_value("portfolio_dynamics_value_scale", -1)
        with self.assertRaisesRegex(settings.SettingsError, "Ожидается"):
            settings.set_value("ofz_lookback_days", "девяносто")

    def test_comma_is_accepted_as_a_decimal_separator(self):
        self.assertEqual(settings.set_value("portfolio_dynamics_tolerance", "0,01"), 0.01)

    def test_quotes_around_a_pasted_path_are_stripped(self):
        """Путь из проводника вставляют вместе с кавычками — это не должно мешать."""
        self.assertEqual(settings.set_value("nim_dir", '"/data/nim"'), Path("/data/nim"))

    def test_unknown_key_is_rejected(self):
        with self.assertRaisesRegex(settings.SettingsError, "Неизвестная настройка"):
            settings.set_value("no_such_setting", "x")
        with self.assertRaisesRegex(settings.SettingsError, "Неизвестная настройка"):
            settings.get("no_such_setting")


class CheckPathsTests(SettingsTestCase):
    def test_missing_and_present_directories_are_distinguished(self):
        with tempfile.TemporaryDirectory() as existing:
            (Path(existing) / "отчёт.xlsx").touch()
            settings.set_value("ovp_dir", existing)
            settings.set_value("nim_dir", "/no/such/directory")

            by_key = {s.key: s for s in settings.check_paths()}

        self.assertTrue(by_key["ovp_dir"].exists)
        self.assertEqual(by_key["ovp_dir"].files, 1)
        self.assertFalse(by_key["nim_dir"].exists)
        self.assertIn("недоступна", by_key["nim_dir"].detail)

    def test_output_folder_is_not_reported_as_a_problem(self):
        """Выходной папки на чистой установке нет — это норма, а не поломка."""
        settings.set_value("jupiter_root", "/no/such/root")
        by_key = {s.key: s for s in settings.check_paths()}

        self.assertFalse(by_key["ovp_output_dir"].exists)
        self.assertFalse(by_key["ovp_output_dir"].is_problem)
        self.assertIn("создана", by_key["ovp_output_dir"].detail)
        # А вот папки-источника не хватает по-настоящему.
        self.assertTrue(by_key["ovp_dir"].is_problem)

    def test_source_folder_without_matching_files_is_visible(self):
        with tempfile.TemporaryDirectory() as empty:
            settings.set_value("ovp_dir", empty)
            status = {s.key: s for s in settings.check_paths()}["ovp_dir"]

        self.assertTrue(status.exists)
        self.assertEqual(status.files, 0)


if __name__ == "__main__":
    unittest.main()
