"""Тесты справочника «портфель -> тип» (portfolio_types.json).

Файл — главный источник разметки: всё, что в нём указано, перебивает и
угадывание по коду, и настройки. Типов в отчёте ровно столько, сколько ключей
в файле (по умолчанию TSS, AFS, HTM и подтип HTM_KUAP), — угаданное «OFZ»
типом не становится.
"""
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
from reports.portfolio_dynamics import etl, type_map  # noqa: E402
from test_portfolio_dynamics_etl import PortfolioDynamicsTestCase  # noqa: E402


class TypeMapTestCase(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self._saved_env = os.environ.get(settings.SETTINGS_FILE_ENV)
        settings_file = self.tmp / "settings.json"
        settings_file.write_text("{}", encoding="utf-8")
        os.environ[settings.SETTINGS_FILE_ENV] = str(settings_file)
        config.reload()
        type_map.forget_cache()
        self.path = type_map.types_file()

    def tearDown(self):
        if self._saved_env is None:
            os.environ.pop(settings.SETTINGS_FILE_ENV, None)
        else:
            os.environ[settings.SETTINGS_FILE_ENV] = self._saved_env
        config.reload()
        type_map.forget_cache()
        self._tmp.cleanup()

    def write(self, data) -> None:
        self.path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        type_map.forget_cache()


class LoadTests(TypeMapTestCase):
    def test_file_lives_next_to_settings(self):
        self.assertEqual(self.path.parent, self.tmp)

    def test_without_file_there_are_four_types(self):
        self.assertEqual(type_map.registry(), ["TSS", "AFS", "HTM", "HTM_KUAP"])
        self.assertIsNone(type_map.explicit_type("AFS_TR_RUR"))

    def test_exact_code_and_mask(self):
        self.write({"AFS": ["AFS_*"], "TSS": ["afs_special"], "HTM": [], "HTM_KUAP": []})

        self.assertEqual(type_map.explicit_type("AFS_TR_RUR"), "AFS")
        self.assertEqual(type_map.explicit_type("AFS_SPECIAL"), "TSS")  # точный код важнее
        self.assertIsNone(type_map.explicit_type("OFZ_PD"))

    def test_type_spellings_are_normalised(self):
        self.write({"TTS": ["A"], "htm-kuap": ["B"], "AFS": [], "HTM": []})

        self.assertEqual(type_map.explicit_type("A"), "TSS")
        self.assertEqual(type_map.explicit_type("B"), "HTM_KUAP")
        self.assertIn("HTM_KUAP", type_map.registry())

    def test_code_in_two_lists_takes_the_subtype(self):
        self.write({"HTM": ["KUAP_1"], "HTM_KUAP": ["KUAP_1"], "AFS": [], "TSS": []})

        with self.assertLogs(etl.logger, level="WARNING"):
            self.assertEqual(type_map.explicit_type("KUAP_1"), "HTM_KUAP")

    def test_broken_json_is_a_loud_error(self):
        self.path.write_text('{"AFS": ["A",]}', encoding="utf-8")
        type_map.forget_cache()

        with self.assertRaises(etl.PortfolioDynamicsError) as ctx:
            type_map.explicit_type("A")
        self.assertIn("строка 1", str(ctx.exception))

    def test_service_keys_are_not_types(self):
        self.write({"_как_заполнять": ["..."], "AFS": [], "_не_размечены": {"X": "AFS?"}})

        self.assertEqual(type_map.registry(), ["AFS"])


class RecordUnmappedTests(TypeMapTestCase):
    def test_file_is_created_with_template(self):
        added = type_map.record_unmapped({"NEW_ONE": "AFS", "ODD": None})

        data = json.loads(self.path.read_text(encoding="utf-8"))
        self.assertEqual(sorted(added), ["NEW_ONE", "ODD"])
        for portfolio_type in ("TSS", "AFS", "HTM", "HTM_KUAP"):
            self.assertEqual(data[portfolio_type], [])
        self.assertEqual(data[type_map.UNMAPPED_KEY], {"NEW_ONE": "AFS?", "ODD": "?"})

    def test_mapped_codes_leave_the_unmapped_section(self):
        type_map.record_unmapped({"A": "AFS", "B": "HTM"})
        data = json.loads(self.path.read_text(encoding="utf-8"))
        data["AFS"].append("A")
        self.write(data)

        type_map.record_unmapped({"B": "HTM"})

        data = json.loads(self.path.read_text(encoding="utf-8"))
        self.assertEqual(data[type_map.UNMAPPED_KEY], {"B": "HTM?"})
        self.assertEqual(data["AFS"], ["A"])


class EtlTypingTests(PortfolioDynamicsTestCase):
    """Разметка из файла доходит до dim_portfolio и объёмов по типам."""

    def setUp(self):
        super().setUp()
        type_map.forget_cache()
        self.path = type_map.types_file()

    def tearDown(self):
        type_map.forget_cache()
        super().tearDown()

    def write(self, data):
        self.path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        type_map.forget_cache()

    def types_of(self, data):
        dim = data.dim_portfolio
        return dict(zip(dim["portfolio_code"], dim["portfolio_type"]))

    def test_file_beats_guessing(self):
        # По коду HTM_ALCO угадывается HTM, но в файле он размечен как AFS.
        self.write({"TSS": [], "AFS": ["HTM_ALCO"], "HTM": [], "HTM_KUAP": []})

        data = self.build(bootstrap=True)

        self.assertEqual(self.types_of(data)["HTM_ALCO"], "AFS")

    def test_file_beats_settings_override(self):
        settings.set_value("portfolio_dynamics_portfolio_types", "HTM_ALCO=TSS")
        config.reload()
        self.write({"TSS": [], "AFS": [], "HTM": ["HTM_ALCO"], "HTM_KUAP": []})

        data = self.build(bootstrap=True)

        self.assertEqual(self.types_of(data)["HTM_ALCO"], "HTM")

    def test_file_retypes_portfolios_of_previous_release(self):
        previous = self.bootstrap_release()
        self.write({"TSS": ["HTM_ALCO"], "AFS": [], "HTM": [], "HTM_KUAP": []})

        data = self.build(previous=previous)

        self.assertEqual(self.types_of(data)["HTM_ALCO"], "TSS")

    def test_kuap_counts_inside_htm(self):
        """HTM_KUAP — подтип: объём HTM включает КУАП, КУАП виден отдельной строкой."""
        self.write({"TSS": [], "AFS": [], "HTM": [], "HTM_KUAP": ["HTM_ALCO"]})

        data = self.build(bootstrap=True)
        history = data.fact_type_daily
        today = history[history["business_date"] == data.business_date]
        volumes = dict(zip(today["portfolio_type"], today["volume_amount"]))

        self.assertGreater(volumes["HTM_KUAP"], 0)
        self.assertAlmostEqual(volumes["HTM"], volumes["HTM_KUAP"])
        dim = data.dim_portfolio.set_index("portfolio_code")
        self.assertFalse(dim.loc["HTM_ALCO", "include_in_total"])  # ИТОГО не задваивается

    def test_unmapped_portfolios_are_written_to_the_file(self):
        self.build(bootstrap=True)

        data = json.loads(self.path.read_text(encoding="utf-8"))
        self.assertEqual(data[type_map.UNMAPPED_KEY], {"AFS_TR_RUR": "AFS?", "HTM_ALCO": "HTM?"})

    def test_junk_types_from_previous_release_are_dropped(self):
        previous = self.bootstrap_release()
        from test_portfolio_dynamics_etl import _append_history
        import datetime as dt
        _append_history(previous, [(dt.date(2026, 1, 5), "OFZ", 5.0)])

        with self.assertLogs(etl.logger, level="WARNING"):
            data = self.build(previous=previous)

        self.assertNotIn("OFZ", set(data.fact_type_daily["portfolio_type"]))
        self.assertLessEqual(set(data.fact_limit["portfolio_type"]),
                             {"TSS", "AFS", "HTM", "HTM_KUAP"})


if __name__ == "__main__":
    unittest.main()
