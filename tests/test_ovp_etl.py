import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd

from reports.ovp import etl
from reports.ovp import report as ovp_report


def _hierarchy_labels():
    labels = []
    for category, subcategories in etl.HIERARCHY.items():
        labels.append(category)
        labels.extend(subcategories)
    return labels


def _consolidated_frame():
    labels = _hierarchy_labels()
    frame = pd.DataFrame([[None] * 12 for _ in range(4 + len(labels))])
    frame.iat[0, 1] = "Прогноз экономической ОВП ПСБ на 02.09.2026 г."

    blocks = (("CNY", 1, 2), ("USD", 4, 5), ("EURO", 7, 8))
    for currency, balance_col, reserve_col in blocks:
        frame.iat[2, balance_col] = currency
        frame.iat[3, balance_col] = "Валютный\nБаланс"
        frame.iat[3, reserve_col] = "Резервы МСФО"
        frame.iat[3, reserve_col + 1] = "ИТОГО"

    frame.iat[3, 10] = "Итого EUR в USD"
    frame.iat[2, 11] = "ИТОГО в USD и EURO"

    for label_idx, label in enumerate(labels):
        row_idx = 4 + label_idx
        if label == "ОВП БФР":
            label = "ОВП БФР за 01.09.2026"
        frame.iat[row_idx, 0] = label
        for block_idx, (_currency, balance_col, reserve_col) in enumerate(blocks, start=1):
            frame.iat[row_idx, balance_col] = block_idx * 1000 + label_idx + 0.9
            frame.iat[row_idx, reserve_col] = -(block_idx * 100 + label_idx + 0.9)
            frame.iat[row_idx, reserve_col + 1] = 999999
        frame.iat[row_idx, 10] = 888888
        frame.iat[row_idx, 11] = 777777

    return frame


def _legacy_frame(currency_offset):
    labels = _hierarchy_labels()
    frame = pd.DataFrame([[None] * 5 for _ in range(2 + len(labels))])
    frame.iat[0, 1] = "01.09.2026"
    frame.iat[0, 3] = "02.09.2026"
    frame.iat[1, 1] = "Валютный Баланс"
    frame.iat[1, 2] = "Резервы МСФО"
    frame.iat[1, 3] = "Валютный Баланс"
    frame.iat[1, 4] = "Резервы МСФО"

    for label_idx, label in enumerate(labels):
        row_idx = 2 + label_idx
        frame.iat[row_idx, 0] = label
        frame.iat[row_idx, 1] = currency_offset + 100 + label_idx
        frame.iat[row_idx, 2] = currency_offset + 200 + label_idx
        frame.iat[row_idx, 3] = currency_offset + 300 + label_idx
        frame.iat[row_idx, 4] = currency_offset + 400 + label_idx
    return frame


class OvpEtlTests(unittest.TestCase):
    def test_single_sheet_is_selected_automatically(self):
        self.assertEqual(etl._resolve_sheet_selection(["ОВП"], None), "ОВП")

    def test_multiple_sheets_require_explicit_selection(self):
        with self.assertRaisesRegex(etl.OvpDataError, "несколько листов"):
            etl._resolve_sheet_selection(["02.09", "03.09"], None)

    def test_explicit_sheet_selection_is_validated(self):
        self.assertEqual(
            etl._resolve_sheet_selection(["02.09", "03.09"], "03.09"),
            "03.09",
        )
        with self.assertRaisesRegex(etl.OvpDataError, "не найден"):
            etl._resolve_sheet_selection(["02.09", "03.09"], "04.09")

    @patch.object(ovp_report.ui.console, "print")
    @patch.object(ovp_report.ui, "ask", return_value="2")
    def test_interactive_mode_offers_sheet_choice(self, _ask, _print):
        selected = ovp_report._select_sheet(
            Path("report.xlsx"),
            ["02.09", "03.09"],
        )

        self.assertEqual(selected, "03.09")

    def test_new_consolidated_layout_keeps_old_output_contract(self):
        result = etl._convert_sheet_matrices([("СВОД", _consolidated_frame())])

        self.assertEqual(list(result.columns), etl.OUT_COLUMNS)
        self.assertEqual(len(result), 90)
        self.assertEqual(result["currency"].drop_duplicates().tolist(), ["CNY", "USD", "EURO"])
        self.assertEqual(result["report_date"].unique().tolist(), ["2026-09-02"])
        self.assertEqual(result.groupby("currency")["ord"].apply(list)["CNY"], list(range(1, 31)))

        first_cny = result.iloc[0]
        self.assertEqual(first_cny["curr_balance"], 1000)
        self.assertEqual(first_cny["reserve_msfo"], -100)

        first_euro = result[result["currency"] == "EURO"].iloc[0]
        self.assertEqual(first_euro["curr_balance"], 3000)
        self.assertEqual(first_euro["reserve_msfo"], -300)
        self.assertNotIn(999999, result["curr_balance"].tolist())

        ovp_bfr = result[result["subcategory"] == "ОВП БФР"]
        self.assertEqual(len(ovp_bfr), 3)
        self.assertEqual(ovp_bfr["report_date"].unique().tolist(), ["2026-09-02"])

    def test_new_layout_accepts_eur_alias_filter(self):
        result = etl._convert_sheet_matrices(
            [("СВОД", _consolidated_frame())],
            currencies=["EUR"],
        )

        self.assertEqual(len(result), 30)
        self.assertEqual(result["currency"].unique().tolist(), ["EURO"])
        self.assertEqual(result["id"].tolist(), list(range(1, 31)))

    def test_legacy_currency_sheets_still_use_latest_date(self):
        sheets = [
            ("CNY", _legacy_frame(0)),
            ("USD", _legacy_frame(1000)),
            ("EURO", _legacy_frame(2000)),
            ("СВОД", pd.DataFrame([["служебный лист"]])),
        ]
        result = etl._convert_sheet_matrices(sheets)

        self.assertEqual(len(result), 90)
        self.assertEqual(result["report_date"].unique().tolist(), ["2026-09-02"])
        self.assertEqual(result.iloc[0]["curr_balance"], 300)
        self.assertEqual(result.iloc[30]["curr_balance"], 1300)
        self.assertEqual(result.iloc[60]["curr_balance"], 2300)

    def test_legacy_full_history_is_preserved(self):
        result = etl._convert_sheet_matrices(
            [("CNY", _legacy_frame(0))],
            full_history_currencies=["CNY"],
        )

        self.assertEqual(len(result), 60)
        self.assertEqual(
            result["report_date"].drop_duplicates().tolist(),
            ["2026-09-01", "2026-09-02"],
        )
        self.assertEqual(result.iloc[0]["curr_balance"], 100)
        self.assertEqual(result.iloc[30]["curr_balance"], 300)


if __name__ == "__main__":
    unittest.main()
