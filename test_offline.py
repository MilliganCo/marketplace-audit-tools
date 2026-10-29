import hashlib
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd

import ozon_cpc_campaign_audit as ozon
import wb_concurents as wb


class OfflineTests(unittest.TestCase):
    def test_token_is_explicit_environment_only(self):
        with patch.dict(os.environ, {}, clear=True), patch("builtins.open", side_effect=AssertionError("No credential file reads")):
            with self.assertRaisesRegex(ValueError, "WB_TOKEN"):
                wb.load_wb_token_from_keys()
        for value, expected in [("synthetic", "Bearer synthetic"), (" Bearer synthetic ", "Bearer synthetic")]:
            with self.subTest(value=value), patch.dict(os.environ, {"WB_TOKEN": value}, clear=True):
                self.assertEqual(wb.load_wb_token_from_keys(), expected)

    def test_rates_support_both_export_formats(self):
        self.assertEqual(ozon.normalize_rate(pd.Series([10, 20, 30])).tolist(), [.1, .2, .3])
        self.assertEqual(ozon.normalize_rate(pd.Series([.1, .2, .3])).tolist(), [.1, .2, .3])
        self.assertEqual(ozon.to_float("1 234,50 ₽"), 1234.5)
        self.assertEqual(ozon.safe_div(5, 0), 0)

    def test_synthetic_workbook_roundtrip_preserves_input(self):
        row = {name: 0 for name in ozon.REQUIRED_COLUMNS}
        row.update({"ID кампании": 101, "Название кампании": "Synthetic example", "Статус": "Активна", "Стратегия": "CPC",
                    "Расход, ₽": 100, "Показы": 1000, "Клики": 100, "CTR, %": 10,
                    "Средняя стоимость клика, ₽": 1, "Заказы, шт.": 5, "Продажи, ₽": 1000, "ДРР, %": 10, "В корзину": 20})
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "synthetic.xlsx"
            pd.DataFrame([row]).to_excel(source, index=False, startrow=2)
            digest = hashlib.sha256(source.read_bytes()).hexdigest()
            with patch("requests.sessions.Session.request", side_effect=AssertionError("Offline test")):
                spreadsheet, text = ozon.analyze(source, .3, root / "analysis")
            result = pd.read_excel(spreadsheet, sheet_name="Активные кампании")
            self.assertEqual(result["ID кампании"].tolist(), [101])
            self.assertEqual(result["Допустимый CPC, ₽"].tolist(), [3])
            self.assertIn("Synthetic example", text.read_text("utf-8"))
            self.assertEqual(hashlib.sha256(source.read_bytes()).hexdigest(), digest)

    def test_empty_active_report_fails_before_outputs(self):
        with patch.object(ozon, "load_report", return_value=pd.DataFrame({"Статус": ["Остановлена"]})):
            with self.assertRaisesRegex(ValueError, "активных"):
                ozon.analyze(Path("unused.xlsx"), .3, Path("unused"))


if __name__ == "__main__":
    unittest.main()
