import os
import unittest
from unittest.mock import patch

from invoice_system.config import Settings, parse_allowed_user_ids, parse_bool, parse_choice, parse_csv, parse_float, parse_int


class ConfigTests(unittest.TestCase):
    def test_parse_allowed_user_ids_accepts_common_separators(self):
        self.assertEqual(parse_allowed_user_ids("123, 456;789 101"), {123, 456, 789, 101})

    def test_parse_allowed_user_ids_ignores_non_numeric_notes(self):
        self.assertEqual(parse_allowed_user_ids("Marco: 123456789, @helper 987"), {123456789, 987})

    def test_parse_allowed_user_ids_empty_allows_all(self):
        self.assertEqual(parse_allowed_user_ids(""), set())
        self.assertEqual(parse_allowed_user_ids(None), set())

    def test_parse_csv_defaults_when_empty(self):
        self.assertEqual(parse_csv("", ("es", "en")), ("es", "en"))
        self.assertEqual(parse_csv("es;en", ("en",)), ("es", "en"))

    def test_parse_bool_accepts_common_values(self):
        self.assertTrue(parse_bool("true"))
        self.assertTrue(parse_bool(" yes "))
        self.assertTrue(parse_bool("ON"))
        self.assertFalse(parse_bool("false", True))
        self.assertFalse(parse_bool(" no ", True))
        self.assertFalse(parse_bool("off", True))

    def test_parse_bool_defaults_for_empty_or_unknown_values(self):
        self.assertTrue(parse_bool("", True))
        self.assertFalse(parse_bool(None, False))
        self.assertTrue(parse_bool("maybe", True))

    def test_parse_float_accepts_numbers_and_whitespace(self):
        self.assertEqual(parse_float(" 0.74 ", 0.62), 0.74)

    def test_parse_float_defaults_when_empty_or_invalid(self):
        self.assertEqual(parse_float("", 0.62), 0.62)
        self.assertEqual(parse_float(None, 0.50), 0.50)
        self.assertEqual(parse_float("not-a-number", 0.50), 0.50)

    def test_parse_int_accepts_numbers_and_defaults_when_invalid(self):
        self.assertEqual(parse_int(" 4 ", 9), 4)
        self.assertEqual(parse_int("", 4), 4)
        self.assertEqual(parse_int(None, 4), 4)
        self.assertEqual(parse_int("oops", 4), 4)

    def test_parse_choice_accepts_known_values_only(self):
        self.assertEqual(parse_choice(" review ", "auto", {"auto", "review"}), "review")
        self.assertEqual(parse_choice("bad", "auto", {"auto", "review"}), "auto")

    def test_settings_from_env_uses_float_defaults_for_invalid_values(self):
        env = {
            key: value
            for key, value in os.environ.items()
            if key
            not in {
                "LOCAL_OCR_CONFIDENCE_THRESHOLD",
                "AMOUNT_TOLERANCE_MXN",
                "AI_VISUAL_COUNT_MIN_OPENCV_CROPS",
                "PAIRING_MODE",
                "QWEN_API_KEY",
                "DASHSCOPE_API_KEY",
                "QWEN_MODEL",
                "QWEN_BASE_URL",
                "ENABLE_QWEN_SCAN",
                "TESSERACT_CMD",
                "TESSERACT_LANG",
                "TESSERACT_PSM",
                "COMPANY_PROFILE",
            }
        }
        env["LOCAL_OCR_CONFIDENCE_THRESHOLD"] = "oops"
        env["AMOUNT_TOLERANCE_MXN"] = " "
        env["AI_VISUAL_COUNT_MIN_OPENCV_CROPS"] = "not-int"
        env["PAIRING_MODE"] = "bad-mode"
        with patch.dict(os.environ, env, clear=True):
            settings = Settings.from_env()
        self.assertEqual(settings.local_confidence_threshold, 0.62)
        self.assertEqual(settings.amount_tolerance, 0.50)
        self.assertEqual(settings.ai_visual_count_min_opencv_crops, 4)
        self.assertEqual(settings.pairing_mode, "auto")

    def test_settings_from_env_reads_company_profile(self):
        with patch.dict(os.environ, {"COMPANY_PROFILE": "acme"}, clear=True):
            settings = Settings.from_env()

        self.assertEqual(settings.company_profile, "acme")

    def test_settings_from_env_reads_qwen_scan_settings(self):
        env = {
            "QWEN_API_KEY": "qwen-token",
            "QWEN_MODEL": "qwen-vl-plus",
            "QWEN_BASE_URL": "https://example.test/v1/chat/completions",
            "ENABLE_QWEN_SCAN": "true",
        }
        with patch.dict(os.environ, env, clear=True):
            settings = Settings.from_env()

        self.assertEqual(settings.qwen_api_key, "qwen-token")
        self.assertEqual(settings.qwen_model, "qwen-vl-plus")
        self.assertEqual(settings.qwen_base_url, "https://example.test/v1/chat/completions")
        self.assertTrue(settings.qwen_scan_enabled)

    def test_settings_from_env_ignores_removed_openai_fallback(self):
        env = {
            "OPENAI_API_KEY": "openai-token",
            "OPENAI_MODEL": "gpt-4.1",
            "ENABLE_CODEX_SCAN": "true",
            "ENABLE_AI_VISUAL_COUNT": "true",
        }
        with patch.dict(os.environ, env, clear=True):
            settings = Settings.from_env()

        self.assertIsNone(settings.openai_api_key)
        self.assertEqual(settings.openai_model, "")
        self.assertFalse(settings.codex_scan_enabled)
        self.assertFalse(settings.ai_visual_count_enabled)

    def test_settings_from_env_reads_tesseract_settings(self):
        env = {
            "TESSERACT_CMD": r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            "TESSERACT_LANG": "eng+spa",
            "TESSERACT_PSM": "11",
        }
        with patch.dict(os.environ, env, clear=True):
            settings = Settings.from_env()

        self.assertEqual(settings.tesseract_cmd, r"C:\Program Files\Tesseract-OCR\tesseract.exe")
        self.assertEqual(settings.tesseract_lang, "eng+spa")
        self.assertEqual(settings.tesseract_psm, "11")


if __name__ == "__main__":
    unittest.main()
