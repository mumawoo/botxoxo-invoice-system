import sys
import types
import unittest
import subprocess
from pathlib import Path
from unittest.mock import patch

from invoice_system.recognizers import PaddleOCRRecognizer, TesseractRecognizer, _tesseract_lines


class RecognizerTests(unittest.TestCase):
    def test_paddleocr_disables_slow_document_preprocessing_models(self):
        calls = {}

        class FakePaddleOCR:
            def __init__(self, **kwargs):
                calls.update(kwargs)

            def predict(self, path):
                return []

        fake_module = types.SimpleNamespace(PaddleOCR=FakePaddleOCR)
        with patch.dict(sys.modules, {"paddleocr": fake_module}):
            PaddleOCRRecognizer().recognize(Path("dummy.jpg"))

        self.assertFalse(calls["use_doc_orientation_classify"])
        self.assertFalse(calls["use_doc_unwarping"])
        self.assertFalse(calls["use_textline_orientation"])

    def test_tesseract_lines_parse_tsv_words_and_confidence(self):
        raw = "level\tconf\ttext\n5\t91.5\tCafe\n5\t80\tXuan\n5\t-1\t\n"

        lines = _tesseract_lines(raw)

        self.assertEqual([line.text for line in lines], ["Cafe", "Xuan"])
        self.assertEqual(lines[0].confidence, 0.915)

    def test_tesseract_recognizer_calls_cli_tsv_output(self):
        raw = "level\tconf\ttext\n5\t90\tCafe\n5\t90\tXuan\n5\t90\tTOTAL\n5\t90\t126.00\n"
        completed = subprocess.CompletedProcess(args=[], returncode=0, stdout=raw, stderr="")

        with patch("invoice_system.recognizers.subprocess.run", return_value=completed) as run:
            result = TesseractRecognizer().recognize(Path("receipt.jpg"))

        self.assertEqual(result.engine, "tesseract")
        self.assertEqual(result.error, "")
        self.assertGreater(result.confidence, 0)
        self.assertIn("tsv", run.call_args.args[0])


if __name__ == "__main__":
    unittest.main()
