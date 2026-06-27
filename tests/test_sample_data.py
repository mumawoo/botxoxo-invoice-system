import tempfile
import unittest
from pathlib import Path

from invoice_system.image_splitter import OpenCVInvoiceSplitter
from invoice_system.sample_data import create_synthetic_multi_receipt, create_synthetic_receipt


class SampleDataTests(unittest.TestCase):
    def test_creates_synthetic_receipt(self):
        with tempfile.TemporaryDirectory() as temp:
            output = create_synthetic_receipt(Path(temp) / "sample.jpg")
            self.assertTrue(output.exists())
            self.assertGreater(output.stat().st_size, 1000)

    def test_creates_synthetic_multi_receipt(self):
        with tempfile.TemporaryDirectory() as temp:
            output = create_synthetic_multi_receipt(Path(temp) / "multi.jpg")
            self.assertTrue(output.exists())
            self.assertGreater(output.stat().st_size, 1000)

    def test_synthetic_multi_receipt_splits_into_two_crops(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            output = create_synthetic_multi_receipt(root / "multi.jpg")
            splitter = OpenCVInvoiceSplitter(root / "crops")

            crops = splitter.split(output)

            self.assertEqual(len(crops), 2)
            self.assertTrue(all(crop.crop_path.exists() for crop in crops))


if __name__ == "__main__":
    unittest.main()
