import unittest
import tempfile
from pathlib import Path

from invoice_system.models import InvoiceRecord
from invoice_system.quality import image_quality_metrics, is_noise_record, poor_image_quality_reason


class QualityTests(unittest.TestCase):
    def test_zero_unknown_is_noise(self):
        self.assertTrue(is_noise_record(InvoiceRecord(total_amount=0, seller="Unknown")))

    def test_zero_known_seller_is_kept_for_review(self):
        self.assertFalse(is_noise_record(InvoiceRecord(total_amount=0, seller="CAFE XUAN")))

    def test_positive_unknown_is_kept_for_review(self):
        self.assertFalse(is_noise_record(InvoiceRecord(total_amount=10, seller="Unknown")))

    def test_low_contrast_image_is_poor_quality(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "flat.jpg"
            _write_flat_image(path)

            self.assertEqual(poor_image_quality_reason(path), "poor image quality")
            metrics = image_quality_metrics(path)
            self.assertIsNotNone(metrics)
            self.assertLess(metrics["contrast"], 12)

    def test_bright_readable_image_is_not_poor_quality(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "receipt.jpg"
            _write_receipt_like_image(path)

            self.assertIsNone(poor_image_quality_reason(path))


def _write_flat_image(path: Path) -> None:
    import cv2
    import numpy as np

    image = np.full((120, 120), 128, dtype=np.uint8)
    ok, encoded = cv2.imencode(".jpg", image)
    if not ok:
        raise RuntimeError("could not encode test image")
    encoded.tofile(str(path))


def _write_receipt_like_image(path: Path) -> None:
    import cv2
    import numpy as np

    image = np.full((300, 500), 255, dtype=np.uint8)
    cv2.putText(image, "CAFE XUAN", (30, 80), cv2.FONT_HERSHEY_SIMPLEX, 1.4, 0, 3)
    cv2.putText(image, "TOTAL $126.00", (30, 160), cv2.FONT_HERSHEY_SIMPLEX, 1.2, 0, 3)
    ok, encoded = cv2.imencode(".jpg", image, [int(cv2.IMWRITE_JPEG_QUALITY), 98])
    if not ok:
        raise RuntimeError("could not encode test image")
    encoded.tofile(str(path))


if __name__ == "__main__":
    unittest.main()
