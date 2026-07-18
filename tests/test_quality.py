import unittest
import tempfile
from pathlib import Path

from invoice_system.models import InvoiceRecord
from invoice_system.quality import (
    image_quality_metrics,
    is_noise_record,
    is_obvious_background_crop,
    poor_image_quality_reason,
    should_delete_failed_crop,
)


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

    def test_dark_low_detail_table_crop_is_filtered_before_remote_ocr(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "table.jpg"
            _write_flat_image(path, value=92)

            self.assertTrue(is_obvious_background_crop(path))

    def test_receipt_like_crop_is_not_filtered_before_remote_ocr(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "receipt.jpg"
            _write_receipt_like_image(path)

            self.assertFalse(is_obvious_background_crop(path))

    def test_dark_blurred_object_is_filtered_before_remote_ocr(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "dark-object.jpg"
            _write_dark_blurred_object(path)

            self.assertTrue(is_obvious_background_crop(path))

    def test_qwen_explicit_non_receipt_result_deletes_failed_crop(self):
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / "large-image.jpg"
            _write_receipt_like_image(path)
            record = InvoiceRecord(
                total_amount=0,
                seller="Unknown",
                remarks="Image does not contain a receipt or invoice; shows a watch strap on a wooden surface.",
            )

            self.assertTrue(
                should_delete_failed_crop(
                    record,
                    path,
                    ocr_text="Image does not contain a receipt or invoice.",
                )
            )


def _write_flat_image(path: Path, value: int = 128) -> None:
    import cv2
    import numpy as np

    image = np.full((120, 120), value, dtype=np.uint8)
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


def _write_dark_blurred_object(path: Path) -> None:
    import cv2
    import numpy as np

    image = np.full((1000, 400), 100, dtype=np.uint8)
    cv2.ellipse(image, (80, 600), (150, 500), 0, 0, 360, 15, -1)
    image = cv2.GaussianBlur(image, (0, 0), 25)
    ok, encoded = cv2.imencode(".jpg", image, [int(cv2.IMWRITE_JPEG_QUALITY), 90])
    if not ok:
        raise RuntimeError("could not encode test image")
    encoded.tofile(str(path))


if __name__ == "__main__":
    unittest.main()
