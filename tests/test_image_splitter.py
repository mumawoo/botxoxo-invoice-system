import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from invoice_system.image_splitter import (
    iter_images,
    _looks_like_edge_noise,
    _merge_boxes,
    _normalize_crop,
    _ocr_orientation_score_from_tsv,
    _ocr_scale_factor,
    _orient_text_upright,
    _overlaps,
    _pad_box,
    _parse_tesseract_rotate,
    _same_receipt_fragment,
    _upright_text_score,
    _windows_path_to_wsl,
)


class ImageSplitterTests(unittest.TestCase):
    def test_iter_images_finds_supported_files_recursively(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            nested = root / "nested"
            nested.mkdir()
            jpg = root / "receipt.JPG"
            png = nested / "receipt.png"
            txt = nested / "note.txt"
            jpg.write_bytes(b"jpg")
            png.write_bytes(b"png")
            txt.write_text("not an image", encoding="utf-8")

            self.assertEqual(set(iter_images(root)), {jpg, png})
            self.assertEqual(iter_images(jpg), [jpg])

    def test_overlaps_uses_smaller_box_ratio(self):
        self.assertTrue(_overlaps((0, 0, 100, 100), (50, 50, 130, 130)))
        self.assertFalse(_overlaps((0, 0, 100, 100), (100, 100, 200, 200)))

    def test_merge_boxes_keeps_neighboring_overlaps_separate(self):
        boxes = [(0, 0, 100, 100), (50, 50, 130, 130), (300, 300, 400, 400)]

        self.assertEqual(_merge_boxes(boxes), [(0, 0, 100, 100), (50, 50, 130, 130), (300, 300, 400, 400)])

    def test_merge_boxes_keeps_outer_box_when_inner_box_is_detected(self):
        boxes = [(654, 753, 887, 1087), (670, 924, 872, 1082)]

        self.assertEqual(_merge_boxes(boxes), [(654, 753, 887, 1087)])

    def test_merge_boxes_still_merges_near_duplicate_regions(self):
        boxes = [(100, 100, 300, 500), (110, 115, 295, 490)]

        self.assertEqual(_merge_boxes(boxes), [(100, 100, 300, 500)])

    def test_merge_boxes_merges_same_receipt_vertical_fragments(self):
        boxes = [(100, 100, 400, 360), (112, 330, 390, 650)]

        self.assertEqual(_merge_boxes(boxes), [(100, 100, 400, 650)])

    def test_same_receipt_fragment_rejects_diagonal_neighboring_receipts(self):
        self.assertFalse(_same_receipt_fragment((0, 0, 100, 100), (50, 50, 130, 130)))

    def test_edge_noise_filter_rejects_short_bottom_strips_only(self):
        self.assertTrue(_looks_like_edge_noise((536, 1137, 960, 1280), 960, 1280))
        self.assertFalse(_looks_like_edge_noise((654, 753, 887, 1087), 960, 1280))

    def test_pad_box_adds_15_percent_and_clamps_to_image_bounds(self):
        padded = _pad_box((100, 100, 300, 500), width=1000, height=1000, ratio=0.15)

        self.assertEqual(padded, (70, 40, 330, 560))
        self.assertEqual(_pad_box((5, 5, 100, 100), width=120, height=120, ratio=0.15), (0, 0, 114, 114))

    def test_normalize_crop_rotates_landscape_and_scales_minimum_to_1500(self):
        import cv2
        import numpy as np

        crop = np.zeros((400, 800, 3), dtype=np.uint8)

        normalized = _normalize_crop(crop, cv2)

        height, width = normalized.shape[:2]
        self.assertGreaterEqual(min(height, width), 1500)
        self.assertGreater(height, width)

    def test_normalize_crop_keeps_portrait_orientation_and_scales(self):
        import cv2
        import numpy as np

        crop = np.zeros((900, 450, 3), dtype=np.uint8)

        normalized = _normalize_crop(crop, cv2)

        height, width = normalized.shape[:2]
        self.assertEqual(min(height, width), 1500)
        self.assertGreater(height, width)

    def test_normalize_crop_flips_upside_down_text_when_detectable(self):
        import cv2
        import numpy as np

        crop = np.full((900, 450, 3), 255, dtype=np.uint8)
        cv2.putText(crop, "TOTAL 123.45", (35, 90), cv2.FONT_HERSHEY_SIMPLEX, 1.1, (0, 0, 0), 3)
        cv2.putText(crop, "CAFE XUAN", (35, 145), cv2.FONT_HERSHEY_SIMPLEX, 1.0, (0, 0, 0), 2)
        upside_down = cv2.rotate(crop, cv2.ROTATE_180)

        normalized = _normalize_crop(upside_down, cv2)

        self.assertGreater(_upright_text_score(normalized, cv2), _upright_text_score(cv2.rotate(normalized, cv2.ROTATE_180), cv2))

    def test_parse_tesseract_rotate(self):
        self.assertEqual(_parse_tesseract_rotate("Rotate: 180\nOrientation confidence: 12.0"), 180)
        self.assertIsNone(_parse_tesseract_rotate("no orientation"))

    def test_orient_text_does_not_flip_without_reliable_ocr_evidence(self):
        import cv2
        import numpy as np

        crop = np.full((900, 450, 3), 255, dtype=np.uint8)
        cv2.putText(crop, "TOTAL 123.45", (35, 760), cv2.FONT_HERSHEY_SIMPLEX, 1.1, (0, 0, 0), 3)

        with patch("invoice_system.image_splitter._tesseract_rotation", return_value=None), patch(
            "invoice_system.image_splitter._tesseract_ocr_rotation", return_value=None
        ):
            normalized = _orient_text_upright(crop, cv2)

        self.assertTrue(np.array_equal(normalized, crop))

    def test_ocr_orientation_score_prefers_confident_receipt_words(self):
        header = "level\tpage_num\tblock_num\tpar_num\tline_num\tword_num\tleft\ttop\twidth\theight\tconf\ttext\n"
        good = header + "\n".join(
            [
                "5\t1\t1\t1\t1\t1\t0\t0\t10\t10\t80\tFAMILY",
                "5\t1\t1\t1\t1\t2\t0\t0\t10\t10\t78\tDOLLAR",
                "5\t1\t1\t1\t1\t3\t0\t0\t10\t10\t70\tTOTAL",
                "5\t1\t1\t1\t1\t4\t0\t0\t10\t10\t82\t$3.32",
            ]
        )
        poor = header + "\n".join(
            [
                "5\t1\t1\t1\t1\t1\t0\t0\t10\t10\t30\tMii",
                "5\t1\t1\t1\t1\t2\t0\t0\t10\t10\t25\tVe",
            ]
        )

        self.assertGreater(_ocr_orientation_score_from_tsv(good), _ocr_orientation_score_from_tsv(poor) + 20)

    def test_ocr_scale_factor_enlarges_small_receipts_before_orientation_ocr(self):
        self.assertGreater(_ocr_scale_factor(681, 324), 2.0)
        self.assertLessEqual(681 * _ocr_scale_factor(681, 324), 2600)

    def test_windows_path_to_wsl(self):
        self.assertEqual(_windows_path_to_wsl(Path(r"C:\Users\donxi\AppData\Local\Temp\crop.png")), "/mnt/c/Users/donxi/AppData/Local/Temp/crop.png")


if __name__ == "__main__":
    unittest.main()
