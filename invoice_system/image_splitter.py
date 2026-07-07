from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

from .models import CropResult, IMAGE_SUFFIXES

_TESSERACT_COMMAND_CACHE: list[str] | None = None


def iter_images(path: Path) -> list[Path]:
    if path.is_file() and path.suffix.lower() in IMAGE_SUFFIXES:
        return [path]
    if not path.exists():
        return []
    return sorted(p for p in path.rglob("*") if p.is_file() and p.suffix.lower() in IMAGE_SUFFIXES)


class OpenCVInvoiceSplitter:
    def __init__(self, crops_dir: Path, *, local_orientation: bool = True) -> None:
        self.crops_dir = crops_dir
        self.local_orientation = local_orientation
        self.crops_dir.mkdir(parents=True, exist_ok=True)

    def split(self, image_path: Path) -> list[CropResult]:
        try:
            return self._split_with_opencv(image_path)
        except Exception:
            return [self._copy_original(image_path, 1)]

    def _split_with_opencv(self, image_path: Path) -> list[CropResult]:
        import cv2
        import numpy as np

        image = _read_image_with_exif_orientation(image_path, cv2, np)
        if image is None:
            return [self._copy_original(image_path, 1)]

        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        blurred = cv2.GaussianBlur(gray, (5, 5), 0)
        edges = cv2.Canny(blurred, 50, 150)
        contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        height, width = image.shape[:2]
        image_area = height * width
        boxes: list[tuple[int, int, int, int]] = []
        for contour in contours:
            x, y, w, h = cv2.boundingRect(contour)
            area = w * h
            box = (x, y, x + w, y + h)
            if image_area * 0.02 <= area <= image_area * 0.80 and w > 80 and h > 80 and not _looks_like_edge_noise(box, width, height):
                boxes.append(box)

        boxes = _merge_boxes(boxes)
        paper_boxes = _paper_region_boxes(image, cv2, np)
        if len(paper_boxes) > len(boxes) and (len(boxes) <= 1 or _box_area(boxes[0]) >= image_area * 0.18):
            boxes = paper_boxes
        if not boxes:
            return [self._copy_original(image_path, 1)]

        boxes.sort(key=lambda box: (box[1], box[0]))
        crops: list[CropResult] = []
        for index, box in enumerate(boxes, start=1):
            x1, y1, x2, y2 = _pad_box(box, width, height, 0.15)
            crop = image[y1:y2, x1:x2]
            crop = _normalize_crop(crop, cv2, local_orientation=self.local_orientation)
            out = self.crops_dir / f"{image_path.stem}_d{index:02d}.jpg"
            ok, encoded = cv2.imencode(".jpg", crop, [int(cv2.IMWRITE_JPEG_QUALITY), 98])
            if ok:
                encoded.tofile(str(out))
                crops.append(CropResult(image_path, out, index))
        return crops or [self._copy_original(image_path, 1)]

    def _copy_original(self, image_path: Path, index: int) -> CropResult:
        out = self.crops_dir / f"{image_path.stem}_d{index:02d}.jpg"
        try:
            import cv2
            import numpy as np

            image = _read_image_with_exif_orientation(image_path, cv2, np)
            if image is None:
                raise ValueError("image decode failed")
            image = _normalize_crop(image, cv2, local_orientation=self.local_orientation)
            ok, encoded = cv2.imencode(".jpg", image, [int(cv2.IMWRITE_JPEG_QUALITY), 98])
            if not ok:
                raise ValueError("jpeg encode failed")
            encoded.tofile(str(out))
        except Exception:
            out = self.crops_dir / f"{image_path.stem}_d{index:02d}{image_path.suffix.lower()}"
            if image_path.resolve() != out.resolve():
                shutil.copy2(image_path, out)
        return CropResult(image_path, out, index)


def _merge_boxes(boxes: list[tuple[int, int, int, int]]) -> list[tuple[int, int, int, int]]:
    merged: list[tuple[int, int, int, int]] = []
    for box in sorted(boxes, key=_box_area, reverse=True):
        if any(_contains(existing, box) for existing in merged):
            continue

        current = box
        next_merged: list[tuple[int, int, int, int]] = []
        for existing in merged:
            if _contains(current, existing):
                continue
            if _near_duplicate(current, existing) or _same_receipt_fragment(current, existing):
                current = _union(current, existing)
            else:
                next_merged.append(existing)
        next_merged.append(current)
        merged = next_merged
    return sorted(merged, key=lambda box: (box[1], box[0]))


def _overlaps(a: tuple[int, int, int, int], b: tuple[int, int, int, int]) -> bool:
    intersection = _intersection_area(a, b)
    if intersection <= 0:
        return False
    smaller = min(_box_area(a), _box_area(b))
    return intersection / max(smaller, 1) > 0.20


def _contains(outer: tuple[int, int, int, int], inner: tuple[int, int, int, int]) -> bool:
    if _box_area(outer) <= _box_area(inner):
        return False
    return _intersection_area(outer, inner) / max(_box_area(inner), 1) >= 0.85


def _near_duplicate(a: tuple[int, int, int, int], b: tuple[int, int, int, int]) -> bool:
    intersection = _intersection_area(a, b)
    if intersection <= 0:
        return False
    smaller = min(_box_area(a), _box_area(b))
    return intersection / max(smaller, 1) >= 0.78


def _same_receipt_fragment(a: tuple[int, int, int, int], b: tuple[int, int, int, int]) -> bool:
    """Merge boxes that look like repeated slices of the same physical receipt."""
    aw, ah = _box_size(a)
    bw, bh = _box_size(b)
    if min(aw, ah, bw, bh) <= 0:
        return False

    horizontal_overlap = _axis_overlap(a[0], a[2], b[0], b[2]) / max(min(aw, bw), 1)
    vertical_overlap = _axis_overlap(a[1], a[3], b[1], b[3]) / max(min(ah, bh), 1)
    x_center_delta = abs(((a[0] + a[2]) / 2) - ((b[0] + b[2]) / 2))
    y_center_delta = abs(((a[1] + a[3]) / 2) - ((b[1] + b[3]) / 2))
    width_ratio = min(aw, bw) / max(aw, bw)
    height_ratio = min(ah, bh) / max(ah, bh)
    vertical_gap = _axis_gap(a[1], a[3], b[1], b[3])
    horizontal_gap = _axis_gap(a[0], a[2], b[0], b[2])
    top_delta = abs(a[1] - b[1])
    bottom_delta = abs(a[3] - b[3])

    same_column_slice = (
        horizontal_overlap >= 0.72
        and width_ratio >= 0.62
        and x_center_delta <= max(aw, bw) * 0.18
        and vertical_gap <= min(ah, bh) * 0.18
    )
    same_row_slice = (
        vertical_overlap >= 0.72
        and height_ratio >= 0.82
        and y_center_delta <= max(ah, bh) * 0.18
        and horizontal_gap <= min(aw, bw) * 0.12
        and top_delta <= max(ah, bh) * 0.14
        and bottom_delta <= max(ah, bh) * 0.14
    )
    return same_column_slice or same_row_slice


def _box_size(box: tuple[int, int, int, int]) -> tuple[int, int]:
    return max(box[2] - box[0], 0), max(box[3] - box[1], 0)


def _axis_overlap(a1: int, a2: int, b1: int, b2: int) -> int:
    return max(0, min(a2, b2) - max(a1, b1))


def _axis_gap(a1: int, a2: int, b1: int, b2: int) -> int:
    if a2 < b1:
        return b1 - a2
    if b2 < a1:
        return a1 - b2
    return 0


def _intersection_area(a: tuple[int, int, int, int], b: tuple[int, int, int, int]) -> int:
    x_left = max(a[0], b[0])
    y_top = max(a[1], b[1])
    x_right = min(a[2], b[2])
    y_bottom = min(a[3], b[3])
    if x_right <= x_left or y_bottom <= y_top:
        return 0
    return (x_right - x_left) * (y_bottom - y_top)


def _box_area(box: tuple[int, int, int, int]) -> int:
    return max(box[2] - box[0], 0) * max(box[3] - box[1], 0)


def _union(a: tuple[int, int, int, int], b: tuple[int, int, int, int]) -> tuple[int, int, int, int]:
    return min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3])


def _looks_like_edge_noise(box: tuple[int, int, int, int], image_width: int, image_height: int) -> bool:
    width = box[2] - box[0]
    height = box[3] - box[1]
    if width <= 0 or height <= 0:
        return True
    aspect = width / height
    touches_bottom = box[3] >= image_height * 0.98
    short_strip = height <= image_height * 0.14
    very_wide = aspect >= 2.7
    return touches_bottom and short_strip and very_wide


def _paper_region_boxes(image, cv2, np) -> list[tuple[int, int, int, int]]:
    height, width = image.shape[:2]
    image_area = height * width
    if image_area <= 0:
        return []
    hsv = cv2.cvtColor(image, cv2.COLOR_BGR2HSV)
    mask = cv2.inRange(hsv, (0, 0, 125), (180, 85, 255))
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (11, 11))
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, kernel, iterations=2)
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, kernel, iterations=1)
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    boxes: list[tuple[int, int, int, int]] = []
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        area = w * h
        box = (x, y, x + w, y + h)
        if image_area * 0.02 <= area <= image_area * 0.70 and w > 80 and h > 80 and not _looks_like_edge_noise(box, width, height):
            boxes.append(box)
    return _merge_boxes(boxes)


def _pad_box(box: tuple[int, int, int, int], width: int, height: int, ratio: float) -> tuple[int, int, int, int]:
    x1, y1, x2, y2 = box
    pad_x = int((x2 - x1) * ratio)
    pad_y = int((y2 - y1) * ratio)
    return max(0, x1 - pad_x), max(0, y1 - pad_y), min(width, x2 + pad_x), min(height, y2 + pad_y)


def _read_image_with_exif_orientation(image_path: Path, cv2, np):
    try:
        from PIL import Image, ImageOps

        with Image.open(image_path) as image:
            image = ImageOps.exif_transpose(image).convert("RGB")
            array = np.array(image)
        return cv2.cvtColor(array, cv2.COLOR_RGB2BGR)
    except Exception:
        return cv2.imdecode(np.fromfile(str(image_path), dtype=np.uint8), cv2.IMREAD_COLOR)


def _normalize_crop(crop, cv2, *, local_orientation: bool = True):
    h, w = crop.shape[:2]
    if w > h:
        clockwise = cv2.rotate(crop, cv2.ROTATE_90_CLOCKWISE)
        counterclockwise = cv2.rotate(crop, cv2.ROTATE_90_COUNTERCLOCKWISE)
        crop = _orient_text_upright(_choose_more_upright(clockwise, counterclockwise, cv2), cv2) if local_orientation else clockwise
        h, w = crop.shape[:2]
    elif local_orientation:
        crop = _orient_text_upright(crop, cv2)
        h, w = crop.shape[:2]
    minimum = min(h, w)
    if minimum < 1500 and minimum > 0:
        scale = 1500 / minimum
        crop = cv2.resize(crop, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_CUBIC)
        if local_orientation:
            crop = _orient_text_upright(crop, cv2)
    return crop


def _orient_text_upright(crop, cv2):
    rotation = _tesseract_rotation(crop, cv2)
    if rotation in {90, 270}:
        return _rotate_degrees(crop, rotation, cv2)
    ocr_rotation = _tesseract_ocr_rotation(crop, cv2)
    if rotation == 180 and ocr_rotation != 180:
        return crop
    if ocr_rotation == 180:
        return _rotate_degrees(crop, 180, cv2)
    return crop


def _tesseract_rotation(crop, cv2) -> int | None:
    if _text_pixel_count(crop, cv2) < 50:
        return None
    command = _tesseract_command()
    if not command:
        return None
    try:
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as handle:
            temp_path = Path(handle.name)
        try:
            cv2.imwrite(str(temp_path), crop)
            output = _run_tesseract_osd(command, temp_path)
        finally:
            temp_path.unlink(missing_ok=True)
    except Exception:
        return None
    return _parse_tesseract_rotate(output)


def _tesseract_command() -> list[str] | None:
    global _TESSERACT_COMMAND_CACHE
    if _TESSERACT_COMMAND_CACHE:
        return list(_TESSERACT_COMMAND_CACHE)
    windows = shutil.which("tesseract")
    if windows:
        _TESSERACT_COMMAND_CACHE = [windows]
        return list(_TESSERACT_COMMAND_CACHE)
    wsl = shutil.which("wsl.exe") or shutil.which("wsl")
    if not wsl:
        return None
    try:
        result = subprocess.run(
            [wsl, "sh", "-lc", "command -v tesseract"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=20,
            check=False,
        )
    except Exception:
        return None
    if result.returncode == 0 and result.stdout.strip():
        _TESSERACT_COMMAND_CACHE = [wsl, "tesseract"]
        return list(_TESSERACT_COMMAND_CACHE)
    return None


def _run_tesseract_osd(command: list[str], image_path: Path) -> str:
    path_text = str(image_path)
    if command[0].lower().endswith("wsl.exe") or command[0].lower().endswith("wsl"):
        path_text = _windows_path_to_wsl(image_path)
    result = subprocess.run(
        [*command, path_text, "stdout", "--psm", "0", "-l", "osd"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=12,
        check=False,
    )
    return "\n".join(part for part in (result.stdout, result.stderr) if part)


def _tesseract_ocr_rotation(crop, cv2) -> int | None:
    if _text_pixel_count(crop, cv2) < 50:
        return None
    command = _tesseract_command()
    if not command:
        return None
    try:
        score_0 = _tesseract_ocr_score(command, crop, cv2)
        score_180 = _tesseract_ocr_score(command, cv2.rotate(crop, cv2.ROTATE_180), cv2)
    except Exception:
        return None
    return 180 if score_180 > score_0 + 8 else None


def _tesseract_ocr_score(command: list[str], crop, cv2) -> float:
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as handle:
        temp_path = Path(handle.name)
    try:
        h, w = crop.shape[:2]
        image = crop
        scale = _ocr_scale_factor(h, w)
        if scale != 1.0:
            image = cv2.resize(crop, (max(1, int(w * scale)), max(1, int(h * scale))), interpolation=cv2.INTER_AREA)
        cv2.imwrite(str(temp_path), image)
        path_text = str(temp_path)
        if command[0].lower().endswith("wsl.exe") or command[0].lower().endswith("wsl"):
            path_text = _windows_path_to_wsl(temp_path)
        result = subprocess.run(
            [*command, path_text, "stdout", "--psm", "6", "-l", "eng", "tsv"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=20,
            check=False,
        )
        return _ocr_orientation_score_from_tsv(result.stdout)
    finally:
        temp_path.unlink(missing_ok=True)


def _ocr_scale_factor(height: int, width: int) -> float:
    if height <= 0 or width <= 0:
        return 1.0
    minimum = min(height, width)
    maximum = max(height, width)
    scale = 1.0
    if minimum < 900:
        scale = 900 / minimum
    if maximum * scale > 2600:
        scale = 2600 / maximum
    return max(scale, 0.1)


def _ocr_orientation_score_from_tsv(output: str) -> float:
    words = 0
    conf_sum = 0.0
    digit_words = 0
    alpha_words = 0
    keyword_hits = 0
    keywords = {"total", "subtotal", "tax", "visa", "mastercard", "family", "dollar", "restaurant", "kitchen", "receipt", "approved"}
    for line in (output or "").splitlines()[1:]:
        cols = line.split("\t")
        if len(cols) < 12:
            continue
        text = cols[11].strip()
        if not text:
            continue
        try:
            confidence = float(cols[10])
        except ValueError:
            confidence = -1.0
        if confidence < 20:
            continue
        words += 1
        conf_sum += confidence
        lowered = text.casefold()
        if any(char.isdigit() for char in text):
            digit_words += 1
        if any(char.isalpha() for char in text):
            alpha_words += 1
        if any(keyword in lowered for keyword in keywords):
            keyword_hits += 1
    if words <= 0:
        return 0.0
    average_confidence = conf_sum / words
    return words + average_confidence * 0.35 + digit_words * 2.0 + alpha_words * 0.5 + keyword_hits * 4.0


def _windows_path_to_wsl(path: Path) -> str:
    resolved = path.resolve()
    drive = resolved.drive.rstrip(":").lower()
    rest = resolved.relative_to(resolved.anchor).as_posix()
    return f"/mnt/{drive}/{rest}"


def _parse_tesseract_rotate(output: str) -> int | None:
    import re

    rotate = re.search(r"Rotate:\s*(\d+)", output or "", re.IGNORECASE)
    if not rotate:
        return None
    value = int(rotate.group(1)) % 360
    return value if value in {0, 90, 180, 270} else None


def _rotate_degrees(image, degrees: int, cv2):
    degrees = degrees % 360
    if degrees == 90:
        return cv2.rotate(image, cv2.ROTATE_90_CLOCKWISE)
    if degrees == 180:
        return cv2.rotate(image, cv2.ROTATE_180)
    if degrees == 270:
        return cv2.rotate(image, cv2.ROTATE_90_COUNTERCLOCKWISE)
    return image


def _choose_more_upright(first, second, cv2):
    first_score = _upright_text_score(first, cv2)
    second_score = _upright_text_score(second, cv2)
    return second if second_score > first_score + 0.03 else first


def _upright_text_score(image, cv2) -> float:
    h, w = image.shape[:2]
    if h <= 0 or w <= 0:
        return 0.0
    scale = min(1.0, 900 / max(h, w))
    if scale < 1.0:
        image = cv2.resize(image, (max(1, int(w * scale)), max(1, int(h * scale))), interpolation=cv2.INTER_AREA)
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    gray = cv2.GaussianBlur(gray, (3, 3), 0)
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 31, 12)
    binary = _likely_text_mask(binary, cv2)
    ys, _ = binary.nonzero()
    if len(ys) < 50:
        return 0.0
    height = binary.shape[0]
    top = (ys < height * 0.42).sum() / len(ys)
    bottom = (ys > height * 0.58).sum() / len(ys)
    center = float(ys.mean()) / max(height, 1)
    first_row = float(ys.min()) / max(height, 1)
    return (top - bottom) + (0.5 - center) * 0.35 + (0.15 - first_row) * 0.20


def _text_pixel_count(image, cv2) -> int:
    h, w = image.shape[:2]
    if h <= 0 or w <= 0:
        return 0
    scale = min(1.0, 900 / max(h, w))
    if scale < 1.0:
        image = cv2.resize(image, (max(1, int(w * scale)), max(1, int(h * scale))), interpolation=cv2.INTER_AREA)
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    gray = cv2.GaussianBlur(gray, (3, 3), 0)
    binary = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 31, 12)
    return int((_likely_text_mask(binary, cv2) > 0).sum())


def _likely_text_mask(binary, cv2):
    count, labels, stats, _ = cv2.connectedComponentsWithStats(binary, 8)
    mask = binary.copy()
    mask[:] = 0
    image_area = binary.shape[0] * binary.shape[1]
    for label in range(1, count):
        x, y, w, h, area = stats[label]
        if area <= 0:
            continue
        if area < 4 or area > image_area * 0.08:
            continue
        if h < 2 or w < 2:
            continue
        if w / max(h, 1) > 35 or h / max(w, 1) > 12:
            continue
        mask[labels == label] = 255
    return mask
