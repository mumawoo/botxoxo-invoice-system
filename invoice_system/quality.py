from __future__ import annotations

from pathlib import Path

from .models import InvoiceRecord


def is_noise_record(record: InvoiceRecord) -> bool:
    seller = (record.seller or "").strip().casefold()
    return record.total_amount <= 0 and seller in {"", "unknown"}


def poor_image_quality_reason(image_path: Path) -> str | None:
    metrics = image_quality_metrics(image_path)
    if metrics is None:
        return None
    brightness = metrics["brightness"]
    contrast = metrics["contrast"]
    sharpness = metrics["sharpness"]
    if contrast < 8:
        return "poor image quality"
    if brightness < 25 and contrast < 25:
        return "poor image quality"
    if brightness > 252 and contrast < 18:
        return "poor image quality"
    if sharpness < 8 and contrast < 25:
        return "poor image quality"
    return None


def image_quality_metrics(image_path: Path) -> dict[str, float] | None:
    try:
        import cv2
        import numpy as np

        image = cv2.imdecode(np.fromfile(str(image_path), dtype=np.uint8), cv2.IMREAD_GRAYSCALE)
        if image is None:
            return None
        return {
            "brightness": float(image.mean()),
            "contrast": float(image.std()),
            "sharpness": float(cv2.Laplacian(image, cv2.CV_64F).var()),
        }
    except Exception:
        return None
