from __future__ import annotations

from pathlib import Path

from .models import InvoiceRecord


def is_noise_record(record: InvoiceRecord) -> bool:
    seller = (record.seller or "").strip().casefold()
    return record.total_amount <= 0 and seller in {"", "unknown"}


def should_delete_failed_crop(record: InvoiceRecord, image_path: Path, *, ocr_text: str = "", ocr_error: str = "") -> bool:
    """Return true only for likely noise, not for readable receipts needing review."""

    if not is_noise_record(record):
        return False
    if is_obvious_background_crop(image_path):
        return True
    text = f"{ocr_text}\n{ocr_error}\n{record.remarks}".casefold()
    if _explicitly_not_a_receipt(text):
        return True
    if _looks_like_receipt_text(text):
        return False
    metrics = image_quality_metrics(image_path)
    if not metrics:
        return True
    width = metrics.get("width", 0.0)
    height = metrics.get("height", 0.0)
    if width >= 900 and height >= 900 and metrics.get("contrast", 0.0) >= 18:
        return False
    return True


def is_obvious_background_crop(image_path: Path) -> bool:
    """Reject only nearly textureless, non-paper regions before remote OCR."""

    metrics = image_quality_metrics(image_path)
    if not metrics:
        return False
    flat_dark_region = (
        metrics.get("bright_fraction", 1.0) < 0.03
        and metrics.get("edge_density", 1.0) < 0.0015
        and metrics.get("contrast", 100.0) < 30.0
    )
    dark_blurred_object = (
        metrics.get("bright_fraction", 1.0) < 0.01
        and metrics.get("edge_density", 1.0) < 0.0015
        and metrics.get("sharpness", 100.0) < 4.0
    )
    return flat_dark_region or dark_blurred_object


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
        edges = cv2.Canny(image, 50, 150)
        return {
            "width": float(image.shape[1]),
            "height": float(image.shape[0]),
            "brightness": float(image.mean()),
            "contrast": float(image.std()),
            "sharpness": float(cv2.Laplacian(image, cv2.CV_64F).var()),
            "bright_fraction": float((image >= 160).mean()),
            "edge_density": float((edges > 0).mean()),
        }
    except Exception:
        return None


def _looks_like_receipt_text(text: str) -> bool:
    keywords = (
        "receipt",
        "invoice",
        "factura",
        "comprobante",
        "importe",
        "total",
        "monto",
        "pago",
        "payment",
        "propina",
        "iva",
        "mxn",
        "pesos",
        "cfe",
        "comision federal",
        "multipagos",
        "tarjeta",
        "autorizacion",
        "referencia",
    )
    return any(keyword in text for keyword in keywords)


def _explicitly_not_a_receipt(text: str) -> bool:
    phrases = (
        "does not contain a receipt",
        "does not contain an invoice",
        "does not contain a receipt or invoice",
        "not a receipt or invoice",
        "no receipt or invoice",
        "no invoice or receipt",
        "not an invoice or receipt",
    )
    return any(phrase in text for phrase in phrases)
