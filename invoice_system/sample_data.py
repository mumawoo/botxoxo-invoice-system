from __future__ import annotations

from pathlib import Path


def create_synthetic_receipt(output_path: Path) -> Path:
    from PIL import Image, ImageDraw, ImageFont

    output_path.parent.mkdir(parents=True, exist_ok=True)
    image = Image.new("RGB", (1800, 2400), "white")
    draw = ImageDraw.Draw(image)
    font_big, font, font_small = _load_fonts()
    lines = [
        ("CAFE XUAN", font_big),
        ("RFC XUA260612AB1", font_small),
        ("Fecha 2026-06-12", font),
        ("Mesa 4 Ticket 12345", font_small),
        ("Consumo alimentos 100.00", font),
        ("IVA 16.00", font),
        ("Propina 10.00", font),
        ("TOTAL $126.00 MXN", font_big),
    ]

    y = 220
    for text, item_font in lines:
        draw.text((190, y), text, fill="black", font=item_font)
        y += 190 if item_font == font_big else 140
    draw.rectangle((110, 110, 1690, 2290), outline="black", width=8)
    image.save(output_path, quality=98)
    return output_path


def create_synthetic_multi_receipt(output_path: Path) -> Path:
    from PIL import Image, ImageDraw

    output_path.parent.mkdir(parents=True, exist_ok=True)
    image = Image.new("RGB", (2600, 1800), (215, 215, 215))
    draw = ImageDraw.Draw(image)
    font_big, font, font_small = _load_fonts()

    _draw_receipt(
        draw,
        box=(160, 150, 1160, 1650),
        lines=[
            ("CAFE XUAN", font_big),
            ("Fecha 2026-06-12", font),
            ("Consumo 100.00", font),
            ("IVA 16.00", font),
            ("TOTAL $116.00 MXN", font_big),
        ],
    )
    _draw_receipt(
        draw,
        box=(1440, 150, 2440, 1650),
        lines=[
            ("PANADERIA LUZ", font_big),
            ("Fecha 2026-06-13", font),
            ("Pan dulce 80.00", font),
            ("IVA 12.80", font),
            ("TOTAL $92.80 MXN", font_big),
        ],
    )
    image.save(output_path, quality=98)
    return output_path


def _draw_receipt(draw, box: tuple[int, int, int, int], lines: list[tuple[str, object]]) -> None:
    x1, y1, x2, y2 = box
    draw.rectangle(box, fill="white", outline="black", width=8)
    y = y1 + 120
    for text, item_font in lines:
        draw.text((x1 + 90, y), text, fill="black", font=item_font)
        y += 170 if y < y2 - 360 else 140


def _load_fonts():
    from PIL import ImageFont

    try:
        return (
            ImageFont.truetype("arial.ttf", 86),
            ImageFont.truetype("arial.ttf", 62),
            ImageFont.truetype("arial.ttf", 50),
        )
    except Exception:
        return ImageFont.load_default(), ImageFont.load_default(), ImageFont.load_default()
