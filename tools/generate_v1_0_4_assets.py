from __future__ import annotations

from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "src" / "assets"
ASSETS.mkdir(parents=True, exist_ok=True)


def font(size: int, bold: bool = False) -> ImageFont.ImageFont:
    candidates = [
        Path(r"C:\Windows\Fonts\segoeuib.ttf" if bold else r"C:\Windows\Fonts\segoeui.ttf"),
        Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
    ]
    for candidate in candidates:
        if candidate.exists():
            return ImageFont.truetype(str(candidate), size)
    return ImageFont.load_default()


def star_points(cx: float, cy: float, outer: float, inner: float) -> list[tuple[float, float]]:
    from math import cos, pi, sin

    points: list[tuple[float, float]] = []
    for index in range(10):
        radius = outer if index % 2 == 0 else inner
        angle = -pi / 2 + index * pi / 5
        points.append((cx + radius * cos(angle), cy + radius * sin(angle)))
    return points


def build_banner() -> None:
    width, height = 620, 78
    image = Image.new("RGBA", (width, height), "white")
    draw = ImageDraw.Draw(image)
    red = "#e21d2f"
    navy = "#101827"
    muted = "#667085"

    draw.rectangle((0, 0, 7, height), fill=red)
    draw.rectangle((0, height - 5, width, height), fill=red)
    draw.polygon(star_points(48, 37, 25, 11), fill=red)
    draw.text((86, 14), "AP HEX CONVERTER TOOL", font=font(20, True), fill=navy)
    draw.text((86, 44), "China Grove, NC  |  Access-control conversion", font=font(10), fill=muted)

    bx1, by1, bx2, by2 = 495, 17, 600, 51
    draw.rounded_rectangle((bx1, by1, bx2, by2), radius=7, fill="#f7f9fc", outline="#cfd7e6", width=1)
    label = "HEX TOOL"
    label_font = font(11, True)
    box = draw.textbbox((0, 0), label, font=label_font)
    text_width = box[2] - box[0]
    text_height = box[3] - box[1]
    draw.text(((bx1 + bx2 - text_width) / 2, (by1 + by2 - text_height) / 2 - 1), label, font=label_font, fill="#0759b8")
    image.save(ASSETS / "ap-window-accent.png")


def build_salesforce_icon() -> None:
    image = Image.new("RGBA", (40, 40), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    blue = "#0d9dda"
    draw.ellipse((4, 12, 24, 32), fill=blue)
    draw.ellipse((12, 7, 31, 31), fill=blue)
    draw.ellipse((21, 13, 36, 31), fill=blue)
    draw.rounded_rectangle((7, 18, 34, 33), radius=7, fill=blue)
    label_font = font(10, True)
    draw.text((12, 17), "SF", font=label_font, fill="white")
    image.save(ASSETS / "icon-salesforce.png")


def main() -> None:
    build_banner()
    build_salesforce_icon()
    print("Generated v1.0.4 banner and Salesforce icon.")


if __name__ == "__main__":
    main()
