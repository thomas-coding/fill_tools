from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
ICON_PNG = ROOT / "app_icon_preview.png"
ICON_ICO = ROOT / "app_icon.ico"


def pick_font(size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    candidates = [
        "C:/Windows/Fonts/msyhbd.ttc",
        "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/simhei.ttf",
        "C:/Windows/Fonts/simsun.ttc",
    ]
    for font_path in candidates:
        try:
            return ImageFont.truetype(font_path, size)
        except Exception:
            continue
    return ImageFont.load_default()


def draw_star(draw: ImageDraw.ImageDraw, cx: int, cy: int, r: int, color: str) -> None:
    points = [
        (cx, cy - r),
        (cx + r // 3, cy - r // 3),
        (cx + r, cy),
        (cx + r // 3, cy + r // 3),
        (cx, cy + r),
        (cx - r // 3, cy + r // 3),
        (cx - r, cy),
        (cx - r // 3, cy - r // 3),
    ]
    draw.polygon(points, fill=color)


def main() -> None:
    size = 512
    bg = "#F5B51E"
    fg = "#1F1F1F"

    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    radius = 118
    pad = 14
    draw.rounded_rectangle((pad, pad, size - pad, size - pad), radius=radius, fill=bg)

    font = pick_font(300)
    text = "盛"
    bbox = draw.textbbox((0, 0), text, font=font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    tx = (size - tw) // 2
    ty = (size - th) // 2 - 10
    draw.text((tx, ty), text, font=font, fill=fg)

    # 右下角做一个小白星，风格接近示例。
    draw_star(draw, int(size * 0.80), int(size * 0.78), 34, "#FFFFFF")
    draw_star(draw, int(size * 0.85), int(size * 0.83), 18, "#FFF6D8")

    img.save(ICON_PNG)
    img.save(ICON_ICO, format="ICO", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])

    print(f"Created: {ICON_ICO}")
    print(f"Preview: {ICON_PNG}")


if __name__ == "__main__":
    main()
