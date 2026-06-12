"""Generuje prostą ikonę aplikacji (app.ico) programowo — bez plików zewnętrznych."""

from PIL import Image, ImageDraw

SIZES = [16, 24, 32, 48, 64, 128, 256]

def draw_icon(size: int) -> Image.Image:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)

    pad = max(1, size // 16)
    # Tło — zaokrąglony prostokąt w firmowych kolorach aplikacji (niebieski)
    d.rounded_rectangle(
        [pad, pad, size - pad, size - pad],
        radius=size // 5,
        fill=(0, 120, 212, 255),       # #0078d4 — kolor przycisku „Uruchom analizę”
        outline=(0, 90, 158, 255),
        width=max(1, size // 32),
    )

    # Stylizowany „wykres/wyciąg”: kilka poziomych pasków
    bar_left  = size * 0.28
    bar_right = size * 0.72
    bar_h     = max(1, size // 16)
    ys = [size * 0.34, size * 0.50, size * 0.66]
    widths = [1.0, 0.7, 0.85]
    for y, w in zip(ys, widths):
        d.rounded_rectangle(
            [bar_left, y - bar_h / 2, bar_left + (bar_right - bar_left) * w, y + bar_h / 2],
            radius=bar_h / 2,
            fill=(255, 255, 255, 235),
        )

    return img


def main():
    base = draw_icon(256)
    imgs = [draw_icon(s) for s in SIZES]
    out = "app.ico"
    base.save(out, format="ICO", sizes=[(s, s) for s in SIZES], append_images=imgs[1:])
    print(f"Done: {out}")


if __name__ == "__main__":
    main()
