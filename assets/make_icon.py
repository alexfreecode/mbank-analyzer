"""Buduje ikonę aplikacji (app.ico) z pliku źródłowego assets/logo.png.

Maskotka „Suma Wpłat” to sum (ryba) owinięty wokół kalkulatora — gra słów
„sum / suma”. Źródło ma białe tło, więc tutaj wycinamy je na przezroczystość,
przycinamy do samej maskotki i zapisujemy wszystkie rozmiary, których szuka
Windows (pasek zadań, pulpit, „Dodaj/usuń programy”).
"""

from pathlib import Path

from PIL import Image

ASSETS = Path(__file__).parent
SOURCE = ASSETS / "logo.png"
OUTPUT = ASSETS / "app.ico"

# Rozmiary wymagane przez Windows. Powyżej rozdzielczości źródła obraz jest
# powiększany, więc większe ikony są tak ostre, jak pozwala na to plik logo.png.
SIZES = [16, 24, 32, 48, 64, 128, 256]

# Progi wycinania białego tła. Piksel czysto biały znika, lekko szary (artefakt
# kompresji na krawędzi) robi się półprzezroczysty, a beż kalkulatora zostaje.
WHITE_FULL = 246      # powyżej — całkowicie przezroczysty
WHITE_EDGE = 224      # poniżej — całkowicie nieprzezroczysty

# Margines wokół maskotki (ułamek boku) — bez niego ikona dotyka krawędzi
MARGIN = 0.04


def _cut_background(img: Image.Image) -> Image.Image:
    """Zamienia białe tło na przezroczystość, z miękką krawędzią."""
    img = img.convert("RGBA")
    px = img.load()
    w, h = img.size
    span = WHITE_FULL - WHITE_EDGE
    for y in range(h):
        for x in range(w):
            r, g, b, _ = px[x, y]
            low = min(r, g, b)
            if low >= WHITE_FULL:
                alpha = 0
            elif low <= WHITE_EDGE:
                alpha = 255
            else:
                alpha = round(255 * (WHITE_FULL - low) / span)
            px[x, y] = (r, g, b, alpha)
    return img


def _square(img: Image.Image) -> Image.Image:
    """Przycina do samej maskotki i wstawia na kwadratowe, przezroczyste tło."""
    bbox = img.getbbox()
    if bbox:
        img = img.crop(bbox)
    side = max(img.size)
    canvas = round(side * (1 + 2 * MARGIN))
    out = Image.new("RGBA", (canvas, canvas), (0, 0, 0, 0))
    out.paste(img, ((canvas - img.width) // 2, (canvas - img.height) // 2))
    return out


def main() -> None:
    if not SOURCE.exists():
        raise SystemExit(f"Brak pliku źródłowego: {SOURCE}")

    base = _square(_cut_background(Image.open(SOURCE)))
    # Zapisujemy z największego rozmiaru — PIL sam skaluje go w dół do
    # pozostałych. Podanie mniejszego obrazu obcięłoby .ico do jego rozmiaru.
    largest = max(SIZES)
    base = base.resize((largest, largest), Image.LANCZOS)
    base.save(OUTPUT, format="ICO", sizes=[(s, s) for s in SIZES])
    print(f"Done: {OUTPUT}  (zrodlo {Image.open(SOURCE).size[0]} px)")


if __name__ == "__main__":
    main()
