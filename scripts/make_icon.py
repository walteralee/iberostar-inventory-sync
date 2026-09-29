"""
Genera el icono de la aplicación (PNG para el frontend e ICO para el
ejecutable y el instalador). Solo hace falta ejecutarlo si se quiere
cambiar el diseño:

    python scripts/make_icon.py
"""

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parents[1]
SIZE = 1024
PRIMARY_DARK = (8, 58, 79)
PRIMARY = (11, 79, 108)
ACCENT = (201, 162, 39)


def build() -> Image.Image:
    gradient = Image.new("RGB", (SIZE, SIZE))
    pixels = gradient.load()

    for y in range(SIZE):
        for x in range(SIZE):
            t = (x + y) / (2 * SIZE)
            pixels[x, y] = tuple(
                round(a + (b - a) * t) for a, b in zip(PRIMARY_DARK, PRIMARY)
            )

    mask = Image.new("L", (SIZE, SIZE), 0)
    ImageDraw.Draw(mask).rounded_rectangle(
        (0, 0, SIZE - 1, SIZE - 1), radius=int(SIZE * 0.22), fill=255
    )

    icon = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    icon.paste(gradient, (0, 0), mask)

    draw = ImageDraw.Draw(icon)

    # Franja dorada inferior, como un libro de registro.
    draw.rounded_rectangle(
        (int(SIZE * 0.22), int(SIZE * 0.74), int(SIZE * 0.78), int(SIZE * 0.79)),
        radius=int(SIZE * 0.025),
        fill=ACCENT,
    )

    font = ImageFont.truetype("C:/Windows/Fonts/segoeuib.ttf", int(SIZE * 0.44))
    draw.text(
        (SIZE / 2, SIZE * 0.45),
        "IB",
        font=font,
        fill=(255, 255, 255),
        anchor="mm",
    )

    return icon


def main() -> None:
    icon = build()

    png_path = ROOT / "app" / "frontend" / "assets" / "icon.png"
    icon.resize((256, 256), Image.LANCZOS).save(png_path)

    ico_path = ROOT / "packaging" / "icon.ico"
    icon.save(
        ico_path,
        sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)],
    )

    print(f"Iconos generados:\n  {png_path}\n  {ico_path}")


if __name__ == "__main__":
    main()
