from __future__ import annotations

import hashlib
import re
from functools import lru_cache
from io import BytesIO
from pathlib import Path

from matplotlib import font_manager
from matplotlib.backends.backend_agg import FigureCanvasAgg
from matplotlib.figure import Figure
from matplotlib.mathtext import MathTextParser
from PIL import Image

from utils.config import TEMP_DIR


RENDER_CACHE_VERSION = "v3-cropped-sans"
DEFAULT_DPI = 300
DEFAULT_FONT_SIZE = 18
DEFAULT_PADDING_INCHES = 0.05
MATH_INLINE_PATTERN = re.compile(r"(?<!\\)\$(.*?)(?<!\\)\$|\\\$", re.DOTALL)
MATH_FONT_FAMILY = "dejavusans"


_parser: MathTextParser | None = None


def _get_parser() -> MathTextParser:
    global _parser
    if _parser is None:
        _parser = MathTextParser("path")
    return _parser


def _get_cache_dir() -> Path:
    cache_dir = TEMP_DIR / "latex_cache"
    cache_dir.mkdir(parents=True, exist_ok=True)
    return cache_dir


def _strip_math_delimiters(tex_string: str) -> str:
    tex_string = (tex_string or "").strip()
    if tex_string.startswith("$$") and tex_string.endswith("$$"):
        return tex_string[2:-2].strip()
    if tex_string.startswith("$") and tex_string.endswith("$"):
        return tex_string[1:-1].strip()
    return tex_string


def _normalize_math(tex_string: str) -> str:
    tex_string = _strip_math_delimiters(tex_string)
    tex_string = tex_string.replace("\r", " ").replace("\n", " ")
    tex_string = re.sub(r"\s+", " ", tex_string).strip()

    replacements = {
        r"\,": " ",
        r"\;": " ",
        r"\!": "",
        r"\textstyle": "",
        r"\displaystyle": "",
        r"\dfrac": r"\frac",
        r"\tfrac": r"\frac",
        r"\bm{": r"\mathbf{",
    }
    for source, target in replacements.items():
        tex_string = tex_string.replace(source, target)

    tex_string = re.sub(r"\\pmb\{([^{}]+)\}", r"\1", tex_string)
    tex_string = re.sub(r"\\text\{([^{}]+)\}", r"\\mathrm{\1}", tex_string)

    return tex_string.strip()


def _render_plain_text_png_bytes(text: str, dpi: int, font_size: int) -> tuple[bytes, int, int]:
    figure = Figure(facecolor="none", edgecolor="none", figsize=(0.01, 0.01))
    canvas = FigureCanvasAgg(figure)
    ax = figure.add_axes([0, 0, 1, 1])
    ax.axis("off")

    font_props = font_manager.FontProperties(family="Microsoft YaHei", size=font_size)
    text_artist = ax.text(0.02, 0.05, text, fontproperties=font_props, color="black", ha="left", va="bottom")
    canvas.draw()
    bbox = text_artist.get_window_extent(renderer=canvas.get_renderer())

    width_inches = max((bbox.width / dpi) + DEFAULT_PADDING_INCHES * 2, 0.2)
    height_inches = max((bbox.height / dpi) + DEFAULT_PADDING_INCHES * 2, 0.2)

    figure = Figure(facecolor="none", edgecolor="none", figsize=(width_inches, height_inches))
    canvas = FigureCanvasAgg(figure)
    ax = figure.add_axes([0, 0, 1, 1])
    ax.axis("off")
    ax.text(0.02, 0.05, text, fontproperties=font_props, color="black", ha="left", va="bottom")

    buffer = BytesIO()
    figure.savefig(
        buffer,
        format="png",
        dpi=dpi,
        facecolor="none",
        edgecolor="none",
        transparent=True,
        bbox_inches="tight",
        pad_inches=0.02,
    )
    buffer.seek(0)
    png_bytes = buffer.read()
    return png_bytes, max(int(width_inches * dpi), 1), max(int(height_inches * dpi), 1)


def _crop_transparent_padding(png_bytes: bytes, dpi: int) -> tuple[bytes, int, int]:
    image = Image.open(BytesIO(png_bytes)).convert("RGBA")
    alpha = image.getchannel("A")
    bbox = alpha.getbbox()
    if bbox:
        image = image.crop(bbox)

    output = BytesIO()
    image.save(output, format="PNG")
    output.seek(0)
    cropped = output.read()
    width, height = image.size
    return cropped, max(width, 1), max(height, 1)


def _render_math_png_bytes(tex_string: str, dpi: int, font_size: int) -> tuple[bytes, int, int]:
    tex_string = _normalize_math(tex_string)
    parser = _get_parser()
    font_props = font_manager.FontProperties(family="DejaVu Sans", size=font_size)
    metrics = parser.parse(tex_string, dpi=dpi, prop=font_props)

    width_inches = max((metrics.width / dpi) + DEFAULT_PADDING_INCHES * 2, 0.2)
    height_inches = max((metrics.height / dpi) + DEFAULT_PADDING_INCHES * 2, 0.2)

    figure = Figure(facecolor="none", edgecolor="none", figsize=(width_inches, height_inches))
    FigureCanvasAgg(figure)
    ax = figure.add_axes([0, 0, 1, 1])
    ax.axis("off")
    ax.text(
        0.02,
        0.05,
        f"${tex_string}$",
        fontproperties=font_props,
        math_fontfamily=MATH_FONT_FAMILY,
        color="black",
        ha="left",
        va="bottom",
    )

    buffer = BytesIO()
    figure.savefig(
        buffer,
        format="png",
        dpi=dpi,
        facecolor="none",
        edgecolor="none",
        transparent=True,
        bbox_inches="tight",
        pad_inches=0.02,
    )
    buffer.seek(0)
    png_bytes = buffer.read()
    cropped_bytes, cropped_width, cropped_height = _crop_transparent_padding(png_bytes, dpi=dpi)
    return cropped_bytes, cropped_width, cropped_height


@lru_cache(maxsize=2048)
def _render_cached(tex_string: str, dpi: int = DEFAULT_DPI, font_size: int = DEFAULT_FONT_SIZE) -> str:
    cache_dir = _get_cache_dir()
    normalized = _normalize_math(tex_string)
    key = hashlib.md5(f"{RENDER_CACHE_VERSION}|{normalized}|{dpi}|{font_size}".encode("utf-8")).hexdigest()
    output_path = cache_dir / f"latex_{key}.png"

    if output_path.exists():
        return str(output_path)

    try:
        png_bytes, _, _ = _render_math_png_bytes(normalized, dpi=dpi, font_size=font_size)
    except Exception:
        png_bytes, _, _ = _render_plain_text_png_bytes(normalized or tex_string, dpi=dpi, font_size=font_size)

    output_path.write_bytes(png_bytes)
    return str(output_path)


class LatexRenderer:
    def render_to_image(self, tex_string: str, font_size: int = DEFAULT_FONT_SIZE, dpi: int = DEFAULT_DPI) -> str:
        tex_string = _strip_math_delimiters(tex_string)
        return _render_cached(tex_string, dpi=dpi, font_size=font_size)

    def render_text_with_latex(self, text: str) -> list[tuple[str, str | None]]:
        parts: list[tuple[str, str | None]] = []
        last_end = 0

        for match in MATH_INLINE_PATTERN.finditer(text or ""):
            start = match.start()
            if start > last_end:
                plain_text = text[last_end:start]
                if plain_text:
                    parts.append((plain_text, None))

            formula = match.group(1)
            if formula is None:
                parts.append(("$", None))
            else:
                parts.append((formula, self.render_to_image(formula)))
            last_end = match.end()

        if last_end < len(text or ""):
            parts.append((text[last_end:], None))

        return parts


latex_renderer = LatexRenderer()
