"""Распознавание текста скана PDF встроенным OCR Windows (русский язык)."""
import asyncio
from dataclasses import dataclass

import cv2
import numpy as np
import pymupdf
from winrt.windows.globalization import Language
from winrt.windows.graphics.imaging import BitmapDecoder
from winrt.windows.media.ocr import OcrEngine
from winrt.windows.storage.streams import DataWriter, InMemoryRandomAccessStream

DPI = 300


@dataclass
class OcrLine:
    text: str
    x: float  # координаты в пикселях отрендеренной страницы
    y: float
    w: float
    h: float

    @property
    def cy(self) -> float:
        return self.y + self.h / 2


@dataclass
class OcrPage:
    lines: list[OcrLine]
    width: int
    height: int
    image: np.ndarray  # очищенное от рамок изображение страницы (grayscale)

    def ocr_region(self, x0: float, y0: float, x1: float, y1: float, zoom: float = 2.0,
                   lang: str = "ru") -> str:
        """Повторно распознаёт фрагмент страницы с увеличением (точнее для мелких деталей)."""
        x0, y0 = max(0, int(x0)), max(0, int(y0))
        x1, y1 = min(self.width, int(x1)), min(self.height, int(y1))
        crop = cv2.resize(self.image[y0:y1, x0:x1], None, fx=zoom, fy=zoom,
                          interpolation=cv2.INTER_CUBIC)
        crop = cv2.copyMakeBorder(crop, 40, 40, 40, 40, cv2.BORDER_CONSTANT, value=255)
        lines = asyncio.run(_ocr_png(cv2.imencode(".png", crop)[1].tobytes(), lang))
        return " ".join(l.text for l in lines)


async def _ocr_png(png: bytes, lang: str) -> list[OcrLine]:
    stream = InMemoryRandomAccessStream()
    writer = DataWriter(stream)
    writer.write_bytes(png)
    await writer.store_async()
    writer.detach_stream()
    stream.seek(0)

    decoder = await BitmapDecoder.create_async(stream)
    bitmap = await decoder.get_software_bitmap_async()

    engine = OcrEngine.try_create_from_language(Language(lang))
    if engine is None:
        raise RuntimeError(
            f"OCR для языка '{lang}' недоступен. Установите языковой пакет в Windows."
        )
    result = await engine.recognize_async(bitmap)

    lines = []
    for line in result.lines:
        words = list(line.words)
        if not words:
            continue
        x0 = min(w.bounding_rect.x for w in words)
        y0 = min(w.bounding_rect.y for w in words)
        x1 = max(w.bounding_rect.x + w.bounding_rect.width for w in words)
        y1 = max(w.bounding_rect.y + w.bounding_rect.height for w in words)
        lines.append(OcrLine(line.text, x0, y0, x1 - x0, y1 - y0))
    return lines


def remove_table_lines(gray: np.ndarray) -> np.ndarray:
    """Стирает линии рамок полей: OCR теряет буквы, чьи хвосты касаются рамки."""
    bw = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 25, 15
    )
    scale = gray.shape[1] / 2480  # ядра подобраны для 300 dpi
    h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(20, int(120 * scale)), 1))
    v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(15, int(70 * scale))))
    lines = cv2.morphologyEx(bw, cv2.MORPH_OPEN, h_kernel) | cv2.morphologyEx(
        bw, cv2.MORPH_OPEN, v_kernel
    )
    lines = cv2.dilate(lines, np.ones((3, 3), np.uint8))
    clean = gray.copy()
    clean[lines > 0] = 255
    return clean


def render_page(page: "pymupdf.Page") -> np.ndarray:
    pix = page.get_pixmap(dpi=DPI, colorspace=pymupdf.csGRAY)
    return np.frombuffer(pix.samples, np.uint8).reshape(pix.height, pix.width).copy()


def ocr_pdf(path: str, lang: str = "ru") -> list[OcrPage]:
    pages = []
    with pymupdf.open(path) as doc:
        for page in doc:
            img = remove_table_lines(render_page(page))
            png = cv2.imencode(".png", img)[1].tobytes()
            lines = asyncio.run(_ocr_png(png, lang))
            pages.append(OcrPage(lines, img.shape[1], img.shape[0], img))
    return pages


if __name__ == "__main__":
    import sys

    for p in ocr_pdf(sys.argv[1]):
        for ln in p.lines:
            print(f"{ln.x:6.0f} {ln.y:6.0f} {ln.w:6.0f}  {ln.text}")
