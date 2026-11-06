"""Формирование «Отчёта о выполненной работе» (.doc) по шаблону через Microsoft Word."""
import shutil
from pathlib import Path

import pythoncom
import win32com.client

from materials import Material, TEMPLATES_DIR, format_qty

TEMPLATE = TEMPLATES_DIR / "отчет_о_работе.doc"
PLACEHOLDER_AFTER = "использованы материалы:"
ATTACHMENT_NOTE = "Список использованных материалов прилагается."

# Константы Word
WD_FORMAT_DOCUMENT = 0  # .doc (Word 97–2003)
WD_FORMAT_PDF = 17
WD_PAGE_BREAK = 7
WD_ALIGN_LEFT, WD_ALIGN_CENTER, WD_ALIGN_RIGHT, WD_ALIGN_JUSTIFY = 0, 1, 2, 3
WD_CELL_ALIGN_VCENTER = 1
WD_LINE_SINGLE = 0


class ReportError(Exception):
    pass


def build_report(out_path: str | Path, reason: str, materials: list[Material],
                 pdf_copy: str | Path | None = None) -> None:
    """Копирует шаблон в out_path и заполняет его. Шрифты берутся из шаблона."""
    out_path = Path(out_path)
    if not TEMPLATE.is_file():
        raise ReportError(f"Не найден шаблон отчёта: {TEMPLATE}")
    try:
        shutil.copyfile(TEMPLATE, out_path)
    except PermissionError:
        raise ReportError(f"Файл занят (открыт в Word?): {out_path}")

    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(str(out_path), False, False, False)
        _fill_reason(doc, reason)
        _append_materials_page(doc, materials)
        doc.Save()
        if pdf_copy:
            doc.ExportAsFixedFormat(str(pdf_copy), WD_FORMAT_PDF)
    finally:
        if doc is not None:
            doc.Close(False)
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()


def _find_paragraph(doc, text: str):
    for i in range(1, doc.Paragraphs.Count + 1):
        p = doc.Paragraphs(i)
        if text.lower() in p.Range.Text.lower():
            return p
    return None


def _fill_reason(doc, reason: str) -> None:
    p = _find_paragraph(doc, PLACEHOLDER_AFTER)
    if p is None:
        raise ReportError(f"В шаблоне не найден текст «{PLACEHOLDER_AFTER}»")
    r = p.Range
    start = r.Start + r.Text.lower().index(PLACEHOLDER_AFTER) + len(PLACEHOLDER_AFTER)
    tail = doc.Range(start, r.End - 1)  # до знака конца абзаца
    body = " ".join(reason.split())
    tail.Text = f" {body} {ATTACHMENT_NOTE}" if body else f" {ATTACHMENT_NOTE}"


def _append_materials_page(doc, materials: list[Material]) -> None:
    font_name = doc.Paragraphs(_find_index(doc, PLACEHOLDER_AFTER)).Range.Font.Name
    # Значения копируем сразу: COM-объект абзаца «съезжает» после вставок в документ.
    sig_par = _find_paragraph(doc, "Подпись ответственного лица")
    signature = None
    if sig_par is not None:
        signature = {
            "text": sig_par.Range.Text.rstrip("\r"),
            "font": sig_par.Range.Font.Name,
            "size": sig_par.Range.Font.Size,
            "align": sig_par.Alignment,
        }

    title = _new_last_paragraph(doc, "Использованные материалы")
    _style_paragraph(title, font_name, 14, bold=True, align=WD_ALIGN_CENTER, after=12)
    title.ParagraphFormat.PageBreakBefore = True  # вторая страница
    title.InsertParagraphAfter()

    has_notes = any(m.note for m in materials)
    headers = ["№ п/п", "Наименование", "Код", "Ед. изм.", "Кол-во"]
    widths = [32, 270, 55, 55, 55]  # пт, в сумме ≈ ширине текста страницы
    if has_notes:
        headers.append("Примечание")
        widths = [28, 200, 45, 40, 45, 110]

    anchor = doc.Paragraphs(doc.Paragraphs.Count).Range
    _style_paragraph(anchor, font_name, 12, bold=False, align=WD_ALIGN_LEFT, after=0)
    rows = [h for h in [headers]] + [
        [str(i), m.name, m.code, m.unit, format_qty(m.qty)] + ([m.note] if has_notes else [])
        for i, m in enumerate(materials, 1)
    ]
    table = doc.Tables.Add(anchor, len(rows), len(headers))
    table.Borders.Enable = True
    table.Range.Font.Name = font_name
    table.Range.Font.Size = 11
    table.Range.Font.Bold = False
    table.Range.ParagraphFormat.SpaceAfter = 0
    table.Range.ParagraphFormat.SpaceBefore = 0
    table.Range.ParagraphFormat.FirstLineIndent = 0
    table.Range.ParagraphFormat.LineSpacingRule = WD_LINE_SINGLE
    table.Range.ParagraphFormat.PageBreakBefore = False
    table.TopPadding = table.BottomPadding = 1
    table.Rows.AllowBreakAcrossPages = False
    table.Rows(1).HeadingFormat = True  # шапка повторяется на следующих страницах
    for c, w in enumerate(widths, 1):
        table.Columns(c).Width = w

    for r, row in enumerate(rows, 1):
        for c, value in enumerate(row, 1):
            cell = table.Cell(r, c)
            cell.Range.Text = value
            cell.VerticalAlignment = WD_CELL_ALIGN_VCENTER
            name_col = headers[c - 1] in ("Наименование", "Примечание")
            cell.Range.ParagraphFormat.Alignment = (
                WD_ALIGN_LEFT if (name_col and r > 1) else WD_ALIGN_CENTER)
    header = table.Rows(1)
    header.Range.Font.Bold = True
    header.Shading.BackgroundPatternColor = 0xEEEEEE  # BGR, светло-серый

    if signature is not None:
        last = _new_last_paragraph(doc, signature["text"])
        _style_paragraph(last, signature["font"], signature["size"], bold=False,
                         align=signature["align"], after=0)
        last.ParagraphFormat.SpaceBefore = 24


def _new_last_paragraph(doc, text: str):
    """Добавляет абзац в конец документа и возвращает его диапазон (без знака абзаца)."""
    doc.Content.InsertParagraphAfter()
    p = doc.Paragraphs(doc.Paragraphs.Count).Range
    p.InsertBefore(text)
    p = doc.Paragraphs(doc.Paragraphs.Count).Range
    return doc.Range(p.Start, p.End - 1) if text else p


def _find_index(doc, text: str) -> int:
    for i in range(1, doc.Paragraphs.Count + 1):
        if text.lower() in doc.Paragraphs(i).Range.Text.lower():
            return i
    return doc.Paragraphs.Count


def _style_paragraph(rng, font: str, size: float, bold: bool, align: int, after: float) -> None:
    rng.Font.Name = font
    rng.Font.Size = size
    rng.Font.Bold = bold
    pf = rng.ParagraphFormat
    pf.Alignment = align
    pf.FirstLineIndent = 0
    pf.LeftIndent = 0
    pf.SpaceBefore = 0
    pf.SpaceAfter = after
    pf.PageBreakBefore = False
