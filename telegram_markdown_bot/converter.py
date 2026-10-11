"""Local document-to-Markdown conversion: no LLM or paid API is used."""
from __future__ import annotations

import csv
import html
import io
import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

import pymupdf
from bs4 import BeautifulSoup
from docx import Document
from docx.table import Table as DocxTable
from docx.text.paragraph import Paragraph as DocxParagraph
from markdownify import markdownify
from openpyxl import load_workbook
from PIL import Image, ImageOps, UnidentifiedImageError
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

SUPPORTED = {'.pdf', '.docx', '.pptx', '.xlsx', '.txt', '.md', '.csv', '.html', '.htm', '.png', '.jpg', '.jpeg', '.webp'}
OFFICE_ZIP = {'.docx', '.pptx', '.xlsx'}
MAX_PDF_PAGES = 40
MAX_OFFICE_UNCOMPRESSED = 100 * 1024 * 1024
MAX_OUTPUT_CHARS = 8_000_000
Image.MAX_IMAGE_PIXELS = 28_000_000


class ConversionError(Exception):
    """Expected input or conversion issue; safe to display to Telegram user."""


@dataclass
class ConversionResult:
    markdown: str
    raw_markdown: str
    warnings: list[str] = field(default_factory=list)


def safe_table(rows: list[list[str]]) -> str:
    """Return a valid GFM table, with no loss of nonempty cell content."""
    if not rows:
        return ''
    width = max(len(row) for row in rows)
    if not width:
        return ''
    def cell(v: str) -> str:
        return str(v).replace('|', r'\|').replace('\r', '').replace('\n', ' / ').strip()
    padded = [[cell(row[i]) if i < len(row) else '' for i in range(width)] for row in rows]
    # Do not treat a data row as a header. Use neutral numbered columns.
    header = [f'عمود {i + 1}' for i in range(width)]
    parts = ['| ' + ' | '.join(header) + ' |', '| ' + ' | '.join(['---'] * width) + ' |']
    for row in padded:
        parts.append('| ' + ' | '.join(row) + ' |')
    return '\n'.join(parts)


def assert_safe_office_zip(path: Path) -> None:
    try:
        with zipfile.ZipFile(path) as archive:
            infos = archive.infolist()
            if len(infos) > 2500 or sum(info.file_size for info in infos) > MAX_OFFICE_UNCOMPRESSED:
                raise ConversionError('ملف Office كبير جدًا بعد فك الضغط؛ لا يمكن معالجته بأمان.')
            if any(info.file_size > 30 * 1024 * 1024 for info in infos):
                raise ConversionError('يحتوي الملف على جزء داخلي كبير جدًا.')
    except zipfile.BadZipFile as ex:
        raise ConversionError('ملف Office تالف أو امتداده غير صحيح.') from ex


def clean_compact(markdown: str) -> str:
    """Remove layout noise only. Never summarize or delete content paragraphs."""
    out: list[str] = []
    for line in markdown.replace('\r\n', '\n').replace('\r', '\n').split('\n'):
        line = line.rstrip()
        if re.fullmatch(r'\s*[-_=*•·.]{5,}\s*', line):
            continue
        if re.fullmatch(r'\s*(?:Page\s+\d+(?:\s+of\s+\d+)?|صفحة\s+\d+(?:\s+من\s+\d+)?)\s*', line, re.I):
            continue
        if not line and out and not out[-1]:
            continue
        out.append(line)
    return '\n'.join(out).strip() + '\n'


def _pdf(path: Path, *, force_ocr: bool) -> tuple[str, list[str]]:
    warnings = []
    try:
        doc = pymupdf.open(str(path))
        try:
            if doc.needs_pass:
                raise ConversionError('الملف محمي بكلمة مرور. أرسل نسخة غير محمية.')
            count = doc.page_count
            if count > MAX_PDF_PAGES:
                raise ConversionError(f'هذا الملف فيه {count} صفحة. الحد الحالي {MAX_PDF_PAGES} صفحة لتجنب توقف البوت.')
            if count == 0:
                raise ConversionError('ملف PDF لا يحتوي على صفحات.')
            lengths = [len(page.get_text().strip()) for page in doc]
            selectable = sum(lengths)
            scanned_pages = sum(length < 10 for length in lengths)
        finally:
            doc.close()
        if scanned_pages:
            warnings.append(f'يوجد {scanned_pages} صفحة دون نص رقمي واضح؛ تعتمد دقة ناتجها على OCR.')
        try:
            import pymupdf4llm
        except ImportError:
            pymupdf4llm = None
        if pymupdf4llm is not None:
            text = pymupdf4llm.to_markdown(
                str(path), ocr_language='ara+eng', force_ocr=force_ocr,
                use_ocr=True, write_images=False, embed_images=False,
                page_separators=False, show_progress=False,
            )
        else:
            # Basic fallback for local diagnostics; Docker includes pymupdf4llm.
            with pymupdf.open(str(path)) as d:
                pages: list[str] = []
                for pg in d:
                    pg_text = pg.get_text('text', sort=True).strip()
                    if force_ocr or not pg_text:
                        pg_text = _image_ocr(_pixmap_image(pg))
                    pages.append(pg_text)
                text = '\n\n'.join(pages)
            warnings.append('لم يتوفر محرك الجداول المتقدم؛ استخدم المتطلبات كاملة للحصول على أفضل نتيجة.')
        if not text or not text.strip():
            raise ConversionError('تعذّر استخراج النص. جرّب /ocr on ثم أرسل الملف مرة ثانية.')
        if force_ocr:
            warnings.append('استُخدم وضع OCR الإجباري؛ راجع الرموز الرياضية والأرقام بعناية.')
        warnings.append('الرسومات والصور غير النصية لا تتحول تلقائيًا إلى وصف دقيق.')
        return str(text), warnings
    except ConversionError:
        raise
    except Exception as ex:
        raise ConversionError('تعذّرت قراءة PDF. تحقق أن الملف صالح، أو جرّب وضع OCR.') from ex


def _pixmap_image(page) -> Image.Image:
    pix = page.get_pixmap(matrix=pymupdf.Matrix(1.9, 1.9), alpha=False)
    return Image.open(io.BytesIO(pix.tobytes('png')))


def _image_ocr(img: Image.Image) -> str:
    try:
        import pytesseract
        img = ImageOps.exif_transpose(img).convert('RGB')
        return pytesseract.image_to_string(img, lang='ara+eng', config='--psm 3').strip()
    except (ImportError, OSError) as ex:
        raise ConversionError('التعرّف على النص من الصور غير متاح. ثبّت Tesseract مع اللغتين ara وeng.') from ex
    except Exception as ex:
        raise ConversionError('فشلت قراءة النص من الصورة؛ تحقق من وضوحها.') from ex


def _image(path: Path) -> tuple[str, list[str]]:
    try:
        with Image.open(path) as img:
            if img.width * img.height > Image.MAX_IMAGE_PIXELS:
                raise ConversionError('أبعاد الصورة كبيرة جدًا؛ أرسل صورة أصغر.')
            result = _image_ocr(img)
    except (Image.DecompressionBombError, UnidentifiedImageError) as ex:
        raise ConversionError('الصورة تالفة أو عالية الدقة بصورة غير آمنة.') from ex
    if not result:
        raise ConversionError('لم يظهر نص يمكن استخراجه من الصورة.')
    return result, ['الصور المعالجة بـOCR قد تحتوي على أخطاء في الأرقام والرموز.']


def _docx(path: Path) -> tuple[str, list[str]]:
    assert_safe_office_zip(path)
    doc = Document(str(path))
    parts = []
    for block in doc.iter_inner_content():
        if isinstance(block, DocxParagraph):
            value = block.text.strip()
            if not value:
                continue
            style = block.style.name.lower() if block.style is not None else ''
            if style.startswith('heading'):
                level = re.search(r'\d+', style)
                level = max(1, min(6, int(level.group()) if level else 1))
                parts.append('#' * level + ' ' + value)
            elif any(s in style for s in ('list bullet', 'list paragraph', 'قائمة نقطية')):
                parts.append('- ' + value)
            elif 'list number' in style or 'قائمة رقمية' in style:
                parts.append('1. ' + value)
            else:
                parts.append(value)
        elif isinstance(block, DocxTable):
            parts.append(safe_table([[c.text for c in row.cells] for row in block.rows]))
    return '\n\n'.join(parts), ['الصور والأشكال المضمّنة في Word لا يتحول معناها تلقائيًا إلى نص.']


def _pptx(path: Path) -> tuple[str, list[str]]:
    assert_safe_office_zip(path)
    presentation = Presentation(str(path))
    parts = []
    for number, slide in enumerate(presentation.slides, start=1):
        parts.append(f'## شريحة {number}')
        for shape in sorted(slide.shapes, key=lambda s: (s.top, s.left)):
            if shape.has_text_frame:
                text = '\n'.join(p.text for p in shape.text_frame.paragraphs if p.text.strip()).strip()
                if text:
                    parts.append(text)
            if shape.has_table:
                table = shape.table
                parts.append(safe_table([[cell.text for cell in row.cells] for row in table.rows]))
    return '\n\n'.join(parts), ['الرسوم البيانية والصور في الشرائح لا يمكن فهمها بالكامل دون نموذج رؤية.']


def _xlsx(path: Path) -> tuple[str, list[str]]:
    assert_safe_office_zip(path)
    try:
        wb = load_workbook(str(path), read_only=True, data_only=True)
        contents = []
        for ws in wb.worksheets:
            contents.append('# ' + ws.title)
            if ws.max_column > 75 or (ws.max_row or 0) * (ws.max_column or 0) > 120_000:
                raise ConversionError('ورقة Excel كبيرة جدًا للتحويل في النسخة الحالية.')
            rows = []
            for row in ws.iter_rows(values_only=True):
                values = ['' if item is None else str(item) for item in row]
                if any(v.strip() for v in values):
                    rows.append(values)
            # Drop completely empty trailing columns without dropping nonempty cells.
            if rows:
                width = max(i + 1 for row in rows for i, v in enumerate(row) if v.strip())
                contents.append(safe_table([r[:width] for r in rows]))
        return '\n\n'.join(contents), ['قد لا تتوفر القيم المحسوبة للصيغ إذا لم يحفظ Excel ناتجها داخل الملف.']
    finally:
        if 'wb' in locals():
            wb.close()


def _text(path: Path) -> tuple[str, list[str]]:
    data = path.read_bytes()
    codecs = ('utf-16',) if data.startswith((b'\xff\xfe', b'\xfe\xff')) else ('utf-8-sig', 'cp1256', 'cp1252')
    for enc in codecs:
        try:
            decoded = data.decode(enc)
            # Avoid accidental UTF-16 on unmarked ASCII bytes.
            if '\x00' in decoded and not data.startswith((b'\xff\xfe', b'\xfe\xff')):
                continue
            return decoded, []
        except UnicodeError:
            pass
    raise ConversionError('تعذر قراءة ترميز الملف النصي.')


def _csv(path: Path) -> tuple[str, list[str]]:
    text, _ = _text(path)
    try:
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=',;\t')
        rows = list(csv.reader(io.StringIO(text), dialect))
    except csv.Error:
        rows = list(csv.reader(io.StringIO(text)))
    if len(rows) > 15_000:
        raise ConversionError('ملف CSV يحتوي على صفوف أكثر من الحد المسموح.')
    return safe_table(rows), []


def _html(path: Path) -> tuple[str, list[str]]:
    text, _ = _text(path)
    soup = BeautifulSoup(text, 'html.parser')
    for element in soup(['script', 'style', 'noscript']):
        element.decompose()
    body = soup.body or soup
    return markdownify(str(body), heading_style='ATX'), []


def convert_file(path: str | Path, mode: str = 'faithful', force_ocr: bool = False) -> ConversionResult:
    """Convert *one* incoming file. Does not transmit data to third parties."""
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix not in SUPPORTED:
        raise ConversionError('نوع الملف غير مدعوم. استخدم PDF أو DOCX أو PPTX أو XLSX أو TXT أو CSV أو HTML أو صورة.')
    if mode not in ('faithful', 'compact'):
        raise ValueError('Unknown conversion mode')
    handlers = {
        '.pdf': lambda: _pdf(path, force_ocr=force_ocr),
        '.docx': lambda: _docx(path),
        '.pptx': lambda: _pptx(path),
        '.xlsx': lambda: _xlsx(path),
        '.txt': lambda: _text(path),
        '.md': lambda: _text(path),
        '.csv': lambda: _csv(path),
        '.html': lambda: _html(path),
        '.htm': lambda: _html(path),
        '.png': lambda: _image(path),
        '.jpg': lambda: _image(path),
        '.jpeg': lambda: _image(path),
        '.webp': lambda: _image(path),
    }
    try:
        raw, warnings = handlers[suffix]()
    except ConversionError:
        raise
    except Exception as ex:
        raise ConversionError('تعذّر تحويل هذا الملف. ربما هو تالف أو لا يحتوي على نص قابل للاستخراج.') from ex
    raw = raw.replace('\r\n', '\n').replace('\r', '\n').strip() + '\n'
    if not raw.strip():
        raise ConversionError('الملف لا يحتوي على نص قابل للاستخراج.')
    out = clean_compact(raw) if mode == 'compact' else raw
    if len(out) > MAX_OUTPUT_CHARS:
        raise ConversionError('الناتج كبير جدًا؛ قسّم الملف إلى أجزاء أصغر.')
    return ConversionResult(markdown=out, raw_markdown=raw, warnings=warnings)


def estimate_tokens(text: str) -> tuple[int, str]:
    """Approximate GPT-family token count; not a guarantee for any specific model."""
    try:
        import tiktoken
        enc = tiktoken.get_encoding('cl100k_base')
        return len(enc.encode(text)), 'cl100k_base'
    except Exception:
        # Conservative approximate fallback; avoid presenting as an exact tokenizer count.
        return max(1, round(len(text) / 3)), 'تقريب بالحروف'
