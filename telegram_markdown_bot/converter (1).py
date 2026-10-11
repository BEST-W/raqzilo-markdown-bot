"""Local document-to-Markdown conversion. No LLM or paid APIs are used.

Improved PDF extraction for 6-panel Arabic handouts, preserving superscripts.
"""
from __future__ import annotations

import csv
import io
import re
import unicodedata
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

SUPPORTED = {'.pdf', '.docx', '.pptx', '.xlsx', '.txt', '.md', '.csv', '.html', '.htm', '.png', '.jpg', '.jpeg', '.webp'}
MAX_PDF_PAGES = 40
MAX_OFFICE_UNCOMPRESSED = 100 * 1024 * 1024
MAX_OUTPUT_CHARS = 8_000_000
Image.MAX_IMAGE_PIXELS = 28_000_000

class ConversionError(Exception):
    """Conversion issue safe to show to the Telegram user."""

@dataclass
class ConversionResult:
    markdown: str
    raw_markdown: str
    warnings: list[str] = field(default_factory=list)


def safe_table(rows: list[list[str]]) -> str:
    if not rows:
        return ''
    width = max(len(row) for row in rows)
    if not width:
        return ''
    def cell(v: str) -> str:
        return str(v).replace('|', r'\|').replace('\r', '').replace('\n', ' / ').strip()
    padded = [[cell(row[i]) if i < len(row) else '' for i in range(width)] for row in rows]
    header = [f'عمود {i+1}' for i in range(width)]
    return '\n'.join(['| ' + ' | '.join(header) + ' |', '| ' + ' | '.join(['---']*width) + ' |'] + [
        '| ' + ' | '.join(row) + ' |' for row in padded])


def assert_safe_office_zip(path: Path) -> None:
    try:
        with zipfile.ZipFile(path) as archive:
            infos = archive.infolist()
            if len(infos) > 2500 or sum(i.file_size for i in infos) > MAX_OFFICE_UNCOMPRESSED:
                raise ConversionError('ملف Office كبير جدًا بعد فك الضغط؛ لا يمكن معالجته بأمان.')
            if any(i.file_size > 30*1024*1024 for i in infos):
                raise ConversionError('يحتوي الملف على جزء داخلي كبير جدًا.')
    except zipfile.BadZipFile as ex:
        raise ConversionError('ملف Office تالف أو امتداده غير صحيح.') from ex


def clean_compact(markdown: str) -> str:
    out = []
    for line in markdown.replace('\r\n','\n').replace('\r','\n').split('\n'):
        line=line.rstrip()
        if re.fullmatch(r'\s*[-_=*•·.]{5,}\s*',line):
            continue
        if re.fullmatch(r'\s*(?:Page\s+\d+(?:\s+of\s+\d+)?|صفحة\s+\d+(?:\s+من\s+\d+)?)\s*',line,re.I):
            continue
        if not line and out and not out[-1]:
            continue
        out.append(line)
    return '\n'.join(out).strip()+'\n'


def _math_line(line: dict) -> str | None:
    """Rebuild positioned math, distinguishing smaller adjacent exponent glyphs.

    Returns None for ordinary text, where the PDF's native RTL order is better.
    """
    spans = line.get('spans', [])
    if not spans:
        return None
    all_text = ''.join(s['text'] for s in spans)
    has_ops = any(ch in all_text for ch in '=×÷')
    numeric_only = len(spans) >= 2 and all(
        ch.isdigit() or ch.isspace() for s in spans for ch in s['text'])
    if not (has_ops or numeric_only):
        return None
    parts = []
    prior = None
    for span in sorted(spans, key=lambda s:s['bbox'][0]):
        text=span['text']
        if re.fullmatch(r'[\u0660-\u0669\u06f0-\u06f90-9]+',text):
            # Arabic numeral sequences in this font are stored opposite to display.
            if len(text)>1:
                text=text[::-1]
            text=''.join(str(unicodedata.digit(ch)) if ch.isdigit() else ch for ch in text)
            if (prior and prior['text'].strip().isdigit()
                    and span['size'] < prior['size'] * .82
                    and abs(span['bbox'][0]-prior['bbox'][2]) < 4):
                parts.append('^'+text)
            else:
                parts.append(text)
        else:
            parts.append(text)
        prior=span
    return ''.join(parts).strip()


def _looks_like_six_panel_handout(page: pymupdf.Page) -> bool:
    """Conservative detection: distinct three columns AND a full horizontal fold gap."""
    w,h=page.rect.width,page.rect.height
    if not (1.32 < h/w < 1.52):
        return False
    blocks=[b for b in page.get_text('dict')['blocks'] if b.get('type')==0]
    if len(blocks)<35:
        return False
    occupied=[0]*6
    top_bottom=[]
    bottom_top=[]
    for block in blocks:
        x0,y0,x1,y1=block['bbox']
        xc=(x0+x1)/2; yc=(y0+y1)/2
        c=min(2,max(0,int(xc/(w/3))))
        r=int(yc>=h/2)
        occupied[r*3+c]+=1
        if r==0: top_bottom.append(y1)
        else: bottom_top.append(y0)
    if not all(count>=4 for count in occupied):
        return False
    if not top_bottom or not bottom_top:
        return False
    # A three-column newspaper does not have a gap across the whole page.
    return (min(bottom_top)-max(top_bottom)) >= .025*h


# Arithmetic-only lines, kept deliberately conservative so words are never
# silently "corrected" based on a language model or a guessed equation.
_SIMPLE_MATH = re.compile(r"^[0-9\s×÷+\-^=()]+$")
_ARABIC_MULTI_NUMBER = re.compile(r"[٠-٩۰-۹]{2,}")


def _latin_digits(s: str) -> str:
    return "".join(str(unicodedata.digit(ch)) if ch.isdigit() else ch for ch in s)


def _inline_superscript(line: dict) -> str | None:
    """Promote only visually confirmed small exponent glyphs in mixed text.

    The glyph stream may list a superscript before the base in RTL PDFs.
    """
    spans=line.get('spans',[])
    if len(spans)<2:
        return None
    pairs={}
    consumed=set()
    for i,base in enumerate(spans):
        if len(base['text'])!=1 or not base['text'].isdigit():
            continue
        for j,exponent in enumerate(spans):
            if i==j or j in consumed or i in consumed or j in pairs:
                continue
            if len(exponent['text'])!=1 or not exponent['text'].isdigit():
                continue
            if not (.45 <= exponent['size']/base['size'] <= .84):
                continue
            if abs(exponent['bbox'][0]-base['bbox'][2])>max(2.0,base['size']*.38):
                continue
            if exponent['origin'][1]>base['origin'][1]-1.0:
                continue
            first=min(i,j)
            pairs[first]=(_latin_digits(base['text'])+'^'+_latin_digits(exponent['text']))
            consumed.update((i,j))
            break
    if not pairs:
        return None
    return ''.join(pairs[i] if i in pairs else ('' if i in consumed else sp['text'])
                   for i,sp in enumerate(spans))


def _infer_reversed_math_digits(page: pymupdf.Page) -> bool:
    """Infer backwards stored Arabic digits only from checked powers/results.

    For a PDF storing e.g. ٣٢ for printed 32, require a neighboring
    arithmetic expression that gives 32 before changing prose numbers.
    """
    lines = []
    for block in page.get_text('dict')['blocks']:
        if block.get('type') != 0:
            continue
        for line in block['lines']:
            raw = ''.join(sp['text'] for sp in line.get('spans', []))
            lines.append((line['bbox'][0], line['bbox'][1], raw, _math_line(line) or ''))
    for x,y,raw,converted in lines:
        m = re.search(r'([0-9]+)\^([0-9]+)',converted)
        if not m:
            continue
        base, exponent = map(int,m.groups())
        if base > 99 or exponent > 12:
            continue
        expected=base**exponent
        if expected < 10 or expected > 999999:
            continue
        for x2,y2,raw2,converted2 in lines:
            if not (0 < y2-y < 65 and abs(x2-x) < 200):
                continue
            match=re.search(r'[٠-٩۰-۹]{2,}',raw2)
            if not match or '=' not in raw2:
                continue
            glyphs=match.group()
            forward=int(_latin_digits(glyphs))
            reverse=int(_latin_digits(glyphs[::-1]))
            if reverse==expected and forward != expected:
                return True
    return False


def _normalize_arabic_line(text: str, *, reverse_multidigit: bool = False) -> str:
    # Correct only well-known PDF glyph-order and Unicode mapping artifacts.
    for old,new in (('األ','الأ'), ('إىل','إلى'),('عىل','على'),
                    ('ثانياًلا','ثانيًا'),('ثالثاًلا','ثالثًا'),('أوˀ','أولًا'),
                    ('مالحظة','ملاحظة'), ('أُأسّ','أُسّ'),('عامˁً','عامًا')):
        text=text.replace(old,new)
    if reverse_multidigit:
        text=_ARABIC_MULTI_NUMBER.sub(lambda m:_latin_digits(m.group()[::-1]), text)
    # RTL punctuation regularly arrives on the opposite end of a line.
    if text.lstrip().startswith('.') and re.search(r'[ء-ي]',text):
        text=text.lstrip()[1:].strip()+'.'
    if text.lstrip().startswith(':') and re.search(r'[ء-ي]',text):
        text=text.lstrip()[1:].strip()+':'
    # A combining mark separated from its base by PDF glyph reordering is not a valid prefix.
    text=re.sub(r'(?<=\s)[\u064b-\u0652]+(?=[ء-ي])','',text)
    text=re.sub(r'([ء-ي][ًٌٍ])\s+(?=ا)',r'\1',text)
    text=re.sub(r'\s+([.,:])',r'\1',text)
    text=re.sub(r'^([ء-ي\s]{3,25})\.\s*(?=[ء-يً-ْ])',r'\1: ',text)
    text=re.sub(r'(\^[0-9]+):\s+هكذا',r'\1 هكذا:',text)
    text=re.sub(r'^\)مثال \(([٠-٩0-9]+)$',r'مثال (\1)',text)
    text=re.sub(r'([.:])\s*(?=\S)',r'\1 ',text)
    return text.strip()


def _merge_arithmetic_lines(lines: list[str]) -> list[str]:
    """Join vertically separated equations, never interpolate missing values."""
    out=[]
    for item in lines:
        s=item.strip()
        if _SIMPLE_MATH.fullmatch(s):
            s=re.sub(r'=(?=[0-9])','= ',s)
            item=s
        if s and out and not out[-1] and len(out)>1 and _SIMPLE_MATH.fullmatch(s) and _SIMPLE_MATH.fullmatch(out[-2]):
            if s.startswith('=') and '^' in out[-2]:
                out.pop()
        if s and out and out[-1] and _SIMPLE_MATH.fullmatch(s) and _SIMPLE_MATH.fullmatch(out[-1]):
            last=out[-1].strip()
            if last.count('×') >= 1 and (s.endswith('=') or s.startswith('=')) and '^' in s:
                power=s.rstrip('=').lstrip('=').strip()
                out[-1]=last+' = '+power
                continue
            if '^' in last and s.startswith('=') and last.count('=') == 0:
                out[-1]=last+' '+s
                continue
            if s.startswith('=') and '=' in last and not last.endswith('='):
                out[-1]=last+' '+s
                continue
        out.append(item)
    return out


def _panel_lines_to_markdown(cell: list[tuple], *, reverse_multidigit: bool) -> list[str]:
    """Stable top-to-bottom grouping, respecting compact definition rows."""
    records=sorted(cell,key=lambda z:(z[0],-z[1]))
    groups=[]
    for y,x,txt,size in records:
        if groups and y - groups[-1]['end'] <= 9.5 and y - groups[-1]['start'] < 18:
            # Merge only when headings or brief fragments share a print row.
            items=groups[-1]['items']
            share_y=abs(y - groups[-1]['start']) <= 6.5
            bold_heading=any(z[3] >= size*1.06 and z[1] > x for z in items)
            is_short=len(txt.strip())<=25
            bold_incoming=(is_short and size>=max(z[3] for z in items)*1.04
                           and x>max(z[1] for z in items)+25)
            if share_y or (is_short and bold_heading) or bold_incoming:
                items.append((y,x,txt,size))
                groups[-1]['end']=y
                continue
        groups.append({'start':y,'end':y,'items':[(y,x,txt,size)]})
    result=[]
    prior=None
    for group in groups:
        items=group['items']
        if prior is not None and group['start']-prior>28:
            result.append('')
        # When a bold right-hand label has a wrapped left-hand description,
        # label first, then continuation in vertical order.
        heading=max(items,key=lambda q:(q[1],q[3]))
        definition_row=(len(items)>=2 and heading[3]>=max(8.0,min(q[3] for q in items)*1.06)
                        and heading[1]>max(q[1] for q in items if q is not heading)+18
                        and len(heading[2])<=25)
        if definition_row:
            children=[heading]+sorted((q for q in items if q is not heading),key=lambda q:(q[0],-q[1]))
        else:
            children=sorted(items,key=lambda q:-q[1])
        parts=[z[2] for z in children]
        if definition_row:
            label=parts[0].strip().lstrip('ّ')
            parts[0]=label.rstrip(': .')+': '
            parts[1:]=[s.lstrip(' .') for s in parts[1:]]
        elif parts and parts[0].lstrip().startswith(':') and len(parts)>1:
            parts[0]=parts[0].lstrip().lstrip(':').strip()+':'
        value=' '.join(parts)
        # Single-line PDF table headings sometimes carry the separator as the
        # first character of their right-to-left definition text.
        if len(items)>=2 and not definition_row:
            value=re.sub(r'^([^.:]{3,25})\.\s+(?=[ء-ي])',r'\1: ',value)
        value=_normalize_arabic_line(value,reverse_multidigit=reverse_multidigit)
        result.append(value)
        prior=group['start']
    return _merge_arithmetic_lines(result)


def _six_panel_markdown(page: pymupdf.Page) -> str:
    """Use physical panels and glyph positions rather than plain line order."""
    w,h=page.rect.width,page.rect.height
    cells={i:[] for i in range(6)}
    reverse_math_numbers=_infer_reversed_math_digits(page)
    for block in page.get_text('dict')['blocks']:
        if block.get('type')!=0:
            continue
        xc=(block['bbox'][0]+block['bbox'][2])/2
        yc=(block['bbox'][1]+block['bbox'][3])/2
        c=min(2,max(0,int(xc/(w/3))))
        r=int(yc>=h/2)
        for line in block['lines']:
            text=_math_line(line)
            if text is None:
                text=_inline_superscript(line)
            if text is None:
                text=''.join(sp['text'] for sp in line.get('spans',[])).strip()
            if text:
                y=line['bbox'][1]
                x=line['bbox'][0]
                size=max((sp['size'] for sp in line['spans']),default=9)
                cells[r*3+c].append((y,x,text,size))
    output=[]
    for n,panel_id in enumerate((2,1,0,5,4,3),1):
        output.append(f'## القسم {n}')
        content=_panel_lines_to_markdown(cells[panel_id],reverse_multidigit=reverse_math_numbers)
        # Remove print fold/page labels, not meaningful lesson content.
        if reverse_math_numbers:
            content=[t for t in content if not re.fullmatch(r'[٠-٩0-9]\s*/\s*[٠-٩0-9]',t)]
        if n==1:
            content=[re.sub(r'^(الدرس)\s+([٠-٩0-9]+)[–-]([٠-٩0-9]+)$',
                            lambda m:f'{m.group(1)} {_latin_digits(m.group(3))}–{_latin_digits(m.group(2))}',t)
                     for t in content]
        output.extend(content)
        output.append('')
    return '\n'.join(output)


def _pixmap_image(page) -> Image.Image:
    pix=page.get_pixmap(matrix=pymupdf.Matrix(1.9,1.9),alpha=False)
    return Image.open(io.BytesIO(pix.tobytes('png')))


def _image_ocr(img: Image.Image) -> str:
    try:
        import pytesseract
        img=ImageOps.exif_transpose(img).convert('RGB')
        return pytesseract.image_to_string(img,lang='ara+eng',config='--psm 3').strip()
    except (ImportError, OSError) as ex:
        raise ConversionError('التعرف على النص من الصور غير متاح. ثبّت Tesseract مع اللغتين ara وeng.') from ex
    except Exception as ex:
        raise ConversionError('فشلت قراءة النص من الصورة؛ تحقق من وضوحها.') from ex


def _pdf(path: Path, *,force_ocr: bool) -> tuple[str,list[str]]:
    warnings=[]
    try:
        with pymupdf.open(str(path)) as doc:
            if doc.needs_pass:
                raise ConversionError('الملف محمي بكلمة مرور. أرسل نسخة غير محمية.')
            count=doc.page_count
            if not count:
                raise ConversionError('ملف PDF لا يحتوي على صفحات.')
            if count>MAX_PDF_PAGES:
                raise ConversionError(f'هذا الملف فيه {count} صفحة. الحد الحالي {MAX_PDF_PAGES} صفحة.')
            lengths=[len(page.get_text().strip()) for page in doc]
            scanned_pages=sum(length<10 for length in lengths)
            if scanned_pages:
                warnings.append(f'يوجد {scanned_pages} صفحة دون نص رقمي واضح؛ تعتمد دقة ناتجها على OCR.')
            # Preserve semantics and equations in a 3-column x 2-row printout.
            # Force OCR overrides this because the text layer may be corrupted.
            if count==1 and not force_ocr and _looks_like_six_panel_handout(doc[0]):
                text=_six_panel_markdown(doc[0])
                warnings.append('تم تقسيم المطوية إلى ستة أقسام بحسب مواضعها، وتحليل الأسس من حجم الحروف. راجع الصيغ المعقدة.')
                if re.search(r'(?m)^\s*[٠-٩0-9]\s+[٠-٩0-9]\s*$', text):
                    warnings.append('توجد أرقام منفصلة في رسومات توضيحية؛ راجع موضع الأساس والأس في PDF الأصلي.')
                return text, warnings
        try:
            import pymupdf4llm
        except ImportError:
            pymupdf4llm=None
        if pymupdf4llm is not None:
            text=pymupdf4llm.to_markdown(
                str(path),ocr_language='ara+eng',force_ocr=force_ocr,
                use_ocr=True,write_images=False,embed_images=False,
                page_separators=False,show_progress=False)
        else:
            with pymupdf.open(str(path)) as doc:
                parts=[]
                for page in doc:
                    s=page.get_text('text',sort=True).strip()
                    if force_ocr or not s:
                        s=_image_ocr(_pixmap_image(page))
                    parts.append(s)
                text='\n\n'.join(parts)
            warnings.append('لم يتوفر محرك الجداول المتقدم؛ استخدم المتطلبات كاملة للحصول على أفضل نتيجة.')
        if not text or not text.strip():
            raise ConversionError('تعذّر استخراج النص. جرّب /ocr on ثم أرسل الملف مرة ثانية.')
        if force_ocr:
            warnings.append('استُخدم وضع OCR الإجباري؛ راجع الرموز الرياضية والأرقام بعناية.')
        warnings.append('الرسومات والصور غير النصية لا تتحول تلقائيًا إلى وصف دقيق.')
        return str(text),warnings
    except ConversionError:
        raise
    except Exception as ex:
        raise ConversionError('تعذرت قراءة PDF. تحقق أن الملف صالح، أو جرّب وضع OCR.') from ex


def _image(path: Path) -> tuple[str,list[str]]:
    try:
        with Image.open(path) as img:
            if img.width*img.height > Image.MAX_IMAGE_PIXELS:
                raise ConversionError('أبعاد الصورة كبيرة جدًا؛ أرسل صورة أصغر.')
            value=_image_ocr(img)
    except (Image.DecompressionBombError,UnidentifiedImageError) as ex:
        raise ConversionError('الصورة تالفة أو عالية الدقة بصورة غير آمنة.') from ex
    if not value:
        raise ConversionError('لم يظهر نص يمكن استخراجه من الصورة.')
    return value,['الصور المعالجة بـOCR قد تحتوي على أخطاء في الأرقام والرموز.']


def _docx(path: Path) -> tuple[str,list[str]]:
    assert_safe_office_zip(path)
    doc=Document(str(path))
    parts=[]
    for block in doc.iter_inner_content():
        if isinstance(block,DocxParagraph):
            value=block.text.strip()
            if not value:continue
            style=block.style.name.lower() if block.style else ''
            if style.startswith('heading'):
                found=re.search(r'\d+',style)
                level=max(1,min(6,int(found.group()) if found else 1))
                parts.append('#'*level+' '+value)
            elif any(s in style for s in ('list bullet','list paragraph','قائمة نقطية')):
                parts.append('- '+value)
            elif 'list number' in style or 'قائمة رقمية' in style:
                parts.append('1. '+value)
            else:parts.append(value)
        elif isinstance(block,DocxTable):
            parts.append(safe_table([[c.text for c in row.cells] for row in block.rows]))
    return '\n\n'.join(parts),['الصور والأشكال المضمّنة في Word لا يتحول معناها تلقائيًا إلى نص.']


def _pptx(path: Path) -> tuple[str,list[str]]:
    assert_safe_office_zip(path)
    presentation=Presentation(str(path))
    parts=[]
    for idx,slide in enumerate(presentation.slides,start=1):
        parts.append(f'## شريحة {idx}')
        for shape in sorted(slide.shapes,key=lambda s:(s.top,s.left)):
            if shape.has_text_frame:
                text='\n'.join(p.text for p in shape.text_frame.paragraphs if p.text.strip()).strip()
                if text:parts.append(text)
            if shape.has_table:
                parts.append(safe_table([[c.text for c in r.cells] for r in shape.table.rows]))
    return '\n\n'.join(parts),['الرسوم البيانية والصور في الشرائح لا يمكن فهمها بالكامل دون نموذج رؤية.']


def _xlsx(path: Path) -> tuple[str,list[str]]:
    assert_safe_office_zip(path)
    wb=load_workbook(str(path),read_only=True,data_only=True)
    try:
        contents=[]
        for ws in wb.worksheets:
            contents.append('# '+ws.title)
            if ws.max_column>75 or (ws.max_row or 0)*(ws.max_column or 0)>120_000:
                raise ConversionError('ورقة Excel كبيرة جدًا للتحويل في النسخة الحالية.')
            rows=[]
            for row in ws.iter_rows(values_only=True):
                values=['' if it is None else str(it) for it in row]
                if any(v.strip() for v in values):rows.append(values)
            if rows:
                width=max(i+1 for row in rows for i,v in enumerate(row) if v.strip())
                contents.append(safe_table([r[:width] for r in rows]))
        return '\n\n'.join(contents),['قد لا تتوفر القيم المحسوبة للصيغ إذا لم يحفظ Excel ناتجها داخل الملف.']
    finally:wb.close()


def _text(path: Path) -> tuple[str,list[str]]:
    data=path.read_bytes()
    codecs=('utf-16',) if data.startswith((b'\xff\xfe',b'\xfe\xff')) else ('utf-8-sig','cp1256','cp1252')
    for enc in codecs:
        try:
            decoded=data.decode(enc)
            if '\x00' in decoded and not data.startswith((b'\xff\xfe',b'\xfe\xff')):continue
            return decoded,[]
        except UnicodeError:pass
    raise ConversionError('تعذر قراءة ترميز الملف النصي.')


def _csv(path: Path) -> tuple[str,list[str]]:
    text,_=_text(path)
    try:
        dialect=csv.Sniffer().sniff(text[:4096],delimiters=',;\t')
        rows=list(csv.reader(io.StringIO(text),dialect))
    except csv.Error:
        rows=list(csv.reader(io.StringIO(text)))
    if len(rows)>15000:raise ConversionError('ملف CSV يحتوي على صفوف أكثر من الحد المسموح.')
    return safe_table(rows),[]


def _html(path: Path) -> tuple[str,list[str]]:
    text,_=_text(path)
    soup=BeautifulSoup(text,'html.parser')
    for element in soup(['script','style','noscript']):element.decompose()
    return markdownify(str(soup.body or soup),heading_style='ATX'),[]


def convert_file(path: str|Path,mode: str='faithful',force_ocr: bool=False) -> ConversionResult:
    path=Path(path)
    suffix=path.suffix.lower()
    if suffix not in SUPPORTED:
        raise ConversionError('نوع الملف غير مدعوم. استخدم PDF أو DOCX أو PPTX أو XLSX أو TXT أو CSV أو HTML أو صورة.')
    if mode not in ('faithful','compact'):
        raise ValueError('Unknown conversion mode')
    handlers={
        '.pdf':lambda:_pdf(path,force_ocr=force_ocr),
        '.docx':lambda:_docx(path),'.pptx':lambda:_pptx(path),
        '.xlsx':lambda:_xlsx(path),'.txt':lambda:_text(path),
        '.md':lambda:_text(path),'.csv':lambda:_csv(path),
        '.html':lambda:_html(path),'.htm':lambda:_html(path),
        '.png':lambda:_image(path),'.jpg':lambda:_image(path),
        '.jpeg':lambda:_image(path),'.webp':lambda:_image(path),
    }
    try:
        raw,warnings=handlers[suffix]()
    except ConversionError:raise
    except Exception as ex:
        raise ConversionError('تعذر تحويل هذا الملف. ربما هو تالف أو لا يحتوي على نص قابل للاستخراج.') from ex
    raw=raw.replace('\r\n','\n').replace('\r','\n').strip()+'\n'
    if not raw.strip():raise ConversionError('الملف لا يحتوي على نص قابل للاستخراج.')
    out=clean_compact(raw) if mode=='compact' else raw
    if len(out)>MAX_OUTPUT_CHARS:
        raise ConversionError('الناتج كبير جدًا؛ قسّم الملف إلى أجزاء أصغر.')
    return ConversionResult(markdown=out,raw_markdown=raw,warnings=warnings)


def estimate_tokens(text: str) -> tuple[int,str]:
    try:
        import tiktoken
        enc=tiktoken.get_encoding('cl100k_base')
        return len(enc.encode(text)),'cl100k_base'
    except Exception:
        return max(1,round(len(text)/3)),'تقريب بالحروف'
