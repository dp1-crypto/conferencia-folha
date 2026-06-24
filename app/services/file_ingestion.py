"""
Deteccao de tipo de arquivo e OCR
"""
import os
from typing import Tuple


def detect_file_type(filename: str, raw: bytes) -> str:
    """Detecta tipo por extensao e magic bytes. Retorna: 'pdf'|'excel'|'word'|'csv'|'txt'|'image'|'unknown'"""
    fn = filename.lower()
    if fn.endswith('.pdf'):
        return 'pdf'
    elif fn.endswith(('.xls', '.xlsx')):
        return 'excel'
    elif fn.endswith(('.doc', '.docx')):
        return 'word'
    elif fn.endswith('.csv'):
        return 'csv'
    elif fn.endswith('.txt'):
        return 'txt'
    elif fn.endswith(('.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif')):
        return 'image'
    # Fallback por magic bytes
    if raw[:4] == b'%PDF':
        return 'pdf'
    if raw[:2] in (b'PK', ):  # ZIP = xlsx/docx
        if b'word/' in raw[:2000]:
            return 'word'
        return 'excel'
    return 'unknown'


def ocr_image(raw: bytes) -> str:
    """Tenta OCR em imagem. Retorna texto ou string vazia se pytesseract nao disponivel."""
    try:
        import pytesseract
        from PIL import Image
        import io
        img = Image.open(io.BytesIO(raw))
        return pytesseract.image_to_string(img, lang='por')
    except ImportError:
        return ''
    except Exception:
        return ''


def ocr_pdf_scanned(raw: bytes) -> str:
    """Tenta OCR em PDF escaneado via pdfplumber + Pillow."""
    try:
        import pdfplumber
        import io
        text_parts = []
        with pdfplumber.open(io.BytesIO(raw)) as pdf:
            for page in pdf.pages:
                t = page.extract_text() or ''
                if t.strip():
                    text_parts.append(t)
                else:
                    # Pagina sem texto — tenta renderizar e fazer OCR
                    try:
                        img = page.to_image(resolution=200).original
                        from PIL import Image
                        import pytesseract
                        text_parts.append(pytesseract.image_to_string(img, lang='por'))
                    except Exception:
                        pass
        return '\n'.join(text_parts)
    except Exception:
        return ''
