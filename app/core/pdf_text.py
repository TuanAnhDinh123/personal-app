"""PDF → Text: trích văn bản từ PDF (Word xuất ra hoặc bản scan) + đối chiếu.

Cách hoạt động:
  - Trang nào có sẵn lớp text (PDF từ Word) → trích thẳng, nhanh & chính xác.
  - Trang nào là ảnh scan (không có lớp text) → OCR bằng Tesseract (offline).
Giao diện chia hai cột: bên trái là ảnh trang gốc, bên phải là văn bản
trích được (sửa trực tiếp) để người dùng đối chiếu rồi sao chép / lưu ra .txt.

Phụ thuộc:
  pip install pymupdf pytesseract
  + cài Tesseract-OCR (kèm gói tiếng Việt) — chỉ cần cho PDF scan.
"""
import os
import threading



try:
    import fitz  # PyMuPDF
    _FITZ_OK = True
except ImportError:
    _FITZ_OK = False

try:
    import pytesseract
    from PIL import Image
    _PYTESSERACT_OK = True
except ImportError:
    _PYTESSERACT_OK = False


# Ngưỡng ký tự để coi một trang là "có lớp text". Dưới ngưỡng → coi là scan.
_MIN_TEXT_CHARS = 12
_OCR_DPI_OPTIONS = {"Fast (200 DPI)": 200, "Balanced (300 DPI)": 300, "Sharp (400 DPI)": 400}

# Nơi installer Tesseract hay đặt file, dùng khi tesseract.exe không có trong PATH
# (bản cài qua winget không tự thêm PATH).
_TESSERACT_DIRS = [
    os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"), "Tesseract-OCR"),
    os.path.join(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"), "Tesseract-OCR"),
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "Tesseract-OCR"),
]


# ---------------------------------------------------------------------------
# Core logic (độc lập với UI, dễ test)
# ---------------------------------------------------------------------------

def _locate_tesseract():
    """Trỏ pytesseract tới tesseract.exe: ưu tiên PATH, sau đó các thư mục cài sẵn."""
    if not _PYTESSERACT_OK:
        return False
    try:
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        pass
    for folder in _TESSERACT_DIRS:
        exe = os.path.join(folder, "tesseract.exe")
        if folder and os.path.isfile(exe):
            pytesseract.pytesseract.tesseract_cmd = exe
            try:
                pytesseract.get_tesseract_version()
                return True
            except Exception:
                continue
    return False


def tesseract_available():
    """True nếu pytesseract import được VÀ tìm thấy file thực thi Tesseract."""
    return _locate_tesseract()


def available_languages():
    """Danh sách mã ngôn ngữ Tesseract đã cài (rỗng nếu chưa có Tesseract)."""
    if not tesseract_available():
        return []
    try:
        return sorted(pytesseract.get_languages(config=""))
    except Exception:
        return []


def missing_languages(lang):
    """Các mã trong chuỗi lang ("vie+eng") chưa có gói .traineddata tương ứng."""
    installed = available_languages()
    if not installed:
        return []
    return [c for c in lang.split("+") if c and c not in installed]


def _ocr_page(page, lang, dpi):
    """Render một trang PDF ra ảnh rồi OCR, trả về chuỗi text."""
    zoom = dpi / 72.0
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
    img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    return pytesseract.image_to_string(img, lang=lang).strip()


def extract_pages(pdf_path, ocr_lang="vie+eng", ocr_dpi=300,
                  force_ocr=False, progress=None):
    """Trích text từng trang. Trả về list dict: {text, source}.

    source: "text"  = lấy từ lớp text có sẵn
            "ocr"   = nhận dạng bằng Tesseract
            "scan"  = là trang scan nhưng chưa OCR được (thiếu Tesseract)
    `progress(done, total, source)` được gọi sau mỗi trang (nếu truyền vào).
    """
    can_ocr = tesseract_available()
    if can_ocr:
        lacking = missing_languages(ocr_lang)
        if lacking:
            raise RuntimeError(
                "Tesseract has no language pack for: " + ", ".join(lacking) + ".\n"
                "Installed: " + ", ".join(available_languages()) + ".\n"
                "Add the missing .traineddata file to Tesseract's tessdata folder, "
                "or change the OCR language.")
    pages = []
    with fitz.open(pdf_path) as doc:
        total = doc.page_count
        for i in range(total):
            page = doc[i]
            text = page.get_text("text").strip()
            source = "text"

            if force_ocr or len(text) < _MIN_TEXT_CHARS:
                if can_ocr:
                    text = _ocr_page(page, ocr_lang, ocr_dpi)
                    source = "ocr"
                elif source == "text" and len(text) < _MIN_TEXT_CHARS:
                    source = "scan"  # cần OCR nhưng không có Tesseract

            pages.append({"text": text, "source": source})
            if progress:
                progress(i + 1, total, source)
    return pages


# ---------------------------------------------------------------------------
# Tool (UI)
# ---------------------------------------------------------------------------

_SOURCE_BADGE = {
    "text": ("Text layer", "info"),
    "ocr": ("OCR", "success"),
    "scan": ("Scan – no OCR", "warning"),
}
