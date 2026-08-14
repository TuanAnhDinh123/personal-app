"""PDF → Markdown: trích văn bản từ PDF, giữ tương đối bố cục gốc để tiện copy.

Hai nhánh xử lý, tự chọn theo từng trang:
  - Trang có sẵn lớp text (PDF xuất từ Word): dựng lại Markdown từ toạ độ chữ —
    tiêu đề theo cỡ chữ, danh sách, bảng (`find_tables`), tự tách 2 cột và bỏ
    header/footer lặp lại. Nhanh, chính xác, chạy offline.
  - Trang là ảnh scan: render ra PNG rồi nhờ Gemini đọc thành Markdown (dùng
    chung API key ở màn hình Settings).

Chỉ phụ thuộc `pymupdf` (wheel thuần, `pip install` được, KHÔNG cần quyền admin
và không cần cài binary ngoài) + `urllib` chuẩn để gọi API. Cố ý không dùng
Tesseract: nó đòi cài chương trình ngoài kèm gói ngôn ngữ, máy không có quyền
admin sẽ tắc.

Mọi thao tác MuPDF trong module này đều đi qua `_LOCK`: MuPDF không thread-safe,
gọi song song từ luồng nền và luồng giao diện sẽ làm app tắt ngang không log.
"""
import base64
import json
import re
import threading
import urllib.error
import urllib.request
from collections import Counter

try:
    import pymupdf as fitz          # tên gói từ 1.24 trở đi
except ImportError:                 # pragma: no cover - bản cũ chỉ có tên `fitz`
    import fitz

from app.core import debuglog

# ---------------------------------------------------------------------------
# Hằng số
# ---------------------------------------------------------------------------

#: Dưới ngưỡng ký tự này thì coi trang là ảnh scan (không có lớp text dùng được).
MIN_TEXT_CHARS = 12

#: Độ phân giải render trang gửi đi OCR. Cao hơn = chữ rõ hơn nhưng ảnh nặng hơn.
OCR_DPI_OPTIONS = {"Fast (150 DPI)": 150, "Balanced (200 DPI)": 200, "Sharp (300 DPI)": 300}

#: Chế độ xử lý người dùng chọn.
MODE_AUTO = "auto"      # trang nào có lớp text thì lấy thẳng, trang scan mới OCR
MODE_TEXT = "text"      # chỉ lấy lớp text, không gọi AI
MODE_OCR = "ocr"        # ép OCR mọi trang (dùng khi lớp text bị lỗi font)

#: Nguồn của text mỗi trang — UI hiển thị thành badge.
SRC_TEXT = "text"
SRC_OCR = "ocr"
SRC_SCAN = "scan"       # là trang scan nhưng chưa OCR được (thiếu API key)
SRC_EMPTY = "empty"     # trang trắng

# Cỡ ảnh tối đa gửi lên API (~4 triệu điểm ảnh) — chặn trang khổ lớn làm phình
# request rồi bị API từ chối.
_MAX_OCR_PIXELS = 4_000_000

_API_URL = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
_RETRY_STATUS = {429, 500, 502, 503, 504}
_MAX_RETRIES = 3

_OCR_PROMPT = (
    "You are an OCR engine. Transcribe every piece of text in this page image "
    "into Markdown.\n"
    "Rules:\n"
    "- Keep the original reading order, top to bottom. If the page has columns, "
    "finish the left column before starting the right one.\n"
    "- Keep the original language exactly as written. Vietnamese text must stay "
    "Vietnamese with correct diacritics. Never translate.\n"
    "- Use Markdown headings for headings, '-' for bullet lists, and Markdown "
    "pipe tables for tables.\n"
    "- Transcribe only. Do not summarise, explain, comment, or add anything that "
    "is not printed on the page.\n"
    "- If the page has no readable text, return an empty response.\n"
    "Return the Markdown only, with no code fence around it."
)

# MuPDF không thread-safe → mọi lời gọi phải nằm trong lock này.
_LOCK = threading.RLock()

# Bit cờ kiểu chữ trong span của get_text("dict").
_FLAG_ITALIC = 2
_FLAG_BOLD = 16

_BULLET_CHARS = "•‣▪●◦·–—-*"
_BULLET_RE = re.compile(r"^\s*([" + re.escape(_BULLET_CHARS) + r"]|o)\s+")
_NUMBERED_RE = re.compile(r"^\s*(\d+|[a-zA-Z]|[ivxIVX]+)\s*[.)]\s+")


class OcrUnavailable(RuntimeError):
    """Cần OCR nhưng chưa cấu hình được (thiếu API key / model)."""


class Cancelled(RuntimeError):
    """Người dùng bấm Hủy giữa chừng."""


class _Transient(RuntimeError):
    """Lỗi API tạm thời (quá tải, giới hạn nhịp, rớt mạng) — nên thử lại."""


# ---------------------------------------------------------------------------
# Mở tài liệu
# ---------------------------------------------------------------------------

def open_document(path):
    """Mở PDF (đã giữ lock). Người gọi tự đóng — dùng làm context manager."""
    with _LOCK:
        return fitz.open(path)


def page_count(path):
    with _LOCK:
        with fitz.open(path) as doc:
            return doc.page_count


def render_png(path, index, max_width=1400):
    """Render một trang ra PNG (bytes) để xem đối chiếu.

    Trả bytes chứ không trả pixmap/QImage: người gọi ở luồng giao diện không
    chạm vào MuPDF nữa, tránh dùng chéo luồng.
    """
    with _LOCK:
        with fitz.open(path) as doc:
            if not 0 <= index < doc.page_count:
                return None
            page = doc[index]
            zoom = max(0.2, min(3.0, max_width / max(1.0, page.rect.width)))
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            return pix.tobytes("png")


# ---------------------------------------------------------------------------
# Dựng lại Markdown từ lớp text
# ---------------------------------------------------------------------------

def _page_lines(page):
    """Mọi dòng chữ của trang: {bbox, text, size}.

    Đơn vị xử lý là DÒNG chứ không phải block: MuPDF hay gộp chữ hai cột nằm
    ngang hàng vào cùng một block, xếp theo block thì hai cột bị trộn vào nhau.
    """
    out = []
    for block in page.get_text("dict").get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            text = _line_text(line)
            if text:
                out.append({"bbox": line["bbox"], "text": text,
                            "size": _line_size(line)})
    return out


def _body_size(pages_lines):
    """Cỡ chữ thân bài = cỡ chiếm nhiều ký tự nhất trong tài liệu.

    Tính theo số ký tự chứ không theo số dòng: tiêu đề ít chữ nhưng cỡ lớn, đếm
    theo dòng thì tài liệu nhiều tiêu đề sẽ chọn nhầm cỡ tiêu đề làm thân bài.
    """
    tally = Counter()
    for lines in pages_lines:
        for line in lines:
            tally[round(line["size"], 1)] += len(line["text"])
    if not tally:
        return 11.0
    return tally.most_common(1)[0][0] or 11.0


def _order_lines(lines, page_rect):
    """Sắp xếp dòng theo THỨ TỰ ĐỌC → trả về danh sách các "mạch" đọc liên tiếp.

    Thứ tự trong content stream của PDF thường lộn xộn — đó chính là lý do text
    trích ra hay bị đảo. Mặc định cả trang là một mạch xếp theo hàng rồi
    trái→phải; trang chia 2 cột thì tách thành các mạch riêng (đầu trang, cột
    trái, cột phải) để đọc hết cột này mới sang cột kia.

    Trả về nhiều mạch thay vì một danh sách phẳng vì mỗi cột có mép phải riêng —
    xem `_right_edge`.
    """
    if not lines:
        return []

    by_row = lambda l: (round(l["bbox"][1] / 3), l["bbox"][0])
    by_top = lambda l: (l["bbox"][1], l["bbox"][0])

    mid = (page_rect.x0 + page_rect.x1) / 2
    wide = [l for l in lines if (l["bbox"][2] - l["bbox"][0]) >= page_rect.width * 0.6]
    narrow = [l for l in lines if l not in wide]
    left = [l for l in narrow if l["bbox"][2] <= mid]
    right = [l for l in narrow if l["bbox"][0] >= mid]

    # Chỉ coi là 2 cột khi không dòng hẹp nào vắt qua đường giữa và cả hai bên
    # đều có đủ nội dung — nếu không, một trang 1 cột có chú thích lệch phải
    # cũng bị chẻ đôi.
    if len(left) + len(right) != len(narrow) or len(left) < 3 or len(right) < 3:
        return [sorted(lines, key=by_row)]

    col_top = min(l["bbox"][1] for l in left + right)
    head = [l for l in wide if l["bbox"][3] <= col_top]
    rest = [l for l in wide if l not in head]
    runs = [sorted(head, key=by_top), sorted(left, key=by_top),
            sorted(right, key=by_top), sorted(rest, key=by_top)]
    return [r for r in runs if r]


def _right_edge(lines):
    """Mép phải của khối chữ — mốc để biết một dòng có bị ngắt do hết chỗ không.

    Lấy phân vị 90 chứ không lấy max: một dòng lẻ thò ra ngoài lề (số trang, chú
    thích) sẽ kéo mốc lệch hẳn, làm mọi dòng còn lại đều bị coi là ngắt cứng.
    """
    edges = sorted(l["bbox"][2] for l in lines)
    return edges[min(len(edges) - 1, int(len(edges) * 0.9))] if edges else 0.0


def _line_text(line):
    """Ghép các span của một dòng, đánh dấu **đậm** / *nghiêng* nếu có."""
    parts = []
    for span in line.get("spans", []):
        text = span.get("text", "")
        if not text.strip():
            parts.append(text)
            continue
        flags = span.get("flags", 0)
        stripped = text.strip()
        lead = text[:len(text) - len(text.lstrip())]
        tail = text[len(text.rstrip()):]
        if flags & _FLAG_BOLD:
            stripped = f"**{stripped}**"
        elif flags & _FLAG_ITALIC:
            stripped = f"*{stripped}*"
        parts.append(lead + stripped + tail)
    # Gộp hai lần đánh dấu liền nhau ("**a** **b**" → "**a b**") cho gọn mắt.
    text = "".join(parts)
    text = re.sub(r"\*\*(\s*)\*\*", r"\1", text)
    text = re.sub(r"(?<!\*)\*(\s*)\*(?!\*)", r"\1", text)
    return text.strip()


def _line_size(line):
    """Cỡ chữ lớn nhất trong dòng — dùng để đoán cấp tiêu đề."""
    return max((s.get("size", 0) for s in line.get("spans", [])), default=0)


def _heading_prefix(size, body, text=""):
    """Trả về '# ' / '## ' / '### ' nếu dòng đủ lớn để coi là tiêu đề.

    Riêng cỡ chữ là chưa đủ: tài liệu trộn nhiều cỡ chữ sẽ biến cả câu văn bình
    thường thành tiêu đề. Câu dài kết thúc bằng dấu câu thì luôn là văn xuôi.
    """
    if len(text) > 120:
        return ""
    if len(text) > 60 and text.rstrip().endswith((".", ",", ";")):
        return ""
    if size >= body * 1.6:
        return "# "
    if size >= body * 1.3:
        return "## "
    if size >= body * 1.12:
        return "### "
    return ""


def _classify(line, body, right_edge):
    """Nhận dạng vai trò của một dòng → ("heading"|"bullet"|"numbered"|"body", md).

    Tiêu đề phải vừa lớn hơn thân bài vừa KHÔNG chạm mép phải của khối: chỉ xét
    cỡ chữ thì một trang đặt nguyên khối văn xuôi ở cỡ lớn (trang tuyên bố miễn
    trừ, lời nói đầu) sẽ biến thành mấy chục dòng tiêu đề liền nhau.
    """
    text = line["text"]
    reaches_margin = line["bbox"][2] >= right_edge - 12
    heading = "" if reaches_margin else _heading_prefix(line["size"], body, text)
    if heading:
        # Tiêu đề đánh số ("1. Tổng quan") vẫn là tiêu đề: xét cỡ chữ TRƯỚC mẫu
        # danh sách, nếu không mọi mục đánh số đều tụt xuống thành gạch đầu dòng.
        return "heading", heading + text.strip("*# ")
    if _BULLET_RE.match(text):
        return "bullet", "- " + _BULLET_RE.sub("", text, count=1).strip()
    if _NUMBERED_RE.match(text):
        return "numbered", text
    return "body", text


def _paragraphs(lines, body, right_edge=None):
    """Gom các dòng đã sắp thứ tự thành đoạn Markdown.

    PDF ngắt dòng theo bề rộng trang chứ không theo câu, nên các dòng liền nhau
    của cùng một đoạn phải nối lại bằng dấu cách — giữ nguyên xuống dòng thì text
    copy ra bị răng cưa. Cắt đoạn khi gặp tiêu đề, gạch đầu dòng, khoảng hở dọc
    rộng bất thường, lề trái đổi, hoặc dòng trước KẾT THÚC SỚM so với mép phải
    của khối: dòng chưa chạm mép phải là do người viết chủ động xuống dòng, nối
    tiếp vào sẽ dính các mục rời rạc (điều khoản, địa chỉ) thành một đoạn dài.
    """
    if right_edge is None:
        right_edge = _right_edge(lines)
    out = []
    buffer = []
    prev = None

    def flush():
        if buffer:
            out.append(" ".join(buffer))
            buffer.clear()

    for line in lines:
        kind, md = _classify(line, body, right_edge)
        if kind == "heading" and buffer and prev is not None:
            # Dòng CUỐI của một đoạn đang viết dở bao giờ cũng ngắn, nên nhìn
            # riêng nó thì giống hệt tiêu đề. Dòng trước chạm mép phải và sát
            # ngay trên → đây là phần đuôi của đoạn, không phải tiêu đề mới.
            gap = line["bbox"][1] - prev["bbox"][3]
            height = max(1.0, prev["bbox"][3] - prev["bbox"][1])
            if gap <= height * 0.7 and prev["bbox"][2] >= right_edge - 12:
                kind, md = "body", line["text"]

        if kind != "body":
            flush()
            out.append(md)
            prev = line
            continue

        if prev is not None:
            gap = line["bbox"][1] - prev["bbox"][3]
            height = max(1.0, prev["bbox"][3] - prev["bbox"][1])
            indent = abs(line["bbox"][0] - prev["bbox"][0])
            short = prev["bbox"][2] < right_edge - 12
            if gap > height * 0.7 or indent > 20 or short:
                flush()
        buffer.append(md)
        prev = line
    flush()
    return [p for p in out if p.strip()]


def _tables_markdown(page):
    """Bảng trên trang: [(bbox, markdown)] — dùng bộ dò bảng sẵn của PyMuPDF."""
    try:
        found = page.find_tables()
    except Exception as exc:                # pragma: no cover - PDF dị dạng
        debuglog.exception("find_tables lỗi", exc)
        return []
    out = []
    for table in getattr(found, "tables", []):
        try:
            md = table.to_markdown().strip()
        except Exception:
            continue
        if md:
            out.append((fitz.Rect(table.bbox), md))
    return out


def _inside_any(bbox, rects):
    """True nếu tâm bbox nằm trong một trong các vùng (bảng) đã cho."""
    cx = (bbox[0] + bbox[2]) / 2
    cy = (bbox[1] + bbox[3]) / 2
    return any(r.x0 <= cx <= r.x1 and r.y0 <= cy <= r.y1 for r in rects)


def _page_markdown(page, lines, body, skip_lines=()):
    """Ghép cả trang: bảng chèn đúng vị trí dọc của nó giữa các đoạn văn."""
    tables = _tables_markdown(page)
    table_rects = [rect for rect, _ in tables]

    kept = [l for l in lines
            if not (table_rects and _inside_any(l["bbox"], table_rects))
            and _norm(l["text"]) not in skip_lines]

    # Bảng đã có Markdown riêng; chèn nó vào đúng vị trí dọc giữa các đoạn văn,
    # nếu không bảng sẽ bị dồn hết xuống cuối trang.
    bounds = sorted((rect.y0, rect.y1, md) for rect, md in tables)
    pieces = []
    index = 0

    for run in _order_lines(kept, page.rect):
        right = _right_edge(run)
        segment = []
        for line in run:
            while index < len(bounds) and line["bbox"][1] >= bounds[index][1]:
                pieces.extend((p, False) for p in _paragraphs(segment, body, right))
                segment = []
                pieces.append((bounds[index][2], True))
                index += 1
            segment.append(line)
        pieces.extend((p, False) for p in _paragraphs(segment, body, right))

    for _, _, md in bounds[index:]:
        pieces.append((md, True))

    return "\n\n".join(_merge_wrapped(pieces)).strip()


_SENTENCE_END = ".!?:;”\"')"


def _merge_wrapped(paragraphs):
    """Nối lại những đoạn bị PDF cắt rời giữa câu.

    Word xuất PDF hay tách một đoạn văn thành nhiều block khi có ngắt dòng mềm;
    để nguyên thì bản Markdown đầy đoạn cụt lủn, dán sang Word phải sửa tay.
    """
    out = []
    for text, is_table in paragraphs:
        text = text.strip()
        if not text:
            continue
        if out and not is_table and _continues(out[-1], text):
            out[-1] = out[-1] + " " + text
        else:
            out.append(text)
    return out


def _continues(prev, nxt):
    """True nếu `nxt` là phần đuôi bị cắt rời của đoạn `prev`."""
    if prev.startswith(("#", "-", "|", ">")) or nxt.startswith(("#", "-", "|", ">")):
        return False
    if _NUMBERED_RE.match(nxt) or prev.rstrip().endswith("|"):
        return False
    tail = prev.rstrip().rstrip("*")
    if not tail or tail[-1] in _SENTENCE_END:
        return False
    head = nxt.lstrip("*")
    return bool(head) and (head[0].islower() or tail.endswith(","))


def _norm(text):
    """Chuẩn hoá một dòng để so khớp header/footer (bỏ số trang, khoảng trắng)."""
    text = re.sub(r"[*#>|`]", "", text)
    text = re.sub(r"\d+", "#", text)
    return re.sub(r"\s+", " ", text).strip().lower()


def _repeated_edge_lines(doc, pages_lines, body):
    """Các dòng header/footer lặp lại gần hết các trang → bỏ khi ghép Markdown.

    PDF từ Word gắn tên công ty / số trang vào mọi trang; giữ lại thì bản text
    cứ vài chục dòng lại chen một dòng rác, copy sang chỗ khác rất khó dùng.

    Ba điều kiện cùng lúc mới loại: nằm trong lề trên/dưới, cỡ chữ KHÔNG lớn hơn
    thân bài, và lặp ở đa số trang. Thiếu điều kiện cỡ chữ thì tiêu đề mục đánh
    số ("2. Tổng quan", "3. Kết luận") ở đầu trang cũng bị coi là header vì sau
    khi chuẩn hoá số chúng trùng nhau.
    """
    total = len(pages_lines)
    if total < 3:
        return set()
    tally = Counter()
    for index, lines in enumerate(pages_lines):
        rect = doc[index].rect
        top = rect.y0 + rect.height * 0.08
        bottom = rect.y1 - rect.height * 0.10
        for line in lines:
            y0, y1 = line["bbox"][1], line["bbox"][3]
            if y1 > top and y0 < bottom:
                continue                    # nằm giữa trang → là nội dung thật
            if line["size"] > body:
                continue                    # chữ to hơn thân bài → là tiêu đề
            if 0 < len(line["text"]) <= 120:
                tally[_norm(line["text"])] += 1
    threshold = max(2, int(total * 0.6))
    return {key for key, count in tally.items() if count >= threshold and key}


# ---------------------------------------------------------------------------
# OCR qua Gemini
# ---------------------------------------------------------------------------

def _ocr_png(png_bytes, api_key, model, timeout=180, should_cancel=None):
    """Gửi ảnh một trang cho Gemini, nhận lại Markdown. Tự thử lại khi 429/5xx."""
    last_error = ""
    for attempt in range(1, _MAX_RETRIES + 1):
        if should_cancel and should_cancel():
            raise Cancelled()
        try:
            return _ocr_png_once(png_bytes, api_key, model, timeout)
        except _Transient as exc:
            last_error = str(exc)
            if attempt == _MAX_RETRIES:
                break
            _sleep(4 * (2 ** (attempt - 1)), should_cancel)
    raise RuntimeError(f"OCR failed after {_MAX_RETRIES} attempts — {last_error}")


def _sleep(seconds, should_cancel=None):
    """Chờ nhưng vẫn bấm Hủy được: cứ 0.2s lại kiểm tra một lần."""
    remaining = seconds
    while remaining > 0:
        if should_cancel and should_cancel():
            raise Cancelled()
        step = min(0.2, remaining)
        threading.Event().wait(step)
        remaining -= step


def _ocr_png_once(png_bytes, api_key, model, timeout):
    body = {
        "contents": [{
            "parts": [
                {"text": _OCR_PROMPT},
                {"inline_data": {
                    "mime_type": "image/png",
                    "data": base64.b64encode(png_bytes).decode("ascii"),
                }},
            ],
        }],
        # Nhiệt độ 0: đây là việc chép lại đúng nguyên văn, không cần sáng tạo.
        "generationConfig": {"temperature": 0.0},
    }
    req = urllib.request.Request(
        _API_URL.format(model=model),
        data=json.dumps(body).encode("utf-8"),
        headers={"Content-Type": "application/json", "x-goog-api-key": api_key},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", "ignore")
        try:
            detail = json.loads(detail).get("error", {}).get("message", detail)
        except ValueError:
            pass
        if exc.code in _RETRY_STATUS:
            raise _Transient(f"HTTP {exc.code}: {detail}") from exc
        raise RuntimeError(f"HTTP {exc.code}: {detail}") from exc
    except urllib.error.URLError as exc:
        raise _Transient(f"Network error: {exc.reason}") from exc

    try:
        text = payload["candidates"][0]["content"]["parts"][0]["text"]
    except (KeyError, IndexError):
        reason = payload.get("promptFeedback", {}).get("blockReason")
        if reason:
            raise RuntimeError(f"The model returned nothing (blocked: {reason})")
        return ""       # trang trắng → model trả rỗng, không phải lỗi
    return _strip_fence(text).strip()


def _strip_fence(text):
    """Bỏ ```markdown ... ``` nếu model vẫn bọc code fence quanh kết quả."""
    stripped = text.strip()
    if stripped.startswith("```"):
        stripped = re.sub(r"^```[a-zA-Z]*\s*\n?", "", stripped)
        stripped = re.sub(r"\n?```\s*$", "", stripped)
    return stripped


# ---------------------------------------------------------------------------
# Điểm vào chính
# ---------------------------------------------------------------------------

def extract(path, mode=MODE_AUTO, ocr_dpi=200, api_key="", model="",
            progress=None, should_cancel=None):
    """Trích cả tài liệu. Trả về list dict {"markdown", "source"} theo từng trang.

    `progress(done, total, source)` gọi sau mỗi trang; `should_cancel()` trả True
    thì dừng ngay và ném Cancelled.

    Toàn bộ phần đọc PDF chạy trong lock — hàm này phải được gọi ở luồng nền,
    còn luồng giao diện chỉ chạm vào PDF qua render_png().
    """
    can_ocr = bool(api_key and model)
    if mode == MODE_OCR and not can_ocr:
        raise OcrUnavailable(
            "OCR needs a Gemini API key. Open Settings and add one, "
            'or switch the mode to "Text layer only".')

    results = []
    with _LOCK:
        with fitz.open(path) as doc:
            total = doc.page_count
            # Đọc trước toàn bộ dòng để tính cỡ chữ thân bài và dò header/footer
            # trên PHẠM VI CẢ TÀI LIỆU — xét từng trang riêng thì trang chỉ có
            # tiêu đề sẽ suy ra cỡ thân bài sai.
            pages_lines = [_page_lines(doc[i]) for i in range(total)]
            body = _body_size(pages_lines)
            skip_lines = _repeated_edge_lines(doc, pages_lines, body)

            for index in range(total):
                if should_cancel and should_cancel():
                    raise Cancelled()
                page = doc[index]
                lines = pages_lines[index]
                raw_len = sum(len(l["text"]) for l in lines)

                need_ocr = mode == MODE_OCR or (
                    mode == MODE_AUTO and raw_len < MIN_TEXT_CHARS)

                if not need_ocr:
                    markdown = _page_markdown(page, lines, body, skip_lines)
                    source = SRC_TEXT if markdown else SRC_EMPTY
                elif not can_ocr:
                    markdown = ""
                    source = SRC_SCAN
                else:
                    png = _page_png(page, ocr_dpi)
                    markdown = _ocr_png(png, api_key, model,
                                        should_cancel=should_cancel)
                    source = SRC_OCR if markdown else SRC_EMPTY

                results.append({"markdown": markdown, "source": source})
                if progress:
                    progress(index + 1, total, source)
    return results


def _page_png(page, dpi):
    """Render trang ra PNG để gửi OCR, tự hạ DPI nếu ảnh quá lớn."""
    zoom = dpi / 72.0
    width = page.rect.width * zoom
    height = page.rect.height * zoom
    if width * height > _MAX_OCR_PIXELS:
        zoom *= (_MAX_OCR_PIXELS / (width * height)) ** 0.5
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
    return pix.tobytes("png")


# ---------------------------------------------------------------------------
# Xuất kết quả
# ---------------------------------------------------------------------------

def join_pages(pages, separators=True):
    """Ghép Markdown của mọi trang thành một tài liệu."""
    parts = []
    for index, page in enumerate(pages):
        text = (page.get("markdown") or "").strip()
        if separators:
            parts.append(f"<!-- Page {index + 1} -->\n\n{text}" if text
                         else f"<!-- Page {index + 1} — no text -->")
        elif text:
            parts.append(text)
    return "\n\n".join(parts).strip() + "\n"


def to_plain_text(markdown):
    """Bỏ ký hiệu Markdown → text trần, cho người muốn dán vào ô chỉ nhận text."""
    lines = []
    for line in markdown.splitlines():
        if re.match(r"^\s*\|[\s:|-]+\|\s*$", line):
            continue                                # dòng kẻ ngăn của bảng
        line = re.sub(r"^<!--.*?-->\s*$", "", line)
        line = re.sub(r"^\s{0,3}#{1,6}\s*", "", line)
        line = re.sub(r"\*\*(.+?)\*\*", r"\1", line)
        line = re.sub(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)", r"\1", line)
        if line.strip().startswith("|"):
            line = "\t".join(c.strip() for c in line.strip().strip("|").split("|"))
        lines.append(line.rstrip())
    return re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip() + "\n"
