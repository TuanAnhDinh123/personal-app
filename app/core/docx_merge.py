"""Điền dữ liệu vào các trường MERGEFIELD của file Word (.docx).

Làm đúng việc mà tính năng **Mail Merge** của Word làm, nhưng chạy thẳng trong
app: mở file mẫu → thay từng trường bằng CHỮ TĨNH → ghi ra file mới. Không cần
cài Word, không cần file Excel trung gian, không phụ thuộc gói ngoài (`zipfile`
+ `xml.etree` đều là thư viện chuẩn).

TRƯỜNG TRONG FILE MẪU LÀ *FIELD* CỦA WORD, KHÔNG PHẢI CHỮ ``<<tên>>``. Word
hiển thị chúng là ``«Full_name»`` nên nhìn trên màn hình tưởng là chữ thường,
nhưng trong XML mỗi trường là một chuỗi run:

    begin ─ instrText " MERGEFIELD Full_name " ─ separate ─ [kết quả] ─ end

nên tìm-thay theo chuỗi ký tự KHÔNG chạm được tới chúng. Ở đây mỗi trường được
thay bằng **một run duy nhất** sao chép định dạng của run kết quả cũ (font, cỡ
chữ, in đậm…), còn các run điều khiển thì bỏ đi — file ra là văn bản thường,
mở bằng Word không còn hỏi cập nhật dữ liệu nữa.

Khóa ngày (``\\@ "dd/MM/yyyy"``) VIẾT TRONG CHÍNH FILE MẪU nên truyền thẳng
``datetime.date`` là được: HR đổi cách hiển thị ngày trong Word là đổi luôn ở
file xuất ra, không phải sửa code.
"""
import copy
import datetime
import os
import re
import shutil
import unicodedata
import zipfile
from typing import NamedTuple
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_R = f"{W}r"
_T = f"{W}t"
_FLDCHAR = f"{W}fldChar"
_INSTRTEXT = f"{W}instrText"
_RPR = f"{W}rPr"
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# Các phần của gói .docx có thể chứa trường: thân bài, đầu/chân trang, chú thích.
# Bảng và text box nằm lồng bên trong chính các phần này nên không phải kể ra.
_TEXT_PARTS = re.compile(
    r"^word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$")

_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

# " MERGEFIELD  Tháng_còn_đóng_BHXH \@ "MM/yyyy" " → tên + khóa định dạng.
# Tên trường là phần đứng trước khóa đầu tiên; chỉ khóa `\@` mới là định dạng
# ngày (`\* MERGEFORMAT` — hay gặp — nói về định dạng CHỮ, không phải ngày).
_MERGEFIELD_RE = re.compile(r"\bMERGEFIELD\s+(.+?)\s*$", re.S)
_SWITCH_RE = re.compile(r"\\[A-Za-z*@#!]\s*(?:\"[^\"]*\"|\S+)")
_DATE_SWITCH_RE = re.compile(r"\\@\s*(?:\"([^\"]*)\"|(\S+))")

# Mã ngày của Word → cách lấy giá trị. Word dùng CHỮ HOA cho tháng (MM) và chữ
# thường cho phút, nên không thể so khớp kiểu không phân biệt hoa/thường.
_DATE_TOKEN_RE = re.compile(r"d{1,2}|M{1,2}|y{4}|y{2}|'[^']*'")
_DATE_TOKENS = {
    "d": lambda v: str(v.day),
    "dd": lambda v: f"{v.day:02d}",
    "M": lambda v: str(v.month),
    "MM": lambda v: f"{v.month:02d}",
    "yy": lambda v: f"{v.year % 100:02d}",
    "yyyy": lambda v: f"{v.year:04d}",
}
_DEFAULT_DATE_FORMAT = "dd/MM/yyyy"


class MergeError(Exception):
    """File mẫu không đọc được (không phải .docx, thiếu word/document.xml…)."""


def norm_name(name) -> str:
    """Chuẩn hóa tên trường để so khớp: NFC · bỏ khoảng trắng thừa · chữ thường.

    NFC là bắt buộc chứ không phải cho đẹp: tên trường trong file mẫu có dấu
    tiếng Việt ("Tháng_còn_đóng_BHXH") và Word ghi ra dạng tổ hợp, còn chuỗi gõ
    trong code là dạng dựng sẵn — hai chuỗi nhìn y hệt nhau mà so bằng `==` thì
    trượt.
    """
    return unicodedata.normalize("NFC", str(name)).strip().lower()


def format_date(value, pattern="") -> str:
    """`date` → chuỗi theo mã định dạng ngày của Word (dd/MM/yyyy, MM/yyyy…).

    Cố tình KHÔNG dùng `strftime`: bỏ số 0 ở đầu (`%-d`) mỗi hệ điều hành viết
    một kiểu (`%#d` trên Windows), mà đây là mã của Word chứ không phải của C.
    Mã lạ (MMMM, dddd…) giữ nguyên chữ — không đoán bừa.
    """
    pattern = (pattern or "").strip() or _DEFAULT_DATE_FORMAT
    out, pos = [], 0
    for m in _DATE_TOKEN_RE.finditer(pattern):
        out.append(pattern[pos:m.start()])
        token = m.group(0)
        if token.startswith("'"):          # 'chữ nguyên văn' trong mã Word
            out.append(token[1:-1])
        else:
            out.append(_DATE_TOKENS[token](value))
        pos = m.end()
    out.append(pattern[pos:])
    return "".join(out)


class _Field(NamedTuple):
    """Một trường MERGEFIELD đã định vị trong danh sách con của một thẻ."""

    name: str
    date_format: str
    begin: int          # chỉ số run `begin`
    sep: int            # chỉ số run `separate` (-1 nếu trường chưa merge lần nào)
    end: int            # chỉ số run `end`


# Thẻ mở của phần tử gốc. `<?xml …?>` và `<!-- … -->` không khớp vì sau "<"
# bắt buộc là chữ cái — nên đây luôn là thẻ mở thật đầu tiên. Thẻ gốc của Word
# dài cỡ 2–3 KB (một đống khai báo namespace); quét dư cho chắc.
_ROOT_TAG_RE = re.compile(r"<([A-Za-z_][\w.:-]*)((?:[^<>\"]|\"[^\"]*\")*)>")
_NS_DECL_RE = re.compile(r'xmlns:([\w.-]+)="([^"]*)"')
_ROOT_SCAN = 32000


def _declared_namespaces(text: str) -> list:
    """Mọi khai báo `xmlns:tiền-tố="uri"` trong phần XML, theo thứ tự gặp.

    Quét CẢ PHẦN chứ không chỉ thẻ gốc: Word khai một số namespace ở phần tử
    CON (hình vẽ DrawingML khai `a` và `a14` ngay tại chỗ dùng). Chỉ đọc thẻ
    gốc thì ET không biết mấy tiền tố đó và tự đặt lại tên thành `ns5`, `ns7` —
    xem `_register_namespaces` để biết vì sao đổi tên là hỏng.
    """
    return _NS_DECL_RE.findall(text)


def _register_namespaces(pairs) -> None:
    """Giữ NGUYÊN tiền tố namespace của file gốc khi ghi lại XML.

    ElementTree tự đặt tên tiền tố (`ns0`, `ns1`…) cho namespace nó chưa biết.
    XML vẫn hợp lệ, nhưng OOXML có mấy thuộc tính mang GIÁ TRỊ LÀ TIỀN TỐ —
    `mc:Ignorable="w14 wp14"`, `<mc:Choice Requires="wps">`, `mc:ProcessContent`
    — chúng không đổi theo. Đổi tên tiền tố là mấy chỗ đó trỏ vào tiền tố không
    còn tồn tại, file sai chuẩn và Word có thể đòi sửa lỗi trước khi mở.
    """
    for prefix, uri in pairs:
        ET.register_namespace(prefix, uri)


def _instr_text(runs) -> str:
    return "".join("".join(t.text or "" for t in r.iter(_INSTRTEXT)) for r in runs)


def _parse_instr(code: str):
    """Mã field → (tên trường, mã định dạng ngày); None nếu không phải MERGEFIELD."""
    m = _MERGEFIELD_RE.search(code)
    if not m:
        return None
    body = m.group(1)
    sw = _DATE_SWITCH_RE.search(body)
    date_format = (sw.group(1) or sw.group(2) or "") if sw else ""
    name = _SWITCH_RE.split(body, maxsplit=1)[0].strip().strip('"')
    return (name, date_format) if name else None


def _read_field(kids, start):
    """Đọc trường bắt đầu ở `kids[start]`; None nếu đó không phải MERGEFIELD.

    Đếm `begin`/`end` theo tầng để field lồng nhau (IF chứa MERGEFIELD) không
    cắt nhầm ở dấu `end` của field con.
    """
    head = kids[start].find(_FLDCHAR) if kids[start].tag == _R else None
    if head is None or head.get(f"{W}fldCharType") != "begin":
        return None
    depth, sep, instr = 0, -1, []
    for i in range(start, len(kids)):
        kid = kids[i]
        if kid.tag != _R:
            continue
        fld = kid.find(_FLDCHAR)
        if fld is None:
            if depth == 1 and sep < 0:
                instr.append(kid)
            continue
        kind = fld.get(f"{W}fldCharType")
        if kind == "begin":
            depth += 1
        elif kind == "separate" and depth == 1:
            sep = i
        elif kind == "end":
            depth -= 1
            if depth == 0:
                parsed = _parse_instr(_instr_text(instr))
                if parsed is None:
                    return None
                return _Field(parsed[0], parsed[1], start, sep, i)
    return None


def _value_text(value, date_format: str) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime.datetime):
        value = value.date()
    if isinstance(value, datetime.date):
        return format_date(value, date_format)
    return str(value)


def _replacement_run(kids, field: _Field, text: str):
    """Một run mang `text`, mượn định dạng của run kết quả cũ.

    Mượn định dạng (chứ không dựng run trắng) để giá trị điền vào giữ đúng
    font/cỡ chữ của chỗ nó đứng — trường ở dòng tiêu đề khác cỡ chữ với trường
    trong thân bài. Trường CHƯA merge lần nào thì không có run kết quả, lấy tạm
    định dạng của run chứa mã field.
    """
    if not text:
        return None
    source = None
    for kid in kids[(field.sep if field.sep >= 0 else field.begin) + 1:field.end]:
        if kid.tag == _R and kid.find(_T) is not None:
            source = kid
            break
    if source is None:
        source = kids[field.begin + 1] if field.begin + 1 < field.end else kids[field.begin]

    run = ET.Element(_R)
    rpr = source.find(_RPR)
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    node = ET.SubElement(run, _T)
    node.text = text
    node.set(_XML_SPACE, "preserve")
    return run


def _merge_element(parent, values: dict, seen: set) -> None:
    """Đệ quy xuống các cấp dưới (bảng, text box) rồi thay mọi MERGEFIELD nằm
    TRỰC TIẾP trong `parent` — trường luôn là chuỗi run ANH EM trong một `<w:p>`."""
    for kid in list(parent):
        _merge_element(kid, values, seen)

    kids = list(parent)
    if not any(k.tag == _R and k.find(_FLDCHAR) is not None for k in kids):
        return

    out, i = [], 0
    while i < len(kids):
        field = _read_field(kids, i) if kids[i].tag == _R else None
        if field is None:
            out.append(kids[i])
            i += 1
            continue
        seen.add(field.name)
        run = _replacement_run(
            kids, field,
            _value_text(values.get(norm_name(field.name)), field.date_format))
        if run is not None:
            out.append(run)
        i = field.end + 1

    parent[:] = out


def _restore_namespaces(new_text: str, original_pairs) -> str:
    """Khai báo lại những namespace mà ElementTree bỏ đi vì "không ai dùng".

    Word khai cả chục namespace của các phiên bản Word khác nhau (w14, w15,
    w16se…) rồi liệt kê chúng trong `mc:Ignorable` — nghĩa là "bản Word nào
    không hiểu mấy tiền tố này thì cứ bỏ qua". Phần lớn không phần tử nào dùng
    tới, nên ET cắt luôn dòng khai báo; `mc:Ignorable` khi đó trỏ vào tiền tố
    KHÔNG TỒN TẠI và file thành sai chuẩn. Chép các khai báo còn thiếu sang.

    ET gom mọi khai báo lên THẺ GỐC, nên bù ở thẻ gốc là đủ — kể cả khai báo mà
    file cũ đặt ở phần tử con: tiền tố có phạm vi cả cây khi khai ở gốc.
    """
    dst = _ROOT_TAG_RE.search(new_text[:_ROOT_SCAN])
    if not dst:
        return new_text
    have = {m.group(1) for m in _NS_DECL_RE.finditer(dst.group(2))}
    missing, added = [], set()
    for prefix, uri in original_pairs:
        if prefix in have or prefix in added:
            continue
        added.add(prefix)
        missing.append(f' xmlns:{prefix}="{uri}"')
    if not missing:
        return new_text
    return new_text[:dst.start(2)] + "".join(missing) + new_text[dst.start(2):]


def _merge_part(raw: bytes, values: dict, seen: set) -> bytes:
    text = raw.decode("utf-8")
    namespaces = _declared_namespaces(text)
    _register_namespaces(namespaces)
    root = ET.fromstring(text)
    _merge_element(root, values, seen)
    merged = _restore_namespaces(ET.tostring(root, encoding="unicode"), namespaces)
    return (_DECLARATION + merged).encode("utf-8")


# ─────────────────────────── gỡ liên kết nguồn dữ liệu ───────────────────────
# File mẫu đang trỏ tới file Excel của HR bằng một câu lệnh OLEDB. Giữ lại thì
# mỗi lần mở file xuất ra Word đều hỏi "chạy câu lệnh SQL?" và đòi đúng đường
# dẫn cũ — trong khi bản xuất ra đã là văn bản tĩnh, không còn gì để merge.
_MAILMERGE_RE = re.compile(r"<w:mailMerge\b.*?</w:mailMerge>|<w:mailMerge\b[^>]*/>", re.S)
_MERGESOURCE_REL_RE = re.compile(r"<Relationship\b[^>]*mailMergeSource[^>]*/>", re.S)


def _strip_data_source(name: str, raw: bytes) -> bytes:
    if name == "word/settings.xml":
        return _MAILMERGE_RE.sub("", raw.decode("utf-8")).encode("utf-8")
    if name == "word/_rels/settings.xml.rels":
        return _MERGESOURCE_REL_RE.sub("", raw.decode("utf-8")).encode("utf-8")
    return raw


# ───────────────────────────────── API dùng ngoài ────────────────────────────
def field_names(template_path) -> list[str]:
    """Tên các trường MERGEFIELD có trong file mẫu (không trùng, giữ thứ tự gặp)."""
    names, seen = [], set()
    try:
        with zipfile.ZipFile(template_path) as zf:
            parts = [n for n in zf.namelist() if _TEXT_PARTS.match(n)]
            for name in parts:
                for code in re.findall(r"<w:instrText[^>]*>(.*?)</w:instrText>",
                                       zf.read(name).decode("utf-8"), re.S):
                    parsed = _parse_instr(code)
                    if parsed and norm_name(parsed[0]) not in seen:
                        seen.add(norm_name(parsed[0]))
                        names.append(parsed[0])
    except (OSError, ValueError, zipfile.BadZipFile) as exc:
        raise MergeError(f"Can't read the Word template: {exc}") from exc
    return names


def render(template_path, out_path, values: dict) -> set:
    """Ghi `out_path` = file mẫu đã điền `values`. Trả về tên các trường CÓ trong
    mẫu mà `values` không có (giá trị của chúng bị bỏ trắng).

    `values` khớp theo tên trường, không phân biệt hoa/thường (xem `norm_name`).
    Giá trị nhận `str`, `datetime.date`/`datetime` (định dạng theo khóa `\\@` của
    chính trường đó) hoặc None/"" — None và "" đều xóa trắng chỗ đó.

    Ghi ra FILE TẠM rồi mới đổi tên: xuất hàng loạt mà lỗi giữa chừng thì không
    để lại file .docx cụt trong thư mục kết quả.
    """
    lookup = {norm_name(k): v for k, v in values.items()}
    seen: set = set()
    tmp_path = f"{out_path}.part"
    try:
        with zipfile.ZipFile(template_path) as src:
            if "word/document.xml" not in src.namelist():
                raise MergeError(f"Not a Word document: {template_path}")
            with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as dst:
                for item in src.infolist():
                    raw = src.read(item.filename)
                    if _TEXT_PARTS.match(item.filename):
                        raw = _merge_part(raw, lookup, seen)
                    else:
                        raw = _strip_data_source(item.filename, raw)
                    dst.writestr(item, raw)
        shutil.move(tmp_path, out_path)
    except (OSError, ValueError, zipfile.BadZipFile, ET.ParseError) as exc:
        _discard(tmp_path)
        raise MergeError(f"Can't build the Word file: {exc}") from exc
    except BaseException:
        _discard(tmp_path)
        raise
    return {name for name in seen if norm_name(name) not in lookup}


def _discard(path) -> None:
    """Dọn file tạm sau khi hỏng giữa chừng; xóa không được cũng kệ."""
    try:
        os.remove(path)
    except OSError:
        pass
