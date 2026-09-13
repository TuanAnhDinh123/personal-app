"""SSOT — BỐ CỤC FILE EXCEL "Personnel Data" của HR.

MỌI hiểu biết của app về file Excel nằm trong danh sách `COLUMNS` bên dưới,
xếp **ĐÚNG THỨ TỰ CỘT của file** (A → CD). Cấu trúc file đổi thì sửa **đúng
một chỗ này** — bốn thứ sau tự đi theo:

| Hàm | Dùng ở đâu |
|---|---|
| `header_map()`      | nhận diện tiêu đề khi Bulk Import / Sync (khớp theo TÊN) |
| `ignored_headers()` | tiêu đề cố tình bỏ qua (không báo "cột lạ") |
| `copy_layout()`     | 82 ô của *Copy row*, dán theo **VỊ TRÍ** cột A → CD |
| `table_columns()`   | thứ tự + nhãn cột trên bảng nhân viên |

Thêm một cột vào file = thêm **một dòng** vào đúng vị trí của nó trong
`COLUMNS`; các ô phía sau của *Copy row* tự dời theo, không phải đếm tay.

**Hai kiểu khớp, đừng lẫn.** Import/Sync khớp theo TÊN cột nên đổi chỗ cột
trong file không ảnh hưởng; còn *Copy row* dán mù ra clipboard nên bám VỊ TRÍ.
Vì vậy `COLUMNS` phải giữ đúng thứ tự file, và `layout_mismatches()` đối chiếu
lại với dòng tiêu đề thật mỗi lần app đọc file để lệch là biết ngay — đây chính
là kiểu hỏng mà bố cục theo vị trí không tự lộ ra.
"""
import re
import unicodedata
from typing import NamedTuple

# Sentinel: cột trong Excel là TEXT, phải tra ra id ở bảng danh mục trước khi
# ghi xuống DB (xem employee_db._MASTER_LOOKUPS).
DEPT_TEXT = "__department_short_name__"    # tra theo departments.short_name
LEVEL_TEXT = "__level_name__"              # tra theo levels.level_name
CC_TEXT = "__cost_center_code__"           # tra theo cost_centers.code
ETYPE_TEXT = "__employee_type_code__"      # tra theo employee_types.code

# Ô Copy row mặc định: dán GIÁ TRỊ của chính cột đó (xem Col.copy).
VALUE = object()

# Ký tự vùng Private Use Area (U+E000–U+F8FF) — font ký hiệu (Wingdings…) lẫn
# vào tiêu đề cột ở file gốc, vd 'Marital Status '. Bỏ đi để tiêu đề chuẩn
# hóa vẫn khớp.
_PUA_RE = re.compile(r"[-]")


def norm(text) -> str:
    """Chuẩn hóa tiêu đề cột để so khớp: NFC · bỏ ký tự PUA · chữ thường · gộp
    khoảng trắng (tiêu đề trong file có cả xuống dòng giữa chừng)."""
    s = unicodedata.normalize("NFC", str(text))
    return " ".join(_PUA_RE.sub("", s).strip().lower().split())


def row_value(row, key):
    """Đọc row[key] an toàn cho cả sqlite3.Row lẫn dict; thiếu/NULL → ""."""
    try:
        value = row[key]
    except (KeyError, IndexError):
        return ""
    return "" if value is None else value


def emergency_contact_cell(row) -> str:
    """Ô "Emergency Contact Name" của file — GỘP tên + SĐT làm một ô.

    File chỉ có một ô cho cả hai (app tách sang hai cột khi import, xem
    `cv_repository.split_contact_name_phone`), nên Copy row phải nối ngược lại
    theo đúng kiểu phổ biến trong file: "tên ⏎ số ĐT". Thiếu vế nào thì dán vế
    còn lại, không để lại dấu xuống dòng thừa.
    """
    name = str(row_value(row, "emergency_contact_name")).strip()
    phone = str(row_value(row, "emergency_contact_phone")).strip()
    return "\n".join(p for p in (name, phone) if p)


class Col(NamedTuple):
    """Một cột của file Excel.

    title      – tiêu đề ĐÚNG NHƯ TRONG FILE (xuống dòng giữa chừng thì viết
                 liền, `norm()` gộp khoảng trắng nên vẫn khớp).
    field      – field bảng `employees` để import, hoặc sentinel cột danh mục.
                 "" = file có cột này nhưng app KHÔNG import (cột phụ trợ, cột
                 công thức, cột đã có nguồn khác).
    aliases    – tên gọi khác của cùng cột ở các bản file khác.
    row_key    – khóa đọc TEXT ra từ kết quả truy vấn; mặc định = `field`.
                 Khác `field` ở 4 cột danh mục (DB lưu id, truy vấn trả text).
    copy       – ô *Copy row*: `VALUE` (dán giá trị cột này) · `None` (ô trống
                 giữ chỗ) · "=công thức" · hàm `f(row) -> str`.
    on_table   – có hiện thành một cột trên bảng nhân viên không.
    table_key  – khóa cột trên bảng nếu KHÁC `row_key` (bảng hiện tên bộ phận
                 đầy đủ trong khi file lưu mã viết tắt).
    label      – nhãn trên bảng & trong thông báo lỗi; mặc định = `title`.
    align      – canh lề ô trên bảng ("w" | "center" | "e").
    in_layout  – cột này có mặt trong BỐ CỤC file hiện tại không. False = chỉ
                 nằm trong từ điển nhận diện (bản file khác từng có cột này),
                 không chiếm ô Copy row và không lên bảng.
    """

    title: str
    field: str = ""
    aliases: tuple = ()
    row_key: str = ""
    copy: object = VALUE
    on_table: bool = True
    table_key: str = ""
    label: str = ""
    align: str = "w"
    in_layout: bool = True

    @property
    def text_key(self) -> str:
        """Khóa đọc giá trị TEXT của cột từ một dòng truy vấn."""
        return self.row_key or self.field

    @property
    def key(self) -> str:
        """Khóa cột trên bảng nhân viên."""
        return self.table_key or self.text_key

    @property
    def heading(self) -> str:
        return self.label or self.title

    @property
    def titles(self) -> tuple:
        """Tiêu đề chính + mọi tên gọi khác."""
        return (self.title, *self.aliases)


class AppCol(NamedTuple):
    """Cột CHỈ CÓ TRONG APP (file Excel không có), chèn vào bảng nhân viên ngay
    sau cột `after` của file — "" nghĩa là chèn ở đầu bảng."""

    key: str
    heading: str
    align: str = "w"
    after: str = ""


# ════════════════════════ BỐ CỤC FILE (A → CD) ══════════════════════════════
# Xếp ĐÚNG THỨ TỰ CỘT TRONG FILE. Chữ cái cột ghi ở comment chỉ để dò cho
# nhanh — thứ tự trong list mới là thứ quyết định.
COLUMNS = [
    Col("Legal Entity (Company)", copy=None, on_table=False),                   # A
    Col("Count", aliases=("STT",), copy="=ROW()-6", on_table=False),            # B
    Col("EC", "code", aliases=("Emp code", "Employee code", "Code"),            # C
        label="Emp code", align="w"),
    Col("GlobalEmpCode", "global_code",                                         # D
        aliases=("GlobalEmp code", "Global Emp Code", "Global code",
                 "global_code"),
        label="Global code"),
    Col("Full name", "full_name", aliases=("FullName",), label="Full name"),    # E
    Col("Surname", "surname", label="Surname"),                                 # F
    Col("Name", "name", label="Name"),                                          # G
    Col("Middle Name (only for vietnam)", "middle_name",                        # H
        aliases=("Middle name",), label="Middle name"),
    Col("Date of Birth", "date_of_birth", aliases=("DOB",),                     # I
        label="Date of birth", align="center"),
    Col("Gender", "gender", label="Gender", align="center"),                    # J
    Col("Phone Number", "phone", aliases=("Phone",), label="Phone"),            # K
    Col("Date of Employment", "date_of_employment",                             # L
        label="Date of employment", align="center"),
    Col("Business Unit (Department)",                                           # M
        aliases=("Business Unit", "Department"), copy=None, on_table=False),
    Col("Department (short)", DEPT_TEXT, row_key="department_short_name",       # N
        table_key="department_name", label="Function (dept.)"),
    Col("Function (Common)", "sub_function", aliases=("Function",),             # O
        label="Function (Common)"),
    Col("New Cost center", CC_TEXT, aliases=("Cost center",),                   # P
        row_key="cost_center_code", label="Cost center", align="center"),
    Col("Full name of manager", "manager_name", label="Manager"),               # Q
    Col("Manager Surname (report directly to)", copy=None, on_table=False),     # R
    Col("Manager Name", copy=None, on_table=False),                             # S
    Col("Manager Middle Name (only for vietnam)", copy=None, on_table=False),   # T
    Col("Job Title (Description)", "job_title", label="Job title"),             # U
    Col("Direct/Indirect", "direct_indirect", label="Direct/Indirect",          # V
        align="center"),
    # Không import (suy ra từ employee_types.collar) nhưng VẪN dán khi Copy row
    # và vẫn là một cột trên bảng.
    Col("BC/WC", copy="collar", table_key="collar", label="BC/WC"),             # W
    Col("IBC/DBC/WC", ETYPE_TEXT, row_key="employee_type_code",                 # X
        label="IBC/DBC/WC", align="center"),
    Col("BY GROUP", "by_group", label="By group"),                              # Y
    Col("Working hour/week", "working_hours_per_week",                          # Z
        label="Working hour/week", align="center"),
    Col("Production Line (Internal)", "production_line",                        # AA
        label="Production line"),
    Col("Job level", LEVEL_TEXT, row_key="level_name", label="Job level",       # AB
        align="center"),
    Col("Operator skill", "operator_skill", label="Operator skill"),            # AC
    Col("Permanent/Temporary contract", "contract_permanency",                  # AD
        label="Perm./Temp.", align="center"),
    Col("Type of contract", "contract_type", label="Type of contract"),         # AE
    Col("Starting date of contract", "contract_start_date",                     # AF
        label="Contract start", align="center"),
    Col("Ending date of contract", "contract_end_date",                         # AG
        label="Contract end", align="center"),
    Col("Termination Date", "termination_date", label="Termination date",       # AH
        align="center"),
    Col("Reason for leaving", "leaving_reason", label="Reason for leaving"),    # AI
    Col("Qualification", "qualification", label="Qualification"),               # AJ
    Col("Qualification (Việt Nam)", "qualification_vn",                         # AK
        aliases=("Qualification (Viet Nam)",), label="Qualification (VN)"),
    Col("Education Level", "education", aliases=("Education",),                 # AL
        label="Education level"),
    Col("Major", "major", label="Major"),                                       # AM
    Col("Year of graduated", "graduation_year", label="Year of graduated",      # AN
        align="center"),
    Col("School name", "school_name", label="School name"),                     # AO
    Col("Place of birth", "place_of_birth", label="Place of birth"),            # AP
    Col("ID no.", "id_no", aliases=("ID no",), label="ID no."),                 # AQ
    Col("Issued date", "id_issued_date", label="ID issued date",                # AR
        align="center"),
    Col("Issued Place", "id_issued_place", label="ID issued place"),            # AS
    Col("Native country", "native_place", label="Native place"),                # AT
    Col("City (address)-theo đc thường trú", "city",                            # AU
        aliases=("City (address)",), label="City"),
    Col("Country (address)", "country", label="Country"),                       # AV
    Col("Bank account no.", "bank_account_no", label="Bank account no."),       # AW
    Col("Bank address", "bank_address", label="Bank address"),                  # AX
    Col("Personal Tax Code", "tax_code", label="Personal tax code"),            # AY
    Col("Dependance", "dependants", label="Dependants", align="center"),        # AZ
    Col("Insurance Book No.", "insurance_book_no",                              # BA
        label="Insurance book no."),
    Col("Passport No.", "passport_no", aliases=("Passport no",),                # BB
        label="Passport no."),
    # Bản file khác đặt tên cột này trùng "Issued date" — khai alias để lần
    # xuất hiện THỨ HAI của tên đó rơi vào đúng field hộ chiếu.
    Col("Issued date2", "passport_issued_date", aliases=("Issued date",),       # BC
        label="Passport issued", align="center"),
    Col("Emergency Contact Name", "emergency_contact_name",                     # BD
        copy=emergency_contact_cell, label="Emergency contact"),
    Col("Relationship", "emergency_contact_relationship",                       # BE
        label="Relationship"),
    Col("Địa chỉ thường trú", "permanent_address", label="Permanent address"),  # BF
    Col("Địa chỉ tạm trú", "temporary_address", label="Temporary address"),     # BG
    Col("Personal Email address", "email", aliases=("Email",),                  # BH
        label="Personal email"),
    Col("Company email", "company_email", label="Company email"),               # BI
    Col("Marital Status", "marital_status", label="Marital status"),            # BJ
    Col("Spouse Name", "spouse_name", label="Spouse name"),                     # BK
    Col("Spouse date", "spouse_dob", label="Spouse date", align="center"),      # BL
    Col("Nation (dân tộc)", copy=None, on_table=False),                         # BM
    Col("Religion", "religion", label="Religion"),                              # BN
    Col("Driving forklift", "driving_forklift", label="Driving forklift",       # BO
        align="center"),
    Col("#ER/ JRF", "er_jrf", aliases=("#ER/JRF",), label="#ER/JRF"),           # BP
    # ── 4 cột "Figures" ── file để CÔNG THỨC, app tính lại mỗi lần truy vấn
    # (cv_repository.EMPLOYEE_COMPUTED_SQL) nên không import; Copy row dán
    # công thức vì giá trị phụ thuộc NGÀY HÔM NAY.
    Col("Year of service", copy="=(TODAY()-[@[Date of Employment]])/365",       # BQ
        table_key="years_of_service", label="Year of service", align="center"),
    Col("Age", copy="=YEAR(NOW())-YEAR([@[Date of Birth]])",                    # BR
        table_key="age", label="Age", align="center"),
    Col("Age range",                                                            # BS
        copy=('=IF([@Age]>50,"Over 50",IF([@Age]>=41,"41-50",'
              'IF([@Age]>=31,"31 - 40",IF([@Age]>=21,"21 - 30","Under 21"))))'),
        table_key="age_range", label="Age range", align="center"),
    Col("Length of service",                                                    # BT
        copy=('=IF([@[Year of service]]>5,"5 and more",IF([@[Year of service]]>3,'
              '"3 - 5 years",IF([@[Year of service]]>=2,"2 - 3 years",'
              'IF([@[Year of service]]>=1,"1- 2 years","less than 1"))))'),
        table_key="length_of_service", label="Length of service"),
    Col("Type of labor", "labor_type", label="Type of labor"),                  # BU
    Col("Eligible -Smart Working Policy Eligible", "smart_working_eligible",    # BV
        label="Smart working", align="center"),
    Col("Changing notes", "changing_notes", label="Changing notes"),            # BW
    Col("Changing date", "changing_date", label="Changing date",                # BX
        align="center"),
    Col("Updated changing date", "updated_changing_date",                       # BY
        label="Updated changing", align="center"),
    Col("Note", "note", label="Note"),                                          # BZ
    Col("Cột tính thâm niên", "seniority_date", label="Seniority date",         # CA
        align="center"),
    Col("Time in Position", "time_in_position", label="Time in position"),      # CB
    Col("Current Position", "current_position", label="Current position"),      # CC
    Col("Birthday", copy="=MONTH([@[Date of Birth]])", on_table=False),         # CD

    # ── Ngoài bố cục: tiêu đề của các BẢN FILE KHÁC ──────────────────────────
    # Nhận diện được khi gặp, nhưng không chiếm ô Copy row và không lên bảng.
    Col("Nationality", "nationality", in_layout=False),
    Col("Marriage status (Yes)", "marriage_status", aliases=("Marriage status",),
        in_layout=False),
    Col("Number of children", "children_count", in_layout=False),
    Col("Children's name", "children_names", in_layout=False),
    Col("Emergency contact phone", "emergency_contact_phone",
        aliases=("Emergency contact number",), in_layout=False),
    # Công thức TRÙNG HỆT cột Age (tiêu đề gây hiểu nhầm) nên cột này đã bỏ khỏi
    # cả app lẫn file. Vẫn khai ở đây để bản file CŨ còn cột đó import được mà
    # không bị báo "cột lạ" — nhưng không chiếm ô Copy row.
    Col("Year of birthday (year)", in_layout=False),
    Col("Level", in_layout=False),                    # cột số phụ trợ, trùng "Job level"
    Col("Job title with level (no use)", in_layout=False),   # file ghi rõ "no use"
    Col("(old) Phone Number", in_layout=False),
    Col("Street (address)", aliases=("Address",), in_layout=False),
    Col("Position status", aliases=("Status",), in_layout=False),   # suy từ Termination Date
]

# Cột chỉ app mới có, chèn vào bảng nhân viên cạnh cột cùng chủ đề.
APP_COLUMNS = [
    AppCol("employee_id", "ID", "center"),
    AppCol("cost_center_group", "CC group", "w", after="cost_center_code"),
    AppCol("work_status", "Status", "center", after="termination_date"),
    AppCol("nationality", "Nationality", "w", after="native_place"),
    AppCol("emergency_contact_phone", "Emergency phone", "w",
           after="emergency_contact_name"),
    AppCol("marriage_status", "Marriage status", "center", after="marital_status"),
    AppCol("children_count", "Children", "center", after="spouse_dob"),
    AppCol("children_names", "Children's names", "w", after="children_count"),
    AppCol("is_interviewer", "Interviewer", "center",
           after="smart_working_eligible"),
]


# ════════════════════════════ SINH RA TỪ COLUMNS ════════════════════════════

def header_map() -> dict:
    """{tiêu đề chuẩn hóa → field} cho các cột CÓ import.

    Tiêu đề trùng nhau ở nhiều cột (vd "Issued date" dùng cho cả CMND lẫn hộ
    chiếu) trả về TUPLE theo thứ tự cột: lần xuất hiện thứ n trong file lấy
    phần tử thứ n.
    """
    out = {}
    for col in COLUMNS:
        if not col.field:
            continue
        for title in col.titles:
            key = norm(title)
            previous = out.get(key)
            if previous is None:
                out[key] = col.field
            elif isinstance(previous, tuple):
                if col.field not in previous:
                    out[key] = previous + (col.field,)
            elif previous != col.field:
                out[key] = (previous, col.field)
    return out


def ignored_headers() -> set:
    """Tiêu đề chuẩn hóa của các cột CỐ TÌNH bỏ qua khi import — gặp thì im
    lặng bỏ, không báo "cột lạ"."""
    return {norm(title) for col in COLUMNS if not col.field
            for title in col.titles}


def copy_layout() -> list:
    """Các ô của *Copy row*, theo đúng thứ tự cột của file.

    Mỗi phần tử: khóa cột (dán giá trị) · `None` (ô trống giữ chỗ) · chuỗi mở
    đầu "=" (công thức) · hàm `f(row) -> str`.
    """
    return [(col.text_key if col.copy is VALUE else col.copy)
            for col in COLUMNS if col.in_layout]


def table_columns() -> list:
    """[(khóa, tiêu đề, canh lề)] cột bảng nhân viên — cột của file theo đúng
    thứ tự file, xen các cột riêng của app tại mốc `after` của chúng."""
    after = {}
    for app_col in APP_COLUMNS:
        after.setdefault(app_col.after, []).append(app_col)

    def spread(anchor):
        """Các cột app gắn sau `anchor` — đệ quy vì một cột app còn có thể làm
        mốc cho cột app khác (Children's names đứng sau Children)."""
        out = []
        for app_col in after.get(anchor, ()):
            out.append((app_col.key, app_col.heading, app_col.align))
            out += spread(app_col.key)
        return out

    out = spread("")
    for col in COLUMNS:
        if not (col.in_layout and col.on_table):
            continue
        out.append((col.key, col.heading, col.align))
        out += spread(col.key)
    return out


def file_titles() -> dict:
    """{field (hoặc sentinel danh mục) → TIÊU ĐỀ ĐÚNG NHƯ TRONG FILE}.

    Dùng cho thông báo kiểu "sửa lại ô này trong file Excel": phải gọi đúng cái
    tên người dùng nhìn thấy khi mở file, không phải nhãn rút gọn trên bảng.
    """
    return {col.field: col.title for col in COLUMNS if col.field}


def file_title(field: str) -> str:
    """Tiêu đề trong file của một field; không phải cột của file thì trả lại
    chính tên field."""
    return file_titles().get(field, field)


def sync_labels() -> dict:
    """{khóa đọc text → nhãn} cho bảng duyệt của lượt đồng bộ.

    Lấy NHÃN TRÊN BẢNG khi khóa đó đúng là một cột của bảng (người dùng vừa
    nhìn thấy nó ở màn hình Employees), còn lại lấy tiêu đề trong file — như
    "Department (short)", thứ mà bảng không hiện vì bảng để tên bộ phận đầy đủ.
    """
    app_headings = {c.key: c.heading for c in APP_COLUMNS}
    out = dict(app_headings)
    for col in COLUMNS:
        key = col.text_key
        # Khóa đã là một cột riêng của app trên bảng (vd "Children's names",
        # cột mà file gộp kiểu khác) → giữ nhãn của bảng.
        if not key or key in app_headings:
            continue
        on_table = col.in_layout and col.on_table and col.key == key
        out[key] = col.heading if on_table else col.title
    return out


def sync_row_keys() -> dict:
    """{sentinel cột danh mục → khóa đọc text từ truy vấn} — lượt đồng bộ so
    theo TEXT trong khi DB lưu id."""
    return {col.field: col.row_key for col in COLUMNS
            if col.row_key and col.field}


def layout_titles() -> list:
    """Tiêu đề các cột theo đúng bố cục file — để đối chiếu với file thật."""
    return [col.title for col in COLUMNS if col.in_layout]


def column_letter(index: int) -> str:
    """0 → 'A', 26 → 'AA' — chỉ để chỉ chỗ trong thông báo lệch bố cục."""
    letters = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def layout_mismatches(header_row) -> list:
    """So dòng tiêu đề THẬT của file với bố cục khai ở `COLUMNS` → list mô tả
    chỗ lệch (rỗng = khớp).

    Import và Sync khớp theo TÊN nên không việc gì, nhưng *Copy row* dán theo
    VỊ TRÍ: file thêm/bớt một cột là mọi ô phía sau dán lệch một ô mà không có
    dấu hiệu gì. Đối chiếu ở đây để biến cái hỏng âm thầm đó thành một dòng
    cảnh báo.
    """
    wanted = [norm(t) for t in layout_titles()]
    found = [norm(t) for t in header_row if t is not None and str(t).strip()]
    issues = []
    for i, (a, b) in enumerate(zip(wanted, found)):
        if a != b:
            issues.append(
                f'column {column_letter(i)}: the file has "{found[i]}" where the '
                f'app expects "{a}"')
            break        # lệch một cột là mọi cột sau lệch theo, báo một dòng đủ
    if not issues and len(found) != len(wanted):
        issues.append(
            f"the file has {len(found)} named columns, the app expects "
            f"{len(wanted)}")
    return issues
