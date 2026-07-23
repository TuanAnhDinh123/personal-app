"""Form nhập liệu tổng quát (frameless, bo góc) — bản Qt của _open_entity_form.

Dựng field từ danh sách `specs`. Mỗi spec là dict {"key","label","kind"[,...]}:
    kind ∈ text | int | decimal | textarea | file | dropdown | choice | section
  - dropdown: "options" = dict {tên → id} (hoặc callable trả dict).
  - choice:   "choices" = list[str]; "allow_empty" (mặc định False).
  - file:     "filetypes" = list[(nhãn, "*.ext *.ext2")].
  - section:  chỉ hiện tiêu đề nhóm.
  - "required": True → bắt buộc nhập.
"""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QComboBox, QFileDialog, QFrame, QHBoxLayout, QLabel, QLineEdit,
    QScrollArea, QTextEdit, QVBoxLayout, QWidget, QDialog,
)

from app_qt import dialogs, theme, widgets

_NONE = "— Chưa chọn —"


def _num(text, kind):
    s = str(text).strip()
    if not s:
        return None
    try:
        return int(float(s)) if kind == "int" else float(s)
    except ValueError:
        return None


def _filter_str(filetypes):
    if not filetypes:
        return "Tất cả (*.*)"
    return ";;".join(f"{label} ({pat})" for label, pat in filetypes)


class FormDialog(QDialog):
    def __init__(self, parent, title, specs, current=None, on_save=None, on_delete=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.setModal(True)
        self._specs = specs
        self._current = current
        self._on_save = on_save
        self._on_delete = on_delete
        self._getters = {}
        self._required = {}
        self._saved = False
        self._drag = None

        shell = QVBoxLayout(self)
        shell.setContentsMargins(24, 24, 24, 24)
        card = QFrame(self)
        card.setObjectName("Dialog")
        card.setMinimumWidth(520)
        widgets.add_shadow(card, blur=48, dy=12, alpha=70)
        shell.addWidget(card)

        lay = QVBoxLayout(card)
        lay.setContentsMargins(22, 20, 22, 18)   # padding thẻ→nội dung chuẩn
        lay.setSpacing(12)

        # header
        head = QHBoxLayout()
        t = QLabel(title); t.setObjectName("DialogTitle")
        head.addWidget(t, 1)
        close = QLabel("✕"); close.setObjectName("DialogClose")
        close.setFixedSize(24, 24); close.setAlignment(Qt.AlignCenter)
        close.setCursor(Qt.PointingHandCursor)
        close.mousePressEvent = lambda _e: self.reject()
        head.addWidget(close, 0, Qt.AlignTop)
        lay.addLayout(head)

        # thân form cuộn được
        form = QWidget()
        self._form_col = QVBoxLayout(form)
        self._form_col.setContentsMargins(0, 0, 8, 0)
        self._form_col.setSpacing(4)
        self._build_fields(form)
        self._form_col.addStretch(1)

        sa = widgets.scroll_area(form)   # viewport trong suốt → không còn nền xám
        sa.setMinimumHeight(360)
        sa.setMaximumHeight(560)
        lay.addWidget(sa, 1)

        # footer
        foot = QHBoxLayout()
        foot.setContentsMargins(0, 6, 0, 0)
        foot.addWidget(widgets.button(card, "Lưu", variant="success", icon="save",
                                      command=self._save))
        foot.addWidget(widgets.button(card, "Hủy", variant="neutral", icon="x",
                                      command=self.reject))
        foot.addStretch(1)
        if on_delete is not None:
            foot.addWidget(widgets.button(card, "Xóa", variant="danger", icon="trash",
                                          command=self._delete))
        lay.addLayout(foot)

    # ------------------------------------------------------------- dựng field
    def _cur(self, key):
        c = self._current
        if c is None:
            return ""
        try:
            v = c[key]
        except (KeyError, IndexError):
            return ""
        return "" if v is None else v

    def _label(self, form, text):
        lbl = QLabel(text, form)
        lbl.setObjectName("FieldLabel")
        lbl.setContentsMargins(0, 8, 0, 2)
        self._form_col.addWidget(lbl)

    def _build_fields(self, form):
        for spec in self._specs:
            kind = spec.get("kind", "text")
            key = spec.get("key")
            label = spec.get("label", "")

            if kind == "section":
                widgets.section_label(form, label)
                continue
            if spec.get("required"):
                self._required[key] = label

            if kind in ("text", "int", "decimal"):
                self._label(form, label)
                edit = QLineEdit(form)
                val = self._cur(key)
                if val != "":
                    edit.setText(str(val))
                self._form_col.addWidget(edit)
                if kind == "text":
                    self._getters[key] = lambda e=edit: e.text().strip() or None
                else:
                    self._getters[key] = lambda e=edit, k=kind: _num(e.text(), k)

            elif kind == "textarea":
                self._label(form, label)
                box = QTextEdit(form)
                box.setAcceptRichText(False)
                box.setFixedHeight(max(60, spec.get("height", 3) * 22))
                box.setPlainText(str(self._cur(key)))
                self._form_col.addWidget(box)
                self._getters[key] = lambda b=box: b.toPlainText().strip() or None

            elif kind == "file":
                self._label(form, label)
                row = QHBoxLayout(); row.setSpacing(8)
                edit = QLineEdit(form); edit.setText(str(self._cur(key)))
                row.addWidget(edit, 1)
                ft = spec.get("filetypes")

                def _pick(e=edit, f=ft):
                    p, _ = QFileDialog.getOpenFileName(self, "Chọn file", "", _filter_str(f))
                    if p:
                        e.setText(p)

                row.addWidget(widgets.button(form, "Chọn…", variant="neutral",
                                             icon="folder", command=_pick))
                self._form_col.addLayout(row)
                self._getters[key] = lambda e=edit: e.text().strip() or None

            elif kind == "dropdown":
                opts = spec["options"]() if callable(spec["options"]) else spec["options"]
                self._label(form, label)
                combo = QComboBox(form)
                combo.addItems([_NONE] + list(opts))
                try:
                    cur_id = self._current[key] if self._current is not None else None
                except (KeyError, IndexError):
                    cur_id = None
                cur_name = next((n for n, i in opts.items() if i == cur_id), _NONE)
                combo.setCurrentText(cur_name)
                self._form_col.addWidget(combo)
                self._getters[key] = lambda c=combo, o=opts: o.get(c.currentText())

            elif kind == "choice":
                choices = list(spec["choices"])
                allow_empty = spec.get("allow_empty", False)
                values = ([_NONE] + choices) if allow_empty else choices
                self._label(form, label)
                combo = QComboBox(form)
                combo.addItems(values)
                cur = self._cur(key)
                combo.setCurrentText(cur if cur else (_NONE if allow_empty else choices[0]))
                self._form_col.addWidget(combo)
                self._getters[key] = lambda c=combo: (None if c.currentText() == _NONE
                                                      else c.currentText())

    # ------------------------------------------------------------- lưu
    def _save(self):
        data = {k: g() for k, g in self._getters.items()}
        for k, lbl in self._required.items():
            if not data.get(k):
                dialogs.warning(self, "Thiếu thông tin", f'Vui lòng nhập "{lbl}".')
                return
        if self._on_save is not None:
            if self._on_save(data) is False:
                return   # giữ form mở (vd người dùng hủy khi báo trùng)
        self._data = data
        self._saved = True
        self.accept()

    def _delete(self):
        # on_delete tự lo xác nhận + xóa; trả về False để giữ form mở (vd hủy xác nhận).
        if self._on_delete is not None and self._on_delete() is not False:
            self.reject()

    # kéo di chuyển
    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            self._drag = e.globalPosition().toPoint() - self.frameGeometry().topLeft()

    def mouseMoveEvent(self, e):
        if self._drag is not None and e.buttons() & Qt.LeftButton:
            self.move(e.globalPosition().toPoint() - self._drag)

    def mouseReleaseEvent(self, e):
        self._drag = None

    def run(self):
        """Trả về data dict nếu lưu (khi không dùng on_save), hoặc True/False."""
        self.exec()
        if self._on_save is not None:
            return self._saved
        return getattr(self, "_data", None) if self._saved else None
