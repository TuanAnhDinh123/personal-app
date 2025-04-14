import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QMessageBox

def xu_ly_chinh():
    # Đây là nơi bạn gọi logic xử lý chính
    # Ví dụ: xử lý file Excel
    try:
        # Ví dụ xử lý giả
        print("Đang xử lý dữ liệu...")
        # Gọi hàm xử lý thực tế ở đây
        # kết_qua = tinh_toan_excel()
        # Giả lập thông báo
        QMessageBox.information(None, "Hoàn tất", "Xử lý thành công!")
    except Exception as e:
        QMessageBox.critical(None, "Lỗi", f"Đã xảy ra lỗi:\n{str(e)}")

app = QApplication(sys.argv)

# Tạo cửa sổ chính
window = QWidget()
window.setWindowTitle("Ứng dụng đơn giản")
window.setFixedSize(300, 150)

# Tạo layout và nút
layout = QVBoxLayout()
button = QPushButton("Bấm để xử lý")
button.clicked.connect(xu_ly_chinh)

layout.addWidget(button)
window.setLayout(layout)

# Hiển thị giao diện
window.show()
sys.exit(app.exec_())
