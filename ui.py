import resources 
from PyQt5 import QtCore, QtGui, QtWidgets
import sys
import platform
import winreg
import ctypes
from ctypes import wintypes

icon_app_link = ":/icons/images/icons/download.png"
icon_setting_link = ":/icons/images/icons/cogwheel.png"

class StyleManager:
    """Quản lý các style cho ứng dụng"""
    @staticmethod
    def get_button_style(theme):
        return f"""
            QPushButton {{
                background-color: {'#2a82da' if theme == 'dark' else '#2a82da'};
                border-radius: 8px;
                color: white;
                font-family: 'Segoe UI';
            }}
            QPushButton:hover {{
                background-color: {'#3a92ea' if theme == 'dark' else '#3a92ea'};
            }}
            QPushButton:pressed {{
                background-color: {'#1a72ca' if theme == 'dark' else '#1a72ca'};
            }}
            QPushButton:disabled {{
                background-color: #505050;
            }}
        """

    @staticmethod
    def get_stop_button_style(theme):
        return f"""
            QPushButton {{
                background-color: {'#d84f4f' if theme == 'dark' else '#ff6b6b'};
                border-radius: 8px;
                color: white;
            }}
            QPushButton:hover {{
                background-color: {'#e85f5f' if theme == 'dark' else '#ff7b7b'};
            }}
            QPushButton:pressed {{
                background-color: {'#c83f3f' if theme == 'dark' else '#ff5b5b'};
            }}
        """

    @staticmethod
    def get_frame_style(theme):
        return f"""
            QFrame {{
                background-color: {'#424242' if theme == 'dark' else "#ffffffff"};
                border: 1px solid {'#555555' if theme == 'dark' else '#cccccc'};
                border-radius: 8px;
            }}
            QLabel {{
                color: {'#ffffff' if theme == 'dark' else '#000000'};
                border: none;
            }}
        """

    @staticmethod
    def get_entry_style(theme):
        return f"""
            QLineEdit, QPlainTextEdit {{
                background-color: {'#353535' if theme == 'dark' else '#ffffff'};
                color: {'#ffffff' if theme == 'dark' else '#000000'};
                border: 1px solid {'#555555' if theme == 'dark' else '#cccccc'};
                border-radius: 7px;
                padding: 5px;
                font-size: 12px;
                font-family: 'Segoe UI';
                selection-background-color: #2a82da;
            }}
            QLineEdit:hover, QPlainTextEdit:hover {{
                border: 1px solid {'#6a6a6a' if theme == 'dark' else '#aaaaaa'};
            }}
        """

    @staticmethod
    def get_checkbox_style(theme):
        return f"""
            QCheckBox {{
                color: {'#ffffff' if theme == 'dark' else '#000000'};
            }}
            QCheckBox::indicator {{
                width: 15px;
                height: 15px;
            }}
            QCheckBox::indicator:unchecked {{
                border: 2px solid {'#aaaaaa' if theme == 'dark' else '#666666'};
                background: {'#353535' if theme == 'dark' else '#ffffff'};
            }}
            QCheckBox::indicator:checked {{
                border: 2px solid {'#aaaaaa' if theme == 'dark' else '#666666'};
                background: {'#0078d7' if theme == 'dark' else '#0078d7'};
            }}
        """


class GearButton(QtWidgets.QPushButton):
    """Nút bánh răng để hiển thị các tùy chọn"""
    def __init__(self, parent=None, size=28):
        super().__init__(parent)
        self.setFixedSize(size, size)
        self.setCursor(QtCore.Qt.PointingHandCursor)
        self._setup_icon(size)
        self._setup_style()

    def _setup_icon(self, size):
        """Cấu hình icon cho nút"""
        try:
            icon = QtGui.QIcon(icon_setting_link)
            if icon.isNull():
                icon = QtGui.QIcon.fromTheme("preferences-system")
        except Exception:
            icon = QtGui.QIcon.fromTheme("preferences-system")
        self.setIcon(icon)
        self.setIconSize(QtCore.QSize(size-4, size-4))

    def _setup_style(self):
        """Cấu hình style cho nút"""
        self.setStyleSheet("""
            QPushButton {
                background: transparent;
                border: none;
            }
            QPushButton:hover {
                background-color: rgba(200, 200, 200, 60);
                border-radius: 14px;
            }
        """)


class BaseToggle(QtWidgets.QCheckBox):
    """Lớp cơ sở cho các toggle button tùy chỉnh"""
    def __init__(
        self, parent=None, width=60,
        bg_color="#b0b0b0", circle_color="#fff",
        active_color="#2a82da", label_on="On", label_off="Off",
        text_color=None  # Thêm tham số text_color
    ):
        super().__init__(parent)
        self._text_color = text_color if text_color else "#222"  # Màu mặc định
        self._setup_ui(width, bg_color, circle_color, active_color, label_on, label_off)
        self._setup_animation()
        self._setup_mouse_states()
        self._connect_signals()

    def _setup_ui(self, width, bg_color, circle_color, active_color, label_on, label_off):
        """Cấu hình giao diện ban đầu"""
        self.setFixedSize(width, 28)
        self.setCursor(QtCore.Qt.PointingHandCursor)

        self._bg_color = bg_color
        self._circle_color = circle_color
        self._active_color = active_color
        self._label_on = label_on
        self._label_off = label_off

        self._create_external_label()

    def _create_external_label(self):
        """Tạo label bên ngoài cho toggle"""
        self._external_label = QtWidgets.QLabel(self.parent())
        self._external_label.setText(self._label_on if self.isChecked() else self._label_off)
        self._external_label.setFont(QtGui.QFont("Segoe UI", 10, QtGui.QFont.Bold))
        self._external_label.setStyleSheet("color: #222; background: transparent;")
        self._external_label.setAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignTop)
        self._external_label.resize(self.width(), 20)
        self._external_label.move(self.x(), self.y() + self.height() + 2)
        self._external_label.show()

    def _setup_animation(self):
        """Cấu hình animation cho toggle"""
        self._bar_margin = 8
        self._bar_height = 12
        self._bar_y = (self.height() - self._bar_height) // 2
        self._bar_width = self.width() - 2 * self._bar_margin
        self._circle_diameter = 20
        self._circle_y = (self.height() - self._circle_diameter) // 2
        self._min_pos = self._bar_margin
        self._max_pos = self._bar_margin + self._bar_width - self._circle_diameter
        self._circle_pos = self._min_pos if not self.isChecked() else self._max_pos
        
        self.animation = QtCore.QPropertyAnimation(self, b"circlePos")
        self.animation.setDuration(200)

    def _setup_mouse_states(self):
        """Cấu hình trạng thái chuột"""
        self._hover = False
        self._pressed = False

    def _connect_signals(self):
        """Kết nối các signal"""
        self.toggled.connect(self.start_animation)
        self.toggled.connect(self.update_external_label)
        self.start_animation(self.isChecked())
        self.update_external_label(self.isChecked())

    def moveEvent(self, event):
        """Xử lý sự kiện di chuyển"""
        if hasattr(self, '_external_label') and self._external_label:
            self._external_label.move(self.x(), self.y() + self.height() + 2)
        super().moveEvent(event)

    def resizeEvent(self, event):
        """Xử lý sự kiện thay đổi kích thước"""
        if hasattr(self, '_external_label') and self._external_label:
            self._external_label.resize(self.width(), 20)
        super().resizeEvent(event)

    def update_external_label(self, checked):
        """Cập nhật label bên ngoài"""
        if hasattr(self, '_external_label') and self._external_label:
            self._external_label.setText(self._label_on if checked else self._label_off)
            # Sử dụng màu sắc từ thuộc tính _text_color
            self._external_label.setStyleSheet(
                f"color: {self._text_color}; background: transparent;"
            )

    def set_text_color(self, color):
        """Thiết lập màu sắc cho text"""
        self._text_color = color
        self.update_external_label(self.isChecked())

    @QtCore.pyqtProperty(int)
    def circlePos(self):
        return self._circle_pos

    @circlePos.setter
    def circlePos(self, pos):
        self._circle_pos = pos
        self.update()

    def start_animation(self, checked):
        """Bắt đầu animation khi toggle thay đổi trạng thái"""
        self.animation.stop()
        end_pos = self._max_pos if checked else self._min_pos
        self.animation.setStartValue(self._circle_pos)
        self.animation.setEndValue(end_pos)
        self.animation.setEasingCurve(QtCore.QEasingCurve.OutCubic)
        self.animation.start()

    def paintEvent(self, event):
        """Vẽ toggle button tùy chỉnh"""
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.Antialiasing)

        self._draw_shadow(painter)
        self._draw_background(painter)
        if self.isChecked():
            self._draw_glow(painter)
        self._draw_circle(painter)

    def _draw_shadow(self, painter):
        """Vẽ shadow cho toggle"""
        shadow_color = QtGui.QColor(0, 0, 0, 30)
        shadow_rect = QtCore.QRectF(self._bar_margin, self._bar_y, self._bar_width, self._bar_height)
        painter.setPen(QtCore.Qt.NoPen)
        painter.setBrush(shadow_color)
        painter.drawRoundedRect(shadow_rect, self._bar_height / 2, self._bar_height / 2)

    def _draw_background(self, painter):
        """Vẽ nền cho toggle"""
        bg_rect = QtCore.QRectF(self._bar_margin, self._bar_y, self._bar_width, self._bar_height)
        bg_color = QtGui.QColor(self._active_color if self.isChecked() else self._bg_color)
        if self._hover:
            bg_color = bg_color.lighter(110)
        if self._pressed:
            bg_color = bg_color.darker(110)
        painter.setBrush(bg_color)
        painter.drawRoundedRect(bg_rect, self._bar_height / 2, self._bar_height / 2)

    def _draw_glow(self, painter):
        """Vẽ hiệu ứng glow khi toggle được bật"""
        glow_color = QtGui.QColor(self._active_color)
        glow_color.setAlpha(80)
        glow_rect = QtCore.QRectF(
            self._circle_pos - 2, self._circle_y - 2,
            self._circle_diameter + 4, self._circle_diameter + 4
        )
        painter.setBrush(glow_color)
        painter.drawEllipse(glow_rect)

    def _draw_circle(self, painter):
        """Vẽ vòng tròn di chuyển của toggle"""
        circle_rect = QtCore.QRectF(
            self._circle_pos, self._circle_y,
            self._circle_diameter, self._circle_diameter
        )
        grad = QtGui.QRadialGradient(circle_rect.center(), self._circle_diameter / 2)
        grad.setColorAt(0, QtGui.QColor("#fff"))
        grad.setColorAt(1, QtGui.QColor("#e0e0e0"))
        painter.setBrush(QtGui.QBrush(grad))
        painter.setPen(QtGui.QPen(QtGui.QColor("#cccccc"), 1))
        painter.drawEllipse(circle_rect)

    def enterEvent(self, event):
        self._hover = True
        self.update()
        super().enterEvent(event)
    
    def closeEvent(self, event):
        if hasattr(self, '_external_label') and self._external_label:
            self._external_label.deleteLater()
        super().closeEvent(event)

    def leaveEvent(self, event):
        self._hover = False
        self.update()
        super().leaveEvent(event)

    def mousePressEvent(self, event):
        self._pressed = True
        self.update()
        super().mousePressEvent(event)

    def mouseReleaseEvent(self, event):
        self._pressed = False
        self.update()
        super().mouseReleaseEvent(event)


class AnimatedToggle(BaseToggle):
    """Toggle button cho chế độ Dark/Light"""
    def __init__(self, parent=None, width=70, **kwargs):
        kwargs.setdefault('label_on', 'Dark')
        kwargs.setdefault('label_off', 'Light')
        super().__init__(parent, width=width, **kwargs)


class FileModeToggle(BaseToggle):
    """Toggle button cho chế độ file"""
    def __init__(self, parent=None, width=70, **kwargs):
        kwargs.setdefault('label_on', 'MAP File')
        kwargs.setdefault('label_off', 'Simple File')
        super().__init__(parent, width=width, **kwargs)


class Ui_MainWindow:
    """Lớp giao diện chính của ứng dụng"""
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(740, 670)
        
        # Central widget
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        
        # Setup các thành phần UI
        self._setup_title()
        self._setup_theme_toggle()
        self._setup_file_mode_toggle()
        self._setup_input_output_frame()
        self._setup_settings_frame()
        self._setup_file_type_frame()
        self._setup_process_frame()
        self._setup_button_frame()
        
        MainWindow.setCentralWidget(self.centralwidget)
        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        
        # Cấu hình ban đầu
        self._initial_configuration()

    def _setup_title(self):
        """Cấu hình tiêu đề ứng dụng"""
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(120, 10, 500, 120))
        font = QtGui.QFont()
        font.setFamily("Segoe UI")
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(60)
        self.label.setFont(font)
        self.label.setObjectName("label")

    def _setup_theme_toggle(self):
        """Cấu hình nút chuyển đổi chế độ Dark/Light"""
        self.theme_toggle = AnimatedToggle(self.centralwidget, width=70)
        self.theme_toggle.setGeometry(QtCore.QRect(625, 10, 70, 28))
        self.theme_toggle.setObjectName("theme_toggle")

    def _setup_file_mode_toggle(self):
        """Cấu hình nút chuyển đổi chế độ file"""
        self.file_mode_toggle = FileModeToggle(self.centralwidget, width=70)
        self.file_mode_toggle.setGeometry(QtCore.QRect(625, 70, 70, 28))
        self.file_mode_toggle.setObjectName("file_mode_toggle")

    def _setup_input_output_frame(self):
        """Cấu hình frame input/output"""
        self.inout_frame = QtWidgets.QFrame(self.centralwidget)
        self.inout_frame.setGeometry(QtCore.QRect(20, 150, 700, 175))
        self.inout_frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.inout_frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.inout_frame.setObjectName("inout_frame")
        
        # Tạo các thành phần trong frame
        self._create_input_output_entries()
        self._create_input_output_labels()
        self._create_input_output_buttons()

    def _create_input_output_entries(self):
        """Tạo các ô nhập liệu"""
        entry_style = StyleManager.get_entry_style("light")
        
        self.output_folder_entry = QtWidgets.QLineEdit(self.inout_frame)
        self.output_folder_entry.setGeometry(QtCore.QRect(140, 120, 470, 35))
        self.output_folder_entry.setStyleSheet(entry_style)
        self.output_folder_entry.setObjectName("output_folder_entry")
        
        self.input_file_entry_2 = QtWidgets.QLineEdit(self.inout_frame)
        self.input_file_entry_2.setGeometry(QtCore.QRect(140, 70, 470, 35))
        self.input_file_entry_2.setStyleSheet(entry_style)
        self.input_file_entry_2.setObjectName("input_file_entry_2")
        
        self.input_file_entry_1 = QtWidgets.QLineEdit(self.inout_frame)
        self.input_file_entry_1.setGeometry(QtCore.QRect(140, 20, 470, 35))
        self.input_file_entry_1.setStyleSheet(entry_style)
        self.input_file_entry_1.setObjectName("input_file_entry_1")

    def _create_input_output_labels(self):
        """Tạo các nhãn cho frame input/output"""
        label_font = QtGui.QFont()
        label_font.setFamily("Segoe UI")
        label_font.setPointSize(10)
        
        self.input_file_label_1 = QtWidgets.QLabel(self.inout_frame)
        self.input_file_label_1.setGeometry(QtCore.QRect(20, 30, 110, 16))
        self.input_file_label_1.setFont(label_font)
        self.input_file_label_1.setObjectName("input_file_label_1")
        self.input_file_label_1.setStyleSheet("border: none;")
        
        self.input_file_label_2 = QtWidgets.QLabel(self.inout_frame)
        self.input_file_label_2.setGeometry(QtCore.QRect(20, 80, 110, 16))
        self.input_file_label_2.setFont(label_font)
        self.input_file_label_2.setObjectName("input_file_label_2")
        self.input_file_label_2.setStyleSheet("border: none;")
        
        self.output_label = QtWidgets.QLabel(self.inout_frame)
        self.output_label.setGeometry(QtCore.QRect(20, 130, 110, 16))
        self.output_label.setFont(label_font)
        self.output_label.setObjectName("output_label")
        self.output_label.setStyleSheet("border: none;")

    def _create_input_output_buttons(self):
        """Tạo các nút cho frame input/output"""
        button_style = """
            QPushButton {
                border: none;
                border-radius: 5px;
                padding: 5px;
                font-size: 12px;
                font-family: 'Segoe UI';
                min-width: 60px;
            }
            QPushButton:hover {
                border: 1px solid #aaaaaa;
            }
            QPushButton:pressed {
                border: 1px solid #2a82da;
            }
            QPushButton:disabled {
                color: #a0a0a0;
            }
        """
        
        self.input_file_button_1 = QtWidgets.QPushButton(self.inout_frame)
        self.input_file_button_1.setGeometry(QtCore.QRect(620, 20, 60, 30))
        self.input_file_button_1.setStyleSheet(button_style)
        self.input_file_button_1.setObjectName("input_file_button_1")
        
        self.input_file_button_2 = QtWidgets.QPushButton(self.inout_frame)
        self.input_file_button_2.setGeometry(QtCore.QRect(620, 70, 60, 30))
        self.input_file_button_2.setStyleSheet(button_style)
        self.input_file_button_2.setObjectName("input_file_button_2")
        
        self.output_button = QtWidgets.QPushButton(self.inout_frame)
        self.output_button.setGeometry(QtCore.QRect(620, 120, 60, 30))
        self.output_button.setStyleSheet(button_style)
        self.output_button.setObjectName("output_button")

    def _setup_settings_frame(self):
        """Cấu hình frame settings"""
        self.settings_frame = QtWidgets.QFrame(self.centralwidget)
        self.settings_frame.setGeometry(QtCore.QRect(20, 350, 700, 90))
        self.settings_frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.settings_frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.settings_frame.setObjectName("settings_frame")
        
        # Tạo các thành phần trong frame
        self._create_settings_entries()
        self._create_settings_labels()

    def _create_settings_entries(self):
        """Tạo các ô nhập liệu cho settings"""
        entry_style = StyleManager.get_entry_style("light")
        
        self.folder_column_entry = QtWidgets.QLineEdit(self.settings_frame)
        self.folder_column_entry.setGeometry(QtCore.QRect(70, 40, 55, 35))
        self.folder_column_entry.setStyleSheet(entry_style)
        self.folder_column_entry.setObjectName("folder_column_entry")
        
        self.item_column_entry = QtWidgets.QLineEdit(self.settings_frame)
        self.item_column_entry.setGeometry(QtCore.QRect(320, 40, 55, 35))
        self.item_column_entry.setStyleSheet(entry_style)
        self.item_column_entry.setObjectName("item_column_entry")
        
        self.rev_column_entry = QtWidgets.QLineEdit(self.settings_frame)
        self.rev_column_entry.setGeometry(QtCore.QRect(570, 40, 55, 35))
        self.rev_column_entry.setStyleSheet(entry_style)
        self.rev_column_entry.setObjectName("rev_column_entry")

        # Căn giữa text
        self.folder_column_entry.setAlignment(QtCore.Qt.AlignCenter)
        self.item_column_entry.setAlignment(QtCore.Qt.AlignCenter)
        self.rev_column_entry.setAlignment(QtCore.Qt.AlignCenter)

    def _create_settings_labels(self):
        """Tạo các nhãn cho settings"""
        label_font = QtGui.QFont()
        label_font.setFamily("Segoe UI")
        label_font.setPointSize(10)
        
        self.folder_column_label = QtWidgets.QLabel(self.settings_frame)
        self.folder_column_label.setGeometry(QtCore.QRect(50, 10, 100, 30))
        self.folder_column_label.setFont(label_font)
        self.folder_column_label.setObjectName("folder_column_label")
        self.folder_column_label.setStyleSheet("border: none;")
        
        self.item_column_label = QtWidgets.QLabel(self.settings_frame)
        self.item_column_label.setGeometry(QtCore.QRect(300, 10, 100, 30))
        self.item_column_label.setFont(label_font)
        self.item_column_label.setObjectName("item_column_label")
        self.item_column_label.setStyleSheet("border: none;")
        
        self.revision_column_label = QtWidgets.QLabel(self.settings_frame)
        self.revision_column_label.setGeometry(QtCore.QRect(550, 10, 100, 30))
        self.revision_column_label.setFont(label_font)
        self.revision_column_label.setObjectName("revision_column_label")
        self.revision_column_label.setStyleSheet("border: none;")

    def _setup_file_type_frame(self):
        """Cấu hình frame loại file"""
        self.file_type_frame = QtWidgets.QFrame(self.centralwidget)
        self.file_type_frame.setGeometry(QtCore.QRect(20, 460, 700, 60))
        self.file_type_frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.file_type_frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.file_type_frame.setObjectName("file_type_frame")
        
        # Tạo các checkbox
        checkbox_font = QtGui.QFont()
        checkbox_font.setFamily("Segoe UI")
        checkbox_font.setPointSize(10)
        
        self.datanote = QtWidgets.QCheckBox(self.file_type_frame)
        self.datanote.setGeometry(QtCore.QRect(50, 20, 100, 20))
        self.datanote.setFont(checkbox_font)
        self.datanote.setObjectName("datanote")
        
        self.rev_drawing = QtWidgets.QCheckBox(self.file_type_frame)
        self.rev_drawing.setGeometry(QtCore.QRect(300, 20, 100, 20))
        self.rev_drawing.setFont(checkbox_font)
        self.rev_drawing.setObjectName("rev_drawing")
        
        self.nx_pdf = QtWidgets.QCheckBox(self.file_type_frame)
        self.nx_pdf.setGeometry(QtCore.QRect(550, 20, 100, 20))
        self.nx_pdf.setFont(checkbox_font)
        self.nx_pdf.setObjectName("nx_pdf")

    def _setup_process_frame(self):
        """Cấu hình frame tiến trình"""
        self.process_frame = QtWidgets.QFrame(self.centralwidget)
        self.process_frame.setGeometry(QtCore.QRect(20, 540, 700, 60))
        self.process_frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.process_frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.process_frame.setObjectName("process_frame")
        
        # Tạo các nhãn
        label_font = QtGui.QFont()
        label_font.setFamily("Segoe UI")
        label_font.setPointSize(12)
        
        self.process_label = QtWidgets.QLabel(self.process_frame)
        self.process_label.setGeometry(QtCore.QRect(20, 2, 660, 30))
        self.process_label.setFont(label_font)
        self.process_label.setObjectName("process_label")
        self.process_label.setStyleSheet("border: none;")
        
        self.noti_label = QtWidgets.QLabel(self.process_frame)
        self.noti_label.setGeometry(QtCore.QRect(20, 27, 660, 30))
        self.noti_label.setFont(label_font)
        self.noti_label.setObjectName("noti_label")
        self.noti_label.setStyleSheet("border: none;")

    def _setup_button_frame(self):
        """Cấu hình frame nút"""
        self.button_frame = QtWidgets.QFrame(self.centralwidget)
        self.button_frame.setGeometry(QtCore.QRect(170, 620, 400, 60))
        self.button_frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.button_frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.button_frame.setObjectName("button_frame")
        
        # Tạo các nút hành động
        action_button_font = QtGui.QFont()
        action_button_font.setFamily("Segoe UI")
        action_button_font.setPointSize(12)
        action_button_font.setBold(True)
        
        self.download_button = QtWidgets.QPushButton(self.button_frame)
        self.download_button.setGeometry(QtCore.QRect(50, 10, 120, 40))
        self.download_button.setFont(action_button_font)
        self.download_button.setObjectName("download_button")
        
        self.stop_button = QtWidgets.QPushButton(self.button_frame)
        self.stop_button.setGeometry(QtCore.QRect(230, 10, 120, 40))
        self.stop_button.setFont(action_button_font)
        self.stop_button.setObjectName("stop_button")

    def _initial_configuration(self):
        """Cấu hình ban đầu cho ứng dụng"""
        self.current_theme = "dark"
        self.set_theme("dark")
        self.file_mode_toggle.setChecked(False)
        self.file_mode_toggle.toggled.emit(False)
        self.input_file_label_2.hide()
        self.input_file_entry_2.hide()
        self.input_file_button_2.hide()
        self.settings_frame.show()
        self.input_file_label_1.setText("Simple File:")
        
        # Điều chỉnh vị trí các frame
        self.settings_frame.move(20, 300)
        self.file_type_frame.move(20, 415)
        self.process_frame.move(20, 500)
        self.button_frame.move(170, 585)

        # Điều chỉnh vị trí các thành phần output
        self.output_label.move(20, 80)
        self.output_folder_entry.move(140, 70)
        self.output_button.move(620, 70)

        # Điều chỉnh kích thước frame input/output
        self.inout_frame.setFixedHeight(125)

    def set_theme(self, theme):
        """Thiết lập theme cho ứng dụng"""
        self.current_theme = theme
        
        # Thiết lập palette
        self._set_palette(theme)
        
        # Thiết lập stylesheet
        self._apply_styles(theme)
        
        # Cập nhật text cho các toggle
        self._update_toggle_texts(theme)

    def _set_palette(self, theme):
        """Thiết lập palette cho ứng dụng"""
        palette = QtGui.QPalette()
        
        if theme == "dark":
            palette.setColor(QtGui.QPalette.Window, QtGui.QColor(53, 53, 53))
            palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.white)
            palette.setColor(QtGui.QPalette.Base, QtGui.QColor(35, 35, 35))
            palette.setColor(QtGui.QPalette.AlternateBase, QtGui.QColor(53, 53, 53))
            palette.setColor(QtGui.QPalette.ToolTipBase, QtCore.Qt.white)
            palette.setColor(QtGui.QPalette.ToolTipText, QtCore.Qt.white)
            palette.setColor(QtGui.QPalette.Text, QtCore.Qt.white)
            palette.setColor(QtGui.QPalette.Button, QtGui.QColor(53, 53, 53))
            palette.setColor(QtGui.QPalette.ButtonText, QtCore.Qt.white)
            palette.setColor(QtGui.QPalette.BrightText, QtCore.Qt.red)
            palette.setColor(QtGui.QPalette.Link, QtGui.QColor(42, 130, 218))
            palette.setColor(QtGui.QPalette.Highlight, QtGui.QColor(42, 130, 218))
            palette.setColor(QtGui.QPalette.HighlightedText, QtCore.Qt.black)
        else:
            palette.setColor(QtGui.QPalette.Window, QtGui.QColor(240, 240, 240))
            palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.black)
            palette.setColor(QtGui.QPalette.Base, QtGui.QColor(255, 255, 255))
            palette.setColor(QtGui.QPalette.AlternateBase, QtGui.QColor(240, 240, 240))
            palette.setColor(QtGui.QPalette.ToolTipBase, QtGui.QColor(255, 255, 220))
            palette.setColor(QtGui.QPalette.ToolTipText, QtCore.Qt.black)
            palette.setColor(QtGui.QPalette.Text, QtCore.Qt.black)
            palette.setColor(QtGui.QPalette.Button, QtGui.QColor(240, 240, 240))
            palette.setColor(QtGui.QPalette.ButtonText, QtCore.Qt.black)
            palette.setColor(QtGui.QPalette.BrightText, QtCore.Qt.red)
            palette.setColor(QtGui.QPalette.Link, QtGui.QColor(0, 120, 215))
            palette.setColor(QtGui.QPalette.Highlight, QtGui.QColor(0, 120, 215))
            palette.setColor(QtGui.QPalette.HighlightedText, QtCore.Qt.white)
            
        QtWidgets.QApplication.setPalette(palette)

    def _apply_styles(self, theme):
        """Áp dụng stylesheet cho các thành phần UI"""
        # Áp dụng style cho các frame
        frame_style = StyleManager.get_frame_style(theme)
        self.inout_frame.setStyleSheet(frame_style)
        self.settings_frame.setStyleSheet(frame_style)
        self.file_type_frame.setStyleSheet(frame_style)
        self.process_frame.setStyleSheet(frame_style)
        self.button_frame.setStyleSheet(frame_style)
        
        # Áp dụng style cho các nút
        button_style = StyleManager.get_button_style(theme)
        stop_button_style = StyleManager.get_stop_button_style(theme)
        
        self.input_file_button_1.setStyleSheet(button_style)
        self.input_file_button_2.setStyleSheet(button_style)
        self.output_button.setStyleSheet(button_style)
        self.download_button.setStyleSheet(button_style)
        self.stop_button.setStyleSheet(stop_button_style)
        
        # Áp dụng style cho các ô nhập liệu
        entry_style = StyleManager.get_entry_style(theme)
        self.output_folder_entry.setStyleSheet(entry_style)
        self.input_file_entry_1.setStyleSheet(entry_style)
        self.input_file_entry_2.setStyleSheet(entry_style)
        self.folder_column_entry.setStyleSheet(entry_style)
        self.item_column_entry.setStyleSheet(entry_style)
        self.rev_column_entry.setStyleSheet(entry_style)
        
        # Áp dụng style cho các checkbox
        checkbox_style = StyleManager.get_checkbox_style(theme)
        self.datanote.setStyleSheet(checkbox_style)
        self.rev_drawing.setStyleSheet(checkbox_style)
        self.nx_pdf.setStyleSheet(checkbox_style)

    def _update_toggle_texts(self, theme):
        """Cập nhật text cho các toggle button độc lập"""
        # Cập nhật theme toggle
        self.theme_toggle.setText("Light Mode" if theme == "dark" else "Dark Mode")
        
        # Cập nhật file mode toggle
        self.file_mode_toggle.setText("MAP File" if self.file_mode_toggle.isChecked() else "Simple File")
        
        # Thiết lập màu sắc text dựa trên theme
        text_color = "#fff" if theme == "dark" else "#222"
        self.theme_toggle.set_text_color(text_color)
        self.file_mode_toggle.set_text_color(text_color)

    def retranslateUi(self, MainWindow):
        """Thiết lập các text cho các thành phần UI"""
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Teamcenter Downloader Tool"))
        self.label.setText(_translate("MainWindow", "<html><head/><body>\n"
"<p align=\"center\"><span style=\" font-size:18pt; font-weight:600;\">Teamcenter Downloader Tool</span></p>\n"
f"<p align=\"center\">• Log in to Teamcenter before running</span></p>\n"
f"<p align=\"center\">• Switch your keyboard to English mode</span></p>\n"
f"<p align=\"center\">• Do not operate while the tool is running</span></p>\n"
"</body></html>"))
        self.input_file_label_1.setText(_translate("MainWindow", "Simple File:"))
        self.input_file_label_2.setText(_translate("MainWindow", "Connector IF file:"))
        self.output_label.setText(_translate("MainWindow", "Output Folder:"))
        self.input_file_button_1.setText(_translate("MainWindow", "Browse"))
        self.input_file_button_2.setText(_translate("MainWindow", "Browse"))
        self.output_button.setText(_translate("MainWindow", "Browse"))
        self.folder_column_label.setText(_translate("MainWindow", "Folder Column"))
        self.item_column_label.setText(_translate("MainWindow", "Item ID Column"))
        self.revision_column_label.setText(_translate("MainWindow", "Revision Column"))
        self.datanote.setText(_translate("MainWindow", "Data Note"))
        self.rev_drawing.setText(_translate("MainWindow", "Rev Drawing"))
        self.nx_pdf.setText(_translate("MainWindow", "NX PDF"))
        self.process_label.setText(_translate("MainWindow", "<html><head/><body><p align=\"center\"><span style=\"\">Ready</span></p></body></html>"))
        self.noti_label.setText(_translate("MainWindow", "<html><head/><body><p align=\"center\"><span style=\"\">Status</span></p></body></html>"))
        self.download_button.setText(_translate("MainWindow", "Download"))
        self.stop_button.setText(_translate("MainWindow", "Stop"))

class TitleBarThemeManager:
    """Quản lý màu sắc thanh tiêu đề trên Windows 10/11"""
    def __init__(self):
        self._dwmapi = None
        self._user32 = None
        self._set_title_bar_theme_light = None
        self._set_title_bar_theme_dark = None
        
        # Chỉ khởi tạo trên Windows 10/11
        if platform.system() == "Windows":
            try:
                self._dwmapi = ctypes.WinDLL('dwmapi')
                self._user32 = ctypes.WinDLL('user32')
                self._setup_dark_mode_functions()
            except Exception as e:
                print(f"Không thể khởi tạo Dark Mode: {e}")

    def _setup_dark_mode_functions(self):
        """Thiết lập các hàm API cần thiết"""
        if not hasattr(self, '_dwmapi') or not self._dwmapi:
            return
            
        # Kiểm tra version Windows (chỉ hỗ trợ từ Windows 10 build 1809 trở lên)
        try:
            from sys import getwindowsversion
            win_version = getwindowsversion()
            if win_version.major < 10 or (win_version.major == 10 and win_version.build < 17763):
                return
        except:
            return
            
        try:
            self._dwmapi.DwmSetWindowAttribute.argtypes = [
                wintypes.HWND,
                ctypes.c_int,
                ctypes.c_void_p,
                ctypes.c_int
            ]
            
            # Hằng số cho Dark Mode
            self.DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            self._set_title_bar_theme_light = lambda hwnd: self._set_dark_mode(hwnd, False)
            self._set_title_bar_theme_dark = lambda hwnd: self._set_dark_mode(hwnd, True)
        except Exception as e:
            print(f"Không thể thiết lập Dark Mode functions: {e}")

    def _set_dark_mode(self, hwnd, dark):
        """Thiết lập chế độ tối/sáng cho thanh tiêu đề"""
        if not hasattr(self._dwmapi, 'DwmSetWindowAttribute'):
            return
        
        value = ctypes.c_int(dark)
        self._dwmapi.DwmSetWindowAttribute(
            hwnd,
            self.DWMWA_USE_IMMERSIVE_DARK_MODE,
            ctypes.byref(value),
            ctypes.sizeof(value)
        )

    def set_title_bar_theme(self, hwnd, theme):
        """Thiết lập theme cho thanh tiêu đề"""
        if not hwnd:
            return
            
        if theme == "dark" and self._set_title_bar_theme_dark:
            self._set_title_bar_theme_dark(hwnd)
        elif theme == "light" and self._set_title_bar_theme_light:
            self._set_title_bar_theme_light(hwnd)

class ThemeManager:
    """Quản lý theme và các tùy chọn giao diện"""
    def __init__(self, main_window, ui):
        self.main_window = main_window
        self.ui = ui
        self._theme = self._detect_system_theme()
        self._toggle_visible = False
        self.title_bar_manager = TitleBarThemeManager()  # Thêm quản lý thanh tiêu đề
        
        # Thêm nút gear và cấu hình ban đầu
        self._setup_gear_button()
        self._initialize_theme_state()
        self._connect_signals()
        
        # Áp dụng theme ban đầu cho thanh tiêu đề
        self._apply_title_bar_theme()

    def _detect_system_theme(self):
        """Phát hiện theme hệ thống"""
        if platform.system() == "Windows":
            try:
                with winreg.OpenKey(
                    winreg.HKEY_CURRENT_USER,
                    r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
                ) as key:
                    apps_use_light_theme, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
                    return "light" if apps_use_light_theme == 1 else "dark"
            except Exception:
                pass
        return "light"

    def _setup_gear_button(self):
        """Cấu hình nút gear"""
        self.gear_btn = GearButton(self.ui.centralwidget)
        self.gear_btn.move(700, 10)
        self.gear_btn.clicked.connect(self.toggle_theme_switch)

    def _initialize_theme_state(self):
        """Khởi tạo trạng thái theme ban đầu"""
        self.ui.theme_toggle.setChecked(self._theme == "dark")
        self.update_theme()
        
        # Ẩn các toggle ban đầu
        self.ui.theme_toggle.hide()
        self.ui.theme_toggle._external_label.hide()
        self.ui.file_mode_toggle.hide()
        self.ui.file_mode_toggle._external_label.hide()

    def _connect_signals(self):
        """Kết nối các signal"""
        self.ui.theme_toggle.toggled.connect(self.on_toggle_changed)
        self.ui.file_mode_toggle.toggled.connect(self.toggle_file_mode)

    def __del__(self):
        """Dọn dẹp khi đối tượng bị hủy"""
        if hasattr(self, '_external_label') and self._external_label:
            self._external_label.deleteLater()

    def toggle_theme_switch(self):
        """Chuyển đổi hiển thị các toggle button"""
        self._toggle_visible = not self._toggle_visible
        for widget in [self.ui.theme_toggle, self.ui.file_mode_toggle]:
            widget.setVisible(self._toggle_visible)
            if hasattr(widget, '_external_label'):
                widget._external_label.setVisible(self._toggle_visible)
            
    def on_toggle_changed(self, checked):
        """Xử lý khi theme toggle thay đổi"""
        self._theme = "dark" if checked else "light"
        self.update_theme()
        
    def toggle_file_mode(self, checked):
        """Chuyển đổi giữa Simple File và MAP File mode"""
        # Hiển thị/ẩn các thành phần tương ứng
        self.ui.input_file_label_2.setVisible(checked)
        self.ui.input_file_entry_2.setVisible(checked)
        self.ui.input_file_button_2.setVisible(checked)
        self.ui.settings_frame.setVisible(not checked)

        # Điều chỉnh vị trí và kích thước các frame
        if checked:  # MAP File mode
            self._adjust_frames_for_map_mode()
        else:  # Simple File mode
            self._adjust_frames_for_simple_mode()

    def _adjust_frames_for_map_mode(self):
        """Điều chỉnh layout cho chế độ MAP File"""
        self.ui.file_type_frame.move(20, 350)
        self.ui.process_frame.move(20, 435)
        self.ui.button_frame.move(170, 520)
        
        self.ui.output_label.move(20, 130)
        self.ui.output_folder_entry.move(140, 120)
        self.ui.output_button.move(620, 120)

        self.ui.inout_frame.setFixedHeight(175)
        self.main_window.resize(740, 605)

        self.ui.input_file_label_1.setText("MAP File:")

    def _adjust_frames_for_simple_mode(self):
        """Điều chỉnh layout cho chế độ Simple File"""
        self.ui.settings_frame.move(20, 300)
        self.ui.file_type_frame.move(20, 415)
        self.ui.process_frame.move(20, 500)
        self.ui.button_frame.move(170, 585)

        self.ui.output_label.move(20, 80)
        self.ui.output_folder_entry.move(140, 70)
        self.ui.output_button.move(620, 70)

        self.ui.inout_frame.setFixedHeight(125)
        self.main_window.resize(740, 670)

        self.ui.input_file_label_1.setText("Simple File:")

    def update_theme(self):
        """Cập nhật theme cho ứng dụng"""
        self.ui.set_theme(self._theme)
        self._apply_title_bar_theme()
    
    def _apply_title_bar_theme(self):
        """Áp dụng theme cho thanh tiêu đề"""
        if not hasattr(self, 'title_bar_manager') or not self.title_bar_manager:
            return
            
        if platform.system() != "Windows":
            return
            
        try:
            # Lấy HWND của cửa sổ
            hwnd = self.main_window.winId()
            if hwnd:
                # Chuyển đổi QWindow thành HWND
                hwnd = int(hwnd)
                self.title_bar_manager.set_title_bar_theme(hwnd, self._theme)
        except Exception as e:
            print(f"Không thể áp dụng theme cho thanh tiêu đề: {e}")

    def handle_click(self, pos):
        """Xử lý sự kiện click chuột để ẩn các toggle"""
        if not self._toggle_visible:
            return
            
        # Kiểm tra xem click có nằm trong vùng toggle hoặc gear button không
        toggle_rect = QtCore.QRect(
            self.ui.theme_toggle.mapToGlobal(QtCore.QPoint(0,0)),
            self.ui.theme_toggle.size()
        )
        gear_rect = QtCore.QRect(
            self.gear_btn.mapToGlobal(QtCore.QPoint(0,0)),
            self.gear_btn.size()
        )
        
        if not toggle_rect.contains(pos) and not gear_rect.contains(pos):
            self.toggle_theme_switch()


class MainWindow(QtWidgets.QMainWindow):
    """Lớp cửa sổ chính của ứng dụng"""
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        
        # Thiết lập icon cho ứng dụng
        self._setup_app_icon()
        
        # Tạo theme manager
        self.theme_manager = ThemeManager(self, self.ui)
        
        # Cài đặt event filter
        self.installEventFilter(self)
        
    def _setup_app_icon(self):
        """Thiết lập icon cho ứng dụng"""
        try:
            # Thử load icon từ file
            app_icon = QtGui.QIcon(icon_app_link)
            if not app_icon.isNull():
                self.setWindowIcon(app_icon)
            else:
                # Fallback: sử dụng icon mặc định từ hệ thống
                fallback_icon = QtGui.QIcon.fromTheme("applications-office")
                self.setWindowIcon(fallback_icon)
        except Exception as e:
            print(f"Không thể tải icon ứng dụng: {e}")
            # Sử dụng icon mặc định nếu có lỗi
            fallback_icon = QtGui.QIcon.fromTheme("applications-office")
            if not fallback_icon.isNull():
                self.setWindowIcon(fallback_icon)

    def showEvent(self, event):
        """Xử lý sự kiện hiển thị cửa sổ"""
        super().showEvent(event)
        # Đợi một chút để cửa sổ hiển thị hoàn toàn trước khi áp dụng theme
        QtCore.QTimer.singleShot(100, self._apply_title_bar_theme)
        
    def _apply_title_bar_theme(self):
        """Áp dụng theme cho thanh tiêu đề"""
        if hasattr(self, 'theme_manager') and hasattr(self.theme_manager, '_apply_title_bar_theme'):
            self.theme_manager._apply_title_bar_theme()
    
    def eventFilter(self, obj, event):
        """Lọc các sự kiện"""
        if event.type() == QtCore.QEvent.MouseButtonPress:
            self.theme_manager.handle_click(event.pos())
        return super().eventFilter(obj, event)


if __name__ == "__main__":
    """Hàm chính khởi chạy ứng dụng"""
    app = QtWidgets.QApplication(sys.argv)
    
    # # Cấu hình style và font
    # app.setStyle("Fusion")
    
    # font = QtGui.QFont()
    # font.setFamily("Segoe UI")
    # font.setPointSize(10)
    # app.setFont(font)
    
    # Tạo và hiển thị cửa sổ chính
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
