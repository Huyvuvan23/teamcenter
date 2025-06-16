# -*- coding: utf-8 -*-

from PyQt5 import QtCore, QtGui, QtWidgets

class GearButton(QtWidgets.QPushButton):
    def __init__(self, parent=None, size=28):
        super().__init__(parent)
        self.setFixedSize(size, size)
        self.setCursor(QtCore.Qt.PointingHandCursor)
        
        # Load the gear icon
        icon = QtGui.QIcon("cogwheel.png")
        self.setIcon(icon)
        self.setIconSize(QtCore.QSize(size-4, size-4))  # Slightly smaller than button size
        
        # Style for transparent background
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

# Insert after import statements
class AnimatedToggle(QtWidgets.QCheckBox):
    def __init__(
        self, parent=None, width=60,
        bg_color="#b0b0b0", circle_color="#fff",
        active_color="#2a82da", label_on="Dark", label_off="Light"
    ):
        super().__init__(parent)
        self.setFixedSize(width, 28)
        self.setCursor(QtCore.Qt.PointingHandCursor)

        # Colors
        self._bg_color = bg_color
        self._circle_color = circle_color
        self._active_color = active_color

        # Labels
        self._label_on = label_on
        self._label_off = label_off

        # External label (below the toggle)
        self._external_label = QtWidgets.QLabel(parent)
        self._external_label.setText(self._label_on if self.isChecked() else self._label_off)
        self._external_label.setFont(QtGui.QFont("Segoe UI", 10, QtGui.QFont.Bold))
        self._external_label.setStyleSheet("color: #222; background: transparent;")
        self._external_label.setAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignTop)
        self._external_label.resize(self.width(), 20)
        self._external_label.move(self.x(), self.y() + self.height() + 2)
        self._external_label.show()

        # Animation
        # Center the bar vertically and the circle inside the bar
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

        # Mouse states
        self._hover = False
        self._pressed = False

        self.toggled.connect(self.start_animation)
        self.toggled.connect(self.update_external_label)
        self.start_animation(self.isChecked())
        self.update_external_label(self.isChecked())

    def moveEvent(self, event):
        if self._external_label:
            self._external_label.move(self.x(), self.y() + self.height() + 2)
        super().moveEvent(event)

    def resizeEvent(self, event):
        if self._external_label:
            self._external_label.resize(self.width(), 20)
        super().resizeEvent(event)

    def update_external_label(self, checked):
        if self._external_label:
            self._external_label.setText(self._label_on if checked else self._label_off)
            if checked:
                self._external_label.setStyleSheet("color: #fff; background: transparent;")
            else:
                self._external_label.setStyleSheet("color: #222; background: transparent;")

    @QtCore.pyqtProperty(int)
    def circlePos(self):
        return self._circle_pos

    @circlePos.setter
    def circlePos(self, pos):
        self._circle_pos = pos
        self.update()

    def start_animation(self, checked):
        self.animation.stop()
        end_pos = self._max_pos if checked else self._min_pos
        self.animation.setStartValue(self._circle_pos)
        self.animation.setEndValue(end_pos)
        self.animation.setEasingCurve(QtCore.QEasingCurve.OutCubic)
        self.animation.start()

    def paintEvent(self, event):
        painter = QtGui.QPainter(self)
        painter.setRenderHint(QtGui.QPainter.Antialiasing)

        # Shadow (short bar)
        shadow_color = QtGui.QColor(0, 0, 0, 30)
        shadow_rect = QtCore.QRectF(self._bar_margin, self._bar_y, self._bar_width, self._bar_height)
        painter.setPen(QtCore.Qt.NoPen)
        painter.setBrush(shadow_color)
        painter.drawRoundedRect(shadow_rect, self._bar_height / 2, self._bar_height / 2)

        # Background (short bar)
        bg_rect = QtCore.QRectF(self._bar_margin, self._bar_y, self._bar_width, self._bar_height)
        bg_color = QtGui.QColor(self._active_color if self.isChecked() else self._bg_color)
        if self._hover:
            bg_color = bg_color.lighter(110)
        if self._pressed:
            bg_color = bg_color.darker(110)
        painter.setBrush(bg_color)
        painter.drawRoundedRect(bg_rect, self._bar_height / 2, self._bar_height / 2)

        # Glow effect when checked
        if self.isChecked():
            glow_color = QtGui.QColor(self._active_color)
            glow_color.setAlpha(80)
            glow_rect = QtCore.QRectF(self._circle_pos - 2, self._circle_y - 2, self._circle_diameter + 4, self._circle_diameter + 4)
            painter.setBrush(glow_color)
            painter.drawEllipse(glow_rect)

        # Circle
        circle_rect = QtCore.QRectF(self._circle_pos, self._circle_y, self._circle_diameter, self._circle_diameter)
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

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setFixedSize(740, 760)
        # Central widget
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        
        # Replace the theme toggle button with this:
        self.theme_toggle = AnimatedToggle(self.centralwidget, width=70)
        self.theme_toggle.setGeometry(QtCore.QRect(625, 10, 70, 28))
        self.theme_toggle.setObjectName("theme_toggle")
        
        # Title label
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(120, 10, 500, 120))
        font = QtGui.QFont()
        font.setFamily("Segoe UI")
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(60)
        self.label.setFont(font)
        self.label.setObjectName("label")
        
        # Input/Output frame
        self.inout_frame = QtWidgets.QFrame(self.centralwidget)
        self.inout_frame.setGeometry(QtCore.QRect(20, 150, 700, 180))
        self.inout_frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.inout_frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.inout_frame.setObjectName("inout_frame")
        
        # Text entries style
        entry_style = """
            QTextEdit, QPlainTextEdit {
                border: 1px solid #ccc;
                border-radius: 4px;
                padding: 5px;
                font-size: 12px;
                font-family: 'Segoe UI';
                selection-background-color: #2a82da;
            }
        """
        
        self.output_folder_entry = QtWidgets.QTextEdit(self.inout_frame)
        self.output_folder_entry.setGeometry(QtCore.QRect(160, 120, 450, 35))
        self.output_folder_entry.setStyleSheet(entry_style)
        self.output_folder_entry.setObjectName("output_folder_entry")
        
        self.input_file_entry_2 = QtWidgets.QTextEdit(self.inout_frame)
        self.input_file_entry_2.setGeometry(QtCore.QRect(160, 70, 450, 35))
        self.input_file_entry_2.setStyleSheet(entry_style)
        self.input_file_entry_2.setObjectName("input_file_entry_2")
        
        self.input_file_entry_1 = QtWidgets.QTextEdit(self.inout_frame)
        self.input_file_entry_1.setGeometry(QtCore.QRect(160, 20, 450, 35))
        self.input_file_entry_1.setStyleSheet(entry_style)
        self.input_file_entry_1.setObjectName("input_file_entry_1")
        
        # Labels
        label_font = QtGui.QFont()
        label_font.setFamily("Segoe UI")
        label_font.setPointSize(10)
        
        self.input_file_label_1 = QtWidgets.QLabel(self.inout_frame)
        self.input_file_label_1.setGeometry(QtCore.QRect(20, 30, 120, 16))
        self.input_file_label_1.setFont(label_font)
        self.input_file_label_1.setObjectName("input_file_label_1")
        self.input_file_label_1.setStyleSheet("border: none;")
        
        self.input_file_label_2 = QtWidgets.QLabel(self.inout_frame)
        self.input_file_label_2.setGeometry(QtCore.QRect(20, 80, 130, 16))
        self.input_file_label_2.setFont(label_font)
        self.input_file_label_2.setObjectName("input_file_label_2")
        self.input_file_label_2.setStyleSheet("border: none;")
        
        self.output_label = QtWidgets.QLabel(self.inout_frame)
        self.output_label.setGeometry(QtCore.QRect(20, 130, 130, 16))
        self.output_label.setFont(label_font)
        self.output_label.setObjectName("output_label")
        self.output_label.setStyleSheet("border: none;")
        
        # Buttons
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
        
        # Settings frame - increased height to fit all elements
        self.settings_frame = QtWidgets.QFrame(self.centralwidget)
        self.settings_frame.setGeometry(QtCore.QRect(20, 350, 700, 150))
        self.settings_frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.settings_frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.settings_frame.setObjectName("settings_frame")
        
        radio_font = QtGui.QFont()
        radio_font.setFamily("Segoe UI")
        radio_font.setPointSize(10)
        
        self.using_simple_file_radio = QtWidgets.QRadioButton(self.settings_frame)
        self.using_simple_file_radio.setGeometry(QtCore.QRect(80, 30, 150, 20))
        self.using_simple_file_radio.setFont(radio_font)
        self.using_simple_file_radio.setObjectName("using_simple_file_radio")
        
        self.using_map_file_radio = QtWidgets.QRadioButton(self.settings_frame)
        self.using_map_file_radio.setGeometry(QtCore.QRect(80, 70, 150, 20))
        self.using_map_file_radio.setFont(radio_font)
        self.using_map_file_radio.setObjectName("using_map_file_radio")
        
        self.folder_column_entry = QtWidgets.QPlainTextEdit(self.settings_frame)
        self.folder_column_entry.setGeometry(QtCore.QRect(550, 20, 80, 30))
        self.folder_column_entry.setStyleSheet(entry_style)
        self.folder_column_entry.setObjectName("folder_column_entry")
        
        self.item_column_entry = QtWidgets.QPlainTextEdit(self.settings_frame)
        self.item_column_entry.setGeometry(QtCore.QRect(550, 60, 80, 30))
        self.item_column_entry.setStyleSheet(entry_style)
        self.item_column_entry.setObjectName("item_column_entry")
        
        self.rev_entry = QtWidgets.QPlainTextEdit(self.settings_frame)
        self.rev_entry.setGeometry(QtCore.QRect(550, 100, 80, 30))
        self.rev_entry.setStyleSheet(entry_style)
        self.rev_entry.setObjectName("rev_entry")
        
        self.folder_column_label = QtWidgets.QLabel(self.settings_frame)
        self.folder_column_label.setGeometry(QtCore.QRect(430, 25, 120, 20))
        self.folder_column_label.setFont(label_font)
        self.folder_column_label.setObjectName("folder_column_label")
        self.folder_column_label.setStyleSheet("border: none;")
        
        self.item_column_label = QtWidgets.QLabel(self.settings_frame)
        self.item_column_label.setGeometry(QtCore.QRect(430, 65, 120, 20))
        self.item_column_label.setFont(label_font)
        self.item_column_label.setObjectName("item_column_label")
        self.item_column_label.setStyleSheet("border: none;")
        
        self.revision_column_label = QtWidgets.QLabel(self.settings_frame)
        self.revision_column_label.setGeometry(QtCore.QRect(430, 105, 120, 20))
        self.revision_column_label.setFont(label_font)
        self.revision_column_label.setObjectName("revision_column_label")
        self.revision_column_label.setStyleSheet("border: none;")
        
        # File type frame - moved down to accommodate larger settings frame
        self.file_type_frame = QtWidgets.QFrame(self.centralwidget)
        self.file_type_frame.setGeometry(QtCore.QRect(20, 520, 700, 60))
        self.file_type_frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.file_type_frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.file_type_frame.setObjectName("file_type_frame")
        
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
        
        # Process frame - moved down
        self.process_frame = QtWidgets.QFrame(self.centralwidget)
        self.process_frame.setGeometry(QtCore.QRect(20, 600, 700, 60))
        self.process_frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.process_frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.process_frame.setObjectName("process_frame")
        
        self.process_label = QtWidgets.QLabel(self.process_frame)
        self.process_label.setGeometry(QtCore.QRect(20, 10, 660, 20))
        self.process_label.setFont(label_font)
        self.process_label.setObjectName("process_label")
        self.process_label.setStyleSheet("border: none;")
        
        self.noti_label = QtWidgets.QLabel(self.process_frame)
        self.noti_label.setGeometry(QtCore.QRect(20, 30, 660, 20))
        self.noti_label.setFont(label_font)
        self.noti_label.setObjectName("noti_label")
        self.noti_label.setStyleSheet("border: none;")
        
        # Button frame - moved down
        self.button_frame = QtWidgets.QFrame(self.centralwidget)
        self.button_frame.setGeometry(QtCore.QRect(170, 680, 400, 60))
        self.button_frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.button_frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.button_frame.setObjectName("button_frame")
        
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
        
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        
        # Set initial theme
        self.current_theme = "dark"
        self.set_theme("dark")

    def set_theme(self, theme):
        """Set light or dark theme"""
        self.current_theme = theme
        
        if theme == "dark":
            # Dark theme
            dark_palette = QtGui.QPalette()
            dark_palette.setColor(QtGui.QPalette.Window, QtGui.QColor(53, 53, 53))
            dark_palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.white)
            dark_palette.setColor(QtGui.QPalette.Base, QtGui.QColor(35, 35, 35))
            dark_palette.setColor(QtGui.QPalette.AlternateBase, QtGui.QColor(53, 53, 53))
            dark_palette.setColor(QtGui.QPalette.ToolTipBase, QtCore.Qt.white)
            dark_palette.setColor(QtGui.QPalette.ToolTipText, QtCore.Qt.white)
            dark_palette.setColor(QtGui.QPalette.Text, QtCore.Qt.white)
            dark_palette.setColor(QtGui.QPalette.Button, QtGui.QColor(53, 53, 53))
            dark_palette.setColor(QtGui.QPalette.ButtonText, QtCore.Qt.white)
            dark_palette.setColor(QtGui.QPalette.BrightText, QtCore.Qt.red)
            dark_palette.setColor(QtGui.QPalette.Link, QtGui.QColor(42, 130, 218))
            dark_palette.setColor(QtGui.QPalette.Highlight, QtGui.QColor(42, 130, 218))
            dark_palette.setColor(QtGui.QPalette.HighlightedText, QtCore.Qt.black)
            
            # Apply palette
            QtWidgets.QApplication.setPalette(dark_palette)
            
            # Additional dark theme styles
            frame_style = """
                QFrame {
                    background-color: #424242;
                    border: 1px solid #555555;
                    border-radius: 8px;
                }
                QLabel {
                    color: #ffffff;
                    border: none;
                }
            """
            
            button_style = """
                QPushButton {
                    background-color: #2a82da;
                    border-radius: 8px;
                    color: white;
                }
                QPushButton:hover {
                    background-color: #3a92ea;
                }
                QPushButton:pressed {
                    background-color: #1a72ca;
                }
                QPushButton:disabled {
                    background-color: #505050;
                }
            """
            
            stop_button_style = """
                QPushButton {
                    background-color: #d84f4f;
                    border-radius: 8px;
                    color: white;
                }
                QPushButton:hover {
                    background-color: #e85f5f;
                }
                QPushButton:pressed {
                    background-color: #c83f3f;
                }
            """
            
            entry_style = """
                QTextEdit, QPlainTextEdit {
                    background-color: #353535;
                    color: #ffffff;
                    border: 1px solid #555555;
                }
                QTextEdit:hover, QPlainTextEdit:hover {
                    border: 1px solid #6a6a6a;
                }
            """
            
            self.theme_toggle.setText("Light Mode")
            
        else:
            # Light theme
            light_palette = QtGui.QPalette()
            light_palette.setColor(QtGui.QPalette.Window, QtGui.QColor(240, 240, 240))
            light_palette.setColor(QtGui.QPalette.WindowText, QtCore.Qt.black)
            light_palette.setColor(QtGui.QPalette.Base, QtGui.QColor(255, 255, 255))
            light_palette.setColor(QtGui.QPalette.AlternateBase, QtGui.QColor(240, 240, 240))
            light_palette.setColor(QtGui.QPalette.ToolTipBase, QtGui.QColor(255, 255, 220))
            light_palette.setColor(QtGui.QPalette.ToolTipText, QtCore.Qt.black)
            light_palette.setColor(QtGui.QPalette.Text, QtCore.Qt.black)
            light_palette.setColor(QtGui.QPalette.Button, QtGui.QColor(240, 240, 240))
            light_palette.setColor(QtGui.QPalette.ButtonText, QtCore.Qt.black)
            light_palette.setColor(QtGui.QPalette.BrightText, QtCore.Qt.red)
            light_palette.setColor(QtGui.QPalette.Link, QtGui.QColor(0, 120, 215))
            light_palette.setColor(QtGui.QPalette.Highlight, QtGui.QColor(0, 120, 215))
            light_palette.setColor(QtGui.QPalette.HighlightedText, QtCore.Qt.white)
            
            # Apply palette
            QtWidgets.QApplication.setPalette(light_palette)
            
            # Additional light theme styles
            frame_style = """
                QFrame {
                    background-color: #ffffff;
                    border: 1px solid #cccccc;
                    border-radius: 8px;
                }
                QLabel {
                    color: #000000;
                    border: none;
                }
            """
            
            button_style = """
                QPushButton {
                    background-color: #2a82da;
                    border-radius: 8px;
                    color: white;
                }
                QPushButton:hover {
                    background-color: #3a92ea;
                }
                QPushButton:pressed {
                    background-color: #1a72ca;
                }
                QPushButton:disabled {
                    background-color: #505050;
                }
            """
            
            stop_button_style = """
                QPushButton {
                    background-color: #ff6b6b;
                    border-radius: 8px;
                    color: white;
                }
                QPushButton:hover {
                    background-color: #ff7b7b;
                }
                QPushButton:pressed {
                    background-color: #ff5b5b;
                }
            """
            
            entry_style = """
                QTextEdit, QPlainTextEdit {
                    background-color: #ffffff;
                    color: #000000;
                    border: 1px solid #cccccc;
                }
                QTextEdit:hover, QPlainTextEdit:hover {
                    border: 1px solid #aaaaaa;
                }
            """
            
            self.theme_toggle.setText("Dark Mode")
        
        # Apply styles
        self.inout_frame.setStyleSheet(frame_style)
        self.settings_frame.setStyleSheet(frame_style)
        self.file_type_frame.setStyleSheet(frame_style)
        self.process_frame.setStyleSheet(frame_style)
        self.button_frame.setStyleSheet(frame_style)
        
        self.input_file_button_1.setStyleSheet(button_style)
        self.input_file_button_2.setStyleSheet(button_style)
        self.output_button.setStyleSheet(button_style)
        self.download_button.setStyleSheet(button_style)
        self.stop_button.setStyleSheet(stop_button_style)
        
        self.output_folder_entry.setStyleSheet(entry_style)
        self.input_file_entry_1.setStyleSheet(entry_style)
        self.input_file_entry_2.setStyleSheet(entry_style)
        self.folder_column_entry.setStyleSheet(entry_style)
        self.item_column_entry.setStyleSheet(entry_style)
        self.rev_entry.setStyleSheet(entry_style)
        
        # Radio buttons and checkboxes
        radio_style = f"""
            QRadioButton {{
                color: {'#ffffff' if theme == 'dark' else '#000000'};
            }}
            QRadioButton::indicator {{
                width: 14px;
                height: 14px;
            }}
            QRadioButton::indicator::unchecked {{
                border: 2px solid {'#aaaaaa' if theme == 'dark' else '#666666'};
                border-radius: 7px;
            }}
            QRadioButton::indicator::checked {{
                border: 2px solid {"#aaaaaa" if theme == 'dark' else '#666666'};
                border-radius: 7px;
                background-color: {"#0078d7" if theme == 'dark' else '#0078d7'};
            }}
        """
        
        checkbox_style = f"""
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
                border: 2px solid {"#aaaaaa" if theme == 'dark' else '#666666'};
                background: {"#0078d7" if theme == 'dark' else '#0078d7'};
            }}
        """
        
        self.using_simple_file_radio.setStyleSheet(radio_style)
        self.using_map_file_radio.setStyleSheet(radio_style)
        self.datanote.setStyleSheet(checkbox_style)
        self.rev_drawing.setStyleSheet(checkbox_style)
        self.nx_pdf.setStyleSheet(checkbox_style)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Teamcenter Downloader"))
        self.label.setText(_translate("MainWindow", "<html><head/><body>\n"
"<p align=\"center\"><span style=\" font-size:18pt; font-weight:600;\">Teamcenter Downloader Tool</span></p>\n"
f"<p align=\"center\">• Log in to Teamcenter before running</span></p>\n"
f"<p align=\"center\">• Switch your keyboard to English mode</span></p>\n"
f"<p align=\"center\">• Do not operate while the tool is running</span></p>\n"
"</body></html>"))
        self.input_file_label_1.setText(_translate("MainWindow", "Input File:"))
        self.input_file_label_2.setText(_translate("MainWindow", "Connector IF file:"))
        self.output_label.setText(_translate("MainWindow", "Output Folder:"))
        self.input_file_button_1.setText(_translate("MainWindow", "Browse"))
        self.input_file_button_2.setText(_translate("MainWindow", "Browse"))
        self.output_button.setText(_translate("MainWindow", "Browse"))
        self.using_simple_file_radio.setText(_translate("MainWindow", "Using Simple File"))
        self.using_map_file_radio.setText(_translate("MainWindow", "Using MAP File"))
        self.folder_column_label.setText(_translate("MainWindow", "Folder Column:"))
        self.item_column_label.setText(_translate("MainWindow", "Item ID Column:"))
        self.revision_column_label.setText(_translate("MainWindow", "Revision Column:"))
        self.datanote.setText(_translate("MainWindow", "Data Note"))
        self.rev_drawing.setText(_translate("MainWindow", "Rev Drawing"))
        self.nx_pdf.setText(_translate("MainWindow", "NX PDF"))
        self.process_label.setText(_translate("MainWindow", "<html><head/><body><p align=\"center\"><span style=\" font-size:9pt;\">Ready</span></p></body></html>"))
        self.noti_label.setText(_translate("MainWindow", "<html><head/><body><p align=\"center\"><span style=\" font-size:9pt;\">Status</span></p></body></html>"))
        self.download_button.setText(_translate("MainWindow", "Download"))
        self.stop_button.setText(_translate("MainWindow", "Stop"))


class ThemeManager:
    def __init__(self, main_window, ui):
        self.main_window = main_window
        self.ui = ui
        # Detect system theme (Windows 10/11)
        self._theme = "light"
        if platform.system() == "Windows":
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize") as key:
                    apps_use_light_theme, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
                    self._theme = "light" if apps_use_light_theme == 1 else "dark"
            except Exception:
                pass
        self._toggle_visible = False
        
        # Add gear button
        self.gear_btn = GearButton(self.ui.centralwidget)
        self.gear_btn.move(700, 10)
        self.gear_btn.clicked.connect(self.toggle_theme_switch)
        
        # Initialize theme state
        self.ui.theme_toggle.setChecked(True if self._theme == "dark" else False)
        self.update_theme()
        
        # Initially hide the theme toggle
        self.ui.theme_toggle.hide()
        self.ui.theme_toggle._external_label.hide()
        
        # Connect signals
        self.ui.theme_toggle.toggled.connect(self.on_toggle_changed)
        
    def toggle_theme_switch(self):
        self._toggle_visible = not self._toggle_visible
        if self._toggle_visible:
            self.ui.theme_toggle.show()
            self.ui.theme_toggle._external_label.show()
        else:
            self.ui.theme_toggle.hide()
            self.ui.theme_toggle._external_label.hide()
            
    def on_toggle_changed(self, checked):
        self._theme = "dark" if checked else "light"
        self.update_theme()
        
    def update_theme(self):
        if self._theme == "dark":
            self.ui.set_theme("dark")
        else:
            self.ui.set_theme("light")
            
    def handle_click(self, pos):
        # Convert toggle position to global coordinates
        toggle_rect = QtCore.QRect(
            self.ui.theme_toggle.mapToGlobal(QtCore.QPoint(0,0)),
            self.ui.theme_toggle.size()
        )
        
        # Convert gear button position to global coordinates 
        gear_rect = QtCore.QRect(
            self.gear_btn.mapToGlobal(QtCore.QPoint(0,0)),
            self.gear_btn.size()
        )
        
        # If click is outside both toggle and gear button, hide the toggle
        if not toggle_rect.contains(pos) and not gear_rect.contains(pos) and self._toggle_visible:
            self.ui.theme_toggle.hide()
            self.ui.theme_toggle._external_label.hide()
            self._toggle_visible = False

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        
        # Create theme manager
        self.theme_manager = ThemeManager(self, self.ui)
        
        # Install event filter for mouse tracking
        self.installEventFilter(self)
        
    def eventFilter(self, obj, event):
        if event.type() == QtCore.QEvent.MouseButtonPress:
            self.theme_manager.handle_click(event.pos())
        return super().eventFilter(obj, event)


if __name__ == "__main__":
    import sys
    import platform
    import winreg
    app = QtWidgets.QApplication(sys.argv)
    
    # Set Fusion style for better appearance
    app.setStyle("Fusion")
    
    # Set default font
    font = QtGui.QFont()
    font.setFamily("Segoe UI")
    font.setPointSize(9)
    app.setFont(font)
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())                                                                                                                                                                        
