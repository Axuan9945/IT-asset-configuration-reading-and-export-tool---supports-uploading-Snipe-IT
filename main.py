# -*- coding: utf-8 -*-

"""
IT 资产信息导出工具 (v2.0)
"""

# 基础库导入
import os
import sys
import wmi
import traceback
import pythoncom
import re
import ctypes
import winreg
import time
import inspect

# UI 主题与风格库
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QHBoxLayout,
                               QVBoxLayout, QPushButton, QFrame, QLabel,
                               QStackedWidget, QTextEdit, QLineEdit, QCheckBox,
                               QFileDialog, QMessageBox, QGridLayout, QProgressBar,
                               QDialog, QListWidget, QDialogButtonBox)
from PySide6.QtCore import QObject, Signal, QThread, Qt, QPropertyAnimation, QEasingCurve, QSize, Property
from PySide6.QtGui import QIcon, QPainter, QColor
import pywinstyles

# Windows API & 打印相关
import win32print

# 本地模块导入
from plugin_manager import PluginManager

# --- 全局定义 ---
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ICON_DIR = os.path.join(SCRIPT_DIR, 'icons')
ILLEGAL_CHARACTERS_RE = re.compile(r'[\x00-\x08\x0B\x0C\x0E-\x1F]')


# --- 核心辅助函数 ---
def is_dark_mode():
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Microsoft\Windows\CurrentVersion\Themes\Personalize')
        value, _ = winreg.QueryValueEx(key, 'AppsUseLightTheme')
        winreg.CloseKey(key)
        return value == 0
    except (FileNotFoundError, OSError):
        return False


def find_next_filename(base_path, base_name="IT资产电脑硬件信息", extension=".xlsx"):
    counter = 1
    while True:
        file_counter = f"{counter:04d}"
        file_name = f"{base_name}-{file_counter}{extension}"
        full_path = os.path.join(base_path, file_name)
        if not os.path.exists(full_path):
            return full_path
        counter += 1


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


# --- 后台工作线程 (最终修正版)---
class Worker(QObject):
    log_message = Signal(str)
    progress_update = Signal(int)
    finished = Signal(object)
    error_message = Signal(str, str)

    def __init__(self, task, *args, **kwargs):
        super().__init__()
        self.task = task
        self.args = args
        self.kwargs = kwargs

    def run(self):
        pythoncom.CoInitialize()
        try:
            # 统一所有任务的调用方式：第一个参数永远是 worker 自身 (self)
            result = self.task(self, *self.args, **self.kwargs)

            self.finished.emit(result)
        except Exception:
            full_traceback = traceback.format_exc()
            self.log_message.emit(f"\n❌ 后台任务发生严重错误。")
            self.log_message.emit(full_traceback)
            self.finished.emit(None)
        finally:
            pythoncom.CoUninitialize()


# --- 打印机选择对话框 ---
class PrinterSelectionDialog(QDialog):
    def __init__(self, printers, parent=None):
        super().__init__(parent)
        self.setWindowTitle("选择打印机")
        self.selected_printer = None
        self.setStyleSheet(parent.styleSheet())
        layout = QVBoxLayout(self)
        self.list_widget = QListWidget()
        self.list_widget.addItems(printers)
        try:
            default_printer = win32print.GetDefaultPrinter()
            items = self.list_widget.findItems(default_printer, Qt.MatchFlag.MatchExactly)
            if items:
                self.list_widget.setCurrentItem(items[0])
        except Exception:
            pass
        self.list_widget.itemDoubleClicked.connect(self.accept)
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(QLabel("请从下方列表选择一台打印机:"))
        layout.addWidget(self.list_widget)
        layout.addWidget(button_box)

    def accept(self):
        item = self.list_widget.currentItem()
        if item: self.selected_printer = item.text()
        super().accept()


# --- 自定义开关控件 ---
class AnimatedSwitch(QWidget):
    toggled = Signal(bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(50, 24)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self._checked = False
        self._offset = 2
        self.radius = 10
        self._on_color = QColor("#4CC2FF")
        self._off_color = QColor("#6A6A6A")
        self._circle_color = QColor("#FFFFFF")
        self.animation = QPropertyAnimation(self, b"offset", self)
        self.animation.setEasingCurve(QEasingCurve.Type.OutBounce)
        self.animation.setDuration(400)

    def isChecked(self):
        return self._checked

    def setChecked(self, checked):
        if self._checked == checked: return
        self._checked = checked
        self.toggled.emit(self._checked)
        self.update_animation()

    def update_animation(self):
        self.animation.stop()
        self.animation.setEndValue(26 if self._checked else 2)
        self.animation.start()

    def _get_offset(self):
        return self._offset

    def _set_offset(self, value):
        self._offset = value; self.update()

    offset = Property(float, _get_offset, _set_offset)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = self.rect()
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(self._on_color if self._checked else self._off_color)
        painter.drawRoundedRect(rect, 12, 12)
        painter.setBrush(self._circle_color)
        painter.drawEllipse(int(self._offset), 2, self.radius * 2, self.radius * 2)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.setChecked(not self.isChecked())
        super().mousePressEvent(event)


# --- QSS 样式表 ---
QSS = {'light': """...""", 'dark': """..."""}  # QSS内容较长，为简洁省略，使用您之前的版本即可


# --- 自定义悬停动画按钮 ---
class HoverAnimatedButton(QPushButton):
    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setObjectName("NavButton")
        self.colors = {
            'light': {'normal': QColor(0, 0, 0, 0), 'hover': QColor(220, 222, 224, 150),
                      'checked': QColor(200, 202, 204, 200)},
            'dark': {'normal': QColor(0, 0, 0, 0), 'hover': QColor(56, 56, 56, 204), 'checked': QColor(70, 70, 70, 230)}
        }
        self.current_theme = 'light'
        self._color = self.colors[self.current_theme]['normal']
        self.animation = QPropertyAnimation(self, b"color", self)
        self.animation.setDuration(200)

    def setColor(self, color):
        self._color = color
        self.setStyleSheet(f"background-color: {self._color.name(QColor.NameFormat.HexArgb)};")

    def color(self):
        return self._color

    color = Property(QColor, color, setColor)

    def setTheme(self, theme):
        self.current_theme = theme; self.updateColor()

    def updateColor(self):
        target_color = self.colors[self.current_theme]['checked'] if self.isChecked() else \
        self.colors[self.current_theme]['normal']
        self.setColor(target_color)

    def enterEvent(self, event):
        if not self.isChecked():
            self.animation.stop()
            self.animation.setEndValue(self.colors[self.current_theme]['hover'])
            self.animation.start()
        super().enterEvent(event)

    def leaveEvent(self, event):
        if not self.isChecked():
            self.animation.stop()
            self.animation.setEndValue(self.colors[self.current_theme]['normal'])
            self.animation.start()
        super().leaveEvent(event)


# --- 主窗口 ---
class MainWindow(QMainWindow):
    def __init__(self, theme):
        super().__init__()
        self.setWindowTitle("IT 资产信息导出工具 (v2.0) Axuan与 Gemini 联合制作")
        self.setGeometry(100, 100, 1100, 800)
        self.current_theme = theme
        self.scanned_data = None
        self.nav_pane_expanded = True
        print("正在初始化插件管理器...")
        self.plugin_manager = PluginManager()
        self.plugin_manager.discover_plugins()
        print("插件加载完成。")
        main_widget = QWidget()
        main_widget.setObjectName("MainWindow")
        self.main_layout = QHBoxLayout(main_widget)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)
        self._load_icons()
        self.nav_pane = self._create_nav_pane()
        self.main_layout.addWidget(self.nav_pane, stretch=0)
        self.content_pane = self._create_content_pane()
        self.main_layout.addWidget(self.content_pane, stretch=1)
        self.setCentralWidget(main_widget)
        self.home_button.setChecked(True)
        self.update_theme()
        self.home_button.updateColor()

    def _load_icons(self):
        self.icons = {}
        icon_paths = {
            "menu": "menu.png", "home": "home.png", "export": "export.png",
            "diagnostics": "ZhengDuan.png", "scan": "search.png",
            "excel": "excel.png", "pdf": "pdf.png", "print": "printer.png",
            "json": "json.png", "csv": "csv.png", "sync": "sync.png"
        }
        if not os.path.isdir(ICON_DIR):
            QMessageBox.critical(self, "图标文件夹缺失", f"错误：找不到 'icons' 文件夹。")
            return
        for name, filename in icon_paths.items():
            path = os.path.join(ICON_DIR, filename)
            if os.path.exists(path):
                self.icons[name] = QIcon(path)

    def _create_nav_pane(self):
        nav_widget = QWidget()
        nav_widget.setObjectName("NavPane")
        self.nav_layout = QVBoxLayout(nav_widget)
        self.nav_layout.setContentsMargins(12, 12, 12, 12)
        self.nav_layout.setSpacing(12)
        header_layout = QHBoxLayout()
        self.hamburger_button = QPushButton()
        self.hamburger_button.setObjectName("NavButton")
        self.hamburger_button.clicked.connect(self.toggle_nav_pane)
        header_layout.addWidget(self.hamburger_button)
        self.app_title = QLabel("资产导出工具")
        self.app_title.setObjectName("AppTitle")
        header_layout.addWidget(self.app_title, stretch=1)
        self.nav_layout.addLayout(header_layout)
        self.home_button = HoverAnimatedButton(" 主页", nav_widget)
        self.diagnostics_button = HoverAnimatedButton(" 系统诊断", nav_widget)
        self.export_button = HoverAnimatedButton(" 报告导出", nav_widget)
        self.nav_buttons = [self.home_button, self.diagnostics_button, self.export_button]
        page_map = {self.home_button: 0, self.diagnostics_button: 1, self.export_button: 2}
        for btn in self.nav_buttons:
            btn.setCheckable(True)
            btn.clicked.connect(lambda checked=False, b=btn: self.stacked_widget.setCurrentIndex(page_map[b]))
            btn.clicked.connect(self.update_nav_selection)
        if self.icons:
            self.hamburger_button.setIcon(self.icons.get("menu"))
            self.hamburger_button.setIconSize(QSize(30, 30))
            self.home_button.setIcon(self.icons.get("home"))
            self.home_button.setIconSize(QSize(24, 24))
            self.diagnostics_button.setIcon(self.icons.get("diagnostics"))
            self.diagnostics_button.setIconSize(QSize(24, 24))
            self.export_button.setIcon(self.icons.get("export"))
            self.export_button.setIconSize(QSize(24, 24))
        self.nav_layout.addWidget(self.home_button)
        self.nav_layout.addWidget(self.diagnostics_button)
        self.nav_layout.addWidget(self.export_button)
        self.nav_layout.addStretch()
        nav_widget.setMinimumWidth(200)
        return nav_widget

    def toggle_nav_pane(self):
        self.nav_pane_expanded = not self.nav_pane_expanded
        end_width = 200 if self.nav_pane_expanded else 60
        self.animation = QPropertyAnimation(self.nav_pane, b"minimumWidth")
        self.animation.setDuration(400)
        self.animation.setStartValue(self.nav_pane.width())
        self.animation.setEndValue(end_width)
        self.animation.setEasingCurve(QEasingCurve.Type.InOutSine)
        self.animation.start()
        self.update_nav_button_text()

    def update_nav_button_text(self):
        if self.nav_pane_expanded:
            self.app_title.show()
            self.home_button.setText(" 主页")
            self.diagnostics_button.setText(" 系统诊断")
            self.export_button.setText(" 报告导出")
        else:
            self.app_title.hide()
            self.home_button.setText("")
            self.diagnostics_button.setText("")
            self.export_button.setText("")

    def _create_content_pane(self):
        content_widget = QWidget()
        content_widget.setObjectName("ContentPane")
        content_layout = QVBoxLayout(content_widget)
        content_layout.setContentsMargins(30, 20, 30, 20)
        self.stacked_widget = QStackedWidget()
        content_layout.addWidget(self.stacked_widget)
        self.stacked_widget.addWidget(self._create_home_page())
        self.stacked_widget.addWidget(self._create_diagnostics_page())
        self.stacked_widget.addWidget(self._create_export_page())
        return content_widget

    def _create_home_page(self):
        page = QWidget()
        page.setObjectName("Page")
        layout = QVBoxLayout(page)
        layout.setSpacing(20)
        title = QLabel("主页与配置")
        title.setObjectName("PageTitle")
        layout.addWidget(title)
        main_horizontal_layout = QHBoxLayout()
        main_horizontal_layout.setSpacing(20)
        left_panel = QVBoxLayout()
        left_panel.setSpacing(20)
        left_panel.setAlignment(Qt.AlignmentFlag.AlignTop)
        config_card = QFrame()
        config_card.setObjectName("Card")
        card_layout = QVBoxLayout(config_card)
        card_layout.setContentsMargins(20, 20, 20, 20)
        card_layout.setSpacing(15)
        header_layout = QHBoxLayout()
        header_label = QLabel("自定义页眉:")
        self.header_edit = QLineEdit()
        header_layout.addWidget(header_label)
        header_layout.addWidget(self.header_edit)
        card_layout.addLayout(header_layout)
        internal_url_layout = QHBoxLayout()
        internal_url_label = QLabel("Snipe-IT 内网URL:")
        self.snipe_internal_url_edit = QLineEdit()
        self.snipe_internal_url_edit.setPlaceholderText("例如: http://192.168.1.100")
        internal_url_layout.addWidget(internal_url_label)
        internal_url_layout.addWidget(self.snipe_internal_url_edit)
        card_layout.addLayout(internal_url_layout)
        external_url_layout = QHBoxLayout()
        external_url_label = QLabel("Snipe-IT 外网URL:")
        self.snipe_external_url_edit = QLineEdit()
        self.snipe_external_url_edit.setPlaceholderText("例如: https://assets.yourcompany.com")
        external_url_layout.addWidget(external_url_label)
        external_url_layout.addWidget(self.snipe_external_url_edit)
        card_layout.addLayout(external_url_layout)
        snipe_key_layout = QHBoxLayout()
        snipe_key_label = QLabel("Snipe-IT API Key:")
        self.snipe_key_edit = QLineEdit()
        self.snipe_key_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.snipe_key_edit.setPlaceholderText("请粘贴您在Snipe-IT生成的API密钥")
        snipe_key_layout.addWidget(snipe_key_label)
        snipe_key_layout.addWidget(self.snipe_key_edit)
        card_layout.addLayout(snipe_key_layout)
        theme_layout = QHBoxLayout()
        self.theme_check = QCheckBox("暗黑模式")
        self.theme_check.setChecked(self.current_theme == 'dark')
        self.theme_check.stateChanged.connect(self.toggle_theme)
        theme_layout.addStretch()
        theme_layout.addWidget(self.theme_check)
        card_layout.addLayout(theme_layout)
        left_panel.addWidget(config_card)
        plugins_card = QFrame()
        plugins_card.setObjectName("Card")
        plugins_layout = QVBoxLayout(plugins_card)
        plugins_layout.setContentsMargins(20, 20, 20, 20)
        plugins_layout.setSpacing(10)
        plugins_layout.addWidget(QLabel("<b>选择要扫描的模块:</b>"))
        self.scan_checkboxes = []
        for plugin in self.plugin_manager.scan_plugins:
            plugin_row = QHBoxLayout()
            plugin_name_label = QLabel(getattr(plugin, 'name', '未命名插件'))
            switch_btn = AnimatedSwitch()
            switch_btn.setChecked(True)
            plugin_row.addWidget(plugin_name_label)
            plugin_row.addStretch()
            plugin_row.addWidget(switch_btn)
            plugins_layout.addLayout(plugin_row)
            self.scan_checkboxes.append((switch_btn, plugin))
        left_panel.addWidget(plugins_card)
        action_card = QFrame()
        action_card.setObjectName("Card")
        card_layout_action = QVBoxLayout(action_card)
        card_layout_action.setContentsMargins(20, 20, 20, 20)
        card_layout_action.setSpacing(15)
        self.scan_button = QPushButton("开始扫描硬件信息")
        self.scan_button.clicked.connect(self.start_scan)
        if self.icons: self.scan_button.setIcon(self.icons.get("scan")); self.scan_button.setIconSize(QSize(20, 20))
        card_layout_action.addWidget(self.scan_button)
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(False)
        card_layout_action.addWidget(self.progress_bar)
        left_panel.addWidget(action_card)
        main_horizontal_layout.addLayout(left_panel, 1)
        right_panel = QVBoxLayout()
        right_panel.setSpacing(10)
        right_panel.setAlignment(Qt.AlignmentFlag.AlignTop)
        log_card = QFrame()
        log_card.setObjectName("Card")
        log_card_layout = QVBoxLayout(log_card)
        log_card_layout.setContentsMargins(20, 20, 20, 20)
        log_label = QLabel("日志输出")
        log_label.setObjectName("LogTitle")
        log_card_layout.addWidget(log_label)
        self.log_text_edit = QTextEdit()
        self.log_text_edit.setReadOnly(True)
        log_card_layout.addWidget(self.log_text_edit, 1)
        signature_label = QLabel("Axuan与 Gemini 联合制作")
        signature_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignBottom)
        signature_label.setStyleSheet("color: gray; font-size: 12px; margin-top: 10px;")
        log_card_layout.addWidget(signature_label)
        right_panel.addWidget(log_card, 1)
        main_horizontal_layout.addLayout(right_panel, 1)
        layout.addLayout(main_horizontal_layout, 1)
        return page

    def _create_diagnostics_page(self):
        page = QWidget()
        page.setObjectName("Page")
        layout = QVBoxLayout(page)
        layout.setSpacing(20)
        title = QLabel("系统诊断")
        title.setObjectName("PageTitle")
        layout.addWidget(title)
        card = QFrame()
        card.setObjectName("Card")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(20, 20, 20, 20)
        card_layout.setSpacing(15)
        self.start_diag_button = QPushButton("开始系统诊断")
        self.start_diag_button.clicked.connect(self.start_diagnostics)
        if self.icons: self.start_diag_button.setIcon(
            self.icons.get("diagnostics")); self.start_diag_button.setIconSize(QSize(20, 20))
        card_layout.addWidget(self.start_diag_button)
        self.diag_results_edit = QTextEdit()
        self.diag_results_edit.setReadOnly(True)
        card_layout.addWidget(self.diag_results_edit, 1)
        layout.addWidget(card, 1)
        return page

    def _create_export_page(self):
        page = QWidget()
        page.setObjectName("Page")
        page_layout = QVBoxLayout(page)
        page_layout.setContentsMargins(20, 0, 20, 0)
        title = QLabel("报告导出与同步")
        title.setObjectName("PageTitle")
        title.setAlignment(Qt.AlignmentFlag.AlignLeft)
        page_layout.addWidget(title)
        card = QFrame()
        card.setObjectName("Card")
        self.card_layout = QGridLayout(card)
        self.card_layout.setContentsMargins(30, 30, 30, 30)
        self.card_layout.setSpacing(20)
        self.export_buttons = {}
        row, col = 0, 0
        for plugin in self.plugin_manager.export_plugins:
            button = QPushButton(getattr(plugin, 'name', '未命名插件'))
            button.setEnabled(self.scanned_data is not None)
            button.setMinimumHeight(60)
            button.clicked.connect(lambda checked=False, p=plugin: self.start_export(p))
            if self.icons and hasattr(plugin, 'icon_name') and plugin.icon_name in self.icons:
                button.setIcon(self.icons[plugin.icon_name])
                button.setIconSize(QSize(24, 24))
            self.card_layout.addWidget(button, row, col)
            self.export_buttons[plugin.name] = button
            col += 1
            if col % 2 == 0: row, col = row + 1, 0
        sync_plugins = self.plugin_manager.get_sync_plugins()
        if sync_plugins:
            self.sync_button = QPushButton(getattr(sync_plugins[0], 'name', '未命名插件'))
            self.sync_button.setEnabled(self.scanned_data is not None)
            self.sync_button.setMinimumHeight(60)
            self.sync_button.clicked.connect(self.start_sync)
            if self.icons and hasattr(sync_plugins[0], 'icon_name') and sync_plugins[0].icon_name in self.icons:
                self.sync_button.setIcon(self.icons[sync_plugins[0].icon_name])
                self.sync_button.setIconSize(QSize(24, 24))
            if col != 0: row, col = row + 1, 0
            self.card_layout.addWidget(self.sync_button, row, col, 1, 2)
        page_layout.addStretch(1)
        page_layout.addWidget(card)
        page_layout.addStretch(1)
        return page

    def update_nav_selection(self):
        sender = self.sender()
        for btn in self.nav_buttons:
            if btn != sender: btn.setChecked(False)
        for btn in self.nav_buttons:
            if isinstance(btn, HoverAnimatedButton): btn.updateColor()

    def update_log(self, message):
        self.log_text_edit.append(message)

    def show_error_message(self, title, text):
        QMessageBox.critical(self, title, text)

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def set_buttons_state(self, is_enabled):
        self.scan_button.setEnabled(is_enabled)
        self.start_diag_button.setEnabled(is_enabled)
        for button in self.export_buttons.values():
            button.setEnabled(is_enabled and self.scanned_data is not None)
        if hasattr(self, 'sync_button'):
            self.sync_button.setEnabled(is_enabled and self.scanned_data is not None)

    def toggle_theme(self):
        self.current_theme = 'dark' if self.theme_check.isChecked() else 'light'
        self.update_theme()

    def update_theme(self):
        self.setStyleSheet(QSS[self.current_theme])
        for btn in self.nav_buttons:
            if isinstance(btn, HoverAnimatedButton):
                btn.setTheme(self.current_theme)
        pywinstyles.apply_style(self, "aero")
        header_color = "#202020" if self.current_theme == 'dark' else "#f3f3f3"
        pywinstyles.change_header_color(self, header_color)

    def start_task(self, task, on_finish, *args, **kwargs):
        self.thread = QThread()
        self.worker = Worker(task, *args, **kwargs)
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(on_finish)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.log_message.connect(self.update_log)
        self.worker.error_message.connect(self.show_error_message)
        self.worker.progress_update.connect(self.update_progress)
        self.thread.start()

    def start_scan(self):
        self.set_buttons_state(False)
        self.scan_button.setText("正在扫描中...")
        self.scanned_data = None
        self.log_text_edit.clear()
        self.stacked_widget.setCurrentIndex(0)
        self.home_button.setChecked(True)
        self.update_nav_selection()
        self.update_log("--- 扫描任务开始 ---")
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        selected_plugins = [p for btn, p in self.scan_checkboxes if btn.isChecked()]
        if not selected_plugins:
            self.update_log("⚠️ 未选择任何扫描模块，任务已取消。")
            self.set_buttons_state(True)
            self.scan_button.setText("开始扫描硬件信息")
            self.progress_bar.setVisible(False)
            return
        self.start_task(_scan_worker_task_plugin, self._scan_finished, selected_plugins)

    def _scan_finished(self, scanned_data):
        self.scanned_data = scanned_data
        self.scan_button.setText("重新扫描硬件信息")
        self.progress_bar.setVisible(False)
        self.set_buttons_state(True)
        if self.scanned_data:
            self.stacked_widget.setCurrentIndex(2)
            self.export_button.setChecked(True)
            self.update_nav_selection()

    def start_sync(self):
        if not self.scanned_data:
            QMessageBox.warning(self, "无数据", "请先扫描硬件信息后再同步。")
            return
        sync_plugins = self.plugin_manager.get_sync_plugins()
        if not sync_plugins:
            QMessageBox.critical(self, "错误", "未找到同步插件。")
            return
        config = {
            'internal_url': self.snipe_internal_url_edit.text().strip(),
            'external_url': self.snipe_external_url_edit.text().strip(),
            'key': self.snipe_key_edit.text().strip()
        }
        print(f"--- 调试信息: 准备传递给插件的配置 ---\n{config}\n------------------------------------")
        self.start_task(sync_plugins[0].sync, self._sync_finished, self.scanned_data, config)

    def _sync_finished(self, result):
        self.update_log("--- 同步任务完成 ---")
        QMessageBox.information(self, "完成", "同步任务已执行，请检查主页日志输出获取详细信息。")

    def start_diagnostics(self):
        self.set_buttons_state(False)
        self.start_diag_button.setText("正在诊断中...")
        self.diag_results_edit.clear()
        self.update_log("\n--- 系统诊断任务开始 ---")
        diag_plugins = self.plugin_manager.get_diagnostic_plugins()
        if not diag_plugins:
            self.update_log("⚠️ 未找到任何诊断模块。")
            self.diag_results_edit.setText("未找到任何诊断模块。")
            self.set_buttons_state(True)
            self.start_diag_button.setText("开始系统诊断")
            return
        self.start_task(_diagnostics_worker_task, self._diagnostics_finished, diag_plugins)

    def _diagnostics_finished(self, results):
        if results:
            report = []
            for plugin_name, result_list in results.items():
                report.append(f"▶️ 插件: {plugin_name}\n" + "=" * 40)
                for res in result_list:
                    report.append(
                        f"  - 任务: {res.get('task', 'N/A')}\n"
                        f"    状态: {res.get('status', 'N/A')}\n"
                        f"    信息: {res.get('message', 'N/A')}\n"
                    )
            self.diag_results_edit.setText("\n".join(report))
        self.update_log("✅ 系统诊断完成。")
        self.set_buttons_state(True)
        self.start_diag_button.setText("重新开始诊断")

    def start_export(self, plugin):
        if self.scanned_data is None:
            QMessageBox.warning(self, "无数据", "请先扫描硬件信息后再导出。")
            return
        if plugin.name == "打印报告":
            self.set_buttons_state(False)
            self.update_log("\n正在准备打印...")
            try:
                printers = [p[2] for p in win32print.EnumPrinters(2)]
                if not printers:
                    QMessageBox.critical(self, "打印错误", "系统中没有找到任何打印机。")
                    self.set_buttons_state(True)
                    return
            except Exception as e:
                QMessageBox.critical(self, "打印错误", f"无法获取打印机列表。\n\n错误: {e}")
                self.set_buttons_state(True)
                return
            dialog = PrinterSelectionDialog(printers, self)
            if dialog.exec():
                if dialog.selected_printer:
                    self.update_log(f"  -> 用户选择了打印机: {dialog.selected_printer}")
                    self.update_log("  -> 正在生成打印预览文件...")
                    self.start_task(_export_worker_task, self._save_finished, plugin, self.scanned_data, "",
                                    self.header_edit.text(), dialog.selected_printer)
                else:
                    self.update_log("  -> 未选择打印机，打印取消。")
                    self.set_buttons_state(True)
            else:
                self.update_log("  -> 用户取消了打印。")
                self.set_buttons_state(True)
        else:
            desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
            suggested_path = find_next_filename(desktop_path, "IT资产电脑硬件信息",
                                                getattr(plugin, 'file_extension', '.txt'))
            file_path, _ = QFileDialog.getSaveFileName(self, f"保存为 {getattr(plugin, 'name', '未知插件')}",
                                                       os.path.join(os.path.expanduser('~'),
                                                                    os.path.basename(suggested_path)),
                                                       getattr(plugin, 'file_filter', 'All Files (*)'))
            if file_path:
                self.set_buttons_state(False)
                self.update_log(f"\n正在使用 '{plugin.name}' 插件导出: {os.path.basename(file_path)}")
                self.start_task(_export_worker_task, self._save_finished, plugin, self.scanned_data, file_path,
                                self.header_edit.text(), None)

    def _save_finished(self, result):
        if isinstance(result, dict) and result.get("action") == "print":
            self.update_log("✅ 打印预览已在默认PDF阅读器中打开。")
            self.show_error_message("请手动打印",
                                    "报告已在您的默认PDF阅读器中打开。\n\n请在该程序中使用打印功能 (通常是按 Ctrl+P) 来完成打印。")
            cleanup_worker = CleanupWorker(result.get("path"), parent=self)
            cleanup_worker.log_message.connect(self.update_log)
            cleanup_worker.start()
        elif isinstance(result, str) and (os.path.exists(result) or "失败" in result):
            if "失败" in result:
                self.update_log(f"❌ 操作失败。返回信息: {result}")
            else:
                self.update_log(f"✅ 操作成功！\n文件路径: {result}")
        else:
            self.update_log(f"❌ 操作失败。返回了未知结果: {result}")
        self.set_buttons_state(True)


# --- 打印后清理文件的线程 ---
class CleanupWorker(QThread):
    log_message = Signal(str)

    def __init__(self, path, parent=None):
        super().__init__(parent); self.path = path

    def run(self):
        self.sleep(15)
        try:
            if self.path and os.path.exists(self.path):
                os.remove(self.path)
                self.log_message.emit("  -> 已清理临时打印文件。")
        except Exception as e:
            self.log_message.emit(f"  -> 清理临时文件失败: {e}")


# --- 后台任务函数 (全局) ---
def _scan_worker_task_plugin(worker, scan_plugins):
    log_signal, progress_signal = worker.log_message.emit, worker.progress_update.emit
    hardware_data = []
    total_steps = len(scan_plugins)
    current_step = 0
    if total_steps == 0: return []
    log_signal("正在连接 WMI 核心服务...")
    try:
        c_wmi = wmi.WMI()
    except Exception as e:
        log_signal(f"❌ WMI 连接失败: {e}"); return None
    for plugin in scan_plugins:
        log_signal(f"--- 正在扫描: {getattr(plugin, 'name', '未命名插件')} ---")
        try:
            result = plugin.scan(c_wmi)
            if result:
                hardware_data.extend(result); log_signal(f"✅ 模块 '{plugin.name}' 扫描成功。")
            else:
                log_signal(f"⚠️ 模块 '{plugin.name}' 扫描完成，但未返回数据。")
        except Exception as e:
            log_signal(f"❌ 模块 '{plugin.name}' 扫描失败: {e}")
        current_step += 1
        progress_signal(int((current_step / total_steps) * 100))
    progress_signal(100)
    return hardware_data


def _diagnostics_worker_task(worker, diag_plugins):
    log_signal = worker.log_message.emit
    all_results = {}
    for plugin in diag_plugins:
        plugin_name = getattr(plugin, 'name', '未命名插件')
        log_signal(f"--- 正在运行诊断: {plugin_name} ---")
        try:
            results = plugin.run_diagnostic()
            all_results[plugin_name] = results
            log_signal(f"✅ 模块 '{plugin_name}' 诊断完成。")
        except Exception as e:
            log_signal(f"❌ 模块 '{plugin.name}' 诊断失败: {e}")
            all_results[plugin_name] = [{'task': '插件执行', 'status': '错误', 'message': str(e)}]
    return all_results


def _export_worker_task(worker, plugin, data, file_path, header_text, printer_name):
    result = None
    try:
        worker.log_message.emit(f"  -> 正在尝试调用插件 '{getattr(plugin, 'name', '未命名插件')}'...")
        try:
            result = plugin.export(data, file_path, header_text, worker.log_message.emit, printer_name)
        except TypeError:
            try:
                result = plugin.export(data, file_path, header_text, worker.log_message.emit)
            except TypeError:
                try:
                    result = plugin.export(data, file_path, header_text)
                except TypeError as final_e:
                    worker.log_message.emit(
                        f"❌ 插件 '{getattr(plugin, 'name', '未知插件')}' 的调用失败，参数不匹配。最终错误: {final_e}")
    except Exception as general_e:
        worker.log_message.emit(f"❌ 插件 {getattr(plugin, 'name', '未知插件')} 导出时发生内部错误: {general_e}")
        worker.log_message.emit(traceback.format_exc())
    return result


# --- 程序入口 ---
if __name__ == "__main__":
    app = QApplication(sys.argv)
    if not is_admin():
        msg_box = QMessageBox(QMessageBox.Icon.Warning, "权限提示",
                              "建议以管理员身份运行以获取最完整的硬件信息（如序列号等）。\n\n是否继续？",
                              QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if msg_box.exec() == QMessageBox.StandardButton.No: sys.exit()
    theme = 'dark' if is_dark_mode() else 'light'
    window = MainWindow(theme)
    window.show()
    sys.exit(app.exec())