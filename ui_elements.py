import os
import sys
import traceback
from pathlib import Path
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QUrl, QPropertyAnimation, QEasingCurve, pyqtProperty, QRectF
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog,
    QTextEdit, QLabel, QSplitter, QGroupBox, QLineEdit, QTabWidget,
    QProgressBar, QHeaderView, QTabBar, QAbstractItemView
)
from PyQt5.QtGui import QDesktopServices, QPainter, QColor

# 导入拆分后的模块
from excel_model import ExcelTableModel, CustomTableView
from file_operations import SearchWorker
from utils import ensure_embedded_excels, resource_path

# -------------------------------------------------
# 滑动TabBar实现
# -------------------------------------------------
class SlidingTabBar(QTabBar):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._current_animated_index = float(self.currentIndex())

        self.animation = QPropertyAnimation(self, b"current_index_animated")
        self.animation.setDuration(300)
        self.animation.setEasingCurve(QEasingCurve.OutCubic)

        self.currentChanged.connect(self._on_tab_changed)

    def _on_tab_changed(self, new_index):
        self.animation.stop()
        self.animation.setStartValue(self._current_animated_index)
        self.animation.setEndValue(float(new_index))
        self.animation.start()

    def paintEvent(self, event):
        super().paintEvent(event)
        
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        current_float_index = self._current_animated_index
        
        left_index = int(current_float_index)
        right_index = left_index + 1

        left_rect = self.tabRect(left_index)
        right_rect = self.tabRect(right_index) if right_index < self.count() else left_rect

        interpolation_factor = current_float_index - left_index

        if left_rect.isValid() and right_rect.isValid():
            interpolated_x = left_rect.x() + (right_rect.x() - left_rect.x()) * interpolation_factor
            interpolated_width = left_rect.width() + (right_rect.width() - left_rect.width()) * interpolation_factor
            
            indicator_height = 3
            interpolated_rect = QRectF(interpolated_x, float(left_rect.height() - indicator_height),
                                         interpolated_width, float(indicator_height))
            
            painter.setBrush(QColor(0, 255, 255))
            painter.setPen(Qt.NoPen)
            painter.drawRect(interpolated_rect.adjusted(5, 0, -5, 0))
        
    def _get_current_index_animated(self):
        return self._current_animated_index

    def _set_current_index_animated(self, index):
        self._current_animated_index = index
        self.update()

    current_index_animated = pyqtProperty(float, _get_current_index_animated, _set_current_index_animated)

# -------------------------------------------------
# UniApp 主应用
# -------------------------------------------------
class UniApp(QWidget):
    SETTINGS_FILE = "last_paths.ini"

    def __init__(self):
        super().__init__()
        self.setWindowTitle('目录管理终端')
        self.resize(1000, 800)
        self.setWindowFlag(Qt.FramelessWindowHint)

        self.excel_le = QLineEdit(self)
        self.target_le = QLineEdit(self)
        self.root_le = QLineEdit(self)

        self.excel_btn = QPushButton('浏览...', self)
        self.target_btn = QPushButton('浏览...', self)
        self.root_btn = QPushButton('浏览...', self)
        self.start_btn = QPushButton('开始执行', self)
        self.create_refresh_excels_btn = QPushButton('创建/刷新 Excel 表', self)
        self.cancel_btn = QPushButton('取消任务', self)

        self.model_origin = ExcelTableModel()
        # --- 关键修改：为更新结果模型添加 is_read_only=True 标志 ---
        self.model_updated = ExcelTableModel(is_read_only=True)
        self.view_origin = CustomTableView(self)
        self.view_updated = CustomTableView(self)

        self.success_edit = QTextEdit(self)
        self.fail_edit = QTextEdit(self)
        
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("准备中... %p%")
        
        self.progress_label = QLabel("当前状态: 等待任务开始...")
        # ================== 核心修改 ==================
        self.progress_label.setWordWrap(True) # 启用自动换行
        # ============================================
        self.progress_label.setAlignment(Qt.AlignCenter)

        self.tab = QTabWidget()
        self.tab.setTabBar(SlidingTabBar())
        
        self.tab.addTab(self._build_work_tab(), "工作终端")
        self.tab.addTab(self._build_excel_tab(self.model_origin, "file_list.xlsx", self.view_origin), "file_list.xlsx")
        self.tab.addTab(self._build_excel_tab(self.model_updated, "file_list_updated.xlsx", self.view_updated), "file_list_updated.xlsx")
        self.tab.addTab(self._build_about_tab(), "关于")
        
        self._build_ui()
        self._signals()
        self._apply_styles()
        self.load_excels()
        self.load_settings()
        self._initial_ui_state()
        
    def _build_work_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)

        form = QGroupBox('路径设置')
        form_layout = QVBoxLayout(form)
        for title, le, btn in [('Excel 列表', self.excel_le, self.excel_btn),
                               ('目标文件夹', self.target_le, self.target_btn),
                               ('查找根目录', self.root_le, self.root_btn)]:
            h_layout = QHBoxLayout()
            h_layout.addWidget(QLabel(title))
            h_layout.addWidget(le)
            h_layout.addWidget(btn)
            form_layout.addLayout(h_layout)
        layout.addWidget(form)

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.create_refresh_excels_btn)
        button_layout.addWidget(self.start_btn)
        button_layout.addWidget(self.cancel_btn)
        layout.addLayout(button_layout)

        layout.addWidget(self.progress_label)
        layout.addWidget(self.progress_bar)

        return widget

    def _build_excel_tab(self, model, title, view_instance):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        view_instance.setModel(model)
        view_instance.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(view_instance)
        
        # --- 仅为可编辑的表格添加保存按钮 ---
        if not model.is_read_only:
            save_btn = QPushButton(f"保存 {title}")
            save_btn.clicked.connect(lambda: self.save_excel(model, title))
            layout.addWidget(save_btn)
        
        return widget

    def _build_about_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        about_text = QTextEdit()
        about_text.setReadOnly(True)
        about_text.setText(
            """
            <h2 style='color:#00FFFF;'>目录管理终端</h2>
            <p style='color:#E0E0E0;'>本工具旨在简化批量文件查找与整理流程，助您高效管理大量文件。</p>
            
            <h3 style='color:#00FFFF;'>使用流程：</h3>
            <ol>
                <li style='color:#E0E0E0;'><b>导入列表：</b> 导入包含目标文件名的 Excel 列表。</li>
                <li style='color:#E0E0E0;'><b>设定路径：</b> 设置文件查找的“根目录”和文件复制的“目标文件夹”。</li>
                <li style='color:#E0E0E0;'><b>一键执行：：</b> 程序将自动搜索并复制文件，同时生成详细的执行报告。</li>
            </ol>
            
            <h3 style='color:#00FFFF;'>特色功能：</h3>
            <ul>
                <li style='color:#E0E0E0;'><b>Excel 驱动：：</b> 通过 Excel 列表进行批量查找与复制，告别手动操作。</li>
                <li style='color:#E0E0E0;'><b>智能匹配：：</b> 支持文件名及“词干”匹配，提高查找成功率。</li>
                <li style='color:#E0E0E0;'><b>实时报告：：</b> 即时查看成功/失败日志，任务完成后自动生成带标记的更新版 Excel 报告。</li>
                <li style='color:#E0E0E0;'><b>内置编辑：：</b> 直接在界面中编辑 Excel 列表，支持复制、粘贴、删除单元格内容。</li>
                <li style='color:#E0E0E0;'><b>提示：</b> 每次输入新表必须点击保存才能被应用</li>
                <li style='color:#E0E0E0;'><b>提示：</b> 首次使用最好点击创建/刷新Excel表</li>
            </ul>
            
            <p style='color:#E0E0E0;'>告别繁琐，让文件管理轻松高效！</p>
            """
        )
        layout.addWidget(about_text)
        
        about_text.setStyleSheet("""
            QTextEdit {
                background: #1A1F36;
                border: 1px solid #00FFFF;
                border-radius: 6px;
                color: #E0E0E0;
                padding: 15px;
                font-family: "Segoe UI", "Microsoft YaHei", "Consolas";
                font-size: 10pt;
            }
        """)
        
        return widget

    def _build_ui(self):
        main = QVBoxLayout(self)

        title_bar = QHBoxLayout()
        title_label = QLabel("目录管理终端")
        title_label.setAlignment(Qt.AlignCenter)
        title_bar.addWidget(title_label)
        title_bar.addStretch()
        close_btn = QPushButton("X")
        close_btn.clicked.connect(self.close)
        title_bar.addWidget(close_btn)
        main.addLayout(title_bar)

        main.addWidget(self.tab)

        log_splitter = QSplitter(Qt.Horizontal)
        log_splitter.addWidget(self._log_group("成功日志", self.success_edit))
        log_splitter.addWidget(self._log_group("失败日志", self.fail_edit))
        main.addWidget(log_splitter)

    def _log_group(self, title, text_edit):
        group = QGroupBox(title)
        layout = QVBoxLayout(group)
        layout.addWidget(text_edit)
        return group

    def _signals(self):
        self.excel_btn.clicked.connect(lambda: self.choose_file(self.excel_le, 'Excel 文件 (*.xlsx)'))
        self.target_btn.clicked.connect(lambda: self.choose_folder(self.target_le))
        self.root_btn.clicked.connect(lambda: self.choose_folder(self.root_le))
        self.start_btn.clicked.connect(self.start_task)
        self.create_refresh_excels_btn.clicked.connect(self._create_and_refresh_excels)
        self.cancel_btn.clicked.connect(self.cancel_task)

    def choose_file(self, line_edit, filt):
        path, _ = QFileDialog.getOpenFileName(self, '选择文件', filter=filt)
        if path:
            line_edit.setText(path)
            self.load_excels()

    def choose_folder(self, line_edit):
        path = QFileDialog.getExistingDirectory(self, '选择文件夹')
        if path:
            line_edit.setText(path)

    def open_excel(self, updated=False):
        excel_path, updated_path = ensure_embedded_excels()
        QDesktopServices.openUrl(QUrl.fromLocalFile(updated_path if updated else excel_path))

    def load_excels(self):
        excel_path, updated_path = ensure_embedded_excels()
        self.model_origin.load(excel_path)
        self.model_updated.load(updated_path)

    def save_excel(self, model, title):
        excel_path, updated_path = ensure_embedded_excels()
        path = excel_path if title == "file_list.xlsx" else updated_path
        model.save(path)
        self.success_edit.append(f"✅ {title} 已保存")

    def _create_and_refresh_excels(self):
        try:
            excel_path, updated_path = ensure_embedded_excels()
            self.load_excels()
            self.success_edit.append("✅ Excel 表格已检测并刷新。")
            
            self.excel_le.setText(excel_path)

        except Exception as e:
            traceback.print_exc()
            self.fail_edit.append(f"❌ 创建/刷新 Excel 表格失败: {e}")

    def _initial_ui_state(self):
        self.start_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.progress_label.setText("当前状态: 等待任务开始...")
        self.progress_bar.setFormat("准备中... %p%")

    def start_task(self):
        excel = self.excel_le.text()
        target = self.target_le.text()
        root = self.root_le.text()
        if not all([excel, target, root]):
            self.fail_edit.append("请确保Excel列表、目标文件夹和查找根目录都已设置！")
            return
        
        self.start_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        
        self.success_edit.clear()
        self.fail_edit.clear()
        self.progress_bar.setValue(0)
        self.progress_label.setText("当前状态: 正在初始化...")

        with open(self.SETTINGS_FILE, "w", encoding="utf-8") as f:
            f.write("\n".join([self.excel_le.text(), self.target_le.text(), self.root_le.text()]))

        self.thread = QThread()
        
        excel_path, updated_excel_path = ensure_embedded_excels()
        self.worker = SearchWorker(excel, target, [root], updated_excel_path)
        
        self.worker.moveToThread(self.thread)
        
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.thread.wait)
        
        self.worker.finished.connect(self.load_excels)
        self.worker.finished.connect(self._on_task_finished)
        
        self.worker.success.connect(lambda s: self.success_edit.append(s))
        self.worker.failed.connect(lambda s: self.fail_edit.append(s))
        self.worker.progress.connect(self.update_progress)

        self.thread.start()

    def cancel_task(self):
        if hasattr(self, 'worker') and self.thread.isRunning():
            self.worker.stop()
            self.cancel_btn.setEnabled(False)
            self.progress_label.setText("当前状态: 正在取消...")
            self.fail_edit.append("用户请求取消任务...")

    def update_progress(self, current, total, message):
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(current)
        
        # ================== 核心修改 ==================
        # 截断路径字符串，以防止窗口被拉伸
        max_len = 50 
        display_message = message
        if len(message) > max_len:
            # 截取前半部分并添加省略号
            display_message = message[:max_len-3] + '...'
            
        self.progress_bar.setFormat(f"{display_message} %p%")
        self.progress_label.setText(f"当前状态: {display_message}")
        # ============================================

    def _on_task_finished(self):
        self.start_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        self.progress_label.setText("当前状态: 任务完成。")
        self.progress_bar.setValue(self.progress_bar.maximum())
        self.progress_bar.setFormat("任务完成！ %p%")

        if hasattr(self, 'thread'):
            del self.thread
        if hasattr(self, 'worker'):
            del self.worker

    def load_settings(self):
        excel_path, _ = ensure_embedded_excels()
        self.excel_le.setText(excel_path)

        try:
            with open(self.SETTINGS_FILE, "r", encoding="utf-8") as f:
                lines = f.read().splitlines()
                if len(lines) >= 3:
                    self.target_le.setText(lines[1])
                    self.root_le.setText(lines[2])
        except FileNotFoundError:
            pass

    def _apply_styles(self):
        self.setStyleSheet("""
/* 全局设置 - 深邃星空主题 */
QWidget {
    background: #0A0A1A; /* 深星空蓝背景 */
    color: #E0E0E0; /* 柔和的白色文本 */
    font-family: "Segoe UI", "Microsoft YaHei", "Consolas";
    font-size: 9pt;
}

/* 标题 - 高亮青色 */
QLabel {
    color: #00FFFF; /* 赛博朋克青色 */
    font-size: 14pt;
    font-weight: bold;
}

/* 按钮 - 霓虹渐变效果 */
QPushButton {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop: 0 #00c6ff, stop: 1 #0078ff);
    border: none;
    border-radius: 6px;
    padding: 6px 18px;
    color: #fff; /* 白色文字以保证对比度 */
    font-weight: bold;
}

QPushButton:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop: 0 #0078ff, stop: 1 #00c6ff);
}

/* 分组框 - 科技感线框 */
QGroupBox {
    border: 1px solid #00FFFF;
    border-radius: 8px;
    margin-top: 6px;
    font-weight: bold;
    color: #00FFFF; /* 标题文字同为青色 */
}

QGroupBox::title {
    subcontrol-origin: margin;
    left: 8px;
    padding: 0 3px 0 3px;
}

/* 输入框 - 终端风格 */
QLineEdit {
    border: 1px solid #00FFFF;
    border-radius: 4px;
    padding: 4px;
    background: #1A1F36; /* 较浅的深蓝背景 */
    color: #00FFFF; /* 青色输入文字 */
}

/* 日志输出框 - 经典绿色荧光字 */
QTextEdit {
    background: #1A1F36;
    border: 1px solid #00FFFF;
    border-radius: 6px;
    color: #39FF14; /* 终端荧光绿 */
    font-family: "Consolas", "Courier New";
}

/* 表格视图 - 数据矩阵风格 */
QTableView {
    background: #1A1F36;
    border: 1px solid #00FFFF;
    border-radius: 6px;
    color: #E0E0E0;
    gridline-color: #007F7F; /* 较暗的青色网格线 */
}

QTableView::item {
    padding: 4px 0;
    border: none;
}

QTableView::item:selected {
    background: #00FFFF; /* 选中时背景为高亮青色 */
    color: #0A0A1A; /* 文字变为深色以保证可读性 */
}

QTableView QTableCornerButton::section {
    background-color: #101428; /* 与表头背景色一致 */
    border: none;
}

/* 表头 - 控制台风格 */
QHeaderView {
    background: #101428; /* 更深的蓝色 */
    color: #E0E0E0;
    border: 1px solid #007F7F;
    border-radius: 4px;
    padding: 4px;
}

QHeaderView::section {
    background: #1A1F36;
    color: #00FFFF;
    border: none;
    padding: 4px;
}

/* 选项卡 - 全息界面风格 */
QTabWidget::pane {
    border: 1px solid #00FFFF; /* 添加边框以包裹内容 */
    border-top-left-radius: 0px;
    border-top-right-radius: 0px;
    border-bottom-left-radius: 8px;
    border-bottom-right-radius: 8px;
    margin-top: -1px; /* 负边距，使内容与tab bar的底部边框对齐 */
    background: #1A1F36; /* 确保内容区域有背景色 */
}

/* 注意：以下 QTabBar::tab 样式将应用于 SlidingTabBar */
QTabBar::tab {
    background: transparent; /* 透明背景 */
    color: #00FFFF;
    padding: 8px 15px; /* 减小左右填充，使页签变窄 */
    margin: 0;
    border: 1px solid #007F7F; /* 暗青色边框 */
    border-bottom: none;
    border-top-left-radius: 4px;
    border-top-right-radius: 4px;
    min-width: 100px; /* 减小最小宽度，使页签变窄 */
    font-size: 10pt; /* 可以适当减小字体大小 */
    font-weight: bold;
}

QTabBar::tab:selected {
    background: #1A1F36; /* 选中时有背景色 */
    color: #00FFFF;
    border: 1px solid #00FFFF; /* 边框变为高亮青色 */
    border-bottom: 1px solid #1A1F36; /* 底部边框与内容区域融合 */
}

QTabBar::tab:!selected:hover {
    color: #FF00FF; /* 未选中项悬停时变为品红色 */
    border-color: #FF00FF;
}
QProgressBar {
    border: 1px solid #00FFFF;
    border-radius: 5px;
    background-color: #1A1F36;
    text-align: center;
    color: #00FFFF;
}

QProgressBar::chunk {
    background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #00FFFF, stop:1 #39FF14);
    border-radius: 5px;
}

QLabel#progress_label {
    color: #E0E0E0;
    font-size: 10pt;
    font-weight: bold;
    padding: 5px;
}
""")
        self.progress_label.setObjectName("progress_label")