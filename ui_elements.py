"""
ui_elements.py

该模块包含应用程序的用户界面元素和逻辑，负责处理用户交互、
布局、样式和与后台任务的通信。
"""
import os
import sys
import traceback
from pathlib import Path
from PyQt5.QtCore import Qt, QThread, QUrl, QPropertyAnimation, QEasingCurve, pyqtProperty, QRectF, QTimer
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog,
    QTextEdit, QLabel, QSplitter, QGroupBox, QLineEdit, QTabWidget,
    QProgressBar, QHeaderView, QTabBar, QAbstractItemView, QComboBox, QApplication
)
from PyQt5.QtGui import QDesktopServices, QPainter, QColor, QIcon, QFontMetrics
from excel_model import ExcelTableModel, CustomTableView
from file_operations import SearchWorker, resource_path
from utils import setup_excel_files
import json

# -------------------------------------------------
# 翻译字典
# -------------------------------------------------
TRANSLATIONS = {
    'zh': {
        'work_terminal': '工作终端',
        'target_files': '目标文件列',
        'execution_results': '执行结果列',
        'settings': '设置',
        'about': '关于',
        'title': '目录管理终端',
        'path_settings': '路径设置',
        'excel_list': 'Excel 列表',
        'target_directory': '目标文件夹',
        'search_root': '查找根目录',
        'browse': '浏览...',
        'create_refresh': '创建/刷新 Excel 表',
        'start': '开始执行',
        'cancel': '取消任务',
        'match_settings': '匹配设置',
        'match_mode': '匹配模式:',
        'exact_match': '精确匹配 (包含)',
        'fuzzy_match': '模糊匹配 (85%)',
        'regex_match': '正则表达式',
        'status_waiting': '当前状态: 等待任务开始...',
        'status_initializing': '当前状态: 正在初始化...',
        'preparing': '准备中... %p%',
        'file_preview': '文件预览: ',
        'save': '保存',
        'add_row': '新增一行',
        'success_log': '成功日志',
        'failure_log': '失败日志',
        'task_completed_msg': '任务完成。',
        'task_completed': '任务完成！ %p%',
        'path_not_set_error': '请确保Excel列表、目标文件夹和查找根目录都已设置！',
        'excel_refreshed_success': 'Excel 表格已检测并刷新。',
        'excel_refresh_fail': '创建/刷新 Excel 表格失败: ',
        'user_cancel': '用户请求取消任务...',
        'language_settings': '语言设置:',
        'chinese': '简体中文',
        'english': 'English',
        'about_text': """
            <h2 style='color:#00FFFF;'>目录管理终端</h2>
            <p style='color:#E0E0E0;'>本工具旨在简化批量文件查找与整理流程，助您高效管理大量文件。</p>
            <h3 style='color:#00FFFF;'>使用流程：</h3>
            <ol>
                <li style='color:#E0E0E0;'><b>导入列表：</b> 导入包含目标文件名的 Excel 列表。</li>
                <li style='color:#E0E0E0;'><b>设定路径：：</b> 设置文件查找的"根目录"和文件复制的"目标文件夹"。</li>
                <li style='color:#E0E0E0;'><b>一键执行：：</b> 程序将自动搜索并复制文件，同时生成详细的执行报告。</li>
            </ol>
            <h3 style='color:#00FFFF;'>特色功能：：</h3>
            <ul>
                <li style='color:#E0E0E0;'><b>Excel 驱动：：</b> 通过 Excel 列表进行批量查找与复制，告别手动操作。</li>
                <li style='color:#E0E0E0;'><b>智能匹配：：</b> **支持精确、模糊和正则表达式三种匹配模式，提高查找成功率。**</li>
                <li style='color:#E0E0E0;'><b>实时报告：：</b> 即时查看成功/失败日志，任务完成后自动生成带标记的更新版 Excel 报告。</li>
                <li style='color:#E0E0E0;'><b>内置编辑：：</b> 直接在界面中编辑 Excel 列表，支持复制、粘贴、删除单元格内容。</li>
                <li style='color:#E0E0E0;'><b>提示：</b> 每次输入新表必须点击保存才能被应用</li>
                <li style='color:#E0E0E0;'><b>提示：</b> 首次使用最好点击创建/刷新Excel表</li>
            </ul>
            <p style='color:#E0E0E0;'>告别繁琐，让文件管理轻松高效！</p>
        """,
        'language_changed': '语言设置已更改为简体中文',
        'status_prefix': '当前状态:',
    },
    'en': {
        'work_terminal': 'Work Terminal',
        'target_files': 'Target Files',
        'execution_results': 'Execution Results',
        'settings': 'Settings',
        'about': 'About',
        'title': 'Directory Management Terminal',
        'path_settings': 'Path Settings',
        'excel_list': 'Excel List',
        'target_directory': 'Target Directory',
        'search_root': 'Search Root',
        'browse': 'Browse...',
        'create_refresh': 'Create/Refresh Excel',
        'start': 'Start',
        'cancel': 'Cancel',
        'match_settings': 'Match Settings',
        'match_mode': 'Match Mode:',
        'exact_match': 'Exact Match (Contains)',
        'fuzzy_match': 'Fuzzy Match (85%)',
        'regex_match': 'Regex',
        'status_waiting': 'Status: Waiting to start...',
        'status_initializing': 'Status: Initializing...',
        'preparing': 'Preparing... %p%',
        'file_preview': 'File Preview: ',
        'save': 'Save',
        'add_row': 'Add Row',
        'success_log': 'Success Log',
        'failure_log': 'Failure Log',
        'task_completed_msg': 'Task completed.',
        'task_completed': 'Task Completed! %p%',
        'path_not_set_error': 'Please ensure Excel list, target folder and search root are all set!',
        'excel_refreshed_success': 'Excel tables detected and refreshed.',
        'excel_refresh_fail': 'Failed to create/refresh Excel tables: ',
        'user_cancel': 'User requested to cancel task...',
        'language_settings': 'Language Settings:',
        'chinese': '简体中文',
        'english': 'English',
        'about_text': """
            <h2 style='color:#00FFFF;'>Directory Management Terminal</h2>
            <p style='color:#E0E0E0;'>This tool is designed to simplify the process of bulk file searching and organization, helping you manage large numbers of files efficiently.</p>
            <h3 style='color:#00FFFF;'>Usage Flow:</h3>
            <ol>
                <li style='color:#E0E0E0;'><b>Import List:</b> Import an Excel list containing the filenames to be searched.</li>
                <li style='color:#E0E0E0;'><b>Set Paths:</b> Define the 'Search Root' directory for finding files and the 'Target Directory' for copying them.</li>
                <li style='color:#E0E0E0;'><b>One-Click Execution:</b> The program will automatically search, copy files, and generate a detailed execution report.</li>
            </ol>
            <h3 style='color:#00FFFF;'>Features:</h3>
            <ul>
                <li style='color:#E0E0E0;'><b>Excel-Driven:</b> Bulk search and copy files using an Excel list, eliminating manual operations.</li>
                <li style='color:#E0E0E0;'><b>Smart Matching:</b> **Supports Exact, Fuzzy, and Regex matching modes to increase success rates.**</li>
                <li style='color:#E0E0E0;'><b>Real-time Reporting:</b> View success/failure logs instantly. An updated Excel report with markings is generated automatically upon completion.</li>
                <li style='color:#E0E0E0;'><b>In-App Editing:</b> Edit Excel lists directly in the interface, with support for copy, paste, and deletion of cell content.</li>
                <li style='color:#E0E0E0;'><b>Note:</b> You must click Save to apply any changes to the Excel sheet.</li>
                <li style='color:#E0E0E0;'><b>Note:</b> It is recommended to click Create/Refresh Excel Table for first-time use.</li>
            </ul>
            <p style='color:#E0E0E0;'>Say goodbye to tedious tasks and hello to efficient file management!</p>
        """,
        'language_changed': 'Language settings have been changed to English',
        'status_prefix': 'Status:',
    }
}


def get_translation(key, language):
    """翻译辅助函数"""
    return TRANSLATIONS.get(language, TRANSLATIONS['zh']).get(key, key)


# -------------------------------------------------
# 滑动TabBar实现
# -------------------------------------------------
class SlidingTabBar(QTabBar):
    """一个自定义的 QTabBar，增加了滑动动画效果和动态光标。"""
    def __init__(self, parent=None):
        """初始化TabBar。"""
        super().__init__(parent)
        self._current_animated_index = float(self.currentIndex())

        self.animation = QPropertyAnimation(self, b"current_index_animated")
        self.animation.setDuration(300)
        self.animation.setEasingCurve(QEasingCurve.OutCubic)
        self.currentChanged.connect(self._on_tab_changed)

    def _on_tab_changed(self, new_index):
        """当页签改变时启动动画。"""
        self.animation.stop()
        self.animation.setStartValue(self._current_animated_index)
        self.animation.setEndValue(float(new_index))
        self.animation.start()

    def paintEvent(self, event):
        """重写绘制事件以实现动画和动态光标效果。"""
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
            
            # --- 核心改动：动态计算光标宽度 ---
            font_metrics = QFontMetrics(self.font())
            current_tab_text = self.tabText(int(self._current_animated_index))
            text_width = font_metrics.boundingRect(current_tab_text).width()
            
            # 增加一些内边距，确保光标比文字稍宽
            indicator_width = text_width + 10  
            
            # 确保光标在当前页签下方居中
            indicator_x = interpolated_x + (interpolated_width - indicator_width) / 2
            indicator_y = float(left_rect.height() - indicator_height)

            painter.setBrush(QColor(0, 255, 255))
            painter.setPen(Qt.NoPen)
            # --- 关键修正：将所有参数都转换为整数 ---
            painter.drawRect(int(indicator_x), int(indicator_y), int(indicator_width), int(indicator_height))
    
    def _get_current_index_animated(self):
        """获取当前动画索引。"""
        return self._current_animated_index

    def _set_current_index_animated(self, index):
        """设置当前动画索引。"""
        self._current_animated_index = index
        self.update()

    current_index_animated = pyqtProperty(float, _get_current_index_animated, _set_current_index_animated)

# -------------------------------------------------
# UniApp 主应用
# -------------------------------------------------
class UniApp(QWidget):
    """
    主应用程序窗口，包含所有 UI 元素和逻辑。
    """
    SETTINGS_FILE = "last_paths.ini"
    CONFIG_FILE = "settings.json"

    def __init__(self):
        """初始化应用程序主窗口。"""
        super().__init__()
        self._language = 'zh' # 默认语言
        self.load_settings()

        self.setWindowTitle(get_translation('title', self._language))
        self.resize(1000, 800)
        self.setWindowIcon(QIcon(resource_path('resources/icon.ico'))) # 设置窗口图标
        self.setWindowFlag(Qt.FramelessWindowHint)

        # 初始化所有 UI 组件
        self.excel_le = QLineEdit(self)
        self.target_le = QLineEdit(self)
        self.root_le = QLineEdit(self)
        self.excel_btn = QPushButton(self)
        self.target_btn = QPushButton(self)
        self.root_btn = QPushButton(self)
        self.start_btn = QPushButton(self)
        self.create_refresh_excels_btn = QPushButton(self)
        self.cancel_btn = QPushButton(self)
        self.match_mode_combo = QComboBox(self)
        
        self.tab_work_label = QLabel()
        self.tab_excel_label = QLabel()
        self.tab_updated_label = QLabel()
        self.tab_settings_label = QLabel()
        self.tab_about_label = QLabel()
        self.tab_work_group_label = QGroupBox()
        self.tab_match_group_label = QGroupBox()
        self.match_mode_label = QLabel()
        self.log_group_success = QGroupBox()
        self.log_group_failure = QGroupBox()
        
        # excel tab widgets
        self.excel_save_btn = QPushButton(self)
        self.excel_add_row_btn = QPushButton(self)
        self.updated_excel_label = QLabel()
        self.origin_excel_label = QLabel()

        # settings tab widgets
        self.lang_label = QLabel()
        self.lang_combo = QComboBox(self)

        # === 关键修改点1：在UI初始化之前调用 setup_excel_files，并保存路径 ===
        self.excel_file_path, self.updated_excel_path = setup_excel_files()

        self.model_origin = ExcelTableModel(is_read_only=False)
        self.model_updated = ExcelTableModel(is_read_only=True)
        self.view_origin = CustomTableView(self)
        self.view_updated = CustomTableView(self)
        self.success_edit = QTextEdit(self)
        self.fail_edit = QTextEdit(self)
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setTextVisible(True)
        self.progress_label = QLabel(self)
        self.progress_label.setWordWrap(True)
        self.progress_label.setAlignment(Qt.AlignCenter)
        self.tab = QTabWidget()
        self.tab.setTabBar(SlidingTabBar())
        self.thread = None
        self.worker = None
        
        # about tab widgets
        self.about_text_edit = QTextEdit(self)

        self._build_ui()
        self._retranslate_ui() # 启动时进行一次UI翻译
        self._signals()
        self._apply_styles()
        # 延迟加载和调整，确保UI完全可见后再执行
        QTimer.singleShot(0, self.load_excels)
        self.load_paths()
        self._initial_ui_state()
        self.excel_le.dragEnterEvent = self.dragEnterEvent
        self.excel_le.dropEvent = lambda event: self.dropEvent(event, self.excel_le)
        self.target_le.dragEnterEvent = self.dragEnterEvent
        self.target_le.dropEvent = lambda event: self.dropEvent(event, self.target_le)
        self.root_le.dragEnterEvent = self.dragEnterEvent
        self.root_le.dropEvent = lambda event: self.dropEvent(event, self.root_le)
        
        # 为按钮设置图标
        down_arrow_path = resource_path('resources/down_arrow.png')
        self.excel_btn.setIcon(QIcon(down_arrow_path))
        self.target_btn.setIcon(QIcon(down_arrow_path))
        self.root_btn.setIcon(QIcon(down_arrow_path))


    def dragEnterEvent(self, event):
      if event.mimeData().hasUrls():
        event.acceptProposedAction()

    def dropEvent(self, event, line_edit):
        urls = event.mimeData().urls()
        if urls:
            path = urls[0].toLocalFile()
        if os.path.exists(path):
            line_edit.setText(path)

    def _build_work_tab(self):
        """构建工作终端页签。"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        form = QGroupBox()
        form_layout = QVBoxLayout(form)

        for title_key, le, btn in [('excel_list', self.excel_le, self.excel_btn),
                               ('target_directory', self.target_le, self.target_btn),
                               ('search_root', self.root_le, self.root_btn)]:
            h_layout = QHBoxLayout()
            h_layout.addWidget(QLabel(get_translation(title_key, self._language)))
            h_layout.addWidget(le)
            h_layout.addWidget(btn)
            form_layout.addLayout(h_layout)
        layout.addWidget(form)
        self.tab_work_group_label = form

        match_mode_group = QGroupBox()
        match_mode_layout = QHBoxLayout(match_mode_group)
        self.match_mode_label = QLabel()
        self.match_mode_combo.setObjectName('match_mode_combo') # 为匹配模式下拉框添加对象名
        self.match_mode_combo.addItems([get_translation('exact_match', self._language), get_translation('fuzzy_match', self._language), get_translation('regex_match', self._language)])
        
        match_mode_layout.addWidget(self.match_mode_label)
        match_mode_layout.addWidget(self.match_mode_combo)
        layout.addWidget(match_mode_group)
        self.tab_match_group_label = match_mode_group

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.create_refresh_excels_btn)
        button_layout.addWidget(self.start_btn)
        button_layout.addWidget(self.cancel_btn)
        layout.addLayout(button_layout)

        layout.addWidget(self.progress_label)
        layout.addWidget(self.progress_bar)
        
        return widget

    def _build_excel_tab(self, model, title, view_instance):
        """构建 Excel 预览页签。"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        header_layout = QHBoxLayout()
        header_label = QLabel()
        header_layout.addWidget(header_label)
        header_layout.addStretch()

        if not model.is_read_only:
            self.excel_save_btn = QPushButton()
            self.excel_save_btn.clicked.connect(lambda: self.save_excel(model, title))
            header_layout.addWidget(self.excel_save_btn)
            
            self.excel_add_row_btn = QPushButton()
            self.excel_add_row_btn.clicked.connect(model.appendRow)
            header_layout.addWidget(self.excel_add_row_btn)
            
            if "file_list.xlsx" in title:
                self.origin_excel_label = header_label
            else:
                self.updated_excel_label = header_label
        else:
            self.updated_excel_label = header_label
        
        layout.addLayout(header_layout)
        
        view_instance.setModel(model)
        view_instance.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        view_instance.horizontalHeader().setMinimumSectionSize(120)
        view_instance.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        
        layout.addWidget(view_instance)
        
        return widget

    def _build_settings_tab(self):
        """构建新的设置页签。"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # 语言设置
        language_group = QGroupBox()
        language_group.setObjectName('lang_group')
        lang_layout = QHBoxLayout(language_group)
        self.lang_label = QLabel()
        self.lang_label.setObjectName('lang_label')  # 为标签设置对象名称，以便样式表定位
        self.lang_combo = QComboBox(self)
        self.lang_combo.setObjectName('lang_combo')  # 为下拉框设置对象名称，以便样式表定位
        self.lang_combo.addItems([get_translation('chinese', self._language), get_translation('english', self._language)])
        self.lang_combo.setCurrentIndex(0 if self._language == 'zh' else 1)
        self.lang_combo.currentIndexChanged.connect(self._change_language)

        lang_layout.addWidget(self.lang_label)
        lang_layout.addWidget(self.lang_combo)
        
        layout.addWidget(language_group)
        layout.addStretch()
        
        return widget

    def _build_about_tab(self):
        """构建关于页签。"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        self.about_text_edit = QTextEdit()
        self.about_text_edit.setReadOnly(True)
        self.about_text_edit.setStyleSheet("""
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
        
        layout.addWidget(self.about_text_edit)
        return widget
        
    def _build_ui(self):
        """构建主界面布局。"""
        main = QVBoxLayout(self)

        title_bar = QHBoxLayout()
        title_label = QLabel()
        title_label.setAlignment(Qt.AlignCenter)
        title_bar.addWidget(title_label)
        title_bar.addStretch()
        close_btn = QPushButton("X")
        close_btn.clicked.connect(self.close)
        title_bar.addWidget(close_btn)
        self.title_label = title_label
        main.addLayout(title_bar)

        self.tab.addTab(self._build_work_tab(), "")
        self.tab.addTab(self._build_excel_tab(self.model_origin, "file_list.xlsx", self.view_origin), "")
        self.tab.addTab(self._build_excel_tab(self.model_updated, "file_list_updated.xlsx", self.view_updated), "")
        self.tab.addTab(self._build_settings_tab(), "")
        self.tab.addTab(self._build_about_tab(), "")
        main.addWidget(self.tab)

        log_splitter = QSplitter(Qt.Horizontal)
        self.log_group_success = self._log_group("", self.success_edit)
        self.log_group_failure = self._log_group("", self.fail_edit)
        log_splitter.addWidget(self.log_group_success)
        log_splitter.addWidget(self.log_group_failure)
        main.addWidget(log_splitter)

    def _log_group(self, title, text_edit):
        """创建一个日志分组框。"""
        group = QGroupBox(title)
        layout = QVBoxLayout(group)
        layout.addWidget(text_edit)
        return group

    def _signals(self):
        """连接所有信号和槽。"""
        self.excel_btn.clicked.connect(lambda: self.choose_file(self.excel_le, 'Excel 文件 (*.xlsx)'))
        self.target_btn.clicked.connect(lambda: self.choose_folder(self.target_le))
        self.root_btn.clicked.connect(lambda: self.choose_folder(self.root_le))
        self.start_btn.clicked.connect(self.start_task)
        self.create_refresh_excels_btn.clicked.connect(self._create_and_refresh_excels)
        self.cancel_btn.clicked.connect(self.cancel_task)
        
        # 当 tab 发生改变时，强制刷新表格行高
        self.tab.currentChanged.connect(self._handle_tab_change)

    def _handle_tab_change(self, index):
        """处理标签页切换事件，并强制刷新表格布局。"""
        if index == 1:  # '目标文件列' 是第二个标签页 (索引为 1)
            self.view_origin.resizeRowsToContents()
        elif index == 2:  # '执行结果列' 是第三个标签页 (索引为 2)
            self.view_updated.resizeRowsToContents()

    def _retranslate_ui(self):
               # 更新工作终端页签内的元素
        self.tab_work_group_label.setTitle(get_translation('path_settings', self._language))
        # 明确保存三个标签的引用（建议在 _build_work_tab 中初始化时保存）
        # 由于您没有在类中保存这些 QLabel 的引用，我们通过布局结构来查找
        form_layout = self.tab_work_group_label.layout()
        # 假设每个 QHBoxLayout 包含一个 QLabel 和 QLineEdit + QPushButton
        for i, title_key in enumerate(['excel_list', 'target_directory', 'search_root']):
            row_layout = form_layout.itemAt(i)
            if row_layout and row_layout.count() > 0:
                label_item = row_layout.itemAt(0)
                if label_item:
                    label = label_item.widget()
                    if isinstance(label, QLabel):
                        label.setText(get_translation(title_key, self._language))
    # 更新按钮文本
        self.excel_btn.setText(get_translation('browse', self._language))
        self.target_btn.setText(get_translation('browse', self._language))
        self.root_btn.setText(get_translation('browse', self._language))
        # 主窗口标题
        self.setWindowTitle(get_translation('title', self._language))
        self.title_label.setText(get_translation('title', self._language))
        
        # 页签标题
        self.tab.setTabText(0, get_translation('work_terminal', self._language))
        self.tab.setTabText(1, get_translation('target_files', self._language))
        self.tab.setTabText(2, get_translation('execution_results', self._language))
        self.tab.setTabText(3, get_translation('settings', self._language))
        self.tab.setTabText(4, get_translation('about', self._language))
        
        # 工作终端页签
        self.tab_work_group_label.setTitle(get_translation('path_settings', self._language))
        self.tab_match_group_label.setTitle(get_translation('match_settings', self._language))
        self.match_mode_label.setText(get_translation('match_mode', self._language))
        self.excel_btn.setText(get_translation('browse', self._language))
        self.target_btn.setText(get_translation('browse', self._language))
        self.root_btn.setText(get_translation('browse', self._language))
        self.create_refresh_excels_btn.setText(get_translation('create_refresh', self._language))
        self.start_btn.setText(get_translation('start', self._language))
        self.cancel_btn.setText(get_translation('cancel', self._language))
        self.progress_bar.setFormat(get_translation('preparing', self._language))
        self.progress_label.setText(get_translation('status_waiting', self._language))
        
        # 匹配模式下拉框
        # 保存当前选择，避免重置
        current_index = self.match_mode_combo.currentIndex()
        self.match_mode_combo.clear()
        self.match_mode_combo.addItems([get_translation('exact_match', self._language), get_translation('fuzzy_match', self._language), get_translation('regex_match', self._language)])
        self.match_mode_combo.setCurrentIndex(current_index)

        # 日志分组框
        self.log_group_success.setTitle(get_translation('success_log', self._language))
        self.log_group_failure.setTitle(get_translation('failure_log', self._language))

        # Excel页签中的按钮和标签
        self.origin_excel_label.setText(f"<b>{get_translation('file_preview', self._language)}file_list.xlsx</b>")
        self.updated_excel_label.setText(f"<b>{get_translation('file_preview', self._language)}file_list_updated.xlsx</b>")
        self.excel_save_btn.setText(get_translation('save', self._language))
        self.excel_add_row_btn.setText(get_translation('add_row', self._language))
        
        # Settings页签
        settings_group = self.findChild(QGroupBox, 'lang_group')
        settings_group.setTitle(get_translation('language_settings', self._language))
        self.lang_label.setText(get_translation('language_settings', self._language))
        
        # 关于页签
        self.about_text_edit.setText(get_translation('about_text', self._language))
        
        # 语言下拉框
        # 在重新翻译UI时，暂时阻止lang_combo的信号，防止递归
        self.lang_combo.blockSignals(True)
        current_lang_index = 0 if self._language == 'zh' else 1
        self.lang_combo.clear()
        self.lang_combo.addItems([get_translation('chinese', self._language), get_translation('english', self._language)])
        self.lang_combo.setCurrentIndex(current_lang_index)
        self.lang_combo.blockSignals(False)


    def _change_language(self, index):
        """处理语言切换。"""
        new_lang = 'zh' if index == 0 else 'en'
        if new_lang != self._language:
            self._language = new_lang
            self.save_settings()
            
            self._retranslate_ui()
            
            self.success_edit.append(get_translation('language_changed', self._language))
            
    def choose_file(self, line_edit, filt):
        """选择文件。"""
        path, file_filter = QFileDialog.getOpenFileName(self, get_translation('browse', self._language), filter=filt)
        if path:
            line_edit.setText(path)
            # 在选择新文件后，手动更新 excel model 的路径
            if line_edit == self.excel_le:
                self.excel_file_path = path
                self.load_excels()


    def choose_folder(self, line_edit):
        """选择文件夹。"""
        path = QFileDialog.getExistingDirectory(self, get_translation('browse', self._language))
        if path:
            line_edit.setText(path)

    def open_excel(self, updated=False):
        """打开 Excel 文件。"""
        QDesktopServices.openUrl(QUrl.fromLocalFile(self.updated_excel_path if updated else self.excel_file_path))

    def load_excels(self):
        """加载 Excel 文件到模型中。"""
        # === 关键修改点2：直接使用已存储的路径来加载模型 ===
        self.model_origin.load(self.excel_file_path)
        self.model_updated.load(self.updated_excel_path)

        # 在加载数据后强制刷新表格视图的行高
        self.view_origin.resizeRowsToContents()
        self.view_updated.resizeRowsToContents()


    def save_excel(self, model, title):
        """保存 Excel 文件。"""
        # === 关键修改点3：直接使用已存储的路径来保存 ===
        path = self.excel_file_path if "file_list.xlsx" in title else self.updated_excel_path
        model.save(path)
        self.success_edit.append(f"✅ {title} {get_translation('save', self._language)}")

    def _create_and_refresh_excels(self):
        """创建或刷新 Excel 文件并加载。"""
        try:
            # === 关键修改点4：调用新函数并更新路径，确保同步 ===
            self.excel_file_path, self.updated_excel_path = setup_excel_files()
            self.load_excels()
            self.success_edit.append(f"✅ {get_translation('excel_refreshed_success', self._language)}")
            self.excel_le.setText(self.excel_file_path)
        except Exception as e:
            traceback.print_exc()
            self.fail_edit.append(f"❌ {get_translation('excel_refresh_fail', self._language)}{e}")

    def _initial_ui_state(self):
        """设置初始 UI 状态。"""
        self.start_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.progress_label.setText(get_translation('status_waiting', self._language))
        self.progress_bar.setFormat(get_translation('preparing', self._language))

    def start_task(self):
        """开始执行任务。"""
        excel = self.excel_le.text()
        target = self.target_le.text()
        root = self.root_le.text()
        if not all([excel, target, root]):
            self.fail_edit.append(get_translation('path_not_set_error', self._language))
            return
        
        self.start_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        
        self.success_edit.clear()
        self.fail_edit.clear()
        self.progress_bar.setValue(0)
        self.progress_label.setText(get_translation('status_initializing', self._language))

        self.save_paths()

        # === 关键修改点5：使用已存储的路径创建 SearchWorker ===
        match_mode_text = self.match_mode_combo.currentText()
        if get_translation('fuzzy_match', self._language) in match_mode_text:
            match_mode = 'fuzzy'
        elif get_translation('regex_match', self._language) in match_mode_text:
            match_mode = 'regex'
        else:
            match_mode = 'exact'
            
        self.worker = SearchWorker(
            excel_path=self.excel_file_path,
            target_dir=target,
            roots=[root],
            updated_excel_path=self.updated_excel_path,
            match_mode=match_mode,
            min_fuzzy_score=85
        )
        
        self.thread = QThread(self)
        self.worker.moveToThread(self.thread)
        
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self._on_task_finished)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)

        self.worker.success.connect(lambda s: self.success_edit.append(s))
        self.worker.failed.connect(lambda s: self.fail_edit.append(s))
        self.worker.progress.connect(self.update_progress)
        self.thread.start()

    def cancel_task(self):
        """取消当前任务。"""
        if self.worker and self.thread.isRunning():
            self.worker.stop()
            self.cancel_btn.setEnabled(False)
            self.progress_label.setText(get_translation('user_cancel', self._language))
            self.fail_edit.append(get_translation('user_cancel', self._language))

    def update_progress(self, current, total, message):
        """更新进度条和标签。"""
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(current)
        
        max_len = 50 
        display_message = message
        if len(message) > max_len:
            display_message = message[:max_len-3] + '...'
            
        self.progress_bar.setFormat(f"{get_translation('status_prefix', self._language)} {display_message} %p%")
        self.progress_label.setText(f"{get_translation('status_prefix', self._language)} {display_message}")

    def _on_task_finished(self):
        """任务完成后的处理。"""
        self.start_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        self.progress_label.setText(get_translation('task_completed_msg', self._language))
        self.progress_bar.setValue(self.progress_bar.maximum())
        self.progress_bar.setFormat(get_translation('task_completed', self._language))
        
        self.load_excels()

    def load_paths(self):
        """加载上次的路径设置。"""
        self.excel_le.setText(self.excel_file_path)

        try:
            with open(self.SETTINGS_FILE, "r", encoding="utf-8") as f:
                lines = f.read().splitlines()
                if len(lines) >= 3:
                    self.target_le.setText(lines[1])
                    self.root_le.setText(lines[2])
        except FileNotFoundError:
            pass

    def save_paths(self):
        """保存当前路径设置。"""
        with open(self.SETTINGS_FILE, "w", encoding="utf-8") as f:
            f.write("\n".join([self.excel_le.text(), self.target_le.text(), self.root_le.text()]))

    def load_settings(self):
        """加载配置文件中的设置，包括语言。"""
        try:
            with open(self.CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
                self._language = config.get('language', 'zh')
        except (FileNotFoundError, json.JSONDecodeError):
            self._language = 'zh'
    
    def save_settings(self):
        """保存设置到配置文件。"""
        config = {'language': self._language}
        try:
            with open(self.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
        except Exception:
            traceback.print_exc()
    
    def _apply_styles(self):
        """应用样式表。"""
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
    min-width: 130px; /* 减小最小宽度，使页签变窄 */
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

/* 修正后的进度条样式 */
QProgressBar {
    border: 1px solid #00FFFF;
    border-radius: 5px;
    background-color: #1A1F36;
    text-align: center;
    color: #E0E0E0;
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

/* 语言设置的标签和下拉框样式 */
QLabel#lang_label, QComboBox#lang_combo {
    font-size: 12pt;
    font-weight: bold;
    color: #E0E0E0;
}
QComboBox {
    border: 1px solid #00FFFF;
    border-radius: 4px;
    padding: 4px;
    background-color: #1A1F36;
    color: #00FFFF;
}

/* 确保匹配模式下拉框的样式与语言下拉框一致 */
QComboBox#match_mode_combo {
    font-size: 9pt; /* 保持与之前的一致性 */
    font-weight: bold;
    color: #00FFFF;
}

QComboBox::drop-down {
    border: 0px;
}

QComboBox::down-arrow {
    image: url(:/icons/down_arrow.png);
    width: 12px;
    height: 12px;
}

QComboBox QAbstractItemView {
    border: 1px solid #00FFFF;
    background-color: #1A1F36;
    color: #E0E0E0;
    selection-background-color: #00FFFF;
    selection-color: #1A1F36;
}
""")
        self.progress_label.setObjectName("progress_label")