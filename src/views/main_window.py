# 使用PySide6构建主界面，包含菜单栏、标签页，并集成控制器。

import sys
import os
from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout,
                               QPushButton, QFileDialog, QLabel,
                               QComboBox, QTextEdit, QProgressBar,
                               QMessageBox, QTabWidget, QListWidget,
                               QListWidgetItem, QHBoxLayout)
from PySide6.QtCore import Qt, Slot
from PySide6.QtGui import QIcon
from src.controllers.main_controller import MainController
from src.utils.logger import setup_logger
import os

logger = setup_logger(__name__)


class MainWindow(QMainWindow):
    def __init__(self, config):
        super().__init__()
        self.config = config
        self.controller = MainController()
        self.setWindowTitle("Excel效率工具")
        icon_path = self.resource_path(os.path.join("resources", "icons", "myapp.ico"))
        self.setWindowIcon(QIcon(icon_path))
        self.resize(
            self.config.getint('APP', 'window_width', 1000),
            self.config.getint('APP', 'window_height', 700)
        )

        # 加载样式表
        self._load_stylesheet()

        # 创建中央控件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # 创建标签页
        self.tab_widget = QTabWidget()
        layout.addWidget(self.tab_widget)

        # 初始化各个功能模块的页面
        self._init_merge_tab()  # 合并功能页
        self._init_match_tab()   # 匹配组件
        self._init_filter_tab()  # 筛选组件
        self._init_reconciliation_tab()   # 对账组件

        # 状态栏添加进度条
        self.status_bar = self.statusBar()
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximumWidth(200)
        self.progress_bar.hide()
        self.status_bar.addPermanentWidget(self.progress_bar)

        # 日志输出控件（可在调试时显示信息）
        self.log_text = QTextEdit()
        self.log_text.setMaximumHeight(100)
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text)

    # --- 辅助函数，用于获取资源文件的正确路径 ---
    def resource_path(self, relative_path):
        """获取资源的绝对路径，兼容开发环境和 PyInstaller 打包后的环境。"""
        try:
            # PyInstaller 创建临时文件夹，将路径存储在 _MEIPASS 中
            base_path = sys._MEIPASS
        except Exception:
            # 如果不是打包环境，则使用当前脚本的目录
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)
    def _load_stylesheet(self):
        """加载QSS样式表"""
        style_path = os.path.join(os.path.dirname(__file__), '..', 'resources', 'styles.qss')
        if os.path.exists(style_path):
            with open(style_path, 'r', encoding='utf-8') as f:
                self.setStyleSheet(f.read())
    def _init_match_tab(self):
        """初始化数据匹配标签页"""
        from src.views.tabs.match_tab import MatchTab # 匹配组件
        self.progress_bar = QProgressBar()
        self.match_tab = MatchTab(
            controller=self.controller,
            status_bar=self.statusBar(),
            log_text=QTextEdit(),
            progress_bar=self.progress_bar
        )
        self.tab_widget.addTab(self.match_tab, "数据匹配")


    def _init_merge_tab(self):
        """初始化“文件合并”标签页"""
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # 文件选择区域
        file_layout = QHBoxLayout()
        self.file_list_widget = QListWidget()
        file_layout.addWidget(self.file_list_widget)

        btn_layout = QVBoxLayout()
        self.btn_add_files = QPushButton("添加文件")
        self.btn_add_files.clicked.connect(self._select_files)
        self.btn_clear = QPushButton("清空列表")
        self.btn_clear.clicked.connect(self.file_list_widget.clear)
        btn_layout.addWidget(self.btn_add_files)
        btn_layout.addWidget(self.btn_clear)
        btn_layout.addStretch()
        file_layout.addLayout(btn_layout)
        layout.addLayout(file_layout)

        # 参数设置
        param_layout = QHBoxLayout()
        param_layout.addWidget(QLabel("合并方式:"))
        self.combo_merge_type = QComboBox()
        self.combo_merge_type.addItems(["纵向合并(追加行)", "横向合并(拼接列)"])
        param_layout.addWidget(self.combo_merge_type)

        param_layout.addWidget(QLabel("输出文件:"))
        self.output_path_edit = QLabel("未选择")
        param_layout.addWidget(self.output_path_edit)
        self.btn_output = QPushButton("浏览...")
        self.btn_output.clicked.connect(self._select_output_file)
        param_layout.addWidget(self.btn_output)
        param_layout.addStretch()
        layout.addLayout(param_layout)

        # 执行按钮
        self.btn_start = QPushButton("开始合并")
        self.btn_start.clicked.connect(self._on_merge_start)
        layout.addWidget(self.btn_start)

        self.tab_widget.addTab(tab, "文件合并")

    def _select_files(self):
        """添加Excel文件到列表"""
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择Excel文件", "",
            "Excel文件 (*.xlsx *.xls);;所有文件 (*.*)"
        )
        for f in files:
            item = QListWidgetItem(f)
            self.file_list_widget.addItem(item)

    def _init_filter_tab(self):
        """初始化多文件筛选标签页"""
        from src.views.tabs.filter_tab import FilterTab # 筛选组件
        self.filter_tab = FilterTab(
            controller=self.controller,
            status_bar=self.statusBar(),
            log_text=QTextEdit(),
            progress_bar=self.progress_bar
        )
        self.tab_widget.addTab(self.filter_tab, "多文件筛选")
    
    def _select_output_file(self):
        """选择输出文件路径"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存合并后的文件", "",
            "Excel文件 (*.xlsx);;Excel 97-2003 (*.xls)"
        )
        if file_path:
            # 自动添加扩展名
            if not (file_path.endswith('.xlsx') or file_path.endswith('.xls')):
                file_path += '.xlsx'
            self.output_path_edit.setText(file_path)

    def _init_reconciliation_tab(self):
        """初始化绿能畅游对账标签页"""
        from src.views.tabs.reconciliation_tab import ReconciliationTab # 对账组件
        self.reconciliation_tab = ReconciliationTab(
            controller=self.controller,
            status_bar=self.statusBar(),
            log_text=QTextEdit(),          # 使用主窗口底部的日志控件
            progress_bar=self.progress_bar
        )
        self.tab_widget.addTab(self.reconciliation_tab, "绿能畅游对账")

    @Slot()
    def _on_merge_start(self):
        """开始合并任务"""
        # 获取参数
        file_count = self.file_list_widget.count()
        if file_count == 0:
            QMessageBox.warning(self, "警告", "请至少选择一个Excel文件")
            return
        file_list = [self.file_list_widget.item(i).text() for i in range(file_count)]

        output_path = self.output_path_edit.text()
        if output_path == "未选择" or not output_path:
            QMessageBox.warning(self, "警告", "请选择输出文件路径")
            return

        merge_type = 'rows' if self.combo_merge_type.currentIndex() == 0 else 'cols'

        # 禁用按钮，显示进度条
        self.btn_start.setEnabled(False)
        self.progress_bar.setRange(0, 0)  # 繁忙状态
        self.progress_bar.show()
        self.log_text.append("开始执行合并任务...")

        # 调用控制器启动线程
        self.controller.start_merge_task(
            file_list, output_path, merge_type,
            progress_callback=self._update_progress,
            message_callback=self._update_message,
            finished_callback=self._on_task_finished,
            error_callback=self._on_task_error
        )

    @Slot(int)
    def _update_progress(self, value):
        """更新进度条（如果业务层支持进度）"""
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(value)

    @Slot(str)
    def _update_message(self, msg):
        """更新状态信息"""
        self.log_text.append(msg)
        self.status_bar.showMessage(msg, 3000)

    @Slot(object)
    def _on_task_finished(self, result):
        """任务完成回调"""
        self.btn_start.setEnabled(True)
        self.progress_bar.hide()
        self.log_text.append(f"任务完成，文件已保存至: {result}")
        QMessageBox.information(self, "完成", f"合并成功！\n输出文件: {result}")

    @Slot(str)
    def _on_task_error(self, error_msg):
        """任务出错回调"""
        self.btn_start.setEnabled(True)
        self.progress_bar.hide()
        self.log_text.append(f"错误: {error_msg}")
        QMessageBox.critical(self, "错误", f"任务执行失败:\n{error_msg}")