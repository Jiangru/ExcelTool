# src/views/tabs/match_tab.py

import os
import pandas as pd
from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout,
                               QPushButton, QFileDialog, QLabel,
                               QComboBox, QListWidget, QAbstractItemView,
                               QLineEdit, QGroupBox, QRadioButton,
                               QButtonGroup, QMessageBox, QProgressBar)
from PySide6.QtCore import Slot, Qt
from src.utils.logger import setup_logger

logger = setup_logger(__name__)


class MatchTab(QWidget):
    """数据匹配合并标签页"""
    def __init__(self, controller, status_bar, log_text, progress_bar):
        super().__init__()
        self.controller = controller
        self.status_bar = status_bar
        self.log_text = log_text
        self.progress_bar = progress_bar

        self.file_a_path = ""
        self.file_b_path = ""
        self.output_path = ""

        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)

        # ---------- 文件选择区 ----------
        file_group = QGroupBox("1. 选择文件")
        file_layout = QVBoxLayout()

        # 表A
        hlayout_a = QHBoxLayout()
        hlayout_a.addWidget(QLabel("表A（主表）："))
        self.label_a = QLabel("未选择")
        hlayout_a.addWidget(self.label_a)
        self.btn_a = QPushButton("浏览...")
        self.btn_a.clicked.connect(self._select_file_a)
        hlayout_a.addWidget(self.btn_a)
        file_layout.addLayout(hlayout_a)

        # 表B
        hlayout_b = QHBoxLayout()
        hlayout_b.addWidget(QLabel("表B（匹配表）："))
        self.label_b = QLabel("未选择")
        hlayout_b.addWidget(self.label_b)
        self.btn_b = QPushButton("浏览...")
        self.btn_b.clicked.connect(self._select_file_b)
        hlayout_b.addWidget(self.btn_b)
        file_layout.addLayout(hlayout_b)

        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # ---------- 匹配键设置区 ----------
        key_group = QGroupBox("2. 设置匹配列")
        key_layout = QVBoxLayout()

        # 表A匹配列
        hlayout_key_a = QHBoxLayout()
        hlayout_key_a.addWidget(QLabel("表A匹配列："))
        self.combo_key_a = QComboBox()
        self.combo_key_a.setMinimumWidth(150)
        hlayout_key_a.addWidget(self.combo_key_a)
        hlayout_key_a.addStretch()
        key_layout.addLayout(hlayout_key_a)

        # 表B匹配列
        hlayout_key_b = QHBoxLayout()
        hlayout_key_b.addWidget(QLabel("表B匹配列："))
        self.combo_key_b = QComboBox()
        self.combo_key_b.setMinimumWidth(150)
        hlayout_key_b.addWidget(self.combo_key_b)
        hlayout_key_b.addStretch()
        key_layout.addLayout(hlayout_key_b)

        # 提示：多列匹配需后续扩展，当前简化版本仅支持单列
        self.label_multi = QLabel("（当前版本支持单列匹配，多列匹配将在后续添加）")
        self.label_multi.setStyleSheet("color: gray;")
        key_layout.addWidget(self.label_multi)

        key_group.setLayout(key_layout)
        layout.addWidget(key_group)

        # ---------- 需添加的列选择区 ----------
        col_group = QGroupBox("3. 选择要合并的表B列")
        col_layout = QVBoxLayout()
        self.col_list_widget = QListWidget()
        self.col_list_widget.setSelectionMode(QAbstractItemView.MultiSelection)
        col_layout.addWidget(self.col_list_widget)

        btn_layout = QHBoxLayout()
        self.btn_select_all = QPushButton("全选")
        self.btn_select_all.clicked.connect(self._select_all_cols)
        self.btn_clear_all = QPushButton("清空")
        self.btn_clear_all.clicked.connect(self._clear_all_cols)
        btn_layout.addWidget(self.btn_select_all)
        btn_layout.addWidget(self.btn_clear_all)
        btn_layout.addStretch()
        col_layout.addLayout(btn_layout)

        col_group.setLayout(col_layout)
        layout.addWidget(col_group)

        # ---------- 连接方式 ----------
        join_group = QGroupBox("4. 连接方式")
        join_layout = QHBoxLayout()
        self.join_left = QRadioButton("左连接（保留表A所有行）")
        self.join_left.setChecked(True)
        self.join_inner = QRadioButton("内连接（仅保留匹配上的行）")
        self.join_right = QRadioButton("右连接（保留表B所有行）")
        self.join_outer = QRadioButton("全连接")
        join_layout.addWidget(self.join_left)
        join_layout.addWidget(self.join_inner)
        join_layout.addWidget(self.join_right)
        join_layout.addWidget(self.join_outer)
        join_layout.addStretch()
        join_group.setLayout(join_layout)
        layout.addWidget(join_group)

        # ---------- 输出设置 ----------
        output_group = QGroupBox("5. 输出设置")
        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("输出文件："))
        self.output_label = QLabel("未选择，将自动生成")
        output_layout.addWidget(self.output_label)
        self.btn_output = QPushButton("浏览...")
        self.btn_output.clicked.connect(self._select_output)
        output_layout.addWidget(self.btn_output)
        output_layout.addStretch()
        output_group.setLayout(output_layout)
        layout.addWidget(output_group)

        # ---------- 执行按钮 ----------
        self.btn_start = QPushButton("开始匹配合并")
        self.btn_start.clicked.connect(self._on_start)
        self.btn_start.setMinimumHeight(40)
        layout.addWidget(self.btn_start)

        layout.addStretch()

        # 初始化状态
        self._update_ui_state()

    # ---------- 槽函数 ----------
    def _select_file_a(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择表A文件", "", "Excel文件 (*.xlsx *.xls)"
        )
        if path:
            self.file_a_path = path
            self.label_a.setText(path)
            self._load_column_names(path, self.combo_key_a, is_a=True)
            # 显式调用 self._update_ui_state()，确保控件状态与文件选择状态同步。
            self._update_ui_state()

    def _select_file_b(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择表B文件", "", "Excel文件 (*.xlsx *.xls)"
        )
        if path:
            self.file_b_path = path
            self.label_b.setText(path)
            self._load_column_names(path, self.combo_key_b, is_a=False)
            self._load_columns_for_selection(path)
            # 显式调用 self._update_ui_state()，确保控件状态与文件选择状态同步。
            self._update_ui_state()

    def _load_column_names(self, file_path, combo_widget, is_a=False):
        """读取Excel的第一行列名，填充到下拉框"""
        try:
            # 仅读取第一行以获取列名
            if file_path.endswith('.xls'):
                df = pd.read_excel(file_path, nrows=0, engine='xlrd')
            else:
                df = pd.read_excel(file_path, nrows=0, engine='openpyxl')
            columns = df.columns.tolist()
            combo_widget.clear()
            combo_widget.addItems(columns)
            # 默认选中第一列（如果有）
            if columns:
                combo_widget.setCurrentIndex(0)
        except Exception as e:
            QMessageBox.warning(self, "读取失败", f"无法读取文件列名：{str(e)}")
            logger.error(f"读取列名失败 {file_path}: {e}")

    def _load_columns_for_selection(self, file_path):
        """读取表B的所有列名，供用户选择要添加的列"""
        try:
            if file_path.endswith('.xls'):
                df = pd.read_excel(file_path, nrows=0, engine='xlrd')
            else:
                df = pd.read_excel(file_path, nrows=0, engine='openpyxl')
            columns = df.columns.tolist()
            self.col_list_widget.clear()
            for col in columns:
                self.col_list_widget.addItem(col)
        except Exception as e:
            QMessageBox.warning(self, "读取失败", f"无法读取表B列名：{str(e)}")
            logger.error(f"读取表B列名失败 {file_path}: {e}")

    def _select_output(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "保存匹配结果", "", "Excel文件 (*.xlsx);;Excel 97-2003 (*.xls)"
        )
        if path:
            if not (path.endswith('.xlsx') or path.endswith('.xls')):
                path += '.xlsx'
            self.output_path = path
            self.output_label.setText(path)

    def _select_all_cols(self):
        self.col_list_widget.selectAll()

    def _clear_all_cols(self):
        self.col_list_widget.clearSelection()

    def _update_ui_state(self):
        """根据文件选择情况更新界面可用性"""
        has_a = bool(self.file_a_path)
        has_b = bool(self.file_b_path)
        self.combo_key_a.setEnabled(has_a)
        self.combo_key_b.setEnabled(has_b)
        self.col_list_widget.setEnabled(has_b)
        self.btn_select_all.setEnabled(has_b)
        self.btn_clear_all.setEnabled(has_b)
        self.btn_start.setEnabled(has_a and has_b)

    @Slot()
    def _on_start(self):
        """开始匹配任务"""
        # 参数校验
        if not self.file_a_path or not self.file_b_path:
            QMessageBox.warning(self, "警告", "请先选择表A和表B文件")
            return

        key_a = self.combo_key_a.currentText()
        if not key_a:
            QMessageBox.warning(self, "警告", "请选择表A的匹配列")
            return

        key_b = self.combo_key_b.currentText()
        if not key_b:
            QMessageBox.warning(self, "警告", "请选择表B的匹配列")
            return

        # 获取要添加的列
        selected_items = self.col_list_widget.selectedItems()
        if selected_items:
            cols_to_add = [item.text() for item in selected_items]
        else:
            # 如果未选择任何列，默认添加所有非匹配键列
            cols_to_add = None  # 业务层会处理为所有非匹配键列

        # 连接方式
        how = 'left'
        if self.join_inner.isChecked():
            how = 'inner'
        elif self.join_right.isChecked():
            how = 'right'
        elif self.join_outer.isChecked():
            how = 'outer'

        # 输出路径
        if not self.output_path:
            # 自动生成路径
            import os
            from pathlib import Path
            base = Path(self.file_a_path).parent
            name = f"{Path(self.file_a_path).stem}_matched.xlsx"
            self.output_path = str(base / name)
            self.output_label.setText(self.output_path)

        # 禁用按钮，显示进度
        self.btn_start.setEnabled(False)
        self.progress_bar.setRange(0, 0)
        self.progress_bar.show()
        self.log_text.append("开始数据匹配合并...")

        # 调用控制器（假设controller有start_match_task方法）
        self.controller.start_match_task(
            file_a=self.file_a_path,
            file_b=self.file_b_path,
            key_a=[key_a],        # 单列，封装为列表
            key_b=[key_b],
            cols_b=cols_to_add,
            how=how,
            output=self.output_path,
            message_callback=self._on_message,
            finished_callback=self._on_finished,
            error_callback=self._on_error
        )

    @Slot(str)
    def _on_message(self, msg):
        self.log_text.append(msg)
        self.status_bar.showMessage(msg, 3000)

    @Slot(object)
    def _on_finished(self, result):
        self.btn_start.setEnabled(True)
        self.progress_bar.hide()
        self.log_text.append(f"匹配完成！文件已保存至：{result}")
        QMessageBox.information(self, "完成", f"数据匹配成功！\n输出文件：{result}")

    @Slot(str)
    def _on_error(self, err_msg):
        self.btn_start.setEnabled(True)
        self.progress_bar.hide()
        self.log_text.append(f"错误：{err_msg}")
        QMessageBox.critical(self, "错误", f"匹配任务失败：\n{err_msg}")
print("模块被成功加载，MatchTab 类存在")