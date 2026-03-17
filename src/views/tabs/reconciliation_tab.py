# src/views/tabs/reconciliation_tab.py

import os
import pandas as pd
from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout,
                               QPushButton, QFileDialog, QLabel,
                               QComboBox, QGroupBox, QMessageBox,
                               QCheckBox, QProgressBar)
from PySide6.QtCore import Slot, Qt
from src.utils.logger import setup_logger

logger = setup_logger(__name__)


class ReconciliationTab(QWidget):
    """绿能畅游对账标签页（含进度提示）"""
    def __init__(self, controller, status_bar, log_text, progress_bar):
        super().__init__()
        self.controller = controller
        self.status_bar = status_bar
        self.log_text = log_text
        self.progress_bar = progress_bar

        self.left_file_path = ""
        self.right_file_path = ""
        self.left_columns = []
        self.right_columns = []
        self.match_file_path = ""
        self.output_path = ""

        self.setup_ui()
        self._update_ui_state()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(15)

        # ---------- 1. 左右文件选择区 ----------
        file_layout = QHBoxLayout()

        # 左侧（达克云）
        left_group = QGroupBox("达克云平台订单（注：需要把原文件中【充电场站名称】列的抬头改为【场站】）")
        left_vbox = QVBoxLayout()
        left_file_row = QHBoxLayout()
        self.left_file_label = QLabel("未选择")
        left_file_row.addWidget(self.left_file_label)
        self.left_btn = QPushButton("上传")
        self.left_btn.clicked.connect(lambda: self._select_file('left'))
        left_file_row.addWidget(self.left_btn)
        left_file_row.addStretch()
        left_vbox.addLayout(left_file_row)

        left_energy_row = QHBoxLayout()
        left_energy_row.addWidget(QLabel("电量列:"))
        self.left_energy_combo = QComboBox()
        left_energy_row.addWidget(self.left_energy_combo)
        left_vbox.addLayout(left_energy_row)

        left_fee_row = QHBoxLayout()
        left_fee_row.addWidget(QLabel("费用列:"))
        self.left_fee_combo = QComboBox()
        left_fee_row.addWidget(self.left_fee_combo)
        left_vbox.addLayout(left_fee_row)

        left_group.setLayout(left_vbox)
        file_layout.addWidget(left_group, 1)

        # 右侧（海汇德）
        right_group = QGroupBox("海汇德平台订单")
        right_vbox = QVBoxLayout()
        right_file_row = QHBoxLayout()
        self.right_file_label = QLabel("未选择")
        right_file_row.addWidget(self.right_file_label)
        self.right_btn = QPushButton("上传")
        self.right_btn.clicked.connect(lambda: self._select_file('right'))
        right_file_row.addWidget(self.right_btn)
        right_file_row.addStretch()
        right_vbox.addLayout(right_file_row)

        right_energy_row = QHBoxLayout()
        right_energy_row.addWidget(QLabel("电量列:"))
        self.right_energy_combo = QComboBox()
        right_energy_row.addWidget(self.right_energy_combo)
        right_vbox.addLayout(right_energy_row)

        right_fee_row = QHBoxLayout()
        right_fee_row.addWidget(QLabel("费用列:"))
        self.right_fee_combo = QComboBox()
        right_fee_row.addWidget(self.right_fee_combo)
        right_vbox.addLayout(right_fee_row)

        right_group.setLayout(right_vbox)
        file_layout.addWidget(right_group, 1)

        layout.addLayout(file_layout)

        # ---------- 2. 场站名称列 ----------
        group_layout = QHBoxLayout()
        group_layout.addWidget(QLabel("充电场站名称列:"))
        self.group_combo = QComboBox()
        self.group_combo.setMinimumWidth(200)
        group_layout.addWidget(self.group_combo)
        group_layout.addStretch()
        layout.addLayout(group_layout)

        # ---------- 3. 外部匹配（剔除）设置 ----------
        self.match_checkbox = QCheckBox("启用外部文件匹配（剔除场站）")
        self.match_checkbox.toggled.connect(self._on_match_toggled)
        layout.addWidget(self.match_checkbox)

        self.match_group = QGroupBox("外部匹配设置")
        self.match_group.setEnabled(False)
        match_layout = QVBoxLayout()

        # 匹配文件选择
        file_match_layout = QHBoxLayout()
        file_match_layout.addWidget(QLabel("匹配文件:"))
        self.match_file_label = QLabel("未选择")
        file_match_layout.addWidget(self.match_file_label)
        self.match_btn = QPushButton("浏览...")
        self.match_btn.clicked.connect(self._select_match_file)
        file_match_layout.addWidget(self.match_btn)
        file_match_layout.addStretch()
        match_layout.addLayout(file_match_layout)

        # 外部文件中的场站列选择
        col_match_layout = QHBoxLayout()
        col_match_layout.addWidget(QLabel("外部文件中的场站列:"))
        self.match_col_combo = QComboBox()
        col_match_layout.addWidget(self.match_col_combo)
        col_match_layout.addStretch()
        match_layout.addLayout(col_match_layout)

        self.match_group.setLayout(match_layout)
        layout.addWidget(self.match_group)

        # ---------- 4. 输出设置 ----------
        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("输出文件:"))
        self.output_label = QLabel("未选择，将自动生成")
        output_layout.addWidget(self.output_label)
        self.output_btn = QPushButton("浏览...")
        self.output_btn.clicked.connect(self._select_output)
        output_layout.addWidget(self.output_btn)
        output_layout.addStretch()
        layout.addLayout(output_layout)

        # ---------- 5. 进度条与提示 ----------
        progress_layout = QHBoxLayout()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.hide()
        self.progress_label = QLabel("正在拼命处理中，请耐心等待。。。")
        self.progress_label.hide()
        progress_layout.addWidget(self.progress_bar)
        progress_layout.addWidget(self.progress_label)
        layout.addLayout(progress_layout)

        # ---------- 6. 执行按钮 ----------
        self.start_btn = QPushButton("开始对账")
        self.start_btn.clicked.connect(self._on_start)
        self.start_btn.setMinimumHeight(40)
        layout.addWidget(self.start_btn)

        layout.addStretch()

    # ---------- 文件选择 ----------
    def _select_file(self, side):
        path, _ = QFileDialog.getOpenFileName(
            self, f"选择{side}文件", "", "Excel文件 (*.xlsx *.xls)"
        )
        if not path:
            return
        if side == 'left':
            self.left_file_path = path
            self.left_file_label.setText(path)
            self._load_columns(path, 'left')
        else:
            self.right_file_path = path
            self.right_file_label.setText(path)
            self._load_columns(path, 'right')
        self._update_ui_state()

    def _load_columns(self, file_path, side):
        try:
            if file_path.endswith('.xls'):
                df = pd.read_excel(file_path, nrows=0, engine='xlrd')
            else:
                df = pd.read_excel(file_path, nrows=0, engine='openpyxl')
            columns = df.columns.tolist()
            if side == 'left':
                self.left_columns = columns
                self.left_energy_combo.clear()
                self.left_energy_combo.addItems(columns)
                self.left_fee_combo.clear()
                self.left_fee_combo.addItems(columns)
                # 设置默认值：电量列优先匹配 "总电量", "订单充电量（度）", "充电量", "总电量"
                self._set_combo_default(self.left_energy_combo, ["总电量", "订单充电量（度）", "充电量", "总电量"])
                self._set_combo_default(self.left_fee_combo, ["总费用", "应收总金额(元)", "金额"])
            else:
                self.right_columns = columns
                self.right_energy_combo.clear()
                self.right_energy_combo.addItems(columns)
                self.right_fee_combo.clear()
                self.right_fee_combo.addItems(columns)
                self._set_combo_default(self.right_energy_combo, ["总电量", "订单充电量（度）", "充电量", "总电量"])
                self._set_combo_default(self.right_fee_combo, ["总费用", "应收总金额(元)", "金额"])

            all_cols = list(set(self.left_columns + self.right_columns))
            self.group_combo.clear()
            self.group_combo.addItems(all_cols)
            self._set_combo_default(self.group_combo, ["场站", "场站名称", "充电站", "站点"])

        except Exception as e:
            QMessageBox.warning(self, "读取失败", f"无法读取文件列名：{str(e)}")
            logger.error(f"读取列名失败 {file_path}: {e}")

    def _set_combo_default(self, combo, default_candidates):
        if isinstance(default_candidates, str):
            default_candidates = [default_candidates]
        for candidate in default_candidates:
            candidate_str = str(candidate)
            index = combo.findText(candidate_str, Qt.MatchFixedString)
            if index >= 0:
                combo.setCurrentIndex(index)
                return
        if combo.count() > 0:
            combo.setCurrentIndex(0)

    # ---------- 外部匹配 ----------
    def _on_match_toggled(self, checked):
        self.match_group.setEnabled(checked)
        if not checked:
            self.match_file_path = ""
            self.match_file_label.setText("未选择")
            self.match_col_combo.clear()

    def _select_match_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "选择匹配文件", "", "Excel文件 (*.xlsx *.xls)"
        )
        if path:
            self.match_file_path = path
            self.match_file_label.setText(path)
            try:
                if path.endswith('.xls'):
                    df = pd.read_excel(path, nrows=0, engine='xlrd')
                else:
                    df = pd.read_excel(path, nrows=0, engine='openpyxl')
                cols = df.columns.tolist()
                self.match_col_combo.clear()
                self.match_col_combo.addItems(cols)
                idx = self.match_col_combo.findText("场站", Qt.MatchFixedString)
                if idx >= 0:
                    self.match_col_combo.setCurrentIndex(idx)
            except Exception as e:
                QMessageBox.warning(self, "读取失败", f"无法读取匹配文件列名：{str(e)}")
                logger.error(f"读取匹配文件列名失败 {path}: {e}")

    # ---------- 输出路径 ----------
    def _select_output(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "保存对账结果", "", "Excel文件 (*.xlsx);;Excel 97-2003 (*.xls)"
        )
        if path:
            if not (path.endswith('.xlsx') or path.endswith('.xls')):
                path += '.xlsx'
            self.output_path = path
            self.output_label.setText(path)

    # ---------- 界面状态更新 ----------
    def _update_ui_state(self):
        has_left = bool(self.left_file_path)
        has_right = bool(self.right_file_path)
        self.start_btn.setEnabled(has_left and has_right)

    # ---------- 执行任务 ----------
    @Slot()
    def _on_start(self):
        # 参数校验
        if not self.left_file_path or not self.right_file_path:
            QMessageBox.warning(self, "警告", "请上传左右两侧文件")
            return

        group_col = self.group_combo.currentText()
        if not group_col:
            QMessageBox.warning(self, "警告", "请选择充电场站名称列")
            return

        left_energy = self.left_energy_combo.currentText()
        left_fee = self.left_fee_combo.currentText()
        if not left_energy or not left_fee:
            QMessageBox.warning(self, "警告", "请完整选择达克云平台的电量列和费用列")
            return

        right_energy = self.right_energy_combo.currentText()
        right_fee = self.right_fee_combo.currentText()
        if not right_energy or not right_fee:
            QMessageBox.warning(self, "警告", "请完整选择海汇德平台的电量列和费用列")
            return

        # 外部匹配配置
        match_config = None
        if self.match_checkbox.isChecked():
            if not self.match_file_path:
                QMessageBox.warning(self, "警告", "请上传匹配文件")
                return
            match_col = self.match_col_combo.currentText()
            if not match_col:
                QMessageBox.warning(self, "警告", "请选择外部文件中的场站列")
                return
            match_config = {
                'match_file': self.match_file_path,
                'match_col': match_col,
                'mode': 'remove'
            }

        # 输出路径
        if not self.output_path:
            import datetime
            default_name = f"对账结果_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            self.output_path = os.path.join(os.path.expanduser("~"), "Desktop", default_name)
            self.output_label.setText(self.output_path)

        # 显示进度条和提示文字
        self.progress_bar.setValue(0)
        self.progress_bar.show()
        self.progress_label.show()
        self.start_btn.setEnabled(False)
        self.log_text.append("开始对账处理...")

        self.controller.start_reconciliation_task(
            left_file=self.left_file_path,
            right_file=self.right_file_path,
            group_col=group_col,
            left_energy_col=left_energy,
            left_fee_col=left_fee,
            right_energy_col=right_energy,
            right_fee_col=right_fee,
            match_config=match_config,
            output_path=self.output_path,
            message_callback=self._on_message,
            finished_callback=self._on_finished,
            error_callback=self._on_error,
            progress_callback=self._on_progress
        )

    @Slot(str)
    def _on_message(self, msg):
        self.log_text.append(msg)
        self.status_bar.showMessage(msg, 3000)

    @Slot(int)
    def _on_progress(self, value):
        """更新进度条和提示文字"""
        self.progress_bar.setValue(value)
        self.progress_label.setText(f"正在拼命处理中，请耐心等待。。。 {value}%")

    @Slot(object)
    def _on_finished(self, result):
        self.start_btn.setEnabled(True)
        self.progress_bar.hide()
        self.progress_label.hide()
        self.log_text.append(f"对账完成！文件已保存至：{result}")
        QMessageBox.information(self, "完成", f"对账成功！\n输出文件：{result}")

    @Slot(str)
    def _on_error(self, err_msg):
        self.start_btn.setEnabled(True)
        self.progress_bar.hide()
        self.progress_label.hide()
        self.log_text.append(f"错误：{err_msg}")
        QMessageBox.critical(self, "错误", f"对账任务失败：\n{err_msg}")