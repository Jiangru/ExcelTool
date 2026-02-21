# src/views/tabs/filter_tab.py

import os
import pandas as pd
from PySide6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout,
                               QPushButton, QFileDialog, QLabel, QRadioButton,
                               QListWidget, QListWidgetItem, QComboBox,
                               QTableWidget, QTableWidgetItem, QHeaderView,
                               QGroupBox, QLineEdit, QMessageBox,
                               QProgressBar, QAbstractItemView, QCheckBox)
from PySide6.QtCore import Slot, Qt
from src.utils.logger import setup_logger

logger = setup_logger(__name__)


class FilterTab(QWidget):
    """多文件筛选与汇总标签页"""
    def __init__(self, controller, status_bar, log_text, progress_bar):
        super().__init__()
        self.controller = controller
        self.status_bar = status_bar
        self.log_text = log_text
        self.progress_bar = progress_bar

        self.file_paths = []          # 已选择的文件列表
        self.all_columns = []         # 所有文件的列名并集（从第一个文件读取）
        self.output_path = ""

        self.setup_ui()
        self._update_ui_state()
        self.match_file_path = None

    def setup_ui(self):
        layout = QVBoxLayout(self)

        # ---------- 1. 文件选择区 ----------
        file_group = QGroupBox("1. 选择多个Excel文件")
        file_layout = QHBoxLayout()

        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        file_layout.addWidget(self.file_list)

        btn_layout = QVBoxLayout()
        self.btn_add_files = QPushButton("添加文件")
        self.btn_add_files.clicked.connect(self._select_files)
        self.btn_remove_selected = QPushButton("移除选中")
        self.btn_remove_selected.clicked.connect(self._remove_selected_files)
        self.btn_clear_all = QPushButton("清空所有")
        self.btn_clear_all.clicked.connect(self._clear_all_files)
        btn_layout.addWidget(self.btn_add_files)
        btn_layout.addWidget(self.btn_remove_selected)
        btn_layout.addWidget(self.btn_clear_all)
        btn_layout.addStretch()
        file_layout.addLayout(btn_layout)

        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # ---------- 2. 筛选条件设置 ----------
        cond_group = QGroupBox("2. 设置筛选条件（多个条件之间为“与”关系）")
        cond_layout = QVBoxLayout()

        # ---------- 2.5 外部文件匹配条件（新增）----------
        self.match_group = QGroupBox("外部文件匹配条件（可选）")
        match_layout = QVBoxLayout()

        # 启用复选框
        self.chk_enable_match = QCheckBox("启用外部文件匹配")
        self.chk_enable_match.toggled.connect(self._on_match_enable_toggled)
        match_layout.addWidget(self.chk_enable_match)

        # 匹配文件选择
        file_match_layout = QHBoxLayout()
        file_match_layout.addWidget(QLabel("匹配文件："))
        self.label_match_file = QLabel("未选择")
        file_match_layout.addWidget(self.label_match_file)
        self.btn_match_file = QPushButton("浏览...")
        self.btn_match_file.clicked.connect(self._select_match_file)
        file_match_layout.addWidget(self.btn_match_file)
        file_match_layout.addStretch()
        match_layout.addLayout(file_match_layout)

        # 列选择
        col_match_layout = QHBoxLayout()
        col_match_layout.addWidget(QLabel("原文件匹配列："))
        self.combo_match_source = QComboBox()
        self.combo_match_source.setMinimumWidth(250)      # 增加宽度，可自行调整数值
        col_match_layout.addWidget(self.combo_match_source)
        col_match_layout.addWidget(QLabel("匹配文件列："))
        self.combo_match_target = QComboBox()
        self.combo_match_target.setMinimumWidth(250)      # 同样增加宽度
        col_match_layout.addWidget(self.combo_match_target)
        col_match_layout.addStretch()
        match_layout.addLayout(col_match_layout)

        # 模式选择（保留/排除）
        mode_layout = QHBoxLayout()
        self.radio_keep = QRadioButton("保留匹配的数据（白名单）")
        self.radio_keep.setChecked(True)
        self.radio_remove = QRadioButton("排除匹配的数据（黑名单）")
        mode_layout.addWidget(self.radio_keep)
        mode_layout.addWidget(self.radio_remove)
        mode_layout.addStretch()
        match_layout.addLayout(mode_layout)

        self.match_group.setLayout(match_layout)
        layout.insertWidget(layout.indexOf(cond_group) + 1, self.match_group)  # 插入在条件组之后

        # 初始化控件状态（默认禁用）
        self._set_match_controls_enabled(False)

        # 条件表格
        self.cond_table = QTableWidget()
        self.cond_table.setColumnCount(4)
        self.cond_table.setHorizontalHeaderLabels(["列名", "运算符", "值", "操作"])
        self.cond_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.cond_table.setEditTriggers(QTableWidget.NoEditTriggers)  # 部分列需自定义控件
        cond_layout.addWidget(self.cond_table)

        # 添加条件按钮
        self.btn_add_condition = QPushButton("+ 添加条件")
        self.btn_add_condition.clicked.connect(self._add_condition_row)
        cond_layout.addWidget(self.btn_add_condition)

        cond_group.setLayout(cond_layout)
        layout.addWidget(cond_group)

        # ---------- 3. Sheet命名列选择 ----------
        name_group = QGroupBox("3. Sheet命名列（取该列第一个非空值作为工作表名称）")
        name_layout = QHBoxLayout()
        name_layout.addWidget(QLabel("选择列："))
        self.combo_sheet_name = QComboBox()
        self.combo_sheet_name.setMinimumWidth(200)
        name_layout.addWidget(self.combo_sheet_name)
        name_layout.addStretch()
        name_group.setLayout(name_layout)
        layout.addWidget(name_group)

        # ---------- 4. 求和列选择 ----------
        sum_group = QGroupBox("4. 选择需要格式化为数字并求和的列（可多选）")
        sum_layout = QVBoxLayout()
        self.sum_list = QListWidget()
        self.sum_list.setSelectionMode(QAbstractItemView.MultiSelection)
        sum_layout.addWidget(self.sum_list)
        btn_sel_layout = QHBoxLayout()
        self.btn_select_all_sum = QPushButton("全选")
        self.btn_select_all_sum.clicked.connect(self._select_all_sum)
        self.btn_clear_sum = QPushButton("清空")
        self.btn_clear_sum.clicked.connect(self._clear_sum)
        btn_sel_layout.addWidget(self.btn_select_all_sum)
        btn_sel_layout.addWidget(self.btn_clear_sum)
        btn_sel_layout.addStretch()
        sum_layout.addLayout(btn_sel_layout)
        sum_group.setLayout(sum_layout)
        layout.addWidget(sum_group)

        # ---------- 5. 输出设置 ----------
        out_group = QGroupBox("5. 输出文件")
        out_layout = QHBoxLayout()
        out_layout.addWidget(QLabel("保存路径："))
        self.label_output = QLabel("未选择，将自动生成")
        out_layout.addWidget(self.label_output)
        self.btn_output = QPushButton("浏览...")
        self.btn_output.clicked.connect(self._select_output)
        out_layout.addWidget(self.btn_output)
        out_layout.addStretch()
        out_group.setLayout(out_layout)
        layout.addWidget(out_group)

        # ---------- 6. 执行按钮 ----------
        self.btn_start = QPushButton("开始筛选与汇总")
        self.btn_start.clicked.connect(self._on_start)
        self.btn_start.setMinimumHeight(40)
        layout.addWidget(self.btn_start)

        layout.addStretch()

        # 初始化添加一行默认条件
        # self._add_condition_row()

    # ---------- 文件操作 ----------
    def _select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择Excel文件", "",
            "Excel文件 (*.xlsx *.xls);;所有文件 (*.*)"
        )
        if files:
            for f in files:
                # 避免重复添加
                existing = [self.file_list.item(i).text() for i in range(self.file_list.count())]
                if f not in existing:
                    self.file_list.addItem(f)
                    self.file_paths.append(f)
            # 读取第一个文件的列名，更新所有列下拉框
            if self.file_paths:
                self._load_columns_from_first_file()
            self._update_ui_state()

    def _remove_selected_files(self):
        selected = self.file_list.selectedItems()
        for item in selected:
            row = self.file_list.row(item)
            self.file_list.takeItem(row)
            self.file_paths.pop(row)
        self._update_columns_after_file_change()
        self._update_ui_state()

    def _clear_all_files(self):
        self.file_list.clear()
        self.file_paths.clear()
        self._update_columns_after_file_change()
        self._update_ui_state()

    def _set_match_controls_enabled(self, enabled):
        """设置外部匹配相关控件的启用状态"""
        self.btn_match_file.setEnabled(enabled)
        self.combo_match_source.setEnabled(enabled)
        self.combo_match_target.setEnabled(enabled)
        self.radio_keep.setEnabled(enabled)
        self.radio_remove.setEnabled(enabled)
        if not enabled:
            self.label_match_file.setText("未选择")
            self.combo_match_source.clear()
            self.combo_match_target.clear()
            self.match_file_path = None  # 新增属性，用于存储匹配文件路径

    def _on_match_enable_toggled(self, checked):
        """启用复选框切换"""
        self._set_match_controls_enabled(checked)
        if checked:
            # 如果之前没有加载过列名，可能需要刷新原文件列名
            self._refresh_match_source_columns()
        else:
            # 清空匹配文件相关数据
            self.match_file_path = None
            self.combo_match_target.clear()

    def _refresh_match_source_columns(self):
        """刷新原文件匹配列下拉框（复用 all_columns）"""
        self.combo_match_source.clear()
        self.combo_match_source.addItems(self.all_columns)
        if self.all_columns:
            self.combo_match_source.setCurrentIndex(0)

    def _select_match_file(self):
        """选择匹配Excel文件"""
        path, _ = QFileDialog.getOpenFileName(
            self, "选择匹配参考文件", "",
            "Excel文件 (*.xlsx *.xls);;所有文件 (*.*)"
        )
        if path:
            self.match_file_path = path
            self.label_match_file.setText(path)
            self.chk_enable_match.setChecked(True)  # 自动启用
            # 读取匹配文件的列名，填充 target 下拉框
            self._load_match_file_columns(path)

    def _load_match_file_columns(self, file_path):
        """读取匹配文件的列名"""
        try:
            if file_path.endswith('.xls'):
                df = pd.read_excel(file_path, nrows=0, engine='xlrd')
            else:
                df = pd.read_excel(file_path, nrows=0, engine='openpyxl')
            columns = df.columns.tolist()
            self.combo_match_target.clear()
            self.combo_match_target.addItems(columns)
            if columns:
                self.combo_match_target.setCurrentIndex(0)
        except Exception as e:
            QMessageBox.warning(self, "读取失败", f"无法读取匹配文件列名：{str(e)}")
            logger.error(f"读取匹配文件列名失败 {file_path}: {e}")

    def _load_columns_from_first_file(self):
        """从第一个文件中读取列名，并填充到列名下啦、sheet命名列、求和列列表"""
        if not self.file_paths:
            return
        try:
            first_file = self.file_paths[0]
            if first_file.endswith('.xls'):
                df = pd.read_excel(first_file, nrows=0, engine='xlrd')
            else:
                df = pd.read_excel(first_file, nrows=0, engine='openpyxl')
            self.all_columns = df.columns.tolist()

            # 更新条件行中的所有列名下啦
            self._refresh_condition_columns()

            # 更新sheet命名列下拉框
            self.combo_sheet_name.clear()
            self.combo_sheet_name.addItems(self.all_columns)
            if self.all_columns:
                self.combo_sheet_name.setCurrentIndex(0)

            # 更新求和列列表
            self.sum_list.clear()
            self.sum_list.addItems(self.all_columns)

            if self.chk_enable_match.isChecked():
                self._refresh_match_source_columns()

        except Exception as e:
            QMessageBox.warning(self, "读取失败", f"无法读取文件列名：{str(e)}")
            logger.error(f"读取列名失败 {first_file}: {e}")

    def _update_columns_after_file_change(self):
        """文件列表变化后，重新加载列名（若还有文件）"""
        if self.file_paths:
            self._load_columns_from_first_file()
        else:
            # 无文件，清空相关控件
            self.all_columns = []
            self._refresh_condition_columns()
            self.combo_sheet_name.clear()
            self.sum_list.clear()
        if self.chk_enable_match.isChecked() and self.file_paths:
            self._refresh_match_source_columns()

    def _refresh_condition_columns(self):
        """刷新所有条件行中的列名下啦"""
        for row in range(self.cond_table.rowCount()):
            combo = self.cond_table.cellWidget(row, 0)
            if combo and isinstance(combo, QComboBox):
                current_text = combo.currentText()
                combo.clear()
                combo.addItems(self.all_columns)
                # 恢复之前选中的列（如果存在）
                if current_text in self.all_columns:
                    combo.setCurrentText(current_text)

    # ---------- 条件行管理 ----------
    def _add_condition_row(self):
        row = self.cond_table.rowCount()
        self.cond_table.insertRow(row)

        # 列名下拉框
        combo_col = QComboBox()
        combo_col.addItems(self.all_columns)
        combo_col.setEditable(False)
        self.cond_table.setCellWidget(row, 0, combo_col)

        # 运算符下拉框
        combo_op = QComboBox()
        operators = ['等于', '不等于', '大于', '大于等于', '小于', '小于等于',
                     '包含', '不包含', '为空', '不为空']
        combo_op.addItems(operators)
        self.cond_table.setCellWidget(row, 1, combo_op)

        # 值输入框（为空/不为空时禁用）
        line_edit = QLineEdit()
        self.cond_table.setCellWidget(row, 2, line_edit)

        # 删除按钮
        btn_del = QPushButton("删除")
        btn_del.clicked.connect(lambda: self._delete_condition_row(row))
        self.cond_table.setCellWidget(row, 3, btn_del)

        # 关联运算符变化事件，控制值输入框的可用性
        combo_op.currentTextChanged.connect(
            lambda text, le=line_edit: self._on_operator_changed(text, le)
        )
        # 初始触发一次
        self._on_operator_changed(combo_op.currentText(), line_edit)

    def _delete_condition_row(self, row):
        self.cond_table.removeRow(row)

    def _on_operator_changed(self, op_text, line_edit):
        """运算符改变时，值为空/不为空时禁用输入框"""
        if op_text in ['为空', '不为空']:
            line_edit.setEnabled(False)
            line_edit.clear()
        else:
            line_edit.setEnabled(True)

    # ---------- 求和列全选/清空 ----------
    def _select_all_sum(self):
        self.sum_list.selectAll()

    def _clear_sum(self):
        self.sum_list.clearSelection()

    # ---------- 输出路径 ----------
    def _select_output(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "保存筛选汇总结果", "",
            "Excel文件 (*.xlsx);;Excel 97-2003 (*.xls)"
        )
        if path:
            if not (path.endswith('.xlsx') or path.endswith('.xls')):
                path += '.xlsx'
            self.output_path = path
            self.label_output.setText(path)

    # ---------- 界面状态更新 ----------
    def _update_ui_state(self):
        """根据文件有无启用/禁用相关控件"""
        has_files = bool(self.file_paths)
        self.btn_remove_selected.setEnabled(has_files)
        self.btn_clear_all.setEnabled(has_files)
        self.cond_table.setEnabled(has_files)
        self.btn_add_condition.setEnabled(has_files)
        self.combo_sheet_name.setEnabled(has_files)
        self.sum_list.setEnabled(has_files)
        self.btn_select_all_sum.setEnabled(has_files)
        self.btn_clear_sum.setEnabled(has_files)
        self.btn_start.setEnabled(has_files)

    # ---------- 执行任务 ----------
    @Slot()
    def _on_start(self):
        """启动筛选任务"""
        # 1. 校验参数
        if not self.file_paths:
            QMessageBox.warning(self, "警告", "请至少选择一个Excel文件")
            return

        # 2. 收集筛选条件
        conditions = []
        for row in range(self.cond_table.rowCount()):
            col_widget = self.cond_table.cellWidget(row, 0)
            op_widget = self.cond_table.cellWidget(row, 1)
            val_widget = self.cond_table.cellWidget(row, 2)

            col = col_widget.currentText() if col_widget else ""
            op = op_widget.currentText() if op_widget else ""
            val = val_widget.text() if val_widget else ""

            if not col:
                continue
            # 值为空且运算符不是为空/不为空时，跳过该条件
            if op not in ['为空', '不为空'] and not val:
                msg = f"第{row+1}行条件「{col} {op}」未填写值，已自动忽略"
                self.log_text.append(f"⚠️ {msg}")
                logger.warning(msg)
                continue

            conditions.append({
                'column': col,
                'operator': op,
                'value': val,
                'logic': 'AND'   # 当前版本固定AND
            })

        # 3. sheet命名列
        sheet_name_col = self.combo_sheet_name.currentText()
        if not sheet_name_col:
            QMessageBox.warning(self, "警告", "请选择Sheet命名列")
            return

        # 4. 求和列
        selected_items = self.sum_list.selectedItems()
        sum_columns = [item.text() for item in selected_items]

        # 5. 输出路径
        if not self.output_path:
            # 自动生成路径
            import datetime
            default_name = f"筛选汇总_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            self.output_path = os.path.join(os.path.expanduser("~"), "Desktop", default_name)
            self.label_output.setText(self.output_path)

        # --- 收集外部匹配参数（如果启用）---
        match_config = None
        if self.chk_enable_match.isChecked():
            print("开始匹配外部文件咯")
            # 校验参数完整性
            if not hasattr(self, 'match_file_path') or not self.match_file_path:
                QMessageBox.warning(self, "警告", "请选择匹配文件")
                return
            source_col = self.combo_match_source.currentText()
            if not source_col:
                QMessageBox.warning(self, "警告", "请选择原文件匹配列")
                return
            target_col = self.combo_match_target.currentText()
            if not target_col:
                QMessageBox.warning(self, "警告", "请选择匹配文件列")
                return
            # 新增：快速预读匹配文件，检查目标列是否有有效值
            try:
                df_temp = pd.read_excel(self.match_file_path, nrows=100, dtype=str)
                if self.combo_match_target.currentText() not in df_temp.columns:
                    QMessageBox.critical(self, "错误", "匹配文件中不存在所选列")
                    return
                non_null_count = df_temp[self.combo_match_target.currentText()].notna().sum()
                if non_null_count == 0:
                    reply = QMessageBox.question(
                        self, "警告",
                        "匹配文件所选列没有有效数据，继续执行将：\n"
                        "- 白名单模式 → 无数据输出\n"
                        "- 黑名单模式 → 保留全部数据\n"
                        "是否继续？",
                        QMessageBox.Yes | QMessageBox.No
                    )
                    if reply == QMessageBox.No:
                        return
            except Exception as e:
                QMessageBox.critical(self, "错误", f"读取匹配文件失败：{str(e)}")
                return
            mode = 'keep' if self.radio_keep.isChecked() else 'remove'
            match_config = {
                'match_file': self.match_file_path,
                'source_column': source_col,
                'target_column': target_col,
                'mode': mode
            }

        # 6. 禁用按钮，显示进度
        self.btn_start.setEnabled(False)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.show()
        self.log_text.append("开始批量筛选与汇总...")

        # 7. 调用控制器
        self.controller.start_filter_task(
            file_paths=self.file_paths,
            conditions=conditions,
            sheet_name_col=sheet_name_col,
            sum_columns=sum_columns,
            output_path=self.output_path,
            match_config=match_config,
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
        try:
            self.progress_bar.setValue(int(value))
        except Exception as e:
            logger.error(f"进度条设置失败: {e}, value={value}")
            self.progress_bar.setValue(0)

    @Slot(object)
    def _on_finished(self, result):
        self.btn_start.setEnabled(True)
        self.progress_bar.hide()
        self.log_text.append(f"筛选汇总完成！文件已保存至：{result}")
        QMessageBox.information(self, "完成", f"批量处理成功！\n输出文件：{result}")

    @Slot(str)
    def _on_error(self, err_msg):
        self.btn_start.setEnabled(True)
        self.progress_bar.hide()
        self.log_text.append(f"错误：{err_msg}")
        QMessageBox.critical(self, "错误", f"任务执行失败：\n{err_msg}")