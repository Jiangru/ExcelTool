# 负责接收界面指令，创建工作线程并调度业务层函数，同时通过信号与界面通信。

from PySide6.QtCore import QObject, Signal
from src.utils.worker_thread import WorkerThread
from src.models.excel_worker import ExcelMergeWorker
from src.utils.logger import setup_logger
from src.controllers.matcher_controller import MatcherController  # 匹配功能
from src.controllers.filter_controller import FilterController # 筛选功能
from src.controllers.reconciliation_controller import ReconciliationController # 绿能对账

logger = setup_logger(__name__)


class MergeTaskThread(WorkerThread):
    """合并任务的具体线程实现"""
    def __init__(self, file_list, output_path, merge_type):
        super().__init__()
        self.file_list = file_list
        self.output_path = output_path
        self.merge_type = merge_type

    def run(self):
        try:
            self.signals.message.emit("开始合并文件...")
            # 这里可以计算进度，简单实现直接调用业务函数
            result = ExcelMergeWorker.merge_files(
                self.file_list, self.output_path, self.merge_type
            )
            self.signals.finished.emit(result)
        except Exception as e:
            logger.exception("合并任务失败")
            self.signals.error.emit(str(e))


class MainController(QObject):
    """主控制器，管理所有业务线程"""
    def __init__(self):
        super().__init__()
        self.current_thread = None
        # 可以持有具体功能的控制器，也可以直接在这里定义方法
        self.matcher_ctrl = MatcherController()  # 实例化匹配控制器
        self.filter_ctrl = FilterController() # 实例化筛选控制器
        self.reconciliation_ctrl = ReconciliationController()   # 实例化绿能对账控制器

    def start_merge_task(self, file_list, output_path, merge_type,
                         progress_callback=None, message_callback=None,
                         finished_callback=None, error_callback=None):
        """启动合并任务"""
        # 如果有正在运行的任务，先停止（实际可做中断处理）
        if self.current_thread and self.current_thread.isRunning():
            self.current_thread.stop()
            self.current_thread.wait()

        self.current_thread = MergeTaskThread(file_list, output_path, merge_type)

        # 连接信号
        if progress_callback:
            self.current_thread.signals.progress.connect(progress_callback)
        if message_callback:
            self.current_thread.signals.message.connect(message_callback)
        if finished_callback:
            self.current_thread.signals.finished.connect(finished_callback)
        if error_callback:
            self.current_thread.signals.error.connect(error_callback)

        self.current_thread.start()

    # 匹配任务
    def start_match_task(self, file_a, file_b, key_a, key_b, cols_b, how, output,
                         message_callback, finished_callback, error_callback):
        """启动匹配任务"""
        self.matcher_ctrl.start_match_task(
            file_a, file_b, key_a, key_b, cols_b, how, output,
            message_callback, finished_callback, error_callback
        )

    # 筛选任务
    def start_filter_task(self, file_paths, conditions, sheet_name_col,
                          sum_columns, output_path, match_config,
                          message_callback, finished_callback,
                          error_callback, progress_callback=None):
        self.filter_ctrl.start_filter_task(
            file_paths=file_paths,
            conditions=conditions,
            sheet_name_col=sheet_name_col,
            sum_columns=sum_columns,
            output_path=output_path,
            match_config=match_config,
            message_callback=message_callback,
            finished_callback=finished_callback,
            error_callback=error_callback,
            progress_callback=progress_callback
        )

        # 绿能对账任务启动方法
    def start_reconciliation_task(self,
                                  left_file, right_file,
                                  group_col,
                                  left_energy_col, left_fee_col,
                                  right_energy_col, right_fee_col,
                                  match_config,
                                  output_path,
                                  message_callback,
                                  finished_callback,
                                  error_callback,
                                  progress_callback=None):
        self.reconciliation_ctrl.start_reconciliation_task(
            left_file, right_file,
            group_col,
            left_energy_col, left_fee_col,
            right_energy_col, right_fee_col,
            match_config,
            output_path,
            message_callback,
            finished_callback,
            error_callback,
            progress_callback
        )