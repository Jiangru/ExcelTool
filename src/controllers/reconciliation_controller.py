# src/controllers/reconciliation_controller.py

from src.utils.worker_thread import WorkerThread
from src.models.excel_reconciliation import ExcelReconciliation
from src.utils.logger import setup_logger

logger = setup_logger(__name__)


class ReconciliationTaskThread(WorkerThread):
    """对账任务线程"""
    def __init__(self, left_file, right_file,
                 group_col,
                 left_energy_col, left_fee_col,
                 right_energy_col, right_fee_col,
                 match_config,
                 output_path):
        super().__init__()
        self.left_file = left_file
        self.right_file = right_file
        self.group_col = group_col
        self.left_energy_col = left_energy_col
        self.left_fee_col = left_fee_col
        self.right_energy_col = right_energy_col
        self.right_fee_col = right_fee_col
        self.match_config = match_config
        self.output_path = output_path

    def run(self):
        try:
            self.signals.message.emit("开始对账处理...")
            # 将线程的进度信号作为回调传递给业务函数
            result = ExcelReconciliation.reconcile(
                left_file=self.left_file,
                right_file=self.right_file,
                group_col=self.group_col,
                left_energy_col=self.left_energy_col,
                left_fee_col=self.left_fee_col,
                right_energy_col=self.right_energy_col,
                right_fee_col=self.right_fee_col,
                match_config=self.match_config,
                output_path=self.output_path,
                progress_callback=self.signals.progress.emit   # 传递进度回调
            )
            self.signals.finished.emit(result)
        except Exception as e:
            logger.exception("对账任务失败")
            self.signals.error.emit(str(e))


class ReconciliationController:
    """对账任务控制器"""
    def __init__(self):
        self.current_thread = None

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
        if self.current_thread and self.current_thread.isRunning():
            self.current_thread.stop()
            self.current_thread.wait()

        self.current_thread = ReconciliationTaskThread(
            left_file, right_file,
            group_col,
            left_energy_col, left_fee_col,
            right_energy_col, right_fee_col,
            match_config,
            output_path
        )
        self.current_thread.signals.message.connect(message_callback)
        self.current_thread.signals.finished.connect(finished_callback)
        self.current_thread.signals.error.connect(error_callback)
        if progress_callback:
            self.current_thread.signals.progress.connect(progress_callback)
        self.current_thread.start()