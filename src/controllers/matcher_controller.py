# 定义匹配任务线程，继承WorkerThread，并通过MainController调度。
# src/controllers/matcher_controller.py

from src.utils.worker_thread import WorkerThread
from src.models.excel_matcher import ExcelMatcher
from src.utils.logger import setup_logger

logger = setup_logger(__name__)


class MatchTaskThread(WorkerThread):
    """数据匹配合并线程"""
    def __init__(self, file_a, file_b, key_a, key_b, cols_b, how, output):
        super().__init__()
        self.file_a = file_a
        self.file_b = file_b
        self.key_a = key_a
        self.key_b = key_b
        self.cols_b = cols_b
        self.how = how
        self.output = output

    def run(self):
        try:
            self.signals.message.emit("开始读取文件...")
            result = ExcelMatcher.match_and_merge(
                file_a=self.file_a,
                file_b=self.file_b,
                key_columns_a=self.key_a,
                key_columns_b=self.key_b,
                columns_b_to_add=self.cols_b,
                how=self.how,
                output_path=self.output
            )
            self.signals.finished.emit(result)
        except Exception as e:
            logger.exception("匹配任务失败")
            self.signals.error.emit(str(e))


class MatcherController:
    """匹配任务控制器（可集成进MainController或独立）"""
    def __init__(self):
        self.current_thread = None

    def start_match_task(self, file_a, file_b, key_a, key_b, cols_b, how, output,
                         message_callback, finished_callback, error_callback):
        # 停止已有任务
        if self.current_thread and self.current_thread.isRunning():
            self.current_thread.stop()
            self.current_thread.wait()

        self.current_thread = MatchTaskThread(
            file_a, file_b, key_a, key_b, cols_b, how, output
        )
        self.current_thread.signals.message.connect(message_callback)
        self.current_thread.signals.finished.connect(finished_callback)
        self.current_thread.signals.error.connect(error_callback)
        self.current_thread.start()