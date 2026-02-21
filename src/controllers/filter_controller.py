# src/controllers/filter_controller.py

from src.utils.worker_thread import WorkerThread
from src.models.excel_filter import ExcelFilter
from src.utils.logger import setup_logger

logger = setup_logger(__name__)


class FilterTaskThread(WorkerThread):
    """多文件筛选任务线程（支持外部匹配）"""
    def __init__(self, file_paths, conditions, sheet_name_col,
                 sum_columns, output_path, match_config=None):
        super().__init__()
        self.file_paths = file_paths
        self.conditions = conditions
        self.sheet_name_col = sheet_name_col
        self.sum_columns = sum_columns
        self.output_path = output_path
        self.match_config = match_config  # 支持外部匹配

    def run(self):
        try:
            self.signals.message.emit("开始批量处理文件...")
            result = ExcelFilter.filter_and_export(
                file_paths=self.file_paths,
                conditions=self.conditions,
                sheet_name_col=self.sheet_name_col,
                sum_columns=self.sum_columns,
                output_path=self.output_path,
                match_config=self.match_config,
                progress_callback=self.signals.progress.emit
            )
            self.signals.finished.emit(result)
        except Exception as e:
            logger.exception("筛选任务失败")
            self.signals.error.emit(str(e))


class FilterController:
    """筛选任务控制器"""
    def __init__(self):
        self.current_thread = None

    def start_filter_task(self, file_paths, conditions, sheet_name_col,
                          sum_columns, output_path,
                          match_config=None,
                          message_callback = None, finished_callback=None,
                          error_callback=None, progress_callback=None):
        # 终止正在运行的任务
        if self.current_thread and self.current_thread.isRunning():
            self.current_thread.stop()
            self.current_thread.wait()

        self.current_thread = FilterTaskThread(
            file_paths, conditions, sheet_name_col, sum_columns, output_path, match_config  # 传递
        )
        self.current_thread.signals.message.connect(message_callback)
        self.current_thread.signals.finished.connect(finished_callback)
        self.current_thread.signals.error.connect(error_callback)
        if progress_callback:
            self.current_thread.signals.progress.connect(progress_callback)
        self.current_thread.start()