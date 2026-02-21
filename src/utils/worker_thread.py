# 继承QThread，提供信号机制，将耗时任务放到子线程执行，避免界面卡死。

from PySide6.QtCore import QThread, Signal


class WorkerSignals(QThread):
    """线程通信信号集"""
    progress = Signal(int)  # 进度百分比
    message = Signal(str)   # 状态信息
    finished = Signal(object)  # 任务完成，返回结果
    error = Signal(str)    # 错误信息


class WorkerThread(QThread):
    """通用工作线程基类，使用时只需继承并重写run方法"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.signals = WorkerSignals()
        self._is_running = True

    def stop(self):
        """提供停止线程的接口"""
        self._is_running = False

    def run(self):
        """子类必须重写此方法，实现具体业务逻辑"""
        pass