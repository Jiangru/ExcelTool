# 负责初始化应用、加载配置、启动主窗口。
import sys
import os

# 将项目根目录加入模块搜索路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PySide6.QtWidgets import QApplication
from PySide6.QtCore import Qt
from src.views.main_window import MainWindow
from src.utils.logger import setup_logger
from src.utils.config_manager import ConfigManager


def main():
    # 初始化日志
    logger = setup_logger()
    logger.info("程序启动")

    # 高DPI适配
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    # QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    # QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps)

    app = QApplication(sys.argv)
    app.processEvents()   # 使画面及时显示

    # 加载配置文件（全局唯一实例）
    config = ConfigManager()

    # 创建主窗口
    window = MainWindow(config)
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()