# 负责读取/写入 config/settings.ini，提供全局配置访问。

import configparser
import os
from pathlib import Path


class ConfigManager:
    """配置文件管理器（单例模式）"""
    _instance = None

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self):
        if not hasattr(self, 'initialized'):
            self.initialized = True
            self.config = configparser.ConfigParser()
            self.config_path = Path(__file__).parent.parent.parent / "config" / "settings.ini"
            self._load_config()

    def _load_config(self):
        """加载配置文件，若不存在则创建默认配置"""
        if not self.config_path.exists():
            self._create_default_config()
        self.config.read(self.config_path, encoding='utf-8')

    def _create_default_config(self):
        """生成默认配置文件"""
        self.config['APP'] = {
            'version': '1.0.0',
            'window_width': '1000',
            'window_height': '700'
        }
        self.config['EXCEL'] = {
            'default_output_folder': './output',
            'max_preview_rows': '100'
        }
        self.config['LOG'] = {
            'level': 'INFO',
            'max_bytes': '1048576',  # 1MB
            'backup_count': '5'
        }
        # 确保配置目录存在
        self.config_path.parent.mkdir(parents=True, exist_ok=True)
        with open(self.config_path, 'w', encoding='utf-8') as f:
            self.config.write(f)

    def get(self, section, option, fallback=None):
        """获取配置值（字符串）"""
        return self.config.get(section, option, fallback=fallback)

    def getint(self, section, option, fallback=None):
        return self.config.getint(section, option, fallback=fallback)

    def getboolean(self, section, option, fallback=None):
        return self.config.getboolean(section, option, fallback=fallback)

    def set(self, section, option, value):
        """设置配置值并保存到文件"""
        if not self.config.has_section(section):
            self.config.add_section(section)
        self.config.set(section, option, str(value))
        with open(self.config_path, 'w', encoding='utf-8') as f:
            self.config.write(f)