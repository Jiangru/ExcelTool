# 这里演示一个简单的Excel合并功能，实际开发可根据需要扩展。

import pandas as pd
from pathlib import Path
from src.utils.logger import setup_logger

logger = setup_logger(__name__)


class ExcelMergeWorker:
    """Excel文件合并业务类"""

    @staticmethod
    def merge_files(file_list, output_path, merge_type='rows'):
        """
        合并多个Excel文件
        :param file_list: 文件路径列表
        :param output_path: 输出文件路径
        :param merge_type: 'rows'纵向合并, 'cols'横向合并
        :return: 成功返回输出路径，失败抛出异常
        """
        all_data = []
        for file in file_list:
            # 根据扩展名选择引擎
            if str(file).endswith('.xls'):
                df = pd.read_excel(file, engine='xlrd')
            else:
                df = pd.read_excel(file, engine='openpyxl')
            all_data.append(df)

        if merge_type == 'rows':
            result = pd.concat(all_data, axis=0, ignore_index=True)
        else:
            result = pd.concat(all_data, axis=1, ignore_index=True)

        # 确保输出目录存在
        output_dir = Path(output_path).parent
        output_dir.mkdir(parents=True, exist_ok=True)

        result.to_excel(output_path, index=False)
        logger.info(f"文件合并成功，保存至: {output_path}")
        return output_path