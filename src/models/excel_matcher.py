# 使用pandas的merge方法，支持多列匹配、自定义连接方式、选择输出列。
# src/models/excel_matcher.py

import pandas as pd
from pathlib import Path
from src.utils.logger import setup_logger

logger = setup_logger(__name__)


class ExcelMatcher:
    """Excel数据匹配合并业务类"""

    @staticmethod
    def match_and_merge(
        file_a: str,
        file_b: str,
        key_columns_a: list,
        key_columns_b: list,
        columns_b_to_add: list = None,
        how: str = 'left',
        output_path: str = None
    ) -> str:
        """
        根据条件匹配两个Excel文件，并将表B指定列合并到表A
        :param file_a: 表A文件路径
        :param file_b: 表B文件路径
        :param key_columns_a: 表A中用于匹配的列名列表
        :param key_columns_b: 表B中用于匹配的列名列表（长度需与key_columns_a一致）
        :param columns_b_to_add: 需要添加到表A的表B列名列表，None表示除匹配列外全部添加
        :param how: 连接方式，left/right/inner/outer，默认left
        :param output_path: 输出文件路径
        :return: 输出文件路径
        """
        # 读取Excel
        df_a = ExcelMatcher._read_excel(file_a)
        df_b = ExcelMatcher._read_excel(file_b)

        # 校验列名是否存在
        ExcelMatcher._validate_columns(df_a, key_columns_a, file_a)
        ExcelMatcher._validate_columns(df_b, key_columns_b, file_b)

        # 确定要合并的列
        if columns_b_to_add is None:
            # 默认添加表B中除匹配键以外的所有列
            columns_b_to_add = [col for col in df_b.columns if col not in key_columns_b]
        else:
            # 确保请求的列存在于表B中
            ExcelMatcher._validate_columns(df_b, columns_b_to_add, file_b)

        # 构建合并用的两个DataFrame
        left = df_a
        right = df_b[key_columns_b + columns_b_to_add]  # 只保留需要的列

        # 执行合并
        result = pd.merge(
            left, right,
            left_on=key_columns_a,
            right_on=key_columns_b,
            how=how,
            suffixes=('', '_fromB')  # 如有重复列名自动添加后缀
        )

        # 保存结果
        output_path = output_path or str(Path(file_a).parent / f"{Path(file_a).stem}_matched.xlsx")
        result.to_excel(output_path, index=False)
        logger.info(f"匹配合并完成，输出文件：{output_path}")
        return output_path

    @staticmethod
    def _read_excel(file_path):
        """统一读取Excel，自动识别引擎"""
        if str(file_path).endswith('.xls'):
            return pd.read_excel(file_path, engine='xlrd')
        else:
            return pd.read_excel(file_path, engine='openpyxl')

    @staticmethod
    def _validate_columns(df, columns, file_name):
        """校验列名是否存在，不存在则抛出ValueError"""
        missing = [col for col in columns if col not in df.columns]
        if missing:
            raise ValueError(f"文件 {file_name} 中缺少列: {missing}")