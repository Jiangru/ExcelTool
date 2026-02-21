# src/models/excel_filter.py

import pandas as pd
import numpy as np
from pathlib import Path
from src.utils.logger import setup_logger

logger = setup_logger(__name__)


class ExcelFilter:
    """多文件筛选与汇总业务类"""

    # 运算符映射到pandas查询表达式
    OPERATOR_MAP = {
        '等于': '==',
        '不等于': '!=',
        '大于': '>',
        '大于等于': '>=',
        '小于': '<',
        '小于等于': '<=',
        '包含': 'contains',
        '不包含': 'not contains',
        '为空': 'isnull',
        '不为空': 'notnull'
    }

    @classmethod
    def filter_and_export(cls, file_paths, conditions, sheet_name_col,
                          sum_columns, output_path,
                          match_config=None, # 外部匹配配置
                          progress_callback=None):
        """
        批量筛选文件并导出为多sheet工作簿
        :param file_paths: list of str, 输入文件路径列表
        :param conditions: list of dict, 筛选条件，格式：
            [{'column': '列名', 'operator': '等于', 'value': '值', 'logic': 'AND'}, ...]
            logic 字段目前保留，后续可实现复杂组合，当前所有条件用AND连接
        :param sheet_name_col: str, 用作sheet名称的列名
        :param sum_columns: list of str, 需要格式化为数字并求和的列名
        :param output_path: str, 输出文件路径
        :param progress_callback: function, 进度回调函数，接收当前进度百分比
        :return: str, 输出文件路径
        """
        # 安全包装进度回调
        def safe_progress(value):
            try:
                if progress_callback is not None:
                    int_value = int(value)
                    progress_callback(int_value)
            except Exception as e:
                logger.error(f"进度回调异常: {e}, value={value}", exc_info=True)

        writer = None
        # total_files = len(file_paths)
        has_any_sheet = False   # 新增：标记是否已写入至少一个sheet

        # --- 预处理外部匹配集合（如果启用）---
        match_set = None
        try:
            writer = pd.ExcelWriter(output_path, engine='openpyxl')
            total_files = len(file_paths)
            if match_config:
                try:
                    match_set = cls._load_match_set(
                        match_config['match_file'],
                        match_config['target_column']
                    )
                    logger.info(f"外部匹配集合加载完成，共 {len(match_set)} 个唯一值")
                except Exception as e:
                    logger.error(f"加载匹配文件失败: {e}")
                    raise ValueError(f"匹配文件处理失败: {e}")
            
            for idx, file_path in enumerate(file_paths):
                try:
                    # 读取Excel
                    df = cls._read_excel(file_path)
                    if df.empty:
                        logger.warning(f"文件为空，跳过: {file_path}")
                        continue

                    # 应用常规条件筛选
                    df_filtered = cls._apply_conditions(df, conditions)

                    # --- 应用外部匹配条件（如果有）---
                    if match_set is not None and match_config:
                        source_col = match_config['source_column']
                        mode = match_config['mode']
                        if source_col not in df_filtered.columns:
                            logger.warning(f"原文件缺少匹配列 {source_col}，跳过外部匹配条件")
                        else:
                            # 将原文件匹配列转为字符串（与集合中存储的类型一致）
                            series = df_filtered[source_col].astype(str).str.strip()
                            if mode == 'keep':
                                mask = series.isin(match_set)
                            else:  # 'remove'
                                mask = ~series.isin(match_set)
                            df_filtered = df_filtered[mask]

                    if df_filtered.empty:
                        logger.warning(f"文件筛选后无数据: {file_path}")
                        # 仍然创建空sheet？这里选择跳过
                        continue

                    # 格式化数字列并计算合计
                    df_formatted, sums = cls._format_and_sum(df_filtered, sum_columns)

                    # 确定sheet名称
                    sheet_name = cls._generate_sheet_name(df, sheet_name_col, idx + 1)

                    output_dir = Path(output_path).parent
                    output_dir.mkdir(parents=True, exist_ok=True)

                    # 写入Excel
                    df_formatted.to_excel(writer, sheet_name=sheet_name, index=False)

                    # --- 添加合计行（仅当有求和列时）---
                    if sum_columns:
                        # 获取当前数据表的所有列名（保持顺序）
                        columns = df_formatted.columns.tolist()
                        # 构造一行与表头列数相同的空行
                        total_row = {col: "" for col in columns}
                        
                        # 在第一列写入"合计"标签
                        if columns:
                            total_row[columns[0]] = "合计"
                        
                        # 将计算好的合计值填入对应的求和列
                        for col in sum_columns:
                            if col in total_row and col in sums:
                                total_row[col] = sums[col]
                        
                        # 转换为单行 DataFrame
                        total_df = pd.DataFrame([total_row])
                        
                        # 计算起始行：数据行从第1行开始（第0行为表头），合计行应放在数据最后一行之后
                        startrow = len(df_formatted) + 1
                        
                        # 写入合计行（不写入索引和表头）
                        total_df.to_excel(writer, sheet_name=sheet_name,
                                        startrow=startrow, index=False, header=False)

                    # 标记已写入至少一个sheet
                    has_any_sheet = True

                    logger.info(f"已处理: {file_path} -> Sheet: {sheet_name}")

                except Exception as e:
                    logger.exception(f"处理文件失败 {file_path}: {e}")
                    # 继续处理下一个文件，不中断整个任务
                    continue

                # 更新进度
                if progress_callback:
                    progress = int((idx + 1) / total_files * 100)
                    safe_progress(progress)

            # 如果没有写入任何工作表，创建一个空白工作表
            if not has_any_sheet:
                logger.warning("所有文件均无符合条件的数据，将创建一个空白工作表")
                empty_df = pd.DataFrame()
                empty_df.to_excel(writer, sheet_name="无数据", index=False)
        except Exception as e:
            logger.exception("数据写入过程中发生异常")
            raise  # 重新抛出，由上层线程捕获并触发 error_callback
        finally:
            # 无论是否发生异常，都尝试关闭 writer（保存文件）
            if writer is not None:
                try:
                    writer.close()
                    logger.info(f"ExcelWriter 已关闭，文件保存至: {output_path}")
                except Exception as close_err:
                    logger.error(f"关闭 ExcelWriter 时出错: {close_err}")
                    # 如果文件已成功写入但关闭失败，我们仍向上层报告“成功”，仅记录错误
                    # 但此时无法再抛出异常，因为 finally 中 raise 会覆盖原始异常
                    # 解决方案：将关闭异常暂存，优先抛出数据写入异常
                    if 'e' not in locals():  # 如果没有数据写入异常，则关闭异常成为主要错误
                        raise close_err
                    else:
                        logger.warning("数据已写入，但文件关闭失败，请检查文件是否被占用")

            # writer.close()
            # logger.info(f"筛选汇总完成，输出文件: {output_path}")
        return output_path
    
    @classmethod
    def _load_match_set(cls, match_file, column):
        """读取匹配文件指定列，返回去重后的字符串集合"""
        df = cls._read_excel(match_file)
        if column not in df.columns:
            raise ValueError(f"匹配文件中不存在列: {column}")
        # 去除空值，转为字符串，去重
        values = df[column].dropna().astype(str).str.strip().unique()
        match_set = set(values)
        # 记录日志（前10个值）
        sample = list(match_set)[:10]
        logger.info(f"匹配集合加载完成，共 {len(match_set)} 个唯一值，示例: {sample}")
        
        if len(match_set) == 0:
            logger.warning("匹配集合为空！请检查匹配文件列是否包含有效数据。")
        
        return match_set

    @staticmethod
    def _read_excel(file_path):
        """读取Excel，自动选择引擎"""
        if str(file_path).endswith('.xls'):
            return pd.read_excel(file_path, engine='xlrd')
        else:
            return pd.read_excel(file_path, engine='openpyxl')

    @classmethod
    def _apply_conditions(cls, df, conditions):
        """应用多个筛选条件（AND）"""
        mask = pd.Series([True] * len(df), index=df.index)
        for cond in conditions:
            col = cond['column']
            op = cond['operator']
            val = cond.get('value', '')

            if col not in df.columns:
                logger.warning(f"列 {col} 不存在，跳过该条件")
                continue

            # 根据运算符构造布尔索引
            if op == '等于':
                mask &= (df[col] == val)
            elif op == '不等于':
                mask &= (df[col] != val)
            elif op == '大于':
                mask &= (pd.to_numeric(df[col], errors='coerce') > float(val))
            elif op == '大于等于':
                mask &= (pd.to_numeric(df[col], errors='coerce') >= float(val))
            elif op == '小于':
                mask &= (pd.to_numeric(df[col], errors='coerce') < float(val))
            elif op == '小于等于':
                mask &= (pd.to_numeric(df[col], errors='coerce') <= float(val))
            elif op == '包含':
                mask &= (df[col].astype(str).str.contains(val, na=False))
            elif op == '不包含':
                mask &= (~df[col].astype(str).str.contains(val, na=False))
            elif op == '为空':
                mask &= (df[col].isnull())
            elif op == '不为空':
                mask &= (df[col].notnull())
        return df[mask]

    @classmethod
    def _format_and_sum(cls, df, sum_columns):
        """格式化数字列，并返回格式化后的DF以及各列合计"""
        df = df.copy()
        sums = {}
        for col in sum_columns:
            if col in df.columns:
                # 转换为数值，无法转换的变为NaN
                df[col] = pd.to_numeric(df[col], errors='coerce')
                # 计算总和
                total = df[col].sum()
                sums[col] = total
            else:
                sums[col] = None
        return df, sums

    @staticmethod
    def _generate_sheet_name(df, col_name, default_index):
        """根据列的第一个非空值生成sheet名，最多31字符（Excel限制）"""
        if col_name and col_name in df.columns:
            # 获取第一个非空值
            first_valid = df[col_name].dropna().iloc[0] if not df[col_name].dropna().empty else None
            if first_valid is not None:
                name = str(first_valid)[:31]  # Excel sheet名最大长度31
                # 去除非法字符
                invalid_chars = r'[]:*?\/\\'
                for ch in invalid_chars:
                    name = name.replace(ch, '_')
                return name if name else f"Sheet{default_index}"
        return f"Sheet{default_index}"