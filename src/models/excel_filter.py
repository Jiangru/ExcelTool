# src/models/excel_filter.py

import pandas as pd
import numpy as np
from pathlib import Path
from src.utils.logger import setup_logger

logger = setup_logger(__name__)


class ExcelFilter:
    """多文件筛选与汇总业务类（支持按列分组导出多个Sheet）"""

    # 运算符映射（保留备用）
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
                          match_config=None,
                          progress_callback=None):
        """
        批量筛选文件，并按指定列的值分组导出到多个Sheet（相同值放在同一个Sheet）
        :param file_paths: list of str, 输入文件路径列表
        :param conditions: list of dict, 筛选条件
        :param sheet_name_col: str, 用于分组的列名，该列的不同值将作为Sheet名称
        :param sum_columns: list of str, 需要格式化为数字并求和的列名（在每个Sheet内分别求和）
        :param output_path: str, 输出文件路径
        :param match_config: dict, 外部匹配配置 {'match_file', 'source_column', 'target_column', 'mode'}
        :param progress_callback: function, 进度回调函数，接收当前进度百分比
        :return: str, 输出文件路径
        """
        def safe_progress(value):
            try:
                if progress_callback is not None:
                    progress_callback(int(value))
            except Exception as e:
                logger.error(f"进度回调异常: {e}, value={value}", exc_info=True)

        # 1. 加载外部匹配集合（如果启用）
        match_set = None
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

        # 2. 遍历所有文件，收集筛选后的数据
        all_filtered_dfs = []
        total_files = len(file_paths)

        for idx, file_path in enumerate(file_paths):
            try:
                # 读取Excel
                df = cls._read_excel(file_path)
                if df.empty:
                    logger.warning(f"文件为空，跳过: {file_path}")
                    continue

                # 应用常规条件筛选
                df_filtered = cls._apply_conditions(df, conditions)

                # 应用外部匹配条件
                if match_set is not None and match_config:
                    source_col = match_config['source_column']
                    mode = match_config['mode']
                    if source_col not in df_filtered.columns:
                        logger.warning(f"原文件缺少匹配列 {source_col}，跳过外部匹配条件")
                    else:
                        # 将匹配列转为字符串并去除首尾空格
                        series = df_filtered[source_col].astype(str).str.strip()
                        if mode == 'keep':
                            mask = series.isin(match_set)
                        else:  # 'remove'
                            mask = ~series.isin(match_set)
                        df_filtered = df_filtered[mask]

                if df_filtered.empty:
                    logger.warning(f"文件筛选后无数据，跳过: {file_path}")
                    continue

                all_filtered_dfs.append(df_filtered)
                logger.info(f"已处理: {file_path} -> 增加 {len(df_filtered)} 行")

            except Exception as e:
                logger.exception(f"处理文件失败 {file_path}: {e}")
                continue

            # 更新进度
            safe_progress(int((idx + 1) / total_files * 100))

        # 3. 合并所有筛选后的数据
        if not all_filtered_dfs:
            logger.warning("所有文件均无符合条件的数据，将创建一个空白工作表")
            final_df = pd.DataFrame()
        else:
            final_df = pd.concat(all_filtered_dfs, ignore_index=True)
            logger.info(f"合并后共 {len(final_df)} 行数据")

        # 4. 输出Excel
        output_dir = Path(output_path).parent
        output_dir.mkdir(parents=True, exist_ok=True)

        # 判断是否启用分组导出
        enable_grouping = (sheet_name_col and sheet_name_col in final_df.columns and not final_df.empty)
        if not enable_grouping:
            # 未启用分组：全部数据写入一个Sheet
            logger.info("未指定分组列或分组列不存在，将全部数据写入单个Sheet")
            df_formatted, total_sums = cls._format_and_sum(final_df, sum_columns)
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df_formatted.to_excel(writer, sheet_name="筛选汇总", index=False)
                if sum_columns and not df_formatted.empty:
                    cls._add_total_row(writer, "筛选汇总", df_formatted, sum_columns, total_sums)
            return output_path

        # 启用分组：按 sheet_name_col 分组导出多个Sheet
        grouped = final_df.groupby(sheet_name_col)
        logger.info(f"按列 '{sheet_name_col}' 分组，共 {len(grouped)} 个组")

        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for group_value, group_df in grouped:
                # 生成合法的Sheet名称
                sheet_name = cls._sanitize_sheet_name(str(group_value))
                # 格式化数字列并计算该组的总和
                group_formatted, group_sums = cls._format_and_sum(group_df, sum_columns)
                # 写入该组数据
                group_formatted.to_excel(writer, sheet_name=sheet_name, index=False)
                # 添加该组的合计行
                if sum_columns and not group_formatted.empty:
                    cls._add_total_row(writer, sheet_name, group_formatted, sum_columns, group_sums)
                logger.info(f"写入Sheet: {sheet_name}，共 {len(group_df)} 行")

        logger.info(f"分组导出完成，输出文件: {output_path}")
        return output_path

    # ----------------------------------------------------------------------
    # 辅助方法
    # ----------------------------------------------------------------------
    @classmethod
    def _load_match_set(cls, match_file, column):
        """读取匹配文件（所有Sheet）指定列，返回去重后的字符串集合"""
        try:
            xl = pd.ExcelFile(match_file)
            all_values = []
            for sheet_name in xl.sheet_names:
                # 假设每个Sheet的第一行为列名
                df = pd.read_excel(match_file, sheet_name=sheet_name, header=0)
                if column in df.columns:
                    values = df[column].dropna().astype(str).str.strip()
                    all_values.append(values)
                    logger.debug(f"从 sheet '{sheet_name}' 读取到 {len(values)} 个值")
                else:
                    logger.warning(f"Sheet '{sheet_name}' 中缺少列 '{column}'，已跳过")
            if not all_values:
                raise ValueError(f"在所有Sheet中均未找到列 '{column}'")
            combined = pd.concat(all_values, ignore_index=True)
            match_set = set(combined)
            logger.info(f"匹配集合加载完成，共 {len(match_set)} 个唯一值（来源于 {len(xl.sheet_names)} 个Sheet）")
            return match_set
        except Exception as e:
            logger.error(f"加载匹配文件失败: {e}")
            raise

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
                df[col] = pd.to_numeric(df[col], errors='coerce')
                sums[col] = df[col].sum()
            else:
                sums[col] = None
        return df, sums

    @staticmethod
    def _sanitize_sheet_name(name):
        """清理Sheet名称中的非法字符，并限制长度31"""
        invalid_chars = r'[]:*?/\\'
        for ch in invalid_chars:
            name = name.replace(ch, '_')
        if len(name) > 31:
            name = name[:31]
        return name if name else "Sheet"

    @classmethod
    def _add_total_row(cls, writer, sheet_name, df, sum_columns, sums):
        """在指定sheet中添加合计行"""
        columns = df.columns.tolist()
        total_row = {col: "" for col in columns}
        if columns:
            total_row[columns[0]] = "合计"
        for col in sum_columns:
            if col in total_row and col in sums:
                total_row[col] = sums[col]
        total_df = pd.DataFrame([total_row])
        startrow = len(df) + 1
        total_df.to_excel(writer, sheet_name=sheet_name,
                          startrow=startrow, index=False, header=False)