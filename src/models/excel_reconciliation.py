# src/models/excel_reconciliation.py

import pandas as pd
from pathlib import Path
from src.utils.logger import setup_logger

logger = setup_logger(__name__)


class ExcelReconciliation:
    """绿能畅游对账业务类"""

    @classmethod
    def reconcile(cls,
                  left_file: str,
                  right_file: str,
                  group_col: str,
                  left_energy_col: str,
                  left_fee_col: str,
                  right_energy_col: str,
                  right_fee_col: str,
                  match_config: dict = None,
                  output_path: str = None,
                  progress_callback=None   # 新增进度回调
                  ) -> str:
        """
        执行对账（支持进度回调）
        """
        # 定义内部进度报告函数
        def report(step):
            if progress_callback:
                progress_callback(step)

        report(5)   # 开始

        # 1. 读取文件
        df_left = cls._read_excel(left_file)
        df_right = cls._read_excel(right_file)
        report(15)

        # 校验必要列是否存在
        cls._validate_columns(df_left, [group_col, left_energy_col, left_fee_col], left_file)
        cls._validate_columns(df_right, [group_col, right_energy_col, right_fee_col], right_file)
        report(25)

        # 2. 数值转换，过滤费用为0的订单
        df_left = cls._prepare_dataframe(df_left, group_col, left_energy_col, left_fee_col)
        df_right = cls._prepare_dataframe(df_right, group_col, right_energy_col, right_fee_col)
        report(35)

        # 3. 外部匹配剔除（如果有）
        if match_config:
            df_left = cls._apply_match_filter(df_left, group_col, match_config, '达克云')
            df_right = cls._apply_match_filter(df_right, group_col, match_config, '海汇德')
        report(45)

        # 4. 左侧分组聚合
        left_agg = df_left.groupby(group_col).agg(
            total_energy=(left_energy_col, 'sum'),
            total_fee=(left_fee_col, 'sum'),
            order_count=(group_col, 'count')
        ).reset_index()
        left_agg.columns = [group_col, '总电量', '总费用', '订单数量']
        report(55)

        # 5. 右侧分组聚合
        right_agg = df_right.groupby(group_col).agg(
            total_energy=(right_energy_col, 'sum'),
            total_fee=(right_fee_col, 'sum'),
            order_count=(group_col, 'count')
        ).reset_index()
        right_agg.columns = [group_col, '总电量', '总费用', '订单数量']
        report(65)

        # 6. 合并左右结果（全外连接）
        merged = pd.merge(left_agg, right_agg, on=group_col, how='outer', suffixes=('_left', '_right')).fillna(0)
        report(75)

        # 7. 计算差值
        for col in ['总电量', '总费用', '订单数量']:
            merged[f'{col}_diff'] = merged[f'{col}_left'] - merged[f'{col}_right']
        report(85)

        # 8. 构建最终报表
        result = pd.DataFrame()
        result[group_col] = merged[group_col]
        result['达克云_总电量'] = merged['总电量_left']
        result['达克云_总费用'] = merged['总费用_left']
        result['达克云_订单数量'] = merged['订单数量_left']
        result['海汇德_总电量'] = merged['总电量_right']
        result['海汇德_总费用'] = merged['总费用_right']
        result['海汇德_订单数量'] = merged['订单数量_right']
        result['差值_总电量'] = merged['总电量_diff']
        result['差值_总费用'] = merged['总费用_diff']
        result['差值_订单数量'] = merged['订单数量_diff']
        report(95)

        # 9. 保存
        if output_path is None:
            output_dir = Path(left_file).parent
            output_path = str(output_dir / f"对账结果_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        else:
            output_dir = Path(output_path).parent
            output_dir.mkdir(parents=True, exist_ok=True)

        result.to_excel(output_path, index=False)
        report(100)
        logger.info(f"对账完成，结果已保存至: {output_path}")
        return output_path

    @classmethod
    def _prepare_dataframe(cls, df, group_col, energy_col, fee_col):
        df = df.copy()
        df[energy_col] = pd.to_numeric(df[energy_col], errors='coerce')
        df[fee_col] = pd.to_numeric(df[fee_col], errors='coerce')
        df.dropna(subset=[group_col, energy_col, fee_col], inplace=True)
        df = df[df[fee_col] != 0]
        return df

    @classmethod
    def _apply_match_filter(cls, df, group_col, match_config, side):
        if match_config is None:
            return df
        match_file = match_config.get('match_file')
        if not match_file:
            return df
        match_col = match_config.get('match_col', group_col)
        try:
            df_match = cls._read_excel(match_file)
            if match_col not in df_match.columns:
                logger.error(f"外部文件缺少列 {match_col}")
                return df
            match_values = set(df_match[match_col].dropna().astype(str).str.strip())
            if not match_values:
                logger.warning("外部匹配集合为空")
                return df
            series = df[group_col].astype(str).str.strip()
            mask = ~series.isin(match_values)   # 剔除匹配上的
            filtered = df[mask]
            logger.info(f"{side} 外部匹配剔除后剩余 {len(filtered)} 行（原 {len(df)} 行）")
            return filtered
        except Exception as e:
            logger.error(f"外部匹配处理失败: {e}")
            return df

    @staticmethod
    def _read_excel(file_path):
        if str(file_path).endswith('.xls'):
            return pd.read_excel(file_path, engine='xlrd')
        else:
            return pd.read_excel(file_path, engine='openpyxl')

    @staticmethod
    def _validate_columns(df, columns, file_name):
        missing = [col for col in columns if col not in df.columns]
        if missing:
            raise ValueError(f"文件 {file_name} 中缺少必要的列: {missing}")