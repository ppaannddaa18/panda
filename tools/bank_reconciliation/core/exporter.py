"""结果导出器"""

from typing import List
from bank_reconciliation.models import MatchResult, Transaction


class ResultExporter:
    """结果导出器（占位实现）"""

    def __init__(self):
        """初始化导出器"""
        pass

    def export_to_excel(
        self,
        match_results: List[MatchResult],
        unmatched_bank: List[Transaction],
        unmatched_account: List[Transaction],
        output_path: str
    ) -> str:
        """
        导出结果到 Excel

        Args:
            match_results: 匹配结果列表
            unmatched_bank: 未匹配的银行交易
            unmatched_account: 未匹配的账务交易
            output_path: 输出文件路径

        Returns:
            输出文件路径
        """
        # 占位实现
        return output_path
