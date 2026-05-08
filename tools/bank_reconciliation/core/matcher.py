"""匹配引擎"""

from typing import List, Tuple
from bank_reconciliation.models import Transaction, MatchResult


class MatchEngine:
    """匹配引擎（占位实现）"""

    def __init__(self):
        """初始化匹配引擎"""
        pass

    def match(
        self,
        bank_transactions: List[Transaction],
        account_transactions: List[Transaction]
    ) -> Tuple[List[MatchResult], List[Transaction], List[Transaction]]:
        """
        执行匹配

        Args:
            bank_transactions: 银行流水交易列表
            account_transactions: 账务明细交易列表

        Returns:
            (匹配结果列表, 未匹配银行交易列表, 未匹配账务交易列表)
        """
        # 占位实现
        return [], bank_transactions, account_transactions
