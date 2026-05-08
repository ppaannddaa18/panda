"""匹配结果数据模型"""

from dataclasses import dataclass, field
from typing import Optional, List
from enum import Enum

from .transaction import Transaction


class MatchStatus(Enum):
    """匹配状态"""
    MATCHED = "已达账项"
    BANK_UNMATCHED = "银行未达"      # 银行有，账务无
    ACCOUNT_UNMATCHED = "企业未达"   # 账务有，银行无
    MANUAL_MATCHED = "手动匹配"
    SPLIT_MATCHED = "拆分匹配"


@dataclass
class SplitPart:
    """拆分匹配的子项"""
    bank_txn: Transaction
    amount: float


@dataclass
class MatchResult:
    """匹配结果"""
    bank_txn: Optional[Transaction] = None      # 银行流水记录
    account_txn: Optional[Transaction] = None   # 账务记录
    status: MatchStatus = MatchStatus.MATCHED   # 匹配状态
    match_score: float = 0.0                    # 匹配置信度 (0-1)
    match_reason: str = ""                      # 匹配原因说明
    split_parts: Optional[List[SplitPart]] = None  # 拆分匹配的子项

    @property
    def is_matched(self) -> bool:
        """是否已匹配"""
        return self.status in (
            MatchStatus.MATCHED,
            MatchStatus.MANUAL_MATCHED,
            MatchStatus.SPLIT_MATCHED
        )

    @property
    def status_display(self) -> str:
        """状态显示文本"""
        return self.status.value