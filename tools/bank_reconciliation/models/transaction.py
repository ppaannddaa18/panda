"""交易记录数据模型"""

from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from typing import Optional, Dict, Any, Tuple


@dataclass
class Transaction:
    """交易记录"""
    date: date                              # 交易日期
    amount: Decimal                         # 金额
    summary: str                            # 摘要
    direction: str                         # 'debit'(借/支出) 或 'credit'(贷/收入)
    source: str                            # 'bank' 或 'account'
    balance: Optional[Decimal] = None       # 余额（银行流水）
    voucher_no: Optional[str] = None        # 凭证号（账务系统）
    raw_data: Dict[str, Any] = field(default_factory=dict)  # 原始数据

    def __post_init__(self):
        """初始化后处理"""
        # 确保 amount 是 Decimal 类型
        if not isinstance(self.amount, Decimal):
            self.amount = Decimal(str(self.amount)).quantize(Decimal("0.01"))
        # 确保 balance 是 Decimal 类型
        if self.balance is not None and not isinstance(self.balance, Decimal):
            self.balance = Decimal(str(self.balance)).quantize(Decimal("0.01"))

    def match_key(self) -> Tuple[Decimal, str]:
        """生成匹配键（金额, 方向）"""
        return (self.amount, self.direction)

    def __repr__(self) -> str:
        return f"Transaction({self.date}, {self.direction}: {self.amount}, '{self.summary[:10]}...')"