"""合并结果模型"""

from dataclasses import dataclass, field
from typing import List, Set, Dict, Any
import pandas as pd

from .source_data import SourceRecord


@dataclass
class MergeResult:
    """合并结果"""
    records: List[SourceRecord] = field(default_factory=list)
    unmatched_accounts: Set[str] = field(default_factory=set)
    pivot_table: pd.DataFrame = field(default_factory=pd.DataFrame)
    statistics: Dict[str, Any] = field(default_factory=dict)

    @property
    def total_amount(self) -> float:
        """总金额"""
        from decimal import Decimal
        return float(sum(r.amount for r in self.records))

    @property
    def company_count(self) -> int:
        """公司数量"""
        return len(set(r.company for r in self.records))

    @property
    def record_count(self) -> int:
        """记录数量"""
        return len(self.records)