"""源数据模型"""

from dataclasses import dataclass, field
from decimal import Decimal
from typing import List, Dict, Any, Optional


@dataclass
class SourceRecord:
    """单条源数据记录"""
    company: str                      # 公司名称
    month: int                        # 月份 (1-12)
    account_name: str                 # 原始科目名称
    standardized_name: str = ""       # 标准化科目名称
    amount: Decimal = Decimal("0.00") # 金额
    source_file: str = ""             # 来源文件路径
    sheet_name: str = ""              # 来源 sheet 名称

    def __post_init__(self):
        """初始化后处理"""
        if not isinstance(self.amount, Decimal):
            self.amount = Decimal(str(self.amount)).quantize(Decimal("0.01"))


@dataclass
class SheetData:
    """单个 Sheet 的数据"""
    sheet_name: str
    records: List[SourceRecord] = field(default_factory=list)
    months_found: List[int] = field(default_factory=list)
    accounts_found: List[str] = field(default_factory=list)


@dataclass
class ParsedFile:
    """解析后的文件数据"""
    file_path: str
    company: str = ""                         # 从文件名识别
    sheets: List[SheetData] = field(default_factory=list)
    parse_errors: List[str] = field(default_factory=list)

    @property
    def all_records(self) -> List[SourceRecord]:
        """获取所有记录"""
        records = []
        for sheet in self.sheets:
            records.extend(sheet.records)
        return records

    @property
    def all_accounts(self) -> List[str]:
        """获取所有科目名称"""
        accounts = set()
        for sheet in self.sheets:
            accounts.update(sheet.accounts_found)
        return list(accounts)