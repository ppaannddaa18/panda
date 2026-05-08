# 银行对账助手实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 构建智能银行对账助手，自动匹配银行流水与企业账务数据，输出对账结果和余额调节表。

**Architecture:** 模块化设计，核心层（匹配引擎、解析器、导出器）与 UI 层分离，数据模型独立。遵循项目现有 tkinter + ttkbootstrap + tkinterdnd2 技术栈。

**Tech Stack:** Python 3.10+, pandas, openpyxl, tkinter, ttkbootstrap, tkinterdnd2, difflib

---

## 文件结构

```
tools/bank_reconciliation/
├── __init__.py              # 模块入口，导出主函数
├── main.py                  # GUI 主程序入口
├── core/
│   ├── __init__.py
│   ├── matcher.py           # 核心匹配引擎
│   ├── parser.py            # Excel 解析器
│   └── exporter.py          # 结果导出器
├── ui/
│   ├── __init__.py
│   ├── main_window.py       # 主窗口
│   ├── file_loader.py       # 文件加载组件
│   ├── column_mapper.py     # 列映射配置对话框
│   ├── match_table.py       # 匹配结果表格
│   └── manual_adjust.py     # 手动调整面板
├── models/
│   ├── __init__.py
│   ├── transaction.py       # 交易记录数据模型
│   └── match_result.py      # 匹配结果数据模型
├── templates/
│   ├── bank_templates.json  # 银行流水模板配置
│   └── account_templates.json # 账务系统模板配置
└── config.py                # 配置管理
```

---

## Task 1: 项目骨架和数据模型

**Files:**
- Create: `tools/bank_reconciliation/__init__.py`
- Create: `tools/bank_reconciliation/config.py`
- Create: `tools/bank_reconciliation/models/__init__.py`
- Create: `tools/bank_reconciliation/models/transaction.py`
- Create: `tools/bank_reconciliation/models/match_result.py`

- [ ] **Step 1: 创建模块目录结构**

```bash
mkdir -p tools/bank_reconciliation/core
mkdir -p tools/bank_reconciliation/ui
mkdir -p tools/bank_reconciliation/models
mkdir -p tools/bank_reconciliation/templates
```

- [ ] **Step 2: 创建模块入口 `__init__.py`**

```python
"""
银行对账助手

自动匹配银行流水与企业账务数据，输出对账结果和余额调节表。
"""

__version__ = "1.0.0"

from .main import main

__all__ = ["main"]
```

- [ ] **Step 3: 创建配置管理 `config.py`**

```python
"""配置管理"""

from dataclasses import dataclass
from typing import Dict, Any
import json
from pathlib import Path

@dataclass
class AppConfig:
    """应用配置"""
    # 匹配参数
    date_tolerance_days: int = 1  # 日期容忍天数
    summary_similarity_threshold: float = 0.6  # 摘要相似度阈值

    # 界面配置
    window_width: int = 1200
    window_height: int = 800

    # 输出配置
    output_dir: str = "output"

    @classmethod
    def load(cls, config_path: str = None) -> "AppConfig":
        """加载配置"""
        if config_path and Path(config_path).exists():
            with open(config_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                return cls(**data)
        return cls()

# 预定义银行模板
BANK_TEMPLATES: Dict[str, Dict[str, str]] = {
    "工商银行": {
        "date": "交易日期",
        "debit": "支出",
        "credit": "收入",
        "balance": "余额",
        "summary": "摘要"
    },
    "建设银行": {
        "date": "交易时间",
        "debit": "借方",
        "credit": "贷方",
        "balance": "余额",
        "summary": "交易摘要"
    },
}

# 预定义账务系统模板
ACCOUNT_TEMPLATES: Dict[str, Dict[str, str]] = {
    "用友": {
        "date": "日期",
        "debit": "借方",
        "credit": "贷方",
        "voucher": "凭证号",
        "summary": "摘要"
    },
    "金蝶": {
        "date": "日期",
        "debit": "借方金额",
        "credit": "贷方金额",
        "voucher": "凭证字号",
        "summary": "摘要"
    },
}
```

- [ ] **Step 4: 创建交易记录模型 `models/transaction.py`**

```python
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
    direction: str                          # 'debit'(借/支出) 或 'credit'(贷/收入)
    source: str                             # 'bank' 或 'account'
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
```

- [ ] **Step 5: 创建匹配结果模型 `models/match_result.py`**

```python
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
```

- [ ] **Step 6: 创建 models 包入口 `models/__init__.py`**

```python
"""数据模型"""

from .transaction import Transaction
from .match_result import MatchStatus, MatchResult, SplitPart

__all__ = ["Transaction", "MatchStatus", "MatchResult", "SplitPart"]
```

- [ ] **Step 7: 提交**

```bash
git add tools/bank_reconciliation/
git commit -m "feat(bank-reconciliation): 添加项目骨架和数据模型

- 创建模块目录结构
- 添加配置管理
- 实现 Transaction 交易记录模型
- 实现 MatchResult 匹配结果模型

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>"
```

---

## Task 2: Excel 解析器

**Files:**
- Create: `tools/bank_reconciliation/core/__init__.py`
- Create: `tools/bank_reconciliation/core/parser.py`

- [ ] **Step 1: 创建 core 包入口 `core/__init__.py`**

```python
"""核心模块"""

from .parser import BankStatementParser, AccountLedgerParser
from .matcher import MatchEngine
from .exporter import ResultExporter

__all__ = [
    "BankStatementParser",
    "AccountLedgerParser",
    "MatchEngine",
    "ResultExporter"
]
```

- [ ] **Step 2: 创建 Excel 解析器 `core/parser.py`**

```python
"""Excel 解析器"""

import pandas as pd
from datetime import date
from decimal import Decimal
from typing import List, Dict, Optional, Any

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))
from bank_reconciliation.models import Transaction
from bank_reconciliation.config import BANK_TEMPLATES, ACCOUNT_TEMPLATES


class BaseParser:
    """解析器基类"""

    def __init__(self, column_mapping: Optional[Dict[str, str]] = None):
        """
        初始化解析器

        Args:
            column_mapping: 自定义列映射，格式如 {"date": "交易日期", "debit": "支出", ...}
        """
        self.column_mapping = column_mapping or {}

    def _clean_amount(self, value: Any) -> Decimal:
        """清洗金额数据"""
        if pd.isna(value) or value == "" or value == "-":
            return Decimal("0.00")
        if isinstance(value, (int, float)):
            return Decimal(str(round(value, 2)))
        # 处理字符串格式金额（去除逗号等）
        if isinstance(value, str):
            value = value.replace(",", "").replace("，", "").strip()
            if value == "" or value == "-":
                return Decimal("0.00")
            return Decimal(value).quantize(Decimal("0.01"))
        return Decimal("0.00")

    def _clean_date(self, value: Any) -> Optional[date]:
        """清洗日期数据"""
        if pd.isna(value):
            return None
        try:
            # 尝试解析各种日期格式
            if isinstance(value, date):
                return value
            if isinstance(value, str):
                # 尝试常见格式
                for fmt in ["%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日"]:
                    try:
                        return pd.to_datetime(value, format=fmt).date()
                    except ValueError:
                        continue
                # pandas 自动解析
                return pd.to_datetime(value).date()
            # Excel 日期序列号
            if isinstance(value, (int, float)):
                return pd.to_datetime("1899-12-30") + pd.Timedelta(days=int(value))
        except Exception:
            return None
        return None

    def _clean_summary(self, value: Any) -> str:
        """清洗摘要数据"""
        if pd.isna(value):
            return ""
        return str(value).strip()


class BankStatementParser(BaseParser):
    """银行流水解析器"""

    def __init__(
        self,
        column_mapping: Optional[Dict[str, str]] = None,
        template_name: Optional[str] = None
    ):
        """
        初始化银行流水解析器

        Args:
            column_mapping: 自定义列映射
            template_name: 预定义模板名称（如 "工商银行"、"建设银行"）
        """
        if template_name and template_name in BANK_TEMPLATES:
            column_mapping = BANK_TEMPLATES[template_name]
        super().__init__(column_mapping)

    def parse(self, file_path: str) -> List[Transaction]:
        """
        解析银行流水 Excel 文件

        Args:
            file_path: Excel 文件路径

        Returns:
            交易记录列表
        """
        df = pd.read_excel(file_path)
        return self._parse_dataframe(df)

    def _parse_dataframe(self, df: pd.DataFrame) -> List[Transaction]:
        """解析 DataFrame"""
        transactions = []

        # 获取列名
        date_col = self.column_mapping.get("date", "交易日期")
        debit_col = self.column_mapping.get("debit", "支出")
        credit_col = self.column_mapping.get("credit", "收入")
        balance_col = self.column_mapping.get("balance", "余额")
        summary_col = self.column_mapping.get("summary", "摘要")

        for idx, row in df.iterrows():
            try:
                txn_date = self._clean_date(row.get(date_col))
                if txn_date is None:
                    continue

                debit = self._clean_amount(row.get(debit_col, 0))
                credit = self._clean_amount(row.get(credit_col, 0))
                balance = self._clean_amount(row.get(balance_col)) if balance_col in row else None
                summary = self._clean_summary(row.get(summary_col, ""))

                # 创建借方交易（支出）
                if debit > 0:
                    txn = Transaction(
                        date=txn_date,
                        amount=debit,
                        summary=summary,
                        direction="debit",
                        source="bank",
                        balance=balance,
                        raw_data=row.to_dict()
                    )
                    transactions.append(txn)

                # 创建贷方交易（收入）
                if credit > 0:
                    txn = Transaction(
                        date=txn_date,
                        amount=credit,
                        summary=summary,
                        direction="credit",
                        source="bank",
                        balance=balance,
                        raw_data=row.to_dict()
                    )
                    transactions.append(txn)

            except Exception as e:
                # 跳过解析错误的行
                print(f"解析第 {idx + 1} 行时出错: {e}")
                continue

        return transactions


class AccountLedgerParser(BaseParser):
    """账务明细解析器"""

    def __init__(
        self,
        column_mapping: Optional[Dict[str, str]] = None,
        template_name: Optional[str] = None
    ):
        """
        初始化账务明细解析器

        Args:
            column_mapping: 自定义列映射
            template_name: 预定义模板名称（如 "用友"、"金蝶"）
        """
        if template_name and template_name in ACCOUNT_TEMPLATES:
            column_mapping = ACCOUNT_TEMPLATES[template_name]
        super().__init__(column_mapping)

    def parse(self, file_path: str) -> List[Transaction]:
        """
        解析账务明细 Excel 文件

        Args:
            file_path: Excel 文件路径

        Returns:
            交易记录列表
        """
        df = pd.read_excel(file_path)
        return self._parse_dataframe(df)

    def _parse_dataframe(self, df: pd.DataFrame) -> List[Transaction]:
        """解析 DataFrame"""
        transactions = []

        # 获取列名
        date_col = self.column_mapping.get("date", "日期")
        debit_col = self.column_mapping.get("debit", "借方")
        credit_col = self.column_mapping.get("credit", "贷方")
        voucher_col = self.column_mapping.get("voucher", "凭证号")
        summary_col = self.column_mapping.get("summary", "摘要")

        for idx, row in df.iterrows():
            try:
                txn_date = self._clean_date(row.get(date_col))
                if txn_date is None:
                    continue

                debit = self._clean_amount(row.get(debit_col, 0))
                credit = self._clean_amount(row.get(credit_col, 0))
                voucher_no = str(row.get(voucher_col, "")).strip() if voucher_col in row else None
                summary = self._clean_summary(row.get(summary_col, ""))

                # 创建借方交易
                if debit > 0:
                    txn = Transaction(
                        date=txn_date,
                        amount=debit,
                        summary=summary,
                        direction="debit",
                        source="account",
                        voucher_no=voucher_no,
                        raw_data=row.to_dict()
                    )
                    transactions.append(txn)

                # 创建贷方交易
                if credit > 0:
                    txn = Transaction(
                        date=txn_date,
                        amount=credit,
                        summary=summary,
                        direction="credit",
                        source="account",
                        voucher_no=voucher_no,
                        raw_data=row.to_dict()
                    )
                    transactions.append(txn)

            except Exception as e:
                print(f"解析第 {idx + 1} 行时出错: {e}")
                continue

        return transactions
```

- [ ] **Step 3: 提交**

```bash
git add tools/bank_reconciliation/core/
git commit -m "feat(bank-reconciliation): 实现 Excel 解析器

- 添加 BankStatementParser 银行流水解析器
- 添加 AccountLedgerParser 账务明细解析器
- 支持预定义模板和自定义列映射
- 实现日期、金额、摘要数据清洗

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>"
```

---

## Task 3: 核心匹配引擎

**Files:**
- Create: `tools/bank_reconciliation/core/matcher.py`

- [ ] **Step 1: 创建匹配引擎 `core/matcher.py`**

```python
"""核心匹配引擎"""

from difflib import SequenceMatcher
from datetime import timedelta
from decimal import Decimal
from typing import List, Dict, Tuple, Optional
from collections import defaultdict

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))
from bank_reconciliation.models import Transaction, MatchResult, MatchStatus, SplitPart
from bank_reconciliation.config import AppConfig


class MatchEngine:
    """核心匹配引擎"""

    def __init__(self, config: Optional[AppConfig] = None):
        """
        初始化匹配引擎

        Args:
            config: 应用配置
        """
        self.config = config or AppConfig()
        self.date_tolerance = timedelta(days=self.config.date_tolerance_days)
        self.similarity_threshold = self.config.summary_similarity_threshold

    def match(
        self,
        bank_txns: List[Transaction],
        account_txns: List[Transaction]
    ) -> List[MatchResult]:
        """
        执行匹配

        Args:
            bank_txns: 银行流水交易列表
            account_txns: 账务明细交易列表

        Returns:
            匹配结果列表
        """
        results: List[MatchResult] = []

        # 按金额和方向分组
        bank_groups = self._group_by_key(bank_txns)
        account_groups = self._group_by_key(account_txns)

        # 记录已匹配的交易
        matched_bank: set = set()
        matched_account: set = set()

        # 第一阶段：金额精确匹配
        for key, bank_list in bank_groups.items():
            if key not in account_groups:
                continue

            account_list = account_groups[key]

            for bank_txn in bank_list:
                best_match = self._find_best_match(bank_txn, account_list, matched_account)

                if best_match:
                    account_txn, score, reason = best_match
                    matched_bank.add(id(bank_txn))
                    matched_account.add(id(account_txn))

                    results.append(MatchResult(
                        bank_txn=bank_txn,
                        account_txn=account_txn,
                        status=MatchStatus.MATCHED,
                        match_score=score,
                        match_reason=reason
                    ))

        # 第二阶段：处理拆分匹配（一条账务匹配多条银行）
        split_results = self._handle_split_matches(
            bank_txns, account_txns, matched_bank, matched_account
        )
        results.extend(split_results)

        # 第三阶段：标记未匹配项
        for txn in bank_txns:
            if id(txn) not in matched_bank:
                results.append(MatchResult(
                    bank_txn=txn,
                    account_txn=None,
                    status=MatchStatus.BANK_UNMATCHED,
                    match_reason="银行有记录，账务无对应记录"
                ))

        for txn in account_txns:
            if id(txn) not in matched_account:
                results.append(MatchResult(
                    bank_txn=None,
                    account_txn=txn,
                    status=MatchStatus.ACCOUNT_UNMATCHED,
                    match_reason="账务有记录，银行无对应记录"
                ))

        return results

    def _group_by_key(self, txns: List[Transaction]) -> Dict[Tuple[Decimal, str], List[Transaction]]:
        """按匹配键（金额, 方向）分组"""
        groups: Dict[Tuple[Decimal, str], List[Transaction]] = defaultdict(list)
        for txn in txns:
            groups[txn.match_key()].append(txn)
        return groups

    def _find_best_match(
        self,
        bank_txn: Transaction,
        account_list: List[Transaction],
        matched: set
    ) -> Optional[Tuple[Transaction, float, str]]:
        """
        在账务列表中找到最佳匹配

        Returns:
            (匹配的交易, 匹配分数, 匹配原因) 或 None
        """
        best_match = None
        best_score = 0.0
        best_reason = ""

        for account_txn in account_list:
            if id(account_txn) in matched:
                continue

            score, reason = self._calculate_match_score(bank_txn, account_txn)

            if score > best_score:
                best_score = score
                best_match = account_txn
                best_reason = reason

        if best_match and best_score > 0:
            return (best_match, best_score, best_reason)
        return None

    def _calculate_match_score(
        self,
        bank_txn: Transaction,
        account_txn: Transaction
    ) -> Tuple[float, str]:
        """
        计算匹配分数

        Returns:
            (分数, 原因)
        """
        reasons = []
        score = 0.0

        # 金额已匹配（分组时已确保）
        reasons.append("金额一致")

        # 日期匹配
        date_diff = abs((bank_txn.date - account_txn.date).days)
        if date_diff == 0:
            score += 0.5
            reasons.append("日期相同")
        elif date_diff <= self.date_tolerance.days:
            score += 0.3
            reasons.append(f"日期相差{date_diff}天")

        # 摘要相似度
        similarity = self._calculate_similarity(bank_txn.summary, account_txn.summary)
        if similarity >= self.similarity_threshold:
            score += similarity * 0.2
            reasons.append(f"摘要相似度{similarity:.0%}")

        reason = "；".join(reasons)
        return (score, reason)

    def _calculate_similarity(self, text1: str, text2: str) -> float:
        """计算文本相似度"""
        if not text1 or not text2:
            return 0.0
        return SequenceMatcher(None, text1, text2).ratio()

    def _handle_split_matches(
        self,
        bank_txns: List[Transaction],
        account_txns: List[Transaction],
        matched_bank: set,
        matched_account: set
    ) -> List[MatchResult]:
        """
        处理拆分匹配

        场景：一条账务记录拆分成多条银行流水
        """
        results = []

        # 找出未匹配的账务记录
        unmatched_account = [
            txn for txn in account_txns
            if id(txn) not in matched_account
        ]

        # 找出未匹配的银行记录
        unmatched_bank = [
            txn for txn in bank_txns
            if id(txn) not in matched_bank
        ]

        # 按方向分组
        for direction in ["debit", "credit"]:
            account_by_dir = [t for t in unmatched_account if t.direction == direction]
            bank_by_dir = [t for t in unmatched_bank if t.direction == direction]

            for account_txn in account_by_dir:
                # 尝试找到多条银行记录，金额之和等于账务金额
                split_parts = self._find_split_combination(
                    account_txn.amount,
                    bank_by_dir,
                    matched_bank
                )

                if split_parts:
                    matched_account.add(id(account_txn))
                    for part in split_parts:
                        matched_bank.add(id(part.bank_txn))

                    results.append(MatchResult(
                        bank_txn=split_parts[0].bank_txn,  # 主记录
                        account_txn=account_txn,
                        status=MatchStatus.SPLIT_MATCHED,
                        match_score=0.8,
                        match_reason=f"拆分匹配：{len(split_parts)}笔银行流水",
                        split_parts=split_parts
                    ))

        return results

    def _find_split_combination(
        self,
        target_amount: Decimal,
        candidates: List[Transaction],
        matched: set
    ) -> Optional[List[SplitPart]]:
        """
        找到金额之和等于目标金额的组合

        简化实现：尝试找 2-3 条记录的组合
        """
        available = [t for t in candidates if id(t) not in matched]

        # 尝试两两组合
        for i, txn1 in enumerate(available):
            for txn2 in available[i+1:]:
                if txn1.amount + txn2.amount == target_amount:
                    return [
                        SplitPart(bank_txn=txn1, amount=float(txn1.amount)),
                        SplitPart(bank_txn=txn2, amount=float(txn2.amount))
                    ]

        # 尝试三三组合
        for i, txn1 in enumerate(available):
            for j, txn2 in enumerate(available[i+1:], i+1):
                for txn3 in available[j+1:]:
                    if txn1.amount + txn2.amount + txn3.amount == target_amount:
                        return [
                            SplitPart(bank_txn=txn1, amount=float(txn1.amount)),
                            SplitPart(bank_txn=txn2, amount=float(txn2.amount)),
                            SplitPart(bank_txn=txn3, amount=float(txn3.amount))
                        ]

        return None

    def manual_match(
        self,
        results: List[MatchResult],
        bank_txn: Transaction,
        account_txn: Transaction
    ) -> List[MatchResult]:
        """
        手动建立匹配

        Args:
            results: 当前匹配结果列表
            bank_txn: 银行流水记录
            account_txn: 账务记录

        Returns:
            更新后的匹配结果列表
        """
        # 移除原有的未匹配记录
        results = [
            r for r in results
            if r.bank_txn != bank_txn and r.account_txn != account_txn
        ]

        # 添加手动匹配记录
        results.append(MatchResult(
            bank_txn=bank_txn,
            account_txn=account_txn,
            status=MatchStatus.MANUAL_MATCHED,
            match_score=1.0,
            match_reason="手动匹配"
        ))

        return results

    def cancel_match(
        self,
        results: List[MatchResult],
        match_result: MatchResult
    ) -> List[MatchResult]:
        """
        取消匹配

        Args:
            results: 当前匹配结果列表
            match_result: 要取消的匹配结果

        Returns:
            更新后的匹配结果列表
        """
        # 移除原匹配
        results = [r for r in results if r != match_result]

        # 添加未匹配记录
        if match_result.bank_txn:
            results.append(MatchResult(
                bank_txn=match_result.bank_txn,
                account_txn=None,
                status=MatchStatus.BANK_UNMATCHED,
                match_reason="匹配已取消"
            ))

        if match_result.account_txn:
            results.append(MatchResult(
                bank_txn=None,
                account_txn=match_result.account_txn,
                status=MatchStatus.ACCOUNT_UNMATCHED,
                match_reason="匹配已取消"
            ))

        return results
```

- [ ] **Step 2: 提交**

```bash
git add tools/bank_reconciliation/core/matcher.py
git commit -m "feat(bank-reconciliation): 实现核心匹配引擎

- 多阶段匹配策略（金额精确+日期模糊+摘要相似度）
- 支持拆分匹配（一条账务匹配多条银行流水）
- 支持手动匹配和取消匹配

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>"
```

---

## Task 4: 结果导出器

**Files:**
- Create: `tools/bank_reconciliation/core/exporter.py`

- [ ] **Step 1: 创建结果导出器 `core/exporter.py`**

```python
"""结果导出器"""

import os
from datetime import datetime
from typing import List, Dict, Any
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

import sys
from pathlib import Path as SysPath
sys.path.insert(0, str(SysPath(__file__).parent.parent.parent))
from bank_reconciliation.models import MatchResult, MatchStatus


class ResultExporter:
    """结果导出器"""

    # 样式定义
    HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    HEADER_FONT = Font(bold=True, color="FFFFFF")
    MATCHED_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    UNMATCHED_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    BORDER = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    def __init__(self, output_dir: str = "output"):
        """
        初始化导出器

        Args:
            output_dir: 输出目录
        """
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def export_all(self, results: List[MatchResult]) -> Dict[str, str]:
        """
        导出所有结果

        Returns:
            输出文件路径字典
        """
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        return {
            "match_result": self.export_match_result(results, timestamp),
            "balance_sheet": self.export_balance_sheet(results, timestamp),
            "unmatched_details": self.export_unmatched_details(results, timestamp),
            "statistics": self.export_statistics(results, timestamp)
        }

    def export_match_result(self, results: List[MatchResult], timestamp: str = None) -> str:
        """导出对账结果 Excel"""
        if timestamp is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        file_path = self.output_dir / f"对账结果_{timestamp}.xlsx"

        # 构建数据
        data = []
        for result in results:
            row = {
                "匹配状态": result.status_display,
                "银行日期": result.bank_txn.date.strftime("%Y-%m-%d") if result.bank_txn else "-",
                "银行摘要": result.bank_txn.summary if result.bank_txn else "-",
                "银行借方": float(result.bank_txn.amount) if result.bank_txn and result.bank_txn.direction == "debit" else "",
                "银行贷方": float(result.bank_txn.amount) if result.bank_txn and result.bank_txn.direction == "credit" else "",
                "账务日期": result.account_txn.date.strftime("%Y-%m-%d") if result.account_txn else "-",
                "账务摘要": result.account_txn.summary if result.account_txn else "-",
                "账务借方": float(result.account_txn.amount) if result.account_txn and result.account_txn.direction == "debit" else "",
                "账务贷方": float(result.account_txn.amount) if result.account_txn and result.account_txn.direction == "credit" else "",
                "凭证号": result.account_txn.voucher_no if result.account_txn else "-",
                "匹配说明": result.match_reason
            }
            data.append(row)

        df = pd.DataFrame(data)

        # 创建工作簿
        wb = Workbook()
        ws = wb.active
        ws.title = "对账结果"

        # 写入数据
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=value)

                # 应用边框
                cell.border = self.BORDER

                # 表头样式
                if r_idx == 1:
                    cell.fill = self.HEADER_FILL
                    cell.font = self.HEADER_FONT
                    cell.alignment = Alignment(horizontal="center")
                else:
                    # 条件格式
                    status = data[r_idx - 2]["匹配状态"]
                    if status in ["已达账项", "手动匹配", "拆分匹配"]:
                        cell.fill = self.MATCHED_FILL
                    else:
                        cell.fill = self.UNMATCHED_FILL

        # 调整列宽
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            ws.column_dimensions[column_letter].width = min(max_length + 2, 30)

        wb.save(file_path)
        return str(file_path)

    def export_balance_sheet(self, results: List[MatchResult], timestamp: str = None) -> str:
        """导出余额调节表"""
        if timestamp is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        file_path = self.output_dir / f"余额调节表_{timestamp}.xlsx"

        wb = Workbook()
        ws = wb.active
        ws.title = "余额调节表"

        # 标题
        ws.merge_cells("A1:F1")
        ws["A1"] = "银行存款余额调节表"
        ws["A1"].font = Font(bold=True, size=14)
        ws["A1"].alignment = Alignment(horizontal="center")

        # 分类统计
        categories = {
            "企业已收银行未收": [],
            "企业已付银行未付": [],
            "银行已收企业未收": [],
            "银行已付企业未付": []
        }

        for result in results:
            if result.status == MatchStatus.ACCOUNT_UNMATCHED:
                if result.account_txn.direction == "credit":
                    categories["企业已收银行未收"].append(result.account_txn)
                else:
                    categories["企业已付银行未付"].append(result.account_txn)
            elif result.status == MatchStatus.BANK_UNMATCHED:
                if result.bank_txn.direction == "credit":
                    categories["银行已收企业未收"].append(result.bank_txn)
                else:
                    categories["银行已付企业未付"].append(result.bank_txn)

        row = 3
        for category, txns in categories.items():
            ws.cell(row=row, column=1, value=category).font = Font(bold=True)
            row += 1

            for txn in txns:
                ws.cell(row=row, column=1, value=txn.date.strftime("%Y-%m-%d"))
                ws.cell(row=row, column=2, value=txn.summary)
                ws.cell(row=row, column=3, value=float(txn.amount))
                row += 1

            # 小计
            total = sum(float(t.amount) for t in txns)
            ws.cell(row=row, column=1, value="小计").font = Font(bold=True)
            ws.cell(row=row, column=3, value=total).font = Font(bold=True)
            row += 2

        # 应用边框
        for row_cells in ws.iter_rows(min_row=3, max_row=row, min_col=1, max_col=3):
            for cell in row_cells:
                cell.border = self.BORDER

        wb.save(file_path)
        return str(file_path)

    def export_unmatched_details(self, results: List[MatchResult], timestamp: str = None) -> str:
        """导出未达账项明细"""
        if timestamp is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        file_path = self.output_dir / f"未达账项明细_{timestamp}.xlsx"

        unmatched = [r for r in results if not r.is_matched]

        data = []
        for result in unmatched:
            if result.bank_txn:
                data.append({
                    "类型": "银行未达",
                    "日期": result.bank_txn.date.strftime("%Y-%m-%d"),
                    "摘要": result.bank_txn.summary,
                    "金额": float(result.bank_txn.amount),
                    "方向": "借方" if result.bank_txn.direction == "debit" else "贷方"
                })
            if result.account_txn:
                data.append({
                    "类型": "企业未达",
                    "日期": result.account_txn.date.strftime("%Y-%m-%d"),
                    "摘要": result.account_txn.summary,
                    "金额": float(result.account_txn.amount),
                    "方向": "借方" if result.account_txn.direction == "debit" else "贷方"
                })

        df = pd.DataFrame(data)
        df.to_excel(file_path, index=False)

        return str(file_path)

    def export_statistics(self, results: List[MatchResult], timestamp: str = None) -> str:
        """导出对账统计报告"""
        if timestamp is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        file_path = self.output_dir / f"对账统计报告_{timestamp}.xlsx"

        # 统计数据
        total = len(results)
        matched = sum(1 for r in results if r.is_matched)
        bank_unmatched = sum(1 for r in results if r.status == MatchStatus.BANK_UNMATCHED)
        account_unmatched = sum(1 for r in results if r.status == MatchStatus.ACCOUNT_UNMATCHED)

        data = {
            "指标": [
                "总记录数",
                "已匹配数",
                "银行未达数",
                "企业未达数",
                "匹配率"
            ],
            "数值": [
                total,
                matched,
                bank_unmatched,
                account_unmatched,
                f"{matched / total * 100:.1f}%" if total > 0 else "0%"
            ]
        }

        df = pd.DataFrame(data)
        df.to_excel(file_path, index=False)

        return str(file_path)
```

- [ ] **Step 2: 提交**

```bash
git add tools/bank_reconciliation/core/exporter.py
git commit -m "feat(bank-reconciliation): 实现结果导出器

- 导出对账结果 Excel（含匹配状态列）
- 导出余额调节表（按未达账项分类）
- 导出未达账项明细
- 导出对账统计报告
- 自动添加边框和条件格式

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>"
```

---

## Task 5: GUI 主窗口

**Files:**
- Create: `tools/bank_reconciliation/ui/__init__.py`
- Create: `tools/bank_reconciliation/ui/main_window.py`
- Create: `tools/bank_reconciliation/ui/file_loader.py`

- [ ] **Step 1: 创建 UI 包入口 `ui/__init__.py`**

```python
"""UI 组件"""

from .main_window import BankReconciliationApp
from .file_loader import FileLoader

__all__ = ["BankReconciliationApp", "FileLoader"]
```

- [ ] **Step 2: 创建文件加载组件 `ui/file_loader.py`**

```python
"""文件加载组件"""

import tkinter as tk
from tkinter import ttk, filedialog
from typing import Optional, Callable
from pathlib import Path

try:
    from tkinterdnd2 import DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False


class FileLoader(ttk.LabelFrame):
    """文件加载组件（支持拖拽）"""

    def __init__(
        self,
        parent: tk.Widget,
        title: str,
        on_file_loaded: Optional[Callable[[str], None]] = None
    ):
        """
        初始化文件加载组件

        Args:
            parent: 父组件
            title: 标题
            on_file_loaded: 文件加载回调
        """
        super().__init__(parent, text=f" {title} ", padding=10)
        self.title = title
        self.on_file_loaded = on_file_loaded
        self.file_path: Optional[str] = None

        self._setup_ui()

    def _setup_ui(self):
        """设置 UI"""
        # 拖拽区域
        self.drop_frame = tk.Frame(
            self,
            bg="#e8f4fd",
            relief="ridge",
            bd=2,
            height=80
        )
        self.drop_frame.pack(fill=tk.X, pady=(0, 5))
        self.drop_frame.pack_propagate(False)

        self.drop_label = tk.Label(
            self.drop_frame,
            text="拖拽文件到此处\n或点击选择文件",
            bg="#e8f4fd",
            fg="#666666",
            font=("Arial", 9)
        )
        self.drop_label.pack(expand=True)

        # 绑定点击事件
        self.drop_frame.bind("<Button-1>", self._on_click)
        self.drop_label.bind("<Button-1>", self._on_click)

        # 绑定拖拽事件
        if HAS_DND:
            self.drop_frame.drop_target_register(DND_FILES)
            self.drop_frame.dnd_bind("<<Drop>>", self._on_drop)
            self.drop_label.drop_target_register(DND_FILES)
            self.drop_label.dnd_bind("<<Drop>>", self._on_drop)

        # 文件信息
        self.file_label = ttk.Label(self, text="未选择文件", foreground="gray")
        self.file_label.pack(fill=tk.X)

    def _on_click(self, event):
        """点击选择文件"""
        file_path = filedialog.askopenfilename(
            title=f"选择{self.title}",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xls"),
                ("所有文件", "*.*")
            ]
        )
        if file_path:
            self._load_file(file_path)

    def _on_drop(self, event):
        """拖拽文件"""
        file_path = event.data
        # 处理 Windows 路径格式
        if file_path.startswith("{") and file_path.endswith("}"):
            file_path = file_path[1:-1]

        if file_path.lower().endswith((".xlsx", ".xls")):
            self._load_file(file_path)

    def _load_file(self, file_path: str):
        """加载文件"""
        self.file_path = file_path
        path = Path(file_path)

        self.file_label.config(
            text=f"✓ {path.name}",
            foreground="green"
        )
        self.drop_label.config(
            text=f"已加载:\n{path.name}",
            fg="green"
        )

        if self.on_file_loaded:
            self.on_file_loaded(file_path)

    def clear(self):
        """清除文件"""
        self.file_path = None
        self.file_label.config(text="未选择文件", foreground="gray")
        self.drop_label.config(
            text="拖拽文件到此处\n或点击选择文件",
            fg="#666666"
        )
```

- [ ] **Step 3: 创建主窗口 `ui/main_window.py`**

```python
"""主窗口"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, List
import threading

try:
    from tkinterdnd2 import TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False
    import tkinter as tk_base
    TkinterDnD = type("TkinterDnD", (), {"Tk": tk_base.Tk})

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))
from bank_reconciliation.models import Transaction, MatchResult, MatchStatus
from bank_reconciliation.core import BankStatementParser, AccountLedgerParser, MatchEngine, ResultExporter
from bank_reconciliation.config import AppConfig, BANK_TEMPLATES, ACCOUNT_TEMPLATES
from .file_loader import FileLoader


class BankReconciliationApp:
    """银行对账助手主应用"""

    def __init__(self, root: Optional[tk.Tk] = None):
        """
        初始化应用

        Args:
            root: Tkinter 根窗口
        """
        if root is None:
            root = TkinterDnD.Tk() if HAS_DND else tk.Tk()

        self.root = root
        self.config = AppConfig()

        # 数据
        self.bank_txns: List[Transaction] = []
        self.account_txns: List[Transaction] = []
        self.match_results: List[MatchResult] = []

        # 解析器和引擎
        self.bank_parser: Optional[BankStatementParser] = None
        self.account_parser: Optional[AccountLedgerParser] = None
        self.match_engine = MatchEngine(self.config)
        self.exporter = ResultExporter(self.config.output_dir)

        self._setup_window()
        self._setup_styles()
        self._setup_ui()

    def _setup_window(self):
        """设置窗口"""
        self.root.title("智能银行对账助手")
        self.root.geometry(f"{self.config.window_width}x{self.config.window_height}")
        self.root.minsize(800, 600)

    def _setup_styles(self):
        """设置样式"""
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("Title.TLabel", font=("Arial", 14, "bold"))
        style.configure("Heading.TLabel", font=("Arial", 10, "bold"))
        style.configure("Success.TLabel", foreground="green")
        style.configure("Warning.TLabel", foreground="orange")
        style.configure("Action.TButton", font=("Arial", 10, "bold"))

    def _setup_ui(self):
        """设置 UI"""
        # 主容器
        self.main_container = ttk.Frame(self.root, padding=10)
        self.main_container.pack(fill=tk.BOTH, expand=True)

        # 标题
        title_frame = ttk.Frame(self.main_container)
        title_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(title_frame, text="智能银行对账助手", style="Title.TLabel").pack(side=tk.LEFT)

        # 文件选择区域
        self._create_file_section()

        # 操作按钮区域
        self._create_action_section()

        # 结果预览区域
        self._create_result_section()

        # 状态栏
        self._create_status_section()

    def _create_file_section(self):
        """创建文件选择区域"""
        file_frame = ttk.Frame(self.main_container)
        file_frame.pack(fill=tk.X, pady=(0, 10))

        # 银行流水加载器
        self.bank_loader = FileLoader(
            file_frame,
            "银行流水",
            on_file_loaded=self._on_bank_file_loaded
        )
        self.bank_loader.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        # 账务数据加载器
        self.account_loader = FileLoader(
            file_frame,
            "账务数据",
            on_file_loaded=self._on_account_file_loaded
        )
        self.account_loader.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))

        # 模板选择
        template_frame = ttk.LabelFrame(self.main_container, text=" 模板设置 ", padding=5)
        template_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(template_frame, text="银行模板:").pack(side=tk.LEFT, padx=(0, 5))
        self.bank_template_var = tk.StringVar(value="自定义")
        self.bank_template_combo = ttk.Combobox(
            template_frame,
            textvariable=self.bank_template_var,
            values=list(BANK_TEMPLATES.keys()) + ["自定义"],
            state="readonly",
            width=12
        )
        self.bank_template_combo.pack(side=tk.LEFT, padx=(0, 20))

        ttk.Label(template_frame, text="账务模板:").pack(side=tk.LEFT, padx=(0, 5))
        self.account_template_var = tk.StringVar(value="用友")
        self.account_template_combo = ttk.Combobox(
            template_frame,
            textvariable=self.account_template_var,
            values=list(ACCOUNT_TEMPLATES.keys()) + ["自定义"],
            state="readonly",
            width=12
        )
        self.account_template_combo.pack(side=tk.LEFT)

    def _create_action_section(self):
        """创建操作按钮区域"""
        action_frame = ttk.Frame(self.main_container)
        action_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(
            action_frame,
            text="开始匹配",
            style="Action.TButton",
            command=self._start_match
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            action_frame,
            text="导出结果",
            command=self._export_results
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            action_frame,
            text="清空数据",
            command=self._clear_data
        ).pack(side=tk.LEFT, padx=5)

    def _create_result_section(self):
        """创建结果预览区域"""
        result_frame = ttk.LabelFrame(self.main_container, text=" 匹配结果预览 ", padding=5)
        result_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 创建 Treeview
        columns = ("status", "bank_date", "bank_amount", "account_date", "account_amount", "summary")
        self.result_tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=15)

        self.result_tree.heading("status", text="状态")
        self.result_tree.heading("bank_date", text="银行日期")
        self.result_tree.heading("bank_amount", text="银行金额")
        self.result_tree.heading("account_date", text="账务日期")
        self.result_tree.heading("account_amount", text="账务金额")
        self.result_tree.heading("summary", text="摘要")

        self.result_tree.column("status", width=80)
        self.result_tree.column("bank_date", width=100)
        self.result_tree.column("bank_amount", width=100)
        self.result_tree.column("account_date", width=100)
        self.result_tree.column("account_amount", width=100)
        self.result_tree.column("summary", width=200)

        # 滚动条
        scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_tree.yview)
        self.result_tree.configure(yscrollcommand=scrollbar.set)

        self.result_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def _create_status_section(self):
        """创建状态栏"""
        status_frame = ttk.Frame(self.main_container)
        status_frame.pack(fill=tk.X)

        self.status_label = ttk.Label(
            status_frame,
            text="就绪 | 银行: 0 笔 | 账务: 0 笔 | 匹配: 0 笔",
            foreground="gray"
        )
        self.status_label.pack(side=tk.LEFT)

    def _on_bank_file_loaded(self, file_path: str):
        """银行文件加载回调"""
        try:
            template = self.bank_template_var.get()
            self.bank_parser = BankStatementParser(template_name=template if template != "自定义" else None)
            self.bank_txns = self.bank_parser.parse(file_path)
            self._update_status()
            messagebox.showinfo("成功", f"已加载 {len(self.bank_txns)} 笔银行流水记录")
        except Exception as e:
            messagebox.showerror("错误", f"加载银行流水失败: {e}")

    def _on_account_file_loaded(self, file_path: str):
        """账务文件加载回调"""
        try:
            template = self.account_template_var.get()
            self.account_parser = AccountLedgerParser(template_name=template if template != "自定义" else None)
            self.account_txns = self.account_parser.parse(file_path)
            self._update_status()
            messagebox.showinfo("成功", f"已加载 {len(self.account_txns)} 笔账务记录")
        except Exception as e:
            messagebox.showerror("错误", f"加载账务数据失败: {e}")

    def _start_match(self):
        """开始匹配"""
        if not self.bank_txns and not self.account_txns:
            messagebox.showwarning("警告", "请先加载银行流水和账务数据")
            return

        # 在后台线程执行匹配
        def do_match():
            self.match_results = self.match_engine.match(self.bank_txns, self.account_txns)
            self.root.after(0, self._on_match_complete)

        threading.Thread(target=do_match, daemon=True).start()
        self.status_label.config(text="正在匹配...")

    def _on_match_complete(self):
        """匹配完成回调"""
        self._display_results()
        self._update_status()

        # 统计
        matched = sum(1 for r in self.match_results if r.is_matched)
        total = len(self.match_results)
        rate = matched / total * 100 if total > 0 else 0

        messagebox.showinfo("匹配完成", f"匹配完成！\n匹配率: {rate:.1f}%\n已匹配: {matched} 笔")

    def _display_results(self):
        """显示匹配结果"""
        # 清空现有数据
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        # 添加结果
        for result in self.match_results:
            status = result.status_display
            bank_date = result.bank_txn.date.strftime("%Y-%m-%d") if result.bank_txn else "-"
            bank_amount = f"{float(result.bank_txn.amount):,.2f}" if result.bank_txn else "-"
            account_date = result.account_txn.date.strftime("%Y-%m-%d") if result.account_txn else "-"
            account_amount = f"{float(result.account_txn.amount):,.2f}" if result.account_txn else "-"
            summary = (result.bank_txn.summary[:20] if result.bank_txn else
                      result.account_txn.summary[:20] if result.account_txn else "-")

            tags = ("matched",) if result.is_matched else ("unmatched",)
            self.result_tree.insert("", tk.END, values=(
                status, bank_date, bank_amount, account_date, account_amount, summary
            ), tags=tags)

        # 设置标签样式
        self.result_tree.tag_configure("matched", background="#C6EFCE")
        self.result_tree.tag_configure("unmatched", background="#FFEB9C")

    def _export_results(self):
        """导出结果"""
        if not self.match_results:
            messagebox.showwarning("警告", "没有匹配结果可导出")
            return

        try:
            files = self.exporter.export_all(self.match_results)
            messagebox.showinfo(
                "导出成功",
                f"已导出以下文件:\n" + "\n".join(files.values())
            )
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {e}")

    def _clear_data(self):
        """清空数据"""
        self.bank_txns = []
        self.account_txns = []
        self.match_results = []
        self.bank_loader.clear()
        self.account_loader.clear()

        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        self._update_status()

    def _update_status(self):
        """更新状态栏"""
        matched = sum(1 for r in self.match_results if r.is_matched)
        self.status_label.config(
            text=f"就绪 | 银行: {len(self.bank_txns)} 笔 | 账务: {len(self.account_txns)} 笔 | 匹配: {matched} 笔"
        )

    def run(self):
        """运行应用"""
        self.root.mainloop()
```

- [ ] **Step 4: 提交**

```bash
git add tools/bank_reconciliation/ui/
git commit -m "feat(bank-reconciliation): 实现 GUI 主窗口

- 创建 FileLoader 文件加载组件（支持拖拽）
- 创建 BankReconciliationApp 主窗口
- 实现文件加载、匹配、导出功能
- 结果预览表格（带条件格式）

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>"
```

---

## Task 6: 程序入口和模板配置

**Files:**
- Create: `tools/bank_reconciliation/main.py`
- Create: `tools/bank_reconciliation/templates/bank_templates.json`
- Create: `tools/bank_reconciliation/templates/account_templates.json`

- [ ] **Step 1: 创建程序入口 `main.py`**

```python
"""
银行对账助手 - 主程序入口

用法:
    python -m bank_reconciliation
    或
    python tools/bank_reconciliation/main.py
"""

import sys
from pathlib import Path

# 添加项目根目录到路径
project_root = Path(__file__).parent.parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from bank_reconciliation.ui import BankReconciliationApp


def main():
    """主函数"""
    app = BankReconciliationApp()
    app.run()


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: 创建银行模板配置 `templates/bank_templates.json`**

```json
{
    "工商银行": {
        "date": "交易日期",
        "debit": "支出",
        "credit": "收入",
        "balance": "余额",
        "summary": "摘要"
    },
    "建设银行": {
        "date": "交易时间",
        "debit": "借方",
        "credit": "贷方",
        "balance": "余额",
        "summary": "交易摘要"
    },
    "农业银行": {
        "date": "交易日期",
        "debit": "支出金额",
        "credit": "存入金额",
        "balance": "账户余额",
        "summary": "交易摘要"
    },
    "中国银行": {
        "date": "交易日期",
        "debit": "借方发生额",
        "credit": "贷方发生额",
        "balance": "余额",
        "summary": "摘要"
    }
}
```

- [ ] **Step 3: 创建账务模板配置 `templates/account_templates.json`**

```json
{
    "用友": {
        "date": "日期",
        "debit": "借方",
        "credit": "贷方",
        "voucher": "凭证号",
        "summary": "摘要"
    },
    "金蝶": {
        "date": "日期",
        "debit": "借方金额",
        "credit": "贷方金额",
        "voucher": "凭证字号",
        "summary": "摘要"
    }
}
```

- [ ] **Step 4: 提交**

```bash
git add tools/bank_reconciliation/main.py tools/bank_reconciliation/templates/
git commit -m "feat(bank-reconciliation): 添加程序入口和模板配置

- 添加 main.py 程序入口
- 添加银行模板配置（工行、建行、农行、中行）
- 添加账务模板配置（用友、金蝶）

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>"
```

---

## Task 7: 测试和验证

**Files:**
- Create: `tests/test_bank_reconciliation.py`

- [ ] **Step 1: 创建测试文件 `tests/test_bank_reconciliation.py`**

```python
"""
银行对账助手测试
"""

import pytest
from datetime import date
from decimal import Decimal

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent / "tools"))

from bank_reconciliation.models import Transaction, MatchResult, MatchStatus
from bank_reconciliation.core import MatchEngine


class TestTransaction:
    """交易记录测试"""

    def test_create_transaction(self):
        """测试创建交易记录"""
        txn = Transaction(
            date=date(2026, 5, 8),
            amount=Decimal("1000.00"),
            summary="测试交易",
            direction="debit",
            source="bank"
        )
        assert txn.date == date(2026, 5, 8)
        assert txn.amount == Decimal("1000.00")
        assert txn.direction == "debit"

    def test_match_key(self):
        """测试匹配键生成"""
        txn = Transaction(
            date=date(2026, 5, 8),
            amount=Decimal("1000.00"),
            summary="测试",
            direction="debit",
            source="bank"
        )
        assert txn.match_key() == (Decimal("1000.00"), "debit")


class TestMatchEngine:
    """匹配引擎测试"""

    def test_exact_match(self):
        """测试精确匹配"""
        engine = MatchEngine()

        bank_txns = [
            Transaction(
                date=date(2026, 5, 8),
                amount=Decimal("1000.00"),
                summary="转账",
                direction="debit",
                source="bank"
            )
        ]

        account_txns = [
            Transaction(
                date=date(2026, 5, 8),
                amount=Decimal("1000.00"),
                summary="转账",
                direction="debit",
                source="account"
            )
        ]

        results = engine.match(bank_txns, account_txns)

        assert len(results) == 1
        assert results[0].status == MatchStatus.MATCHED

    def test_date_tolerance_match(self):
        """测试日期容差匹配"""
        engine = MatchEngine()

        bank_txns = [
            Transaction(
                date=date(2026, 5, 8),
                amount=Decimal("1000.00"),
                summary="转账",
                direction="debit",
                source="bank"
            )
        ]

        account_txns = [
            Transaction(
                date=date(2026, 5, 9),  # 日期相差1天
                amount=Decimal("1000.00"),
                summary="转账",
                direction="debit",
                source="account"
            )
        ]

        results = engine.match(bank_txns, account_txns)

        assert len(results) == 1
        assert results[0].status == MatchStatus.MATCHED

    def test_unmatched(self):
        """测试未匹配项"""
        engine = MatchEngine()

        bank_txns = [
            Transaction(
                date=date(2026, 5, 8),
                amount=Decimal("1000.00"),
                summary="转账",
                direction="debit",
                source="bank"
            )
        ]

        account_txns = [
            Transaction(
                date=date(2026, 5, 8),
                amount=Decimal("2000.00"),  # 金额不同
                summary="转账",
                direction="debit",
                source="account"
            )
        ]

        results = engine.match(bank_txns, account_txns)

        assert len(results) == 2
        statuses = [r.status for r in results]
        assert MatchStatus.BANK_UNMATCHED in statuses
        assert MatchStatus.ACCOUNT_UNMATCHED in statuses

    def test_manual_match(self):
        """测试手动匹配"""
        engine = MatchEngine()

        bank_txn = Transaction(
            date=date(2026, 5, 8),
            amount=Decimal("1000.00"),
            summary="转账",
            direction="debit",
            source="bank"
        )

        account_txn = Transaction(
            date=date(2026, 5, 8),
            amount=Decimal("2000.00"),
            summary="转账",
            direction="debit",
            source="account"
        )

        results = [
            MatchResult(bank_txn=bank_txn, status=MatchStatus.BANK_UNMATCHED),
            MatchResult(account_txn=account_txn, status=MatchStatus.ACCOUNT_UNMATCHED)
        ]

        new_results = engine.manual_match(results, bank_txn, account_txn)

        # 检查手动匹配结果
        matched = [r for r in new_results if r.status == MatchStatus.MANUAL_MATCHED]
        assert len(matched) == 1
        assert matched[0].bank_txn == bank_txn
        assert matched[0].account_txn == account_txn
```

- [ ] **Step 2: 运行测试**

```bash
cd C:/Users/30726/OneDrive/panda
python -m pytest tests/test_bank_reconciliation.py -v
```

Expected: 所有测试通过

- [ ] **Step 3: 提交**

```bash
git add tests/test_bank_reconciliation.py
git commit -m "test(bank-reconciliation): 添加单元测试

- 测试交易记录创建和匹配键生成
- 测试精确匹配和日期容差匹配
- 测试未匹配项标记
- 测试手动匹配功能

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>"
```

---

## 验证计划

1. **运行测试**: `python -m pytest tests/test_bank_reconciliation.py -v`
2. **启动程序**: `python -m bank_reconciliation` 或 `python tools/bank_reconciliation/main.py`
3. **准备测试数据**: 创建银行流水 Excel 和账务 Excel 各一份
4. **验证功能**:
   - 拖拽/选择文件加载
   - 自动匹配
   - 结果预览
   - 导出结果
