# Excel 多表一键合并工具箱实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 开发一个 GUI 工具，自动读取文件夹内所有格式相近的 Excel 文件，智能匹配科目名称，合并生成标准化汇总表。

**Architecture:** 采用分层架构，复用现有 bank_reconciliation 模块的设计模式。数据层使用 dataclass 定义模型，核心层实现解析、匹配、导出逻辑，UI 层使用 ttkbootstrap + tkinterdnd2 构建图形界面。

**Tech Stack:** Python 3.8+, pandas, openpyxl, ttkbootstrap, tkinterdnd2, pytest

---

## 文件结构

```
tools/excel_merger/
├── __init__.py           # 模块入口，导出主要类
├── config.py             # 配置管理、内置科目映射表
├── main.py               # 主程序入口
├── core/
│   ├── __init__.py       # 导出解析器、匹配器、导出器
│   ├── parser.py         # Excel 解析器（公司、月份、科目识别）
│   ├── matcher.py        # 科目名称匹配引擎
│   └── exporter.py       # 结果导出器（透视表、未匹配报告）
├── models/
│   ├── __init__.py       # 导出数据模型
│   ├── source_data.py    # SourceRecord, ParsedFile, SheetData
│   └── merge_result.py   # MergeResult
├── templates/
│   └── default_templates.json  # 内置模板配置
└── ui/
    ├── __init__.py       # 导出 UI 组件
    ├── main_window.py    # 主窗口
    └── mapping_editor.py # 科目映射编辑器

tests/
└── test_excel_merger.py  # 单元测试
```

---

## Task 1: 创建模块目录结构和数据模型

**Files:**
- Create: `tools/excel_merger/__init__.py`
- Create: `tools/excel_merger/models/__init__.py`
- Create: `tools/excel_merger/models/source_data.py`
- Create: `tools/excel_merger/models/merge_result.py`

- [ ] **Step 1: 创建目录结构**

```bash
mkdir -p tools/excel_merger/core tools/excel_merger/models tools/excel_merger/templates tools/excel_merger/ui
```

- [ ] **Step 2: 创建模块入口 `tools/excel_merger/__init__.py`**

```python
"""Excel 多表合并工具箱"""

from .models import SourceRecord, ParsedFile, SheetData, MergeResult

__all__ = ["SourceRecord", "ParsedFile", "SheetData", "MergeResult"]
```

- [ ] **Step 3: 创建模型模块入口 `tools/excel_merger/models/__init__.py`**

```python
"""数据模型"""

from .source_data import SourceRecord, ParsedFile, SheetData
from .merge_result import MergeResult

__all__ = ["SourceRecord", "ParsedFile", "SheetData", "MergeResult"]
```

- [ ] **Step 4: 创建源数据模型 `tools/excel_merger/models/source_data.py`**

```python
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
```

- [ ] **Step 5: 创建合并结果模型 `tools/excel_merger/models/merge_result.py`**

```python
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
```

- [ ] **Step 6: 提交**

```bash
git add tools/excel_merger/
git commit -m "feat(excel-merger): 添加数据模型

- 添加 SourceRecord 单条记录模型
- 添加 SheetData 和 ParsedFile 文件解析模型
- 添加 MergeResult 合并结果模型

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>"
```

---

## Task 2: 实现配置管理

**Files:**
- Create: `tools/excel_merger/config.py`
- Create: `tools/excel_merger/templates/default_templates.json`

- [ ] **Step 1: 创建配置模块 `tools/excel_merger/config.py`**

```python
"""配置管理"""

from dataclasses import dataclass, field
from typing import Dict, List
import json
from pathlib import Path


@dataclass
class AppConfig:
    """应用配置"""
    # 界面配置
    window_width: int = 1000
    window_height: int = 700

    # 输出配置
    output_dir: str = "output"

    # 识别配置
    company_patterns: List[str] = field(default_factory=lambda: [
        "{公司}.xlsx",
        "{公司}_{年份}.xlsx",
        "{年份}年{公司}.xlsx",
        "{公司}分公司.xlsx",
        "{公司}子公司.xlsx"
    ])

    # 月份关键词
    month_keywords: Dict[int, List[str]] = field(default_factory=lambda: {
        1: ["1月", "一月", "Jan", "January", "01", "1"],
        2: ["2月", "二月", "Feb", "February", "02", "2"],
        3: ["3月", "三月", "Mar", "March", "03", "3"],
        4: ["4月", "四月", "Apr", "April", "04", "4"],
        5: ["5月", "五月", "May", "05", "5"],
        6: ["6月", "六月", "Jun", "June", "06", "6"],
        7: ["7月", "七月", "Jul", "July", "07", "7"],
        8: ["8月", "八月", "Aug", "August", "08", "8"],
        9: ["9月", "九月", "Sep", "September", "09", "9"],
        10: ["10月", "十月", "Oct", "October", "10"],
        11: ["11月", "十一月", "Nov", "November", "11"],
        12: ["12月", "十二月", "Dec", "December", "12"]
    })

    # 科目列关键词
    account_keywords: List[str] = field(default_factory=lambda: [
        "科目", "项目", "名称", "摘要", "费用项目", "成本项目", "收支项目"
    ])

    @classmethod
    def load(cls, config_path: str = None) -> "AppConfig":
        """加载配置"""
        if config_path and Path(config_path).exists():
            with open(config_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                return cls(**data)
        return cls()


# 内置科目映射表
DEFAULT_ACCOUNT_MAPPING: Dict[str, str] = {
    # 收入类
    "销售收入": "主营业务收入",
    "产品销售": "主营业务收入",
    "销售": "主营业务收入",
    "服务收入": "其他业务收入",
    "其他收入": "其他业务收入",
    "营业外收入": "营业外收入",

    # 成本类
    "销售成本": "主营业务成本",
    "产品成本": "主营业务成本",
    "成本": "主营业务成本",
    "服务成本": "其他业务成本",

    # 费用类 - 管理费用
    "管理费用": "管理费用",
    "办公费": "管理费用-办公费",
    "差旅费": "管理费用-差旅费",
    "交通费": "管理费用-交通费",
    "人员工资": "管理费用-工资",
    "工资": "管理费用-工资",
    "福利费": "管理费用-福利费",
    "折旧费": "管理费用-折旧费",
    "水电费": "管理费用-水电费",
    "物业费": "管理费用-物业费",
    "通讯费": "管理费用-通讯费",
    "招待费": "管理费用-招待费",

    # 费用类 - 销售费用
    "销售费用": "销售费用",
    "广告费": "销售费用-广告费",
    "推广费": "销售费用-推广费",
    "运费": "销售费用-运费",

    # 费用类 - 财务费用
    "财务费用": "财务费用",
    "利息支出": "财务费用-利息支出",
    "手续费": "财务费用-手续费",

    # 税费
    "税费": "税金及附加",
    "增值税": "应交税费-增值税",
    "所得税": "所得税费用",
}
```

- [ ] **Step 2: 创建模板目录 `tools/excel_merger/templates/default_templates.json`**

```json
{
  "templates": [
    {
      "name": "标准收入成本表",
      "description": "标准收入成本表模板",
      "account_column_keywords": ["科目", "项目", "名称"],
      "month_column_pattern": "{月份}",
      "data_start_row": 1
    }
  ]
}
```

- [ ] **Step 3: 提交**

```bash
git add tools/excel_merger/config.py tools/excel_merger/templates/
git commit -m "feat(excel-merger): 添加配置管理和内置科目映射

- 添加 AppConfig 配置类
- 添加内置科目映射表 DEFAULT_ACCOUNT_MAPPING
- 添加月份识别关键词
- 添加模板配置文件

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>"
```

---

## Task 3: 实现 Excel 解析器

**Files:**
- Create: `tools/excel_merger/core/__init__.py`
- Create: `tools/excel_merger/core/parser.py`

- [ ] **Step 1: 创建核心模块入口 `tools/excel_merger/core/__init__.py`**

```python
"""核心模块"""

from .parser import ExcelParser
from .matcher import AccountMatcher
from .exporter import ResultExporter

__all__ = ["ExcelParser", "AccountMatcher", "ResultExporter"]
```

- [ ] **Step 2: 创建解析器 `tools/excel_merger/core/parser.py`**

```python
"""Excel 解析器"""

import os
import re
from pathlib import Path
from typing import List, Optional, Dict, Any
from decimal import Decimal

import pandas as pd

import sys
sys.path.insert(0, str(Path(__file__).parent.parent.parent))
from excel_merger.models import SourceRecord, ParsedFile, SheetData
from excel_merger.config import AppConfig, DEFAULT_ACCOUNT_MAPPING


class ExcelParser:
    """Excel 文件解析器"""

    def __init__(self, config: AppConfig = None):
        """
        初始化解析器

        Args:
            config: 应用配置
        """
        self.config = config or AppConfig()
        self.account_mapping = DEFAULT_ACCOUNT_MAPPING.copy()

    def parse_folder(self, folder_path: str) -> List[ParsedFile]:
        """
        解析文件夹内所有 Excel 文件

        Args:
            folder_path: 文件夹路径

        Returns:
            解析后的文件列表
        """
        results = []
        folder = Path(folder_path)

        for file_path in folder.glob("*.xlsx"):
            if file_path.name.startswith("~$"):  # 跳过临时文件
                continue
            try:
                parsed = self.parse_file(str(file_path))
                results.append(parsed)
            except Exception as e:
                # 记录错误但继续处理其他文件
                error_file = ParsedFile(file_path=str(file_path))
                error_file.parse_errors.append(str(e))
                results.append(error_file)

        return results

    def parse_file(self, file_path: str) -> ParsedFile:
        """
        解析单个 Excel 文件

        Args:
            file_path: 文件路径

        Returns:
            解析后的文件数据
        """
        result = ParsedFile(file_path=file_path)

        # 从文件名识别公司
        result.company = self._extract_company_from_filename(file_path)

        # 读取所有 sheet
        xls = pd.ExcelFile(file_path)
        for sheet_name in xls.sheet_names:
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                sheet_data = self._parse_sheet(df, sheet_name, result.company, file_path)
                result.sheets.append(sheet_data)
            except Exception as e:
                result.parse_errors.append(f"Sheet '{sheet_name}': {str(e)}")

        return result

    def _extract_company_from_filename(self, file_path: str) -> str:
        """
        从文件名提取公司名称

        Args:
            file_path: 文件路径

        Returns:
            公司名称
        """
        filename = Path(file_path).stem

        # 移除常见后缀
        name = re.sub(r'(\d{4}年?|年|\d+月|分公司|子公司|报表|数据)', '', filename)
        name = re.sub(r'[_\-]', '', name)

        return name.strip() if name.strip() else filename

    def _parse_sheet(
        self,
        df: pd.DataFrame,
        sheet_name: str,
        company: str,
        file_path: str
    ) -> SheetData:
        """
        解析单个 Sheet

        Args:
            df: DataFrame
            sheet_name: Sheet 名称
            company: 公司名称
            file_path: 文件路径

        Returns:
            Sheet 数据
        """
        result = SheetData(sheet_name=sheet_name)

        if df.empty:
            return result

        # 识别科目列
        account_col = self._find_account_column(df)
        if account_col is None:
            result.accounts_found = []
            return result

        # 识别月份列
        month_cols = self._find_month_columns(df, account_col)

        # 提取科目列表
        result.accounts_found = df[account_col].dropna().astype(str).str.strip().tolist()

        # 从 sheet 名识别月份
        sheet_month = self._extract_month_from_sheet_name(sheet_name)

        # 解析数据行
        for idx, row in df.iterrows():
            account_name = str(row.get(account_col, "")).strip()
            if not account_name or account_name in ["合计", "小计", "总计", "nan"]:
                continue

            # 如果有月份列，按月份列提取数据
            if month_cols:
                for month, col in month_cols.items():
                    amount = self._clean_amount(row.get(col))
                    if amount > 0:
                        record = SourceRecord(
                            company=company,
                            month=month,
                            account_name=account_name,
                            amount=amount,
                            source_file=file_path,
                            sheet_name=sheet_name
                        )
                        result.records.append(record)
                        if month not in result.months_found:
                            result.months_found.append(month)
            # 否则使用 sheet 名中的月份
            elif sheet_month:
                # 查找金额列（非科目列的第一个数值列）
                amount_col = self._find_amount_column(df, account_col)
                if amount_col:
                    amount = self._clean_amount(row.get(amount_col))
                    if amount > 0:
                        record = SourceRecord(
                            company=company,
                            month=sheet_month,
                            account_name=account_name,
                            amount=amount,
                            source_file=file_path,
                            sheet_name=sheet_name
                        )
                        result.records.append(record)
                        if sheet_month not in result.months_found:
                            result.months_found.append(sheet_month)

        return result

    def _find_account_column(self, df: pd.DataFrame) -> Optional[str]:
        """
        查找科目列

        Args:
            df: DataFrame

        Returns:
            科目列名或 None
        """
        for col in df.columns:
            col_str = str(col).lower()
            for keyword in self.config.account_keywords:
                if keyword.lower() in col_str:
                    return col
        # 如果没找到，返回第一列
        return df.columns[0] if len(df.columns) > 0 else None

    def _find_month_columns(
        self,
        df: pd.DataFrame,
        account_col: str
    ) -> Dict[int, str]:
        """
        查找月份列

        Args:
            df: DataFrame
            account_col: 科目列名

        Returns:
            月份到列名的映射
        """
        month_cols = {}

        for col in df.columns:
            if col == account_col:
                continue
            col_str = str(col)
            month = self._extract_month_from_string(col_str)
            if month:
                month_cols[month] = col

        return month_cols

    def _extract_month_from_string(self, text: str) -> Optional[int]:
        """
        从字符串提取月份

        Args:
            text: 输入字符串

        Returns:
            月份 (1-12) 或 None
        """
        text = str(text).strip()

        for month, keywords in self.config.month_keywords.items():
            for keyword in keywords:
                if keyword in text:
                    return month

        # 尝试匹配数字月份
        match = re.search(r'(\d{1,2})月?', text)
        if match:
            month = int(match.group(1))
            if 1 <= month <= 12:
                return month

        return None

    def _extract_month_from_sheet_name(self, sheet_name: str) -> Optional[int]:
        """
        从 Sheet 名称提取月份

        Args:
            sheet_name: Sheet 名称

        Returns:
            月份 (1-12) 或 None
        """
        return self._extract_month_from_string(sheet_name)

    def _find_amount_column(
        self,
        df: pd.DataFrame,
        account_col: str
    ) -> Optional[str]:
        """
        查找金额列

        Args:
            df: DataFrame
            account_col: 科目列名

        Returns:
            金额列名或 None
        """
        for col in df.columns:
            if col == account_col:
                continue
            # 检查是否为数值列
            if pd.api.types.is_numeric_dtype(df[col]):
                return col
        return None

    def _clean_amount(self, value: Any) -> Decimal:
        """
        清洗金额数据

        Args:
            value: 原始值

        Returns:
            Decimal 金额
        """
        if pd.isna(value) or value == "" or value == "-":
            return Decimal("0.00")
        if isinstance(value, (int, float)):
            return Decimal(str(round(value, 2)))
        if isinstance(value, str):
            value = value.replace(",", "").replace("，", "").strip()
            if value == "" or value == "-":
                return Decimal("0.00")
            try:
                return Decimal(value).quantize(Decimal("0.01"))
            except:
                return Decimal("0.00")
        return Decimal("0.00")
```

- [ ] **Step 3: 提交**

```bash
git add tools/excel_merger/core/
git commit -m "feat(excel-merger): 实现 Excel 解析器

- 添加 ExcelParser 类
- 实现文件夹遍历和文件解析
- 实现公司名称识别（从文件名）
- 实现月份识别（列标题、Sheet名）
- 实现科目列识别

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>"
```

---

## Task 4: 实现科目匹配引擎

**Files:**
- Create: `tools/excel_merger/core/matcher.py`

- [ ] **Step 1: 创建科目匹配器 `tools/excel_merger/core/matcher.py`**

```python
"""科目匹配引擎"""

from typing import List, Dict, Set
from pathlib import Path
import json

import sys
sys.path.insert(0, str(Path(__file__).parent.parent.parent))
from excel_merger.models import SourceRecord, ParsedFile, MergeResult
from excel_merger.config import DEFAULT_ACCOUNT_MAPPING


class AccountMatcher:
    """科目名称匹配器"""

    def __init__(self, custom_mapping: Dict[str, str] = None):
        """
        初始化匹配器

        Args:
            custom_mapping: 自定义科目映射
        """
        self.account_mapping = DEFAULT_ACCOUNT_MAPPING.copy()
        if custom_mapping:
            self.account_mapping.update(custom_mapping)

    def match(self, parsed_files: List[ParsedFile]) -> MergeResult:
        """
        匹配所有文件的科目名称

        Args:
            parsed_files: 解析后的文件列表

        Returns:
            合并结果
        """
        result = MergeResult()

        # 收集所有记录
        all_records: List[SourceRecord] = []
        for parsed in parsed_files:
            all_records.extend(parsed.all_records)

        # 匹配科目名称
        unmatched: Set[str] = set()
        for record in all_records:
            standardized = self._match_account(record.account_name)
            record.standardized_name = standardized
            if standardized == record.account_name and record.account_name not in self.account_mapping:
                unmatched.add(record.account_name)

        result.records = all_records
        result.unmatched_accounts = unmatched

        return result

    def _match_account(self, account_name: str) -> str:
        """
        匹配单个科目名称

        Args:
            account_name: 原始科目名称

        Returns:
            标准化科目名称
        """
        name = account_name.strip()

        # 精确匹配
        if name in self.account_mapping:
            return self.account_mapping[name]

        # 模糊匹配（包含关系）
        for key, value in self.account_mapping.items():
            if key in name or name in key:
                return value

        # 未匹配，返回原始名称
        return name

    def add_mapping(self, original: str, standardized: str):
        """
        添加科目映射

        Args:
            original: 原始科目名称
            standardized: 标准化科目名称
        """
        self.account_mapping[original] = standardized

    def remove_mapping(self, original: str):
        """
        删除科目映射

        Args:
            original: 原始科目名称
        """
        if original in self.account_mapping:
            del self.account_mapping[original]

    def save_mapping(self, file_path: str):
        """
        保存科目映射到文件

        Args:
            file_path: 文件路径
        """
        path = Path(file_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(self.account_mapping, f, ensure_ascii=False, indent=2)

    def load_mapping(self, file_path: str):
        """
        从文件加载科目映射

        Args:
            file_path: 文件路径
        """
        path = Path(file_path)
        if path.exists():
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                self.account_mapping.update(data)
```

- [ ] **Step 2: 提交**

```bash
git add tools/excel_merger/core/matcher.py
git commit -m "feat(excel-merger): 实现科目匹配引擎

- 添加 AccountMatcher 类
- 实现科目名称精确匹配和模糊匹配
- 支持自定义映射的添加和删除
- 支持映射的保存和加载

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>"
```

---

## Task 5: 实现结果导出器

**Files:**
- Create: `tools/excel_merger/core/exporter.py`

- [ ] **Step 1: 创建导出器 `tools/excel_merger/core/exporter.py`**

```python
"""结果导出器"""

from datetime import datetime
from typing import List, Dict
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

import sys
sys.path.insert(0, str(Path(__file__).parent.parent.parent))
from excel_merger.models import MergeResult, SourceRecord


class ResultExporter:
    """结果导出器"""

    # 样式定义
    HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    HEADER_FONT = Font(bold=True, color="FFFFFF")
    TOTAL_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
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

    def export_all(self, result: MergeResult) -> Dict[str, str]:
        """
        导出所有结果

        Args:
            result: 合并结果

        Returns:
            输出文件路径字典
        """
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        return {
            "pivot_table": self.export_pivot_table(result, timestamp),
            "unmatched": self.export_unmatched_report(result, timestamp)
        }

    def export_pivot_table(self, result: MergeResult, timestamp: str = None) -> str:
        """
        导出透视表

        Args:
            result: 合并结果
            timestamp: 时间戳

        Returns:
            文件路径
        """
        if timestamp is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        file_path = self.output_dir / f"合并汇总表_{timestamp}.xlsx"

        # 构建透视表数据
        records = result.records
        if not records:
            # 创建空文件
            wb = Workbook()
            ws = wb.active
            ws.title = "汇总表"
            ws["A1"] = "无数据"
            wb.save(file_path)
            return str(file_path)

        # 创建 DataFrame
        data = []
        for r in records:
            data.append({
                "科目": r.standardized_name,
                "公司": r.company,
                "月份": r.month,
                "金额": float(r.amount)
            })

        df = pd.DataFrame(data)

        # 创建透视表：科目 × (公司-月份)
        pivot = df.pivot_table(
            index="科目",
            columns=["公司", "月份"],
            values="金额",
            aggfunc="sum",
            fill_value=0
        )

        # 创建工作簿
        wb = Workbook()
        ws = wb.active
        ws.title = "汇总表"

        # 写入表头
        ws["A1"] = "科目"

        # 获取所有公司和月份组合
        companies = sorted(df["公司"].unique())
        months = sorted(df["月份"].unique())

        # 写入多级表头
        col_idx = 2
        company_start_cols = {}  # 记录每个公司的起始列

        for company in companies:
            company_start_cols[company] = col_idx
            for month in months:
                ws.cell(row=1, column=col_idx, value=company)
                ws.cell(row=2, column=col_idx, value=f"{month}月")
                col_idx += 1

        # 合并公司名称单元格
        for company in companies:
            start_col = company_start_cols[company]
            end_col = start_col + len(months) - 1
            if len(months) > 1:
                ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=end_col)

        # 应用表头样式
        for col in range(1, col_idx):
            cell = ws.cell(row=1, column=col)
            cell.fill = self.HEADER_FILL
            cell.font = self.HEADER_FONT
            cell.alignment = Alignment(horizontal="center")
            cell.border = self.BORDER

            cell2 = ws.cell(row=2, column=col)
            cell2.fill = self.HEADER_FILL
            cell2.font = self.HEADER_FONT
            cell2.alignment = Alignment(horizontal="center")
            cell2.border = self.BORDER

        # 写入数据
        accounts = sorted(df["科目"].unique())
        for row_idx, account in enumerate(accounts, start=3):
            ws.cell(row=row_idx, column=1, value=account).border = self.BORDER

            for col_offset, company in enumerate(companies):
                for month_offset, month in enumerate(months):
                    col = 2 + col_offset * len(months) + month_offset
                    try:
                        value = pivot.loc[account, (company, month)]
                        cell = ws.cell(row=row_idx, column=col, value=value)
                        cell.number_format = '#,##0.00'
                    except KeyError:
                        cell = ws.cell(row=row_idx, column=col, value=0)
                    cell.border = self.BORDER

        # 添加合计行
        total_row = 3 + len(accounts)
        ws.cell(row=total_row, column=1, value="合计").font = Font(bold=True)
        ws.cell(row=total_row, column=1).fill = self.TOTAL_FILL
        ws.cell(row=total_row, column=1).border = self.BORDER

        for col in range(2, col_idx):
            # 计算列合计
            total = 0
            for row in range(3, total_row):
                val = ws.cell(row=row, column=col).value
                if val:
                    total += val
            cell = ws.cell(row=total_row, column=col, value=total)
            cell.number_format = '#,##0.00'
            cell.font = Font(bold=True)
            cell.fill = self.TOTAL_FILL
            cell.border = self.BORDER

        # 调整列宽
        ws.column_dimensions["A"].width = 20
        for col in range(2, col_idx):
            ws.column_dimensions[ws.cell(row=2, column=col).column_letter].width = 12

        wb.save(file_path)
        return str(file_path)

    def export_unmatched_report(self, result: MergeResult, timestamp: str = None) -> str:
        """
        导出未匹配项报告

        Args:
            result: 合并结果
            timestamp: 时间戳

        Returns:
            文件路径
        """
        if timestamp is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        file_path = self.output_dir / f"未匹配项报告_{timestamp}.xlsx"

        if not result.unmatched_accounts:
            # 创建空报告
            wb = Workbook()
            ws = wb.active
            ws.title = "未匹配项"
            ws["A1"] = "所有科目均已匹配"
            wb.save(file_path)
            return str(file_path)

        # 统计未匹配科目出现次数
        account_counts: Dict[str, int] = {}
        account_sources: Dict[str, List[str]] = {}

        for record in result.records:
            if record.account_name in result.unmatched_accounts:
                if record.account_name not in account_counts:
                    account_counts[record.account_name] = 0
                    account_sources[record.account_name] = []
                account_counts[record.account_name] += 1
                source = f"{record.company} - {record.source_file}"
                if source not in account_sources[record.account_name]:
                    account_sources[record.account_name].append(source)

        # 创建报告
        wb = Workbook()
        ws = wb.active
        ws.title = "未匹配项"

        # 表头
        headers = ["原始科目名称", "出现次数", "来源文件"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.fill = self.HEADER_FILL
            cell.font = self.HEADER_FONT
            cell.border = self.BORDER

        # 数据
        for row, (account, count) in enumerate(sorted(account_counts.items()), 2):
            ws.cell(row=row, column=1, value=account).border = self.BORDER
            ws.cell(row=row, column=2, value=count).border = self.BORDER
            ws.cell(row=row, column=3, value="; ".join(account_sources[account][:3])).border = self.BORDER

        # 调整列宽
        ws.column_dimensions["A"].width = 30
        ws.column_dimensions["B"].width = 12
        ws.column_dimensions["C"].width = 50

        wb.save(file_path)
        return str(file_path)
```

- [ ] **Step 2: 提交**

```bash
git add tools/excel_merger/core/exporter.py
git commit -m "feat(excel-merger): 实现结果导出器

- 添加 ResultExporter 类
- 实现透视表导出（科目 × 公司月份）
- 实现未匹配项报告导出
- 支持多级表头和样式化输出

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>"
```

---

## Task 6: 实现 GUI 主窗口

**Files:**
- Create: `tools/excel_merger/ui/__init__.py`
- Create: `tools/excel_merger/ui/main_window.py`

- [ ] **Step 1: 创建 UI 模块入口 `tools/excel_merger/ui/__init__.py`**

```python
"""UI 模块"""

from .main_window import ExcelMergerApp
from .mapping_editor import MappingEditor

__all__ = ["ExcelMergerApp", "MappingEditor"]
```

- [ ] **Step 2: 创建主窗口 `tools/excel_merger/ui/main_window.py`**

```python
"""主窗口"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Optional, List
from pathlib import Path
import threading

try:
    from tkinterdnd2 import TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False
    import tkinter as tk_base
    TkinterDnD = type("TkinterDnD", (), {"Tk": tk_base.Tk})

import sys
sys.path.insert(0, str(Path(__file__).parent.parent.parent))
from excel_merger.models import ParsedFile, MergeResult
from excel_merger.core import ExcelParser, AccountMatcher, ResultExporter
from excel_merger.config import AppConfig
from .mapping_editor import MappingEditor


class ExcelMergerApp:
    """Excel 多表合并工具主应用"""

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
        self.folder_path: str = ""
        self.parsed_files: List[ParsedFile] = []
        self.merge_result: Optional[MergeResult] = None

        # 核心组件
        self.parser = ExcelParser(self.config)
        self.matcher = AccountMatcher()
        self.exporter = ResultExporter(self.config.output_dir)

        self._setup_window()
        self._setup_styles()
        self._setup_ui()

    def _setup_window(self):
        """设置窗口"""
        self.root.title("Excel 多表合并工具箱")
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
        style.configure("Action.TButton", font=("Arial", 10, "bold"))

    def _setup_ui(self):
        """设置 UI"""
        # 主容器
        self.main_container = ttk.Frame(self.root, padding=10)
        self.main_container.pack(fill=tk.BOTH, expand=True)

        # 标题
        title_frame = ttk.Frame(self.main_container)
        title_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(title_frame, text="Excel 多表合并工具箱", style="Title.TLabel").pack(side=tk.LEFT)

        # 文件夹选择区域
        self._create_folder_section()

        # 文件预览和设置区域
        self._create_preview_section()

        # 结果预览区域
        self._create_result_section()

        # 操作按钮区域
        self._create_action_section()

        # 状态栏
        self._create_status_section()

    def _create_folder_section(self):
        """创建文件夹选择区域"""
        folder_frame = ttk.LabelFrame(self.main_container, text=" 源文件文件夹 ", padding=10)
        folder_frame.pack(fill=tk.X, pady=(0, 10))

        # 文件夹路径
        path_frame = ttk.Frame(folder_frame)
        path_frame.pack(fill=tk.X)

        self.folder_entry = ttk.Entry(path_frame, width=60)
        self.folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        ttk.Button(
            path_frame,
            text="选择文件夹",
            command=self._select_folder
        ).pack(side=tk.LEFT)

        # 文件计数
        self.file_count_label = ttk.Label(folder_frame, text="未选择文件夹", foreground="gray")
        self.file_count_label.pack(anchor=tk.W, pady=(5, 0))

    def _create_preview_section(self):
        """创建预览区域"""
        preview_frame = ttk.Frame(self.main_container)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 左侧：文件列表
        left_frame = ttk.LabelFrame(preview_frame, text=" 文件列表 ", padding=5)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        # 文件 Treeview
        columns = ("filename", "company", "records", "status")
        self.file_tree = ttk.Treeview(left_frame, columns=columns, show="headings", height=8)

        self.file_tree.heading("filename", text="文件名")
        self.file_tree.heading("company", text="公司")
        self.file_tree.heading("records", text="记录数")
        self.file_tree.heading("status", text="状态")

        self.file_tree.column("filename", width=200)
        self.file_tree.column("company", width=100)
        self.file_tree.column("records", width=80)
        self.file_tree.column("status", width=80)

        scrollbar = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=scrollbar.set)

        self.file_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 右侧：设置
        right_frame = ttk.LabelFrame(preview_frame, text=" 合并设置 ", padding=5)
        right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(5, 0))

        ttk.Button(
            right_frame,
            text="编辑科目映射...",
            command=self._open_mapping_editor
        ).pack(fill=tk.X, pady=5)

    def _create_result_section(self):
        """创建结果预览区域"""
        result_frame = ttk.LabelFrame(self.main_container, text=" 合并结果预览 ", padding=5)
        result_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 结果 Treeview
        columns = ("account", "company", "month", "amount")
        self.result_tree = ttk.Treeview(result_frame, columns=columns, show="headings", height=10)

        self.result_tree.heading("account", text="科目")
        self.result_tree.heading("company", text="公司")
        self.result_tree.heading("month", text="月份")
        self.result_tree.heading("amount", text="金额")

        self.result_tree.column("account", width=200)
        self.result_tree.column("company", width=100)
        self.result_tree.column("month", width=80)
        self.result_tree.column("amount", width=120)

        scrollbar = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.result_tree.yview)
        self.result_tree.configure(yscrollcommand=scrollbar.set)

        self.result_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    def _create_action_section(self):
        """创建操作按钮区域"""
        action_frame = ttk.Frame(self.main_container)
        action_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(
            action_frame,
            text="开始合并",
            style="Action.TButton",
            command=self._start_merge
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

    def _create_status_section(self):
        """创建状态栏"""
        status_frame = ttk.Frame(self.main_container)
        status_frame.pack(fill=tk.X)

        self.status_label = ttk.Label(
            status_frame,
            text="就绪 | 请选择包含 Excel 文件的文件夹",
            foreground="gray"
        )
        self.status_label.pack(side=tk.LEFT)

    def _select_folder(self):
        """选择文件夹"""
        folder = filedialog.askdirectory(title="选择包含 Excel 文件的文件夹")
        if folder:
            self.folder_path = folder
            self.folder_entry.delete(0, tk.END)
            self.folder_entry.insert(0, folder)
            self._load_files()

    def _load_files(self):
        """加载文件"""
        if not self.folder_path:
            return

        self.status_label.config(text="正在加载文件...")
        self.root.update()

        # 解析文件
        self.parsed_files = self.parser.parse_folder(self.folder_path)

        # 更新文件列表
        self._update_file_tree()

        # 更新状态
        total_records = sum(len(f.all_records) for f in self.parsed_files)
        success_count = sum(1 for f in self.parsed_files if not f.parse_errors)

        self.file_count_label.config(
            text=f"已加载 {len(self.parsed_files)} 个文件，{total_records} 条记录（{success_count} 个成功）"
        )
        self.status_label.config(text=f"就绪 | 已加载 {len(self.parsed_files)} 个文件")

    def _update_file_tree(self):
        """更新文件列表"""
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)

        for parsed in self.parsed_files:
            filename = Path(parsed.file_path).name
            company = parsed.company
            records = len(parsed.all_records)
            status = "成功" if not parsed.parse_errors else "有错误"

            self.file_tree.insert("", tk.END, values=(filename, company, records, status))

    def _start_merge(self):
        """开始合并"""
        if not self.parsed_files:
            messagebox.showwarning("警告", "请先选择包含 Excel 文件的文件夹")
            return

        self.status_label.config(text="正在合并...")
        self.root.update()

        # 在后台线程执行
        def do_merge():
            self.merge_result = self.matcher.match(self.parsed_files)
            self.root.after(0, self._on_merge_complete)

        threading.Thread(target=do_merge, daemon=True).start()

    def _on_merge_complete(self):
        """合并完成回调"""
        self._display_results()
        self._update_status()

        unmatched_count = len(self.merge_result.unmatched_accounts)
        messagebox.showinfo(
            "合并完成",
            f"合并完成！\n"
            f"总记录数: {self.merge_result.record_count}\n"
            f"公司数量: {self.merge_result.company_count}\n"
            f"未匹配科目: {unmatched_count}"
        )

    def _display_results(self):
        """显示合并结果"""
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        if not self.merge_result:
            return

        # 只显示前 100 条记录
        for record in self.merge_result.records[:100]:
            self.result_tree.insert("", tk.END, values=(
                record.standardized_name,
                record.company,
                f"{record.month}月",
                f"{float(record.amount):,.2f}"
            ))

    def _export_results(self):
        """导出结果"""
        if not self.merge_result:
            messagebox.showwarning("警告", "请先执行合并操作")
            return

        try:
            files = self.exporter.export_all(self.merge_result)
            messagebox.showinfo(
                "导出成功",
                f"已导出以下文件:\n" + "\n".join(files.values())
            )
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {e}")

    def _clear_data(self):
        """清空数据"""
        self.folder_path = ""
        self.parsed_files = []
        self.merge_result = None

        self.folder_entry.delete(0, tk.END)
        self.file_count_label.config(text="未选择文件夹")

        for item in self.file_tree.get_children():
            self.file_tree.delete(item)

        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        self.status_label.config(text="就绪 | 请选择包含 Excel 文件的文件夹")

    def _update_status(self):
        """更新状态栏"""
        if self.merge_result:
            self.status_label.config(
                text=f"就绪 | 记录: {self.merge_result.record_count} | "
                     f"公司: {self.merge_result.company_count} | "
                     f"未匹配科目: {len(self.merge_result.unmatched_accounts)}"
            )

    def _open_mapping_editor(self):
        """打开科目映射编辑器"""
        MappingEditor(self.root, self.matcher)

    def run(self):
        """运行应用"""
        self.root.mainloop()
```

- [ ] **Step 3: 提交**

```bash
git add tools/excel_merger/ui/
git commit -m "feat(excel-merger): 实现 GUI 主窗口

- 添加 ExcelMergerApp 主应用类
- 实现文件夹选择和文件列表预览
- 实现合并操作和结果预览
- 支持导出结果

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>"
```

---

## Task 7: 实现科目映射编辑器

**Files:**
- Create: `tools/excel_merger/ui/mapping_editor.py`

- [ ] **Step 1: 创建映射编辑器 `tools/excel_merger/ui/mapping_editor.py`**

```python
"""科目映射编辑器"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from excel_merger.core import AccountMatcher


class MappingEditor:
    """科目映射编辑器对话框"""

    def __init__(self, parent: tk.Tk, matcher: "AccountMatcher"):
        """
        初始化编辑器

        Args:
            parent: 父窗口
            matcher: 科目匹配器
        """
        self.matcher = matcher

        # 创建对话框
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("科目映射编辑器")
        self.dialog.geometry("600x400")
        self.dialog.transient(parent)
        self.dialog.grab_set()

        self._setup_ui()
        self._load_mappings()

    def _setup_ui(self):
        """设置 UI"""
        # 主容器
        main_frame = ttk.Frame(self.dialog, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 映射列表
        list_frame = ttk.LabelFrame(main_frame, text=" 当前映射 ", padding=5)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Treeview
        columns = ("original", "standardized")
        self.mapping_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=15)

        self.mapping_tree.heading("original", text="原始科目名称")
        self.mapping_tree.heading("standardized", text="标准化科目名称")

        self.mapping_tree.column("original", width=250)
        self.mapping_tree.column("standardized", width=250)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.mapping_tree.yview)
        self.mapping_tree.configure(yscrollcommand=scrollbar.set)

        self.mapping_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 添加/编辑区域
        edit_frame = ttk.LabelFrame(main_frame, text=" 添加/编辑映射 ", padding=5)
        edit_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(edit_frame, text="原始科目:").grid(row=0, column=0, padx=5, pady=5)
        self.original_entry = ttk.Entry(edit_frame, width=30)
        self.original_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(edit_frame, text="标准科目:").grid(row=0, column=2, padx=5, pady=5)
        self.standardized_entry = ttk.Entry(edit_frame, width=30)
        self.standardized_entry.grid(row=0, column=3, padx=5, pady=5)

        ttk.Button(edit_frame, text="添加", command=self._add_mapping).grid(row=0, column=4, padx=5, pady=5)
        ttk.Button(edit_frame, text="删除", command=self._delete_mapping).grid(row=0, column=5, padx=5, pady=5)

        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)

        ttk.Button(button_frame, text="保存", command=self._save_mappings).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="关闭", command=self.dialog.destroy).pack(side=tk.RIGHT, padx=5)

        # 绑定选择事件
        self.mapping_tree.bind("<<TreeviewSelect>>", self._on_select)

    def _load_mappings(self):
        """加载映射列表"""
        for item in self.mapping_tree.get_children():
            self.mapping_tree.delete(item)

        for original, standardized in self.matcher.account_mapping.items():
            self.mapping_tree.insert("", tk.END, values=(original, standardized))

    def _add_mapping(self):
        """添加映射"""
        original = self.original_entry.get().strip()
        standardized = self.standardized_entry.get().strip()

        if not original or not standardized:
            messagebox.showwarning("警告", "请输入原始科目名称和标准科目名称")
            return

        self.matcher.add_mapping(original, standardized)
        self._load_mappings()

        self.original_entry.delete(0, tk.END)
        self.standardized_entry.delete(0, tk.END)

    def _delete_mapping(self):
        """删除映射"""
        selected = self.mapping_tree.selection()
        if not selected:
            messagebox.showwarning("警告", "请选择要删除的映射")
            return

        item = selected[0]
        values = self.mapping_tree.item(item, "values")
        original = values[0]

        self.matcher.remove_mapping(original)
        self._load_mappings()

    def _on_select(self, event):
        """选择映射项"""
        selected = self.mapping_tree.selection()
        if selected:
            item = selected[0]
            values = self.mapping_tree.item(item, "values")

            self.original_entry.delete(0, tk.END)
            self.original_entry.insert(0, values[0])

            self.standardized_entry.delete(0, tk.END)
            self.standardized_entry.insert(0, values[1])

    def _save_mappings(self):
        """保存映射"""
        try:
            self.matcher.save_mapping("output/account_mapping.json")
            messagebox.showinfo("成功", "科目映射已保存")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {e}")
```

- [ ] **Step 2: 提交**

```bash
git add tools/excel_merger/ui/mapping_editor.py
git commit -m "feat(excel-merger): 实现科目映射编辑器

- 添加 MappingEditor 对话框类
- 支持查看、添加、删除科目映射
- 支持保存映射到文件

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>"
```

---

## Task 8: 创建主程序入口

**Files:**
- Create: `tools/excel_merger/main.py`

- [ ] **Step 1: 创建主程序入口 `tools/excel_merger/main.py`**

```python
"""Excel 多表合并工具箱 - 主程序入口"""

import sys
from pathlib import Path

# 添加项目根目录到路径
sys.path.insert(0, str(Path(__file__).parent.parent))

from excel_merger.ui import ExcelMergerApp


def main():
    """主函数"""
    app = ExcelMergerApp()
    app.run()


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: 提交**

```bash
git add tools/excel_merger/main.py
git commit -m "feat(excel-merger): 添加主程序入口

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>"
```

---

## Task 9: 编写单元测试

**Files:**
- Create: `tests/test_excel_merger.py`

- [ ] **Step 1: 创建测试文件 `tests/test_excel_merger.py`**

```python
"""
Excel 多表合并工具箱测试
"""

import pytest
from decimal import Decimal
from datetime import datetime
import tempfile
from pathlib import Path

import sys
sys.path.insert(0, str(Path(__file__).parent.parent / "tools"))

from excel_merger.models import SourceRecord, ParsedFile, SheetData, MergeResult
from excel_merger.core import ExcelParser, AccountMatcher, ResultExporter
from excel_merger.config import AppConfig, DEFAULT_ACCOUNT_MAPPING


class TestSourceRecord:
    """源数据记录测试"""

    def test_create_record(self):
        """测试创建记录"""
        record = SourceRecord(
            company="北京分公司",
            month=1,
            account_name="销售收入",
            amount=Decimal("10000.00")
        )
        assert record.company == "北京分公司"
        assert record.month == 1
        assert record.account_name == "销售收入"
        assert record.amount == Decimal("10000.00")

    def test_amount_conversion(self):
        """测试金额自动转换"""
        record = SourceRecord(
            company="测试",
            month=1,
            account_name="测试",
            amount=1000  # int 类型
        )
        assert isinstance(record.amount, Decimal)
        assert record.amount == Decimal("1000.00")


class TestParsedFile:
    """解析文件测试"""

    def test_all_records(self):
        """测试获取所有记录"""
        file = ParsedFile(file_path="test.xlsx", company="测试公司")

        sheet1 = SheetData(sheet_name="1月")
        sheet1.records = [
            SourceRecord(company="测试公司", month=1, account_name="收入", amount=Decimal("100")),
            SourceRecord(company="测试公司", month=1, account_name="成本", amount=Decimal("50"))
        ]

        sheet2 = SheetData(sheet_name="2月")
        sheet2.records = [
            SourceRecord(company="测试公司", month=2, account_name="收入", amount=Decimal("200"))
        ]

        file.sheets = [sheet1, sheet2]

        assert len(file.all_records) == 3
        assert len(file.all_accounts) == 3


class TestMergeResult:
    """合并结果测试"""

    def test_statistics(self):
        """测试统计属性"""
        result = MergeResult()
        result.records = [
            SourceRecord(company="北京", month=1, account_name="收入", amount=Decimal("100")),
            SourceRecord(company="上海", month=1, account_name="收入", amount=Decimal("200")),
            SourceRecord(company="北京", month=2, account_name="收入", amount=Decimal("150"))
        ]

        assert result.record_count == 3
        assert result.company_count == 2
        assert result.total_amount == 450.0


class TestAccountMatcher:
    """科目匹配器测试"""

    def test_exact_match(self):
        """测试精确匹配"""
        matcher = AccountMatcher()

        assert matcher._match_account("销售收入") == "主营业务收入"
        assert matcher._match_account("办公费") == "管理费用-办公费"

    def test_fuzzy_match(self):
        """测试模糊匹配"""
        matcher = AccountMatcher()

        # 包含关系匹配
        result = matcher._match_account("销售部门办公费")
        # 应该匹配到包含"办公费"的映射
        assert "办公费" in result or result == "销售部门办公费"

    def test_no_match(self):
        """测试未匹配"""
        matcher = AccountMatcher()

        result = matcher._match_account("未知科目XYZ")
        assert result == "未知科目XYZ"

    def test_add_mapping(self):
        """测试添加映射"""
        matcher = AccountMatcher()

        matcher.add_mapping("新科目", "标准科目")
        assert matcher._match_account("新科目") == "标准科目"

    def test_match_files(self):
        """测试匹配文件"""
        matcher = AccountMatcher()

        file = ParsedFile(file_path="test.xlsx", company="测试")
        sheet = SheetData(sheet_name="Sheet1")
        sheet.records = [
            SourceRecord(company="测试", month=1, account_name="销售收入", amount=Decimal("100")),
            SourceRecord(company="测试", month=1, account_name="未知科目", amount=Decimal("50"))
        ]
        file.sheets = [sheet]

        result = matcher.match([file])

        assert result.record_count == 2
        assert len(result.unmatched_accounts) == 1
        assert "未知科目" in result.unmatched_accounts


class TestResultExporter:
    """结果导出器测试"""

    def test_export_pivot_table(self):
        """测试导出透视表"""
        with tempfile.TemporaryDirectory() as tmpdir:
            exporter = ResultExporter(tmpdir)

            result = MergeResult()
            result.records = [
                SourceRecord(
                    company="北京", month=1, account_name="收入",
                    standardized_name="主营业务收入", amount=Decimal("100")
                ),
                SourceRecord(
                    company="北京", month=2, account_name="收入",
                    standardized_name="主营业务收入", amount=Decimal("200")
                ),
                SourceRecord(
                    company="上海", month=1, account_name="收入",
                    standardized_name="主营业务收入", amount=Decimal("150")
                )
            ]

            file_path = exporter.export_pivot_table(result, "test")
            assert Path(file_path).exists()

    def test_export_unmatched_report(self):
        """测试导出未匹配报告"""
        with tempfile.TemporaryDirectory() as tmpdir:
            exporter = ResultExporter(tmpdir)

            result = MergeResult()
            result.unmatched_accounts = {"未知科目1", "未知科目2"}
            result.records = [
                SourceRecord(company="测试", month=1, account_name="未知科目1", amount=Decimal("100"))
            ]

            file_path = exporter.export_unmatched_report(result, "test")
            assert Path(file_path).exists()


class TestExcelParser:
    """Excel 解析器测试"""

    def test_extract_company_from_filename(self):
        """测试从文件名提取公司名称"""
        parser = ExcelParser()

        assert parser._extract_company_from_filename("北京分公司.xlsx") == "北京"
        assert parser._extract_company_from_filename("上海_2024.xlsx") == "上海"
        assert parser._extract_company_from_filename("2024年广州.xlsx") == "广州"

    def test_extract_month_from_string(self):
        """测试从字符串提取月份"""
        parser = ExcelParser()

        assert parser._extract_month_from_string("1月") == 1
        assert parser._extract_month_from_string("一月") == 1
        assert parser._extract_month_from_string("Jan") == 1
        assert parser._extract_month_from_string("01") == 1
        assert parser._extract_month_from_string("12月") == 12
        assert parser._extract_month_from_string("无月份") is None

    def test_clean_amount(self):
        """测试金额清洗"""
        parser = ExcelParser()

        assert parser._clean_amount(1000) == Decimal("1000.00")
        assert parser._clean_amount("1,000.50") == Decimal("1000.50")
        assert parser._clean_amount("") == Decimal("0.00")
        assert parser._clean_amount("-") == Decimal("0.00")
        assert parser._clean_amount(None) == Decimal("0.00")
```

- [ ] **Step 2: 运行测试验证**

```bash
pytest tests/test_excel_merger.py -v
```

- [ ] **Step 3: 提交**

```bash
git add tests/test_excel_merger.py
git commit -m "test(excel-merger): 添加单元测试

- 添加数据模型测试
- 添加科目匹配器测试
- 添加结果导出器测试
- 添加解析器测试

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>"
```

---

## Task 10: 集成测试和文档

**Files:**
- Update: `README.md`

- [ ] **Step 1: 运行完整测试套件**

```bash
pytest tests/ -v
```

- [ ] **Step 2: 更新 README.md**

在 README.md 中添加 Excel 多表合并工具箱的说明：

```markdown
### Excel 多表合并工具箱

**位置**: `tools/excel_merger/`

**功能**: 自动读取文件夹内所有格式相近的 Excel，智能匹配科目名称，合并成标准化汇总表。

**使用方法**:
```bash
python -m tools.excel_merger.main
```

**特性**:
- 自动识别公司名称（从文件名）
- 智能识别月份（列标题、Sheet名）
- 内置常用科目映射表
- 支持自定义科目映射
- 生成科目×公司月份透视表
- 导出未匹配项报告
```

- [ ] **Step 3: 最终提交**

```bash
git add README.md
git commit -m "docs: 更新 README 添加 Excel 多表合并工具箱说明

Co-Authored-By: Claude Opus 4.7 <noreply@anthropic.com>"
```

---

## 验证方案

1. **单元测试**: 运行 `pytest tests/test_excel_merger.py -v` 验证所有测试通过
2. **集成测试**:
   - 准备测试数据文件夹，包含多个 Excel 文件
   - 运行 `python -m tools.excel_merger.main`
   - 选择测试文件夹
   - 检查预览结果
   - 导出并验证输出文件
3. **手动验证**:
   - 检查透视表格式是否正确
   - 检查未匹配项报告是否完整
   - 检查科目映射编辑器功能
