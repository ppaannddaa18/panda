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
