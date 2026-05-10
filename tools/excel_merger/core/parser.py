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