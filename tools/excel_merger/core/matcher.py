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
