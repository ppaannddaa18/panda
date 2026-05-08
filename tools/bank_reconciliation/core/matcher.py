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
