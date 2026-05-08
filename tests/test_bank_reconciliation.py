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
