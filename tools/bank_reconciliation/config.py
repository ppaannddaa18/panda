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