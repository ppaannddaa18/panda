"""
银行对账助手

自动匹配银行流水与企业账务数据，输出对账结果和余额调节表。
"""

__version__ = "1.0.0"

from .main import main

__all__ = ["main"]