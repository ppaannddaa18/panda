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
