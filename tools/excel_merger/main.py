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
