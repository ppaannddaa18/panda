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
